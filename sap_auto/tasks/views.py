import os, mimetypes
import re
import json
import importlib
import time

from django.shortcuts import render, get_object_or_404
from django.http import JsonResponse, FileResponse, Http404
from django.utils import timezone
from django.views.decorators.http import require_POST, require_GET
from django.views.decorators.csrf import csrf_exempt

from .models import TaskConfig, TaskLog, UserPermission


# ===== Helper: Lấy permission của user =====

def get_user_permission(username):
    """Lấy UserPermission theo username, trả về None nếu không có"""
    if not username:
        return None
    try:
        return UserPermission.objects.get(username=username, is_active=True)
    except UserPermission.DoesNotExist:
        return None


def get_username_from_request(request):
    """Lấy username từ request - ưu tiên session mẹ"""
    username = getattr(request, 'parent_username', None)
    if username:
        return username
    return ''


# ===== API: SAP Password =====

@require_GET
def api_get_sap_password(request):
    """API: Lấy mật khẩu SAP đã lưu của user hiện tại"""
    username = get_username_from_request(request)
    user_perm = get_user_permission(username)

    if not user_perm:
        return JsonResponse({'has_password': False, 'password': ''})

    return JsonResponse({
        'has_password': bool(user_perm.sap_password),
        'password': user_perm.sap_password or '',
    })


@csrf_exempt
@require_POST
def api_save_sap_password(request):
    """API: Lưu hoặc xóa mật khẩu SAP của user hiện tại"""
    username = get_username_from_request(request)
    user_perm = get_user_permission(username)

    if not user_perm:
        return JsonResponse({'error': 'Không có quyền'}, status=403)

    try:
        body = json.loads(request.body) if request.body else {}
    except Exception:
        body = {}

    password = body.get('password', '').strip()

    user_perm.sap_password = password
    user_perm.save(update_fields=['sap_password'])

    return JsonResponse({
        'status': 'saved' if password else 'cleared',
        'has_password': bool(password),
    })


# ===== API: Tasks =====

def api_tasks(request):
    """API: Danh sách tasks theo quyền"""
    username = get_username_from_request(request)
    user_perm = get_user_permission(username)

    if not user_perm:
        return JsonResponse({'tasks': []})

    tasks = user_perm.get_accessible_tasks()

    data = []
    for t in tasks:
        perm = user_perm.get_effective_permission(t)
        last_run = t.last_run
        data.append({
            'id': t.id,
            'module': t.module,
            'name': t.name,
            'tcode': t.tcode,
            'description': t.description,
            'watch_folder': t.watch_folder,
            'folder_template': t.folder_template,
            'file_pattern': t.file_pattern,
            'filename_template': t.filename_template,
            'file_regex': t.file_regex,
            'handler_module': t.handler_module,
            'param1': t.param1,
            'param2': t.param2,
            'status': t.status,
            'auto_enabled': t.auto_enabled,
            'can_edit': perm['can_edit'] if perm else False,
            'can_run': perm['can_run'] if perm else False,
            'can_delete': perm['can_delete'] if perm else False,
            'last_run': {
                'status': last_run.status,
                'time': last_run.executed_at.isoformat(),
                'message': last_run.message,
            } if last_run else None,
        })

    return JsonResponse({'tasks': data})


@csrf_exempt
@require_POST
def api_task_run(request, pk):
    """API: Chạy task thủ công"""
    username = get_username_from_request(request)
    user_perm = get_user_permission(username)
    task = get_object_or_404(TaskConfig, pk=pk)

    if not user_perm:
        return JsonResponse({'error': 'Không có quyền'}, status=403)

    perm = user_perm.get_effective_permission(task)
    if not perm:
        return JsonResponse({'error': 'Không có quyền truy cập task này'}, status=403)

    if not perm['can_run']:
        return JsonResponse({'error': 'Không có quyền chạy task'}, status=403)

    # ===== Đọc manual_password và save_password từ body =====
    try:
        body = json.loads(request.body) if request.body else {}
    except Exception:
        body = {}

    manual_password = body.get('manual_password', '').strip()
    save_password = body.get('save_password', False)

    if not manual_password:
        return JsonResponse({'error': 'Vui lòng nhập mật khẩu SAP'}, status=400)

    # Lưu hoặc xóa mật khẩu theo lựa chọn của user
    if save_password:
        user_perm.sap_password = manual_password
        user_perm.save(update_fields=['sap_password'])
    else:
        if user_perm.sap_password:
            user_perm.sap_password = ''
            user_perm.save(update_fields=['sap_password'])

    # run_options: dùng username đang login + password vừa nhập
    run_options = {
        'username': username,
        'password': manual_password,
    }
    # =======================================================

    folder = task.resolve_folder()
    filename = task.resolve_filename()

    # Nếu có filename cụ thể → xử lý đúng file đó
    if filename:
        filepath = os.path.join(folder, filename)

        if not os.path.exists(filepath):
            return JsonResponse({'error': f'File không tồn tại: {filepath}'}, status=400)

        start = time.time()
        try:
            handler_func = _import_handler(task.handler_module)
            result = handler_func(filepath, session=None, task=task, run_options=run_options)
            duration = round(time.time() - start, 2)

            # Set status partial nếu success nhưng có dòng lỗi
            errors = result.get('errors', [])
            log_status = result.get('status', 'success')
            if log_status == 'success' and errors:
                log_status = 'partial'

            TaskLog.objects.create(
                task=task, filename=filename, filepath=filepath,
                status=log_status,
                message=result.get('message', ''),
                rows_processed=result.get('rows', 0),
                duration=duration,
                executed_by=username,
            )
            return JsonResponse({
                'status': 'done',
                'filename': filename,
                'result': log_status,
                'message': result.get('message', ''),
                'errors': errors,
            })
        except Exception as e:
            TaskLog.objects.create(
                task=task, filename=filename, filepath=filepath,
                status='error', message=str(e),
                duration=round(time.time() - start, 2),
                executed_by=username,
            )
            return JsonResponse({'status': 'error', 'filename': filename, 'message': str(e)})

    # Nếu không có filename cụ thể → scan thư mục
    if not folder or not os.path.exists(folder):
        return JsonResponse({'error': f'Thư mục không tồn tại: {folder}'}, status=400)

    results = []

    for fname in os.listdir(folder):
        fpath = os.path.join(folder, fname)
        if not os.path.isfile(fpath):
            continue
        if task.file_regex and not re.match(task.file_regex, fname):
            continue

        start = time.time()
        try:
            handler_func = _import_handler(task.handler_module)
            result = handler_func(fpath, session=None, task=task, run_options=run_options)
            duration = round(time.time() - start, 2)

            # Set status partial nếu success nhưng có dòng lỗi
            errors = result.get('errors', [])
            log_status = result.get('status', 'success')
            if log_status == 'success' and errors:
                log_status = 'partial'

            TaskLog.objects.create(
                task=task, filename=fname, filepath=fpath,
                status=log_status,
                message=result.get('message', ''),
                rows_processed=result.get('rows', 0),
                duration=duration,
                executed_by=username,
            )
            results.append({
                'filename': fname,
                'status': log_status,
                'message': result.get('message'),
                'errors': errors,
            })
        except Exception as e:
            TaskLog.objects.create(
                task=task, filename=fname, filepath=fpath,
                status='error', message=str(e),
                duration=round(time.time() - start, 2),
                executed_by=username,
            )
            results.append({'filename': fname, 'status': 'error', 'message': str(e)})

    if not results:
        return JsonResponse({'status': 'info', 'message': 'Không có file mới'})
    return JsonResponse({'status': 'done', 'results': results, 'total': len(results)})


@csrf_exempt
@require_POST
def api_task_toggle(request, pk):
    """API: Bật/tắt auto (Admin only)"""
    username = get_username_from_request(request)
    user_perm = get_user_permission(username)
    task = get_object_or_404(TaskConfig, pk=pk)

    if not user_perm:
        return JsonResponse({'error': 'Không có quyền'}, status=403)

    if not user_perm.is_admin:
        return JsonResponse({'error': 'Chỉ Admin được bật/tắt Auto'}, status=403)

    task.auto_enabled = not task.auto_enabled
    task.save(update_fields=['auto_enabled'])
    return JsonResponse({'id': task.id, 'auto_enabled': task.auto_enabled})


@csrf_exempt
@require_POST
def api_task_create(request):
    """API: Tạo task mới"""
    username = get_username_from_request(request)
    user_perm = get_user_permission(username)

    if not user_perm:
        return JsonResponse({'error': 'Không có quyền'}, status=403)

    try:
        data = json.loads(request.body)
        module = data.get('module', 'OTHER')

        can_create = False

        module_perm = user_perm.get_module_permission(module)
        if module_perm and module_perm.can_edit:
            can_create = True

        if not can_create:
            group_module_perm = user_perm.get_group_module_permission(module)
            if group_module_perm and group_module_perm.can_edit:
                can_create = True

        if user_perm.is_admin:
            can_create = True

        if not can_create:
            return JsonResponse({'error': f'Không có quyền tạo task trong module {module}'}, status=403)

        task = TaskConfig.objects.create(
            module=module,
            name=data['name'],
            tcode=data.get('tcode', ''),
            description=data.get('description', ''),
            watch_folder=data.get('watch_folder', ''),
            folder_template=data.get('folder_template', ''),
            file_pattern=data.get('file_pattern', ''),
            filename_template=data.get('filename_template', ''),
            file_regex=data.get('file_regex', ''),
            handler_module=data.get('handler_module', 'tasks.handlers.default_handler'),
            param1=data.get('param1', ''),
            param2=data.get('param2', ''),
        )
        return JsonResponse({'id': task.id, 'status': 'created'})
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=400)


@csrf_exempt
@require_POST
def api_task_update(request, pk):
    """API: Cập nhật task"""
    username = get_username_from_request(request)
    user_perm = get_user_permission(username)
    task = get_object_or_404(TaskConfig, pk=pk)

    if not user_perm:
        return JsonResponse({'error': 'Không có quyền'}, status=403)

    perm = user_perm.get_effective_permission(task)
    if not perm or not perm['can_edit']:
        return JsonResponse({'error': 'Không có quyền sửa task'}, status=403)

    try:
        data = json.loads(request.body)
        task.module = data.get('module', task.module)
        task.name = data.get('name', task.name)
        task.tcode = data.get('tcode', task.tcode)
        task.description = data.get('description', task.description)
        task.watch_folder = data.get('watch_folder', task.watch_folder)
        task.folder_template = data.get('folder_template', task.folder_template)
        task.file_pattern = data.get('file_pattern', task.file_pattern)
        task.filename_template = data.get('filename_template', task.filename_template)
        task.file_regex = data.get('file_regex', task.file_regex)
        task.handler_module = data.get('handler_module', task.handler_module)
        task.param1 = data.get('param1', task.param1)
        task.param2 = data.get('param2', task.param2)
        task.save()
        return JsonResponse({'id': task.id, 'status': 'updated'})
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=400)


@csrf_exempt
@require_POST
def api_task_delete(request, pk):
    """API: Xóa task"""
    username = get_username_from_request(request)
    user_perm = get_user_permission(username)
    task = get_object_or_404(TaskConfig, pk=pk)

    if not user_perm:
        return JsonResponse({'error': 'Không có quyền'}, status=403)

    perm = user_perm.get_effective_permission(task)
    if not perm or not perm['can_delete']:
        return JsonResponse({'error': 'Không có quyền xóa task'}, status=403)

    task.delete()
    return JsonResponse({'status': 'deleted'})


def api_task_logs(request, pk):
    """API: Log của task"""
    username = get_username_from_request(request)
    user_perm = get_user_permission(username)
    task = get_object_or_404(TaskConfig, pk=pk)

    if not user_perm or not user_perm.can_access_task(task):
        return JsonResponse({'error': 'Không có quyền'}, status=403)

    logs = task.logs.all()[:100]
    data = [{
        'id': l.id,
        'filename': l.filename,
        'status': l.status,
        'message': l.message,
        'rows': l.rows_processed,
        'duration': l.duration,
        'executed_by': l.executed_by,
        'time': l.executed_at.isoformat(),
    } for l in logs]
    return JsonResponse({'logs': data})


def api_scan_status(request):
    """API: Trạng thái scanner"""
    import subprocess
    try:
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq python.exe', '/FO', 'CSV'],
            capture_output=True, text=True, timeout=5
        )
        scanner_running = 'scan_tasks' in result.stdout
    except Exception:
        scanner_running = False
    return JsonResponse({'scanner_running': scanner_running})


# ===== Dashboard =====

def dashboard(request):
    """Dashboard chính - yêu cầu login từ chương trình mẹ."""
    username = getattr(request, 'parent_username', None)
    module_filter = request.GET.get('module', '')

    if not username:
        return render(request, 'tasks/no_login.html')

    user_perm = get_user_permission(username)

    if not user_perm:
        return render(request, 'tasks/no_permission.html', {'username': username})

    tasks = user_perm.get_accessible_tasks()

    if module_filter:
        tasks = tasks.filter(module=module_filter)

    tasks_with_perm = []
    for task in tasks:
        perm = user_perm.get_effective_permission(task)
        tasks_with_perm.append({
            'task': task,
            'perm': perm,
        })

    all_modules = user_perm.get_accessible_modules()
    module_order = [code for code, name in TaskConfig.MODULE_CHOICES]
    all_modules_sorted = sorted(all_modules, key=lambda x: module_order.index(x) if x in module_order else 999)
    modules = [{'code': m, 'name': dict(TaskConfig.MODULE_CHOICES).get(m, m)} for m in all_modules_sorted]

    all_tasks = user_perm.get_accessible_tasks()
    stats = {
        'total': all_tasks.count(),
        'active': all_tasks.filter(status='active').count(),
        'paused': all_tasks.filter(status='paused').count(),
        'today_success': TaskLog.objects.filter(
            task__in=all_tasks,
            executed_at__date=timezone.now().date(),
            status='success'
        ).count(),
        'today_errors': TaskLog.objects.filter(
            task__in=all_tasks,
            executed_at__date=timezone.now().date(),
            status='error'
        ).count(),
    }

    return render(request, 'tasks/dashboard.html', {
        'tasks_with_perm': tasks_with_perm,
        'stats': stats,
        'modules': modules,
        'current_module': module_filter,
        'module_choices': TaskConfig.MODULE_CHOICES,
        'username': username,
        'user_perm': user_perm,
    })


def task_detail(request, pk):
    """Chi tiết task + lịch sử log"""
    username = getattr(request, 'parent_username', None)

    if not username:
        return render(request, 'tasks/no_login.html')

    user_perm = get_user_permission(username)
    if not user_perm:
        return render(request, 'tasks/no_permission.html', {'username': username})

    task = get_object_or_404(TaskConfig, pk=pk)

    perm = user_perm.get_effective_permission(task)
    if not perm or not perm.get('can_view', False):
        return render(request, 'tasks/no_permission.html', {'username': username})

    logs = task.logs.all()[:50]

    return render(request, 'tasks/detail.html', {
        'task': task,
        'logs': logs,
        'username': username,
        'user_perm': user_perm,
        'perm': perm,
    })


def _import_handler(handler_path):
    parts = handler_path.rsplit('.', 1)
    if len(parts) != 2:
        raise ValueError(f"Invalid handler path: {handler_path}")
    module_path, func_name = parts
    module = importlib.import_module(module_path)
    return getattr(module, func_name)


def api_task_check_file(request, pk):
    """API: Kiểm tra file tồn tại trước khi chạy"""
    task = get_object_or_404(TaskConfig, pk=pk)
    folder = task.resolve_folder()
    filename = task.resolve_filename()
    if filename:
        filepath = os.path.join(folder, filename)
        if not os.path.exists(filepath):
            return JsonResponse({'exists': False, 'message': f'File không tồn tại: {filepath}'})
    elif folder and not os.path.exists(folder):
        return JsonResponse({'exists': False, 'message': f'Thư mục không tồn tại: {folder}'})
    return JsonResponse({'exists': True})


def download_sample(request, pk):
    task = get_object_or_404(TaskConfig, pk=pk)
    filepath = task.sample_file_path
    if not filepath or not os.path.exists(filepath):
        raise Http404("File mẫu không tồn tại")
    filename = os.path.basename(filepath)
    response = FileResponse(open(filepath, 'rb'), as_attachment=True, filename=filename)
    return response


def api_upload_file(request, pk):
    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)
    task = get_object_or_404(TaskConfig, pk=pk)
    uploaded_file = request.FILES.get('file')
    if not uploaded_file:
        return JsonResponse({'error': 'Không có file'}, status=400)
    folder = task.resolve_folder()
    if not folder or not os.path.exists(folder):
        return JsonResponse({'error': f'Thư mục không tồn tại: {folder}'}, status=400)

    # ===== Xóa file cũ trước khi lưu file mới =====
    old_filename = task.resolve_filename()
    if old_filename:
        old_filepath = os.path.join(folder, old_filename)
        if os.path.exists(old_filepath) and old_filepath != os.path.join(folder, uploaded_file.name):
            try:
                os.remove(old_filepath)
                log.info(f"[upload] Đã xóa file cũ: {old_filepath}")
            except Exception as e:
                log.warning(f"[upload] Không xóa được file cũ: {e}")
    # =================================================

    # Lưu file mới
    filepath = os.path.join(folder, uploaded_file.name)
    with open(filepath, 'wb') as f:
        for chunk in uploaded_file.chunks():
            f.write(chunk)

    # Cập nhật tên file vào DB
    task.file_pattern = uploaded_file.name
    task.filename_template = ''
    task.save(update_fields=['file_pattern', 'filename_template'])

    return JsonResponse({'ok': True, 'message': f'Đã upload: {uploaded_file.name}', 'filename': uploaded_file.name})