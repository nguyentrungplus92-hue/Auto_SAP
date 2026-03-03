"""
Django Management Command: scan_tasks
======================================
Chạy background scanner quét thư mục và trigger task tự động.
Hỗ trợ 2 chế độ:
- interval: quét liên tục theo chu kỳ (30s, 5 phút, ...)
- daily/weekly: chạy đúng giờ hẹn (8:00 sáng, ...)

Cách chạy:
    python manage.py scan_tasks              # Chạy liên tục
    python manage.py scan_tasks --once       # Chạy 1 lần rồi dừng
"""
import os
import re
import time
import importlib
import logging
from datetime import datetime, timedelta, date

from django.core.management.base import BaseCommand
from django.utils import timezone

from tasks.models import TaskConfig, TaskLog
from tasks.notifications import send_task_notification

log = logging.getLogger(__name__)


class Command(BaseCommand):
    help = 'Quét thư mục và tự động chạy SAP tasks khi phát hiện file mới'

    def add_arguments(self, parser):
        parser.add_argument(
            '--once',
            action='store_true',
            help='Chỉ quét 1 lần rồi dừng',
        )

    def handle(self, *args, **options):
        run_once = options['once']

        self.stdout.write("=" * 60)
        self.stdout.write(self.style.SUCCESS(" SAP Auto Scanner - Khởi động"))
        self.stdout.write("=" * 60)

        # Hiển thị danh sách task active
        active_tasks = TaskConfig.objects.filter(auto_enabled=True, status='active')
        for task in active_tasks:
            schedule_info = self._get_schedule_info(task)
            self.stdout.write(f"  📂 [{task.tcode}] {task.name}")
            self.stdout.write(f"     Lịch:      {schedule_info}")
            self.stdout.write(f"     Đường dẫn: {task.resolve_folder()}")

        if not active_tasks.exists():
            self.stdout.write(self.style.WARNING("  Chưa có task nào active!"))
            return

        self.stdout.write(f"\n  Đang chờ file mới... (Ctrl+C để dừng)\n")

        # Dict lưu thời điểm quét cuối của mỗi task (cho mode interval)
        last_scan = {}

        try:
            while True:
                now = datetime.now()
                today = now.date()
                current_time = now.time()
                current_weekday = now.weekday()  # 0=Monday, 6=Sunday
                
                # Reload tasks từ DB
                tasks = TaskConfig.objects.filter(auto_enabled=True, status='active')
                
                for task in tasks:
                    should_run = False
                    
                    if task.schedule_mode == 'interval':
                        # Mode interval: quét theo chu kỳ
                        should_run = self._check_interval(task, last_scan)
                        
                    elif task.schedule_mode == 'daily':
                        # Mode daily: chạy đúng giờ mỗi ngày
                        should_run = self._check_daily(task, today, current_time)
                        
                    elif task.schedule_mode == 'weekly':
                        # Mode weekly: chạy đúng giờ các ngày trong tuần
                        should_run = self._check_weekly(task, today, current_time, current_weekday)
                    
                    if should_run:
                        try:
                            self._scan_task(task)
                        except Exception as e:
                            log.error(f"[{task.tcode}] Scanner error: {e}")
                        
                        # Cập nhật thời điểm quét cuối (cho interval mode)
                        last_scan[task.id] = time.time()
                
                if run_once:
                    break
                
                # Sleep 10 giây rồi check lại
                time.sleep(10)
                
        except KeyboardInterrupt:
            self.stdout.write(self.style.SUCCESS("\n  Đã dừng scanner."))

    def _check_interval(self, task, last_scan):
        """Kiểm tra task interval có cần chạy không"""
        now = time.time()
        interval = task.scan_interval or 30
        
        if task.id in last_scan:
            elapsed = now - last_scan[task.id]
            if elapsed < interval:
                return False
        return True

    def _check_daily(self, task, today, current_time):
        """Kiểm tra task daily có cần chạy không"""
        if not task.scheduled_time:
            return False
        
        # Đã chạy hôm nay chưa?
        if task.last_scheduled_run == today:
            return False
        
        # Đã đến giờ chưa?
        scheduled = task.scheduled_time
        # Cho phép sai số 5 phút
        now_minutes = current_time.hour * 60 + current_time.minute
        scheduled_minutes = scheduled.hour * 60 + scheduled.minute
        
        if now_minutes >= scheduled_minutes and now_minutes <= scheduled_minutes + 5:
            self.stdout.write(f"  ⏰ [{task.tcode}] Đến giờ chạy: {scheduled.strftime('%H:%M')}")
            # Cập nhật ngay để tránh chạy lại
            task.last_scheduled_run = today
            task.save(update_fields=['last_scheduled_run'])
            return True
        
        return False

    def _check_weekly(self, task, today, current_time, current_weekday):
        """Kiểm tra task weekly có cần chạy không"""
        if not task.scheduled_time:
            return False
        
        # Đã chạy hôm nay chưa?
        if task.last_scheduled_run == today:
            return False
        
        # Hôm nay có trong danh sách ngày chạy không?
        scheduled_days = [int(d.strip()) for d in task.scheduled_days.split(',') if d.strip().isdigit()]
        
        # Convert: user format 1=T2, 2=T3, ... 5=T6, 6=T7, 0=CN
        # Python weekday: 0=Monday, 1=Tuesday, ... 6=Sunday
        python_to_user = {0: 1, 1: 2, 2: 3, 3: 4, 4: 5, 5: 6, 6: 0}
        user_weekday = python_to_user[current_weekday]
        
        if user_weekday not in scheduled_days:
            return False
        
        # Đã đến giờ chưa?
        scheduled = task.scheduled_time
        now_minutes = current_time.hour * 60 + current_time.minute
        scheduled_minutes = scheduled.hour * 60 + scheduled.minute
        
        if now_minutes >= scheduled_minutes and now_minutes <= scheduled_minutes + 5:
            self.stdout.write(f"  ⏰ [{task.tcode}] Đến giờ chạy: {scheduled.strftime('%H:%M')}")
            # Cập nhật ngay để tránh chạy lại
            task.last_scheduled_run = today
            task.save(update_fields=['last_scheduled_run'])
            return True
        
        return False

    def _get_schedule_info(self, task):
        """Trả về thông tin lịch dễ đọc"""
        if task.schedule_mode == 'interval':
            return f"Mỗi {self._format_interval(task.scan_interval)}"
        elif task.schedule_mode == 'daily':
            time_str = task.scheduled_time.strftime('%H:%M') if task.scheduled_time else 'N/A'
            return f"Hàng ngày lúc {time_str}"
        elif task.schedule_mode == 'weekly':
            time_str = task.scheduled_time.strftime('%H:%M') if task.scheduled_time else 'N/A'
            days_map = {1: 'T2', 2: 'T3', 3: 'T4', 4: 'T5', 5: 'T6', 6: 'T7', 0: 'CN'}
            days = task.scheduled_days or ''
            day_names = [days_map.get(int(d.strip()), d) for d in days.split(',') if d.strip().isdigit()]
            return f"{', '.join(day_names)} lúc {time_str}"
        return "N/A"

    def _format_interval(self, seconds):
        """Format interval thành dạng dễ đọc"""
        if not seconds:
            seconds = 30
        if seconds < 60:
            return f"{seconds} giây"
        elif seconds < 3600:
            minutes = seconds // 60
            return f"{minutes} phút"
        elif seconds < 86400:
            hours = seconds // 3600
            return f"{hours} giờ"
        else:
            days = seconds // 86400
            return f"{days} ngày"

    def _scan_task(self, task):
        """Quét 1 task - tìm file mới trong thư mục (hỗ trợ dynamic path)"""
        # Dùng resolve_folder() để tự động tính đường dẫn theo thời gian
        folder = task.resolve_folder()

        if not folder or not os.path.exists(folder):
            return

        # Lấy danh sách file đã xử lý
        processed_files = set(
            task.logs.values_list('filepath', flat=True)
        )

        # Quét thư mục
        for filename in os.listdir(folder):
            filepath = os.path.join(folder, filename)

            # Chỉ xử lý file
            if not os.path.isfile(filepath):
                continue

            # Đã xử lý rồi?
            if filepath in processed_files:
                continue

            # Khớp regex pattern?
            if not re.match(task.file_regex, filename):
                continue

            # Đợi file copy xong
            if not self._wait_file_ready(filepath):
                log.warning(f"[{task.tcode}] File chưa sẵn sàng: {filename}")
                continue

            # === Chạy handler ===
            self.stdout.write(
                self.style.HTTP_INFO(f"  [{task.tcode}] Phát hiện: {filename}")
            )
            self._execute_task(task, filepath, filename)

    def _execute_task(self, task, filepath, filename):
        """Gọi handler function và lưu log"""
        start_time = time.time()

        try:
            # Import handler dynamically
            handler_func = self._import_handler(task.handler_module)

            # Gọi handler (chưa có SAP session ở đây,
            # session sẽ được quản lý trong handler)
            result = handler_func(filepath, session=None)

            duration = round(time.time() - start_time, 2)

            # Lưu log
            TaskLog.objects.create(
                task=task,
                filename=filename,
                filepath=filepath,
                status=result.get('status', 'success'),
                message=result.get('message', ''),
                rows_processed=result.get('rows', 0),
                duration=duration,
            )

            status_display = result.get('status', 'unknown')
            if status_display == 'success':
                self.stdout.write(
                    self.style.SUCCESS(f"    ✅ {result.get('message', 'OK')} ({duration}s)")
                )
                # Gửi email thông báo thành công (nếu bật)
                log_entry = task.logs.order_by('-executed_at').first()
                if log_entry and task.notify_on_success:
                    send_task_notification(task, log_entry, 'success')
            elif status_display == 'error':
                self.stdout.write(
                    self.style.ERROR(f"    ❌ {result.get('message', 'Error')} ({duration}s)")
                )
                # Gửi email thông báo lỗi
                log_entry = task.logs.order_by('-executed_at').first()
                if log_entry and task.notify_on_error:
                    send_task_notification(task, log_entry, 'error')
            else:
                self.stdout.write(f"    ⏭️  {result.get('message', '')} ({duration}s)")

        except Exception as e:
            duration = round(time.time() - start_time, 2)
            log_entry = TaskLog.objects.create(
                task=task,
                filename=filename,
                filepath=filepath,
                status='error',
                message=str(e),
                duration=duration,
            )
            self.stdout.write(
                self.style.ERROR(f"    ❌ Exception: {e} ({duration}s)")
            )
            # Gửi email thông báo lỗi
            if task.notify_on_error:
                send_task_notification(task, log_entry, 'error')

    def _import_handler(self, handler_path):
        """Import handler function từ string path"""
        # "tasks.handlers.exchange_rate" -> module="tasks.handlers", func="exchange_rate"
        parts = handler_path.rsplit('.', 1)
        if len(parts) != 2:
            raise ValueError(f"Invalid handler path: {handler_path}")

        module_path, func_name = parts
        module = importlib.import_module(module_path)
        func = getattr(module, func_name)

        if not callable(func):
            raise ValueError(f"{handler_path} is not callable")

        return func

    def _wait_file_ready(self, filepath, timeout=30):
        """Đợi file copy xong (size ổn định)"""
        prev_size = -1
        waited = 0
        while waited < timeout:
            try:
                curr_size = os.path.getsize(filepath)
                if curr_size == prev_size and curr_size > 0:
                    return True
                prev_size = curr_size
            except OSError:
                pass
            time.sleep(1)
            waited += 1
        return False
