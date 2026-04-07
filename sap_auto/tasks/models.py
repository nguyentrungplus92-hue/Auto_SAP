from django.db import models
from django.utils import timezone
from datetime import datetime


# ===== Task Config =====

class TaskConfig(models.Model):
    MODULE_CHOICES = [
        ('MM', 'MM - Material Management'),
        ('PP', 'PP - Production Planning'),
        ('FICO', 'FICO - Finance & Controlling'),
        ('SD', 'SD - Sales & Distribution'),
        ('QM', 'QM - Quality Management'),
        ('WM', 'WM - Warehouse Management'),
        ('HR', 'HR - Human Resources'),
        ('BASIS', 'BASIS - Basis/Admin'),
        ('OTHER', 'OTHER - Other'),
    ]
    
    STATUS_CHOICES = [
        ('active', 'Đang hoạt động'),
        ('paused', 'Tạm dừng'),
    ]
    
    SCHEDULE_MODE_CHOICES = [
        ('interval', 'Liên tục (theo interval)'),
        ('daily', 'Hàng ngày (1 lần/ngày)'),
        ('weekly', 'Hàng tuần'),
    ]

    module = models.CharField("Module SAP", max_length=20, choices=MODULE_CHOICES, default='OTHER')
    name = models.CharField("Tên task", max_length=200)
    tcode = models.CharField("T-Code SAP", max_length=20, blank=True)
    description = models.TextField("Mô tả", blank=True)
    
    watch_folder = models.CharField("Thư mục theo dõi", max_length=500, blank=True)
    folder_template = models.CharField("Template đường dẫn", max_length=500, blank=True,
        help_text="VD: G:\\Data\\FY{yyyy}\\{mm}.{yyyy}\\ - Placeholder: {yyyy} {yy} {mm} {dd} {mmm} {FY}")
    
    file_pattern = models.CharField("Pattern tên file", max_length=200, blank=True)
    filename_template = models.CharField("Template tên file", max_length=300, blank=True,
        help_text="VD: Report_{dd}-{mmm}-{yyyy}.xlsx")
    file_regex = models.CharField("Regex match file", max_length=300, blank=True)
    
    handler_module = models.CharField("Handler module", max_length=200, 
        default='tasks.handlers.default_handler')
    param1 = models.CharField("Tham số 1", max_length=200, blank=True)
    param2 = models.CharField("Tham số 2", max_length=200, blank=True)
    sap_user = models.ForeignKey('SAPUser', on_delete=models.SET_NULL, null=True, blank=True,
        verbose_name="SAP User", help_text="User SAP để chạy task này (dùng cho auto run)")
    
    status = models.CharField("Trạng thái", max_length=20, choices=STATUS_CHOICES, default='active')
    auto_enabled = models.BooleanField("Bật auto scan", default=False)
    
    schedule_mode = models.CharField("Chế độ lịch", max_length=20, choices=SCHEDULE_MODE_CHOICES, default='interval')
    scheduled_time = models.TimeField("Giờ chạy", null=True, blank=True, help_text="Cho daily/weekly")
    scheduled_days = models.CharField("Ngày chạy", max_length=50, blank=True, help_text="0-6 (T2-CN), VD: 0,1,2,3,4")
    last_scheduled_run = models.DateField("Lần chạy scheduled cuối", null=True, blank=True)
    
    notify_on_error = models.BooleanField("Gửi mail khi lỗi", default=False)
    notify_on_success = models.BooleanField("Gửi mail khi thành công", default=False)
    notify_emails = models.TextField("Email nhận thông báo", blank=True, help_text="Mỗi email 1 dòng")

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['module', 'name']
        verbose_name = "Task Config"
        verbose_name_plural = "Task Configs"

    def __str__(self):
        return f"[{self.module}] {self.name}"
    
    def resolve_folder(self, date=None):
        if not self.folder_template:
            return self.watch_folder
        return self._resolve_template(self.folder_template, date)
    
    def resolve_filename(self, date=None):
        if not self.filename_template:
            return self.file_pattern
        return self._resolve_template(self.filename_template, date)
    
    def _resolve_template(self, template, date=None):
        if not template:
            return ''
        d = date or timezone.localtime(timezone.now())
        fy = d.year if d.month >= 4 else d.year - 1
        replacements = {
            '{yyyy}': str(d.year),
            '{yy}': str(d.year)[2:],
            '{mm}': f'{d.month:02d}',
            '{dd}': f'{d.day:02d}',
            '{mmm}': d.strftime('%b'),
            '{mmmm}': d.strftime('%B'),
            '{FY}': f'FY{fy}',
            '{fy}': str(fy),
            '{mm.yyyy}': f'{d.month:02d}.{d.year}',
        }
        result = template
        for k, v in replacements.items():
            result = result.replace(k, v)
        return result

    @property
    def last_run(self):
        return self.logs.first()
    
    @property
    def email_list(self):
        if not self.notify_emails:
            return []
        return [e.strip() for e in self.notify_emails.strip().split('\n') if e.strip()]


# ===== Task Log =====

class TaskLog(models.Model):
    STATUS_CHOICES = [
        ('success', 'Thành công'),
        ('partial', 'Có dòng lỗi'),
        ('error', 'Lỗi'),
        ('skipped', 'Bỏ qua'),
    ]

    task = models.ForeignKey(TaskConfig, on_delete=models.CASCADE, related_name='logs')
    filename = models.CharField("Tên file", max_length=300)
    filepath = models.CharField("Đường dẫn đầy đủ", max_length=500)
    status = models.CharField("Trạng thái", max_length=20, choices=STATUS_CHOICES)
    message = models.TextField("Thông báo", blank=True)
    rows_processed = models.IntegerField("Số dòng xử lý", default=0)
    duration = models.FloatField("Thời gian (s)", default=0)
    executed_by = models.CharField("Người thực hiện", max_length=100, blank=True)
    executed_at = models.DateTimeField("Thời điểm", auto_now_add=True)

    class Meta:
        ordering = ['-executed_at']
        verbose_name = "Task Log"
        verbose_name_plural = "Task Logs"

    def __str__(self):
        return f"{self.task.name} - {self.filename} - {self.status}"


# ===== SAP User =====

class SAPUser(models.Model):
    """Thông tin đăng nhập SAP - dùng cho auto run"""
    
    client = models.CharField("Client", max_length=10, help_text="SAP Client (VD: 100, 200)")
    username = models.CharField("User SAP", max_length=50)
    password = models.CharField("Password", max_length=200)
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        verbose_name = "SAP User"
        verbose_name_plural = "SAP Users"
        unique_together = ['client', 'username']
    
    def __str__(self):
        return f"{self.client} - {self.username}"


# ===== User Group (Nhóm người dùng) =====

class UserGroup(models.Model):
    """Nhóm người dùng - dùng để gán quyền chung cho nhiều user"""
    
    name = models.CharField("Tên nhóm", max_length=100, unique=True)
    description = models.TextField("Mô tả", blank=True)
    is_active = models.BooleanField("Kích hoạt", default=True)
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        verbose_name = "Nhóm người dùng"
        verbose_name_plural = "Nhóm người dùng"
        ordering = ['name']
    
    def __str__(self):
        return self.name
    
    def get_accessible_modules(self):
        """Lấy danh sách module nhóm được phép truy cập"""
        return list(self.module_permissions.values_list('module', flat=True))
    
    def get_accessible_tasks(self):
        """Lấy danh sách task nhóm được phép truy cập"""
        modules = self.get_accessible_modules()
        task_ids = list(self.task_permissions.values_list('task_id', flat=True))
        
        from django.db.models import Q
        return TaskConfig.objects.filter(
            Q(module__in=modules) | Q(id__in=task_ids)
        ).distinct()


# ===== Group Module Permission =====

class GroupModulePermission(models.Model):
    """Quyền theo module cho nhóm"""
    
    group = models.ForeignKey(
        UserGroup, 
        on_delete=models.CASCADE, 
        related_name='module_permissions'
    )
    module = models.CharField("Module", max_length=20, choices=TaskConfig.MODULE_CHOICES)
    
    can_view = models.BooleanField("Được xem", default=True)
    can_edit = models.BooleanField("Được sửa", default=False)
    can_run = models.BooleanField("Được chạy", default=True)
    can_delete = models.BooleanField("Được xóa", default=False)
    
    class Meta:
        verbose_name = "Group Module Permission"
        verbose_name_plural = "Group Module Permissions"
        unique_together = ['group', 'module']
    
    def __str__(self):
        perms = []
        if self.can_view: perms.append('Xem')
        if self.can_edit: perms.append('Sửa')
        if self.can_run: perms.append('Chạy')
        if self.can_delete: perms.append('Xóa')
        return f"{self.group.name} - {self.module}: {', '.join(perms) or 'Không có quyền'}"


# ===== Group Task Permission =====

class GroupTaskPermission(models.Model):
    """Quyền cụ thể cho từng task của nhóm"""
    
    group = models.ForeignKey(
        UserGroup, 
        on_delete=models.CASCADE, 
        related_name='task_permissions'
    )
    task = models.ForeignKey(
        TaskConfig, 
        on_delete=models.CASCADE, 
        related_name='group_permissions'
    )
    
    can_view = models.BooleanField("Được xem", default=True)
    can_edit = models.BooleanField("Được sửa", default=False)
    can_run = models.BooleanField("Được chạy", default=True)
    can_delete = models.BooleanField("Được xóa", default=False)
    
    class Meta:
        verbose_name = "Group Task Permission"
        verbose_name_plural = "Group Task Permissions"
        unique_together = ['group', 'task']
    
    def __str__(self):
        perms = []
        if self.can_view: perms.append('Xem')
        if self.can_edit: perms.append('Sửa')
        if self.can_run: perms.append('Chạy')
        if self.can_delete: perms.append('Xóa')
        return f"{self.group.name} - {self.task.name}: {', '.join(perms) or 'Không có quyền'}"


# ===== User Permission (Thông tin cơ bản) =====

class UserPermission(models.Model):
    """Thông tin user và quyền admin"""
    
    username = models.CharField(
        "Username",
        max_length=100,
        unique=True,
        help_text="Username từ chương trình mẹ"
    )
    display_name = models.CharField("Tên hiển thị", max_length=200, blank=True)
    is_admin = models.BooleanField(
        "Quyền Admin",
        default=False,
        help_text="Cho phép truy cập trang Admin"
    )
    is_active = models.BooleanField("Kích hoạt", default=True)

    # ===== Mật khẩu SAP riêng của user (dùng khi chạy thủ công) =====
    sap_password = models.CharField(
        "Mật khẩu SAP",
        max_length=200,
        blank=True,
        help_text="Mật khẩu SAP của user này, dùng khi chạy task thủ công. "
                  "Khác với mật khẩu của SAPUser (dùng cho auto run)."
    )
    # =================================================================
    
    # Thêm quan hệ với nhóm
    groups = models.ManyToManyField(
        UserGroup,
        blank=True,
        related_name='users',
        verbose_name="Nhóm"
    )
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        verbose_name = "User"
        verbose_name_plural = "Users"
    
    def __str__(self):
        admin_tag = " [ADMIN]" if self.is_admin else ""
        groups_str = ", ".join(self.groups.values_list('name', flat=True))
        if groups_str:
            return f"{self.username}{admin_tag} ({groups_str})"
        return f"{self.username}{admin_tag}"
    
    def get_module_permission(self, module):
        """Lấy quyền cho module cụ thể (user level)"""
        try:
            return self.module_permissions.get(module=module)
        except ModulePermission.DoesNotExist:
            return None
    
    def get_task_permission(self, task):
        """Lấy quyền cho task cụ thể (user level)"""
        try:
            return self.task_permissions.get(task=task)
        except TaskPermission.DoesNotExist:
            return None
    
    def get_group_module_permission(self, module):
        """Lấy quyền module từ các nhóm user thuộc về"""
        for group in self.groups.filter(is_active=True):
            try:
                perm = group.module_permissions.get(module=module)
                return perm
            except GroupModulePermission.DoesNotExist:
                continue
        return None
    
    def get_group_task_permission(self, task):
        """Lấy quyền task từ các nhóm user thuộc về"""
        for group in self.groups.filter(is_active=True):
            try:
                perm = group.task_permissions.get(task=task)
                return perm
            except GroupTaskPermission.DoesNotExist:
                continue
        return None
    
    def get_effective_permission(self, task):
        """
        Lấy quyền hiệu lực cho task.
        Thứ tự ưu tiên: Admin > User Task > User Module > Group Task > Group Module > None
        """
        if self.is_admin:
            return {
                'can_view': True,
                'can_edit': True,
                'can_run': True,
                'can_delete': True,
                'source': 'admin',
            }
        
        task_perm = self.get_task_permission(task)
        if task_perm:
            return {
                'can_view': task_perm.can_view,
                'can_edit': task_perm.can_edit,
                'can_run': task_perm.can_run,
                'can_delete': task_perm.can_delete,
                'source': 'user_task',
            }
        
        module_perm = self.get_module_permission(task.module)
        if module_perm:
            return {
                'can_view': module_perm.can_view,
                'can_edit': module_perm.can_edit,
                'can_run': module_perm.can_run,
                'can_delete': module_perm.can_delete,
                'source': 'user_module',
            }
        
        group_task_perm = self.get_group_task_permission(task)
        if group_task_perm:
            return {
                'can_view': group_task_perm.can_view,
                'can_edit': group_task_perm.can_edit,
                'can_run': group_task_perm.can_run,
                'can_delete': group_task_perm.can_delete,
                'source': 'group_task',
            }
        
        group_module_perm = self.get_group_module_permission(task.module)
        if group_module_perm:
            return {
                'can_view': group_module_perm.can_view,
                'can_edit': group_module_perm.can_edit,
                'can_run': group_module_perm.can_run,
                'can_delete': group_module_perm.can_delete,
                'source': 'group_module',
            }
        
        return None
    
    def can_access_task(self, task):
        """Kiểm tra có quyền xem task không"""
        perm = self.get_effective_permission(task)
        return perm is not None and perm.get('can_view', False)
    
    def get_accessible_modules(self):
        """Lấy danh sách module được phép truy cập"""
        if self.is_admin:
            return [code for code, name in TaskConfig.MODULE_CHOICES]
        
        modules = set()
        modules.update(self.module_permissions.values_list('module', flat=True))
        
        for group in self.groups.filter(is_active=True):
            modules.update(group.module_permissions.values_list('module', flat=True))
        
        user_task_modules = TaskConfig.objects.filter(
            id__in=self.task_permissions.values_list('task_id', flat=True)
        ).values_list('module', flat=True)
        modules.update(user_task_modules)
        
        for group in self.groups.filter(is_active=True):
            group_task_modules = TaskConfig.objects.filter(
                id__in=group.task_permissions.values_list('task_id', flat=True)
            ).values_list('module', flat=True)
            modules.update(group_task_modules)
        
        return list(modules)

    def get_accessible_tasks(self):
        """Lấy danh sách task được phép truy cập"""
        from django.db.models import Q
        
        if self.is_admin:
            return TaskConfig.objects.all()
        
        user_modules = set(self.module_permissions.values_list('module', flat=True))
        
        group_modules = set()
        for group in self.groups.filter(is_active=True):
            group_modules.update(group.module_permissions.values_list('module', flat=True))
        
        all_modules = user_modules | group_modules
        
        user_task_ids = set(self.task_permissions.values_list('task_id', flat=True))
        
        group_task_ids = set()
        for group in self.groups.filter(is_active=True):
            group_task_ids.update(group.task_permissions.values_list('task_id', flat=True))
        
        all_task_ids = user_task_ids | group_task_ids
        
        return TaskConfig.objects.filter(
            Q(module__in=all_modules) | Q(id__in=all_task_ids)
        ).distinct()


# ===== Module Permission (User level) =====

class ModulePermission(models.Model):
    """Quyền theo module cho user"""
    
    user = models.ForeignKey(
        UserPermission, 
        on_delete=models.CASCADE, 
        related_name='module_permissions'
    )
    module = models.CharField("Module", max_length=20, choices=TaskConfig.MODULE_CHOICES)
    
    can_view = models.BooleanField("Được xem", default=True)
    can_edit = models.BooleanField("Được sửa", default=False)
    can_run = models.BooleanField("Được chạy", default=True)
    can_delete = models.BooleanField("Được xóa", default=False)
    
    class Meta:
        verbose_name = "Module Permission"
        verbose_name_plural = "Module Permissions"
        unique_together = ['user', 'module']
    
    def __str__(self):
        perms = []
        if self.can_view: perms.append('Xem')
        if self.can_edit: perms.append('Sửa')
        if self.can_run: perms.append('Chạy')
        if self.can_delete: perms.append('Xóa')
        return f"{self.user.username} - {self.module}: {', '.join(perms) or 'Không có quyền'}"


# ===== Task Permission (User level - Override module permission) =====

class TaskPermission(models.Model):
    """Quyền cụ thể cho từng task (ghi đè quyền module)"""
    
    user = models.ForeignKey(
        UserPermission, 
        on_delete=models.CASCADE, 
        related_name='task_permissions'
    )
    task = models.ForeignKey(
        TaskConfig, 
        on_delete=models.CASCADE, 
        related_name='user_permissions'
    )
    
    can_view = models.BooleanField("Được xem", default=True)
    can_edit = models.BooleanField("Được sửa", default=False)
    can_run = models.BooleanField("Được chạy", default=True)
    can_delete = models.BooleanField("Được xóa", default=False)
    
    class Meta:
        verbose_name = "Task Permission"
        verbose_name_plural = "Task Permissions"
        unique_together = ['user', 'task']
    
    def __str__(self):
        perms = []
        if self.can_view: perms.append('Xem')
        if self.can_edit: perms.append('Sửa')
        if self.can_run: perms.append('Chạy')
        if self.can_delete: perms.append('Xóa')
        return f"{self.user.username} - {self.task.name}: {', '.join(perms) or 'Không có quyền'}"