from django.contrib import admin
from .models import (
    TaskConfig, TaskLog, SAPUser,
    UserPermission, ModulePermission, TaskPermission,
    UserGroup, GroupModulePermission, GroupTaskPermission
)


# ===== Group Permissions Inline =====

class GroupModulePermissionInline(admin.TabularInline):
    model = GroupModulePermission
    extra = 1
    fields = ['module', 'can_view', 'can_edit', 'can_run', 'can_delete']
    verbose_name = "Module Permission"
    verbose_name_plural = "Module Permissions"


class GroupTaskPermissionInline(admin.TabularInline):
    model = GroupTaskPermission
    extra = 0
    fields = ['task', 'can_view', 'can_edit', 'can_run', 'can_delete']
    autocomplete_fields = ['task']
    verbose_name = "Task Permission"
    verbose_name_plural = "Task Permissions"


# ===== User Group Admin =====

@admin.register(UserGroup)
class UserGroupAdmin(admin.ModelAdmin):
    list_display = ['name', 'description', 'is_active', 'user_count', 'module_count', 'task_count', 'updated_at']
    list_filter = ['is_active']
    search_fields = ['name', 'description']
    list_editable = ['is_active']
    inlines = [GroupModulePermissionInline, GroupTaskPermissionInline]

    fieldsets = (
        ('Thông tin nhóm', {
            'fields': ('name', 'description', 'is_active')
        }),
    )

    def user_count(self, obj):
        count = obj.users.count()
        return count if count > 0 else '-'
    user_count.short_description = "Số user"

    def module_count(self, obj):
        count = obj.module_permissions.count()
        return count if count > 0 else '-'
    module_count.short_description = "Modules"

    def task_count(self, obj):
        count = obj.task_permissions.count()
        return count if count > 0 else '-'
    task_count.short_description = "Tasks"


# ===== User Permissions Inline =====

class ModulePermissionInline(admin.TabularInline):
    model = ModulePermission
    extra = 1
    fields = ['module', 'can_view', 'can_edit', 'can_run', 'can_delete']
    verbose_name = "Module Permission (User)"
    verbose_name_plural = "Module Permissions (User level - ghi đè Group)"


class TaskPermissionInline(admin.TabularInline):
    model = TaskPermission
    extra = 0
    fields = ['task', 'can_view', 'can_edit', 'can_run', 'can_delete']
    autocomplete_fields = ['task']
    verbose_name = "Task Permission (User)"
    verbose_name_plural = "Task Permissions (User level - ghi đè tất cả)"


# ===== User Permission Admin =====

@admin.register(UserPermission)
class UserPermissionAdmin(admin.ModelAdmin):
    list_display = ['username', 'display_name', 'is_admin', 'group_list', 'get_modules', 'has_sap_password', 'is_active']
    list_filter = ['is_admin', 'is_active', 'groups']
    search_fields = ['username', 'display_name']
    list_editable = ['is_admin', 'is_active']
    filter_horizontal = ['groups']
    inlines = [ModulePermissionInline, TaskPermissionInline]

    fieldsets = (
        ('Thông tin cơ bản', {
            'fields': ('username', 'display_name', 'is_active')
        }),
        ('Quyền', {
            'fields': ('is_admin', 'groups'),
            'description': 'Admin có toàn quyền. Nhóm cung cấp quyền mặc định, có thể ghi đè bằng permission bên dưới.'
        }),
        ('Mật khẩu SAP (chạy thủ công)', {
            'fields': ('sap_password',),
            'description': 'Mật khẩu SAP của user này khi chạy task thủ công. '
                           'Khác với mật khẩu SAPUser (dùng cho auto run). '
                           'User có thể tự lưu qua giao diện web.',
            'classes': ('collapse',),
        }),
    )

    def group_list(self, obj):
        groups = list(obj.groups.values_list('name', flat=True))
        return ', '.join(groups) if groups else '-'
    group_list.short_description = 'Nhóm'

    def get_modules(self, obj):
        if obj.is_admin:
            return '(All - Admin)'
        modules = obj.get_accessible_modules()
        return ', '.join(sorted(modules)) if modules else '-'
    get_modules.short_description = 'Modules'

    def has_sap_password(self, obj):
        return '✅' if obj.sap_password else '❌'
    has_sap_password.short_description = 'Có MK SAP'


# ===== Module Permission Admin =====

@admin.register(ModulePermission)
class ModulePermissionAdmin(admin.ModelAdmin):
    list_display = ['user', 'module', 'can_view', 'can_edit', 'can_run', 'can_delete']
    list_filter = ['module', 'can_view', 'can_edit', 'can_run', 'can_delete']
    search_fields = ['user__username', 'user__display_name']
    list_editable = ['can_view', 'can_edit', 'can_run', 'can_delete']


# ===== Task Permission Admin =====

@admin.register(TaskPermission)
class TaskPermissionAdmin(admin.ModelAdmin):
    list_display = ['user', 'task', 'can_view', 'can_edit', 'can_run', 'can_delete']
    list_filter = ['can_view', 'can_edit', 'can_run', 'can_delete', 'task__module']
    search_fields = ['user__username', 'task__name']
    list_editable = ['can_view', 'can_edit', 'can_run', 'can_delete']
    autocomplete_fields = ['task']


# ===== Group Module Permission Admin =====

@admin.register(GroupModulePermission)
class GroupModulePermissionAdmin(admin.ModelAdmin):
    list_display = ['group', 'module', 'can_view', 'can_edit', 'can_run', 'can_delete']
    list_filter = ['group', 'module', 'can_view', 'can_edit', 'can_run', 'can_delete']
    search_fields = ['group__name']
    list_editable = ['can_view', 'can_edit', 'can_run', 'can_delete']


# ===== Group Task Permission Admin =====

@admin.register(GroupTaskPermission)
class GroupTaskPermissionAdmin(admin.ModelAdmin):
    list_display = ['group', 'task', 'can_view', 'can_edit', 'can_run', 'can_delete']
    list_filter = ['group', 'can_view', 'can_edit', 'can_run', 'can_delete', 'task__module']
    search_fields = ['group__name', 'task__name']
    list_editable = ['can_view', 'can_edit', 'can_run', 'can_delete']
    autocomplete_fields = ['task']


# ===== Task Config Admin =====

@admin.register(TaskConfig)
class TaskConfigAdmin(admin.ModelAdmin):
    list_display = ['name', 'module', 'tcode', 'sap_user', 'status', 'auto_enabled', 'schedule_mode']
    list_filter = ['module', 'status', 'auto_enabled', 'schedule_mode', 'sap_user']
    search_fields = ['name', 'tcode', 'description']
    list_editable = ['status', 'auto_enabled']
    autocomplete_fields = ['sap_user']

    fieldsets = (
        ('Thông tin cơ bản', {
            'fields': ('module', 'name', 'tcode', 'description', 'sap_user')
        }),
        ('Đường dẫn', {
            'fields': ('watch_folder', 'folder_template', 'file_pattern', 'filename_template', 'file_regex')
        }),
        ('Handler', {
            'fields': ('handler_module', 'param1', 'param2')
        }),
        ('Trạng thái & Lịch', {
            'fields': ('status', 'auto_enabled', 'schedule_mode', 'scheduled_time', 'scheduled_days')
        }),
        ('Thông báo', {
            'fields': ('notify_on_error', 'notify_on_success', 'notify_emails'),
            'classes': ('collapse',)
        }),
    )


# ===== Task Log Admin =====

@admin.register(TaskLog)
class TaskLogAdmin(admin.ModelAdmin):
    list_display = ['task', 'filename', 'status', 'rows_processed', 'duration', 'executed_by', 'executed_at']
    list_filter = ['status', 'task__module', 'executed_at']
    search_fields = ['filename', 'task__name', 'message']
    readonly_fields = ['task', 'filename', 'filepath', 'status', 'message', 'rows_processed', 'duration', 'executed_by', 'executed_at']
    date_hierarchy = 'executed_at'


# ===== SAP User Admin =====

@admin.register(SAPUser)
class SAPUserAdmin(admin.ModelAdmin):
    list_display = ['client', 'username', 'updated_at']
    list_filter = ['client']
    search_fields = ['client', 'username']
    ordering = ['client', 'username']
