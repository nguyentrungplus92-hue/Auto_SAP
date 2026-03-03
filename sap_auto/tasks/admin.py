from django.contrib import admin
from .models import TaskConfig, TaskLog, UserPermission, ModulePermission, TaskPermission, SAPUser


# ===== Inline Permissions =====

class ModulePermissionInline(admin.TabularInline):
    model = ModulePermission
    extra = 1
    fields = ['module', 'can_view', 'can_edit', 'can_run', 'can_delete']


class TaskPermissionInline(admin.TabularInline):
    model = TaskPermission
    extra = 0
    fields = ['task', 'can_view', 'can_edit', 'can_run', 'can_delete']
    autocomplete_fields = ['task']


# ===== User Permission Admin =====

@admin.register(UserPermission)
class UserPermissionAdmin(admin.ModelAdmin):
    list_display = ['username', 'display_name', 'is_admin', 'get_modules', 'is_active']
    list_filter = ['is_admin', 'is_active']
    search_fields = ['username', 'display_name']
    list_editable = ['is_admin', 'is_active']
    inlines = [ModulePermissionInline, TaskPermissionInline]
    
    def get_modules(self, obj):
        modules = obj.get_accessible_modules()
        return ', '.join(modules) if modules else '-'
    get_modules.short_description = 'Modules'


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
