from django.urls import path
from . import views

app_name = 'tasks'

urlpatterns = [
    # Pages
    path('', views.dashboard, name='dashboard'),
    path('task/<int:pk>/', views.task_detail, name='detail'),

    # API
    path('api/tasks/', views.api_tasks, name='api_tasks'),
    path('api/tasks/create/', views.api_task_create, name='api_task_create'),
    path('api/tasks/<int:pk>/update/', views.api_task_update, name='api_task_update'),
    path('api/tasks/<int:pk>/logs/', views.api_task_logs, name='api_task_logs'),
    path('api/tasks/<int:pk>/toggle/', views.api_task_toggle, name='api_task_toggle'),
    path('api/tasks/<int:pk>/run/', views.api_task_run, name='api_task_run'),
    path('api/tasks/<int:pk>/delete/', views.api_task_delete, name='api_task_delete'),
    path('api/scanner/status/', views.api_scan_status, name='api_scan_status'),

    # SAP Password
    path('api/sap-password/', views.api_get_sap_password, name='api_get_sap_password'),
    path('api/sap-password/save/', views.api_save_sap_password, name='api_save_sap_password'),

    # Check file
    path('api/tasks/<int:pk>/check-file/', views.api_task_check_file),

]
