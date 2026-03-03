"""
Email Notification Utility
==========================
Gửi email thông báo khi task hoàn thành hoặc gặp lỗi.
"""
import logging
from django.core.mail import send_mail, EmailMessage
from django.template.loader import render_to_string
from django.conf import settings

log = logging.getLogger(__name__)


def send_task_notification(task, log_entry, notify_type='error'):
    """
    Gửi email thông báo cho task.
    
    Args:
        task: TaskConfig instance
        log_entry: TaskLog instance
        notify_type: 'error' hoặc 'success'
    """
    # Kiểm tra có cần gửi không
    if notify_type == 'error' and not task.notify_on_error:
        return
    if notify_type == 'success' and not task.notify_on_success:
        return
    
    recipients = task.email_list
    if not recipients:
        log.warning(f"[{task.tcode}] Không có email để gửi thông báo")
        return
    
    # Tạo nội dung email
    if notify_type == 'error':
        subject = f"[LỖI] SAP Task: {task.name} ({task.tcode})"
        status_text = "GẶP LỖI"
        status_color = "#dc3545"
    else:
        subject = f"[OK] SAP Task: {task.name} ({task.tcode})"
        status_text = "THÀNH CÔNG"
        status_color = "#28a745"
    
    # Nội dung HTML
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
            .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
            .header {{ background: {status_color}; color: white; padding: 15px 20px; border-radius: 8px 8px 0 0; }}
            .header h1 {{ margin: 0; font-size: 18px; }}
            .content {{ background: #f8f9fa; padding: 20px; border: 1px solid #dee2e6; border-top: none; border-radius: 0 0 8px 8px; }}
            .info-row {{ margin-bottom: 12px; }}
            .label {{ font-weight: bold; color: #666; display: inline-block; width: 120px; }}
            .value {{ color: #333; }}
            .error-box {{ background: #fff3f3; border: 1px solid #ffcdd2; padding: 12px; border-radius: 4px; margin-top: 15px; }}
            .error-box pre {{ margin: 0; white-space: pre-wrap; word-wrap: break-word; font-size: 13px; color: #c62828; }}
            .footer {{ margin-top: 20px; font-size: 12px; color: #999; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>SAP Auto Task - {status_text}</h1>
            </div>
            <div class="content">
                <div class="info-row">
                    <span class="label">Task:</span>
                    <span class="value">{task.name}</span>
                </div>
                <div class="info-row">
                    <span class="label">T-Code:</span>
                    <span class="value">{task.tcode or 'N/A'}</span>
                </div>
                <div class="info-row">
                    <span class="label">File:</span>
                    <span class="value">{log_entry.filename}</span>
                </div>
                <div class="info-row">
                    <span class="label">Đường dẫn:</span>
                    <span class="value">{log_entry.filepath}</span>
                </div>
                <div class="info-row">
                    <span class="label">Thời gian:</span>
                    <span class="value">{log_entry.executed_at.strftime('%d/%m/%Y %H:%M:%S')}</span>
                </div>
                <div class="info-row">
                    <span class="label">Số dòng xử lý:</span>
                    <span class="value">{log_entry.rows_processed}</span>
                </div>
                <div class="info-row">
                    <span class="label">Thời gian chạy:</span>
                    <span class="value">{log_entry.duration}s</span>
                </div>
                
                {"<div class='error-box'><strong>Chi tiết lỗi:</strong><pre>" + log_entry.message + "</pre></div>" if notify_type == 'error' and log_entry.message else ""}
                
                <div class="footer">
                    <p>Email này được gửi tự động từ hệ thống SAP Auto Tasks.</p>
                    <p>Vui lòng không reply email này.</p>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    
    # Nội dung text (fallback)
    text_content = f"""
SAP Auto Task - {status_text}
==============================

Task: {task.name}
T-Code: {task.tcode or 'N/A'}
File: {log_entry.filename}
Đường dẫn: {log_entry.filepath}
Thời gian: {log_entry.executed_at.strftime('%d/%m/%Y %H:%M:%S')}
Số dòng xử lý: {log_entry.rows_processed}
Thời gian chạy: {log_entry.duration}s

{"Chi tiết lỗi: " + log_entry.message if notify_type == 'error' and log_entry.message else ""}

---
Email này được gửi tự động từ hệ thống SAP Auto Tasks.
    """
    
    try:
        email = EmailMessage(
            subject=subject,
            body=html_content,
            from_email=settings.DEFAULT_FROM_EMAIL,
            to=recipients,
        )
        email.content_subtype = 'html'
        email.send(fail_silently=False)
        
        log.info(f"[{task.tcode}] Đã gửi email thông báo tới: {', '.join(recipients)}")
        return True
        
    except Exception as e:
        log.error(f"[{task.tcode}] Lỗi gửi email: {e}")
        return False


def send_test_email(to_email):
    """
    Gửi email test để kiểm tra cấu hình SMTP.
    
    Usage trong shell:
        from tasks.notifications import send_test_email
        send_test_email('your-email@company.com')
    """
    try:
        send_mail(
            subject='[TEST] SAP Auto Tasks - Email Test',
            message='Nếu bạn nhận được email này, cấu hình SMTP đã hoạt động!',
            from_email=settings.DEFAULT_FROM_EMAIL,
            recipient_list=[to_email],
            fail_silently=False,
        )
        print(f"✅ Đã gửi email test tới: {to_email}")
        return True
    except Exception as e:
        print(f"❌ Lỗi gửi email: {e}")
        return False
