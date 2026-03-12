"""
Email Notifications for SAP Auto Tasks
======================================
Gửi email thông báo khi task hoàn thành hoặc lỗi.
Sử dụng SMTP server nội bộ (không cần authentication).
"""
import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from django.conf import settings
from django.utils import timezone

log = logging.getLogger(__name__)

# SMTP Configuration
SMTP_HOST = getattr(settings, 'EMAIL_HOST', '157.8.1.154')
SMTP_PORT = getattr(settings, 'EMAIL_PORT', 25)
DEFAULT_FROM = getattr(settings, 'DEFAULT_FROM_EMAIL', 'psnv.isg@vn.panasonic.com')


def send_task_notification(task, log_entry, status_type='success'):
    """
    Gửi email thông báo kết quả task
    
    Args:
        task: TaskConfig object
        log_entry: TaskLog object
        status_type: 'success' hoặc 'error'
    """
    # Kiểm tra có email recipients không
    recipients = get_recipients(task, status_type)
    if not recipients:
        log.info(f"[EMAIL] No recipients for {task.name} ({status_type})")
        return False
    
    try:
        # Tạo nội dung email
        subject, body_html, body_text = build_email_content(task, log_entry, status_type)
        
        # Gửi email
        result = send_email(
            to_emails=recipients,
            subject=subject,
            body_html=body_html,
            body_text=body_text
        )
        
        if result:
            log.info(f"[EMAIL] Sent {status_type} notification for {task.name} to {recipients}")
        else:
            log.warning(f"[EMAIL] Failed to send notification for {task.name}")
        
        return result
        
    except Exception as e:
        log.error(f"[EMAIL] Error sending notification: {e}")
        return False


def get_recipients(task, status_type):
    """
    Lấy danh sách email recipients từ task config
    
    Args:
        task: TaskConfig object
        status_type: 'success' hoặc 'error'
    
    Returns:
        list of email addresses
    """
    recipients = []
    
    # Lấy từ field notify_emails của task (mỗi email 1 dòng)
    if hasattr(task, 'notify_emails') and task.notify_emails:
        # Split theo dòng mới hoặc dấu phẩy
        for line in task.notify_emails.strip().split('\n'):
            for email in line.split(','):
                email = email.strip()
                if email and '@' in email:
                    recipients.append(email)
    
    return list(set(recipients))  # Remove duplicates


def build_email_content(task, log_entry, status_type):
    """
    Tạo nội dung email (subject, body HTML, body text)
    
    Returns:
        tuple: (subject, body_html, body_text)
    """
    # Thông tin cơ bản
    task_name = task.name
    tcode = task.tcode
    filename = log_entry.filename if log_entry else 'N/A'
    message = log_entry.message if log_entry else 'N/A'
    rows = log_entry.rows_processed if log_entry else 0
    duration = log_entry.duration if log_entry else 0
    executed_at = log_entry.executed_at if log_entry else timezone.now()
    
    # Format thời gian
    time_str = timezone.localtime(executed_at).strftime('%d/%m/%Y %H:%M:%S')
    
    if status_type == 'success':
        subject = f"[SAP Auto] ✓ {task_name} - Thành công"
        status_text = "THÀNH CÔNG"
        status_color = "#28a745"  # Green
        icon = "✓"
    else:
        subject = f"[SAP Auto] ✗ {task_name} - Lỗi"
        status_text = "LỖI"
        status_color = "#dc3545"  # Red
        icon = "✗"
    
    # HTML Body
    body_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: Arial, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }}
            .container {{ max-width: 600px; margin: 0 auto; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
            .header {{ background: {status_color}; color: white; padding: 20px; text-align: center; }}
            .header h1 {{ margin: 0; font-size: 24px; }}
            .content {{ padding: 20px; }}
            .info-table {{ width: 100%; border-collapse: collapse; }}
            .info-table tr {{ border-bottom: 1px solid #eee; }}
            .info-table td {{ padding: 12px 8px; }}
            .info-table td:first-child {{ font-weight: bold; color: #666; width: 140px; }}
            .message-box {{ background: #f8f9fa; border-left: 4px solid {status_color}; padding: 15px; margin-top: 15px; }}
            .footer {{ background: #f8f9fa; padding: 15px; text-align: center; font-size: 12px; color: #666; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>{icon} {status_text}</h1>
            </div>
            <div class="content">
                <table class="info-table">
                    <tr>
                        <td>Task:</td>
                        <td><strong>{task_name}</strong></td>
                    </tr>
                    <tr>
                        <td>TCode:</td>
                        <td>{tcode}</td>
                    </tr>
                    <tr>
                        <td>File:</td>
                        <td>{filename}</td>
                    </tr>
                    <tr>
                        <td>Rows processed:</td>
                        <td>{rows}</td>
                    </tr>
                    <tr>
                        <td>Duration:</td>
                        <td>{duration} seconds</td>
                    </tr>
                    <tr>
                        <td>Time:</td>
                        <td>{time_str}</td>
                    </tr>
                </table>
                
                <div class="message-box">
                    <strong>Message:</strong><br>
                    {message}
                </div>
            </div>
            <div class="footer">
                SAP Auto Tasks - Automated Notification<br>
                Do not reply to this email.
            </div>
        </div>
    </body>
    </html>
    """
    
    # Plain text Body (fallback)
    body_text = f"""
{status_text}: {task_name}
{'=' * 50}

Task:           {task_name}
TCode:          {tcode}
File:           {filename}
Rows processed: {rows}
Duration:       {duration} seconds
Time:           {time_str}

Message:
{message}

---
SAP Auto Tasks - Automated Notification
    """
    
    return subject, body_html, body_text


def send_email(to_emails, subject, body_html, body_text=None, from_email=None):
    """
    Gửi email qua SMTP server nội bộ
    
    Args:
        to_emails: list of email addresses hoặc string (comma-separated)
        subject: Email subject
        body_html: HTML body
        body_text: Plain text body (optional)
        from_email: From address (optional, uses DEFAULT_FROM)
    
    Returns:
        bool: True if sent successfully
    """
    # Normalize recipients
    if isinstance(to_emails, str):
        to_emails = [e.strip() for e in to_emails.split(',') if e.strip()]
    
    if not to_emails:
        log.warning("[EMAIL] No recipients provided")
        return False
    
    # From address
    if not from_email:
        from_email = DEFAULT_FROM
    
    try:
        # Tạo message
        msg = MIMEMultipart('alternative')
        msg['Subject'] = Header(subject, 'utf-8')
        msg['From'] = from_email
        msg['To'] = ', '.join(to_emails)
        
        # Attach plain text
        if body_text:
            part_text = MIMEText(body_text, 'plain', 'utf-8')
            msg.attach(part_text)
        
        # Attach HTML
        part_html = MIMEText(body_html, 'html', 'utf-8')
        msg.attach(part_html)
        
        # Gửi qua SMTP
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.sendmail(from_email, to_emails, msg.as_string())
        
        log.info(f"[EMAIL] Sent to {to_emails}")
        return True
        
    except smtplib.SMTPException as e:
        log.error(f"[EMAIL] SMTP error: {e}")
        return False
    except Exception as e:
        log.error(f"[EMAIL] Error: {e}")
        return False


def send_test_email(to_email):
    """
    Gửi email test để kiểm tra cấu hình SMTP
    
    Args:
        to_email: Email address to send test
    
    Returns:
        bool: True if sent successfully
    """
    subject = "[SAP Auto] Test Email"
    
    body_html = """
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"></head>
    <body style="font-family: Arial, sans-serif; padding: 20px;">
        <h2 style="color: #28a745;">✓ Email Test Successful!</h2>
        <p>This is a test email from SAP Auto Tasks.</p>
        <p>If you received this email, the SMTP configuration is working correctly.</p>
        <hr>
        <p style="color: #666; font-size: 12px;">
            SMTP Server: {host}:{port}<br>
            Time: {time}
        </p>
    </body>
    </html>
    """.format(
        host=SMTP_HOST,
        port=SMTP_PORT,
        time=timezone.localtime(timezone.now()).strftime('%d/%m/%Y %H:%M:%S')
    )
    
    body_text = f"""
Email Test Successful!

This is a test email from SAP Auto Tasks.
If you received this email, the SMTP configuration is working correctly.

SMTP Server: {SMTP_HOST}:{SMTP_PORT}
Time: {timezone.localtime(timezone.now()).strftime('%d/%m/%Y %H:%M:%S')}
    """
    
    return send_email(to_email, subject, body_html, body_text)