"""
Django Management Command: scan_tasks
======================================
Background scanner for watching folders and triggering SAP tasks automatically.
Supports 2 modes:
- interval: scan continuously (30s, 5 min, ...)
- daily/weekly: run at scheduled time (8:00 AM, ...)

Usage:
    python manage.py scan_tasks              # Run continuously
    python manage.py scan_tasks --once       # Run once then stop
"""
import os
import re
import sys
import io
import time
import importlib
import logging
from datetime import datetime, timedelta, date

from django.core.management.base import BaseCommand
from django.utils import timezone

from tasks.models import TaskConfig, TaskLog
from tasks.notifications import send_task_notification

# Fix Unicode output for Windows console
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

log = logging.getLogger(__name__)


class Command(BaseCommand):
    help = 'Scan folders and auto-run SAP tasks when new files are detected'

    def add_arguments(self, parser):
        parser.add_argument(
            '--once',
            action='store_true',
            help='Scan once then stop',
        )

    def handle(self, *args, **options):
        run_once = options['once']

        self.stdout.write("=" * 60)
        self.stdout.write(self.style.SUCCESS(" SAP Auto Scanner - Started"))
        self.stdout.write("=" * 60)

        # Show active tasks
        active_tasks = TaskConfig.objects.filter(auto_enabled=True, status='active')
        for task in active_tasks:
            schedule_info = self._get_schedule_info(task)
            self.stdout.write(f"  [>] [{task.tcode}] {task.name}")
            self.stdout.write(f"      Schedule: {schedule_info}")
            self.stdout.write(f"      Path:     {task.resolve_folder()}")
        
        if not active_tasks.exists():
            self.stdout.write(self.style.WARNING("  No active tasks found!"))
            if run_once:
                return
            # Continue waiting even if no tasks

        self.stdout.write(f"\n  Waiting for new files... (Ctrl+C to stop)\n")

        # Dict to store last scan time for each task (for interval mode)
        last_scan = {}

        try:
            while True:
                now = datetime.now()
                today = now.date()
                current_time = now.time()
                current_weekday = now.weekday()  # 0=Monday, 6=Sunday
                
                # Reload tasks from DB
                tasks = TaskConfig.objects.filter(auto_enabled=True, status='active')
                
                for task in tasks:
                    should_run = False
                    
                    if task.schedule_mode == 'interval':
                        # Interval mode: scan by cycle
                        should_run = self._check_interval(task, last_scan)
                        
                    elif task.schedule_mode == 'daily':
                        # Daily mode: run at scheduled time every day
                        should_run = self._check_daily(task, today, current_time)
                        
                    elif task.schedule_mode == 'weekly':
                        # Weekly mode: run at scheduled time on specific days
                        should_run = self._check_weekly(task, today, current_time, current_weekday)
                    
                    if should_run:
                        try:
                            self._scan_task(task)
                        except Exception as e:
                            log.error(f"[{task.tcode}] Scanner error: {e}")
                        
                        # Update last scan time (for interval mode)
                        last_scan[task.id] = time.time()
                
                if run_once:
                    break
                
                # Sleep 10 seconds then check again
                time.sleep(10)
                
        except KeyboardInterrupt:
            self.stdout.write(self.style.SUCCESS("\n  Scanner stopped."))

    def _check_interval(self, task, last_scan):
        """Check if interval task should run"""
        now = time.time()
        interval = getattr(task, 'scan_interval', 30) or 30
        
        if task.id in last_scan:
            elapsed = now - last_scan[task.id]
            if elapsed < interval:
                return False
        return True

    def _check_daily(self, task, today, current_time):
        """Check if daily task should run"""
        if not task.scheduled_time:
            return False
        
        # Already ran today?
        if task.last_scheduled_run == today:
            return False
        
        # Time to run?
        scheduled = task.scheduled_time
        # Allow 5 minute window
        now_minutes = current_time.hour * 60 + current_time.minute
        scheduled_minutes = scheduled.hour * 60 + scheduled.minute
        
        if now_minutes >= scheduled_minutes and now_minutes <= scheduled_minutes + 5:
            self.stdout.write(f"  [TIME] [{task.tcode}] Scheduled time: {scheduled.strftime('%H:%M')}")
            # Update immediately to prevent re-run
            task.last_scheduled_run = today
            task.save(update_fields=['last_scheduled_run'])
            return True
        
        return False

    def _check_weekly(self, task, today, current_time, current_weekday):
        """Check if weekly task should run"""
        if not task.scheduled_time:
            return False
        
        # Already ran today?
        if task.last_scheduled_run == today:
            return False
        
        # Is today in scheduled days list?
        scheduled_days = [int(d.strip()) for d in task.scheduled_days.split(',') if d.strip().isdigit()]
        
        # Convert: user format 1=Mon, 2=Tue, ... 5=Fri, 6=Sat, 0=Sun
        # Python weekday: 0=Monday, 1=Tuesday, ... 6=Sunday
        python_to_user = {0: 1, 1: 2, 2: 3, 3: 4, 4: 5, 5: 6, 6: 0}
        user_weekday = python_to_user[current_weekday]
        
        if user_weekday not in scheduled_days:
            return False
        
        # Time to run?
        scheduled = task.scheduled_time
        now_minutes = current_time.hour * 60 + current_time.minute
        scheduled_minutes = scheduled.hour * 60 + scheduled.minute
        
        if now_minutes >= scheduled_minutes and now_minutes <= scheduled_minutes + 5:
            self.stdout.write(f"  [TIME] [{task.tcode}] Scheduled time: {scheduled.strftime('%H:%M')}")
            # Update immediately to prevent re-run
            task.last_scheduled_run = today
            task.save(update_fields=['last_scheduled_run'])
            return True
        
        return False

    def _get_schedule_info(self, task):
        """Return readable schedule info"""
        if task.schedule_mode == 'interval':
            interval = getattr(task, 'scan_interval', 30) or 30
            return f"Every {self._format_interval(interval)}"
        elif task.schedule_mode == 'daily':
            time_str = task.scheduled_time.strftime('%H:%M') if task.scheduled_time else 'N/A'
            return f"Daily at {time_str}"
        elif task.schedule_mode == 'weekly':
            time_str = task.scheduled_time.strftime('%H:%M') if task.scheduled_time else 'N/A'
            days_map = {1: 'Mon', 2: 'Tue', 3: 'Wed', 4: 'Thu', 5: 'Fri', 6: 'Sat', 0: 'Sun'}
            days = task.scheduled_days or ''
            day_names = [days_map.get(int(d.strip()), d) for d in days.split(',') if d.strip().isdigit()]
            return f"{', '.join(day_names)} at {time_str}"
        return "N/A"

    def _format_interval(self, seconds):
        """Format interval to readable string"""
        if not seconds:
            seconds = 30
        if seconds < 60:
            return f"{seconds} seconds"
        elif seconds < 3600:
            minutes = seconds // 60
            return f"{minutes} minutes"
        elif seconds < 86400:
            hours = seconds // 3600
            return f"{hours} hours"
        else:
            days = seconds // 86400
            return f"{days} days"

    def _scan_task(self, task):
        """Scan a task - find new files in folder (supports dynamic path)"""
        # Use resolve_folder() to auto-calculate path by time
        folder = task.resolve_folder()

        if not folder or not os.path.exists(folder):
            return

        # Get list of processed files
        processed_files = set(
            task.logs.values_list('filepath', flat=True)
        )

        # Scan folder
        for filename in os.listdir(folder):
            filepath = os.path.join(folder, filename)

            # Only process files
            if not os.path.isfile(filepath):
                continue

            # Already processed? (DISABLED FOR TESTING)
            if filepath in processed_files:
                continue

            # Match regex pattern?
            if task.file_regex and not re.match(task.file_regex, filename):
                continue

            # Wait for file to finish copying
            if not self._wait_file_ready(filepath):
                log.warning(f"[{task.tcode}] File not ready: {filename}")
                continue

            # === Run handler ===
            self.stdout.write(
                self.style.HTTP_INFO(f"  [{task.tcode}] Detected: {filename}")
            )
            self._execute_task(task, filepath, filename)

    def _execute_task(self, task, filepath, filename):
        """Call handler function and save log"""
        start_time = time.time()

        try:
            # Import handler dynamically
            handler_func = self._import_handler(task.handler_module)

            # Call handler (no SAP session here,
            # session will be managed inside handler)
            result = handler_func(filepath, session=None, task=task)

            duration = round(time.time() - start_time, 2)

            # Save log
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
                    self.style.SUCCESS(f"    [OK] {result.get('message', 'OK')} ({duration}s)")
                )
                # Send success notification email (if enabled)
                log_entry = task.logs.order_by('-executed_at').first()
                if log_entry and task.notify_on_success:
                    send_task_notification(task, log_entry, 'success')
            elif status_display == 'error':
                self.stdout.write(
                    self.style.ERROR(f"    [ERROR] {result.get('message', 'Error')} ({duration}s)")
                )
                # Send error notification email
                log_entry = task.logs.order_by('-executed_at').first()
                if log_entry and task.notify_on_error:
                    send_task_notification(task, log_entry, 'error')
            else:
                self.stdout.write(f"    [SKIP] {result.get('message', '')} ({duration}s)")

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
                self.style.ERROR(f"    [ERROR] Exception: {e} ({duration}s)")
            )
            # Send error notification email
            if task.notify_on_error:
                send_task_notification(task, log_entry, 'error')

    def _import_handler(self, handler_path):
        """Import handler function from string path"""
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
        """Wait for file to finish copying (size stable)"""
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