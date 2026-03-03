"""
Middleware đọc session từ database chương trình mẹ.

Logic:
1. Nếu có session mẹ → lấy username, auto-login, kiểm tra UserPermission
2. Nếu không có session mẹ → cho phép Django Admin login bình thường
"""
from django.db import connections
from django.http import HttpResponseForbidden


class ParentSessionMiddleware:
    """
    Đọc session cookie, ưu tiên database mẹ, fallback database con.
    """
    
    def __init__(self, get_response):
        self.get_response = get_response
    
    def __call__(self, request):
        request.parent_username = self.get_username(request)
        response = self.get_response(request)
        return response
    
    def get_username(self, request):
        """Lấy username từ session mẹ hoặc con"""
        try:
            session_key = request.COOKIES.get('sessionid', '')
            if not session_key:
                return None
            
            # 1. Thử database mẹ trước
            username = self.get_username_from_parent(session_key)
            if username:
                return username
            
            # 2. Fallback sang database con (khi login Admin con trực tiếp)
            username = self.get_username_from_child(session_key)
            if username:
                return username
            
            return None
                
        except Exception as e:
            print(f"[Session] Error: {e}")
            return None
    
    def get_username_from_parent(self, session_key):
        """Lấy username từ database mẹ"""
        try:
            from django.db import connections
            with connections['parent_db'].cursor() as cursor:
                cursor.execute(
                    "SELECT session_data FROM django_session WHERE session_key = %s AND expire_date > NOW()",
                    [session_key]
                )
                row = cursor.fetchone()
                if not row:
                    return None
                
                session_data = row[0]
                
                from django.contrib.sessions.backends.db import SessionStore
                store = SessionStore()
                data = store.decode(session_data)
                
                user_id = data.get('_auth_user_id')
                if not user_id:
                    return None
                
                cursor.execute("SELECT username FROM auth_user WHERE id = %s", [user_id])
                row = cursor.fetchone()
                username = row[0] if row else None
                
                if username:
                    print(f"[Session] Parent login: {username}")
                return username
        except Exception as e:
            print(f"[Session] Parent DB error: {e}")
            return None
    
    def get_username_from_child(self, session_key):
        """Lấy username từ database con (khi login Admin con)"""
        try:
            from django.contrib.sessions.models import Session
            from django.contrib.auth.models import User
            from django.utils import timezone
            
            sess = Session.objects.filter(
                session_key=session_key,
                expire_date__gt=timezone.now()
            ).first()
            
            if not sess:
                return None
            
            from django.contrib.sessions.backends.db import SessionStore
            store = SessionStore()
            data = store.decode(sess.session_data)
            
            user_id = data.get('_auth_user_id')
            if not user_id:
                return None
            
            user = User.objects.filter(id=user_id).first()
            if user:
                print(f"[Session] Child login: {user.username}")
                return user.username
            
            return None
        except Exception as e:
            print(f"[Session] Child DB error: {e}")
            return None


class AdminPermissionMiddleware:
    """
    Kiểm tra quyền admin:
    - Có session mẹ → kiểm tra UserPermission, auto-login
    - Không có session mẹ → cho phép Django Admin login bình thường
    """
    
    def __init__(self, get_response):
        self.get_response = get_response
    
    def __call__(self, request):
        if not request.path.startswith('/admin/'):
            return self.get_response(request)
        
        parent_username = getattr(request, 'parent_username', None)
        
        # === TRƯỜNG HỢP 1: Có session từ chương trình mẹ ===
        if parent_username:
            from tasks.models import UserPermission
            try:
                perm = UserPermission.objects.get(username=parent_username, is_active=True)
                
                if not perm.is_admin:
                    return HttpResponseForbidden(
                        f'<h1>403 - Không có quyền truy cập</h1>'
                        f'<p>User <b>{parent_username}</b> không có quyền Admin.</p>'
                        f'<p><a href="/">Quay lại Dashboard</a></p>'
                    )
                
                # Auto-login Django user
                if not hasattr(request, 'user') or not request.user.is_authenticated:
                    self.auto_login(request, parent_username, perm.is_admin)
                    
            except UserPermission.DoesNotExist:
                return HttpResponseForbidden(
                    f'<h1>403 - Không có quyền truy cập</h1>'
                    f'<p>User <b>{parent_username}</b> chưa được phân quyền.</p>'
                    f'<p><a href="/">Quay lại Dashboard</a></p>'
                )
            
            return self.get_response(request)
        
        # === TRƯỜNG HỢP 2: Không có session mẹ → Django Admin bình thường ===
        # Cho phép login/logout và mọi thao tác Django Admin như bình thường
        return self.get_response(request)
    
    def auto_login(self, request, username, is_superuser=False):
        """Tự động login Django user"""
        try:
            from django.contrib.auth.models import User
            from django.contrib.auth import login
            
            user, created = User.objects.get_or_create(
                username=username,
                defaults={
                    'is_staff': True,
                    'is_superuser': is_superuser,
                }
            )
            
            if not user.is_staff:
                user.is_staff = True
                user.save()
            
            login(request, user, backend='django.contrib.auth.backends.ModelBackend')
            print(f"[Admin] Auto-login: {username}")
            
        except Exception as e:
            print(f"[Admin] Auto-login failed: {e}")
