"""
SAP Auto Tasks - Django Settings
"""
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

# === SECRET_KEY phải giống chương trình mẹ ===
SECRET_KEY = 'django-insecure-hh4_o(87l9d3*48q$xehb0vyr)k(+5r%*ts1qb+#p1u32-xe9o'

DEBUG = True

ALLOWED_HOSTS = ['*']

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'tasks',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'tasks.middleware.ParentSessionMiddleware',      # Đọc session từ DB mẹ
    'tasks.middleware.AdminPermissionMiddleware',    # Kiểm tra quyền admin
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'sap_auto.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [BASE_DIR / 'templates'],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'sap_auto.wsgi.application'

# === DATABASES ===
DATABASES = {
    'default': {
        # Database riêng của chương trình con (SAP Auto)
        'ENGINE': 'django.db.backends.postgresql',
        'NAME': 'SAP_Auto',
        'USER': 'postgres',
        'PASSWORD': 'admin',
        'HOST': 'localhost',
        'PORT': '5432',
    },
    'parent_db': {
        # Database chương trình mẹ (SCM_NAVI) - chỉ để đọc session
        'ENGINE': 'django.db.backends.postgresql',
        'NAME': 'SCM_control',
        'USER': 'postgres',
        'PASSWORD': 'admin',
        'HOST': 'localhost',
        'PORT': '5432',
    },
}

LANGUAGE_CODE = 'vi'
TIME_ZONE = 'Asia/Ho_Chi_Minh'
USE_I18N = True
USE_TZ = True

STATIC_URL = '/static/'
STATICFILES_DIRS = [BASE_DIR / 'static']
STATIC_ROOT = BASE_DIR / 'staticfiles'

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# # === Session & Cookie Settings ===
# SESSION_COOKIE_NAME = 'sap_auto_sessionid'
# CSRF_COOKIE_NAME = 'sap_auto_csrftoken'
# SESSION_COOKIE_SAMESITE = 'Lax'
# CSRF_COOKIE_SAMESITE = 'Lax'
# SESSION_COOKIE_AGE = 3600  # 1 giờ
# SESSION_EXPIRE_AT_BROWSER_CLOSE = False
# SESSION_SAVE_EVERY_REQUEST = True
# SESSION_COOKIE_DOMAIN = None  # Không chia sẻ cookie
# CSRF_COOKIE_DOMAIN = None

# === SAP Auto Scanner Settings ===
SAP_SCANNER = {
    'SCAN_INTERVAL': 30,
    'FILE_WAIT_TIMEOUT': 30,
    'LOG_DIR': BASE_DIR / 'logs',
}

# ===== EMAIL CONFIGURATION =====
EMAIL_BACKEND = 'django.core.mail.backends.smtp.EmailBackend'
EMAIL_HOST = '157.8.1.154'
EMAIL_PORT = 25
EMAIL_USE_TLS = False
EMAIL_USE_SSL = False
EMAIL_HOST_USER = ''  # Không cần authentication
EMAIL_HOST_PASSWORD = ''  # Không cần authentication
DEFAULT_FROM_EMAIL = 'psnv.isg@vn.panasonic.com'