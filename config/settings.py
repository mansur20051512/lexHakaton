import os
from pathlib import Path
from dotenv import load_dotenv

# BASE_DIR yo‘li
BASE_DIR = Path(__file__).resolve().parent.parent

# .env faylini yuklash
load_dotenv(BASE_DIR / '.env')

# Xavfsizlik sozlamalari
SECRET_KEY = os.getenv('SECRET_KEY', 'django-insecure-adolat-ai-lexhakaton-uzbekistan-legaltech-2026-super-secure')
DEBUG = os.getenv('DEBUG', 'True').lower() in ('true', '1', 'yes')

ALLOWED_HOSTS = ['*']

# O‘rnatilgan ilovalar
INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    
    # Bizning huquqiy AI ilovamiz
    'legal.apps.LegalConfig',
    
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'config.urls'

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

WSGI_APPLICATION = 'config.wsgi.application'

# Ma'lumotlar bazasi (Xakaton uchun tezkor va mustaqil SQLite3)
DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}

# Parol tekshiruvlari
AUTH_PASSWORD_VALIDATORS = [
    {
        'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator',
    },
]

# Xalqaro sozlamalar va O‘zbekiston vaqti
LANGUAGE_CODE = 'uz'
TIME_ZONE = 'Asia/Tashkent'
USE_I18N = True
USE_TZ = True

# Statik va Media fayllar
STATIC_URL = '/static/'
STATICFILES_DIRS = [BASE_DIR / 'static']
STATIC_ROOT = BASE_DIR / 'staticfiles'

MEDIA_URL = '/media/'
MEDIA_ROOT = BASE_DIR / 'media'

# Default primary key field type
DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# Open-Source & VoiceLab LLM sozlamalari
VOICE_MODE = os.getenv('VOICE_MODE', 'voicelab')
VOICELAB_API_KEY = os.getenv('VOICELAB_API_KEY', '')
VOICELAB_BASE_URL = os.getenv('VOICELAB_BASE_URL', 'https://api.voicelab.uz')
VOICELAB_LLM_MODEL = os.getenv('VOICELAB_LLM_MODEL', 'aisha-orbit')

LLM_API_URL = os.getenv('LLM_API_URL', f"{VOICELAB_BASE_URL}/v1/chat/completions")
LLM_API_KEY = os.getenv('LLM_API_KEY', VOICELAB_API_KEY)
LLM_MODEL_NAME = os.getenv('LLM_MODEL_NAME', VOICELAB_LLM_MODEL)
LLM_TIMEOUT = int(os.getenv('LLM_TIMEOUT', '60'))
