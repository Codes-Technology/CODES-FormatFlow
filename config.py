import os
import secrets

from dotenv import load_dotenv
from urllib.parse import quote_plus

load_dotenv()


class Config:
    # New dev database connection
    DB_HOST = os.getenv('DB_HOST', 'localhost')
    DB_USER = os.getenv('DB_USER', 'root')
    DB_PASS = os.getenv('DB_PASS', '')
    DB_NAME = os.getenv('DB_NAME', 'formatflow_dev')
    BASE_UPLOAD_URL = os.getenv('BASE_UPLOAD_URL', 'http://localhost:5000')

    _DB_PASS_ENCODED = quote_plus(DB_PASS) if DB_PASS else ''
    if _DB_PASS_ENCODED:
        SQLALCHEMY_DATABASE_URI = (
            f"mysql+pymysql://{DB_USER}:{_DB_PASS_ENCODED}@{DB_HOST}/{DB_NAME}"
        )
    else:
        SQLALCHEMY_DATABASE_URI = f"mysql+pymysql://{DB_USER}@{DB_HOST}/{DB_NAME}"

    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ECHO = False

    SECRET_KEY = os.getenv('SECRET_KEY', 'dev-secret-key-2024')
    JWT_SECRET_KEY = os.getenv('JWT_SECRET_KEY', 'dev-jwt-key-2024')
    JWT_TOKEN_LOCATION = ['cookies']
    JWT_COOKIE_SECURE = False
    JWT_ACCESS_TOKEN_EXPIRES = 86400

    APP_INSTANCE_ID = os.getenv('APP_INSTANCE_ID', secrets.token_urlsafe(32))

    ADOBE_CLIENT_ID = os.getenv('ADOBE_CLIENT_ID')
    ADOBE_CLIENT_SECRET = os.getenv('ADOBE_CLIENT_SECRET')

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    TEMPLATE_DOCX = os.path.join(BASE_DIR, 'Letter_pad.docx')
    STORAGE_DIR = os.path.join(BASE_DIR, 'storage')
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'file_storage')
    ALLOWED_EXTENSIONS = {'pdf', 'txt', 'docx'}
    MAX_FILE_SIZE = 16 * 1024 * 1024


DB_HOST = Config.DB_HOST
DB_USER = Config.DB_USER
DB_PASS = Config.DB_PASS
DB_NAME = Config.DB_NAME
BASE_UPLOAD_URL = Config.BASE_UPLOAD_URL
SQLALCHEMY_DATABASE_URI = Config.SQLALCHEMY_DATABASE_URI
SQLALCHEMY_TRACK_MODIFICATIONS = Config.SQLALCHEMY_TRACK_MODIFICATIONS
SQLALCHEMY_ECHO = Config.SQLALCHEMY_ECHO
SECRET_KEY = Config.SECRET_KEY
JWT_SECRET_KEY = Config.JWT_SECRET_KEY
JWT_TOKEN_LOCATION = Config.JWT_TOKEN_LOCATION
JWT_COOKIE_SECURE = Config.JWT_COOKIE_SECURE
JWT_ACCESS_TOKEN_EXPIRES = Config.JWT_ACCESS_TOKEN_EXPIRES
APP_INSTANCE_ID = Config.APP_INSTANCE_ID
ADOBE_CLIENT_ID = Config.ADOBE_CLIENT_ID
ADOBE_CLIENT_SECRET = Config.ADOBE_CLIENT_SECRET
BASE_DIR = Config.BASE_DIR
TEMPLATE_DOCX = Config.TEMPLATE_DOCX
STORAGE_DIR = Config.STORAGE_DIR
UPLOAD_FOLDER = Config.UPLOAD_FOLDER
ALLOWED_EXTENSIONS = Config.ALLOWED_EXTENSIONS
MAX_FILE_SIZE = Config.MAX_FILE_SIZE
