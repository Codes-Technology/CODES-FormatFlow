import os
import uuid
from dotenv import load_dotenv
from urllib.parse import quote_plus  

load_dotenv()

# Database Configuration
DB_USER = os.getenv('DB_USER', 'root')
DB_PASSWORD = os.getenv('DB_PASSWORD', '')
DB_HOST = os.getenv('DB_HOST', 'localhost')
DB_DATABASE = os.getenv('DB_DATABASE', 'prisma_db')
APP_INSTANCE_ID = 'supersecretkey12345678901234567890abcddfsfsd'
# URL-encode password to handle special characters (@, #, !, etc.)
DB_PASSWORD_ENCODED = quote_plus(DB_PASSWORD) if DB_PASSWORD else ''

# Build SQLAlchemy Database URI with encoded password
SQLALCHEMY_DATABASE_URI = (
    f"mysql+pymysql://{DB_USER}:{DB_PASSWORD_ENCODED}@{DB_HOST}/{DB_DATABASE}"
)

SQLALCHEMY_TRACK_MODIFICATIONS = False

#(Adobe, folders, etc.)
ADOBE_CLIENT_ID = os.getenv('ADOBE_CLIENT_ID')
ADOBE_CLIENT_SECRET = os.getenv('ADOBE_CLIENT_SECRET')
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ALLOWED_EXTENSIONS = {'pdf', 'txt', 'docx'}
MAX_FILE_SIZE = 16 * 1024 * 1024
TEMPLATE_DOCX = os.path.join(BASE_DIR, 'Letter_pad.docx')
# Authentication
SECRET_KEY = os.getenv('SECRET_KEY', 'a9f3c2e1d4b7a0f8e3c6d8b2a5f1e7c4d9b3a6f2e8c1d5b4a7f0e2c3d6b8a1f4we')

