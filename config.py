import os
from dotenv import load_dotenv

load_dotenv() # Loads the variables from the .env file

# Adobe API Credentials
ADOBE_CLIENT_ID = os.getenv('ADOBE_CLIENT_ID')
ADOBE_CLIENT_SECRET = os.getenv('ADOBE_CLIENT_SECRET')

# Base directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Upload and output folders
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'output')

# Allowed file extensions
ALLOWED_EXTENSIONS = {'pdf', 'txt', 'docx'}

# Max file size (16 MB)
MAX_FILE_SIZE = 16 * 1024 * 1024

# Template document path
TEMPLATE_DOCX = os.path.join(BASE_DIR, 'Letter_pad.docx')

# Flask secret key (change in production)
SECRET_KEY = 'dev-secret-key-change-in-production'

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
