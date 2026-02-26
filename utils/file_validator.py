import os
from typing import IO
from werkzeug.utils import secure_filename
from config import ALLOWED_EXTENSIONS, MAX_FILE_SIZE

def allowed_file(filename: str) -> bool:
    """Check if the file extension exists and is in the allowed list."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def validate_file_size(file_stream: IO) -> bool:
    """
    Measure file stream size without loading it into memory.
    Returns True if size is <= MAX_FILE_SIZE.
    """
    # Fast-forward to the end to get the byte count
    file_stream.seek(0, os.SEEK_END)
    file_size = file_stream.tell()
    
    # Rewind the cursor back to the start so the file can be read/saved later
    file_stream.seek(0)
    
    return file_size <= MAX_FILE_SIZE


def get_safe_filename(filename: str) -> str:
    """Strip dangerous characters and slashes from a filename."""
    return secure_filename(filename)