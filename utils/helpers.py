# Dependencies
import re
import uuid
import hashlib
from pathlib import Path
from datetime import datetime
from datetime import timedelta
from utils.logger import get_logger


# Setup Logging
logger = get_logger(__name__)


def generate_unique_id() -> str:
    """
    Generate unique ID for files/reports
    """
    unique_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"
    
    return unique_id


def calculate_hash(file_path: Path) -> str:
    """
    Calculate SHA256 hash of file
    """
    sha256 = hashlib.sha256()

    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            sha256.update(chunk)
    
    hash = sha256.hexdigest()

    return hash


def format_filesize(size_bytes: int) -> str:
    """
    Format file size in human-readable format
    """
    for unit in ['B', 'KB', 'MB', 'GB']:
        if (size_bytes < 1024.0):
            return f"{size_bytes:.2f} {unit}"

        size_bytes /= 1024.0

    file_size = f"{size_bytes:.2f} TB"
    
    return file_size


def cleanup_old_files(directory: Path, days: int = 7) -> int:
    """
    Clean up files older than specified days
    
    Arguments:
    ----------
        directory { Path } : Directory to clean
        
        days      { int }  : Files older than this will be deleted
    
    Returns:
    --------
            { int }        : Number of files deleted
    """
    if not directory.exists():
        return 0
    
    cutoff  = datetime.now() - timedelta(days = days)
    deleted = 0
    
    for file_path in directory.iterdir():
        if file_path.is_file():
            file_time = datetime.fromtimestamp(file_path.stat().st_mtime)
            
            if (file_time < cutoff):
                try:
                    file_path.unlink()
                    deleted += 1
                    logger.debug(f"Deleted old file: {file_path.name}")

                except Exception as e:
                    logger.error(f"Failed to delete {file_path.name}: {e}")
    
    if (deleted > 0):
        logger.info(f"Cleaned up {deleted} files from {directory.name}")
    
    return deleted


def safe_filename(filename: str) -> str:
    """
    Sanitize filename for safe storage
    """
    # Remove any path components
    filename = Path(filename).name
    
    # Replace unsafe characters
    filename = re.sub(r'[^\w\s.-]', '', filename)
    
    # Limit length
    if (len(filename) > 255):
        name, ext = filename.rsplit('.', 1) if '.' in filename else (filename, '')
        filename  = name[:250] + ('.' + ext if ext else '')
        
    return filename