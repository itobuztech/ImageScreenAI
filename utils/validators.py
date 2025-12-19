# Dependencies
import magic
from PIL import Image
from pathlib import Path
from typing import Tuple
from utils.logger import get_logger
from config.settings import settings
from config.constants import MIN_IMAGE_DIMENSION
from config.constants import MAX_IMAGE_DIMENSION


# Setup Logging
logger = get_logger(__name__)


class ValidationError(Exception):
    """
    Custom validation error
    """
    pass


class ImageValidator:
    """
    Validate uploaded images
    """
    @staticmethod
    def validate_file_size(file_size: int) -> None:
        """
        Validate file size
        """
        if (file_size > settings.max_file_size_bytes):
            raise ValidationError(f"File size {file_size} bytes exceeds maximum {settings.max_file_size_bytes} bytes")

        if (file_size == 0):
            raise ValidationError("File is empty")
    

    @staticmethod
    def validate_file_extension(filename: str) -> None:
        """
        Validate file extension
        """
        extension = Path(filename).suffix.lower()
        
        if (extension not in settings.ALLOWED_EXTENSIONS):
            raise ValidationError(f"File extension {extension} not allowed. Allowed: {', '.join(settings.ALLOWED_EXTENSIONS)}")
    

    @staticmethod
    def validate_image_content(file_path: Path) -> Tuple[int, int]:
        """
        Validate image can be opened and get dimensions
        """
        try:
            with Image.open(file_path) as image:
                width, height = image.size
                
                # Validate dimensions
                if ((width < MIN_IMAGE_DIMENSION) or (height < MIN_IMAGE_DIMENSION)):
                    raise ValidationError(f"Image dimensions ({width}x{height}) too small. Minimum: {MIN_IMAGE_DIMENSION}px")
                
                if ((width > MAX_IMAGE_DIMENSION) or (height > MAX_IMAGE_DIMENSION)):
                    raise ValidationError(f"Image dimensions ({width}x{height}) too large. Maximum: {MAX_IMAGE_DIMENSION}px")
                
                # Verify format
                if (image.format.lower() not in ['jpeg', 'png', 'webp']):
                    raise ValidationError(f"Unsupported image format: {image.format}")
                
                return width, height
                
        except ValidationError:
            raise

        except Exception as e:
            raise ValidationError(f"Cannot open image: {str(e)}")
    

    @staticmethod
    def validate_mime_type(file_path: Path) -> None:
        """
        Validate MIME type matches image
        """
        try:
            mime = magic.from_file(str(file_path), mime = True)

            if (not mime.startswith('image/')):
                raise ValidationError(f"File is not an image. MIME type: {mime}")
        
        except Exception as e:
            logger.warning(f"MIME type validation failed: {e}")
            # Don't fail if python-magic is not available
    

    @classmethod
    def validate_image(cls, file_path: Path, filename: str, file_size: int) -> Tuple[int, int]:
        """
        Comprehensive image validation
        """
        cls.validate_file_size(file_size)
        cls.validate_file_extension(filename)

        dimensions = cls.validate_image_content(file_path)
        cls.validate_mime_type(file_path)  # Optional, commented out if python-magic not available
        
        logger.debug(f"Validated image: {filename} ({dimensions[0]}x{dimensions[1]})")

        return dimensions