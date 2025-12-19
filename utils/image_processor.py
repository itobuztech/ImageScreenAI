# Dependencies
import cv2
import numpy as np
from PIL import Image
from pathlib import Path
from typing import Tuple
from typing import Optional
from utils.logger import get_logger
from config.constants import LUMINANCE_WEIGHTS


# Setup Logging
logger = get_logger(__name__)


class ImageProcessor:
    """
    Image loading and preprocessing utilities
    """
    @staticmethod
    def load_image(file_path: Path) -> np.ndarray:
        """
        Load image as numpy array in RGB format
        
        Arguments:
        ----------
            file_path { Path } : Path of the image file needs to be loaded

        Returns:
        --------
            { np.ndarray }     : Image array in RGB format (H, W, 3)
        """
        try:
            image = cv2.imread(str(file_path))
            
            if image is None:
                raise ValueError(f"Failed to load image: {file_path}")
            
            # Convert BGR to RGB
            image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)

            logger.debug(f"Loaded image: {file_path.name} shape={image.shape}")
            return image

        except Exception as e:
            logger.error(f"Error loading image {file_path}: {e}")
            raise

    
    @staticmethod
    def rgb_to_luminance(image: np.ndarray) -> np.ndarray:
        """
        Convert RGB image to luminance using ITU-R BT.709 standard
        
        Arguments:
        ----------
            image { np.ndarray } : RGB image array (H, W, 3)
        
        Returns:
        --------
             { np.ndarray }      : Luminance array (H, W)
        """
        if ((image.ndim != 3) or (image.shape[2] != 3)):
            raise ValueError(f"Expected RGB image (H, W, 3), got shape {image.shape}")
        
        r, g, b   = LUMINANCE_WEIGHTS
        
        luminance = r * image[:, :, 0] + g * image[:, :, 1] + b * image[:, :, 2]
        
        return luminance.astype(np.float32)

    
    @staticmethod
    def compute_gradients(luminance: np.ndarray) -> Tuple[np.ndarray, np.ndarray]:
        """
        Compute Sobel gradients
        
        Arguments:
        ----------
            luminance { np.ndarray } : Luminance array (H, W)
        
        Returns:
        --------
                   { tuple }         : Tuple of (gradient_x, gradient_y)
        """
        gx = cv2.Sobel(luminance, cv2.CV_64F, 1, 0, ksize = 3)
        gy = cv2.Sobel(luminance, cv2.CV_64F, 0, 1, ksize = 3)

        return gx, gy
    

    @staticmethod
    def normalize_image(image: np.ndarray) -> np.ndarray:
        """
        Normalize image to [0, 1] range
        """
        normalized_image = image.astype(np.float32) / 255.0
        
        return normalized_image
    

    @staticmethod
    def resize_if_needed(image: np.ndarray, max_dimension: int = 2048) -> np.ndarray:
        """
        Resize image if larger than max_dimension while maintaining aspect ratio
        
        Arguments:
        ----------
            image        { np.ndarray } : Input image

            max_dimension   { int }     : Maximum dimension (width or height)
        
        Returns:
        --------
                 { np.ndarray }         : Resized image if needed, otherwise original
        """
        h, w = image.shape[:2]

        if (max(h, w) <= max_dimension):
            return image
        
        scale   = max_dimension / max(h, w)
        new_w   = int(w * scale)
        new_h   = int(h * scale)
        
        resized = cv2.resize(image, (new_w, new_h), interpolation = cv2.INTER_AREA)
        
        logger.debug(f"Resized image from {w}x{h} to {new_w}x{new_h}")

        return resized
    
    @staticmethod
    def extract_patches(image: np.ndarray, patch_size: int, stride: int, max_patches: Optional[int] = None) -> np.ndarray:
        """
        Extract patches from image
        
        Arguments:
        ----------
            image      { np.ndarray } : Input image (H, W) or (H, W, C)
            
            patch_size    { int }     : Size of patches
            
            stride        { int }     : Stride between patches
            
            max_patches   { int }     : Maximum number of patches to extract
        
        Returns:
        --------
                 { np.ndarray }       : Array of patches
        """
        h, w    = image.shape[:2]
        patches = list()
        
        for y in range(0, h - patch_size + 1, stride):
            for x in range(0, w - patch_size + 1, stride):
                patch = image[y:y+patch_size, x:x+patch_size]

                patches.append(patch)
                
                if (max_patches and (len(patches) >= max_patches)):
                    return np.array(patches)
        
        return np.array(patches)
