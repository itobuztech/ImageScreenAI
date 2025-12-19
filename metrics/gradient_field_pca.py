# Dependencies
import numpy as np
from utils.logger import get_logger
from config.schemas import MetricResult
from config.constants import MetricType
from utils.image_processor import ImageProcessor
from config.constants import GRADIENT_FIELD_PCA_PARAMS

# Suppress NumPy warning 
np.seterr(divide  = 'ignore', 
          invalid = 'ignore',
         )


# Setup Logging
logger = get_logger(__name__)


class GradientFieldPCADetector:
    """
    Detects AI-generated images by analyzing gradient field consistency. Real photos have consistent gradient 
    patterns shaped by physics (lighting, optics). Diffusion models struggle to maintain physically consistent
    gradients due to denoising
    
    Core principle:
    ---------------
    - Real photos : Gradients align with physical light sources (low-dimensional structure)
    - AI images   : Gradients are inconsistent due to patch-based denoising (high-dimensional)
    
    Method:
    -------
    1. Convert to luminance
    2. Compute Sobel gradients (Gx, Gy)
    3. Flatten to gradient vectors per pixel
    4. Compute covariance matrix
    5. PCA eigenvalue analysis
    """
    def __init__(self):
        """
        Initialize Gradient-Field PCA Detector class
        """
        self._range          = np.random.default_rng(seed = GRADIENT_FIELD_PCA_PARAMS.RANDOM_SEED)
        self.image_processor = ImageProcessor()
    

    def detect(self, image: np.ndarray) -> MetricResult:
        """
        Run gradient PCA detection
        
        Arguments:
        ----------
            image { np.ndarray } : RGB image array (H, W, 3)
        
        Returns:
        --------
            { MetricResult }     : Structured metric result containing:
                                   - score      : Suspicion score [0.0, 1.0] (0 = natural, 1 = suspicious)
                                   - confidence : Confidence of this metric's assessment [0.0, 1.0]
                                   - details    : Explainability metadata for UI and reports
        """
        try:
            logger.debug(f"Running gradient PCA detection on image shape {image.shape}")
            
            # Convert image to luminance
            luminance             = self.image_processor.rgb_to_luminance(image = image)
            
            # Compute gradients
            gx, gy                = self.image_processor.compute_gradients(luminance = luminance)
            
            # Flatten and sample gradient vectors
            gradient_vectors      = self._prepare_and_sample_gradients(gx = gx, 
                                                                       gy = gy,
                                                                      )
            
            # Perform PCA
            eigenvalue_ratio      = self._compute_eigenvalue_ratio(gradient_vectors = gradient_vectors)

            if ((len(gradient_vectors) < GRADIENT_FIELD_PCA_PARAMS.MIN_SAMPLES) or (eigenvalue_ratio == GRADIENT_FIELD_PCA_PARAMS.NEUTRAL_SCORE)):
                return MetricResult(metric_type = MetricType.GRADIENT,
                                    score       = GRADIENT_FIELD_PCA_PARAMS.NEUTRAL_SCORE,
                                    confidence  = 0.0,
                                    details     = {"reason"           : "insufficient_gradient_information",
                                                   "original_pixels"  : int(gx.size),
                                                   "filtered_vectors" : int(len(gradient_vectors)),
                                                  },
                                   )
            
            # Convert to suspicion score
            suspicion_score       = self._eigenvalue_to_suspicion(eigenvalue_ratio = eigenvalue_ratio)

            # Confidence inverted relative to suspicion: High eigenvalue_ratio = natural, High suspicion_score = AI-like
            confidence            = abs(eigenvalue_ratio - GRADIENT_FIELD_PCA_PARAMS.EIGENVALUE_RATIO_THRESHOLD)
            normalized_confidence = np.clip((confidence / GRADIENT_FIELD_PCA_PARAMS.EIGENVALUE_RATIO_THRESHOLD), 0.0, 1.0)
            
            logger.debug(f"Gradient PCA: eigenvalue_ratio={eigenvalue_ratio:.3f}, suspicion_score={suspicion_score:.3f}")
            
            return MetricResult(metric_type = MetricType.GRADIENT,
                                score       = float(suspicion_score),
                                confidence  = float(normalized_confidence),
                                details     = {"gradient_vectors_sampled" : len(gradient_vectors),
                                               "eigenvalue_ratio"         : float(eigenvalue_ratio),
                                               "threshold"                : GRADIENT_FIELD_PCA_PARAMS.EIGENVALUE_RATIO_THRESHOLD,
                                               "original_pixels"          : int(gx.size),
                                               "filtered_vectors"         : int(len(gradient_vectors)),
                                              },
                               )
            
        except Exception as e:
            logger.error(f"Gradient PCA detection failed: {e}")
            
            # Return neutral score on error
            return MetricResult(metric_type = MetricType.GRADIENT,
                                score       = GRADIENT_FIELD_PCA_PARAMS.NEUTRAL_SCORE,
                                confidence  = 0.0,
                                details     = {"error" : "Gradient PCA detection failed"},
                               )
    

    def _prepare_and_sample_gradients(self, gx: np.ndarray, gy: np.ndarray) -> np.ndarray:
        """
        Flatten gradients into vectors and sample
        
        Arguments:
        ----------
            gx { np.ndarray } : Gradient in x direction

            gy { np.ndarray } : Gradient in y direction
        
        Returns:
        --------
            { np.ndarray }    : Array of gradient vectors (N, 2) where N <= SAMPLE_SIZE
        """
        # Flatten to vectors
        gx_flat                   = gx.flatten()
        gy_flat                   = gy.flatten()

        # Stack into (N, 2) array
        gradient_vectors          = np.stack([gx_flat, gy_flat], axis = 1)
        original_n                = len(gradient_vectors)

        # Remove zero gradients (uniform regions)
        magnitude                 = np.linalg.norm(gradient_vectors, axis = 1)
        non_zero_mask             = (magnitude > GRADIENT_FIELD_PCA_PARAMS.MAGNITUDE_THRESHOLD)
        finite_mask               = np.isfinite(gradient_vectors).all(axis = 1)

        # Filtering Gradient Vector
        filtered_gradient_vectors = gradient_vectors[non_zero_mask & finite_mask]
        filtered_n                = len(filtered_gradient_vectors)
        
        # Sample if too many points without replacement
        if (len(filtered_gradient_vectors) > GRADIENT_FIELD_PCA_PARAMS.SAMPLE_SIZE):
            indices                   = self._range.choice(a       = len(filtered_gradient_vectors), 
                                                           size    = GRADIENT_FIELD_PCA_PARAMS.SAMPLE_SIZE, 
                                                           replace = False,
                                                          )

            sampled_gradient_vectors  = filtered_gradient_vectors[indices]
        
        else:
            sampled_gradient_vectors  = filtered_gradient_vectors


        sampled_n = len(sampled_gradient_vectors)

        logger.debug(f"Gradient PCA sampling: original={original_n}, filtered={filtered_n}, sampled={sampled_n}")
        
        return sampled_gradient_vectors
    

    def _compute_eigenvalue_ratio(self, gradient_vectors: np.ndarray) -> float:
        """
        Compute ratio of first eigenvalue to total variance
        
        -  Lower ratio  = more diffuse structure = suspicious
        -  Higher ratio = concentrated structure = natural
        
        Arguments:
        ----------
            gradient_vectors { np.ndarray } : Array of gradient vectors (N, 2)
        
        Returns:
        --------
                     { float }              : Ratio of first eigenvalue to sum of eigenvalues
        """
        if (len(gradient_vectors) < GRADIENT_FIELD_PCA_PARAMS.MIN_SAMPLES):
            logger.warning("Insufficient gradient samples for PCA")
            return GRADIENT_FIELD_PCA_PARAMS.NEUTRAL_SCORE
        
        # Compute covariance matrix
        covariance       = np.cov(m    = gradient_vectors.T,
                                  bias = True,
                                 )
        
        # Compute eigenvalues
        eigenvalues      = np.linalg.eigvalsh(covariance)

        # Sort in descending order
        eigenvalues      = np.sort(eigenvalues)[::-1]  
        
        # Ratio of largest eigenvalue to sum
        total_variance   = np.sum(eigenvalues)
        
        if (total_variance < GRADIENT_FIELD_PCA_PARAMS.VARIANCE_THRESHOLD):
            return GRADIENT_FIELD_PCA_PARAMS.NEUTRAL_SCORE
        
        eigenvalue_ratio = eigenvalues[0] / total_variance

        return float(eigenvalue_ratio)
    

    def _eigenvalue_to_suspicion(self, eigenvalue_ratio: float) -> float:
        """
        Convert eigenvalue ratio to suspicion score
        
        - Real photos : High ratio (0.85-0.95) -> Low suspicion
        - AI images   : Low ratio (0.50-0.75) -> High suspicion
        
        Arguments:
        ----------
            eigenvalue_ratio { float } : PCA eigenvalue ratio
        
        Returns:
        --------
                    { float }          : Suspicion score [0.0, 1.0]
        """
        # Invert and scale: higher ratio = lower suspicion
        # Real photos typically have ratio > 0.85 & AI images typically have ratio < 0.75
        if (eigenvalue_ratio >= GRADIENT_FIELD_PCA_PARAMS.EIGENVALUE_RATIO_THRESHOLD):
            # Strong gradient alignment = likely real
            suspicion = max(0.0, (1.0 - eigenvalue_ratio) * 2.0)

        else:
            # Weak alignment = suspicious
            suspicion = 1.0 - (eigenvalue_ratio / GRADIENT_FIELD_PCA_PARAMS.EIGENVALUE_RATIO_THRESHOLD)
        
        return float(np.clip(suspicion, 0.0, 1.0))