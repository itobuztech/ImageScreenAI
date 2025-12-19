# Dependencies
import numpy as np
from scipy.stats import entropy
from utils.logger import get_logger
from config.schemas import MetricResult
from config.constants import MetricType
from utils.image_processor import ImageProcessor
from config.constants import TEXTURE_ANALYSIS_PARAMS

# Suppress NumPy warning 
np.seterr(divide  = 'ignore', 
          invalid = 'ignore',
         )
         

# Setup Logging
logger = get_logger(__name__)


class TextureAnalyzer:
    """
    Statistical texture analysis for AI detection
    
    Core principle:
    ---------------
    - Real photos : Natural texture variation (random but structured)
    - AI images   : Either too smooth or repetitive patterns
    
    Method:
    -------
    1. Extract local patches
    2. Compute texture features (contrast, entropy)
    3. Analyze texture consistency and distribution
    4. Detect unnaturally smooth regions
    """
    def __init__(self):
        """
        Initialize TextureAnalyzer Class
        """
        self.patch_size      = TEXTURE_ANALYSIS_PARAMS.PATCH_SIZE
        self.n_patches       = TEXTURE_ANALYSIS_PARAMS.N_PATCHES
        self.image_processor = ImageProcessor()
        self._rng            = np.random.default_rng(seed = TEXTURE_ANALYSIS_PARAMS.RANDOM_SEED)
    

    def detect(self, image: np.ndarray) -> MetricResult:
        """
        Run texture analysis
        
        Arguments:
        ----------
            image { np.ndarray } : RGB image array (H, W, 3)
        
        Returns:
        --------
            { MetricResult }     : Structured Texture-domain metric result containing:
                                   - score      : Suspicion score [0.0, 1.0]
                                   - confidence : Reliability of texture evidence
                                   - details    : Texture forensics and statistics
        """
        try:
            logger.debug(f"Running texture analysis on image shape {image.shape}")
            
            # Convert to luminance
            luminance                          = self.image_processor.rgb_to_luminance(image = image)
            
            # Extract patches
            patches                            = self._extract_patches(luminance = luminance)
            
            if (len(patches) == 0):
                logger.warning("No patches extracted for texture analysis")
                return MetricResult(metric_type = MetricType.TEXTURE,
                                    score       = TEXTURE_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                                    confidence  = 0.0,
                                    details     = {"reason": "no_patches_extracted"},
                                   )           
            
            # Compute texture features
            texture_features, texture_metadata = self._compute_texture_features(patches = patches)
            
            # Analyze for anomalies
            texture_score, texture_details     = self._analyze_texture_anomalies(features = texture_features,
                                                                                 metadata = texture_metadata,
                                                                                )
            
            # Calculate Confidence
            confidence                         = float(np.clip((abs(texture_score - TEXTURE_ANALYSIS_PARAMS.NEUTRAL_SCORE) * 2.0), 0.0, 1.0))
            
            logger.debug(f"Texture analysis: Texture Score={texture_score:.3f}, patches={len(patches)}")
            
            return MetricResult(metric_type = MetricType.TEXTURE,
                                score       = float(texture_score),
                                confidence  = confidence,
                                details     = {"patches_total" : int(len(patches)),
                                               **texture_metadata,
                                               **texture_details,
                                              },
                               )
            
        except Exception as e:
            logger.error(f"Texture analysis failed: {e}")
            
            # Return neutral score on error
            return MetricResult(metric_type = MetricType.TEXTURE,
                                score       = TEXTURE_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                                confidence  = 0.0,
                                details     = {"error": "texture_analysis_failed"},
                               )
    

    def _extract_patches(self, luminance: np.ndarray) -> np.ndarray:
        """
        Extract random patches from image
        """
        h, w = luminance.shape
        
        if ((h < self.patch_size) or (w < self.patch_size)):
            logger.warning(f"Image too small for patch size {self.patch_size}")
            return np.array([])
        
        patches = list()
        
        for _ in range(self.n_patches):
            y     = self._rng.integers(0, h - self.patch_size)
            x     = self._rng.integers(0, w - self.patch_size)
            
            patch = luminance[y:y+self.patch_size, x:x+self.patch_size]
            patches.append(patch)
        
        return np.array(patches)
    

    def _compute_texture_features(self, patches: np.ndarray) -> tuple[dict, dict]:
        """
        Compute texture features for each patch
        
        Features:
        ---------
        1. Local contrast (standard deviation)
        2. Entropy (randomness)
        3. Smoothness (inverse of variance)
        4. Edge density
        
        Arguments:
        ----------
            patches { np.ndarray } : Array of patches
        
        Returns:
        --------
            { tuple }              : A tuple containing
                                     - A dictionary of feature arrays
                                     - A dictionary of texture analysis metadata
        """
        contrasts       = list()
        entropies       = list()
        smoothnesses    = list()
        edge_densities  = list()
        uniform_skipped = 0
        
        for patch in patches:
            pmin = patch.min()
            pmax = patch.max()

            if ((pmax - pmin < 1e-6)):
                # skip fully uniform patch entirely
                uniform_skipped += 1
                continue 
             
            # Contrast (std deviation)
            contrast = np.std(patch)
            contrasts.append(contrast)
            
            # Entropy (using histogram)
            hist, _  = np.histogram(patch,
                                    bins  = TEXTURE_ANALYSIS_PARAMS.HISTOGRAM_BINS,
                                    range = TEXTURE_ANALYSIS_PARAMS.HISTOGRAM_RANGE,
                                   )

            hist     = hist / (np.sum(hist) + 1e-10)
            ent      = entropy(hist + 1e-10)
            entropies.append(ent)
            
            # Smoothness (inverse of variance, scaled)
            variance   = np.var(patch)
            smoothness = 1.0 / (1.0 + variance)
            smoothnesses.append(smoothness)
            
            # Edge density (using Sobel)
            gx, gy       = self.image_processor.compute_gradients(luminance = patch)
            gradient_mag = np.sqrt(gx**2 + gy**2)

            edge_density = np.mean(gradient_mag > TEXTURE_ANALYSIS_PARAMS.EDGE_THRESHOLD) 
            edge_densities.append(edge_density)

        # Construct results in proper format
        features = {"contrast"     : np.array(contrasts),
                    "entropy"      : np.array(entropies),
                    "smoothness"   : np.array(smoothnesses),
                    "edge_density" : np.array(edge_densities),
                   }

        metadata = {"patches_used"            : int(len(contrasts)),
                    "uniform_patches_skipped" : int(uniform_skipped),
                   }

        
        return features, metadata
    

    def _analyze_texture_anomalies(self, features: dict, metadata: dict) -> tuple[float, dict]:
        """
        Analyze texture features for AI generation artifacts
        
        Checks:
        -------
        1. Excessive smoothness (too many overly smooth patches)
        2. Entropy distribution (too uniform = suspicious)
        3. Contrast consistency
        
        Arguments:
        ----------
            features { dict } : Dictionary of texture features

            metadata { dict } : Dictionary of texture analysis metadata
        
        Returns:
        --------
            { tuple }         : A tuple containing:
                                - Suspicion score [0.0, 1.0]
                                - Texture statistics
        """
        contrast     = features['contrast']
        entropy_vals = features['entropy']
        smoothness   = features['smoothness']
        edge_density = features['edge_density']

        if ((len(contrast) == 0) or (len(entropy_vals) == 0) or (len(smoothness) == 0) or (len(edge_density) == 0)):
            logger.debug("All texture features filtered out; returning neutral score")
            return (TEXTURE_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                    {"reason": "all_texture_features_filtered"},
                   )

        # Early exit: all patches nearly uniform
        if (np.all(contrast < 1e-6)):
            logger.debug("All texture patches near-uniform; returning neutral score")
            return (TEXTURE_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                    {"reason": "all_patches_near_uniform"},
                   )
        
        # Smoothness Analysis
        smooth_ratio       = np.mean(smoothness > TEXTURE_ANALYSIS_PARAMS.SMOOTHNESS_THRESHOLD)
        smoothness_anomaly = 0.0
        
        if (smooth_ratio > TEXTURE_ANALYSIS_PARAMS.SMOOTH_RATIO_THRESHOLD):
            # More than 40% very smooth patches
            smoothness_anomaly = min(1.0, (smooth_ratio - TEXTURE_ANALYSIS_PARAMS.SMOOTH_RATIO_THRESHOLD) * TEXTURE_ANALYSIS_PARAMS.SMOOTH_RATIO_SCALE)
        
        # Entropy distribution Analysis
        entropy_cv      = np.std(entropy_vals) / (np.mean(entropy_vals) + 1e-10)
        entropy_anomaly = 0.0
        
        if (entropy_cv < TEXTURE_ANALYSIS_PARAMS.ENTROPY_CV_THRESHOLD):  
            # Too uniform
            entropy_anomaly = (TEXTURE_ANALYSIS_PARAMS.ENTROPY_CV_THRESHOLD - entropy_cv) * TEXTURE_ANALYSIS_PARAMS.ENTROPY_SCALE
        
        # Contrast distribution Analysis
        contrast_cv      = np.std(contrast) / (np.mean(contrast) + 1e-10)
        contrast_anomaly = 0.0
        
        if (contrast_cv < TEXTURE_ANALYSIS_PARAMS.CONTRAST_CV_LOW):
            # Too uniform
            contrast_anomaly = (TEXTURE_ANALYSIS_PARAMS.CONTRAST_CV_LOW - contrast_cv) * TEXTURE_ANALYSIS_PARAMS.CONTRAST_LOW_SCALE 
        
        elif (contrast_cv > TEXTURE_ANALYSIS_PARAMS.CONTRAST_CV_HIGH):
            # Too variable (suspicious)
            contrast_anomaly = min(1.0, (contrast_cv - TEXTURE_ANALYSIS_PARAMS.CONTRAST_CV_HIGH) * TEXTURE_ANALYSIS_PARAMS.CONTRAST_HIGH_SCALE)

        
        # Edge density consistency Analysis
        edge_cv      = np.std(edge_density) / (np.mean(edge_density) + 1e-10)
        edge_anomaly = 0.0
        
        if (edge_cv < TEXTURE_ANALYSIS_PARAMS.EDGE_CV_THRESHOLD):
            edge_anomaly = (TEXTURE_ANALYSIS_PARAMS.EDGE_CV_THRESHOLD - edge_cv) * TEXTURE_ANALYSIS_PARAMS.EDGE_SCALE

        # Clipping Sub-anomalies
        smoothness_anomaly = np.clip(smoothness_anomaly, 0.0, 1.0)
        entropy_anomaly    = np.clip(entropy_anomaly, 0.0, 1.0)
        contrast_anomaly   = np.clip(contrast_anomaly, 0.0, 1.0)
        edge_anomaly       = np.clip(edge_anomaly, 0.0, 1.0)
        
        # Combine scores
        weights            = TEXTURE_ANALYSIS_PARAMS.SUBMETRIC_WEIGHTS
        texture_score      = (weights['smoothness_anomaly'] * smoothness_anomaly + weights['entropy_anomaly'] * entropy_anomaly + weights['contrast_anomaly'] * contrast_anomaly + weights['edge_anomaly'] * edge_anomaly)
        final_score        = float(np.clip(texture_score, 0.0, 1.0))

        detailed_stats     = {"smooth_ratio"      : float(smooth_ratio),
                              "entropy_mean"      : float(np.mean(entropy_vals)),
                              "entropy_cv"        : float(entropy_cv),
                              "contrast_mean"     : float(np.mean(contrast)),
                              "contrast_cv"       : float(contrast_cv),
                              "edge_density_mean" : float(np.mean(edge_density)),
                              "edge_cv"           : float(edge_cv),
                             }

        logger.debug(f"Texture scores - smoothness: {smoothness_anomaly:.3f}, entropy: {entropy_anomaly:.3f}, contrast: {contrast_anomaly:.3f}, edge: {edge_anomaly:.3f}")
        
        return final_score, detailed_stats