# Dependencies
import numpy as np
from utils.logger import get_logger
from config.schemas import MetricResult
from config.constants import MetricType
from utils.image_processor import ImageProcessor
from config.constants import NOISE_ANALYSIS_PARAMS

# Suppress NumPy warning 
np.seterr(divide  = 'ignore', 
          invalid = 'ignore',
         )

         
# Setup Logging
logger = get_logger(__name__)


class NoiseAnalyzer:
    """
    Noise pattern analysis for AI detection
    
    Core principle:
    ---------------
    - Real photos : Sensor noise follows Poisson distribution (shot noise) + Gaussian (read noise)
    - AI images   : Too uniform, artificially smooth, or completely missing noise
    
    Method:
    -------
    1. Extract local patches
    2. Estimate noise variance in each patch
    3. Analyze noise consistency and distribution
    4. Check for unnatural uniformity
    """
    def __init__(self):
        self.image_processor = ImageProcessor()
    

    def detect(self, image: np.ndarray) -> MetricResult:
        """
        Run noise pattern analysis
        
        Arguments:
        ----------
            image { np.ndarray } : RGB image array (H, W, 3)
        
        Returns:
        --------
            { MetricResult }     : Structured Noise-domain metric result containing:
                                   - score      : Suspicion score [0.0, 1.0]
                                   - confidence : Reliability of noise evidence
                                   - details    : Noise related diagnostics
        """
        try:
            logger.debug(f"Running noise analysis on image shape {image.shape}")
            
            # Convert to luminance
            luminance                  = self.image_processor.rgb_to_luminance(image = image)
            
            # Extract patches
            patches                    = self._extract_patches(luminance = luminance)
            
            if (len(patches) == 0):
                logger.warning("No patches extracted for noise analysis")
                return MetricResult(metric_type = MetricType.NOISE,
                                    score       = NOISE_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                                    confidence  = 0.0,
                                    details     = {"reason": "no_patches_extracted"},
                                   )
            
            # Estimate noise in each patch
            noise_estimates, mad_values, laplacian_energy = self._estimate_noise_per_patch(patches = patches)
            
            # Filter Noise Estimates, MAD and Laplacian Energy for finite values only
            filtered_mask                                 = np.isfinite(noise_estimates)
            filtered_noise_estimates                      = noise_estimates[filtered_mask]
            filtered_mad                                  = mad_values[filtered_mask]
            filtered_laplacian_energy                     = laplacian_energy[filtered_mask]
            
            if (len(filtered_noise_estimates) < NOISE_ANALYSIS_PARAMS.MIN_ESTIMATES):
                logger.debug("Insufficient valid noise estimates after filtering")

                return MetricResult(metric_type = MetricType.NOISE,
                                    score       = NOISE_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                                    confidence  = 0.0,
                                    details     = {"reason"        : "insufficient_noise_estimates",
                                                   "patches_total" : int(len(patches)),
                                                   "patches_valid" : int(len(filtered_noise_estimates)),
                                                  },
                                   )

            logger.debug(f"Noise patches: total={len(patches)}, valid={len(filtered_noise_estimates)}")
            
            # Analyze noise distribution
            noise_score, noise_details                    = self._analyze_noise_distribution(noise_estimates  = filtered_noise_estimates,
                                                                                             mad_values       = filtered_mad,
                                                                                             laplacian_energy = filtered_laplacian_energy,
                                                                                            )
            
            # Confidence: distance from neutral
            confidence                                    = float(np.clip((abs(noise_score - NOISE_ANALYSIS_PARAMS.NEUTRAL_SCORE) * 2.0), 0.0, 1.0))
            
            logger.debug(f"Noise analysis: score={noise_score:.3f}, patches={len(patches)}, valid={len(filtered_noise_estimates)}")
            
            return MetricResult(metric_type = MetricType.NOISE,
                                score       = float(noise_score),
                                confidence  = confidence,
                                details     = {"patches_total" : int(len(patches)),
                                               "patches_valid" : int(len(filtered_noise_estimates)),
                                               **noise_details,
                                              },
                               )
            
        except Exception as e:
            logger.error(f"Noise analysis failed: {e}")
            
            # Return neutral score on error
            return MetricResult(metric_type = MetricType.NOISE,
                                score       = NOISE_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                                confidence  = 0.0,
                                details     = {"error": "noise_analysis_failed"},
                               )
                        

    def _extract_patches(self, luminance: np.ndarray) -> np.ndarray:
        """
        Extract patches from image for local noise estimation
        
        Arguments:
        ----------
            luminance { np.ndarray } : Luminance channel (H, W)
        
        Returns:
        --------
            { np.ndarray }           : Array of patches
        """
        patches = self.image_processor.extract_patches(image       = luminance,
                                                       patch_size  = NOISE_ANALYSIS_PARAMS.PATCH_SIZE,
                                                       stride      = NOISE_ANALYSIS_PARAMS.STRIDE,
                                                       max_patches = NOISE_ANALYSIS_PARAMS.SAMPLES,
                                                      )
        
        return patches
    

    def _estimate_noise_per_patch(self, patches: np.ndarray) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
        """
        Estimate noise variance in each patch using median absolute deviation
        
        Uses Median Absolute Deviation (MAD) which is robust to edges/textures
        
        Arguments:
        ----------
            patches { np.ndarray } : Array of image patches (N, patch_size, patch_size)
        
        Returns:
        --------
                 { tuple }         : A tuple containing
                                     - Array of noise estimates per patch
                                     - Array of MAD values
                                     - Array of Laplacian Energy Values
        """
        noise_estimates          = list()
        mad_values               = list()
        laplacian_energy_values  = list()
        
        for patch in patches:
            # Skip patches with too much structure (edges, textures)
            variance = np.var(patch)
            
            if (variance < NOISE_ANALYSIS_PARAMS.VARIANCE_LOW_THRESHOLD):     
                # Too uniform, skip
                continue
            
            if (variance > NOISE_ANALYSIS_PARAMS.VARIANCE_HIGH_THRESHOLD):  
                # Too much structure, skip
                continue
            
            # Use Median Absolute Deviation for robust noise estimation
            laplacian   = self._apply_laplacian(patch = patch)
            mad         = np.median(np.abs(laplacian - np.median(laplacian)))
            
            # Convert MAD to noise standard deviation estimate: For Gaussian noise: σ ≈ 1.4826 × MAD
            noise_std   = NOISE_ANALYSIS_PARAMS.MAD_TO_STD_FACTOR * mad

            # Calculate Laplacian Energy
            lap_energy  = float(np.mean(laplacian ** 2))
            
            # Append corresponding values to their storages
            mad_values.append(mad)
            noise_estimates.append(noise_std)
            laplacian_energy_values.append(lap_energy)
        
        return np.array(noise_estimates), np.array(mad_values), np.array(laplacian_energy_values)
    

    def _apply_laplacian(self, patch: np.ndarray) -> np.ndarray:
        """
        Apply Laplacian filter to isolate high-frequency noise
        
        Arguments:
        ----------
            patch { np.ndarray } : Image patch
        
        Returns:
        --------
            { np.ndarray }       : Laplacian-filtered patch
        """
        # Simple 3x3 Laplacian kernel
        kernel = np.array([[0,  1, 0],
                          [1, -4, 1],
                          [0,  1, 0]],
                         )
        
        # Pad patch
        padded = np.pad(patch, 1, mode = 'reflect')
        
        # Apply convolution
        h, w   = patch.shape
        result = np.zeros_like(patch)
        
        for i in range(h):
            for j in range(w):
                region        = padded[i:i+3, j:j+3]
                result[i, j]  = np.sum(region * kernel)
        
        return result
    

    def _analyze_noise_distribution(self, noise_estimates: np.ndarray, mad_values: np.ndarray, laplacian_energy: np.ndarray,) -> tuple[float, dict]:
        """
        Analyze noise distribution for anomalies
        
        Checks:
        -------
        1. Coefficient of variation (consistency)
        2. Overall noise level (too low = suspicious)
        3. Distribution shape (too uniform = suspicious)
        
        Arguments:
        ----------
            noise_estimates  { np.ndarray } : Array of noise standard deviations

            mad_values       { np.ndarray } : Array of MAD values

            laplacian_energy { np.ndarray } : Array of Laplacian Energy Values 
        
        Returns:
        --------
                    { tuple }              : A tuple containing:
                                             - Suspicion score [0.0, 1.0]
                                             - Noise Distribution detailed diagnostics
        """
        if (len(noise_estimates) < NOISE_ANALYSIS_PARAMS.MIN_ESTIMATES):
            return (NOISE_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                    {"reason": "insufficient_noise_samples"},
                   )
        
        # Remove outliers (keep middle 80%)
        q10                 = np.percentile(noise_estimates, NOISE_ANALYSIS_PARAMS.OUTLIER_PERCENTILE_LOW)
        q90                 = np.percentile(noise_estimates, NOISE_ANALYSIS_PARAMS.OUTLIER_PERCENTILE_HIGH)
        filtered            = noise_estimates[(noise_estimates >= q10) & (noise_estimates <= q90)]
        
        if (len(filtered) < NOISE_ANALYSIS_PARAMS.MIN_FILTERED_SAMPLES):
            return (NOISE_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                    {"reason": "insufficient_filtered_samples"},
                   ) 
        
        mean_noise          = np.mean(filtered)
        std_noise           = np.std(filtered)
        
        # Coefficient of Variation (CV) Analysis
        cv                  = std_noise / (mean_noise + 1e-10)
        cv_anomaly          = 0.0
        
        if (cv < NOISE_ANALYSIS_PARAMS.CV_UNIFORM_THRESHOLD):
            # Too uniform
            cv_anomaly = (NOISE_ANALYSIS_PARAMS.CV_UNIFORM_THRESHOLD - cv) * NOISE_ANALYSIS_PARAMS.CV_UNIFORM_SCALE
        
        elif (cv > NOISE_ANALYSIS_PARAMS.CV_VARIABLE_THRESHOLD):
            # Too variable
            cv_anomaly = min(1.0, (cv - NOISE_ANALYSIS_PARAMS.CV_VARIABLE_THRESHOLD) * NOISE_ANALYSIS_PARAMS.CV_VARIABLE_SCALE)
        
        # Overall noise level Analysis
        noise_level_anomaly = 0.0

        if (mean_noise < NOISE_ANALYSIS_PARAMS.LEVEL_CLEAN_THRESHOLD):
            # Too clean
            noise_level_anomaly = (NOISE_ANALYSIS_PARAMS.LEVEL_CLEAN_THRESHOLD - mean_noise) / NOISE_ANALYSIS_PARAMS.LEVEL_CLEAN_THRESHOLD
        
        elif (mean_noise < NOISE_ANALYSIS_PARAMS.LEVEL_LOW_THRESHOLD):
            # Slightly low
            noise_level_anomaly = (NOISE_ANALYSIS_PARAMS.LEVEL_LOW_THRESHOLD - mean_noise) / NOISE_ANALYSIS_PARAMS.LEVEL_LOW_THRESHOLD * 0.5

        
        # Distribution shape Analysis
        q25                 = np.percentile(filtered, NOISE_ANALYSIS_PARAMS.IQR_PERCENTILE_LOW)
        q75                 = np.percentile(filtered, NOISE_ANALYSIS_PARAMS.IQR_PERCENTILE_HIGH)
        iqr                 = q75 - q25
        iqr_ratio           = iqr / (mean_noise + 1e-10)
        
        iqr_anomaly         = 0.0
        
        if (iqr_ratio < NOISE_ANALYSIS_PARAMS.IQR_THRESHOLD):
            iqr_anomaly = (NOISE_ANALYSIS_PARAMS.IQR_THRESHOLD - iqr_ratio) * NOISE_ANALYSIS_PARAMS.IQR_SCALE

        # Clip sub-anomalies for safety
        cv_anomaly          = np.clip(cv_anomaly, 0.0, 1.0)
        noise_level_anomaly = np.clip(noise_level_anomaly, 0.0, 1.0)
        iqr_anomaly         = np.clip(iqr_anomaly, 0.0, 1.0)

        # Combine scores
        weights             = NOISE_ANALYSIS_PARAMS.SUBMETRIC_WEIGHTS
        combined_score      = (weights['cv_anomaly'] * cv_anomaly + weights['noise_level_anomaly'] * noise_level_anomaly + weights['iqr_anomaly'] * iqr_anomaly)
        final_score         = float(np.clip(combined_score, 0.0, 1.0))

        # Calculate Forensic Stats
        mad_mean            = float(np.mean(mad_values)) if len(mad_values) else 0.0
        laplacian_energy_mu = float(np.mean(laplacian_energy)) if len(laplacian_energy) else 0.0

        noise_details_dict  = {"mean_noise"          : float(mean_noise),
                               "std_noise"           : float(std_noise),
                               "cv"                  : float(cv),
                               "cv_anomaly"          : float(cv_anomaly),
                               "noise_level_anomaly" : float(noise_level_anomaly),
                               "iqr_ratio"           : float(iqr_ratio),
                               "iqr_anomaly"         : float(iqr_anomaly),
                               "mad_mean"            : mad_mean,
                               "laplacian_energy"    : laplacian_energy_mu,
                              }

        logger.debug(f"Noise scores - CV: {cv:.3f}, mean: {mean_noise:.3f}, IQR ratio: {iqr_ratio:.3f}")

        return final_score, noise_details_dict
        