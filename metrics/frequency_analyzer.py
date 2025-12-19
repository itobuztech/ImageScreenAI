# Dependencies
import numpy as np
from scipy import fft
from utils.logger import get_logger
from config.schemas import MetricResult
from config.constants import MetricType
from utils.image_processor import ImageProcessor
from config.constants import FREQUENCY_ANALYSIS_PARAMS

# Suppress NumPy warning 
np.seterr(divide  = 'ignore', 
          invalid = 'ignore',
         )
         

# Setup Logging
logger = get_logger(__name__)


class FrequencyAnalyzer:
    """
    FFT-based frequency domain analysis for AI detection
    
    Core principle:
    ---------------
    - Real photos : Smooth frequency falloff (natural optical blur)
    - AI images   : Unnatural frequency spikes or gaps (artifacts from generation)
    
    Method:
    -------
    1. Convert to luminance
    2. Compute 2D FFT
    3. Compute radial frequency spectrum
    4. Analyze high-frequency content and distribution patterns
    """
    def __init__(self):
        self.image_processor = ImageProcessor()
    

    def detect(self, image: np.ndarray) -> MetricResult:
        """
        Run frequency domain analysis
        
        Arguments:
        ----------
            image { np.ndarray } : RGB image array (H, W, 3)
        
        Returns:
        --------
            { MetricResult }     : Structured frequency-domain metric result containing:
                                   - score      : Suspicion score [0.0, 1.0]
                                   - confidence : Reliability of frequency evidence
                                   - details    : FFT and spectrum diagnostics
        """
        try:
            logger.debug(f"Running frequency analysis on image shape {image.shape}")
            
            # Convert to luminance
            luminance                   = self.image_processor.rgb_to_luminance(image = image)

            # Normalize luminance (remove DC component for FFT stability)
            normalized_luminance        = luminance - np.mean(luminance)

            if not np.any(normalized_luminance):
                logger.debug("FFT skipped: zero-variance luminance")
                
                return MetricResult(metric_type = MetricType.FREQUENCY,
                                    score       = FREQUENCY_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                                    confidence  = 0.0,
                                    details     = {"reason": "zero_variance_luminance"}
                                   )
            
            # Compute FFT on normalized_luminance
            fft_magnitude               = self._compute_fft_magnitude(luminance = normalized_luminance)
            
            # Analyze radial frequency spectrum
            radial_spectrum             = self._compute_radial_spectrum(fft_magnitude = fft_magnitude)
            
            # Detect anomalies
            anomaly_score, freq_details = self._analyze_frequency_anomalies(radial_spectrum = radial_spectrum)
            
            logger.debug(f"Frequency analysis: Anomaly Score={anomaly_score:.3f}")
            
            # Distance from neutral = stronger evidence = higher confidence
            confidence                  = float(np.clip((abs(anomaly_score - FREQUENCY_ANALYSIS_PARAMS.NEUTRAL_SCORE) * 2.0), 0.0, 1.0))
            
            return MetricResult(metric_type = MetricType.FREQUENCY,
                                score       = float(anomaly_score),
                                confidence  = confidence,
                                details     = {"spectrum_bins" : int(len(radial_spectrum)),
                                                **freq_details,
                                              }
                               )
            
        except Exception as e:
            logger.error(f"Frequency analysis failed: {e}")

            # Return neutral score on error
            return MetricResult(metric_type = MetricType.FREQUENCY,
                                score       = FREQUENCY_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                                confidence  = 0.0,
                                details     = {"error" : "frequency_analysis_failed"},
                               )
    

    def _compute_fft_magnitude(self, luminance: np.ndarray) -> np.ndarray:
        """
        Compute 2D FFT magnitude spectrum
        
        Arguments:
        ----------
            luminance { np.ndarray } : Luminance channel (H, W)
        
        Returns:
        --------
            { np.ndarray }           : FFT magnitude spectrum (centered)
        """
        # Compute 2D FFT
        f             = fft.fft2(luminance)
        
        # Shift zero frequency to center
        f_shifted     = fft.fftshift(f)
        
        # Compute magnitude spectrum
        magnitude     = np.abs(f_shifted)
        
        # Log scale for better visualization
        magnitude_log = np.log1p(magnitude)
        
        return magnitude_log
    

    def _compute_radial_spectrum(self, fft_magnitude: np.ndarray) -> np.ndarray:
        """
        Compute radial average of frequency spectrum
        
        Arguments:
        ----------
            fft_magnitude { np.ndarray } : FFT magnitude spectrum
        
        Returns:
        --------
            { np.ndarray }               : Radial spectrum (1D array)
        """
        h, w                       = fft_magnitude.shape
        center_y, center_x         = h // 2, w // 2
        
        # Create coordinate grids
        y, x                       = np.ogrid[:h, :w]
        
        # Compute radial distances from center
        r                          = np.sqrt((x - center_x)**2 + (y - center_y)**2).astype(int)
        
        # Maximum radius
        max_radius                 = min(center_x, center_y)
        
        # Compute radial bins
        bins                       = np.linspace(0, max_radius, FREQUENCY_ANALYSIS_PARAMS.BINS + 1)
        radial_spectrum            = np.zeros(FREQUENCY_ANALYSIS_PARAMS.BINS)
        
        # Average magnitude in each radial bin
        for i in range(FREQUENCY_ANALYSIS_PARAMS.BINS):
            mask = (r >= bins[i]) & (r < bins[i + 1])
            
            if np.any(mask):
                radial_spectrum[i] = np.mean(fft_magnitude[mask])
        
        return radial_spectrum
    

    def _analyze_frequency_anomalies(self, radial_spectrum: np.ndarray) -> tuple[float, dict]:
        """
        Analyze frequency spectrum for AI generation artifacts
        
        Checks:
        -------
        1. High-frequency content (AI images often have unnatural HF energy)
        2. Frequency distribution smoothness
        3. Spectral slope deviation from natural images
        
        Arguments:
        ----------
            radial_spectrum { np.ndarray } : Radial frequency spectrum
        
        Returns:
        --------
                { tuple }                  : A tuple containing
                                             - Suspicion score [0.0, 1.0], and
                                             - frequency details in a dictionary
        """
        if (len(radial_spectrum) < FREQUENCY_ANALYSIS_PARAMS.MIN_SPECTRUM_SAMPLES):
            return (FREQUENCY_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                    {"reason"        : "insufficient_frequency_samples",
                     "spectrum_bins" : int(len(radial_spectrum)),
                    }
                   )
        
        # Normalize spectrum
        spectrum_norm    = radial_spectrum / (np.max(radial_spectrum) + 1e-10)
        
        # High-frequency Energy Analysis
        high_freq_start  = int(len(spectrum_norm) * FREQUENCY_ANALYSIS_PARAMS.HIGH_FREQ_THRESHOLD)

        if (high_freq_start >= len(spectrum_norm) - 1):
            return (FREQUENCY_ANALYSIS_PARAMS.NEUTRAL_SCORE, 
                    {"reason" : "invalid_frequency_partition"}
                   )

        high_freq_energy = np.mean(spectrum_norm[high_freq_start:])
        low_freq_energy  = np.mean(spectrum_norm[:high_freq_start])
        
        hf_ratio         = high_freq_energy / (low_freq_energy + 1e-10)
        
        # Natural images : HF ratio typically 0.1-0.3
        # AI images      : Can be higher (0.3-0.6) or lower (<0.1)
        hf_anomaly       = 0.0
        
        if (hf_ratio > FREQUENCY_ANALYSIS_PARAMS.HF_RATIO_UPPER):
            hf_anomaly  = min(1.0, (hf_ratio - FREQUENCY_ANALYSIS_PARAMS.HF_RATIO_UPPER) * FREQUENCY_ANALYSIS_PARAMS.HF_UPPER_SCALE)
        
        elif (hf_ratio < FREQUENCY_ANALYSIS_PARAMS.HF_RATIO_LOWER):
            hf_anomaly  = min(1.0, (FREQUENCY_ANALYSIS_PARAMS.HF_RATIO_LOWER - hf_ratio) * FREQUENCY_ANALYSIS_PARAMS.HF_LOWER_SCALE)
        
        # Spectral Smoothness Analysis
        spectral_diff   = np.abs(np.diff(spectrum_norm))
        roughness       = np.mean(spectral_diff)
        roughness_score = np.clip(roughness * FREQUENCY_ANALYSIS_PARAMS.ROUGHNESS_SCALE, 0.0, 1.0)
        
        # Power Law Deviation Analysis
        x               = np.arange(1, len(spectrum_norm) + 1)
        log_spectrum    = np.log(spectrum_norm + 1e-10)
        log_x           = np.log(x)
        
        # Linear fit in log-log space
        coeffs          = np.polyfit(log_x, log_spectrum, 1)
        fitted          = np.polyval(coeffs, log_x)
        deviation       = np.mean(np.abs(log_spectrum - fitted))
        deviation_score = np.clip(deviation * FREQUENCY_ANALYSIS_PARAMS.DEVIATION_SCALE, 0.0, 1.0)
        
        # Combine scores
        weights         = FREQUENCY_ANALYSIS_PARAMS.SUBMETRIC_WEIGHTS

        combined_score  = (weights['hf_anomaly'] * hf_anomaly + weights['roughness'] * roughness_score + weights['deviation'] * deviation_score)

        final_score     = float(np.clip(combined_score, 0.0, 1.0))
        
        frequency_dict  = {"low_freq_energy"     : float(low_freq_energy),
                           "high_freq_energy"    : float(high_freq_energy),
                           "hf_ratio"            : float(hf_ratio),
                           "hf_anomaly"          : float(hf_anomaly),
                           "roughness"           : float(roughness),
                           "roughness_score"     : float(roughness_score),
                           "spectral_deviation"  : float(deviation),
                           "deviation_score"     : float(deviation_score),
                           "high_freq_start_bin" : int(high_freq_start),
                          }

        logger.debug(f"FFT scores - HF anomaly: {hf_anomaly:.3f}, roughness: {roughness_score:.3f}, deviation: {deviation_score:.3f}")
        
        return (final_score, frequency_dict)
