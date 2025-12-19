# Dependencies
import numpy as np
from utils.logger import get_logger
from config.schemas import MetricResult
from config.constants import MetricType
from utils.image_processor import ImageProcessor
from config.constants import COLOR_ANALYSIS_PARAMS

# Suppress NumPy warning 
np.seterr(divide  = 'ignore', 
          invalid = 'ignore',
         )


# Setup Logging
logger = get_logger(__name__)


class ColorAnalyzer:
    """
    Color distribution analysis for AI detection

    Core principle:
    ---------------
    - Real photos : Natural color distributions constrained by physics
    - AI images   : Can create unnatural saturation, hue shifts, or impossible color relationships

    Method:
    -------
    1. Convert to multiple color spaces (RGB, HSV)
    2. Analyze color histogram distributions
    3. Check for oversaturation
    4. Detect unnatural color relationships
    """
    def __init__(self):
        self.image_processor = ImageProcessor()


    def detect(self, image: np.ndarray) -> MetricResult:
        """
        Run color distribution analysis
        
        Arguments:
        ----------
            image { np.ndarray } : RGB image array (H, W, 3)
        
        Returns:
        --------
            { MetricResult }     : Structured Color-domain metric result containing:
                                   - score      : Suspicion score [0.0, 1.0]
                                   - confidence : Reliability of color analysis evidence
                                   - details    : Color Analysis forensics and statistics
        """
        try:
            logger.debug(f"Running color analysis on image shape {image.shape}")
            
            # Normalize image to [0, 1]
            image_norm                           = self.image_processor.normalize_image(image = image)
            
            # Convert to HSV
            hsv                                  = self._rgb_to_hsv(rgb = image_norm)
            
            # Analyze saturation
            saturation_score, saturation_details = self._analyze_saturation(hsv = hsv)
            
            # Analyze color histogram
            histogram_score, histogram_details   = self._analyze_color_histogram(rgb = image_norm)
            
            # Analyze hue distribution
            hue_score, hue_details               = self._analyze_hue_distribution(hsv = hsv)
            
            # Combine scores
            weights                              = COLOR_ANALYSIS_PARAMS.MAIN_WEIGHTS
            final_score                          = (weights['saturation'] * saturation_score + weights['histogram'] * histogram_score + weights['hue'] * hue_score)

            # Calculate Confidence
            confidence                           = float(np.clip((abs(final_score - COLOR_ANALYSIS_PARAMS.NEUTRAL_SCORE) * 2.0), 0.0, 1.0))
            
            logger.debug(f"Color analysis: saturation={saturation_score:.3f}, histogram={histogram_score:.3f}, hue={hue_score:.3f}, Score={final_score:.3f}")
            
            return MetricResult(metric_type = MetricType.COLOR,
                                score       = float(final_score),
                                confidence  = confidence,
                                details     = {"saturation_stats" : saturation_details,
                                               "histogram_stats"  : histogram_details,
                                               "hue_stats"        : hue_details,
                                              },
                               )
            
        except Exception as e:
            logger.error(f"Color analysis failed: {e}")

            # Return neutral score on error
            return MetricResult(metric_type = MetricType.COLOR,
                                score       = COLOR_ANALYSIS_PARAMS.NEUTRAL_SCORE,
                                confidence  = 0.0,
                                details     = {"error": "color_analysis_failed"},
                               )


    def _rgb_to_hsv(self, rgb: np.ndarray) -> np.ndarray:
        """
        Convert RGB to HSV color space
        
        Arguments:
        ----------
            rgb { np.ndarray } : RGB image normalized to [0, 1]
        
        Returns:
        --------
            { np.ndarray }     : HSV image (H in [0, 360], S and V in [0, 1])
        """
        r, g, b = rgb[:, :, 0], rgb[:, :, 1], rgb[:, :, 2]
        
        maxc    = np.maximum(np.maximum(r, g), b)
        minc    = np.minimum(np.minimum(r, g), b)
        delta   = maxc - minc
        
        # Value
        v       = maxc
        
        # Saturation
        s       = np.where(maxc != 0, delta / maxc, 0)
        
        # Hue
        h       = np.zeros_like(maxc)
        
        # Red is max
        mask    = (maxc == r) & (delta != 0)
        h[mask] = 60 * (((g[mask] - b[mask]) / delta[mask]) % 6)
        
        # Green is max
        mask    = (maxc == g) & (delta != 0)
        h[mask] = 60 * (((b[mask] - r[mask]) / delta[mask]) + 2)
        
        # Blue is max
        mask    = (maxc == b) & (delta != 0)
        h[mask] = 60 * (((r[mask] - g[mask]) / delta[mask]) + 4)
        
        hsv     = np.stack([h, s, v], axis = 2)
        
        return hsv


    def _analyze_saturation(self, hsv: np.ndarray) -> tuple[float, dict]:
        """
        Analyze saturation distribution for anomalies
        
        Real photos: Most pixels have moderate saturation (0.2-0.7)
        AI images: Can have too many highly saturated pixels (>0.8)
        
        Arguments:
        ----------
            hsv { np.ndarray } : HSV image
        
        Returns:
        --------
            { tuple }          : A tuple containing:
                                 - Suspicion score [0.0, 1.0]
                                 - Saturation Stats
        """
        saturation          = hsv[:, :, 1]
        
        if (np.mean(saturation) < 0.05):
            logger.debug("Low global saturation; skipping saturation analysis")
            return COLOR_ANALYSIS_PARAMS.NEUTRAL_SCORE, {"reason": "insufficient_color_information"}

        # Compute saturation statistics
        mean_sat            = np.mean(saturation)
        high_sat_ratio      = np.mean(saturation > COLOR_ANALYSIS_PARAMS.SAT_HIGH_THRESHOLD)
        very_high_sat_ratio = np.mean(saturation > COLOR_ANALYSIS_PARAMS.SAT_VERY_HIGH_THRESHOLD)
        
        # Overall saturation level Analysis
        mean_anomaly        = 0.0
        
        if (mean_sat > COLOR_ANALYSIS_PARAMS.SAT_MEAN_THRESHOLD):
            mean_anomaly = min(1.0, (mean_sat - COLOR_ANALYSIS_PARAMS.SAT_MEAN_THRESHOLD) * COLOR_ANALYSIS_PARAMS.SAT_MEAN_SCALE)
        
        # High saturation pixels Analysis
        high_sat_anomaly = 0.0
        
        if (high_sat_ratio > COLOR_ANALYSIS_PARAMS.HIGH_SAT_RATIO_THRESHOLD):
            high_sat_anomaly = min(1.0, (high_sat_ratio - COLOR_ANALYSIS_PARAMS.HIGH_SAT_RATIO_THRESHOLD) * COLOR_ANALYSIS_PARAMS.HIGH_SAT_SCALE)

        # Very high saturation Analysis (clipping)
        clip_anomaly = 0.0
        
        if (very_high_sat_ratio > COLOR_ANALYSIS_PARAMS.CLIP_RATIO_THRESHOLD):
            clip_anomaly = min(1.0, (very_high_sat_ratio - COLOR_ANALYSIS_PARAMS.CLIP_RATIO_THRESHOLD) * COLOR_ANALYSIS_PARAMS.CLIP_SCALE)

        # Combine Scores
        weights          = COLOR_ANALYSIS_PARAMS.SAT_SUBMETRIC_WEIGHTS

        color_score      = (weights['mean_anomaly'] * mean_anomaly + weights['high_sat_anomaly'] * high_sat_anomaly + weights['clip_anomaly'] * clip_anomaly)

        final_score      = float(np.clip(color_score, 0.0, 1.0))

        saturation_stats = {"mean_saturation"     : float(mean_sat),
                            "high_sat_ratio"      : float(high_sat_ratio),
                            "very_high_sat_ratio" : float(very_high_sat_ratio),
                            "mean_anomaly"        : float(mean_anomaly),
                            "high_sat_anomaly"    : float(high_sat_anomaly),
                            "clip_anomaly"        : float(clip_anomaly),
                           }

        logger.debug(f"Saturation - mean: {mean_sat:.3f}, high_ratio: {high_sat_ratio:.3f}, clip_ratio: {very_high_sat_ratio:.3f}")
        
        return final_score, saturation_stats


    def _analyze_color_histogram(self, rgb: np.ndarray) -> tuple[float, dict]:
        """
        Analyze RGB histogram distributions for anomalies
        
        Arguments:
        ----------
            rgb { np.ndarray } : RGB image normalized to [0, 1]
        
        Returns:
        --------
            { tuple }          : A tuple containing:
                                 - Suspicion score [0.0, 1.0]
                                 - Histogram Analysis stats
        """
        anomalies      = list()
        roughness_vals = list()
        low_clip_vals  = list()
        high_clip_vals = list()
        
        for channel_idx, channel_name in enumerate(['R', 'G', 'B']):
            channel = rgb[:, :, channel_idx]
            
            # Compute histogram
            hist, bins = np.histogram(channel, 
                                      bins  = COLOR_ANALYSIS_PARAMS.HISTOGRAM_BINS, 
                                      range = COLOR_ANALYSIS_PARAMS.HISTOGRAM_RANGE,
                                     )

            hist       = hist / (np.sum(hist) + 1e-10)
            
            # Measure histogram roughness
            hist_diff  = np.abs(np.diff(hist))
            roughness  = np.mean(hist_diff)
            roughness_vals.append(roughness)

            # High roughness = suspicious
            if (roughness > COLOR_ANALYSIS_PARAMS.ROUGHNESS_THRESHOLD):
                anomalies.append(np.clip(((roughness - COLOR_ANALYSIS_PARAMS.ROUGHNESS_THRESHOLD) * COLOR_ANALYSIS_PARAMS.ROUGHNESS_SCALE), 0.0, 1.0))
            
            # Check for clipping (peaks at extremes)
            low_clip  = hist[0] + hist[1]
            high_clip = hist[-1] + hist[-2]

            # Append values to their respective storages
            low_clip_vals.append(low_clip)
            high_clip_vals.append(high_clip)
            
            if (low_clip > COLOR_ANALYSIS_PARAMS.CLIP_THRESHOLD):
                # More than 10% near black
                anomalies.append(min(1.0, (low_clip - COLOR_ANALYSIS_PARAMS.CLIP_THRESHOLD) * COLOR_ANALYSIS_PARAMS.CLIP_SCALE_FACTOR))

            if (high_clip > COLOR_ANALYSIS_PARAMS.CLIP_THRESHOLD):
                # More than 10% near white
                anomalies.append(min(1.0, (high_clip - COLOR_ANALYSIS_PARAMS.CLIP_THRESHOLD) * COLOR_ANALYSIS_PARAMS.CLIP_SCALE_FACTOR))

        if (len(anomalies) == 0):
            logger.debug("No color histogram anomalies detected")
            return COLOR_ANALYSIS_PARAMS.NEUTRAL_SCORE, {"reason": "insufficient_color_information"}
        
        # Take mean of detected anomalies
        score           = np.mean(anomalies)
        final_score     = float(np.clip(score, 0.0, 1.0))

        histogram_stats = {"roughness_mean"    : float(np.mean(roughness_vals)),
                           "low_clip_mean"     : float(np.mean(low_clip_vals)),
                           "high_clip_mean"    : float(np.mean(high_clip_vals)),
                           "channels_analyzed" : 3,
                          }
        
        return final_score, histogram_stats


    def _analyze_hue_distribution(self, hsv: np.ndarray) -> tuple[float, dict]:
        """
        Analyze hue distribution for unnatural patterns
        
        Arguments:
        ----------
            hsv { np.ndarray } : HSV image
        
        Returns:
        --------
            { tuple }          : A tuple containing:
                                 - Suspicion score [0.0, 1.0]
                                 - hue analysis stats
        """
        hue            = hsv[:, :, 0]
        saturation     = hsv[:, :, 1]
        
        # Only consider pixels with sufficient saturation (avoid gray)
        saturated_mask = saturation > COLOR_ANALYSIS_PARAMS.HUE_SAT_MASK_THRESHOLD
        
        if (np.sum(saturated_mask) < COLOR_ANALYSIS_PARAMS.HUE_MIN_PIXELS):
            # Not enough colored pixels to analyze
            return COLOR_ANALYSIS_PARAMS.NEUTRAL_SCORE, {"reason": "insufficient_color_information"}
        
        hue_saturated         = hue[saturated_mask]

        # Prevents false positives on monotone objects
        if (np.ptp(hue_saturated) < 5.0):
            logger.debug("Hue range too narrow; returning neutral score")
            return COLOR_ANALYSIS_PARAMS.NEUTRAL_SCORE, {"reason": "insufficient_color_information"}
        
        # Compute hue histogram
        hist, bins            = np.histogram(a     = hue_saturated, 
                                             bins  = COLOR_ANALYSIS_PARAMS.HUE_BINS, 
                                             range = COLOR_ANALYSIS_PARAMS.HUE_RANGE,
                                            )

        hist                  = hist / (np.sum(hist) + 1e-10)
        
        # Unnatural hue concentration Analysis
        sorted_hist           = np.sort(hist)[::-1]
        top3_concentration    = np.sum(sorted_hist[:3])
        concentration_anomaly = 0.0
         
        if (top3_concentration > COLOR_ANALYSIS_PARAMS.HUE_CONCENTRATION_THRESHOLD): 
            # More than 60% in 3 hue bins
            concentration_anomaly = min(1.0, (top3_concentration - COLOR_ANALYSIS_PARAMS.HUE_CONCENTRATION_THRESHOLD) * COLOR_ANALYSIS_PARAMS.HUE_CONCENTRATION_SCALE)
        
        # Hue gaps Analysis
        zero_bins             = np.sum(hist < COLOR_ANALYSIS_PARAMS.HUE_EMPTY_BIN_THRESHOLD)
        gap_ratio             = zero_bins / len(hist)
        gap_anomaly           = 0.0
        
        if (gap_ratio > COLOR_ANALYSIS_PARAMS.HUE_GAP_RATIO_THRESHOLD):  
            # More than 40% empty bins
            gap_anomaly = min(1.0, (gap_ratio - COLOR_ANALYSIS_PARAMS.HUE_GAP_RATIO_THRESHOLD) * COLOR_ANALYSIS_PARAMS.HUE_GAP_SCALE)
        
        weights               = COLOR_ANALYSIS_PARAMS.HUE_SUBMETRIC_WEIGHTS
        score                 = (weights['concentration_anomaly'] * concentration_anomaly + weights['gap_anomaly'] * gap_anomaly)
        final_score           = float(np.clip(score, 0.0, 1.0))

        hue_stats             = {"top3_concentration"    : float(top3_concentration),
                                 "gap_ratio"             : float(gap_ratio),
                                 "concentration_anomaly" : float(concentration_anomaly),
                                 "gap_anomaly"           : float(gap_anomaly),
                                }

        logger.debug(f"Hue - concentration: {top3_concentration:.3f}, gap_ratio: {gap_ratio:.3f}")
        
        return final_score, hue_stats