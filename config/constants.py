# Dependencies
from enum import Enum
from dataclasses import dataclass


class DetectionStatus(str, Enum):
    """
    Overall detection status
    """
    LIKELY_AUTHENTIC = "LIKELY_AUTHENTIC"
    REVIEW_REQUIRED  = "REVIEW_REQUIRED"


class SignalStatus(str, Enum):
    """
    Individual signal status
    """
    PASSED  = "passed"
    WARNING = "warning"
    FLAGGED = "flagged"


class FileFormat(str, Enum):
    """
    Supported file formats
    """
    JPG  = ".jpg"
    JPEG = ".jpeg"
    PNG  = ".png"
    WEBP = ".webp"


class MetricType(str, Enum):
    """
    Detection metric types
    """
    GRADIENT  = "gradient"
    FREQUENCY = "frequency"
    NOISE     = "noise"
    TEXTURE   = "texture"
    COLOR     = "color"



# Signal thresholds
SIGNAL_THRESHOLDS          = {SignalStatus.FLAGGED : 0.7,
                              SignalStatus.WARNING : 0.4,
                              SignalStatus.PASSED  : 0.0,
                             }

# Metric explanations
METRIC_EXPLANATIONS        = {MetricType.GRADIENT  : {'high'     : "Detected irregular gradient patterns typical of diffusion models. Natural photos show consistent lighting gradients shaped by physics.",
                                                      'moderate' : "Some gradient inconsistencies detected. May indicate AI generation or heavy editing.",
                                                      'normal'   : "Gradient patterns are consistent with natural lighting and camera optics."
                                                     },
                              MetricType.FREQUENCY : {'high'     : "Unusual frequency distribution detected. AI-generated images often show unnatural spectral patterns.",
                                                      'moderate' : "Frequency patterns show some irregularities. Requires further review.",
                                                      'normal'   : "Frequency distribution matches expected patterns for authentic photographs."
                                                     },
                              MetricType.NOISE     : {'high'     : "Noise pattern is unnaturally uniform. Real camera sensors produce characteristic noise patterns.",
                                                      'moderate' : "Noise distribution shows some anomalies. May indicate synthetic generation.",
                                                      'normal'   : "Noise characteristics are consistent with genuine camera sensor behavior."
                                                     },
                              MetricType.TEXTURE   : {'high'     : "Detected suspiciously smooth regions. Natural photos have organic texture variation.",
                                                      'moderate' : "Some texture regions appear overly uniform. Further analysis recommended.",
                                                      'normal'   : "Texture variation is within expected ranges for authentic photographs."
                                                     },
                              MetricType.COLOR     : {'high'     : "Color distribution shows impossible or highly unlikely patterns.",
                                                      'moderate' : "Some color histogram irregularities detected.",
                                                      'normal'   : "Color distribution is within normal ranges for real photographs."
                                                     }
                             }

# Basic Image Processing Constants
MIN_IMAGE_DIMENSION        = 64
MAX_IMAGE_DIMENSION        = 8192
LUMINANCE_WEIGHTS          = (0.2126, 0.7152, 0.0722)  # ITU-R BT.709
IMAGE_RESIZE_MAX_DIMENSION = 1024


# Gradient-Field PCA Detection Parameters
@dataclass(frozen = True)
class GradientFieldPCAParams:
    """
    Parameters for Gradient-Field PCA detection
    """
    # Random Seed For Reproducibility 
    RANDOM_SEED                : int   = 1234

    # NEUTRAL_SCORE
    NEUTRAL_SCORE              : float = 0.5

    # PCA Configuration
    SAMPLE_SIZE                : int   = 10000  # Max gradient samples for PCA
    
    # Thresholds
    MAGNITUDE_THRESHOLD        : float = 1e-6   # Minimum gradient magnitude to consider
    MIN_SAMPLES                : int   = 10     # Minimum samples required for PCA
    VARIANCE_THRESHOLD         : float = 1e-10  # Minimum total variance
    EIGENVALUE_RATIO_THRESHOLD : float = 0.85   # Real photos typically > 0.85



# Frequency Analysis Parameters
@dataclass(frozen = True)
class FrequencyAnalysisParams:
    """
    Parameters for FFT-based frequency analysis
    """
    # NEUTRAL_SCORE
    NEUTRAL_SCORE       : float = 0.5

    # FFT Configuration
    BINS                : int   = 64
    HIGH_FREQ_THRESHOLD : float = 0.6     # Radial position where high-freq starts
    
    # Analysis Thresholds
    MIN_SPECTRUM_SAMPLES : int   = 10
    HF_RATIO_UPPER       : float = 0.35   # High-frequency ratio upper bound
    HF_RATIO_LOWER       : float = 0.08   # High-frequency ratio lower bound
    
    # Scaling Factors
    HF_UPPER_SCALE       : float = 3.0
    HF_LOWER_SCALE       : float = 5.0
    ROUGHNESS_SCALE      : float = 10.0
    DEVIATION_SCALE      : float = 2.0
    
    # Sub-metric Weights
    SUBMETRIC_WEIGHTS    : dict  = None
    
    def __post_init__(self):
        if self.SUBMETRIC_WEIGHTS is None:
            object.__setattr__(self, 'SUBMETRIC_WEIGHTS', {'hf_anomaly' : 0.4,
                                                           'roughness'  : 0.3,
                                                           'deviation'  : 0.3,
                                                          }
                              )


# Noise Analysis Parameters
@dataclass(frozen = True)
class NoiseAnalysisParams:
    """
    Parameters for noise pattern analysis
    """
    # NEUTRAL SCORE 
    NEUTRAL_SCORE            : float = 0.5

    # Patch Configuration
    PATCH_SIZE               : int   = 32
    STRIDE                   : int   = 16
    SAMPLES                  : int   = 100
    
    # Variance Thresholds
    VARIANCE_LOW_THRESHOLD   : float = 1.0     # Skip too uniform patches
    VARIANCE_HIGH_THRESHOLD  : float = 1000.0  # Skip too structured patches
    
    # MAD Conversion
    MAD_TO_STD_FACTOR        : float = 1.4826  # Gaussian: σ ≈ 1.4826 × MAD
    
    # Distribution Analysis
    MIN_ESTIMATES            : int   = 10
    MIN_FILTERED_SAMPLES     : int   = 5
    OUTLIER_PERCENTILE_LOW   : int   = 10
    OUTLIER_PERCENTILE_HIGH  : int   = 90
    
    # CV (Coefficient of Variation) Thresholds
    CV_UNIFORM_THRESHOLD     : float = 0.15
    CV_VARIABLE_THRESHOLD    : float = 1.2
    CV_UNIFORM_SCALE         : float = 5.0
    CV_VARIABLE_SCALE        : float = 2.0
    
    # Noise Level Thresholds
    LEVEL_CLEAN_THRESHOLD    : float = 1.5
    LEVEL_LOW_THRESHOLD      : float = 2.5
    
    # IQR Analysis
    IQR_THRESHOLD            : float = 0.3
    IQR_SCALE                : float = 2.0
    IQR_PERCENTILE_LOW       : int   = 25
    IQR_PERCENTILE_HIGH      : int   = 75
    
    # Sub-metric Weights
    SUBMETRIC_WEIGHTS        : dict  = None
    
    def __post_init__(self):
        if self.SUBMETRIC_WEIGHTS is None:
            object.__setattr__(self, 'SUBMETRIC_WEIGHTS', {'cv_anomaly'          : 0.4,
                                                           'noise_level_anomaly' : 0.4,
                                                           'iqr_anomaly'         : 0.2,
                                                          }
                              )


# Texture Analysis Parameters
@dataclass(frozen = True)
class TextureAnalysisParams:
    """
    Parameters for texture analysis
    """
    # Random Seed for reproducibility
    RANDOM_SEED                : int   = 1234

    # Neutral Score 
    NEUTRAL_SCORE              : float = 0.5

    # Patch Configuration
    PATCH_SIZE                 : int   = 64
    N_PATCHES                  : int   = 50
    
    # Histogram Configuration
    HISTOGRAM_BINS             : int   = 32
    HISTOGRAM_RANGE            : tuple = (0, 255)
    
    # Edge Detection
    EDGE_THRESHOLD             : float = 10.0
    
    # Smoothness Analysis
    SMOOTHNESS_THRESHOLD       : float = 0.5
    SMOOTH_RATIO_THRESHOLD     : float = 0.4
    SMOOTH_RATIO_SCALE         : float = 2.5
    
    # Entropy Analysis
    ENTROPY_CV_THRESHOLD       : float = 0.15
    ENTROPY_SCALE              : float = 5.0
    
    # Contrast Analysis
    CONTRAST_CV_LOW            : float = 0.3
    CONTRAST_CV_HIGH           : float = 1.5
    CONTRAST_LOW_SCALE         : float = 2.0
    CONTRAST_HIGH_SCALE        : float = 0.5
    
    # Edge Density Analysis
    EDGE_CV_THRESHOLD          : float = 0.4
    EDGE_SCALE                 : float = 1.5
    
    # Sub-metric Weights
    SUBMETRIC_WEIGHTS          : dict  = None
    
    def __post_init__(self):
        if self.SUBMETRIC_WEIGHTS is None:
            object.__setattr__(self, 'SUBMETRIC_WEIGHTS', {'smoothness_anomaly' : 0.35,
                                                           'entropy_anomaly'    : 0.25,
                                                           'contrast_anomaly'   : 0.25,
                                                           'edge_anomaly'       : 0.15,
                                                          }
                              )


# Color Analysis Parameters
@dataclass(frozen = True)
class ColorAnalysisParams:
    """
    Parameters for color distribution analysis
    """
    # Random Seed for reproducibility
    RANDOM_SEED                  : int   = 1234

    # Neutral Score 
    NEUTRAL_SCORE                : float = 0.5
    # Saturation Analysis
    SAT_HIGH_THRESHOLD           : float = 0.8
    SAT_VERY_HIGH_THRESHOLD      : float = 0.95
    SAT_MEAN_THRESHOLD           : float = 0.65
    SAT_MEAN_SCALE               : float = 3.0
    HIGH_SAT_RATIO_THRESHOLD     : float = 0.20
    HIGH_SAT_SCALE               : float = 2.5
    CLIP_RATIO_THRESHOLD         : float = 0.05
    CLIP_SCALE                   : float = 10.0
    
    # Histogram Analysis
    HISTOGRAM_BINS               : int   = 64
    HISTOGRAM_RANGE              : tuple = (0, 1)
    ROUGHNESS_THRESHOLD          : float = 0.015
    ROUGHNESS_SCALE              : float = 50.0
    CLIP_THRESHOLD               : float = 0.10
    CLIP_SCALE_FACTOR            : float = 5.0
    
    # Hue Analysis
    HUE_SAT_MASK_THRESHOLD       : float = 0.2
    HUE_MIN_PIXELS               : int   = 100
    HUE_BINS                     : int   = 36
    HUE_RANGE                    : tuple = (0, 360)
    HUE_CONCENTRATION_THRESHOLD  : float = 0.6
    HUE_CONCENTRATION_SCALE      : float = 2.5
    HUE_EMPTY_BIN_THRESHOLD      : float = 0.01
    HUE_GAP_RATIO_THRESHOLD      : float = 0.4
    HUE_GAP_SCALE                : float = 1.5
    
    # Sub-metric Weights
    SAT_SUBMETRIC_WEIGHTS        : dict  = None
    HUE_SUBMETRIC_WEIGHTS        : dict  = None
    MAIN_WEIGHTS                 : dict  = None
    
    def __post_init__(self):
        if self.SAT_SUBMETRIC_WEIGHTS is None:
            object.__setattr__(self, 'SAT_SUBMETRIC_WEIGHTS', {'mean_anomaly'     : 0.3,
                                                               'high_sat_anomaly' : 0.4,
                                                               'clip_anomaly'     : 0.3,
                                                              }
                              )

        if self.HUE_SUBMETRIC_WEIGHTS is None:
            object.__setattr__(self, 'HUE_SUBMETRIC_WEIGHTS', {'concentration_anomaly' : 0.6,
                                                               'gap_anomaly'           : 0.4,
                                                              }
                              )

        if self.MAIN_WEIGHTS is None:
            object.__setattr__(self, 'MAIN_WEIGHTS', {'saturation' : 0.4,
                                                      'histogram'  : 0.35,
                                                      'hue'        : 0.25,
                                                     }
                              )



# Singleton instances for parameter classes
GRADIENT_FIELD_PCA_PARAMS = GradientFieldPCAParams()
FREQUENCY_ANALYSIS_PARAMS = FrequencyAnalysisParams()
NOISE_ANALYSIS_PARAMS     = NoiseAnalysisParams()
TEXTURE_ANALYSIS_PARAMS   = TextureAnalysisParams()
COLOR_ANALYSIS_PARAMS     = ColorAnalysisParams()


