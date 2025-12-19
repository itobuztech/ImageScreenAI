# Dependencies
import time
import numpy as np
from typing import List
from pathlib import Path
from types import MappingProxyType
from utils.logger import get_logger
from config.settings import settings
from config.schemas import MetricResult
from config.constants import MetricType
from config.constants import SignalStatus
from config.schemas import AnalysisResult
from config.schemas import DetectionSignal
from config.constants import DetectionStatus
from config.constants import SIGNAL_THRESHOLDS
from utils.image_processor import ImageProcessor
from config.constants import METRIC_EXPLANATIONS
from metrics.noise_analyzer import NoiseAnalyzer
from metrics.color_analyzer import ColorAnalyzer
from metrics.texture_analyzer import TextureAnalyzer
from features.threshold_manager import ThresholdManager
from config.constants import IMAGE_RESIZE_MAX_DIMENSION
from metrics.frequency_analyzer import FrequencyAnalyzer
from metrics.gradient_field_pca import GradientFieldPCADetector

# Suppress NumPy warning 
np.seterr(divide  = 'ignore', 
          invalid = 'ignore',
         )

         
# Setup Logging
logger = get_logger(__name__)


class MetricsAggregator:
    """
    Main detector that orchestrates all detection methods

    Combines multiple unsupervised metrics:
    ----------------------------------------
    1. Gradient-Field PCA
    2. Frequency Domain Analysis (FFT)
    3. Noise Pattern Analysis
    4. Texture Analysis
    5. Color Distribution Analysis

    Note: Each metric produces a suspicion score [0.0, 1.0] : scores are combined using weighted average to produce final assessment
    """
    def __init__(self, threshold_manager: ThresholdManager | None = None):
        """
        Initialize all detectors
        """
        logger.info("Initializing AI Image Detector")

        # Optional runtime threshold manager
        self.threshold_manager           = threshold_manager
        
        self.gradient_field_pca_detector = GradientFieldPCADetector()
        self.frequency_analyzer          = FrequencyAnalyzer()
        self.noise_analyzer              = NoiseAnalyzer()
        self.texture_analyzer            = TextureAnalyzer()
        self.color_analyzer              = ColorAnalyzer()
        self.image_processor             = ImageProcessor()

        # Create detector registry
        self.detector_registry          = MappingProxyType({MetricType.GRADIENT  : ("Gradient Field PCA", self.gradient_field_pca_detector),
                                                            MetricType.FREQUENCY : ("Frequency Analysis", self.frequency_analyzer),
                                                            MetricType.NOISE     : ("Noise Analysis", self.noise_analyzer),
                                                            MetricType.TEXTURE   : ("Texture Analysis", self.texture_analyzer),
                                                            MetricType.COLOR     : ("Color Analysis", self.color_analyzer),
                                                          })
        
        # Get metric weights either from runtime UI or default to settings
        self.weights                    = (self.threshold_manager.get_metric_weights() if self.threshold_manager else settings.get_metric_weights())
        
        logger.info(f"Metric weights: {self.weights}")


    def analyze_image(self, image_path: Path, filename: str, image_size: tuple) -> AnalysisResult:
        """
        Analyze single image for AI generation
        
        Arguments:
        ----------
            image_path { Path }  : Path to image file
            
            filename   { str }   : Original filename
            
            image_size { tuple } : (width, height) tuple
        
        Returns:
        --------
            { AnalysisResult }   : AnalysisResult with detection outcome
        """
        logger.info(f"Analyzing image: {filename}")
        
        start_time = time.time()
        
        try:
            # Load image
            image           = self.image_processor.load_image(file_path = image_path)
            
            # Resize if needed for performance
            image           = self.image_processor.resize_if_needed(image         = image, 
                                                                    max_dimension = IMAGE_RESIZE_MAX_DIMENSION,
                                                                   )
            
            # Run all detectors and get raw scores
            metric_results  = self._run_all_detectors(image = image)
            
            # Create signals from scores (aggregator's responsibility)
            signals         = self._create_signals_from_scores(metric_results = metric_results)
            
            # Aggregate results
            overall_score   = self._aggregate_scores(metric_results = metric_results)
            
            # Determine status
            status          = self._determine_status(overall_score = overall_score)
            
            # Calculate processing time
            processing_time = time.time() - start_time
            
            # Create result
            result          = AnalysisResult(filename        = filename,
                                             overall_score   = overall_score,
                                             status          = status,
                                             confidence      = int(overall_score * 100),
                                             signals         = signals,
                                             metric_results  = metric_results,
                                             processing_time = processing_time,
                                             image_size      = image_size,
                                            )
            
            logger.info(f"Analysis complete for {filename}: status={status.value}, score={overall_score:.3f}, time={processing_time:.2f}s")
            
            return result
            
        except Exception as e:
            logger.error(f"Analysis failed for {filename}: {e}")
            raise


    def _run_all_detectors(self, image: np.ndarray) -> dict[MetricType, MetricResult]:
        """
        Run all detection methods and collect raw scores
        
        Arguments:
        ----------
            image { np.ndarray } : RGB image array
        
        Returns:
        --------
                  { dict }       : Dictionary mapping MetricType to MetricResult
        """
        metric_results = dict()
        
        # Run eaach detector one by one
        for metric_type, (detector_name, detector) in self.detector_registry.items():
            try:
                result                      = detector.detect(image = image)
                result.metric_type          = metric_type
                metric_results[metric_type] = result
                
                logger.debug(f"{detector_name} | {metric_type.value} | score={result.score:.3f} | confidence={result.confidence:.3f}")
                
            except Exception as e:
                logger.error(f"{detector.__class__.__name__} failed: {e}")

                # Same Failure Score by all metrics with same confidence 
                metric_results[metric_type] = MetricResult(metric_type = metric_type,
                                                           score       = settings.REVIEW_THRESHOLD,
                                                           confidence  = 0.0,
                                                           details     = {"error": "detector_failed"},
                                                          )
        
        return metric_results


    def _create_signals_from_scores(self, metric_results: dict) -> List[DetectionSignal]:
        """
        Convert MetricResults to DetectionSignals with status and explanations
        
        This is the aggregator's responsibility - metrics don't know about signals
        
        Arguments:
        ----------
            metric_results { dict }   : Dictionary mapping MetricType to float score
        
        Returns:
        --------
                    { list }          : List of complete detection signals
        """
        signals           = list()

        signal_thresholds = (self.threshold_manager.get_signal_thresholds() if self.threshold_manager else SIGNAL_THRESHOLDS)
        
        for metric_type, result in metric_results.items():
            # Extract score of the metric
            score = result.score

            # Determine status based on thresholds
            if (score >= signal_thresholds[SignalStatus.FLAGGED]):
                status   = SignalStatus.FLAGGED
                severity = 'high'
                
            elif (score >= signal_thresholds[SignalStatus.WARNING]):
                status   = SignalStatus.WARNING
                severity = 'moderate'
                
            else:
                status   = SignalStatus.PASSED
                severity = 'normal'
            
            # Get explanation from constants
            explanation = METRIC_EXPLANATIONS[metric_type][severity]
            
            # Create signal
            signal      = DetectionSignal(name        = self.detector_registry[metric_type][0],
                                          metric_type = metric_type,
                                          score       = score,
                                          status      = status,
                                          explanation = explanation,
                                         )
            
            signals.append(signal)
        
        # Sort signals by score (highest first)
        signals.sort(key = lambda s: s.score, reverse = True)
        
        return signals


    def _aggregate_scores(self, metric_results: dict) -> float:
        """
        Aggregate individual metric scores using weighted average
        
        Arguments:
        ----------
            metric_results { dict } : Dictionary mapping MetricType to float score
        
        Returns:
        --------
                { float }           : Overall suspicion score [0.0, 1.0]
        """
        total_score  = 0.0
        total_weight = 0.0
        
        for metric_type, result in metric_results.items():
            weight        = self.weights.get(metric_type, 0.0)
            total_score  += result.score * weight
            total_weight += weight
        
        
        # Get Aggregated Score
        if (total_weight > 0):
            # Normalize
            overall_score = total_score / total_weight
        
        else:
            # Neutral if no valid weights
            overall_score = 0.5  
        
        logger.debug(f"Aggregated score: {overall_score:.3f}")
        
        return float(np.clip(overall_score, 0.0, 1.0))


    def _determine_status(self, overall_score: float) -> DetectionStatus:
        """
        Determine binary status from overall score
        
        Arguments:
        ----------
            overall_score { float } : Aggregated suspicion score
        
        Returns:
        --------
            { DetectionStatus }     : LIKELY_AUTHENTIC or REVIEW_REQUIRED
        """
        # Extract review threshold either from threshold_manager or deault to settings value
        review_threshold = (self.threshold_manager.get_review_threshold() if self.threshold_manager else settings.REVIEW_THRESHOLD)

        if (overall_score >= review_threshold):
            return DetectionStatus.REVIEW_REQUIRED
        
        else:
            return DetectionStatus.LIKELY_AUTHENTIC