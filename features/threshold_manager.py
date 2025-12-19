# Dependencies
from typing import Dict
from utils.logger import get_logger
from config.settings import settings
from config.constants import MetricType
from config.constants import SignalStatus
from config.constants import SIGNAL_THRESHOLDS


# Setup Logging
logger = get_logger(__name__)


class ThresholdManager:
    """
    Manage detection thresholds dynamically
    
    Purpose:
    --------
    Allows runtime adjustment of detection thresholds for:
    - A/B testing different sensitivity levels
    - Calibration based on real-world performance
    - Custom thresholds for specific use cases
    - Environment-specific tuning (production vs staging)
    
    Note: Changes are runtime-only and not persisted
    """
    def __init__(self):
        """
        Initialize Threshold Manager with current settings
        """
        self._review_threshold  = settings.REVIEW_THRESHOLD
        self._signal_thresholds = dict(SIGNAL_THRESHOLDS)
        self._metric_weights    = dict(settings.get_metric_weights())
        
        logger.info(f"ThresholdManager initialized: review_threshold={self._review_threshold}")
    

    def get_review_threshold(self) -> float:
        """
        Get current review threshold
        
        Returns:
        --------
            { float } : Current threshold [0.0, 1.0]
        """
        return self._review_threshold
    

    def set_review_threshold(self, new_threshold: float) -> bool:
        """
        Set new review threshold
        
        Arguments:
        ----------
            new_threshold { float } : New threshold value [0.0, 1.0]
        
        Returns:
        --------
            { bool }                : Success status
        """
        if not (0.0 <= new_threshold <= 1.0):
            logger.error(f"Invalid threshold: {new_threshold} (must be between 0.0 and 1.0)")
            return False
        
        old_threshold          = self._review_threshold
        self._review_threshold = new_threshold
        
        logger.info(f"Review threshold changed: {old_threshold:.2f} → {new_threshold:.2f}")
        
        return True
    

    def adjust_sensitivity(self, sensitivity: str) -> bool:
        """
        Adjust sensitivity using preset levels
        
        Arguments:
        ----------
            sensitivity { str } : One of 'conservative', 'balanced', 'aggressive'
        
        Returns:
        --------
            { bool }            : Success status
        """
        presets = {'conservative' : 0.75,  # Fewer false positives, may miss some AI
                   'balanced'     : 0.65,  # Recommended default
                   'aggressive'   : 0.55,  # Catch more AI, more false positives
                  }
        
        if (sensitivity not in presets):
            logger.error(f"Invalid sensitivity: {sensitivity}. Must be one of {list(presets.keys())}")
            return False
        
        new_threshold = presets[sensitivity]
        success       = self.set_review_threshold(new_threshold = new_threshold)
        
        if success:
            logger.info(f"Sensitivity set to '{sensitivity}' (threshold={new_threshold})")
        
        return success
    

    def get_signal_thresholds(self) -> Dict[SignalStatus, float]:
        """
        Get current signal thresholds
        
        Returns:
        --------
            { dict } : Signal status → threshold mapping
        """
        return self._signal_thresholds.copy()
    

    def set_signal_threshold(self, status: SignalStatus, threshold: float) -> bool:
        """
        Set threshold for specific signal status
        
        Arguments:
        ----------
            status    { SignalStatus } : Signal status to modify
            
            threshold { float }        : New threshold [0.0, 1.0]
        
        Returns:
        --------
            { bool }                   : Success status
        """
        if not (0.0 <= threshold <= 1.0):
            logger.error(f"Invalid threshold: {threshold}")
            return False
        
        old_threshold                    = self._signal_thresholds.get(status)
        self._signal_thresholds[status]  = threshold
        
        logger.info(f"Signal threshold for {status.value}: {old_threshold:.2f} → {threshold:.2f}")
        
        return True
    

    def get_metric_weights(self) -> Dict[MetricType, float]:
        """
        Get current metric weights
        
        Returns:
        --------
            { dict } : Metric type → weight mapping
        """
        return self._metric_weights.copy()
    

    def set_metric_weight(self, metric: MetricType, weight: float) -> bool:
        """
        Set weight for specific metric
        
        Arguments:
        ----------
            metric { MetricType } : Metric to modify
            
            weight   { float }    : New weight [0.0, 1.0]
        
        Returns:
        --------
            { bool }              : Success status
        """
        if not (0.0 <= weight <= 1.0):
            logger.error(f"Invalid weight: {weight}")
            return False
        
        old_weight                   = self._metric_weights.get(metric, 0.0)
        self._metric_weights[metric] = weight
        
        # Validate total weight
        total_weight                 = sum(self._metric_weights.values())
        
        if not (0.99 <= total_weight <= 1.01):
            logger.warning(f"Total metric weights = {total_weight:.3f} (should sum to 1.0)")
        
        logger.info(f"Metric weight for {metric.value}: {old_weight:.2f} → {weight:.2f}")
        
        return True
    

    def set_all_metric_weights(self, weights: Dict[MetricType, float]) -> bool:
        """
        Set all metric weights at once (ensures sum = 1.0)
        
        Arguments:
        ----------
            weights { dict } : Complete metric weights mapping
        
        Returns:
        --------
            { bool }         : Success status
        """
        # Validate input
        if (not all(0.0 <= w <= 1.0 for w in weights.values())):
            logger.error("All weights must be between 0.0 and 1.0")
            return False
        
        total_weight         = sum(weights.values())
        
        if not (0.99 <= total_weight <= 1.01):
            logger.error(f"Weights must sum to 1.0, got {total_weight:.3f}")
            return False
        
        self._metric_weights = dict(weights)
        
        logger.info(f"All metric weights updated: {self._metric_weights}")
        
        return True
    

    def get_recommendations(self, score: float) -> Dict[str, str]:
        """
        Get action recommendations based on score
        
        Arguments:
        ----------
            score { float } : Overall suspicion score [0.0, 1.0]
        
        Returns:
        --------
            { dict }        : Recommendation details
        """
        if (score >= 0.85):
            return {"priority"   : "HIGH",
                    "action"     : "Immediate manual verification recommended",
                    "confidence" : "Very high likelihood of AI generation",
                    "next_steps" : "Forensic analysis, reverse image search, metadata inspection",
                   }
        
        elif (score >= 0.70):
            return {"priority"   : "MEDIUM",
                    "action"     : "Manual verification recommended",
                    "confidence" : "High likelihood of AI generation",
                    "next_steps" : "Visual inspection, compare with similar authentic images",
                   }
        
        elif (score >= 0.50):
            return {"priority"   : "LOW",
                    "action"     : "Optional review",
                    "confidence" : "Moderate indicators of AI generation",
                    "next_steps" : "May be heavily edited real photo, check source",
                   }
        
        else:
            return {"priority"   : "NONE",
                    "action"     : "No immediate action needed",
                    "confidence" : "Low likelihood of AI generation",
                    "next_steps" : "Likely authentic, proceed normally",
                   }
    

    def get_current_config(self) -> Dict[str, object]:
        """
        Get complete current configuration
        
        Returns:
        --------
            { dict } : All current threshold and weight settings
        """
        return {"review_threshold"  : self._review_threshold,
                "signal_thresholds" : self._signal_thresholds.copy(),
                "metric_weights"    : self._metric_weights.copy(),
               }
    

    def reset_to_defaults(self) -> None:
        """
        Reset all thresholds to default settings
        """
        self._review_threshold  = settings.REVIEW_THRESHOLD
        self._signal_thresholds = dict(SIGNAL_THRESHOLDS)
        self._metric_weights    = dict(settings.get_metric_weights())
        
        logger.info("All thresholds reset to default values")