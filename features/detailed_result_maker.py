# Dependencies
import pandas as pd
from typing import Dict
from typing import List
from typing import Optional
from utils.logger import get_logger
from config.constants import MetricType
from config.constants import SignalStatus
from config.schemas import AnalysisResult
from config.constants import SIGNAL_THRESHOLDS


# Setup Logging
logger = get_logger(__name__)


class DetailedResultMaker:
    """
    Extract and format detailed analysis results for UI and reporting
    
    Purpose:
    --------
    - Extracts all intermediate metrics from MetricResult objects
    - Formats data for tabular display in UI
    - Provides rich metadata for PDF/CSV reports
    - No re-computation - just data extraction and formatting
    
    Output Formats:
    ---------------
    1. Structured dictionaries for UI
    2. Pandas DataFrames for reports
    3. Hierarchical JSON for API 
    """
    def __init__(self, signal_thresholds: dict | None = None):
        """
        Initialize Detailed Result Maker
        """
        self.metric_display_names = {MetricType.GRADIENT  : "Gradient-Field PCA",
                                     MetricType.FREQUENCY : "Frequency Domain (FFT)",
                                     MetricType.NOISE     : "Noise Pattern Analysis",
                                     MetricType.TEXTURE   : "Texture Statistics",
                                     MetricType.COLOR     : "Color Distribution",
                                    }

        self.signal_thresholds    = signal_thresholds or SIGNAL_THRESHOLDS
        
        logger.debug("DetailedResultMaker initialized")
    

    def extract_detailed_results(self, analysis_result: AnalysisResult) -> Dict:
        """
        Extract all detailed results from AnalysisResult
        
        Arguments:
        ----------
            analysis_result { AnalysisResult } : Complete analysis result
        
        Returns:
        --------
            { dict }                           : Comprehensive detailed results
        """
        logger.debug(f"Extracting detailed results for: {analysis_result.filename}")
        
        detailed = {"filename"         : analysis_result.filename,
                    "overall_summary"  : self._extract_overall_summary(analysis_result = analysis_result),
                    "metrics_detailed" : self._extract_all_metrics(analysis_result = analysis_result),
                    "metadata"         : self._extract_metadata(analysis_result = analysis_result),
                   }
        
        logger.debug(f"Extracted {len(detailed['metrics_detailed'])} metric details")
        
        return detailed
    

    def create_detailed_table(self, analysis_result: AnalysisResult) -> pd.DataFrame:
        """
        Create detailed metrics table as DataFrame
        
        Arguments:
        ----------
            analysis_result { AnalysisResult } : Complete analysis result
        
        Returns:
        --------
            { DataFrame }                      : Tabular detailed results
        """
        rows = list()
        
        for metric_type, metric_result in analysis_result.metric_results.items():
            display_name = self.metric_display_names.get(metric_type, metric_type.value)
            
            row          = {"Metric"      : display_name,
                            "Score"       : round(metric_result.score, 3),
                            "Confidence"  : round(metric_result.confidence, 3) if metric_result.confidence is not None else "N/A",
                            "Status"      : self._score_to_status(score = metric_result.score),
                           }
            
            # Add key details from each metric
            details      = self._extract_key_details(metric_type   = metric_type,
                                                     metric_result = metric_result,
                                                    )
            
            row.update(details)
            rows.append(row)
        
        # Dump rows into a pandas dataframe for structured result
        dataframe = pd.DataFrame(data = rows)
        
        logger.debug(f"Created detailed table with {len(dataframe)} rows, {len(dataframe.columns)} columns")
        
        return dataframe
    

    def create_report_data(self, analysis_result: AnalysisResult) -> Dict:
        """
        Create rich data structure for report generation
        
        Arguments:
        ----------
            analysis_result { AnalysisResult } : Complete analysis result
        
        Returns:
        --------
            { dict }                           : Report-ready data structure
        """
        report_data = {"header"             : self._create_report_header(analysis_result = analysis_result),
                       "overall_assessment" : self._create_overall_assessment(analysis_result = analysis_result),
                       "metric_breakdown"   : self._create_metric_breakdown(analysis_result = analysis_result),
                       "forensic_details"   : self._create_forensic_details(analysis_result = analysis_result),
                       "recommendations"    : self._create_recommendations(analysis_result = analysis_result),
                      }
        
        logger.debug(f"Created report data for: {analysis_result.filename}")
        
        return report_data
    

    def _extract_overall_summary(self, analysis_result: AnalysisResult) -> Dict:
        """
        Extract overall summary information
        """
        timestamp = getattr(analysis_result, "timestamp", None)

        return {"filename"        : analysis_result.filename,
                "status"          : analysis_result.status.value,
                "overall_score"   : round(analysis_result.overall_score, 3),
                "confidence"      : analysis_result.confidence,
                "processing_time" : round(analysis_result.processing_time, 2),
                "image_size"      : f"{analysis_result.image_size[0]}×{analysis_result.image_size[1]}",
                "timestamp"       : timestamp.isoformat() if timestamp else None,
               }
    

    def _extract_all_metrics(self, analysis_result: AnalysisResult) -> List[Dict]:
        """
        Extract detailed information for all metrics
        """
        metrics_detailed = list()
        
        for metric_type, metric_result in analysis_result.metric_results.items():
            metric_detail = {"metric_type"    : metric_type.value,
                             "display_name"   : self.metric_display_names.get(metric_type, metric_type.value),
                             "score"          : round(metric_result.score, 3),
                             "confidence"     : round(metric_result.confidence, 3) if metric_result.confidence is not None else None,
                             "status"         : self._score_to_status(score = metric_result.score),
                             "details"        : metric_result.details or {},
                             "interpretation" : self._interpret_metric(metric_type   = metric_type,
                                                                       metric_result = metric_result,
                                                                      ),
                            }
            
            metrics_detailed.append(metric_detail)
        
        # Sort by score (highest first)
        metrics_detailed.sort(key = lambda x: x['score'], reverse = True)
        
        return metrics_detailed
    

    def _extract_metadata(self, analysis_result: AnalysisResult) -> Dict:
        """
        Extract processing metadata
        """
        return {"total_metrics"   : len(analysis_result.metric_results),
                "flagged_metrics" : sum(1 for s in analysis_result.signals if s.status.value == 'flagged'),
                "warning_metrics" : sum(1 for s in analysis_result.signals if s.status.value == 'warning'),
                "passed_metrics"  : sum(1 for s in analysis_result.signals if s.status.value == 'passed'),
                "avg_confidence"  : self._calculate_avg_confidence(analysis_result = analysis_result),
               }
    

    def _extract_key_details(self, metric_type: MetricType, metric_result) -> Dict:
        """
        Extract key details specific to each metric type
        """
        details = metric_result.details or {}
        
        if (metric_type == MetricType.GRADIENT):
            return {"Eigenvalue_Ratio" : details.get('eigenvalue_ratio', 'N/A'),
                    "Vectors_Sampled"  : details.get('gradient_vectors_sampled', 'N/A'),
                   }
        
        elif (metric_type == MetricType.FREQUENCY):
            return {"HF_Ratio"        : details.get('hf_ratio', 'N/A'),
                    "HF_Anomaly"      : details.get('hf_anomaly', 'N/A'),
                    "Spectrum_Bins"   : details.get('spectrum_bins', 'N/A'),
                   }
        
        elif (metric_type == MetricType.NOISE):
            return {"Mean_Noise"      : details.get('mean_noise', 'N/A'),
                    "CV"              : details.get('cv', 'N/A'),
                    "Patches_Valid"   : details.get('patches_valid', 'N/A'),
                   }
        
        elif (metric_type == MetricType.TEXTURE):
            return {"Smooth_Ratio"    : details.get('smooth_ratio', 'N/A'),
                    "Contrast_Mean"   : details.get('contrast_mean', 'N/A'),
                    "Patches_Used"    : details.get('patches_used', 'N/A'),
                   }
        
        elif (metric_type == MetricType.COLOR):
            sat_stats = details.get('saturation_stats', {})
            return {"Mean_Saturation" : sat_stats.get('mean_saturation', 'N/A'),
                    "High_Sat_Ratio"  : sat_stats.get('high_sat_ratio', 'N/A'),
                   }
        
        return {}
    

    def _interpret_metric(self, metric_type: MetricType, metric_result) -> str:
        """
        Provide human-readable interpretation of metric result
        """
        score   = metric_result.score
        details = metric_result.details or {}
        
        if (metric_type == MetricType.GRADIENT):
            eig_ratio = details.get('eigenvalue_ratio')
            
            if eig_ratio:
                return f"Eigenvalue ratio of {eig_ratio:.3f} ({'high' if eig_ratio > 0.85 else 'low'} alignment)"
            
            return "Gradient structure analysis"
        
        elif (metric_type == MetricType.FREQUENCY):
            hf_ratio = details.get('hf_ratio')
            
            if hf_ratio:
                return f"High-freq ratio: {hf_ratio:.3f} ({'elevated' if hf_ratio > 0.35 else 'low' if hf_ratio < 0.08 else 'normal'})"
            
            return "Frequency spectrum analysis"
        
        elif (metric_type == MetricType.NOISE):
            mean_noise = details.get('mean_noise')
            
            if mean_noise:
                return f"Mean noise: {mean_noise:.2f} ({'low' if mean_noise < 1.5 else 'normal'})"
            
            return "Noise pattern analysis"
        
        elif (metric_type == MetricType.TEXTURE):
            smooth_ratio = details.get('smooth_ratio')
            
            if smooth_ratio is not None:
                return f"Smooth regions: {smooth_ratio:.1%} ({'excessive' if smooth_ratio > 0.4 else 'normal'})"
            
            return "Texture variation analysis"
        
        elif (metric_type == MetricType.COLOR):
            sat_stats = details.get('saturation_stats', {})
            mean_sat  = sat_stats.get('mean_saturation')
            
            if mean_sat:
                return f"Mean saturation: {mean_sat:.2f} ({'high' if mean_sat > 0.65 else 'normal'})"
            
            return "Color distribution analysis"
        
        return "Analysis complete"
    
    
    def _create_report_header(self, analysis_result: AnalysisResult) -> Dict:
        """
        Create report header section
        """
        return {"filename"        : analysis_result.filename,
                "analysis_date"   : analysis_result.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                "image_size"      : f"{analysis_result.image_size[0]} × {analysis_result.image_size[1]} pixels",
                "processing_time" : f"{analysis_result.processing_time:.2f} seconds",
               }
    

    def _create_overall_assessment(self, analysis_result: AnalysisResult) -> Dict:
        """
        Create overall assessment section
        """
        return {"status"       : analysis_result.status.value,
                "score"        : round(analysis_result.overall_score * 100, 1),
                "confidence"   : analysis_result.confidence,
                "verdict"      : "REVIEW REQUIRED" if analysis_result.status.value == "REVIEW_REQUIRED" else "LIKELY AUTHENTIC",
                "risk_level"   : self._calculate_risk_level(score = analysis_result.overall_score),
               }


    def _create_metric_breakdown(self, analysis_result: AnalysisResult) -> List[Dict]:
        """
        Create detailed metric breakdown for report
        """
        breakdown = list()
        
        for signal in analysis_result.signals:
            metric_result = analysis_result.metric_results.get(signal.metric_type)
            
            item          = {"metric"       : signal.name,
                             "score"        : f"{signal.score * 100:.1f}%",
                             "status"       : signal.status.value.upper(),
                             "confidence"   : f"{metric_result.confidence * 100:.1f}%" if metric_result.confidence else "N/A",
                             "explanation"  : signal.explanation,
                             "key_findings" : self.extract_key_findings(metric_type   = signal.metric_type,
                                                                        metric_result = metric_result,
                                                                       ),
                            }
            
            breakdown.append(item)
        
        return breakdown


    def _create_forensic_details(self, analysis_result: AnalysisResult) -> Dict:
        """
        Create forensic details section
        """
        forensic = dict()
        
        for metric_type, metric_result in analysis_result.metric_results.items():
            metric_name           = self.metric_display_names.get(metric_type, metric_type.value)
            forensic[metric_name] = metric_result.details or {"note": "No detailed forensics available"}
        
        return forensic


    def _create_recommendations(self, analysis_result: AnalysisResult) -> Dict:
        """
        Create recommendations section
        """
        score = analysis_result.overall_score
        
        if (score >= 0.85):
            return {"action"      : "Immediate manual verification required",
                    "priority"    : "HIGH",
                    "next_steps"  : ["Forensic analysis", "Reverse image search", "Metadata inspection", "Expert review"],
                    "confidence"  : "Very high likelihood of AI generation",
                   }
        
        elif (score >= 0.70):
            return {"action"      : "Manual verification recommended",
                    "priority"    : "MEDIUM",
                    "next_steps"  : ["Visual inspection", "Compare with authentic samples", "Check source provenance"],
                    "confidence"  : "High likelihood of AI generation",
                   }
        
        elif (score >= 0.50):
            return {"action"      : "Optional review suggested",
                    "priority"    : "LOW",
                    "next_steps"  : ["May be edited photo", "Verify image source", "Check for inconsistencies"],
                    "confidence"  : "Moderate indicators present",
                   }
        
        else:
            return {"action"      : "No immediate action required",
                    "priority"    : "NONE",
                    "next_steps"  : ["Proceed with normal workflow"],
                    "confidence"  : "Low likelihood of AI generation",
                   }


    def _score_to_status(self, score: float) -> str:
        """
        Convert score to status label
        """
        if (score >= self.signal_thresholds[SignalStatus.FLAGGED]):
            return "FLAGGED"

        elif (score >= self.signal_thresholds[SignalStatus.WARNING]):
            return "WARNING"
        
        else:
            return "PASSED"


    def _calculate_avg_confidence(self, analysis_result: AnalysisResult) -> float:
        """
        Calculate average confidence across all metrics
        """
        confidences = [mr.confidence for mr in analysis_result.metric_results.values() if mr.confidence is not None]
        
        return round(sum(confidences) / len(confidences), 3) if confidences else 0.0


    def _calculate_risk_level(self, score: float) -> str:
        """
        Calculate risk level from score
        """
        if (score >= 0.85):
            return "CRITICAL"

        elif (score >= 0.70):
            return "HIGH"
        
        elif (score >= 0.50):
            return "MEDIUM"
        
        else:
            return "LOW"


    def extract_key_findings(self, metric_type: MetricType, metric_result) -> List[str]:
        """
        Extract human-readable key forensic findings for a given metric used by:
        - Detailed UI views
        - CSV reports
        - JSON reports
        """
        findings = list()
        details  = metric_result.details or {}
        
        if (metric_type == MetricType.GRADIENT):
            eig_ratio = details.get('eigenvalue_ratio')
            
            if eig_ratio:
                findings.append(f"Eigenvalue ratio: {eig_ratio:.3f}")
            
            vectors = details.get('gradient_vectors_sampled')
            
            if vectors:
                findings.append(f"Analyzed {vectors} gradient vectors")
        
        elif (metric_type == MetricType.FREQUENCY):
            hf_ratio = details.get('hf_ratio')
            
            if hf_ratio:
                findings.append(f"High-frequency ratio: {hf_ratio:.3f}")
            
            roughness = details.get('roughness')
            if roughness:
                findings.append(f"Spectral roughness: {roughness:.3f}")
        
        elif (metric_type == MetricType.NOISE):
            mean_noise = details.get('mean_noise')
            
            if mean_noise:
                findings.append(f"Mean noise level: {mean_noise:.2f}")
            
            cv = details.get('cv')
            
            if cv:
                findings.append(f"Coefficient of variation: {cv:.3f}")
        
        elif (metric_type == MetricType.TEXTURE):
            smooth_ratio = details.get('smooth_ratio')
            
            if smooth_ratio:
                findings.append(f"Smooth patches: {smooth_ratio:.1%}")
            
            contrast_mean = details.get('contrast_mean')
            
            if contrast_mean:
                findings.append(f"Average contrast: {contrast_mean:.2f}")
        
        elif (metric_type == MetricType.COLOR):
            sat_stats = details.get('saturation_stats', {})
            mean_sat  = sat_stats.get('mean_saturation')
            
            if mean_sat:
                findings.append(f"Mean saturation: {mean_sat:.2f}")
            
            high_sat = sat_stats.get('high_sat_ratio')
            
            if high_sat:
                findings.append(f"High saturation pixels: {high_sat:.1%}")
        
        return findings if findings else ["Analysis complete"]