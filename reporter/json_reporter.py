# Dependencies
import json
from typing import Dict
from typing import List
from pathlib import Path
from typing import Optional
from datetime import datetime
from utils.logger import get_logger
from config.settings import settings
from config.schemas import AnalysisResult
from utils.helpers import generate_unique_id
from config.schemas import BatchAnalysisResult
from features.detailed_result_maker import DetailedResultMaker


# Setup Logging
logger = get_logger(__name__)


class JSONReporter:
    """
    Professional JSON report generator
    
    Features:
    ---------
    - Machine-readable structured format
    - API-friendly output
    - Complete data preservation
    - Pretty-printed for readability
    - Nested structure for complex data
    """
    def __init__(self):
        """
        Initialize JSON Reporter
        """
        self.detailed_maker = DetailedResultMaker()
        logger.debug("JSONReporter initialized")
    

    def export_batch(self, batch_result: BatchAnalysisResult, output_dir: Optional[Path] = None, include_detailed: bool = True) -> Path:
        """
        Export batch analysis as JSON
        
        Arguments:
        ----------
            batch_result     { BatchAnalysisResult } : Complete batch analysis result
            
            output_dir       { Path }                : Output directory (defaults to settings.REPORTS_DIR)
            
            include_detailed { bool }                : Include detailed forensic data
        
        Returns:
        --------
                        { Path }                     : Path to generated JSON file
        """
        output_dir  = output_dir or settings.REPORTS_DIR
        report_id   = generate_unique_id()
        filename    = f"batch_report_{report_id}.json"
        output_path = output_dir / filename

        output_dir.mkdir(parents = True, exist_ok = True)
        
        logger.info(f"Generating batch JSON: {filename}")
        
        try:
            # Build JSON structure
            data = self._build_batch_json(batch_result     = batch_result,
                                          include_detailed = include_detailed,
                                         )
            
            # Write to file
            with open(output_path, 'w', encoding = 'utf-8') as f:
                json.dump(obj          = data, 
                          fp           = f, 
                          indent       = 4, 
                          ensure_ascii = False, 
                          default      = str,
                         )
            
            logger.info(f"Batch JSON generated: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Failed to generate batch JSON: {e}")
            raise
    

    def export_single(self, result: AnalysisResult, output_dir: Optional[Path] = None, include_detailed: bool = True) -> Path:
        """
        Export single image analysis as JSON
        
        Arguments:
        ----------
            result           { AnalysisResult } : Single image analysis result
            
            output_dir            { Path }      : Output directory (defaults to settings.REPORTS_DIR)
            
            include_detailed      { bool }      : Include detailed forensic data
        
        Returns:
        --------
                      { Path }                  : Path to generated JSON file
        """
        output_dir  = output_dir or settings.REPORTS_DIR
        report_id   = generate_unique_id()
        filename    = f"single_report_{report_id}.json"
        output_path = output_dir / filename

        output_dir.mkdir(parents = True, exist_ok = True)
        
        logger.info(f"Generating single image JSON: {filename}")
        
        try:
            # Build JSON structure
            data = self._build_single_json(result           = result,
                                           include_detailed = include_detailed,
                                          )
            
            # Write to file
            with open(output_path, 'w', encoding = 'utf-8') as f:
                json.dump(obj          = data, 
                          fp           = f, 
                          indent       = 4, 
                          ensure_ascii = False, 
                          default      = str,
                         )
            
            logger.info(f"Single image JSON generated: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Failed to generate single image JSON: {e}")
            raise
    

    def export_api_response(self, result: AnalysisResult) -> Dict:
        """
        Generate API-friendly JSON response (in-memory, no file)
        
        Arguments:
        ----------
            result { AnalysisResult } : Analysis result
        
        Returns:
        --------
                   { dict }           : API response dictionary
        """
        return {"success"   : True,
                "timestamp" : datetime.now().isoformat(),
                "version"   : settings.VERSION,
                "data"      : self._build_single_json(result           = result,
                                                      include_detailed = False,
                                                     ),
               }
    

    def _build_batch_json(self, batch_result: BatchAnalysisResult, include_detailed: bool) -> Dict:
        """
        Build complete batch JSON structure
        """
        data = {"report_metadata" : self._build_metadata(report_type = "Batch Analysis",
                                                         timestamp   = batch_result.timestamp,
                                                        ),
                "batch_summary"   : self._build_batch_summary(batch_result = batch_result),
                "results"         : [],
               }
        
        # Add each image result
        for result in batch_result.results:
            image_data = self._build_image_data(result           = result,
                                                include_detailed = include_detailed,
                                               )
            data["results"].append(image_data)
        
        return data
    

    def _build_single_json(self, result: AnalysisResult, include_detailed: bool) -> Dict:
        """
        Build single image JSON structure
        """
        data = {"report_metadata" : self._build_metadata(report_type = "Single Image Analysis",
                                                         timestamp   = result.timestamp,
                                                        ),
                "analysis"        : self._build_image_data(result           = result,
                                                           include_detailed = include_detailed,
                                                          ),
               }
        
        return data
    

    def _build_metadata(self, report_type: str, timestamp: datetime) -> Dict:
        """
        Build report metadata section
        """
        return {"report_type"    : report_type,
                "generated_at"   : timestamp.isoformat(),
                "generator"      : "AI Image Screener",
                "version"        : settings.VERSION,
                "format_version" : "1.0",
               }
    

    def _build_batch_summary(self, batch_result: BatchAnalysisResult) -> Dict:
        """
        Build batch summary section
        """
        return {"total_images"          : batch_result.total_images,
                "processed"             : batch_result.processed,
                "failed"                : batch_result.failed,
                "success_rate"          : batch_result.summary.get('success_rate', 0),
                "statistics"            : {"likely_authentic" : batch_result.summary.get('likely_authentic', 0),
                                           "review_required"  : batch_result.summary.get('review_required', 0),
                                           "avg_score"        : batch_result.summary.get('avg_score', 0.0),
                                           "avg_confidence"   : batch_result.summary.get('avg_confidence', 0),
                                           "avg_proc_time"    : batch_result.summary.get('avg_proc_time', 0.0),
                                          },
                "total_processing_time" : round(batch_result.total_processing_time, 2),
               }
    

    def _build_image_data(self, result: AnalysisResult, include_detailed: bool) -> Dict:
        """
        Build complete image data structure
        """
        image_data = {"filename"     : result.filename,
                      "status"       : result.status.value,
                      "overall"      : {"score"           : round(result.overall_score, 3),
                                        "confidence"      : result.confidence,
                                        "interpretation"  : self._interpret_score(score = result.overall_score),
                                       },
                      "image_info"   : {"size"            : {"width"  : result.image_size[0],
                                                             "height" : result.image_size[1],
                                                            },
                                        "processing_time" : round(result.processing_time, 2),
                                        "timestamp"       : result.timestamp.isoformat(),
                                       },
                      "signals"      : self._build_signals_data(result = result),
                     }
        
        # Add detailed forensics if requested
        if include_detailed:
            image_data["forensics"]       = self._build_forensics_data(result = result)
            image_data["recommendations"] = self._build_recommendations(result = result)
        
        return image_data
    

    def _build_signals_data(self, result: AnalysisResult) -> List[Dict]:
        """
        Build signals data structure
        """
        signals = list()
        
        for signal in result.signals:
            metric_result = result.metric_results.get(signal.metric_type)
            
            signal_data   = {"metric_name" : signal.name,
                             "metric_type" : signal.metric_type.value,
                             "score"       : round(signal.score, 3),
                             "status"      : signal.status.value,
                             "confidence"  : round(metric_result.confidence, 3) if (metric_result and metric_result.confidence is not None) else None,
                             "explanation" : signal.explanation,
                            }
            
            signals.append(signal_data)
        
        return signals
    

    def _build_forensics_data(self, result: AnalysisResult) -> Dict:
        """
        Build detailed forensics data structure
        """
        forensics = dict()
        
        for metric_type, metric_result in result.metric_results.items():
            metric_name                  = self.detailed_maker.metric_display_names.get(metric_type, metric_type.value)
            
            forensics[metric_type.value] = {"display_name" : metric_name,
                                            "score"        : round(metric_result.score, 3),
                                            "confidence"   : round(metric_result.confidence, 3) if (metric_result and metric_result.confidence is not None) else None,
                                            "details"      : metric_result.details or {},
                                            "key_findings" : self.detailed_maker.extract_key_findings(metric_type   = metric_type,
                                                                                                      metric_result = metric_result,
                                                                                                     ),
                                           }
        
        return forensics
    

    def _build_recommendations(self, result: AnalysisResult) -> Dict:
        """
        Build recommendations structure
        """
        score = result.overall_score
        
        if (score >= 0.85):
            return {"action"      : "Immediate manual verification required",
                    "priority"    : "HIGH",
                    "risk_level"  : "CRITICAL",
                    "next_steps"  : ["Forensic analysis", "Reverse image search", "Metadata inspection"],
                    "confidence"  : "Very high likelihood of AI generation",
                   }
        
        elif (score >= 0.70):
            return {"action"      : "Manual verification recommended",
                    "priority"    : "MEDIUM",
                    "risk_level"  : "HIGH",
                    "next_steps"  : ["Visual inspection", "Compare with authentic samples"],
                    "confidence"  : "High likelihood of AI generation",
                   }
        
        elif (score >= 0.50):
            return {"action"      : "Optional review suggested",
                    "priority"    : "LOW",
                    "risk_level"  : "MEDIUM",
                    "next_steps"  : ["Verify image source", "Check for inconsistencies"],
                    "confidence"  : "Moderate indicators present",
                   }
        
        else:
            return {"action"      : "No immediate action required",
                    "priority"    : "NONE",
                    "risk_level"  : "LOW",
                    "next_steps"  : ["Proceed with normal workflow"],
                    "confidence"  : "Low likelihood of AI generation",
                   }
    

    def _interpret_score(self, score: float) -> str:
        """
        Interpret score for human readability
        """
        if (score >= 0.85):
            return "Very high suspicion"

        elif (score >= 0.70):
            return "High suspicion"
        
        elif (score >= 0.50):
            return "Moderate suspicion"
        
        elif (score >= 0.30):
            return "Low suspicion"
        
        else:
            return "Very low suspicion"
