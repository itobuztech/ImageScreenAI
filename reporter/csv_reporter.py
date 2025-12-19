# Dependencies
import csv
from pathlib import Path
from typing import Optional
from datetime import datetime
from utils.logger import get_logger
from config.settings import settings
from config.constants import MetricType
from config.schemas import AnalysisResult
from utils.helpers import generate_unique_id
from config.constants import DetectionStatus
from config.schemas import BatchAnalysisResult
from features.detailed_result_maker import DetailedResultMaker


# Setup Logging
logger = get_logger(__name__)


class CSVReporter:
    """
    Professional CSV report generator
    
    Features:
    ---------
    - Single image detailed reports
    - Batch summary reports with statistics
    - Detailed forensic data export
    - Excel-compatible formatting
    - UTF-8 encoding with BOM for international compatibility
    """
    def __init__(self):
        """
        Initialize CSV Reporter
        """
        self.detailed_maker = DetailedResultMaker()
        logger.debug("CSVReporter initialized")
    

    def export_batch_summary(self, batch_result: BatchAnalysisResult, output_dir: Optional[Path] = None) -> Path:
        """
        Export batch analysis summary as CSV
        
        Arguments:
        ----------
            batch_result { BatchAnalysisResult } : Complete batch analysis result
            
            output_dir   { Path }                : Output directory (defaults to settings.REPORTS_DIR)
        
        Returns:
        --------
                       { Path }                  : Path to generated CSV file
        """
        output_dir  = output_dir or settings.REPORTS_DIR
        report_id   = generate_unique_id()
        filename    = f"batch_summary_{report_id}.csv"
        output_path = output_dir / filename
        
        logger.info(f"Generating batch summary CSV: {filename}")
        
        try:
            with open(output_path, 'w', newline = '', encoding = 'utf-8-sig') as f:
                writer = csv.writer(f)
                
                # Report Header
                self._write_report_header(writer      = writer,
                                          report_type = "Batch Analysis Summary",
                                          timestamp   = batch_result.timestamp,
                                         )
                
                # Batch Statistics
                self._write_batch_statistics(writer       = writer,
                                             batch_result = batch_result,
                                            )
                
                # Main Results Table
                self._write_batch_results_table(writer       = writer,
                                                batch_result = batch_result,
                                               )
                
                # Footer
                self._write_footer(writer = writer)
            
            logger.info(f"Batch summary CSV generated: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Failed to generate batch summary CSV: {e}")
            raise
    

    def export_batch_detailed(self, batch_result: BatchAnalysisResult, output_dir: Optional[Path] = None) -> Path:
        """
        Export detailed batch analysis with forensic data
        
        Arguments:
        ----------
            batch_result { BatchAnalysisResult } : Complete batch analysis result
            
            output_dir   { Path }                : Output directory (defaults to settings.REPORTS_DIR)
        
        Returns:
        --------
                      { Path }                   : Path to generated CSV file
        """
        output_dir  = output_dir or settings.REPORTS_DIR
        report_id   = generate_unique_id()
        filename    = f"batch_detailed_{report_id}.csv"
        output_path = output_dir / filename
        
        logger.info(f"Generating detailed batch CSV: {filename}")
        
        try:
            with open(output_path, 'w', newline = '', encoding = 'utf-8-sig') as f:
                writer = csv.writer(f)
                
                # Report Header
                self._write_report_header(writer      = writer,
                                          report_type = "Detailed Batch Analysis",
                                          timestamp   = batch_result.timestamp,
                                         )
                
                # Process each image with full details
                for idx, result in enumerate(batch_result.results, 1):
                    self._write_detailed_image_section(writer        = writer,
                                                       result        = result,
                                                       image_number  = idx,
                                                       total_images  = batch_result.processed,
                                                      )
                    
                    # Add separator between images
                    if (idx < batch_result.processed):
                        writer.writerow([])
                        writer.writerow(['=' * 100])
                        writer.writerow([])
                
                # Footer
                self._write_footer(writer = writer)
            
            logger.info(f"Detailed batch CSV generated: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Failed to generate detailed batch CSV: {e}")
            raise
    

    def export_single_detailed(self, result: AnalysisResult, output_dir: Optional[Path] = None) -> Path:
        """
        Export single image detailed analysis as CSV
        
        Arguments:
        ----------
            result     { AnalysisResult } : Single image analysis result
            
            output_dir { Path }           : Output directory (defaults to settings.REPORTS_DIR)
        
        Returns:
        --------
                     { Path }             : Path to generated CSV file
        """
        output_dir  = output_dir or settings.REPORTS_DIR
        report_id   = generate_unique_id()
        filename    = f"single_analysis_{report_id}.csv"
        output_path = output_dir / filename
        
        logger.info(f"Generating single image CSV: {filename}")
        
        try:
            with open(output_path, 'w', newline = '', encoding = 'utf-8-sig') as f:
                writer = csv.writer(f)
                
                # Report Header
                self._write_report_header(writer     = writer,
                                          report_type = "Single Image Analysis",
                                          timestamp   = result.timestamp,
                                         )
                
                # Image Details
                self._write_detailed_image_section(writer       = writer,
                                                   result       = result,
                                                   image_number = 1,
                                                   total_images = 1,
                                                  )
                
                # Footer
                self._write_footer(writer = writer)
            
            logger.info(f"Single image CSV generated: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Failed to generate single image CSV: {e}")
            raise
    

    def export_metrics_comparison(self, batch_result: BatchAnalysisResult, output_dir: Optional[Path] = None) -> Path:
        """
        Export metrics comparison table across all images
        
        Arguments:
        ----------
            batch_result { BatchAnalysisResult } : Complete batch analysis result
            
            output_dir   { Path }                : Output directory (defaults to settings.REPORTS_DIR)
        
        Returns:
        --------
                       { Path }                  : Path to generated CSV file
        """
        output_dir  = output_dir or settings.REPORTS_DIR
        report_id   = generate_unique_id()
        filename    = f"metrics_comparison_{report_id}.csv"
        output_path = output_dir / filename
        
        logger.info(f"Generating metrics comparison CSV: {filename}")
        
        try:
            with open(output_path, 'w', newline = '', encoding = 'utf-8-sig') as f:
                writer = csv.writer(f)
                
                # Report Header
                self._write_report_header(writer     = writer,
                                          report_type = "Metrics Comparison",
                                          timestamp   = batch_result.timestamp,
                                         )
                
                # Comparison Table Header
                writer.writerow(['Metrics Comparison Across All Images'])
                writer.writerow([])
                
                header = ['Filename', 
                          'Overall Score', 
                          'Analysis Status', 
                          'Gradient Analysis Score', 
                          'Gradient Analysis Confidence', 
                          'Frequency Analysis Score', 
                          'Frequency Analysis Confidence',
                          'Noise Analysis Score', 
                          'Noise Analysis Confidence',
                          'Texture Analysis Score', 
                          'Texture Analysis Confidence',
                          'Color Analysis Score', 
                          'Color Analysis Confidence',
                          'Processing Time',
                         ]
                
                writer.writerow(header)
                
                # Data rows
                for result in batch_result.results:
                    row = [result.filename,
                           f"{result.overall_score:.3f}",
                           result.status.value,
                          ]
                    
                    # Add each metric's score and confidence
                    for metric_type in [MetricType.GRADIENT, MetricType.FREQUENCY, MetricType.NOISE, MetricType.TEXTURE, MetricType.COLOR]:
                        metric_result = result.metric_results.get(metric_type)
                        
                        if metric_result:
                            row.append(f"{metric_result.score:.3f}")
                            row.append(f"{metric_result.confidence:.3f}" if metric_result.confidence is not None else "N/A")
                        
                        else:
                            row.extend(["N/A", "N/A"])
                    
                    row.append(f"{result.processing_time:.2f}s")
                    writer.writerow(row)
                
                # Footer
                writer.writerow([])
                self._write_footer(writer = writer)
            
            logger.info(f"Metrics comparison CSV generated: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Failed to generate metrics comparison CSV: {e}")
            raise
    

    def _write_report_header(self, writer, report_type: str, timestamp: datetime) -> None:
        """
        Write CSV report header
        """
        writer.writerow(['=' * 100])
        writer.writerow([f'AI Image Screener - {report_type}'])
        writer.writerow([f'Generated: {timestamp.strftime("%Y-%m-%d %H:%M:%S")}'])
        writer.writerow([f'Version: {settings.VERSION}'])
        writer.writerow(['=' * 100])
        writer.writerow([])
    

    def _write_batch_statistics(self, writer, batch_result: BatchAnalysisResult) -> None:
        """
        Write batch statistics section
        """
        writer.writerow(['BATCH STATISTICS'])
        writer.writerow([])
        
        stats = [['Total Images', batch_result.total_images],
                 ['Successfully Processed', batch_result.processed],
                 ['Failed', batch_result.failed],
                 ['Success Rate', f"{batch_result.summary.get('success_rate', 0)}%"],
                 ['' , ''],
                 ['Likely Authentic', batch_result.summary.get('likely_authentic', 0)],
                 ['Review Required', batch_result.summary.get('review_required', 0)],
                 ['', ''],
                 ['Average Score', f"{batch_result.summary.get('avg_score', 0):.3f}"],
                 ['Average Confidence', f"{batch_result.summary.get('avg_confidence', 0)}%"],
                 ['Total Processing Time', f"{batch_result.total_processing_time:.2f}s"],
                 ['Average Time per Image', f"{batch_result.summary.get('avg_proc_time', 0):.2f}s"],
                ]
        
        for row in stats:
            writer.writerow(row)
        
        writer.writerow([])
        writer.writerow(['=' * 100])
        writer.writerow([])
    

    def _write_batch_results_table(self, writer, batch_result: BatchAnalysisResult) -> None:
        """
        Write batch results main table
        """
        writer.writerow(['ANALYSIS RESULTS'])
        writer.writerow([])
        
        # Table Header
        header = ['Filename', 
                  'Image Size',
                  'Analysis Status', 
                  'Overall Score', 
                  'Analysis Confidence (%)', 
                  'Top Warning Signals', 
                  'Recommendation', 
                  'Processing Time (s)', 
                 ]

        writer.writerow(header)
        
        # Data rows
        for result in batch_result.results:
            # Get top warning signals
            top_signals = [s.name for s in result.signals if s.status.value in ['flagged', 'warning']][:2]
            signals_str = "; ".join(top_signals) if top_signals else "All tests passed"
            
            # Recommendation
            if (result.status == DetectionStatus.REVIEW_REQUIRED):
                recommendation = "Manual verification recommended"

            else:
                recommendation = "No further action needed"
            
            row = [result.filename,
                   f"{result.image_size[0]}×{result.image_size[1]}",
                   result.status.value,
                   f"{result.overall_score:.3f}",
                   f"{result.confidence}%",
                   signals_str,
                   recommendation,
                   f"{result.processing_time:.2f}", 
                  ]
            
            writer.writerow(row)
        
        writer.writerow([])
    

    def _write_detailed_image_section(self, writer, result: AnalysisResult, image_number: int, total_images: int) -> None:
        """
        Write detailed section for single image
        """
        writer.writerow([f'IMAGE {image_number} OF {total_images}'])
        writer.writerow([])
        
        # Basic Information
        writer.writerow(['BASIC INFORMATION'])
        writer.writerow(['Filename', result.filename])
        writer.writerow(['Status', result.status.value])
        writer.writerow(['Overall Score', f"{result.overall_score:.3f}"])
        writer.writerow(['Confidence', f"{result.confidence}%"])
        writer.writerow(['Image Size', f"{result.image_size[0]}×{result.image_size[1]}"])
        writer.writerow(['Processing Time', f"{result.processing_time:.2f}s"])
        writer.writerow(['Timestamp', result.timestamp.isoformat()])
        writer.writerow([])
        
        # Detection Signals
        writer.writerow(['DETECTION SIGNALS'])
        writer.writerow([])
        writer.writerow(['Metric Name', 'Metric Score', 'Analysis Status', 'Metric Confidence', 'Metric Explanation'])
        
        for signal in result.signals:
            metric_result  = result.metric_results.get(signal.metric_type)
            confidence_str = f"{metric_result.confidence:.3f}" if metric_result.confidence is not None else "N/A"
            
            writer.writerow([signal.name,
                             f"{signal.score:.3f}",
                             signal.status.value.upper(),
                             confidence_str,
                             signal.explanation.replace("\n", " "),
                           ])
        
        writer.writerow([])
        
        # Detailed Forensics
        writer.writerow(['FORENSIC DETAILS'])
        writer.writerow([])

        for metric_type in MetricType:
            metric_result = result.metric_results.get(metric_type)
            
            if not metric_result:
                continue

            metric_name = self.detailed_maker.metric_display_names.get(metric_type, metric_type.value)
            
            writer.writerow([f'--- {metric_name} ---'])
            writer.writerow(['Score', f"{metric_result.score:.3f}"])
            writer.writerow(['Confidence', f"{metric_result.confidence:.3f}" if metric_result.confidence is not None else "N/A"])
            
            # Write details
            if metric_result.details:
                for key, value in metric_result.details.items():
                    if isinstance(value, dict):
                        writer.writerow([f"  {key}:", ""])
                        for sub_key, sub_value in value.items():
                            writer.writerow([f"    {sub_key}", str(sub_value)])
                    
                    else:
                        writer.writerow([f"  {key}", str(value)])
            
            writer.writerow([])
        
        # Recommendation
        writer.writerow(['RECOMMENDATION'])
        writer.writerow([])
        
        if (result.status == DetectionStatus.REVIEW_REQUIRED):
            writer.writerow(['Action', 'Manual verification recommended'])
            writer.writerow(['Priority', 'HIGH' if (result.overall_score >= 0.85) else 'MEDIUM'])
            writer.writerow(['Next Steps', 'Forensic analysis, reverse image search, metadata inspection'])
        
        else:
            writer.writerow(['Action', 'No immediate action needed'])
            writer.writerow(['Priority', 'LOW'])
            writer.writerow(['Next Steps', 'Proceed with normal workflow'])
        
        writer.writerow([])
    

    def _write_footer(self, writer) -> None:
        """
        Write CSV report footer
        """
        writer.writerow(['=' * 100])
        writer.writerow(['Report generated by AI Image Screener'])
        writer.writerow(['For questions or support, contact: support@aiimagescreener.com'])
        writer.writerow(['DISCLAIMER: Results are indicative and should be verified manually for critical applications'])
        writer.writerow(['=' * 100])
