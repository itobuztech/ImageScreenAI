# Dependencies
import time
from typing import List
from typing import Dict
from typing import Tuple
from pathlib import Path
from typing import Callable
from utils.logger import get_logger
from config.settings import settings
from config.schemas import AnalysisResult
from concurrent.futures import TimeoutError
from concurrent.futures import as_completed
from config.constants import DetectionStatus
from config.schemas import BatchAnalysisResult
from metrics.aggregator import MetricsAggregator
from concurrent.futures import ThreadPoolExecutor
from features.threshold_manager import ThresholdManager


# Setup Logging
logger = get_logger(__name__)


class BatchProcessor:
    """
    Process multiple images in parallel or sequential mode
    
    Features:
    ---------
    - Parallel processing using ThreadPoolExecutor
    - Sequential fallback for single images or disabled parallel mode
    - Automatic error handling and recovery
    - Progress tracking and logging
    """
    def __init__(self, threshold_manager: ThresholdManager):
        """
        Initialize Batch Processor
        """
        # Instantiate threshold manager
        self.threshold_manager = threshold_manager

        # Initialize aggregator
        self.aggregator        = MetricsAggregator(threshold_manager = threshold_manager)
            
        # Fix number of workers 
        self.max_workers       = settings.MAX_WORKERS if settings.PARALLEL_PROCESSING else 1
        
        logger.info(f"BatchProcessor initialized with max_workers={self.max_workers}, parallel={settings.PARALLEL_PROCESSING}")
    

    def process_batch(self, image_files: List[Dict[str, any]], on_progress: Callable[[int, int, str], None] | None = None) -> BatchAnalysisResult:
        """
        Process multiple images with automatic parallel/sequential switching
        
        Arguments:
        ----------
            image_files   { list }    : List of dicts with keys:
                                        - 'path'     : Path object
                                        - 'filename' : str
                                        - 'size'     : tuple (width, height)

            on_progress { Callablel } : Optional callback invoked after each image is processed
        
        Returns:
        --------
            { BatchAnalysisResult } : Complete batch analysis result
        """
        start_time   = time.time()
        total_images = len(image_files)
        
        logger.info(f"Starting batch processing of {total_images} images")
        
        # Validate input
        if (total_images == 0):
            logger.warning("Empty batch provided")
            return self._create_empty_batch_result()
        
        if (total_images > settings.MAX_BATCH_SIZE):
            logger.error(f"Batch size {total_images} exceeds maximum {settings.MAX_BATCH_SIZE}")
            raise ValueError(f"Batch size {total_images} exceeds maximum allowed {settings.MAX_BATCH_SIZE}")
        
        # Choose processing strategy
        if (settings.PARALLEL_PROCESSING and (total_images > 1)):
            results, failed = self._process_parallel(image_files = image_files,
                                                     on_progress = on_progress,
                                                    )
        
        else:
            results, failed = self._process_sequential(image_files = image_files,
                                                       on_progress = on_progress,
                                                      )
        
        total_time           = time.time() - start_time
        
        # Create batch result
        batch_result         = BatchAnalysisResult(total_images          = total_images,
                                                   processed             = len(results),
                                                   failed                = failed,
                                                   results               = results,
                                                   total_processing_time = total_time,
                                                  )
        
        # Calculate summary statistics
        batch_result.summary = self._calculate_summary(results = results,
                                                       total   = total_images,
                                                      )
        
        logger.info(f"Batch processing complete: {len(results)}/{total_images} successful, {failed} failed in {total_time:.2f}s")
        
        return batch_result
    

    def _process_parallel(self, image_files: List[Dict], on_progress: Callable[[int, int, str], None] | None = None) -> Tuple[List[AnalysisResult], int]:
        """
        Process images in parallel using ThreadPoolExecutor
        
        Arguments:
        ----------
            image_files   { list }    : List of image file dictionaries

            on_progress { Callablel } : Optional callback invoked after each image is processed
        
        Returns:
        --------
            { tuple }            : (results_list, failed_count)
        """
        results = list()
        failed  = 0
        
        logger.debug(f"Using parallel processing with {self.max_workers} workers")
        
        with ThreadPoolExecutor(max_workers = self.max_workers) as executor:
            # Submit all tasks
            future_to_file = {executor.submit(self.process_single,
                                              image['path'],
                                              image['filename'],
                                              image['size'],
                                             ): image for image in image_files
                             }
            
            # Collect results as they complete
            completed = 0

            for future in as_completed(future_to_file):
                completed += 1
                image      = future_to_file[future]

                if on_progress:
                    on_progress(completed, len(image_files), image["filename"])
                
                try:
                    result = future.result(timeout = settings.PROCESSING_TIMEOUT)
                    
                    if result:
                        results.append(result)
                        logger.debug(f"✓ Completed: {image['filename']}")
                    
                    else:
                        failed += 1
                        logger.warning(f"✗ Failed: {image['filename']} (returned None)")
                
                except TimeoutError:
                    failed += 1
                    logger.error(f"✗ Timeout: {image['filename']} (exceeded {settings.PROCESSING_TIMEOUT}s)")
                
                except Exception as e:
                    failed += 1
                    logger.error(f"✗ Error: {image['filename']} - {e}")
        
        return results, failed
    

    def _process_sequential(self, image_files: List[Dict], on_progress: Callable[[int, int, str], None] | None = None) -> Tuple[List[AnalysisResult], int]:
        """
        Process images sequentially (fallback or single image)
        
        Arguments:
        ----------
            image_files   { list }   : List of image file dictionaries

            on_progress { Callabel } : Optional callback invoked after each image is processed
        
        Returns:
        --------
            { tuple }            : (results_list, failed_count)
        """
        results = list()
        failed  = 0
        
        logger.debug("Using sequential processing")
        
        for idx, image in enumerate(image_files, 1):
            try:
                if on_progress:
                    on_progress(idx, len(image_files), image["filename"])
                
                result = self.process_single(image_path = image['path'],
                                             filename   = image['filename'],
                                             image_size = image['size'],
                                            )
                
                if result:
                    results.append(result)
                    logger.debug(f"✓ Completed: {image['filename']}")
                
                else:
                    failed += 1
                    logger.warning(f"✗ Failed: {image['filename']} (returned None)")
            
            except Exception as e:
                failed += 1
                logger.error(f"✗ Error: {image['filename']} - {e}")
        
        return results, failed
    

    def process_single(self, image_path: Path, filename: str, image_size: Tuple[int, int]) -> AnalysisResult:
        """
        Process single image (called by both parallel and sequential)
        
        Arguments:
        ----------
            image_path { Path }  : Path to image file
            
            filename   { str }   : Original filename
            
            image_size { tuple } : (width, height)
        
        Returns:
        --------
            { AnalysisResult }   : Analysis result or None on error
        """
        try:
            return self.aggregator.analyze_image(image_path = image_path,
                                                 filename   = filename,
                                                 image_size = image_size,
                                                )
        
        except Exception as e:
            logger.error(f"Failed to process {filename}: {e}", exc_info = True)
            return None
    

    def _calculate_summary(self, results: List[AnalysisResult], total: int) -> Dict[str, int]:
        """
        Calculate summary statistics from results
        
        Arguments:
        ----------
            results { list } : List of analysis results
            
            total   { int }  : Total number of images
        
        Returns:
        --------
            { dict }         : Summary statistics
        """
        # Calculate processing stats
        likely_authentic = sum(1 for r in results if (r.status == DetectionStatus.LIKELY_AUTHENTIC))
        review_required  = sum(1 for r in results if (r.status == DetectionStatus.REVIEW_REQUIRED))

        processed        = len(results)
        failed           = total - processed
        success_rate     = int((processed / total * 100) if (total > 0) else 0)
        
        # Calculate average scores
        avg_score        = sum(r.overall_score for r in results) / len(results) if results else 0.0
        avg_confidence   = sum(r.confidence for r in results) / len(results) if results else 0
        avg_proc_time    = sum(r.processing_time for r in results) / len(results) if results else 0.0
        
        return {"likely_authentic" : likely_authentic,
                "review_required"  : review_required,
                "success_rate"     : success_rate,
                "processed"        : processed,
                "failed"           : failed,
                "avg_score"        : round(avg_score, 3),
                "avg_confidence"   : int(avg_confidence),
                "avg_proc_time"    : round(avg_proc_time, 2),
               }
    

    def _create_empty_batch_result(self) -> BatchAnalysisResult:
        """
        Create empty batch result for edge cases
        
        Returns:
        --------
            { BatchAnalysisResult } : Empty batch result
        """
        return BatchAnalysisResult(total_images          = 0,
                                   processed             = 0,
                                   failed                = 0,
                                   results               = [],
                                   summary               = {"likely_authentic" : 0,
                                                            "review_required"  : 0,
                                                            "success_rate"     : 0,
                                                           },
                                   total_processing_time = 0.0,
                                  )
