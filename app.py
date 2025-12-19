# Dependencies
import uuid
import shutil
import signal
import uvicorn
import traceback
from typing import List
from typing import Dict
from pathlib import Path
from fastapi import File
from typing import Optional
from fastapi import Request
from fastapi import FastAPI
from fastapi import UploadFile
from fastapi import HTTPException
from utils.logger import get_logger
from config.settings import settings
from fastapi.responses import Response
from config.schemas import APIResponse
from config.schemas import AnalysisResult
from fastapi.responses import HTMLResponse
from fastapi.responses import JSONResponse
from utils.validators import ImageValidator
from fastapi.staticfiles import StaticFiles
from utils.helpers import generate_unique_id
from reporter.csv_reporter import CSVReporter
from reporter.pdf_reporter import PDFReporter
from config.schemas import BatchAnalysisResult
from reporter.json_reporter import JSONReporter
from utils.image_processor import ImageProcessor
from fastapi.middleware.cors import CORSMiddleware
from features.batch_processor import BatchProcessor
from features.threshold_manager import ThresholdManager


# Logging
logger = get_logger(__name__)


# FastAPI App Definition
app = FastAPI(title       = "ImageScreenAI",
              version     = settings.VERSION,
              description = "First-pass AI image screening tool for bulk workflows",
             )


# Serve static assets (if any later)
app.mount("/ui", StaticFiles(directory = "ui"), name = "ui")

# CORS (UI + API)
app.add_middleware(CORSMiddleware,
                   allow_origins     = ["*"],
                   allow_credentials = True,
                   allow_methods     = ["*"],
                   allow_headers     = ["*"],
                  )

# Runtime State
SESSION_STORE: Dict[str, Dict] = {}

# Component Initialization
image_validator   = ImageValidator()
image_processor   = ImageProcessor()

threshold_manager = ThresholdManager()
threshold_manager = threshold_manager
batch_processor   = BatchProcessor(threshold_manager = threshold_manager)

json_reporter     = JSONReporter()
csv_reporter      = CSVReporter()
pdf_reporter      = PDFReporter()

UPLOAD_DIR        = settings.UPLOAD_DIR
CACHE_DIR         = settings.CACHE_DIR
REPORTS_DIR       = settings.REPORTS_DIR

for d in [UPLOAD_DIR, CACHE_DIR, REPORTS_DIR]:
    d.mkdir(parents  = True, 
            exist_ok = True,
           )


# Utility: Progress Callback
def _progress_callback(batch_id: str):
    def callback(image_idx: int, total: int, filename: str):
        session = SESSION_STORE.get(batch_id)
        if (not session or (session.get("status") != "processing")):
            return

        session["progress"] = {"current"  : image_idx,
                               "total"    : total,
                               "filename" : filename,
                              }
    return callback


# Utility: Housekeeping
def cleanup_temp_files():
    try:
        for folder in [UPLOAD_DIR, CACHE_DIR]:
            for item in folder.iterdir():
                if item.is_file():
                    item.unlink(missing_ok = True)

        logger.info("Temporary files cleaned")

    except Exception as e:
        logger.warning(f"Cleanup failed: {e}")


def shutdown_handler(*_):
    logger.warning("Shutdown signal received — cleaning up")
    cleanup_temp_files()


signal.signal(signal.SIGINT, shutdown_handler)
signal.signal(signal.SIGTERM, shutdown_handler)


# Error Handling
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.error(f"Unhandled error: {exc}")
    logger.debug(traceback.format_exc())
    
    return JSONResponse(status_code = 500,
                        content     = APIResponse(success = False,
                                                  message = "Internal server error",
                                                 ).model_dump()
                       )


# Home
@app.get("/", response_class = HTMLResponse)
def serve_frontend():
    index_path = Path("ui/index.html")

    if not index_path.exists():
        raise HTTPException(status_code = 404,
                            detail      = "UI not found",
                           )

    return index_path.read_text(encoding = "utf-8")


# Health Check
@app.get("/health")
def health():
    return {"status"  : "ok",
            "version" : settings.VERSION,
           }


# Single Image Analysis
@app.post("/analyze/image")
async def analyze_single_image(file: UploadFile = File(...)):
    image_id   = generate_unique_id()
    image_path = UPLOAD_DIR / f"{image_id}_{file.filename}"

    image_validator.validate_image(file_path = image_path,
                                   filename  = file.filename,
                                   file_size = file.size,
                                  )

    try:
        with open(image_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        image                  = image_processor.load_image(image_path)

        # image is a NumPy array → shape = (H, W, C) or (H, W)
        height, width          = image.shape[:2]

        result: AnalysisResult = batch_processor.process_single(image      = image_path,
                                                                filename   = file.filename,
                                                                image_size = (width, height),
                                                               )

        return APIResponse(success = True,
                           message = "Image analysis completed",
                           data    = result.model_dump(),
                          )

    finally:
        image_path.unlink(missing_ok = True)


# Batch Image Analysis
@app.post("/analyze/batch")
async def analyze_batch(files: List[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code = 400, 
                            detail      = "No files provided",
                           )

    batch_id                = str(uuid.uuid4())

    SESSION_STORE[batch_id] = {"status"   : "processing",
                               "progress" : {"current" : 0, 
                                             "total"   : len(files),
                                            },
                              }

    image_entries           = list()

    try:
        for file in files:
            uid           = generate_unique_id()
            path          = UPLOAD_DIR / f"{uid}_{file.filename}"

            with open(path, "wb") as f:
                shutil.copyfileobj(file.file, f)
            
            image         = image_processor.load_image(path)
            height, width = image.shape[:2]

            image_validator.validate_image(file_path = path,
                                           filename  = file.filename,
                                           file_size = file.size,
                                          )

            image_entries.append({"path"     : path,
                                  "filename" : file.filename,
                                  "size"     : (width, height),
                                })

        batch_result: BatchAnalysisResult = batch_processor.process_batch(image_files = image_entries,
                                                                          on_progress = _progress_callback(batch_id),
                                                                         )

        SESSION_STORE[batch_id]           = {"status"   : "completed",
                                             "progress" : SESSION_STORE[batch_id]["progress"],
                                             "result"   : batch_result,       
                                            }

        
        return APIResponse(success = True,
                           message = "Batch analysis completed",
                           data    = {"batch_id" : batch_id,
                                      "result"   : batch_result.model_dump(),
                                     },
                          )

    except KeyboardInterrupt:
        SESSION_STORE[batch_id] = {"status"   : "interrupted",
                                   "progress" : SESSION_STORE[batch_id]["progress"],
                                  }

        raise HTTPException(status_code = 499, 
                            detail      = "Processing interrupted",
                           )

    except Exception as e:
        logger.error(f"Batch {batch_id} failed: {e}", exc_info = True)
        
        SESSION_STORE[batch_id] = {"status" : "failed",
                                   "error"  : str(e),
                                  }

        raise HTTPException(status_code = 500, 
                            detail      = "Batch processing failed",
                           )

    finally:
        for item in image_entries:
            Path(item["path"]).unlink(missing_ok = True)


# Batch Progress
@app.get("/batch/{batch_id}/progress")
def batch_progress(batch_id: str):
    session = SESSION_STORE.get(batch_id)
    
    if not session:
        raise HTTPException(status_code = 404,
                            detail      = "Batch not found",
                           )
    
    return session


# Report Downloads
@app.api_route("/report/csv/{batch_id}", methods = ["GET", "POST"])
def export_csv(batch_id: str):
    session = SESSION_STORE.get(batch_id)

    if (not session or ("result" not in session)):
        raise HTTPException(status_code = 404, 
                            detail      = "Batch result not found",
                           )

    path = csv_reporter.export_batch_detailed(session["result"])
    
    # Read the file and send it as a download
    with open(path, "rb") as f:
        content = f.read()
    
    # Clean up the file after sending
    path.unlink(missing_ok = True)
    
    return Response(content    = content,
                    media_type = "text/csv",
                    headers    = {"Content-Disposition" : f"attachment; filename=ai_screener_report_{batch_id}.csv",
                                  "Content-Type"        : "text/csv"
                                 }
                   )


@app.api_route("/report/pdf/{batch_id}", methods = ["GET", "POST"])
def export_pdf(batch_id: str):
    session = SESSION_STORE.get(batch_id)

    if (not session or ("result" not in session)):
        raise HTTPException(status_code = 404, 
                            detail      = "Batch result not found",
                           )

    path = pdf_reporter.export_batch(session["result"])
    
    # Read the file and send it as a download
    with open(path, "rb") as f:
        content = f.read()
    
    # Clean up the file after sending
    path.unlink(missing_ok = True)
    
    return Response(content    = content,
                    media_type = "application/pdf",
                    headers    = {"Content-Disposition" : f"attachment; filename=ai_screener_report_{batch_id}.pdf",
                                  "Content-Type"        : "application/pdf"
                                 }
                   )



# ==================== MAIN ====================
if __name__ == "__main__":
    # Explicit startup log (forces log file creation)
    logger.info("Starting AI Image Screener API Server")

    uvicorn.run("app:app",                    
                host      = settings.HOST,
                port      = settings.PORT,
                reload    = settings.DEBUG,
                log_level = settings.LOG_LEVEL.lower(),
                workers   = 1 if settings.DEBUG else settings.WORKERS,
               )