# Dependencies
from typing import Set
from pathlib import Path
from config.constants import MetricType
from pydantic_settings import BaseSettings
from pydantic_settings import SettingsConfigDict


class Settings(BaseSettings):
    """
    Application settings with environment variable support
    """
    model_config                   = SettingsConfigDict(env_file          = '.env',
                                                        env_file_encoding = 'utf-8',
                                                        case_sensitive    = False,
                                                       )
    
    # Application
    APP_NAME            : str      = "ImageScreenAI"
    VERSION             : str      = "1.0.0"
    DEBUG               : bool     = False
    LOG_LEVEL           : str      = "INFO"

    # Server Configuration
    HOST                : str      = "localhost"
    PORT                : int      = 8005
    WORKERS             : int      = 4
    
    # File processing 
    MAX_FILE_SIZE_MB    : int      = 10
    MAX_BATCH_SIZE      : int      = 50
    ALLOWED_EXTENSIONS  : Set[str] = {".jpg", ".jpeg", ".png", ".webp"}
    
    # Detection thresholds
    REVIEW_THRESHOLD    : float    = 0.65
    
    # Metric weights (must sum to 1.0)
    GRADIENT_WEIGHT     : float    = 0.30
    FREQUENCY_WEIGHT    : float    = 0.25
    NOISE_WEIGHT        : float    = 0.20
    TEXTURE_WEIGHT      : float    = 0.15
    COLOR_WEIGHT        : float    = 0.10
    
    # Processing
    ENABLE_CACHING      : bool     = True
    PROCESSING_TIMEOUT  : int      = 30
    PARALLEL_PROCESSING : bool     = True
    MAX_WORKERS         : int      = 4
    
    # Paths
    BASE_DIR            : Path     = Path(__file__).parent.parent
    UPLOAD_DIR          : Path     = BASE_DIR / "data" / "uploads"
    REPORTS_DIR         : Path     = BASE_DIR / "data" / "reports"
    CACHE_DIR           : Path     = BASE_DIR / "data" / "cache"
    LOGS_DIR            : Path     = BASE_DIR / "logs"
    
    

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self._create_directories()
        self._validate_weights()

    
    def _create_directories(self):
        """
        Ensure all required directories exist
        """
        for directory in [self.UPLOAD_DIR, self.REPORTS_DIR, self.CACHE_DIR, self.LOGS_DIR]:
            directory.mkdir(parents  = True, 
                            exist_ok = True,
                           )
    
    def _validate_weights(self):
        """
        Validate metric weights sum to 1.0
        """
        total = (self.GRADIENT_WEIGHT +
                 self.FREQUENCY_WEIGHT +
                 self.NOISE_WEIGHT +
                 self.TEXTURE_WEIGHT +
                 self.COLOR_WEIGHT
                )

        if (not (0.99 <= total <= 1.01)):
            raise ValueError(f"Metric weights must sum to 1.0, got {total}")
    

    @property
    def max_file_size_bytes(self) -> int:
        return self.MAX_FILE_SIZE_MB * 1024 * 1024
    

    def get_metric_weights(self) -> dict:
        """
        Get all metric weights as dictionary
        """
        return {MetricType.GRADIENT  : self.GRADIENT_WEIGHT,
                MetricType.FREQUENCY : self.FREQUENCY_WEIGHT,
                MetricType.NOISE     : self.NOISE_WEIGHT,
                MetricType.TEXTURE   : self.TEXTURE_WEIGHT,
                MetricType.COLOR     : self.COLOR_WEIGHT
               }


# Singleton 
settings = Settings()
