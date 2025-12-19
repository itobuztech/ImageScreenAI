# Architecture Documentation

## Table of Contents
1. [System Overview](#system-overview)
2. [Overall Architecture](#overall-architecture)
3. [Data Pipeline](#data-pipeline)
4. [Component Details](#component-details)
5. [Product Architecture](#product-architecture)
6. [Technology Stack](#technology-stack)

---

## System Overview

ImageScreenAI is a multi-metric ensemble system designed for first-pass screening of potentially AI-generated images in production workflows. The system processes images through five independent statistical detectors, aggregates their outputs, and provides actionable binary decisions with full explainability.

**Design Principles:**
- No single metric dominates decisions
- All intermediate data preserved for explainability
- Parallel processing for batch efficiency
- Zero external ML model dependencies
- Transparent, auditable decision logic

---

## Overall Architecture

```mermaid
graph TB
    subgraph "Frontend Layer"
        UI[Web UI<br/>Single Page HTML]
    end
    
    subgraph "API Layer"
        API[FastAPI Server<br/>app.py]
        CORS[CORS Middleware]
        ERROR[Error Handler]
    end
    
    subgraph "Processing Layer"
        VALIDATOR[Image Validator<br/>utils/validators.py]
        BATCH[Batch Processor<br/>features/batch_processor.py]
        THRESH[Threshold Manager<br/>features/threshold_manager.py]
    end
    
    subgraph "Detection Layer"
        AGG[Metrics Aggregator<br/>metrics/aggregator.py]
        
        subgraph "Independent Metrics"
            M1[Gradient PCA<br/>gradient_field_pca.py]
            M2[Frequency FFT<br/>frequency_analyzer.py]
            M3[Noise Pattern<br/>noise_analyzer.py]
            M4[Texture Stats<br/>texture_analyzer.py]
            M5[Color Distribution<br/>color_analyzer.py]
        end
    end
    
    subgraph "Reporting Layer"
        DETAIL[DetailedResultMaker<br/>features/detailed_result_maker.py]
        CSV[CSV Reporter]
        JSON[JSON Reporter]
        PDF[PDF Reporter]
    end
    
    subgraph "Storage Layer"
        UPLOAD[(Temp Upload<br/>data/uploads/)]
        CACHE[(Cache<br/>data/cache/)]
        REPORTS[(Reports<br/>data/reports/)]
    end
    
    UI --> API
    API --> VALIDATOR
    VALIDATOR --> BATCH
    BATCH --> AGG
    AGG --> M1 & M2 & M3 & M4 & M5
    M1 & M2 & M3 & M4 & M5 --> AGG
    AGG --> THRESH
    THRESH --> DETAIL
    DETAIL --> CSV & JSON & PDF
    
    API -.-> UPLOAD
    BATCH -.-> CACHE
    CSV & JSON & PDF -.-> REPORTS
    
    style UI fill:#e1f5ff
    style API fill:#fff4e1
    style AGG fill:#ffe1e1
    style DETAIL fill:#e1ffe1
```

---

## Data Pipeline

```mermaid
flowchart LR
    subgraph "Input Stage"
        A[Image Upload] --> B{Validation}
        B -->|Pass| C[Temp Storage]
        B -->|Fail| Z1[Error Response]
    end
    
    subgraph "Preprocessing"
        C --> D[Load Image<br/>RGB Array]
        D --> E[Resize if Needed<br/>max 1024px]
        E --> F[Convert to<br/>Luminance]
    end
    
    subgraph "Parallel Metric Execution"
        F --> G1[Gradient<br/>Analysis]
        F --> G2[Frequency<br/>Analysis]
        F --> G3[Noise<br/>Analysis]
        F --> G4[Texture<br/>Analysis]
        F --> G5[Color<br/>Analysis]
    end
    
    subgraph "Score Aggregation"
        G1 --> H[Weighted<br/>Ensemble]
        G2 --> H
        G3 --> H
        G4 --> H
        G5 --> H
        H --> I[Overall Score<br/>0.0 - 1.0]
    end
    
    subgraph "Decision Logic"
        I --> J{Score vs<br/>Threshold}
        J -->|>= 0.65| K1[REVIEW<br/>REQUIRED]
        J -->|< 0.65| K2[LIKELY<br/>AUTHENTIC]
    end
    
    subgraph "Output Stage"
        K1 --> L[Detailed Result<br/>Assembly]
        K2 --> L
        L --> M[Signal Status<br/>Per Metric]
        M --> N[Explainability<br/>Generation]
        N --> O[Report Export<br/>CSV/JSON/PDF]
    end
    
    style B fill:#ffcccc
    style H fill:#cce5ff
    style J fill:#ffffcc
    style O fill:#ccffcc
```

---

## Component Details

### 1. Configuration Layer (`config/`)

```mermaid
classDiagram
    class Settings {
        +str APP_NAME
        +float REVIEW_THRESHOLD
        +dict METRIC_WEIGHTS
        +int MAX_WORKERS
        +get_metric_weights()
        +_validate_weights()
    }
    
    class Constants {
        <<enumeration>>
        +MetricType
        +SignalStatus
        +DetectionStatus
        +SIGNAL_THRESHOLDS
        +METRIC_EXPLANATIONS
    }
    
    class Schemas {
        +MetricResult
        +DetectionSignal
        +AnalysisResult
        +BatchAnalysisResult
    }
    
    Settings --> Constants: uses
    Schemas --> Constants: references
```

**Key Configuration Files:**
- `settings.py`: Runtime settings, environment variables, validation
- `constants.py`: Enums, thresholds, metric parameters, explanations
- `schemas.py`: Pydantic models for type safety and validation

---

### 2. Metrics Layer (`metrics/`)

```mermaid
graph TD
    subgraph "Gradient-Field PCA"
        A1[RGB → Luminance] --> A2[Sobel Gradients]
        A2 --> A3[Sample Vectors<br/>n=10000]
        A3 --> A4[PCA Analysis]
        A4 --> A5[Eigenvalue Ratio]
        A5 --> A6{Ratio < 0.85?}
        A6 -->|Yes| A7[High Suspicion]
        A6 -->|No| A8[Low Suspicion]
    end
    
    subgraph "Frequency Analysis"
        B1[Luminance] --> B2[2D FFT]
        B2 --> B3[Radial Spectrum<br/>64 bins]
        B3 --> B4[HF Energy Ratio]
        B4 --> B5[Spectral Roughness]
        B5 --> B6[Power Law Deviation]
        B6 --> B7[Weighted Anomaly]
    end
    
    subgraph "Noise Analysis"
        C1[Luminance] --> C2[Extract Patches<br/>32×32, stride=16]
        C2 --> C3[Laplacian Filter]
        C3 --> C4[MAD Estimation]
        C4 --> C5[CV Analysis]
        C5 --> C6[IQR Analysis]
        C6 --> C7[Uniformity Score]
    end
    
    style A1 fill:#ffe1e1
    style B1 fill:#e1e1ff
    style C1 fill:#e1ffe1
```

**Metric Weights (Default):**
```
Gradient:  30%
Frequency: 25%
Noise:     20%
Texture:   15%
Color:     10%
```

---

### 3. Processing Pipeline

```mermaid
sequenceDiagram
    participant UI
    participant API
    participant BatchProcessor
    participant MetricsAggregator
    participant Metric1
    participant Metric2
    participant ThresholdManager
    participant DetailedResultMaker
    
    UI->>API: Upload Batch (n images)
    API->>BatchProcessor: process_batch()
    
    loop For Each Image
        BatchProcessor->>MetricsAggregator: analyze_image()
        
        par Parallel Execution
            MetricsAggregator->>Metric1: detect()
            MetricsAggregator->>Metric2: detect()
        end
        
        Metric1-->>MetricsAggregator: MetricResult(score, confidence, details)
        Metric2-->>MetricsAggregator: MetricResult(score, confidence, details)
        
        MetricsAggregator->>MetricsAggregator: _aggregate_scores()
        MetricsAggregator->>ThresholdManager: _determine_status()
        ThresholdManager-->>MetricsAggregator: DetectionStatus
        
        MetricsAggregator-->>BatchProcessor: AnalysisResult
        BatchProcessor->>UI: Progress Update
    end
    
    BatchProcessor->>DetailedResultMaker: extract_detailed_results()
    DetailedResultMaker-->>BatchProcessor: Detailed Report Data
    
    BatchProcessor-->>API: BatchAnalysisResult
    API-->>UI: JSON Response + batch_id
```

---

### 4. Metric Execution Detail

```mermaid
flowchart TB
    subgraph "Single Metric Execution"
        A[Input: RGB Image<br/>H×W×3] --> B[Preprocessing<br/>Normalization/Conversion]
        
        B --> C[Feature Extraction]
        
        C --> D1[Sub-metric 1]
        C --> D2[Sub-metric 2]
        C --> D3[Sub-metric 3]
        
        D1 --> E[Sub-score 1<br/>0.0 - 1.0]
        D2 --> F[Sub-score 2<br/>0.0 - 1.0]
        D3 --> G[Sub-score 3<br/>0.0 - 1.0]
        
        E --> H[Weighted Combination]
        F --> H
        G --> H
        
        H --> I[Final Metric Score]
        I --> J[Confidence Calculation]
        
        J --> K[MetricResult Object]
        K --> L{Valid?}
        L -->|Yes| M[Return to Aggregator]
        L -->|No| N[Return Neutral Score<br/>0.5 + 0 confidence]
    end
    
    style A fill:#e1f5ff
    style I fill:#ffe1e1
    style K fill:#e1ffe1
```

**Example: Noise Analysis Sub-metrics**
- CV Anomaly: 40% weight
- Noise Level Anomaly: 40% weight  
- IQR Anomaly: 20% weight

---

## Product Architecture

```mermaid
graph TB
    subgraph "User Interfaces"
        WEB[Web UI<br/>Browser-based]
        API_CLIENT[API Clients<br/>Programmatic Access]
    end
    
    subgraph "Core Engine"
        SCREEN[Screening Engine<br/>Multi-metric Ensemble]
        THRESH_MGR[Threshold Manager<br/>Sensitivity Control]
    end
    
    subgraph "Reporting System"
        DETAIL[Detailed Analysis]
        EXPORT[Multi-format Export<br/>CSV/JSON/PDF]
    end
    
    subgraph "Use Cases"
        UC1[Content Moderation<br/>Pipelines]
        UC2[Journalism<br/>Verification]
        UC3[Stock Photo<br/>Platforms]
        UC4[Legal/Compliance<br/>Workflows]
    end
    
    WEB --> SCREEN
    API_CLIENT --> SCREEN
    
    SCREEN --> THRESH_MGR
    THRESH_MGR --> DETAIL
    DETAIL --> EXPORT
    
    EXPORT -.->|Feeds| UC1
    EXPORT -.->|Feeds| UC2
    EXPORT -.->|Feeds| UC3
    EXPORT -.->|Feeds| UC4
    
    style SCREEN fill:#ff6b6b
    style EXPORT fill:#4ecdc4
    style UC1 fill:#ffe66d
    style UC2 fill:#ffe66d
    style UC3 fill:#ffe66d
    style UC4 fill:#ffe66d
```

---

## Technology Stack

```mermaid
graph LR
    subgraph "Backend"
        B1[Python 3.11+]
        B2[FastAPI]
        B3[Pydantic]
        B4[NumPy/SciPy]
        B5[OpenCV]
        B6[Pillow]
    end
    
    subgraph "Frontend"
        F1[HTML5]
        F2[Vanilla JavaScript]
        F3[CSS3]
    end
    
    subgraph "Reporting"
        R1[ReportLab PDF]
        R2[CSV stdlib]
        R3[JSON stdlib]
    end
    
    subgraph "Infrastructure"
        I1[Uvicorn ASGI]
        I2[File-based Storage]
        I3[In-memory Sessions]
    end
    
    B2 --> B1
    B3 --> B1
    B4 --> B1
    B5 --> B1
    B6 --> B1
    
    F1 --> F2
    F2 --> F3
    
    R1 --> B1
    R2 --> B1
    R3 --> B1
    
    I1 --> B2
    I2 --> B1
    I3 --> B2
    
    style B1 fill:#3776ab
    style B2 fill:#009688
    style F1 fill:#e34c26
    style F2 fill:#f0db4f
```

**Key Dependencies:**
- **FastAPI**: Async API framework
- **NumPy/SciPy**: Numerical computation
- **OpenCV**: Image processing and filtering
- **Pillow**: Image loading and validation
- **ReportLab**: PDF generation
- **Pydantic**: Data validation and serialization

---

## Performance Characteristics

### Processing Times (Average)
- Single image analysis: **2-4 seconds**
- Batch processing (10 images): **15-25 seconds** (parallel)
- Report generation: **1-3 seconds**

### Resource Usage
- Memory per image: **50-150 MB**
- Max concurrent workers: **4** (configurable)
- Temp storage: **~10 MB per image**

### Scalability Considerations
- **Current**: Single-server deployment
- **Bottleneck**: CPU-bound metric computation
- **Future**: Distributed processing via task queue (Celery/RabbitMQ)

---

## Security & Privacy

1. **No data persistence**: Uploaded images deleted after processing
2. **Local processing**: No external API calls
3. **Stateless design**: No user tracking
4. **Input validation**: File type, size, dimension checks
5. **Timeout protection**: 30s per-image limit

---

## Deployment Architecture

```mermaid
graph TB
    subgraph "Production Deployment"
        LB[Load Balancer<br/>Nginx/Traefik]
        
        subgraph "Application Servers"
            APP1[FastAPI Instance 1<br/>4 workers]
            APP2[FastAPI Instance 2<br/>4 workers]
        end
        
        subgraph "Shared Storage"
            NFS[Shared NFS Mount<br/>reports/ cache/]
        end
        
        subgraph "Monitoring"
            LOGS[Log Aggregation<br/>ELK/Loki]
            METRICS[Metrics<br/>Prometheus]
        end
    end
    
    CLIENT[Clients] --> LB
    LB --> APP1
    LB --> APP2
    
    APP1 -.-> NFS
    APP2 -.-> NFS
    
    APP1 -.-> LOGS
    APP2 -.-> LOGS
    
    APP1 -.-> METRICS
    APP2 -.-> METRICS
    
    style LB fill:#4ecdc4
    style APP1 fill:#ff6b6b
    style APP2 fill:#ff6b6b
    style NFS fill:#95e1d3
```

**Recommended Setup:**
- **Web Server**: Nginx (reverse proxy)
- **App Server**: Uvicorn (ASGI)
- **Process Manager**: Systemd or Supervisor
- **Monitoring**: Prometheus + Grafana
- **Logging**: Structured JSON logs to ELK stack

---

*Document Version: 1.0*  
*Last Updated: December 2025*