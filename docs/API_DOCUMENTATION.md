# API Documentation

## Base Information

**Base URL**: `http://localhost:8005`  
**API Version**: `1.0.0`  
**Protocol**: HTTP/HTTPS  
**Content Type**: `application/json` (default)

---

## Table of Contents

1. [Authentication](#authentication)
2. [Health Check](#health-check)
3. [Single Image Analysis](#single-image-analysis)
4. [Batch Image Analysis](#batch-image-analysis)
5. [Batch Progress Tracking](#batch-progress-tracking)
6. [Report Export](#report-export)
7. [Error Handling](#error-handling)
8. [Rate Limits](#rate-limits)
9. [Data Models](#data-models)

---

## Authentication

**Current Version**: No authentication required (intended for internal deployment)

**Future Versions**: API key authentication planned
```bash
# Planned header format
Authorization: Bearer <api_key>
```

---

## Health Check

### `GET /health`

Check if the API server is operational.

**Request**
```bash
curl -X GET http://localhost:8005/health
```

**Response** (`200 OK`)
```json
{
  "status": "ok",
  "version": "1.0.0"
}
```

---

## Single Image Analysis

### `POST /analyze/image`

Analyze a single image for AI-generation indicators.

**Request**

```bash
curl -X POST http://localhost:8005/analyze/image \
  -F "file=@/path/to/image.jpg"
```

**Parameters**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file` | File | Yes | Image file (JPG/PNG/WEBP, max 10MB) |

**Response** (`200 OK`)

```json
{
  "success": true,
  "message": "Image analysis completed",
  "data": {
    "filename": "example.jpg",
    "status": "REVIEW_REQUIRED",
    "overall_score": 0.73,
    "confidence": 73,
    "signals": [
      {
        "name": "Gradient Field PCA",
        "metric_type": "gradient",
        "score": 0.81,
        "status": "flagged",
        "explanation": "Detected irregular gradient patterns typical of diffusion models. Natural photos show consistent lighting gradients shaped by physics."
      },
      {
        "name": "Frequency Analysis",
        "metric_type": "frequency",
        "score": 0.68,
        "status": "warning",
        "explanation": "Frequency patterns show some irregularities. Requires further review."
      },
      {
        "name": "Noise Analysis",
        "metric_type": "noise",
        "score": 0.72,
        "status": "flagged",
        "explanation": "Noise pattern is unnaturally uniform. Real camera sensors produce characteristic noise patterns."
      },
      {
        "name": "Texture Analysis",
        "metric_type": "texture",
        "score": 0.65,
        "status": "warning",
        "explanation": "Some texture regions appear overly uniform. Further analysis recommended."
      },
      {
        "name": "Color Analysis",
        "metric_type": "color",
        "score": 0.54,
        "status": "warning",
        "explanation": "Some color histogram irregularities detected."
      }
    ],
    "metric_results": {
      "gradient": {
        "metric_type": "gradient",
        "score": 0.81,
        "confidence": 0.87,
        "details": {
          "eigenvalue_ratio": 0.72,
          "gradient_vectors_sampled": 10000,
          "threshold": 0.85
        }
      },
      "frequency": {
        "metric_type": "frequency",
        "score": 0.68,
        "confidence": 0.65,
        "details": {
          "hf_ratio": 0.38,
          "hf_anomaly": 0.45,
          "roughness": 0.032,
          "spectral_deviation": 0.21
        }
      },
      "noise": {
        "metric_type": "noise",
        "score": 0.72,
        "confidence": 0.78,
        "details": {
          "mean_noise": 1.12,
          "cv": 0.18,
          "patches_valid": 42,
          "patches_total": 100
        }
      },
      "texture": {
        "metric_type": "texture",
        "score": 0.65,
        "confidence": 0.71,
        "details": {
          "smooth_ratio": 0.45,
          "contrast_mean": 18.3,
          "entropy_mean": 4.2,
          "patches_used": 50
        }
      },
      "color": {
        "metric_type": "color",
        "score": 0.54,
        "confidence": 0.58,
        "details": {
          "saturation_stats": {
            "mean_saturation": 0.68,
            "high_sat_ratio": 0.23,
            "very_high_sat_ratio": 0.06
          },
          "histogram_stats": {
            "roughness_mean": 0.021,
            "channels_analyzed": 3
          },
          "hue_stats": {
            "top3_concentration": 0.58,
            "gap_ratio": 0.32
          }
        }
      }
    },
    "processing_time": 2.34,
    "image_size": [1920, 1080],
    "timestamp": "2024-12-19T14:32:15.123456"
  },
  "timestamp": "2024-12-19T14:32:15.123456"
}
```

**Status Values**
- `LIKELY_AUTHENTIC`: Score < 0.65 (default threshold)
- `REVIEW_REQUIRED`: Score >= 0.65

**Signal Status Values**
- `passed`: Score < 0.40
- `warning`: Score >= 0.40 and < 0.70
- `flagged`: Score >= 0.70

---

## Batch Image Analysis

### `POST /analyze/batch`

Analyze multiple images in a single request with parallel processing.

**Request**

```bash
curl -X POST http://localhost:8005/analyze/batch \
  -F "files=@image1.jpg" \
  -F "files=@image2.png" \
  -F "files=@image3.webp"
```

**Parameters**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `files` | File[] | Yes | Multiple image files (max 50 per batch) |

**Response** (`200 OK`)

```json
{
  "success": true,
  "message": "Batch analysis completed",
  "data": {
    "batch_id": "550e8400-e29b-41d4-a716-446655440000",
    "result": {
      "total_images": 3,
      "processed": 3,
      "failed": 0,
      "results": [
        {
          "filename": "image1.jpg",
          "status": "REVIEW_REQUIRED",
          "overall_score": 0.73,
          "confidence": 73,
          "signals": [...],
          "metric_results": {...},
          "processing_time": 2.1,
          "image_size": [1920, 1080],
          "timestamp": "2024-12-19T14:32:15.123456"
        },
        {
          "filename": "image2.png",
          "status": "LIKELY_AUTHENTIC",
          "overall_score": 0.42,
          "confidence": 42,
          "signals": [...],
          "metric_results": {...},
          "processing_time": 2.3,
          "image_size": [2048, 1536],
          "timestamp": "2024-12-19T14:32:17.234567"
        },
        {
          "filename": "image3.webp",
          "status": "LIKELY_AUTHENTIC",
          "overall_score": 0.38,
          "confidence": 38,
          "signals": [...],
          "metric_results": {...},
          "processing_time": 1.9,
          "image_size": [1024, 768],
          "timestamp": "2024-12-19T14:32:19.345678"
        }
      ],
      "summary": {
        "likely_authentic": 2,
        "review_required": 1,
        "success_rate": 100,
        "processed": 3,
        "failed": 0,
        "avg_score": 0.510,
        "avg_confidence": 51,
        "avg_proc_time": 2.10
      },
      "total_processing_time": 6.3,
      "timestamp": "2024-12-19T14:32:19.345678"
    }
  },
  "timestamp": "2024-12-19T14:32:19.345678"
}
```

**Batch Constraints**
- Maximum images per batch: **50**
- Maximum file size per image: **10 MB**
- Timeout per image: **30 seconds**
- Total batch timeout: **15 minutes**

---

## Batch Progress Tracking

### `GET /batch/{batch_id}/progress`

Track the progress of a batch analysis job.

**Request**

```bash
curl -X GET http://localhost:8005/batch/550e8400-e29b-41d4-a716-446655440000/progress
```

**Response - Processing** (`200 OK`)

```json
{
  "status": "processing",
  "progress": {
    "current": 7,
    "total": 10,
    "filename": "image_007.jpg"
  }
}
```

**Response - Completed** (`200 OK`)

```json
{
  "status": "completed",
  "progress": {
    "current": 10,
    "total": 10,
    "filename": "image_010.jpg"
  },
  "result": {
    "total_images": 10,
    "processed": 10,
    "failed": 0,
    "results": [...],
    "summary": {...},
    "total_processing_time": 21.4,
    "timestamp": "2024-12-19T14:35:22.123456"
  }
}
```

**Response - Failed** (`200 OK`)

```json
{
  "status": "failed",
  "error": "Processing timeout exceeded"
}
```

**Status Values**
- `processing`: Batch is currently being analyzed
- `completed`: All images processed successfully
- `failed`: Batch processing encountered fatal error
- `interrupted`: Processing was manually stopped

---

## Report Export

### CSV Export

#### `GET /report/csv/{batch_id}` or `POST /report/csv/{batch_id}`

Download detailed batch analysis as CSV file.

**Request**

```bash
curl -X GET http://localhost:8005/report/csv/550e8400-e29b-41d4-a716-446655440000 \
  -o report.csv
```

**Response**

- Content-Type: `text/csv`
- File download with comprehensive analysis data
- Includes: per-image results, metric breakdowns, forensic details

**CSV Structure**
```
BATCH STATISTICS
Total Images,10
Successfully Processed,10
Failed,0
...

ANALYSIS RESULTS
Filename,Status,Overall Score,Confidence,Processing Time
image1.jpg,REVIEW_REQUIRED,0.73,73,2.1
image2.png,LIKELY_AUTHENTIC,0.42,42,2.3
...

IMAGE 1 DETAILED ANALYSIS
Metric Name,Score,Status,Explanation
Gradient Field PCA,0.81,flagged,Detected irregular gradient patterns...
...
```

---

### PDF Export

#### `GET /report/pdf/{batch_id}` or `POST /report/pdf/{batch_id}`

Download detailed batch analysis as PDF report.

**Request**

```bash
curl -X GET http://localhost:8005/report/pdf/550e8400-e29b-41d4-a716-446655440000 \
  -o report.pdf
```

**Response**

- Content-Type: `application/pdf`
- Professional formatted report with:
  - Executive summary
  - Per-image analysis sections
  - Visual metric breakdowns
  - Forensic details
  - Recommendations

---

## Error Handling

### Error Response Format

All errors return a standardized JSON structure:

```json
{
  "success": false,
  "message": "Error description",
  "error": "Detailed error message",
  "timestamp": "2024-12-19T14:32:15.123456"
}
```

### HTTP Status Codes

| Code | Meaning | Description |
|------|---------|-------------|
| `200` | OK | Request successful |
| `400` | Bad Request | Invalid input (file format, size, etc.) |
| `404` | Not Found | Batch ID not found |
| `413` | Payload Too Large | File size exceeds 10MB |
| `422` | Unprocessable Entity | Validation error |
| `499` | Client Closed Request | Processing interrupted |
| `500` | Internal Server Error | Server-side processing error |

### Common Error Scenarios

**File Too Large**
```json
{
  "success": false,
  "message": "Validation error",
  "error": "File size 12582912 bytes exceeds maximum 10485760 bytes",
  "timestamp": "2024-12-19T14:32:15.123456"
}
```

**Unsupported Format**
```json
{
  "success": false,
  "message": "Validation error",
  "error": "File extension .gif not allowed. Allowed: .jpg, .jpeg, .png, .webp",
  "timestamp": "2024-12-19T14:32:15.123456"
}
```

**Batch Not Found**
```json
{
  "success": false,
  "message": "Batch not found",
  "error": null,
  "timestamp": "2024-12-19T14:32:15.123456"
}
```

**Processing Timeout**
```json
{
  "success": false,
  "message": "Processing timeout",
  "error": "Image analysis exceeded 30 second timeout",
  "timestamp": "2024-12-19T14:32:45.123456"
}
```

---

## Rate Limits

**Current Version**: No rate limiting implemented

**Recommended Production Limits**:
- Single image analysis: **60 requests/minute per IP**
- Batch analysis: **10 requests/minute per IP**
- Report downloads: **30 requests/minute per IP**

---

## Data Models

### MetricResult

```typescript
{
  metric_type: "gradient" | "frequency" | "noise" | "texture" | "color",
  score: number,        // 0.0 - 1.0
  confidence: number,   // 0.0 - 1.0
  details: object       // Metric-specific forensic data
}
```

### DetectionSignal

```typescript
{
  name: string,
  metric_type: "gradient" | "frequency" | "noise" | "texture" | "color",
  score: number,        // 0.0 - 1.0
  status: "passed" | "warning" | "flagged",
  explanation: string
}
```

### AnalysisResult

```typescript
{
  filename: string,
  status: "LIKELY_AUTHENTIC" | "REVIEW_REQUIRED",
  overall_score: number,      // 0.0 - 1.0
  confidence: number,         // 0 - 100
  signals: DetectionSignal[],
  metric_results: {
    [key: string]: MetricResult
  },
  processing_time: number,    // seconds
  image_size: [number, number],
  timestamp: string           // ISO 8601 format
}
```

### BatchAnalysisResult

```typescript
{
  total_images: number,
  processed: number,
  failed: number,
  results: AnalysisResult[],
  summary: {
    likely_authentic: number,
    review_required: number,
    success_rate: number,     // percentage
    processed: number,
    failed: number,
    avg_score: number,
    avg_confidence: number,
    avg_proc_time: number
  },
  total_processing_time: number,
  timestamp: string
}
```

---

## Usage Examples

### Python

```python
import requests

# Single image analysis
with open('image.jpg', 'rb') as f:
    response = requests.post(
        'http://localhost:8005/analyze/image',
        files={'file': f}
    )
    result = response.json()
    print(f"Status: {result['data']['status']}")
    print(f"Score: {result['data']['overall_score']}")

# Batch analysis
files = [
    ('files', open('img1.jpg', 'rb')),
    ('files', open('img2.png', 'rb')),
    ('files', open('img3.webp', 'rb'))
]
response = requests.post(
    'http://localhost:8005/analyze/batch',
    files=files
)
batch_result = response.json()
batch_id = batch_result['data']['batch_id']

# Download CSV report
csv_response = requests.get(f'http://localhost:8005/report/csv/{batch_id}')
with open('report.csv', 'wb') as f:
    f.write(csv_response.content)
```

### JavaScript (Node.js)

```javascript
const FormData = require('form-data');
const fs = require('fs');
const axios = require('axios');

// Single image analysis
const form = new FormData();
form.append('file', fs.createReadStream('image.jpg'));

axios.post('http://localhost:8005/analyze/image', form, {
  headers: form.getHeaders()
})
.then(response => {
  console.log('Status:', response.data.data.status);
  console.log('Score:', response.data.data.overall_score);
})
.catch(error => {
  console.error('Error:', error.response.data);
});

// Batch analysis
const batchForm = new FormData();
batchForm.append('files', fs.createReadStream('img1.jpg'));
batchForm.append('files', fs.createReadStream('img2.png'));

axios.post('http://localhost:8005/analyze/batch', batchForm, {
  headers: batchForm.getHeaders()
})
.then(response => {
  const batchId = response.data.data.batch_id;
  console.log('Batch ID:', batchId);
  
  // Download PDF report
  return axios.get(`http://localhost:8005/report/pdf/${batchId}`, {
    responseType: 'arraybuffer'
  });
})
.then(pdfResponse => {
  fs.writeFileSync('report.pdf', pdfResponse.data);
  console.log('Report downloaded');
});
```

### cURL

```bash
# Single image
curl -X POST http://localhost:8005/analyze/image \
  -F "file=@image.jpg" \
  | jq '.data.status, .data.overall_score'

# Batch processing
curl -X POST http://localhost:8005/analyze/batch \
  -F "files=@img1.jpg" \
  -F "files=@img2.png" \
  -F "files=@img3.webp" \
  | jq '.data.batch_id'

# Progress tracking
curl -X GET http://localhost:8005/batch/{batch_id}/progress

# Download reports
curl -X GET http://localhost:8005/report/csv/{batch_id} -o report.csv
curl -X GET http://localhost:8005/report/pdf/{batch_id} -o report.pdf
```

---

## Changelog

### Version 1.0.0 (Current)
- Initial API release
- Single and batch image analysis
- CSV, JSON, PDF export
- Progress tracking
- Multi-metric ensemble detection

### Planned Features
- API key authentication
- Webhook callbacks for async processing
- Custom threshold configuration per request
- Historical analysis lookup
- Metrics-only API endpoints

---

*API Documentation Version: 1.0*  
*Last Updated: December 2025*  