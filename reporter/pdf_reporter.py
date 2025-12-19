# Dependencies
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple
from datetime import datetime
from utils.logger import get_logger
from config.settings import settings
from reportlab.platypus import Table, Spacer, Paragraph, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
from reportlab.platypus import TableStyle
from config.schemas import AnalysisResult, BatchAnalysisResult
from utils.helpers import generate_unique_id
from config.constants import DetectionStatus, MetricType, SignalStatus
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate
from reportlab.lib.styles import getSampleStyleSheet
from features.detailed_result_maker import DetailedResultMaker
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit


# Setup Logging
logger = get_logger(__name__)


class PDFReporter:
    """
    Enhanced Enterprise PDF Report Generator
    
    Design Philosophy:
    ------------------
    - Single page for 1 image (strict) with ALL details
    - 2 pages for 2-5 images with comprehensive details
    - Matrix-style pivot tables for 5+ images with full metric comparison
    - Visual hierarchy with color coding
    - No wasted whitespace, compact design
    - Complete details from JSON including explanations and reasons
    """
    
    # Enhanced Color scheme
    COLOR_PRIMARY = colors.HexColor('#2C3E50')      # Dark blue-grey
    COLOR_SECONDARY = colors.HexColor('#34495E')    # Darker blue
    COLOR_SUCCESS = colors.HexColor('#27AE60')      # Green
    COLOR_WARNING = colors.HexColor('#F39C12')      # Orange
    COLOR_DANGER = colors.HexColor('#E74C3C')       # Red
    COLOR_INFO = colors.HexColor('#3498DB')         # Blue
    COLOR_NEUTRAL = colors.HexColor('#95A5A6')      # Grey
    COLOR_HEADER_BG = colors.HexColor('#2C3E50')    # Dark header
    COLOR_ALT_ROW = colors.HexColor('#F8F9FA')      # Very light grey
    COLOR_ROW_HIGHLIGHT = colors.HexColor('#ECF0F1') # Light highlight
    
    def __init__(self):
        self.detailed_maker = DetailedResultMaker()
        self.styles = self._build_styles()
        logger.debug("Enhanced PDFReporter initialized")

    def export_single(self, result: AnalysisResult, output_dir: Optional[Path] = None) -> Path:
        """Export single image as comprehensive 1-page report"""
        output_dir = output_dir or settings.REPORTS_DIR
        output_dir.mkdir(parents=True, exist_ok=True)
        
        report_id = generate_unique_id()
        filename = f"single_analysis_{report_id}.pdf"
        output_path = output_dir / filename
        
        logger.info(f"Generating comprehensive single PDF: {filename}")
        
        # Use LETTER with minimal margins
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=LETTER,
            rightMargin=20,
            leftMargin=20,
            topMargin=15,
            bottomMargin=25
        )
        
        story = []
        self._add_watermarked_header(story, "Single Image Analysis Report", result.timestamp)
        self._add_comprehensive_single_image(story, result)
        self._add_footer(story)
        
        # Build with watermark
        def add_watermark(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica', 40)
            canvas.setFillColor(colors.HexColor('#F0F0F0'))
            canvas.rotate(45)
            canvas.drawString(200, 100, "AI IMAGE SCREENER")
            canvas.restoreState()
            # Add page number
            canvas.setFont('Helvetica', 8)
            canvas.setFillColor(colors.grey)
            canvas.drawRightString(LETTER[0] - 20, 15, f"Page {doc.page}")
        
        doc.build(story, onFirstPage=add_watermark, onLaterPages=add_watermark)
        return output_path

    def export_batch(self, batch_result: BatchAnalysisResult, output_dir: Optional[Path] = None) -> Path:
        """Export batch with intelligent layout based on count"""
        output_dir = output_dir or settings.REPORTS_DIR
        output_dir.mkdir(parents=True, exist_ok=True)
        
        report_id = generate_unique_id()
        filename = f"batch_analysis_{report_id}.pdf"
        output_path = output_dir / filename
        
        num_images = len(batch_result.results)
        logger.info(f"Generating batch PDF: {filename} ({num_images} images)")
        
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=LETTER,
            rightMargin=20,
            leftMargin=20,
            topMargin=15,
            bottomMargin=25
        )
        
        story = []
        self._add_watermarked_header(story, f"Batch Analysis Report ({num_images} Images)", batch_result.timestamp)
        
        if num_images == 1:
            # Single image (should use export_single but just in case)
            self._add_comprehensive_single_image(story, batch_result.results[0])
        
        elif 1 < num_images <= 5:
            # 2-page comprehensive batch report
            self._add_detailed_batch_summary(story, batch_result)
            story.append(Spacer(1, 6))
            
            # Add comprehensive details for each image
            for idx, result in enumerate(batch_result.results, 1):
                if idx > 1:
                    story.append(Spacer(1, 4))
                self._add_comprehensive_image_mini(story, result, idx, num_images)
                
                # Add page break after 2-3 images depending on content
                if idx == 3 and num_images > 3:
                    story.append(PageBreak())
                    self._add_watermarked_header(story, f"Batch Analysis Report ({num_images} Images) - Continued", batch_result.timestamp)
        
        else:
            # 5+ images: Matrix/pivot style with comprehensive tables
            self._add_batch_summary_matrix(story, batch_result)
            story.append(Spacer(1, 6))
            
            # Add comprehensive metric comparison tables
            self._add_comprehensive_metric_tables(story, batch_result.results)
            
            # Add page break if needed
            if num_images > 8:
                story.append(PageBreak())
                self._add_watermarked_header(story, f"Batch Analysis Report ({num_images} Images) - Continued", batch_result.timestamp)
                self._add_signal_summary_tables(story, batch_result.results)
        
        self._add_footer(story)
        
        # Build with watermark
        def add_watermark(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica', 40)
            canvas.setFillColor(colors.HexColor('#F0F0F0'))
            canvas.rotate(45)
            canvas.drawString(200, 100, "AI IMAGE SCREENER")
            canvas.restoreState()
            # Add page number
            canvas.setFont('Helvetica', 8)
            canvas.setFillColor(colors.grey)
            canvas.drawRightString(LETTER[0] - 20, 15, f"Page {doc.page}")
        
        doc.build(story, onFirstPage=add_watermark, onLaterPages=add_watermark)
        return output_path

    def _build_styles(self):
        """Build comprehensive style definitions"""
        styles = getSampleStyleSheet()
        
        # Title styles
        styles.add(ParagraphStyle(
            name='ReportTitle',
            fontSize=16,
            textColor=self.COLOR_PRIMARY,
            alignment=TA_CENTER,
            spaceAfter=8,
            fontName='Helvetica-Bold',
            leading=20
        ))
        
        styles.add(ParagraphStyle(
            name='SectionTitle',
            fontSize=12,
            textColor=self.COLOR_SECONDARY,
            spaceBefore=10,
            spaceAfter=6,
            fontName='Helvetica-Bold',
            leftIndent=0
        ))
        
        styles.add(ParagraphStyle(
            name='SubSectionTitle',
            fontSize=10,
            textColor=self.COLOR_INFO,
            spaceBefore=8,
            spaceAfter=4,
            fontName='Helvetica-Bold',
            leftIndent=0
        ))
        
        # Text styles
        styles.add(ParagraphStyle(
            name='CustomBodyText',
            fontSize=8,
            leading=10,
            spaceAfter=3,
            alignment=TA_LEFT
        ))
        
        styles.add(ParagraphStyle(
            name='SmallText',
            fontSize=7,
            leading=9,
            spaceAfter=2
        ))
        
        styles.add(ParagraphStyle(
            name='TableCell',
            fontSize=7,
            leading=9,
            spaceAfter=0
        ))
        
        styles.add(ParagraphStyle(
            name='TableCellBold',
            fontSize=7,
            leading=9,
            spaceAfter=0,
            fontName='Helvetica-Bold'
        ))
        
        styles.add(ParagraphStyle(
            name='TableCellSmall',
            fontSize=6.5,
            leading=8,
            spaceAfter=0
        ))
        
        styles.add(ParagraphStyle(
            name='Timestamp',
            fontSize=7,
            textColor=colors.grey,
            alignment=TA_RIGHT
        ))
        
        styles.add(ParagraphStyle(
            name='CustomBullet',
            fontSize=7,
            leading=9,
            leftIndent=10,
            spaceAfter=1
        ))
        
        styles.add(ParagraphStyle(
            name='WarningText',
            fontSize=8,
            textColor=self.COLOR_WARNING,
            leading=10,
            spaceAfter=3
        ))
        
        styles.add(ParagraphStyle(
            name='AlertText',
            fontSize=8,
            textColor=self.COLOR_DANGER,
            leading=10,
            spaceAfter=3,
            fontName='Helvetica-Bold'
        ))
        
        return styles

    def _add_watermarked_header(self, story, title: str, timestamp: datetime):
        """Header with centered title and timestamp"""
        # Title centered
        story.append(Paragraph("🔍 AI Image Screener", self.styles['ReportTitle']))
        
        # Subtitle with timestamp
        data = [[
            Paragraph(title, self.styles['SubSectionTitle']),
            Paragraph(f"Generated: {timestamp.strftime('%Y-%m-%d %H:%M:%S')}", self.styles['Timestamp'])
        ]]
        
        table = Table(data, colWidths=[400, 150])
        table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
        ]))
        story.append(table)
        
        story.append(Spacer(1, 4))

    def _add_comprehensive_single_image(self, story, result: AnalysisResult):
        """Comprehensive single page layout with ALL details"""
        
        # 1. Basic Information Table
        story.append(Paragraph("Basic Information", self.styles['SectionTitle']))
        
        overview_data = [
            [
                Paragraph("<b>Filename:</b>", self.styles['TableCellBold']),
                Paragraph(result.filename, self.styles['TableCell']),
                Paragraph("<b>Image Size:</b>", self.styles['TableCellBold']),
                Paragraph(f"{result.image_size[0]} × {result.image_size[1]} pixels", self.styles['TableCell'])
            ],
            [
                Paragraph("<b>Overall Status:</b>", self.styles['TableCellBold']),
                Paragraph(self._get_status_html(result.status.value), self.styles['TableCell']),
                Paragraph("<b>Overall Score:</b>", self.styles['TableCellBold']),
                Paragraph(f"<font color='{self._get_score_color(result.overall_score)}'><b>{result.overall_score:.3f}</b></font>", self.styles['TableCell'])
            ],
            [
                Paragraph("<b>Confidence:</b>", self.styles['TableCellBold']),
                Paragraph(f"{result.confidence}%", self.styles['TableCell']),
                Paragraph("<b>Processing Time:</b>", self.styles['TableCellBold']),
                Paragraph(f"{result.processing_time:.3f} seconds", self.styles['TableCell'])
            ],
            [
                Paragraph("<b>Analysis Timestamp:</b>", self.styles['TableCellBold']),
                Paragraph(result.timestamp.strftime("%Y-%m-%d %H:%M:%S"), self.styles['TableCell']),
                Paragraph("", self.styles['TableCell']),
                Paragraph("", self.styles['TableCell'])
            ]
        ]
        
        table = Table(overview_data, colWidths=[80, 200, 80, 200])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), self.COLOR_ALT_ROW),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3)
        ]))
        story.append(table)
        story.append(Spacer(1, 8))
        
        # 2. Detection Signals (Comprehensive)
        story.append(Paragraph("Detection Signals Analysis", self.styles['SectionTitle']))
        
        signal_data = [[
            Paragraph("<font color='white'><b>Metric</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Score</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Status</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Confidence</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Explanation</b></font>", self.styles['TableCellBold'])
        ]]
        
        for signal in result.signals:
            metric_result = result.metric_results.get(signal.metric_type)
            confidence = metric_result.confidence if metric_result and metric_result.confidence is not None else "N/A"
            
            signal_data.append([
                Paragraph(signal.name, self.styles['TableCell']),
                Paragraph(f"<b>{signal.score:.3f}</b>", self.styles['TableCell']),
                Paragraph(self._get_signal_status_html(signal.status.value), self.styles['TableCell']),
                Paragraph(f"{confidence:.3f}" if isinstance(confidence, (int, float)) else str(confidence), self.styles['TableCell']),
                Paragraph(signal.explanation, self.styles['TableCellSmall'])
            ])
        
        table = Table(signal_data, colWidths=[90, 45, 60, 55, 250])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.COLOR_HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, self.COLOR_ALT_ROW]),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3)
        ]))
        story.append(table)
        story.append(Spacer(1, 8))
        
        # 3. Detailed Forensic Analysis (All metrics with full details)
        story.append(Paragraph("Detailed Forensic Analysis", self.styles['SectionTitle']))
        
        # Process each metric type
        metric_order = ['gradient', 'frequency', 'noise', 'texture', 'color']
        metric_display_names = {
            'gradient': 'Gradient-Field PCA',
            'frequency': 'Frequency Analysis (FFT)',
            'noise': 'Noise Pattern Analysis',
            'texture': 'Texture Statistics',
            'color': 'Color Distribution'
        }
        
        for metric_key in metric_order:
            if metric_key not in result.metric_results:
                continue
                
            metric_result = result.metric_results[metric_key]
            details = metric_result.details or {}
            
            story.append(Paragraph(metric_display_names.get(metric_key, metric_key), self.styles['SubSectionTitle']))
            
            # Create metric summary row
            summary_data = [[
                Paragraph("<b>Metric Score:</b>", self.styles['TableCellBold']),
                Paragraph(f"<b>{metric_result.score:.3f}</b>", self.styles['TableCell']),
                Paragraph("<b>Confidence:</b>", self.styles['TableCellBold']),
                Paragraph(f"{metric_result.confidence:.3f}" if metric_result.confidence is not None else "N/A", self.styles['TableCell'])
            ]]
            
            summary_table = Table(summary_data, colWidths=[70, 50, 70, 50])
            summary_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), self.COLOR_ALT_ROW),
                ('GRID', (0, 0), (-1, 0), 0.5, colors.grey),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
            ]))
            story.append(summary_table)
            story.append(Spacer(1, 4))
            
            # Create detailed table for this metric
            if details:
                # Get appropriate headers and data for this metric
                headers, rows = self._get_metric_details_table(metric_key, details)
                
                if headers and rows:
                    table_data = [headers] + rows
                    col_widths = self._get_metric_column_widths(metric_key)
                    
                    detail_table = Table(table_data, colWidths=col_widths)
                    detail_table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), self.COLOR_SECONDARY),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, self.COLOR_ALT_ROW]),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 7),
                        ('LEFTPADDING', (0, 0), (-1, -1), 4),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                        ('TOPPADDING', (0, 0), (-1, -1), 2),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
                    ]))
                    story.append(detail_table)
                    story.append(Spacer(1, 6))
                else:
                    # Handle nested dictionaries
                    bullet_points = self._format_details_as_bullets(details)
                    for bullet in bullet_points:
                        story.append(Paragraph(bullet, self.styles['Bullet']))
                    story.append(Spacer(1, 4))
            else:
                story.append(Paragraph("No detailed forensic data available.", self.styles['SmallText']))
                story.append(Spacer(1, 4))
        
        # 4. Recommendation
        story.append(Paragraph("Recommendation", self.styles['SectionTitle']))
        
        if result.overall_score >= 0.85:
            rec_color = self.COLOR_DANGER
            rec_text = "⚠️ <b>CRITICAL</b>: Immediate manual verification required"
            next_steps = ["Forensic analysis", "Reverse image search", "Metadata inspection", "Expert review"]
        elif result.overall_score >= 0.70:
            rec_color = self.COLOR_WARNING
            rec_text = "⚠️ <b>HIGH RISK</b>: Manual verification recommended"
            next_steps = ["Visual inspection", "Compare with authentic samples", "Check source provenance"]
        elif result.overall_score >= 0.50:
            rec_color = colors.HexColor('#F1C40F')
            rec_text = "⚠️ <b>MEDIUM RISK</b>: Optional review suggested"
            next_steps = ["May be edited photo", "Verify image source", "Check for inconsistencies"]
        else:
            rec_color = self.COLOR_SUCCESS
            rec_text = "✅ <b>LOW RISK</b>: No immediate action required"
            next_steps = ["Proceed with normal workflow"]
        
        rec_data = [[Paragraph(rec_text, self.styles['AlertText'])]]
        rec_table = Table(rec_data, colWidths=[560])
        rec_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), rec_color),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ('RIGHTPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('ROUNDEDCORNERS', [10, 10, 10, 10])
        ]))
        story.append(rec_table)
        
        # Next steps as bullet points
        story.append(Spacer(1, 4))
        for step in next_steps:
            story.append(Paragraph(f"• {step}", self.styles['Bullet']))

    def _add_detailed_batch_summary(self, story, batch_result: BatchAnalysisResult):
        """Detailed batch summary for 2-5 images"""
        summary = batch_result.summary
        
        story.append(Paragraph("Batch Summary", self.styles['SectionTitle']))
        
        summary_data = [
            [
                Paragraph("<b>Total Images:</b>", self.styles['TableCellBold']),
                Paragraph(str(batch_result.total_images), self.styles['TableCell']),
                Paragraph("<b>Processed:</b>", self.styles['TableCellBold']),
                Paragraph(f"<font color='green'>{batch_result.processed}</font>", self.styles['TableCell']),
                Paragraph("<b>Failed:</b>", self.styles['TableCellBold']),
                Paragraph(f"<font color='red'>{batch_result.failed}</font>" if batch_result.failed > 0 else str(batch_result.failed), self.styles['TableCell'])
            ],
            [
                Paragraph("<b>Authentic:</b>", self.styles['TableCellBold']),
                Paragraph(f"<font color='green'>{summary.get('likely_authentic', 0)}</font>", self.styles['TableCell']),
                Paragraph("<b>Review Required:</b>", self.styles['TableCellBold']),
                Paragraph(f"<font color='orange'>{summary.get('review_required', 0)}</font>", self.styles['TableCell']),
                Paragraph("<b>Success Rate:</b>", self.styles['TableCellBold']),
                Paragraph(f"{summary.get('success_rate', 0)}%", self.styles['TableCell'])
            ],
            [
                Paragraph("<b>Average Score:</b>", self.styles['TableCellBold']),
                Paragraph(f"{summary.get('avg_score', 0):.3f}", self.styles['TableCell']),
                Paragraph("<b>Average Confidence:</b>", self.styles['TableCellBold']),
                Paragraph(f"{summary.get('avg_confidence', 0)}%", self.styles['TableCell']),
                Paragraph("<b>Avg Processing Time:</b>", self.styles['TableCellBold']),
                Paragraph(f"{summary.get('avg_proc_time', 0):.2f}s", self.styles['TableCell'])
            ],
            [
                Paragraph("<b>Total Processing Time:</b>", self.styles['TableCellBold']),
                Paragraph(f"{batch_result.total_processing_time:.2f}s", self.styles['TableCell']),
                Paragraph("", self.styles['TableCell']),
                Paragraph("", self.styles['TableCell']),
                Paragraph("", self.styles['TableCell']),
                Paragraph("", self.styles['TableCell'])
            ]
        ]
        
        table = Table(summary_data, colWidths=[80, 60, 80, 60, 80, 60])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), self.COLOR_ALT_ROW),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3)
        ]))
        story.append(table)

    def _add_comprehensive_image_mini(self, story, result: AnalysisResult, idx: int, total: int):
        """Comprehensive mini section for 2-5 images batch"""
        story.append(Paragraph(f"Image {idx}/{total}: {result.filename}", self.styles['SubSectionTitle']))
        
        # Basic info table
        info_data = [
            [
                Paragraph("<b>Status:</b>", self.styles['TableCellBold']),
                Paragraph(self._get_status_html(result.status.value), self.styles['TableCell']),
                Paragraph("<b>Score:</b>", self.styles['TableCellBold']),
                Paragraph(f"{result.overall_score:.3f}", self.styles['TableCell']),
                Paragraph("<b>Confidence:</b>", self.styles['TableCellBold']),
                Paragraph(f"{result.confidence}%", self.styles['TableCell'])
            ],
            [
                Paragraph("<b>Size:</b>", self.styles['TableCellBold']),
                Paragraph(f"{result.image_size[0]}×{result.image_size[1]}", self.styles['TableCell']),
                Paragraph("<b>Time:</b>", self.styles['TableCellBold']),
                Paragraph(f"{result.processing_time:.3f}s", self.styles['TableCell']),
                Paragraph("", self.styles['TableCell']),
                Paragraph("", self.styles['TableCell'])
            ]
        ]
        
        info_table = Table(info_data, colWidths=[50, 100, 50, 60, 60, 60])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), self.COLOR_ALT_ROW),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
        ]))
        story.append(info_table)
        story.append(Spacer(1, 4))
        
        # Signals summary
        signal_data = [[
            Paragraph("<font color='white'><b>Metric</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Score</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Status</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Explanation</b></font>", self.styles['TableCellBold'])
        ]]
        
        for signal in result.signals:
            signal_data.append([
                Paragraph(signal.name[:20], self.styles['TableCellSmall']),
                Paragraph(f"{signal.score:.3f}", self.styles['TableCellSmall']),
                Paragraph(self._get_signal_status_html(signal.status.value), self.styles['TableCellSmall']),
                Paragraph(signal.explanation[:60] + "..." if len(signal.explanation) > 60 else signal.explanation, self.styles['TableCellSmall'])
            ])
        
        sig_table = Table(signal_data, colWidths=[90, 45, 60, 265])
        sig_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.COLOR_HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, self.COLOR_ALT_ROW]),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('FONTSIZE', (0, 0), (-1, -1), 6.5),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
        ]))
        story.append(sig_table)
        story.append(Spacer(1, 6))

    def _add_batch_summary_matrix(self, story, batch_result: BatchAnalysisResult):
        """Matrix-style summary for 5+ images"""
        story.append(Paragraph("Batch Overview Matrix", self.styles['SectionTitle']))
        
        # Header
        header = [
            Paragraph("<font color='white'><b>#</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Filename</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Image Size</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Score</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Status</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Top Signal</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Confidence</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Time(s)</b></font>", self.styles['TableCellBold'])
        ]
        
        data = [header]
        
        for idx, result in enumerate(batch_result.results, 1):
            top_signal = max(result.signals, key=lambda s: s.score) if result.signals else None
            
            data.append([
                Paragraph(str(idx), self.styles['TableCell']),
                Paragraph(result.filename, self.styles['TableCellSmall']),
                Paragraph(f"{result.image_size[0]}×{result.image_size[1]}", self.styles['TableCell']),
                Paragraph(f"{result.overall_score:.3f}", self.styles['TableCell']),
                Paragraph(self._get_status_html(result.status.value), self.styles['TableCell']),
                Paragraph(f"{top_signal.name}: {top_signal.score:.2f}" if top_signal else "N/A", self.styles['TableCellSmall']),
                Paragraph(f"{result.confidence}%", self.styles['TableCell']),
                Paragraph(f"{result.processing_time:.2f}", self.styles['TableCell'])
            ])
        
        table = Table(data, colWidths=[25, 150, 70, 50, 70, 110, 55, 40])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.COLOR_HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, self.COLOR_ALT_ROW]),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('ALIGN', (2, 1), (7, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
        ]))
        story.append(table)

    def _add_comprehensive_metric_tables(self, story, results: List[AnalysisResult]):
        """Comprehensive metric comparison tables for 5+ images"""
        metric_configs = {
            'gradient': {
                'title': 'Gradient-Field PCA Analysis',
                'headers': ['Filename', 'Eigenvalue Ratio', 'Vectors Sampled', 'Original Pixels', 'Filtered Vectors', 'Threshold', 'Score', 'Confidence'],
                'extractors': [
                    lambda d: d.get('eigenvalue_ratio', 'N/A'),
                    lambda d: d.get('gradient_vectors_sampled', 'N/A'),
                    lambda d: d.get('original_pixels', 'N/A'),
                    lambda d: d.get('filtered_vectors', 'N/A'),
                    lambda d: d.get('threshold', 'N/A')
                ]
            },
            'frequency': {
                'title': 'Frequency Analysis (FFT)',
                'headers': ['Filename', 'HF Ratio', 'Roughness', 'Low Freq Energy', 'High Freq Energy', 'HF Anomaly', 'Spectral Deviation', 'Score', 'Confidence'],
                'extractors': [
                    lambda d: d.get('hf_ratio', 'N/A'),
                    lambda d: d.get('roughness', 'N/A'),
                    lambda d: d.get('low_freq_energy', 'N/A'),
                    lambda d: d.get('high_freq_energy', 'N/A'),
                    lambda d: d.get('hf_anomaly', 'N/A'),
                    lambda d: d.get('spectral_deviation', 'N/A')
                ]
            },
            'noise': {
                'title': 'Noise Pattern Analysis',
                'headers': ['Filename', 'Mean Noise', 'CV', 'Patches Total', 'Patches Valid', 'Noise Level Anomaly', 'Reason', 'Score', 'Confidence'],
                'extractors': [
                    lambda d: d.get('mean_noise', 'N/A'),
                    lambda d: d.get('cv', 'N/A'),
                    lambda d: d.get('patches_total', 'N/A'),
                    lambda d: d.get('patches_valid', 'N/A'),
                    lambda d: d.get('noise_level_anomaly', 'N/A'),
                    lambda d: d.get('reason', 'N/A')
                ]
            },
            'texture': {
                'title': 'Texture Statistics',
                'headers': ['Filename', 'Smooth Ratio', 'Contrast Mean', 'Entropy Mean', 'Patches Used', 'Edge Density', 'Contrast CV', 'Score', 'Confidence'],
                'extractors': [
                    lambda d: d.get('smooth_ratio', 'N/A'),
                    lambda d: d.get('contrast_mean', 'N/A'),
                    lambda d: d.get('entropy_mean', 'N/A'),
                    lambda d: d.get('patches_used', 'N/A'),
                    lambda d: d.get('edge_density_mean', 'N/A'),
                    lambda d: d.get('contrast_cv', 'N/A')
                ]
            },
            'color': {
                'title': 'Color Distribution',
                'headers': ['Filename', 'Mean Saturation', 'High Sat Ratio', 'Top3 Concentration', 'Gap Ratio', 'Histogram Roughness', 'Reason', 'Score', 'Confidence'],
                'extractors': [
                    lambda d: self._extract_color_detail(d, 'mean_saturation'),
                    lambda d: self._extract_color_detail(d, 'high_sat_ratio'),
                    lambda d: self._extract_color_detail(d, 'top3_concentration'),
                    lambda d: self._extract_color_detail(d, 'gap_ratio'),
                    lambda d: self._extract_color_detail(d, 'roughness_mean'),
                    lambda d: self._extract_color_reason(d)
                ]
            }
        }
        
        for metric_key, config in metric_configs.items():
            story.append(Spacer(1, 8))
            story.append(Paragraph(config['title'], self.styles['SectionTitle']))
            
            # Build header row
            header = [Paragraph(f"<font color='white'><b>{h}</b></font>", self.styles['TableCellBold']) for h in config['headers']]
            data = [header]
            
            # Build data rows
            for result in results:
                metric_result = result.metric_results.get(metric_key)
                if not metric_result:
                    continue
                    
                details = metric_result.details or {}
                row = [Paragraph(result.filename, self.styles['TableCellSmall'])]
                
                # Extract details using the extractor functions
                for extractor in config['extractors']:
                    value = extractor(details)
                    if isinstance(value, float):
                        row.append(Paragraph(f"{value:.3f}", self.styles['TableCell']))
                    else:
                        row.append(Paragraph(str(value), self.styles['TableCell']))
                
                # Add score and confidence
                row.append(Paragraph(f"{metric_result.score:.3f}", self.styles['TableCell']))
                row.append(Paragraph(f"{metric_result.confidence:.3f}" if metric_result.confidence else "N/A", self.styles['TableCell']))
                
                data.append(row)
            
            # Calculate column widths
            col_widths = [150] + [60] * (len(config['headers']) - 3) + [50, 50]
            
            table = Table(data, colWidths=col_widths)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), self.COLOR_HEADER_BG),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, self.COLOR_ALT_ROW]),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (1, 1), (-3, -1), 'CENTER'),
                ('FONTSIZE', (0, 0), (-1, -1), 6.5),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
            ]))
            story.append(table)

    def _add_signal_summary_tables(self, story, results: List[AnalysisResult]):
        """Signal summary tables for large batches"""
        story.append(Paragraph("Signal Status Summary", self.styles['SectionTitle']))
        
        # Count signals by status
        status_counts = {'flagged': 0, 'warning': 0, 'passed': 0}
        metric_counts = {}
        
        for result in results:
            for signal in result.signals:
                status = signal.status.value
                status_counts[status] = status_counts.get(status, 0) + 1
                
                metric_type = signal.metric_type.value
                metric_counts[metric_type] = metric_counts.get(metric_type, 0) + 1
        
        # Status summary table
        story.append(Paragraph("Signal Status Distribution", self.styles['SubSectionTitle']))
        
        status_data = [[
            Paragraph("<font color='white'><b>Status</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Count</b></font>", self.styles['TableCellBold']),
            Paragraph("<font color='white'><b>Percentage</b></font>", self.styles['TableCellBold'])
        ]]
        
        total_signals = sum(status_counts.values())
        for status, count in status_counts.items():
            percentage = (count / total_signals * 100) if total_signals > 0 else 0
            status_data.append([
                Paragraph(self._get_signal_status_html(status), self.styles['TableCell']),
                Paragraph(str(count), self.styles['TableCell']),
                Paragraph(f"{percentage:.1f}%", self.styles['TableCell'])
            ])
        
        status_table = Table(status_data, colWidths=[100, 80, 80])
        status_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.COLOR_HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, self.COLOR_ALT_ROW]),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3)
        ]))
        story.append(status_table)

    def _get_metric_details_table(self, metric_key: str, details: dict) -> Tuple[List, List]:
        """Get appropriate headers and data rows for a metric details table"""
        headers_map = {
            'gradient': ['Parameter', 'Value', 'Description'],
            'frequency': ['Parameter', 'Value', 'Description'],
            'noise': ['Parameter', 'Value', 'Description'],
            'texture': ['Parameter', 'Value', 'Description'],
            'color': ['Parameter', 'Value', 'Description']
        }
        
        # Define what to show for each metric
        metric_parameters = {
            'gradient': [
                ('eigenvalue_ratio', 'Eigenvalue Ratio', 'Ratio of eigenvalues indicating gradient alignment'),
                ('gradient_vectors_sampled', 'Vectors Sampled', 'Number of gradient vectors analyzed'),
                ('original_pixels', 'Original Pixels', 'Total pixels in the image'),
                ('filtered_vectors', 'Filtered Vectors', 'Vectors after filtering'),
                ('threshold', 'Threshold', 'Detection threshold value')
            ],
            'frequency': [
                ('hf_ratio', 'HF Ratio', 'High-frequency energy ratio'),
                ('roughness', 'Roughness', 'Spectral roughness measure'),
                ('low_freq_energy', 'Low Freq Energy', 'Low frequency energy'),
                ('high_freq_energy', 'High Freq Energy', 'High frequency energy'),
                ('hf_anomaly', 'HF Anomaly', 'High-frequency anomaly score'),
                ('spectral_deviation', 'Spectral Deviation', 'Deviation from normal spectrum')
            ],
            'noise': [
                ('mean_noise', 'Mean Noise', 'Average noise level'),
                ('cv', 'CV', 'Coefficient of variation'),
                ('patches_total', 'Patches Total', 'Total patches analyzed'),
                ('patches_valid', 'Patches Valid', 'Valid patches for analysis'),
                ('noise_level_anomaly', 'Noise Anomaly', 'Noise level anomaly score'),
                ('reason', 'Reason', 'Analysis reason or limitation')
            ],
            'texture': [
                ('smooth_ratio', 'Smooth Ratio', 'Ratio of smooth texture patches'),
                ('contrast_mean', 'Contrast Mean', 'Average texture contrast'),
                ('entropy_mean', 'Entropy Mean', 'Average texture entropy'),
                ('patches_used', 'Patches Used', 'Texture patches used in analysis'),
                ('edge_density_mean', 'Edge Density', 'Average edge density'),
                ('contrast_cv', 'Contrast CV', 'Contrast coefficient of variation')
            ],
            'color': [
                ('saturation_stats', 'Saturation Stats', 'Color saturation statistics'),
                ('histogram_stats', 'Histogram Stats', 'Color histogram statistics'),
                ('hue_stats', 'Hue Stats', 'Hue distribution statistics')
            ]
        }
        
        headers = [Paragraph(f"<font color='white'><b>{h}</b></font>", self.styles['TableCellBold']) for h in headers_map.get(metric_key, ['Parameter', 'Value'])]
        rows = []
        
        params = metric_parameters.get(metric_key, [])
        for param_key, display_name, description in params:
            if param_key in details:
                value = details[param_key]
                if isinstance(value, dict):
                    # Handle nested dictionaries
                    for sub_key, sub_value in value.items():
                        if sub_key != 'reason' or sub_value:
                            rows.append([
                                Paragraph(f"  {sub_key}", self.styles['TableCellSmall']),
                                Paragraph(str(sub_value), self.styles['TableCellSmall']),
                                Paragraph("", self.styles['TableCellSmall'])
                            ])
                else:
                    rows.append([
                        Paragraph(display_name, self.styles['TableCell']),
                        Paragraph(str(value), self.styles['TableCell']),
                        Paragraph(description, self.styles['TableCellSmall'])
                    ])
            elif param_key == 'saturation_stats' and 'saturation_stats' in details:
                sat_stats = details['saturation_stats']
                if isinstance(sat_stats, dict):
                    if 'reason' in sat_stats:
                        rows.append([
                            Paragraph('Saturation Analysis', self.styles['TableCell']),
                            Paragraph(sat_stats['reason'], self.styles['TableCell']),
                            Paragraph('Reason for saturation analysis result', self.styles['TableCellSmall'])
                        ])
                    else:
                        for stat_key, stat_value in sat_stats.items():
                            rows.append([
                                Paragraph(f"  {stat_key}", self.styles['TableCellSmall']),
                                Paragraph(str(stat_value), self.styles['TableCellSmall']),
                                Paragraph("", self.styles['TableCellSmall'])
                            ])
        
        return headers, rows

    def _get_metric_column_widths(self, metric_key: str) -> List:
        """Get appropriate column widths for metric tables"""
        width_map = {
            'gradient': [120, 80, 280],
            'frequency': [120, 80, 280],
            'noise': [120, 80, 280],
            'texture': [120, 80, 280],
            'color': [120, 80, 280]
        }
        return width_map.get(metric_key, [120, 80, 280])

    def _format_details_as_bullets(self, details: dict, indent: int = 0) -> List[str]:
        """Format nested details as bullet points"""
        bullets = []
        prefix = "  " * indent
        
        for key, value in details.items():
            if isinstance(value, dict):
                bullets.append(f"{prefix}• {key}:")
                bullets.extend(self._format_details_as_bullets(value, indent + 1))
            elif isinstance(value, list):
                bullets.append(f"{prefix}• {key}:")
                for item in value:
                    bullets.append(f"{prefix}  - {item}")
            else:
                formatted_value = f"{value:.3f}" if isinstance(value, float) else str(value)
                bullets.append(f"{prefix}• {key}: {formatted_value}")
        
        return bullets

    def _extract_color_detail(self, details: dict, key: str) -> Any:
        """Extract color detail from nested structure"""
        if 'saturation_stats' in details and isinstance(details['saturation_stats'], dict):
            if key in details['saturation_stats']:
                return details['saturation_stats'][key]
        
        if 'histogram_stats' in details and isinstance(details['histogram_stats'], dict):
            if key in details['histogram_stats']:
                return details['histogram_stats'][key]
        
        if 'hue_stats' in details and isinstance(details['hue_stats'], dict):
            if key in details['hue_stats']:
                return details['hue_stats'][key]
        
        return 'N/A'

    def _extract_color_reason(self, details: dict) -> str:
        """Extract reason from color details"""
        if 'saturation_stats' in details and isinstance(details['saturation_stats'], dict):
            if 'reason' in details['saturation_stats']:
                return details['saturation_stats']['reason']
        
        if 'hue_stats' in details and isinstance(details['hue_stats'], dict):
            if 'reason' in details['hue_stats']:
                return details['hue_stats']['reason']
        
        return ''

    def _get_status_html(self, status: str) -> str:
        """Return colored status HTML"""
        if status == "REVIEW_REQUIRED":
            return f"<font color='{self.COLOR_DANGER}'><b>⚠ REVIEW REQUIRED</b></font>"
        elif status == "LIKELY_AUTHENTIC":
            return f"<font color='{self.COLOR_SUCCESS}'><b>✓ LIKELY AUTHENTIC</b></font>"
        else:
            return f"<font color='{self.COLOR_INFO}'><b>{status}</b></font>"

    def _get_signal_status_html(self, status: str) -> str:
        """Return signal status badge HTML"""
        if status == "flagged":
            return f"<font color='{self.COLOR_DANGER}'><b>🔴 FLAGGED</b></font>"
        elif status == "warning":
            return f"<font color='{self.COLOR_WARNING}'><b>🟠 WARNING</b></font>"
        else:
            return f"<font color='{self.COLOR_SUCCESS}'><b>🟢 PASSED</b></font>"

    def _get_score_color(self, score: float) -> str:
        """Get color based on score"""
        if score >= 0.85:
            return self.COLOR_DANGER.toHex()
        elif score >= 0.70:
            return self.COLOR_WARNING.toHex()
        elif score >= 0.50:
            return colors.HexColor('#F1C40F').toHex()
        else:
            return self.COLOR_SUCCESS.toHex()

    def _add_footer(self, story):
        """Add footer with cautions and watermark notice"""
        story.append(Spacer(1, 10))
        
        # Caution box
        caution_text = "⚠️ <b>CAUTION</b>: Results are indicative and should be verified manually for critical applications. " \
                      "For questions or support, contact: support@aiimagescreener.com"
        
        caution_data = [[Paragraph(caution_text, self.styles['SmallText'])]]
        caution_table = Table(caution_data, colWidths=[560])
        caution_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#FFF3CD')),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#856404')),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#FFEEBA')),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ('RIGHTPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4)
        ]))
        story.append(caution_table)
        
        # Watermark notice
        story.append(Spacer(1, 4))
        watermark_notice = Paragraph(
            "<i>Document contains security watermark. Unauthorized duplication prohibited.</i>",
            ParagraphStyle(
                name='WatermarkNotice',
                fontSize=6,
                textColor=colors.grey,
                alignment=TA_CENTER
            )
        )
        story.append(watermark_notice)