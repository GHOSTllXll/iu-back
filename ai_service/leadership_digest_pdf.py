# backend/ai_service/leadership_digest_pdf.py
"""
Pipeline Modules Phase 2 (VP Sign-Off Gate & Leadership Digest): renders the
exact dict _get_leadership_digest_data() (see views.py) returns into a PDF,
for the Leadership Digest page's "Download as PDF" button.

Uses reportlab rather than WeasyPrint — per the client's explicit decision,
this server's cPanel hosting has no system-level Cairo/Pango libraries that
WeasyPrint requires, while reportlab is pure Python with no native
dependencies (see requirements.txt's own comment on this choice).

Deliberately takes the ALREADY-COMPUTED digest dict as input rather than
re-querying anything itself — the one source of truth for this data is
_get_leadership_digest_data() in views.py, so the PDF can never show
different numbers than the page it was downloaded from. This module's only
job is turning that dict into bytes.
"""
import io
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable,
)

# Brand gold used throughout the dashboard UI (#D4AF37) — kept here as the
# one accent color in an otherwise plain black-on-white printable document.
_GOLD = colors.HexColor('#B8912C')  # darkened slightly from #D4AF37 for print contrast
_DARK = colors.HexColor('#1a1a1a')
_MUTED = colors.HexColor('#666666')
_RED = colors.HexColor('#B02A2A')
_AMBER = colors.HexColor('#9C6B0A')

PIPELINE_STAGE_LABELS = {
    'new_incoming': 'New Incoming',
    'underwriting': 'Underwriting',
    'vp_review': 'VP Review',
    'offer_submitted': 'Offer Submitted',
    'dead_passed': 'Dead / Passed',
}


def _styles():
    base = getSampleStyleSheet()
    return {
        'title': ParagraphStyle('DigestTitle', parent=base['Title'], textColor=_DARK, fontSize=20, spaceAfter=2),
        'subtitle': ParagraphStyle('DigestSubtitle', parent=base['Normal'], textColor=_MUTED, fontSize=9, spaceAfter=18),
        'section': ParagraphStyle('DigestSection', parent=base['Heading2'], textColor=_GOLD, fontSize=13, spaceBefore=18, spaceAfter=8),
        'body': ParagraphStyle('DigestBody', parent=base['Normal'], fontSize=9.5, leading=13),
        'empty': ParagraphStyle('DigestEmpty', parent=base['Normal'], fontSize=9.5, textColor=_MUTED, leading=13),
    }


def _stage_counts_table(stage_counts: dict, style) -> Table:
    rows = [['Stage', 'Deals']]
    for stage_key, label in PIPELINE_STAGE_LABELS.items():
        rows.append([label, str(stage_counts.get(stage_key, 0))])

    table = Table(rows, colWidths=[3.2 * inch, 1.5 * inch])
    table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9.5),
        ('TEXTCOLOR', (0, 0), (-1, 0), _DARK),
        ('LINEBELOW', (0, 0), (-1, 0), 0.75, _DARK),
        ('LINEBELOW', (0, 1), (-1, -1), 0.25, colors.HexColor('#dddddd')),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    return table


def _awaiting_sign_off_table(rows_in: list, style) -> Table:
    rows = [['Deal', 'Workspace', 'Analyst', 'Days in VP Review']]
    for r in rows_in:
        rows.append([
            r['property_name'],
            r.get('workspace') or '—',
            r.get('assigned_analyst') or '—',
            str(r['days_in_vp_review']),
        ])
    table = Table(rows, colWidths=[2.2 * inch, 1.4 * inch, 1.9 * inch, 1.2 * inch])
    table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TEXTCOLOR', (0, 0), (-1, 0), _DARK),
        ('LINEBELOW', (0, 0), (-1, 0), 0.75, _DARK),
        ('LINEBELOW', (0, 1), (-1, -1), 0.25, colors.HexColor('#dddddd')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    return table


def _stale_deals_table(rows_in: list, style) -> Table:
    rows = [['Deal', 'Workspace', 'Stage', 'Days Since Activity']]
    for r in rows_in:
        rows.append([
            r['property_name'],
            r.get('workspace') or '—',
            r['pipeline_stage_display'],
            str(r['days_since_activity']),
        ])
    table = Table(rows, colWidths=[2.2 * inch, 1.4 * inch, 1.6 * inch, 1.5 * inch])
    table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TEXTCOLOR', (0, 0), (-1, 0), _DARK),
        ('LINEBELOW', (0, 0), (-1, 0), 0.75, _DARK),
        ('LINEBELOW', (0, 1), (-1, -1), 0.25, colors.HexColor('#dddddd')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    return table


def _variance_alerts_table(rows_in: list, style) -> Table:
    rows = [['Deal', 'Field', 'Change', 'Variance']]
    for r in rows_in:
        rows.append([
            r.get('property_name') or 'Untitled',
            r['field_label'],
            f"{r.get('original_value') or '(empty)'} → {r.get('new_value') or '(empty)'}",
            f"{r['variance_pct']:.1f}%",
        ])
    table = Table(rows, colWidths=[1.8 * inch, 1.6 * inch, 2.2 * inch, 1.1 * inch])
    table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('TEXTCOLOR', (0, 0), (-1, 0), _DARK),
        ('TEXTCOLOR', (3, 1), (3, -1), _RED),
        ('LINEBELOW', (0, 0), (-1, 0), 0.75, _DARK),
        ('LINEBELOW', (0, 1), (-1, -1), 0.25, colors.HexColor('#dddddd')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    return table


def render_leadership_digest_pdf(digest_data: dict, generated_for_email: str = '') -> bytes:
    """
    digest_data: exactly what _get_leadership_digest_data() in views.py
    returns. generated_for_email: whose digest this is (shown in the
    subtitle only — never used to re-scope or re-query anything here).
    Returns the finished PDF as bytes.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=LETTER,
        topMargin=0.75 * inch, bottomMargin=0.75 * inch,
        leftMargin=0.75 * inch, rightMargin=0.75 * inch,
    )
    style = _styles()
    story = []

    generated_at = digest_data.get('generated_at') or ''
    try:
        generated_dt = datetime.fromisoformat(generated_at)
        generated_label = generated_dt.strftime('%A, %B %d, %Y at %H:%M UTC')
    except (TypeError, ValueError):
        generated_label = generated_at

    story.append(Paragraph('Leadership Morning Digest', style['title']))
    subtitle = f"Generated {generated_label}"
    if generated_for_email:
        subtitle += f" for {generated_for_email}"
    story.append(Paragraph(subtitle, style['subtitle']))
    story.append(HRFlowable(width='100%', thickness=0.75, color=_DARK, spaceAfter=4))

    # --- Pipeline at a Glance ---
    story.append(Paragraph('Pipeline at a Glance', style['section']))
    story.append(_stage_counts_table(digest_data.get('stage_counts') or {}, style))

    # --- Awaiting VP Sign-Off ---
    story.append(Paragraph('Awaiting VP Sign-Off', style['section']))
    awaiting = digest_data.get('awaiting_sign_off') or []
    if awaiting:
        story.append(_awaiting_sign_off_table(awaiting, style))
    else:
        story.append(Paragraph('No deals currently waiting on VP Sign-Off.', style['empty']))

    # --- Stale Deals ---
    threshold = digest_data.get('stale_deal_threshold_days', 7)
    story.append(Paragraph(f'Stale Deals (no activity in {threshold}+ days)', style['section']))
    stale = digest_data.get('stale_deals') or []
    if stale:
        story.append(_stale_deals_table(stale, style))
    else:
        story.append(Paragraph('No stale deals — everything visible to you has had recent activity.', style['empty']))

    # --- Variance Alerts ---
    story.append(Paragraph('Variance Alerts', style['section']))
    alerts = digest_data.get('variance_alerts') or []
    if alerts:
        story.append(_variance_alerts_table(alerts, style))
        story.append(Spacer(1, 4))
        story.append(Paragraph(
            'Showing the 10 most recent qualifying changes. See the Deal Pipeline '
            'page for the complete variance alert feed.', style['empty']
        ))
    else:
        story.append(Paragraph('No variance alerts above threshold.', style['empty']))

    doc.build(story)
    return buf.getvalue()
