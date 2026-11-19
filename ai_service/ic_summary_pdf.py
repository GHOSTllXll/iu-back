# backend/ai_service/ic_summary_pdf.py
"""
Pipeline Modules Phase 2 follow-up: a per-deal "IC Summary" one-pager — the
frozen numbers a VP just signed off on, in a form that's easy to hand to an
Investment Committee. See ICSummaryPDFView in views.py for how this is
assembled (always from the LOCKED analysis, never the project's current
one — same "frozen, not a moving target" principle as Draft Broker Email's
use of locked_metrics).

Uses reportlab for the same reason as leadership_digest_pdf.py: this
server's cPanel hosting has no Cairo/Pango for WeasyPrint.

Deliberately takes already-computed data (the project, its lock, the
locked analysis' metrics, and the SAME _compute_deal_scorecard() /
_compute_debt_sizing() results PropertyProjectDetailView.get() uses) as
input rather than recomputing anything itself — this module's only job is
turning that data into bytes.
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

_GOLD = colors.HexColor('#B8912C')
_DARK = colors.HexColor('#1a1a1a')
_MUTED = colors.HexColor('#666666')
_RED = colors.HexColor('#B02A2A')
_GREEN = colors.HexColor('#1E7A3D')

PIPELINE_STAGE_LABELS = {
    'new_incoming': 'New Incoming',
    'underwriting': 'Underwriting',
    'vp_review': 'VP Review',
    'offer_submitted': 'Offer Submitted',
    'dead_passed': 'Dead / Passed',
}


def _fmt_money(value) -> str:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return '—'
    return f"${value:,.0f}"


def _fmt_pct(value) -> str:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return '—'
    return f"{value:.2f}%"


def _fmt_ratio(value) -> str:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return '—'
    return f"{value:.2f}x"


def _fmt_date(value) -> str:
    if not value:
        return '—'
    try:
        return datetime.fromisoformat(value).strftime('%B %d, %Y')
    except (TypeError, ValueError):
        return str(value)


def _styles():
    base = getSampleStyleSheet()
    return {
        'title': ParagraphStyle('ICTitle', parent=base['Title'], textColor=_DARK, fontSize=19, spaceAfter=2),
        'subtitle': ParagraphStyle('ICSubtitle', parent=base['Normal'], textColor=_MUTED, fontSize=9, spaceAfter=14),
        'section': ParagraphStyle('ICSection', parent=base['Heading2'], textColor=_GOLD, fontSize=12.5, spaceBefore=14, spaceAfter=6),
        'body': ParagraphStyle('ICBody', parent=base['Normal'], fontSize=9.5, leading=13),
        'flag': ParagraphStyle('ICFlag', parent=base['Normal'], fontSize=9, leading=13, textColor=_RED),
        'footer': ParagraphStyle('ICFooter', parent=base['Normal'], fontSize=7.5, textColor=_MUTED),
    }


def _kv_table(rows_in: list, col_widths=None) -> Table:
    """A simple 2-column label/value table — used for every section on
    this one-pager except the flags list."""
    table = Table(rows_in, colWidths=col_widths or [2.3 * inch, 3.4 * inch])
    table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9.5),
        ('TEXTCOLOR', (0, 0), (0, -1), _DARK),
        ('TEXTCOLOR', (1, 0), (1, -1), colors.HexColor('#222222')),
        ('LINEBELOW', (0, 0), (-1, -1), 0.25, colors.HexColor('#eeeeee')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    return table


def render_ic_summary_pdf(project, lock, metrics: dict, deal_scorecard: dict, debt_sizing,
                           waterfall_summary: dict = None) -> bytes:
    """
    project: a PropertyProject instance (visibility already checked by the
             caller — see ICSummaryPDFView).
    lock: that project's ExcelBaselineLock (guaranteed non-None by the
          caller; this is what makes the numbers below "frozen").
    metrics: the LOCKED analysis's metrics dict (lock.analysis.metrics).
    deal_scorecard: _compute_deal_scorecard()'s return value, computed by
                     the caller against the same locked analysis.
    debt_sizing: _compute_debt_sizing()'s return value (or None — no
                 usable NOI to size a loan against).
    waterfall_summary: V2 Module 3 — lock.waterfall_summary (the FROZEN
                        Scenario Builder result captured at lock time),
                        or None if the analyst never ran "Calculate"
                        before locking. See waterfall_calculator.py.
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

    metrics = metrics or {}
    deal_scorecard = deal_scorecard or {}
    property_metadata = metrics.get('property_metadata') or {}

    story.append(Paragraph('Investment Committee Summary', style['title']))
    story.append(Paragraph(
        f"{project.property_name or 'Untitled Deal'} — generated {datetime.utcnow().strftime('%B %d, %Y')}",
        style['subtitle']
    ))
    story.append(HRFlowable(width='100%', thickness=0.75, color=_DARK, spaceAfter=4))

    # --- Deal Overview ---
    story.append(Paragraph('Deal Overview', style['section']))
    story.append(_kv_table([
        ['Property', project.property_name or 'Untitled'],
        ['Address', property_metadata.get('address') or '—'],
        ['Workspace', project.workspace.name if project.workspace_id else '—'],
        ['Assigned Analyst', project.assigned_analyst.email if project.assigned_analyst_id else '—'],
        ['Pipeline Stage', PIPELINE_STAGE_LABELS.get(project.pipeline_stage, project.pipeline_stage)],
        ['Asking Price', _fmt_money(project.asking_price)],
        ['Target Closing / Offer Date', _fmt_date(
            project.target_closing_or_offer_date.isoformat() if project.target_closing_or_offer_date else None
        )],
    ]))

    # --- Underwriting Snapshot (frozen at lock time) ---
    story.append(Paragraph('Underwriting Snapshot (Frozen Baseline)', style['section']))
    price_per_unit_or_sf = deal_scorecard.get('price_per_unit_or_sf')
    basis = deal_scorecard.get('price_per_unit_or_sf_basis')
    price_per_label = (
        f"Price per {'Unit' if basis == 'unit' else 'SF'}" if basis else 'Price per Unit/SF'
    )
    story.append(_kv_table([
        ['Net Operating Income', _fmt_money(deal_scorecard.get('noi'))],
        ['In-Place Cap Rate', _fmt_pct(deal_scorecard.get('in_place_cap_rate'))],
        ['Operating Expense Ratio', _fmt_pct(deal_scorecard.get('opex_ratio'))],
        [price_per_label, _fmt_money(price_per_unit_or_sf)],
    ]))

    # --- Debt Sizing ---
    story.append(Paragraph('Debt Sizing', style['section']))
    if debt_sizing:
        story.append(_kv_table([
            ['Max Loan by DSCR', _fmt_money(debt_sizing.get('max_loan_by_dscr'))],
            ['Max Loan by LTV', _fmt_money(debt_sizing.get('max_loan_by_ltv'))],
            ['Constrained Max Loan', _fmt_money(debt_sizing.get('constrained_max_loan'))],
            ['Estimated DSCR', _fmt_ratio(debt_sizing.get('estimated_dscr'))],
            ['Assumptions Used', (
                f"{_fmt_pct(debt_sizing.get('interest_rate'))} rate, "
                f"{debt_sizing.get('amortization_years', '—')}yr amortization, "
                f"{_fmt_ratio(debt_sizing.get('min_dscr'))} min DSCR, "
                f"{_fmt_pct(debt_sizing.get('target_ltv'))} target LTV"
            )],
        ]))
    else:
        story.append(Paragraph(
            'No usable Net Operating Income to size a loan against for this deal.', style['body']
        ))

    # --- Financing & Equity Waterfall (V2 Module 3, if a scenario was
    # calculated before this baseline was locked) ---
    if waterfall_summary:
        story.append(Paragraph('Financing & Equity Waterfall (LP)', style['section']))
        story.append(_kv_table([
            ['LP IRR', _fmt_pct(waterfall_summary.get('lp_irr_pct'))],
            ['LP Equity Multiple', _fmt_ratio(waterfall_summary.get('lp_equity_multiple'))],
            ['LP Cash-on-Cash (Yr 1)', _fmt_pct(waterfall_summary.get('lp_cash_on_cash_year1_pct'))],
            ['LP Cash-on-Cash (Avg.)', _fmt_pct(waterfall_summary.get('lp_cash_on_cash_avg_pct'))],
            ['LP Equity Invested', _fmt_money(waterfall_summary.get('lp_equity_invested'))],
            ['Projected Sale Price', _fmt_money(waterfall_summary.get('sale_price'))],
        ]))
        story.append(Paragraph(
            waterfall_summary.get('modeling_convention') or '', style['footer']
        ))

    # --- Data Quality / Missing Documents (if any) ---
    missing_documents = deal_scorecard.get('missing_documents') or []
    data_quality_flags = deal_scorecard.get('data_quality_flags') or []
    if missing_documents or data_quality_flags:
        story.append(Paragraph('Data Quality Notes', style['section']))
        for doc_label in missing_documents:
            story.append(Paragraph(f"• Missing document: {doc_label}", style['flag']))
        for flag in data_quality_flags:
            story.append(Paragraph(f"• {flag['message']}", style['flag']))

    # --- Sign-Off / Lock Footer ---
    story.append(Spacer(1, 16))
    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#cccccc'), spaceAfter=6))
    locked_by_email = lock.locked_by.email if lock.locked_by else 'Unknown'
    story.append(Paragraph(
        f"Baseline locked {lock.locked_at.strftime('%B %d, %Y')} by {locked_by_email} — "
        f"snapshot {lock.snapshot_sha256[:12]}… This summary reflects that frozen baseline, "
        f"not any later re-analysis.",
        style['footer']
    ))

    doc.build(story)
    return buf.getvalue()
