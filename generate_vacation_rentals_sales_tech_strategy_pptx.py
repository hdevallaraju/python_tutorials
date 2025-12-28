from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Iterable, Optional

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


@dataclass(frozen=True)
class Theme:
    bg: RGBColor = RGBColor(10, 23, 46)  # deep navy
    panel: RGBColor = RGBColor(16, 38, 74)  # slightly lighter navy
    panel_2: RGBColor = RGBColor(19, 46, 90)  # alternate panel
    accent: RGBColor = RGBColor(32, 201, 151)  # teal
    accent_2: RGBColor = RGBColor(255, 159, 67)  # amber
    text: RGBColor = RGBColor(245, 247, 250)  # near-white
    muted: RGBColor = RGBColor(171, 184, 201)  # muted gray-blue

    title_font: str = "Calibri"
    body_font: str = "Calibri"


THEME = Theme()


def _in(x: float) -> Inches:
    return Inches(x)


def set_bg(slide, color: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_header(slide, title: str, subtitle: Optional[str] = None) -> None:
    # Title
    left, top, width, height = _in(0.7), _in(0.45), _in(12.0), _in(0.8)
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.name = THEME.title_font
    run.font.size = Pt(38)
    run.font.bold = True
    run.font.color.rgb = THEME.text

    # Accent line
    line = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, _in(0.7), _in(1.28), _in(3.6), _in(0.06)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = THEME.accent
    line.line.fill.background()

    if subtitle:
        left, top, width, height = _in(0.72), _in(1.38), _in(11.8), _in(0.5)
        stb = slide.shapes.add_textbox(left, top, width, height)
        stf = stb.text_frame
        stf.clear()
        sp = stf.paragraphs[0]
        srun = sp.add_run()
        srun.text = subtitle
        srun.font.name = THEME.body_font
        srun.font.size = Pt(16)
        srun.font.color.rgb = THEME.muted


def add_footer(slide, right_text: str, page_num: int) -> None:
    # subtle footer divider
    div = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, _in(0.7), _in(7.05), _in(12.0), _in(0.02)
    )
    div.fill.solid()
    div.fill.fore_color.rgb = RGBColor(26, 56, 102)
    div.line.fill.background()

    # left footer
    ltb = slide.shapes.add_textbox(_in(0.7), _in(7.10), _in(8.0), _in(0.35))
    ltf = ltb.text_frame
    ltf.clear()
    lp = ltf.paragraphs[0]
    lrun = lp.add_run()
    lrun.text = "Vacation Rentals — Sales & Technology Strategy (Draft)"
    lrun.font.name = THEME.body_font
    lrun.font.size = Pt(10)
    lrun.font.color.rgb = THEME.muted

    # right footer
    rtb = slide.shapes.add_textbox(_in(9.0), _in(7.10), _in(3.7), _in(0.35))
    rtf = rtb.text_frame
    rtf.clear()
    rp = rtf.paragraphs[0]
    rp.alignment = PP_ALIGN.RIGHT
    rrun = rp.add_run()
    rrun.text = f"{right_text}   |   {page_num}"
    rrun.font.name = THEME.body_font
    rrun.font.size = Pt(10)
    rrun.font.color.rgb = THEME.muted


def add_panel(slide, left: float, top: float, width: float, height: float, color: RGBColor) -> None:
    shp = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, _in(left), _in(top), _in(width), _in(height)
    )
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.color.rgb = RGBColor(39, 78, 142)
    shp.line.width = Pt(1)
    return shp


def add_panel_title(slide, text: str, left: float, top: float, width: float) -> None:
    tb = slide.shapes.add_textbox(_in(left), _in(top), _in(width), _in(0.35))
    tf = tb.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.name = THEME.body_font
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = THEME.text


def add_bullets(
    slide,
    bullets: Iterable[str],
    left: float,
    top: float,
    width: float,
    height: float,
    font_size: int = 14,
    color: RGBColor = THEME.text,
    line_spacing: float = 1.1,
) -> None:
    tb = slide.shapes.add_textbox(_in(left), _in(top), _in(width), _in(height))
    tf = tb.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP

    def _set_bullet(paragraph, char: str = "•") -> None:
        p = paragraph._p  # noqa: SLF001 (python-pptx internal)
        pPr = p.get_or_add_pPr()
        # remove existing bullet settings to avoid conflicts
        for tag in ("a:buNone", "a:buAutoNum", "a:buBlip", "a:buChar"):
            for el in pPr.findall(qn(tag)):
                pPr.remove(el)
        buChar = OxmlElement("a:buChar")
        buChar.set("char", char)
        pPr.append(buChar)

    first = True
    for item in bullets:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.text = item
        p.level = 0
        p.font.name = THEME.body_font
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.space_after = Pt(4)
        p.line_spacing = line_spacing
        _set_bullet(p)


def add_kpi_row(slide, items: list[tuple[str, str]], left: float, top: float, width: float) -> None:
    # Creates 3 KPI cards evenly.
    gap = 0.25
    card_w = (width - gap * (len(items) - 1)) / len(items)
    for i, (label, value) in enumerate(items):
        x = left + i * (card_w + gap)
        add_panel(slide, x, top, card_w, 1.15, THEME.panel_2)
        # value
        vtb = slide.shapes.add_textbox(_in(x + 0.25), _in(top + 0.18), _in(card_w - 0.5), _in(0.45))
        vtf = vtb.text_frame
        vtf.clear()
        vp = vtf.paragraphs[0]
        vp.alignment = PP_ALIGN.LEFT
        vr = vp.add_run()
        vr.text = value
        vr.font.name = THEME.title_font
        vr.font.size = Pt(24)
        vr.font.bold = True
        vr.font.color.rgb = THEME.accent
        # label
        ltb = slide.shapes.add_textbox(_in(x + 0.25), _in(top + 0.70), _in(card_w - 0.5), _in(0.35))
        ltf = ltb.text_frame
        ltf.clear()
        lp = ltf.paragraphs[0]
        lr = lp.add_run()
        lr.text = label
        lr.font.name = THEME.body_font
        lr.font.size = Pt(12)
        lr.font.color.rgb = THEME.muted


def slide_title(prs: Presentation, title: str, subtitle: str) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    set_bg(slide, THEME.bg)

    # big title
    tb = slide.shapes.add_textbox(_in(0.9), _in(2.05), _in(12.0), _in(1.2))
    tf = tb.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title
    r.font.name = THEME.title_font
    r.font.size = Pt(48)
    r.font.bold = True
    r.font.color.rgb = THEME.text

    # subtitle
    stb = slide.shapes.add_textbox(_in(0.92), _in(3.15), _in(12.0), _in(0.8))
    stf = stb.text_frame
    stf.clear()
    sp = stf.paragraphs[0]
    sr = sp.add_run()
    sr.text = subtitle
    sr.font.name = THEME.body_font
    sr.font.size = Pt(18)
    sr.font.color.rgb = THEME.muted

    # accent bar
    bar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, _in(0.9), _in(1.55), _in(1.1), _in(0.10))
    bar.fill.solid()
    bar.fill.fore_color.rgb = THEME.accent
    bar.line.fill.background()

    # small right-side motif
    motif = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, _in(10.0), _in(5.2), _in(2.7), _in(1.6))
    motif.fill.solid()
    motif.fill.fore_color.rgb = THEME.panel
    motif.line.fill.background()
    mtb = slide.shapes.add_textbox(_in(10.25), _in(5.45), _in(2.25), _in(1.1))
    mtf = mtb.text_frame
    mtf.clear()
    mp = mtf.paragraphs[0]
    mp.alignment = PP_ALIGN.LEFT
    mr = mp.add_run()
    mr.text = "Sales\n+\nTechnology"
    mr.font.name = THEME.title_font
    mr.font.size = Pt(22)
    mr.font.bold = True
    mr.font.color.rgb = THEME.text


def slide_exec_summary(prs: Presentation, page_num: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME.bg)
    add_header(slide, "Executive Summary", "What this strategy delivers for growth and operational control")

    add_kpi_row(
        slide,
        items=[
            ("Demand generation", "More qualified demand"),
            ("Conversion & revenue", "Higher direct share + RevPAR"),
            ("Operational excellence", "Faster ops + better insights"),
        ],
        left=0.7,
        top=2.1,
        width=12.0,
    )

    add_panel(slide, 0.7, 3.5, 12.0, 3.2, THEME.panel)
    add_panel_title(slide, "Strategy pillars (simple & board-level)", 1.0, 3.65, 11.4)
    add_bullets(
        slide,
        bullets=[
            "Multi-channel sales engine: OTAs + partnerships + direct booking (balanced for volume and margins).",
            "A single source of truth: PMS + Channel Manager + Booking Engine + Admin/Analytics for real-time decisions.",
            "AI-enabled growth: faster creative, smarter targeting, and lead identification — with human review and brand guardrails.",
            "Revenue management: dynamic pricing to keep “price-for-value” competitive while protecting profitability.",
        ],
        left=1.0,
        top=4.05,
        width=11.5,
        height=2.5,
        font_size=14,
        color=THEME.text,
    )

    add_footer(slide, right_text=str(date.today().strftime("%d %b %Y")), page_num=page_num)


def slide_sales_overview(prs: Presentation, page_num: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME.bg)
    add_header(slide, "Sales Strategy", "A channel mix designed for scale, margins, and repeatability")

    # Funnel blocks
    blocks = [
        ("Acquire", "Reach new high-intent travelers\nand corporate groups"),
        ("Convert", "Reduce friction in booking\nand improve trust signals"),
        ("Optimize", "Use analytics + pricing to\nimprove unit economics"),
    ]
    x, y, w, h, gap = 0.7, 2.1, 12.0, 1.55, 0.25
    bw = (w - gap * (len(blocks) - 1)) / len(blocks)
    for i, (t, d) in enumerate(blocks):
        bx = x + i * (bw + gap)
        add_panel(slide, bx, y, bw, h, THEME.panel)
        add_panel_title(slide, t, bx + 0.25, y + 0.20, bw - 0.5)
        tb = slide.shapes.add_textbox(_in(bx + 0.25), _in(y + 0.65), _in(bw - 0.5), _in(0.8))
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = d
        r.font.name = THEME.body_font
        r.font.size = Pt(12)
        r.font.color.rgb = THEME.muted

    # Key principles
    add_panel(slide, 0.7, 3.9, 12.0, 2.8, THEME.panel_2)
    add_panel_title(slide, "Guiding principles", 1.0, 4.05, 11.4)
    add_bullets(
        slide,
        bullets=[
            "Protect brand: consistent premium positioning across OTAs, partners, and direct channels.",
            "Balance volume vs. margin: use OTAs for demand, prioritize direct + repeat where possible.",
            "Track attribution: understand what drives bookings (channel, campaign, partner, lead source).",
            "Standardize partner operations: clear rate cards, commission rules, and service SLAs.",
        ],
        left=1.0,
        top=4.45,
        width=11.5,
        height=2.1,
        font_size=14,
        color=THEME.text,
    )

    add_footer(slide, right_text=str(date.today().strftime("%d %b %Y")), page_num=page_num)


def slide_main_channels(prs: Presentation, page_num: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME.bg)
    add_header(slide, "Main Sales Channels", "Purpose-driven channels with clear outcomes")

    cards = [
        (
            "OTAs (Premium + Scale)",
            "Purpose: Fill demand reliably.\nDescription: List curated inventory on premium OTAs (e.g., MMT Luxe Selection, Airbnb Luxe, goIbibo, Booking.com) to capture high-intent travelers.",
            THEME.panel,
        ),
        (
            "Collaborations (Partners)",
            "Purpose: Access qualified networks.\nDescription: Work with travel agents, luxury rental agencies, brands, and MICE firms to unlock group and repeat bookings.",
            THEME.panel_2,
        ),
        (
            "Direct Booking Portal",
            "Purpose: Improve margins + loyalty.\nDescription: A branded website/portal with a fast booking flow, trust signals, and support to increase direct share over time.",
            THEME.panel,
        ),
        (
            "Paid Social Campaigns",
            "Purpose: Generate demand quickly.\nDescription: Always-on + seasonal campaigns (Instagram/YouTube/Facebook) optimized for qualified leads and bookings.",
            THEME.panel_2,
        ),
        (
            "AI Lead Identification",
            "Purpose: Build a scalable outbound engine.\nDescription: Use AI to find & prioritize leads on LinkedIn/Instagram (profiles, interests, signals), then follow a human-led outreach playbook.",
            THEME.panel,
        ),
    ]

    # Layout: 2 columns of cards + one wide card at bottom
    left_x, right_x = 0.7, 6.85
    card_w, card_h = 5.85, 1.55
    y0, gap_y = 2.05, 0.25

    # first 4 cards
    for idx in range(4):
        cx = left_x if idx % 2 == 0 else right_x
        cy = y0 + (idx // 2) * (card_h + gap_y)
        title, body, col = cards[idx]
        add_panel(slide, cx, cy, card_w, card_h, col)
        add_panel_title(slide, title, cx + 0.25, cy + 0.18, card_w - 0.5)
        tb = slide.shapes.add_textbox(_in(cx + 0.25), _in(cy + 0.55), _in(card_w - 0.5), _in(0.95))
        tf = tb.text_frame
        tf.clear()
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = body
        r.font.name = THEME.body_font
        r.font.size = Pt(12)
        r.font.color.rgb = THEME.muted

    # bottom wide card
    title, body, col = cards[4]
    cy = y0 + 2 * (card_h + gap_y)
    add_panel(slide, 0.7, cy, 12.0, 1.55, col)
    add_panel_title(slide, title, 0.95, cy + 0.18, 11.5)
    tb = slide.shapes.add_textbox(_in(0.95), _in(cy + 0.55), _in(11.55), _in(0.95))
    tf = tb.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = body
    r.font.name = THEME.body_font
    r.font.size = Pt(12)
    r.font.color.rgb = THEME.muted

    add_footer(slide, right_text=str(date.today().strftime("%d %b %Y")), page_num=page_num)


def slide_collaborations(prs: Presentation, page_num: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME.bg)
    add_header(slide, "Collaboration Plan", "Partnerships designed to add volume, segments, and repeat business")

    cols = [
        (
            "Travel Agents & Luxury Rental Agencies",
            [
                "Purpose: Access premium clientele and ready-to-buy demand.",
                "Targets: Elite Homes; Homes & Villas by Marriott (and similar).",
                "Operating model: curated inventory + clear commission + service SLAs.",
            ],
        ),
        (
            "Vacation Rental Brands (Partial Partnerships)",
            [
                "Purpose: Expand distribution without full acquisition cost.",
                "Targets: Lohono; Amã Stays & Trails; StayVista.",
                "Model: commission-based, partial partnership on select inventory/periods.",
            ],
        ),
        (
            "MICE Companies (Groups & Corporate)",
            [
                "Purpose: Unlock high-value group bookings and offsites.",
                "Targets: SOS Party Offsite; Wizard Events; Thomas Cook; MICEKart.",
                "Model: packaged experiences + pre-negotiated rates + event coordination playbook.",
            ],
        ),
    ]

    left = 0.7
    top = 2.1
    width = 12.0
    gap = 0.25
    col_w = (width - gap * 2) / 3

    for i, (title, bullets) in enumerate(cols):
        x = left + i * (col_w + gap)
        add_panel(slide, x, top, col_w, 4.6, THEME.panel if i % 2 == 0 else THEME.panel_2)
        add_panel_title(slide, title, x + 0.25, top + 0.22, col_w - 0.5)
        add_bullets(
            slide,
            bullets=bullets,
            left=x + 0.25,
            top=top + 0.80,
            width=col_w - 0.5,
            height=3.7,
            font_size=12,
            color=THEME.muted,
            line_spacing=1.1,
        )

    # Simple governance note
    gov = slide.shapes.add_textbox(_in(0.7), _in(6.85), _in(12.0), _in(0.3))
    gtf = gov.text_frame
    gtf.clear()
    gp = gtf.paragraphs[0]
    gr = gp.add_run()
    gr.text = "Governance: rate parity rules, partner SLAs, lead attribution, and monthly performance reviews."
    gr.font.name = THEME.body_font
    gr.font.size = Pt(11)
    gr.font.color.rgb = THEME.muted

    add_footer(slide, right_text=str(date.today().strftime("%d %b %Y")), page_num=page_num)


def slide_tech_stack(prs: Presentation, page_num: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME.bg)
    add_header(slide, "Technology Stack", "Systems that enable scale, accuracy, and measurable performance")

    # Core boxes (simple architecture)
    add_panel(slide, 0.7, 2.2, 12.0, 4.7, THEME.panel)
    add_panel_title(slide, "Target architecture (high-level)", 1.0, 2.35, 11.4)

    # Row 1: Guest-facing + core commerce + core ops
    portal = add_panel(slide, 1.0, 2.95, 3.15, 1.2, THEME.panel_2)
    add_panel_title(slide, "Booking / Reservation Portal", 1.2, 3.05, 2.75)
    t1 = slide.shapes.add_textbox(_in(1.2), _in(3.38), _in(2.75), _in(0.7))
    tf1 = t1.text_frame
    tf1.text = "Purpose: drive direct bookings.\nDescription: fast search, availability, payments, support."
    for p in tf1.paragraphs:
        for r in p.runs:
            r.font.name = THEME.body_font
            r.font.size = Pt(11)
            r.font.color.rgb = THEME.muted

    engine = add_panel(slide, 4.35, 2.95, 2.55, 1.2, THEME.panel_2)
    add_panel_title(slide, "Booking Engine", 4.55, 3.05, 2.15)
    t2 = slide.shapes.add_textbox(_in(4.55), _in(3.38), _in(2.15), _in(0.7))
    tf2 = t2.text_frame
    tf2.text = "Purpose: convert reliably.\nDescription: pricing, policies, payments, confirmations."
    for p in tf2.paragraphs:
        for r in p.runs:
            r.font.name = THEME.body_font
            r.font.size = Pt(11)
            r.font.color.rgb = THEME.muted

    pms = add_panel(slide, 7.05, 2.95, 2.75, 1.2, THEME.panel_2)
    add_panel_title(slide, "PMS", 7.25, 3.05, 2.35)
    t3 = slide.shapes.add_textbox(_in(7.25), _in(3.38), _in(2.35), _in(0.7))
    tf3 = t3.text_frame
    tf3.text = "Purpose: single source of truth.\nDescription: inventory, reservations, housekeeping, owners."
    for p in tf3.paragraphs:
        for r in p.runs:
            r.font.name = THEME.body_font
            r.font.size = Pt(11)
            r.font.color.rgb = THEME.muted

    cm = add_panel(slide, 9.95, 2.95, 1.75, 1.2, THEME.panel_2)
    add_panel_title(slide, "Channel Mgr", 10.10, 3.05, 1.45)
    t6 = slide.shapes.add_textbox(_in(10.10), _in(3.38), _in(1.45), _in(0.7))
    tf6 = t6.text_frame
    tf6.text = "Purpose: distribute & sync.\nDescription: OTA/partner rate & availability sync."
    for p in tf6.paragraphs:
        for r in p.runs:
            r.font.name = THEME.body_font
            r.font.size = Pt(10)
            r.font.color.rgb = THEME.muted

    # Row 2: Admin & analytics
    aa = add_panel(slide, 1.0, 4.45, 6.05, 1.25, THEME.panel_2)
    add_panel_title(slide, "Administration & Analytics Portal", 1.2, 4.55, 5.6)
    t4 = slide.shapes.add_textbox(_in(1.2), _in(4.88), _in(5.6), _in(0.8))
    tf4 = t4.text_frame
    tf4.text = "Purpose: visibility + control.\nDescription: booking funnel, channel ROI, ops KPIs, owner reporting."
    for p in tf4.paragraphs:
        for r in p.runs:
            r.font.name = THEME.body_font
            r.font.size = Pt(11)
            r.font.color.rgb = THEME.muted

    # Row 2: Accounting
    ac = add_panel(slide, 7.35, 4.45, 4.35, 1.25, THEME.panel_2)
    add_panel_title(slide, "Accounting Integration (Tally)", 7.55, 4.55, 3.95)
    t5 = slide.shapes.add_textbox(_in(7.55), _in(4.88), _in(3.95), _in(0.8))
    tf5 = t5.text_frame
    tf5.text = "Purpose: accuracy + compliance.\nDescription: automated postings, reconciliations, owner payouts."
    for p in tf5.paragraphs:
        for r in p.runs:
            r.font.name = THEME.body_font
            r.font.size = Pt(11)
            r.font.color.rgb = THEME.muted

    # Note
    note = slide.shapes.add_textbox(_in(0.7), _in(6.95), _in(12.0), _in(0.35))
    ntf = note.text_frame
    ntf.clear()
    np = ntf.paragraphs[0]
    nr = np.add_run()
    nr.text = "Key requirement: clean integrations + consistent IDs (property, rate plan, channel) to avoid operational leakage."
    nr.font.name = THEME.body_font
    nr.font.size = Pt(11)
    nr.font.color.rgb = THEME.muted

    add_footer(slide, right_text=str(date.today().strftime("%d %b %Y")), page_num=page_num)


def slide_dynamic_pricing(prs: Presentation, page_num: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME.bg)
    add_header(slide, "Dynamic Pricing", "Balance ‘price-for-value’ while maximizing revenue per available night")

    add_panel(slide, 0.7, 2.1, 12.0, 2.0, THEME.panel)
    add_panel_title(slide, "Purpose", 1.0, 2.25, 11.4)
    add_bullets(
        slide,
        bullets=[
            "Improve RevPAR by pricing each property for demand, seasonality, events, and competitor benchmarks.",
            "Reduce manual overrides by using rule-based guardrails and exception workflows.",
        ],
        left=1.0,
        top=2.60,
        width=11.5,
        height=1.3,
        font_size=14,
        color=THEME.muted,
    )

    add_panel(slide, 0.7, 4.25, 12.0, 2.5, THEME.panel_2)
    add_panel_title(slide, "How it works (simple)", 1.0, 4.40, 11.4)
    add_bullets(
        slide,
        bullets=[
            "Inputs: occupancy pace, booking lead time, local events, competitor price ranges, property quality signals.",
            "Outputs: recommended price bands + automated updates via PMS/Channel Manager.",
            "Controls: min/max price, brand positioning tiers, blackout rules, and human approvals for high-impact changes.",
        ],
        left=1.0,
        top=4.78,
        width=11.5,
        height=1.85,
        font_size=13,
        color=THEME.text,
    )

    add_footer(slide, right_text=str(date.today().strftime("%d %b %Y")), page_num=page_num)


def slide_ai_capabilities(prs: Presentation, page_num: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME.bg)
    add_header(slide, "AI Enablement", "Speed + consistency, with human-in-the-loop controls")

    add_panel(slide, 0.7, 2.1, 12.0, 4.6, THEME.panel)
    add_panel_title(slide, "Capabilities mapped to business outcomes", 1.0, 2.25, 11.4)

    left = 1.0
    top = 2.75
    width = 11.4
    col_gap = 0.25
    col_w = (width - col_gap) / 2

    # Left column
    add_panel(slide, left, top, col_w, 1.9, THEME.panel_2)
    add_panel_title(slide, "AI Content Generation (Ads & Campaigns)", left + 0.25, top + 0.20, col_w - 0.5)
    add_bullets(
        slide,
        bullets=[
            "Purpose: faster creative production.",
            "Description: generate copy + variants + hooks tailored to persona, property, season, and channel.",
            "Control: brand tone prompts + approval before publishing.",
        ],
        left=left + 0.25,
        top=top + 0.65,
        width=col_w - 0.5,
        height=1.2,
        font_size=12,
        color=THEME.muted,
    )

    add_panel(slide, left, top + 2.05, col_w, 1.9, THEME.panel_2)
    add_panel_title(slide, "AI Workflows (Human-in-loop Posting)", left + 0.25, top + 2.25, col_w - 0.5)
    add_bullets(
        slide,
        bullets=[
            "Purpose: consistent posting cadence.",
            "Description: schedule, generate captions, propose hashtags, and adapt formats for each platform.",
            "Control: reviewer checklist + escalation for sensitive content.",
        ],
        left=left + 0.25,
        top=top + 2.70,
        width=col_w - 0.5,
        height=1.2,
        font_size=12,
        color=THEME.muted,
    )

    # Right column
    rx = left + col_w + col_gap
    add_panel(slide, rx, top, col_w, 1.9, THEME.panel_2)
    add_panel_title(slide, "AI Lead Identification (LinkedIn/Instagram)", rx + 0.25, top + 0.20, col_w - 0.5)
    add_bullets(
        slide,
        bullets=[
            "Purpose: scalable prospecting.",
            "Description: find leads using signals (role, location, interests, intent) and prioritize outreach lists.",
            "Control: outreach scripts + compliance checks.",
        ],
        left=rx + 0.25,
        top=top + 0.65,
        width=col_w - 0.5,
        height=1.2,
        font_size=12,
        color=THEME.muted,
    )

    add_panel(slide, rx, top + 2.05, col_w, 1.9, THEME.panel_2)
    add_panel_title(slide, "Measurement & Learning", rx + 0.25, top + 2.25, col_w - 0.5)
    add_bullets(
        slide,
        bullets=[
            "Purpose: improve ROI.",
            "Description: link campaigns/leads to bookings, then iterate creative, targeting, and pricing.",
            "Control: dashboards + monthly review cadence.",
        ],
        left=rx + 0.25,
        top=top + 2.70,
        width=col_w - 0.5,
        height=1.2,
        font_size=12,
        color=THEME.muted,
    )

    add_footer(slide, right_text=str(date.today().strftime("%d %b %Y")), page_num=page_num)


def slide_rollout(prs: Presentation, page_num: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME.bg)
    add_header(slide, "90-Day Rollout Plan", "Deliver quick wins while building the scalable foundation")

    add_panel(slide, 0.7, 2.1, 12.0, 4.2, THEME.panel)
    add_panel_title(slide, "Phased execution", 1.0, 2.25, 11.4)

    phases = [
        ("Weeks 1–3", "Foundation", ["Finalize channel strategy & partner targets", "Select PMS/Channel Manager/Booking Engine", "Define rate plans + brand guardrails"]),
        ("Weeks 4–7", "Build & Integrate", ["Implement PMS + Channel Manager + booking flow", "Set up analytics portal + attribution tracking", "Tally integration plan & mapping"]),
        ("Weeks 8–13", "Launch & Optimize", ["Launch campaigns + partner onboarding", "Dynamic pricing rules + review process", "AI workflows + content/lead operations cadence"]),
    ]
    left, top, width = 1.0, 2.85, 11.4
    gap = 0.25
    card_w = (width - gap * 2) / 3
    for i, (when, name, bullets) in enumerate(phases):
        x = left + i * (card_w + gap)
        add_panel(slide, x, top, card_w, 3.2, THEME.panel_2)

        # phase label
        tb = slide.shapes.add_textbox(_in(x + 0.25), _in(top + 0.20), _in(card_w - 0.5), _in(0.35))
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = when
        r.font.name = THEME.body_font
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = THEME.accent

        add_panel_title(slide, name, x + 0.25, top + 0.52, card_w - 0.5)
        add_bullets(
            slide,
            bullets=bullets,
            left=x + 0.25,
            top=top + 1.02,
            width=card_w - 0.5,
            height=2.0,
            font_size=12,
            color=THEME.muted,
        )

    # Decision asks
    add_panel(slide, 0.7, 6.45, 12.0, 0.8, THEME.panel_2)
    add_panel_title(slide, "Board decisions needed", 1.0, 6.55, 11.4)
    add_bullets(
        slide,
        bullets=[
            "Approve target channel mix and partnership priorities.",
            "Approve tech stack procurement (PMS, Channel Manager, Booking Engine, Analytics) and integration budget.",
            "Approve AI governance (human review, brand policies, data privacy).",
        ],
        left=1.0,
        top=6.85,
        width=11.5,
        height=0.45,
        font_size=11,
        color=THEME.text,
    )

    add_footer(slide, right_text=str(date.today().strftime("%d %b %Y")), page_num=page_num)


def main() -> None:
    prs = Presentation()
    # widescreen default; keep as-is for professional display

    title = "Vacation Rentals — Sales & Technology Strategy"
    subtitle = "Board-ready overview: channel mix, partnerships, and the enabling tech stack"
    slide_title(prs, title=title, subtitle=f"{subtitle}\n{date.today().strftime('%d %b %Y')}")

    page = 1
    slide_exec_summary(prs, page_num=page)
    page += 1
    slide_sales_overview(prs, page_num=page)
    page += 1
    slide_main_channels(prs, page_num=page)
    page += 1
    slide_collaborations(prs, page_num=page)
    page += 1
    slide_tech_stack(prs, page_num=page)
    page += 1
    slide_dynamic_pricing(prs, page_num=page)
    page += 1
    slide_ai_capabilities(prs, page_num=page)
    page += 1
    slide_rollout(prs, page_num=page)

    out = "/workspace/Vacation_Rentals_Sales_and_Technology_Strategy.pptx"
    prs.save(out)
    print(f"Wrote: {out}")


if __name__ == "__main__":
    main()

