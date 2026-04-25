#!/usr/bin/env python3
"""
System Design PDF Generator: Interactive Chat Agent
Uses ReportLab to produce a professional design document.
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.lib.colors import (
    HexColor, black, white, grey, lightgrey
)
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.platypus.flowables import Flowable
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.graphics.shapes import Drawing, Rect, String, Line, Polygon, Circle, Group, Path
from reportlab.graphics import renderPDF
from reportlab.graphics.charts.barcharts import VerticalBarChart
import os

# ── Brand colours ──────────────────────────────────────────────────────────────
C_DEEP_BLUE   = HexColor("#1a237e")   # headings / accent
C_MID_BLUE    = HexColor("#283593")   # sub-headings
C_LIGHT_BLUE  = HexColor("#e8eaf6")   # row tints
C_ACCENT      = HexColor("#ff6f00")   # highlights / badges
C_GREEN       = HexColor("#2e7d32")
C_TEAL        = HexColor("#00695c")
C_PURPLE      = HexColor("#6a1b9a")
C_RED         = HexColor("#c62828")
C_GREY_BG     = HexColor("#f5f5f5")
C_BORDER      = HexColor("#90a4ae")
C_TEXT        = HexColor("#212121")

PAGE_W, PAGE_H = A4
MARGIN = 2 * cm

# ── Styles ─────────────────────────────────────────────────────────────────────
def make_styles():
    base = getSampleStyleSheet()
    styles = {}

    styles['doc_title'] = ParagraphStyle(
        'doc_title', parent=base['Title'],
        fontSize=28, textColor=white, spaceAfter=6,
        fontName='Helvetica-Bold', alignment=TA_CENTER
    )
    styles['doc_subtitle'] = ParagraphStyle(
        'doc_subtitle', parent=base['Normal'],
        fontSize=13, textColor=HexColor("#bbdefb"), spaceAfter=4,
        fontName='Helvetica', alignment=TA_CENTER
    )
    styles['h1'] = ParagraphStyle(
        'h1', parent=base['Heading1'],
        fontSize=18, textColor=C_DEEP_BLUE, spaceBefore=18, spaceAfter=8,
        fontName='Helvetica-Bold', borderPad=4,
        borderWidth=0, leading=22
    )
    styles['h2'] = ParagraphStyle(
        'h2', parent=base['Heading2'],
        fontSize=13, textColor=C_MID_BLUE, spaceBefore=12, spaceAfter=6,
        fontName='Helvetica-Bold'
    )
    styles['h3'] = ParagraphStyle(
        'h3', parent=base['Heading3'],
        fontSize=11, textColor=C_TEAL, spaceBefore=8, spaceAfter=4,
        fontName='Helvetica-Bold'
    )
    styles['body'] = ParagraphStyle(
        'body', parent=base['Normal'],
        fontSize=9.5, textColor=C_TEXT, spaceAfter=5,
        fontName='Helvetica', leading=14, alignment=TA_JUSTIFY
    )
    styles['body_l'] = ParagraphStyle(
        'body_l', parent=styles['body'], alignment=TA_LEFT
    )
    styles['bullet'] = ParagraphStyle(
        'bullet', parent=base['Normal'],
        fontSize=9.5, textColor=C_TEXT, spaceAfter=3,
        fontName='Helvetica', leading=14, leftIndent=12,
        bulletIndent=0
    )
    styles['code'] = ParagraphStyle(
        'code', parent=base['Code'],
        fontSize=8.5, textColor=HexColor("#37474f"),
        backColor=C_GREY_BG, fontName='Courier',
        leftIndent=10, rightIndent=10,
        borderWidth=1, borderColor=C_BORDER,
        borderPad=6, spaceAfter=8
    )
    styles['caption'] = ParagraphStyle(
        'caption', parent=base['Normal'],
        fontSize=8, textColor=grey, alignment=TA_CENTER,
        fontName='Helvetica-Oblique', spaceAfter=6
    )
    styles['badge'] = ParagraphStyle(
        'badge', parent=base['Normal'],
        fontSize=8, textColor=white, fontName='Helvetica-Bold',
        alignment=TA_CENTER
    )
    styles['toc'] = ParagraphStyle(
        'toc', parent=base['Normal'],
        fontSize=10, textColor=C_DEEP_BLUE, spaceAfter=3,
        fontName='Helvetica', leftIndent=0
    )
    styles['toc_sub'] = ParagraphStyle(
        'toc_sub', parent=base['Normal'],
        fontSize=9, textColor=C_MID_BLUE, spaceAfter=2,
        fontName='Helvetica', leftIndent=14
    )
    return styles

# ── Custom Flowables ────────────────────────────────────────────────────────────
class TitleBlock(Flowable):
    """Full-width gradient-style title banner."""
    def __init__(self, width, height=120):
        Flowable.__init__(self)
        self.width = width
        self.height = height

    def draw(self):
        c = self.canv
        # Background gradient approximation with two rects
        c.setFillColor(C_DEEP_BLUE)
        c.rect(0, 0, self.width, self.height, fill=1, stroke=0)
        c.setFillColor(HexColor("#0d47a1"))
        c.rect(0, 0, self.width, self.height * 0.5, fill=1, stroke=0)
        # Decorative accent line
        c.setStrokeColor(C_ACCENT)
        c.setLineWidth(4)
        c.line(0, self.height - 6, self.width, self.height - 6)
        c.setLineWidth(1)
        # Title text
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 26)
        c.drawCentredString(self.width / 2, self.height - 44,
                            "Interactive Chat Agent")
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(self.width / 2, self.height - 66,
                            "System Design Document")
        c.setFillColor(HexColor("#bbdefb"))
        c.setFont("Helvetica", 9)
        c.drawCentredString(self.width / 2, self.height - 84,
                            "Version 1.0  ·  March 2026  ·  Research Mode  ·  Confidential")


class SectionBanner(Flowable):
    def __init__(self, title, width, color=None):
        Flowable.__init__(self)
        self.title = title
        self.width = width
        self.height = 26
        self.color = color or C_DEEP_BLUE

    def draw(self):
        c = self.canv
        c.setFillColor(self.color)
        c.roundRect(0, 0, self.width, self.height, 4, fill=1, stroke=0)
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(10, 8, self.title)


class ArchDiagram(Flowable):
    """Hand-drawn architecture block diagram."""
    def __init__(self, width, height=320):
        Flowable.__init__(self)
        self.width = width
        self.height = height

    def _box(self, c, x, y, w, h, label, sublabel="", fill=None, text_color=white):
        fill = fill or C_DEEP_BLUE
        c.setFillColor(fill)
        c.setStrokeColor(HexColor("#78909c"))
        c.setLineWidth(1)
        c.roundRect(x, y, w, h, 5, fill=1, stroke=1)
        c.setFillColor(text_color)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawCentredString(x + w / 2, y + h / 2 + (5 if sublabel else 1), label)
        if sublabel:
            c.setFont("Helvetica", 7)
            c.drawCentredString(x + w / 2, y + h / 2 - 7, sublabel)

    def _arrow(self, c, x1, y1, x2, y2):
        c.setStrokeColor(C_BORDER)
        c.setLineWidth(1.5)
        c.line(x1, y1, x2, y2)
        # arrowhead
        import math
        angle = math.atan2(y2 - y1, x2 - x1)
        size = 6
        c.setFillColor(C_BORDER)
        p = c.beginPath()
        p.moveTo(x2, y2)
        p.lineTo(x2 - size * math.cos(angle - 0.4),
                 y2 - size * math.sin(angle - 0.4))
        p.lineTo(x2 - size * math.cos(angle + 0.4),
                 y2 - size * math.sin(angle + 0.4))
        p.close()
        c.drawPath(p, fill=1, stroke=0)

    def draw(self):
        c = self.canv
        W = self.width
        H = self.height

        BW = 105   # box width
        BH = 36    # box height
        COL_X = [10, W*0.28, W*0.54, W*0.78]

        # ── Row 1: Client Layer ──────────────────────────────────────────
        row1_y = H - 50
        self._box(c, COL_X[0], row1_y, BW, BH, "Web Browser",    "React / Vue UI",       HexColor("#1565c0"))
        self._box(c, COL_X[1], row1_y, BW, BH, "Mobile App",     "iOS / Android",        HexColor("#1565c0"))
        self._box(c, COL_X[2], row1_y, BW, BH, "Slack / Teams",  "Webhook Adapters",     HexColor("#1565c0"))
        self._box(c, COL_X[3], row1_y, BW, BH, "API Consumer",   "REST / gRPC client",   HexColor("#1565c0"))

        c.setFillColor(HexColor("#e3f2fd"))
        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString(W/2, row1_y + BH + 5, "▲  CLIENT LAYER")

        # ── Row 2: API Gateway ────────────────────────────────────────────
        row2_y = row1_y - 60
        gw_w = W - 20
        self._box(c, 10, row2_y, gw_w, BH, "API GATEWAY  (Kong / AWS API Gateway / Nginx)",
                  "Auth · Rate Limiting · Load Balancing · SSL Termination", HexColor("#4527a0"))

        # arrows down
        for cx in [COL_X[0]+BW/2, COL_X[1]+BW/2, COL_X[2]+BW/2, COL_X[3]+BW/2]:
            self._arrow(c, cx, row1_y, cx, row2_y + BH)

        # ── Row 3: Orchestration + Context ───────────────────────────────
        row3_y = row2_y - 65
        orch_w = W * 0.52
        ctx_w  = W * 0.38
        orch_x = 10
        ctx_x  = W * 0.58

        self._box(c, orch_x, row3_y, orch_w, BH,
                  "Chat Orchestration Service",
                  "Session Mgmt · Prompt Building · Tool Dispatch",
                  HexColor("#00695c"))
        self._box(c, ctx_x, row3_y, ctx_w, BH,
                  "Context & Memory Manager",
                  "Sliding Window · Vector Store Lookup",
                  HexColor("#6a1b9a"))

        self._arrow(c, W/2, row2_y, W/2, row3_y + BH)
        self._arrow(c, orch_x + orch_w, row3_y + BH/2, ctx_x, row3_y + BH/2)

        # ── Row 4: LLM + Tools ───────────────────────────────────────────
        row4_y = row3_y - 65
        llm_w  = BW + 20
        tool_w = W - llm_w - 30

        self._box(c, 10, row4_y, llm_w, BH,
                  "LLM Engine",
                  "Claude Opus 4.6 / GPT-4o",
                  C_ACCENT, text_color=white)
        self._box(c, llm_w + 20, row4_y, tool_w, BH,
                  "Tool & Plugin Layer",
                  "WebSearch · Code Exec · Calendar · DB · CRM · Custom MCP",
                  HexColor("#c62828"))

        self._arrow(c, orch_x + orch_w/2, row3_y, orch_x + orch_w/2, row4_y + BH)
        self._arrow(c, llm_w + 20 + tool_w/2, row3_y, llm_w + 20 + tool_w/2, row4_y + BH)

        # ── Row 5: Storage + Infra ────────────────────────────────────────
        row5_y = row4_y - 65
        ncols = 4
        sw = (W - 20) / ncols - 5
        stores = [
            ("Message Store",  "PostgreSQL / DynamoDB",  HexColor("#37474f")),
            ("Vector DB",      "Pinecone / Weaviate",    HexColor("#37474f")),
            ("Cache",          "Redis / Elasticache",    HexColor("#37474f")),
            ("Blob / Files",   "S3 / GCS",               HexColor("#37474f")),
        ]
        for i, (lbl, sub, col) in enumerate(stores):
            sx = 10 + i * (sw + 5)
            self._box(c, sx, row5_y, sw, BH, lbl, sub, col)

        mid_orch = orch_x + orch_w / 2
        self._arrow(c, mid_orch, row4_y, mid_orch, row5_y + BH)

        # ── Row 6: Observability ─────────────────────────────────────────
        row6_y = row5_y - 55
        obs_w = W - 20
        self._box(c, 10, row6_y, obs_w, 28,
                  "Observability  (Prometheus · Grafana · OpenTelemetry · Sentry · LangSmith)",
                  "", HexColor("#4e342e"))

        # Layer labels on left margin
        labels = [
            (row1_y, "CLIENTS"),
            (row2_y, "GATEWAY"),
            (row3_y, "LOGIC"),
            (row4_y, "AI/TOOLS"),
            (row5_y, "STORAGE"),
            (row6_y, "OBS"),
        ]
        for (ly, lbl) in labels:
            c.saveState()
            c.setFillColor(HexColor("#78909c"))
            c.setFont("Helvetica-Bold", 6)
            c.drawRightString(8, ly + BH/2, lbl)
            c.restoreState()


class DataFlowDiagram(Flowable):
    """Sequence-style data flow."""
    def __init__(self, width, height=180):
        Flowable.__init__(self)
        self.width = width
        self.height = height

    def draw(self):
        c = self.canv
        W = self.width
        H = self.height

        actors = ["User", "Gateway", "Orchestrator", "LLM", "Tools", "DB"]
        n = len(actors)
        spacing = (W - 40) / (n - 1)
        xs = [20 + i * spacing for i in range(n)]
        colors = [HexColor("#1565c0"), HexColor("#4527a0"), HexColor("#00695c"),
                  C_ACCENT, HexColor("#c62828"), HexColor("#37474f")]

        top_y = H - 20
        bottom_y = 20

        # Lifelines
        for i, (x, col) in enumerate(zip(xs, colors)):
            c.setStrokeColor(col)
            c.setLineWidth(1.5)
            c.setDash([3, 3])
            c.line(x, top_y - 12, x, bottom_y)
            c.setDash([])
            # Actor box
            c.setFillColor(col)
            c.setStrokeColor(HexColor("#546e7a"))
            c.setLineWidth(0.5)
            bw, bh = 55, 14
            c.roundRect(x - bw/2, top_y - bh, bw, bh, 2, fill=1, stroke=1)
            c.setFillColor(white)
            c.setFont("Helvetica-Bold", 6.5)
            c.drawCentredString(x, top_y - bh + 4, actors[i])

        def seq_arrow(from_i, to_i, y, label, dashed=False, color=None):
            x1, x2 = xs[from_i], xs[to_i]
            col = color or HexColor("#455a64")
            c.setStrokeColor(col)
            c.setLineWidth(1)
            if dashed:
                c.setDash([4, 2])
            c.line(x1, y, x2, y)
            c.setDash([])
            # arrowhead
            import math
            dx = 1 if x2 > x1 else -1
            ah = 5
            c.setFillColor(col)
            p = c.beginPath()
            p.moveTo(x2, y)
            p.lineTo(x2 - dx * ah, y + 3)
            p.lineTo(x2 - dx * ah, y - 3)
            p.close()
            c.drawPath(p, fill=1, stroke=0)
            c.setFillColor(HexColor("#37474f"))
            c.setFont("Helvetica", 6)
            mid_x = (x1 + x2) / 2
            c.drawCentredString(mid_x, y + 2, label)

        step_y = top_y - 30
        step_gap = 18

        steps = [
            (0, 1, "1. HTTPS POST /chat {message, session_id}"),
            (1, 2, "2. Auth + rate-check → forward"),
            (2, 5, "3. Load session history"),
            (5, 2, "4. Return context (msgs)", True),
            (2, 3, "5. messages.create(prompt+context)"),
            (3, 4, "6. Tool call: web_search(query)", False, C_RED),
            (4, 3, "7. Search results JSON", True, C_RED),
            (3, 2, "8. Final response (stream)", True),
            (2, 5, "9. Persist assistant turn"),
            (2, 0, "10. SSE stream → user", True, HexColor("#1565c0")),
        ]

        for tup in steps:
            fi, ti = tup[0], tup[1]
            lbl = tup[2]
            dashed = tup[3] if len(tup) > 3 else False
            color = tup[4] if len(tup) > 4 else None
            seq_arrow(fi, ti, step_y, lbl, dashed, color)
            step_y -= step_gap


# ── Helper: coloured badge table cell ──────────────────────────────────────────
def badge(text, color=C_ACCENT):
    return Paragraph(
        f'<font color="white"><b>{text}</b></font>',
        ParagraphStyle('b', fontSize=7, backColor=color,
                       alignment=TA_CENTER, fontName='Helvetica-Bold',
                       borderPad=2)
    )


# ── Document builder ────────────────────────────────────────────────────────────
def build_pdf(out_path: str):
    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=MARGIN,
        title="Interactive Chat Agent – System Design",
        author="System Designer",
        subject="Architecture & Engineering Specification",
    )

    S = make_styles()
    TW = PAGE_W - 2 * MARGIN   # usable text width
    story = []

    # ── COVER ────────────────────────────────────────────────────────────────────
    story.append(TitleBlock(TW))
    story.append(Spacer(1, 18))

    meta = [
        ["Document Type", "System Design Specification"],
        ["Scope", "Interactive Chat Agent (full-stack)"],
        ["Status", "FINAL – Research Draft"],
        ["Date", "March 2026"],
        ["Effort", "Research Mode"],
        ["Author", "System Designer (AI-assisted)"],
    ]
    meta_t = Table(meta, colWidths=[TW * 0.30, TW * 0.70])
    meta_t.setStyle(TableStyle([
        ('FONTNAME',  (0,0),(-1,-1), 'Helvetica'),
        ('FONTSIZE',  (0,0),(-1,-1), 9),
        ('FONTNAME',  (0,0),(0,-1),  'Helvetica-Bold'),
        ('TEXTCOLOR', (0,0),(0,-1),  C_DEEP_BLUE),
        ('BACKGROUND',(0,0),(0,-1),  C_LIGHT_BLUE),
        ('GRID',      (0,0),(-1,-1), 0.5, C_BORDER),
        ('TOPPADDING',(0,0),(-1,-1), 5),
        ('BOTTOMPADDING',(0,0),(-1,-1), 5),
        ('LEFTPADDING',(0,0),(-1,-1), 8),
    ]))
    story.append(meta_t)
    story.append(PageBreak())

    # ── TABLE OF CONTENTS ────────────────────────────────────────────────────────
    story.append(SectionBanner("Table of Contents", TW, C_DEEP_BLUE))
    story.append(Spacer(1, 8))
    toc_entries = [
        ("1.", "Executive Summary", False),
        ("2.", "System Overview & Goals", False),
        ("3.", "High-Level Architecture Diagram", False),
        ("4.", "Component Design Specifications", False),
        ("  4.1", "Client Layer", True),
        ("  4.2", "API Gateway", True),
        ("  4.3", "Chat Orchestration Service", True),
        ("  4.4", "LLM Engine", True),
        ("  4.5", "Tool & Plugin Layer", True),
        ("  4.6", "Context & Memory Manager", True),
        ("  4.7", "Storage Layer", True),
        ("  4.8", "Observability & Monitoring", True),
        ("5.", "Data Flow Diagram (Sequence)", False),
        ("6.", "API Design & Data Schemas", False),
        ("7.", "Non-Functional Requirements", False),
        ("8.", "Scaling & Cost Estimates", False),
        ("9.", "Real-World Examples & Analogues", False),
        ("10.", "References", False),
    ]
    for (num, title, indent) in toc_entries:
        st = S['toc_sub'] if indent else S['toc']
        story.append(Paragraph(f"{num}  {title}", st))
    story.append(PageBreak())

    # ────────────────────────────────────────────────────────────────────────────
    # 1. EXECUTIVE SUMMARY
    # ────────────────────────────────────────────────────────────────────────────
    story.append(SectionBanner("1. Executive Summary", TW, C_DEEP_BLUE))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "This document presents the complete system design for a production-grade <b>Interactive Chat Agent</b> "
        "— a conversational AI system capable of handling multi-turn dialogues, calling external tools, "
        "maintaining long-lived session context, and serving multiple client surfaces (web, mobile, messaging "
        "platforms). The design targets sub-2 s median response latency, 99.9% availability, and horizontal "
        "scalability to ≥10 000 concurrent sessions.", S['body']
    ))
    story.append(Paragraph(
        "The architecture follows a <b>layered, microservices-oriented</b> pattern with a clear separation of "
        "concerns: client adapters → API gateway → chat orchestration → LLM inference → tool execution → storage. "
        "All components are stateless where possible and communicate over HTTPS/gRPC, enabling independent "
        "scaling and blue/green deployments.", S['body']
    ))

    kpis = [
        ["KPI", "Target", "Measurement"],
        ["Median response latency", "< 2 s (TTFB for streaming)", "p50 of /chat endpoint"],
        ["P99 response latency", "< 8 s", "p99 of /chat endpoint"],
        ["Availability", "99.9 % (≈43 min/month)", "Uptime monitor"],
        ["Concurrent sessions", "≥ 10 000", "Load test"],
        ["Context window utilised", "Up to 200 K tokens", "Token counter"],
        ["Tool success rate", "≥ 95 %", "Tool execution logs"],
        ["Monthly token cost", "Tracked per tenant", "Usage dashboard"],
    ]
    t = Table(kpis, colWidths=[TW*0.38, TW*0.32, TW*0.30])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_DEEP_BLUE),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8.5),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, C_LIGHT_BLUE]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 6),
    ]))
    story.append(Spacer(1, 6))
    story.append(t)
    story.append(Paragraph("Table 1 – Key Performance Indicators", S['caption']))
    story.append(PageBreak())

    # ────────────────────────────────────────────────────────────────────────────
    # 2. SYSTEM OVERVIEW & GOALS
    # ────────────────────────────────────────────────────────────────────────────
    story.append(SectionBanner("2. System Overview & Goals", TW, C_MID_BLUE))
    story.append(Spacer(1, 6))

    story.append(Paragraph("<b>What is an Interactive Chat Agent?</b>", S['h3']))
    story.append(Paragraph(
        "An interactive chat agent is a software system where a Large Language Model (LLM) serves as the "
        "reasoning core. Unlike a simple chatbot with hand-crafted decision trees, an agent can: "
        "(a) maintain multi-turn conversation history; "
        "(b) autonomously decide to call external tools (web search, databases, APIs, code execution); "
        "(c) synthesise tool results into coherent, grounded answers; and "
        "(d) escalate to human operators when confidence is low.", S['body']
    ))

    story.append(Paragraph("<b>Core Use Cases</b>", S['h3']))
    use_cases = [
        ["Use Case", "Description", "Example Products"],
        ["Customer Support Bot", "Answers FAQs, handles tickets, escalates edge cases", "Intercom Fin, Zendesk AI"],
        ["Developer Assistant", "Code completion, debugging, documentation generation", "GitHub Copilot, Cursor"],
        ["Enterprise Knowledge Bot", "RAG over internal docs, policy Q&A", "Notion AI, Glean"],
        ["Personal AI Assistant", "Calendar, email, reminders, web research", "Claude.ai, ChatGPT Plus"],
        ["Data Analysis Agent", "Run SQL, produce charts, summarise trends", "Julius AI, Hex Magic"],
    ]
    uc_t = Table(use_cases, colWidths=[TW*0.28, TW*0.42, TW*0.30])
    uc_t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_MID_BLUE),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8.5),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, C_LIGHT_BLUE]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 6),
    ]))
    story.append(uc_t)
    story.append(Paragraph("Table 2 – Core Use Cases", S['caption']))

    story.append(Paragraph("<b>Design Principles</b>", S['h3']))
    principles = [
        "• <b>Stateless services</b> – conversation state lives in the DB, not server memory",
        "• <b>Streaming-first</b> – use SSE/WebSocket to surface tokens progressively (improves perceived latency)",
        "• <b>Tool-augmented reasoning</b> – the LLM is the brain; tools are the hands",
        "• <b>Graceful degradation</b> – tool failures return errors to the LLM, not to the user",
        "• <b>Observability by default</b> – every request is traced end-to-end",
        "• <b>Multi-tenancy</b> – sessions, billing, and data isolated per tenant/user",
        "• <b>Safety guardrails</b> – prompt injection detection, content filtering, PII redaction",
    ]
    for p in principles:
        story.append(Paragraph(p, S['bullet']))
    story.append(PageBreak())

    # ────────────────────────────────────────────────────────────────────────────
    # 3. ARCHITECTURE DIAGRAM
    # ────────────────────────────────────────────────────────────────────────────
    story.append(SectionBanner("3. High-Level Architecture Diagram", TW, C_TEAL))
    story.append(Spacer(1, 8))
    story.append(ArchDiagram(TW, 330))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "Figure 1 – Six-layer architecture of the Interactive Chat Agent platform. "
        "Arrows represent synchronous HTTP/gRPC calls. The Storage layer is shared across services "
        "via internal VPC endpoints. All external traffic enters through the API Gateway.", S['caption']
    ))
    story.append(Spacer(1, 10))

    story.append(Paragraph("<b>Layer Summary</b>", S['h3']))
    layers = [
        ["Layer", "Technologies", "Responsibility"],
        ["Client Layer", "React, React Native, Slack SDK, REST/gRPC",
         "Render UI, capture user input, display streamed tokens"],
        ["API Gateway", "Kong, AWS API GW, Nginx+Lua",
         "Auth (JWT/OAuth2), rate limiting, routing, SSL termination"],
        ["Chat Orchestration", "Python FastAPI / Node.js (custom microservice)",
         "Session management, prompt assembly, tool dispatch, streaming relay"],
        ["LLM Engine", "Claude Opus 4.6, GPT-4o, Gemini (abstracted)",
         "Natural language understanding, reasoning, response generation"],
        ["Tool & Plugin Layer", "MCP servers, LangChain tools, custom functions",
         "Web search, code execution, DB queries, calendar, CRM, custom APIs"],
        ["Context & Memory", "Vector DB + Redis + sliding-window logic",
         "Retrieve relevant past context, manage token budget"],
        ["Storage", "PostgreSQL, Pinecone/Weaviate, Redis, S3",
         "Persist messages, embeddings, sessions, blobs"],
        ["Observability", "Prometheus, Grafana, OpenTelemetry, LangSmith",
         "Metrics, traces, LLM-specific evals, alerting"],
    ]
    lt = Table(layers, colWidths=[TW*0.22, TW*0.36, TW*0.42])
    lt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_TEAL),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, C_LIGHT_BLUE]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 6),
        ('VALIGN',     (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(lt)
    story.append(Paragraph("Table 3 – Layer Summary", S['caption']))
    story.append(PageBreak())

    # ────────────────────────────────────────────────────────────────────────────
    # 4. COMPONENT SPECS
    # ────────────────────────────────────────────────────────────────────────────
    story.append(SectionBanner("4. Component Design Specifications", TW, C_DEEP_BLUE))
    story.append(Spacer(1, 8))

    # 4.1 Client Layer
    story.append(Paragraph("4.1  Client Layer", S['h2']))
    story.append(Paragraph(
        "The client layer encompasses every surface through which end-users or systems interact with the agent. "
        "All clients communicate exclusively with the API Gateway via HTTPS/WSS – they have no direct knowledge "
        "of backend services.", S['body']
    ))
    client_spec = [
        ["Client", "Protocol", "Auth", "Streaming", "Libraries / SDK"],
        ["Web (React)", "HTTPS + SSE", "JWT (Bearer)", "Yes – EventSource",
         "React Query, Tailwind, shadcn/ui"],
        ["Mobile (RN)", "HTTPS + SSE", "JWT + biometrics", "Yes – fetch streams",
         "React Native, Expo"],
        ["Slack / Teams", "Webhooks (POST)", "Slack signing secret", "No (edit messages)",
         "Bolt SDK (Slack), Bot Framework (Teams)"],
        ["CLI / API consumer", "gRPC / REST", "API Key", "Yes – gRPC stream",
         "Any HTTP client"],
        ["Voice (Alexa/GH)", "HTTPS POST + TTS", "OAuth2", "No",
         "Alexa Skills Kit, SSML"],
    ]
    ct = Table(client_spec, colWidths=[TW*0.16, TW*0.13, TW*0.17, TW*0.13, TW*0.41])
    ct.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HexColor("#1565c0")),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, C_LIGHT_BLUE]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
        ('VALIGN',     (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(ct)
    story.append(Paragraph("Table 4 – Client Surface Matrix", S['caption']))

    # 4.2 API Gateway
    story.append(Paragraph("4.2  API Gateway", S['h2']))
    gw_data = [
        ["Concern", "Mechanism", "Real-World Tool"],
        ["Authentication", "JWT RS256 / API-Key header", "Kong + Keycloak / AWS Cognito"],
        ["Rate Limiting", "Token-bucket per user/IP", "Kong Rate Limiting plugin"],
        ["Load Balancing", "Round-robin + health check", "AWS ALB / Nginx upstream"],
        ["SSL Termination", "TLS 1.3, HSTS, OCSP stapling", "ACM certs + Nginx"],
        ["Request Routing", "Path-based + header-based rules", "Kong routes"],
        ["CORS", "Allowlist of origins", "Nginx headers"],
        ["Request Logging", "Structured JSON → stdout → CloudWatch", "Fluentd + CW Logs"],
        ["Prompt Injection Guard", "Regex + ML classifier at gateway level", "Custom Lambda authoriser"],
    ]
    gwt = Table(gw_data, colWidths=[TW*0.25, TW*0.40, TW*0.35])
    gwt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HexColor("#4527a0")),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, C_LIGHT_BLUE]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
    ]))
    story.append(gwt)
    story.append(Paragraph("Table 5 – API Gateway Responsibilities", S['caption']))

    # 4.3 Chat Orchestration
    story.append(Paragraph("4.3  Chat Orchestration Service", S['h2']))
    story.append(Paragraph(
        "The orchestration service is the central coordinator. It is a <b>stateless Python FastAPI</b> application "
        "(or Node.js) that receives validated requests from the gateway and performs the following sequence:", S['body']
    ))
    orch_steps = [
        "① Load or create <b>Session</b> record (session_id → PostgreSQL)",
        "② Fetch recent <b>message history</b> (last N turns or until token budget reached)",
        "③ <b>Vector search</b> – retrieve top-K relevant memory chunks from Pinecone",
        "④ <b>Assemble prompt</b> – system prompt + retrieved context + history + user message",
        "⑤ Call <b>LLM Engine</b> with tool definitions; stream response",
        "⑥ <b>Detect tool calls</b> in response → dispatch to Tool Layer → inject results",
        "⑦ <b>Relay tokens</b> to client via SSE as they arrive",
        "⑧ <b>Persist</b> the completed assistant turn to message store",
        "⑨ <b>Embed</b> the turn asynchronously (background worker) and upsert to vector DB",
    ]
    for s in orch_steps:
        story.append(Paragraph(s, S['bullet']))

    story.append(Spacer(1, 6))
    story.append(Paragraph("<b>Session Object Schema</b>", S['h3']))
    story.append(Paragraph(
        '<font face="Courier" size="8" color="#37474f">'
        'session_id   UUID PRIMARY KEY<br/>'
        'tenant_id    UUID NOT NULL<br/>'
        'user_id      UUID NOT NULL<br/>'
        'created_at   TIMESTAMPTZ NOT NULL<br/>'
        'updated_at   TIMESTAMPTZ NOT NULL<br/>'
        'metadata     JSONB          -- channel, locale, persona<br/>'
        'token_count  INTEGER        -- running total for budget mgmt<br/>'
        'state        ENUM(active, paused, archived)'
        '</font>', S['code']
    ))

    # 4.4 LLM Engine
    story.append(Paragraph("4.4  LLM Engine", S['h2']))
    story.append(Paragraph(
        "The LLM Engine is an abstraction layer that wraps one or more foundation model APIs. "
        "The default model is <b>Claude Opus 4.6</b> (Anthropic) with adaptive thinking enabled. "
        "The abstraction allows fallback to OpenAI GPT-4o or Google Gemini without code changes.", S['body']
    ))

    llm_spec = [
        ["Parameter", "Value", "Rationale"],
        ["Model", "claude-opus-4-6 (default)", "Best reasoning + tool use + 200 K context"],
        ["Thinking", "adaptive (thinking={type:'adaptive'})", "Dynamic CoT; replaces fixed budget_tokens"],
        ["Streaming", "Always enabled", "SSE streaming → low TTFB"],
        ["max_tokens", "8 192 (configurable)", "Balances cost vs. verbosity"],
        ["System prompt", "≤ 2 000 tokens (cached)", "Persona, safety rules, tool guidance"],
        ["Tool choice", "auto (model decides)", "Allows zero-tool responses when appropriate"],
        ["Context window", "200 K tokens (1 M beta)", "Supports very long conversations"],
        ["Prompt caching", "Ephemeral cache on system prompt", "~90 % reduction on cached tokens"],
        ["Fallback model", "claude-haiku-4-5", "Cost guard for simple classification tasks"],
        ["Cost (Opus 4.6)", "$5/1M input · $25/1M output", "Plan for ~1 K tokens avg per turn"],
    ]
    lt2 = Table(llm_spec, colWidths=[TW*0.28, TW*0.38, TW*0.34])
    lt2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_ACCENT),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, HexColor("#fff8e1")]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
    ]))
    story.append(lt2)
    story.append(Paragraph("Table 6 – LLM Engine Configuration", S['caption']))
    story.append(PageBreak())

    # 4.5 Tool Layer
    story.append(Paragraph("4.5  Tool & Plugin Layer", S['h2']))
    story.append(Paragraph(
        "Tools extend the LLM's capabilities beyond text generation. Each tool is defined by a "
        "<b>JSON Schema</b> and exposed either as a native function (beta tool runner) or via an "
        "<b>MCP (Model Context Protocol)</b> server. The orchestrator executes tool calls and "
        "returns results to the LLM in the next API turn.", S['body']
    ))
    tools_data = [
        ["Tool Name", "Type", "Trigger", "Backend", "Latency Budget"],
        ["web_search", "Server-side (Anthropic)", "Research questions", "Anthropic infra", "< 3 s"],
        ["web_fetch", "Server-side (Anthropic)", "URL retrieval", "Anthropic infra", "< 5 s"],
        ["code_execution", "Server-side (Anthropic)", "Math, data analysis", "Sandboxed container", "< 10 s"],
        ["sql_query", "User-defined", "DB lookups", "Read-only replica", "< 2 s"],
        ["calendar_event", "User-defined / MCP", "Scheduling", "Google Calendar API", "< 1 s"],
        ["send_email", "User-defined (guarded)", "Notifications", "SendGrid / SES", "< 2 s"],
        ["crm_lookup", "User-defined", "Customer data", "Salesforce REST", "< 2 s"],
        ["vector_search", "User-defined", "Memory retrieval", "Pinecone gRPC", "< 200 ms"],
        ["image_gen", "User-defined", "Visual content", "DALL-E / Stable Diff", "< 15 s"],
        ["human_escalate", "User-defined", "Low confidence", "Intercom ticket API", "< 500 ms"],
    ]
    tt = Table(tools_data, colWidths=[TW*0.20, TW*0.20, TW*0.22, TW*0.22, TW*0.16])
    tt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HexColor("#c62828")),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 7.5),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, HexColor("#ffebee")]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING',(0,0),(-1,-1), 3),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
    ]))
    story.append(tt)
    story.append(Paragraph("Table 7 – Tool Catalogue", S['caption']))

    # 4.6 Context & Memory
    story.append(Paragraph("4.6  Context & Memory Manager", S['h2']))
    story.append(Paragraph(
        "Managing context is critical for coherence over long sessions. The memory manager implements "
        "a three-tier strategy:", S['body']
    ))
    mem_tiers = [
        ["Tier", "Mechanism", "Scope", "Latency"],
        ["In-Context (Hot)", "Last N turns in the API prompt", "Current session only", "0 ms"],
        ["Episodic (Warm)", "Embedding search in vector DB", "All past sessions of the user", "50–200 ms"],
        ["Summary (Cold)", "LLM-generated summary stored in DB", "Sessions > 20 turns", "Async"],
    ]
    mt = Table(mem_tiers, colWidths=[TW*0.22, TW*0.36, TW*0.28, TW*0.14])
    mt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HexColor("#6a1b9a")),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, HexColor("#f3e5f5")]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
    ]))
    story.append(mt)
    story.append(Paragraph("Table 8 – Memory Tiers", S['caption']))

    story.append(Paragraph(
        "<b>Token Budget Algorithm:</b> The orchestrator targets a token budget "
        "(e.g. 50 K tokens) for the assembled prompt. It fills slots in order: "
        "system prompt → summary → episodic chunks (ranked by relevance) → recent turns. "
        "If the budget is exceeded, oldest history is dropped first.", S['body']
    ))

    # 4.7 Storage
    story.append(Paragraph("4.7  Storage Layer", S['h2']))
    storage_data = [
        ["Store", "Product (AWS / GCP)", "Data Model", "Access Pattern"],
        ["Relational DB", "PostgreSQL / Aurora", "sessions, messages, users, tenants", "OLTP reads/writes"],
        ["Vector DB", "Pinecone / Weaviate", "Embedding vectors + metadata", "ANN similarity search"],
        ["Cache", "Redis / ElastiCache", "Session hot state, rate limit counters", "Sub-ms key-value"],
        ["Blob Store", "S3 / GCS", "Uploaded files, generated images, audio", "Signed URL access"],
        ["Message Queue", "SQS / Pub/Sub", "Async embedding jobs, notifications", "At-least-once delivery"],
        ["Time-Series", "InfluxDB / CloudWatch", "Latency, token counts, error rates", "Aggregation queries"],
    ]
    st2 = Table(storage_data, colWidths=[TW*0.18, TW*0.26, TW*0.32, TW*0.24])
    st2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HexColor("#37474f")),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, C_LIGHT_BLUE]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
        ('VALIGN',     (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(st2)
    story.append(Paragraph("Table 9 – Storage Systems", S['caption']))

    # 4.8 Observability
    story.append(Paragraph("4.8  Observability & Monitoring", S['h2']))
    obs_data = [
        ["Signal", "Tool", "Key Metrics"],
        ["Metrics", "Prometheus + Grafana", "Request rate, latency p50/p95/p99, error rate, token usage"],
        ["Traces", "OpenTelemetry + Jaeger/Tempo", "End-to-end request trace, tool call spans, LLM latency"],
        ["Logs", "Fluentd → CloudWatch / Loki", "Structured JSON logs, session_id as correlation key"],
        ["LLM Evals", "LangSmith / Braintrust", "Faithfulness, answer relevance, tool accuracy"],
        ["Alerts", "PagerDuty / OpsGenie", "Error rate > 1%, p99 > 10 s, cost anomaly"],
        ["User Feedback", "Thumbs up/down in UI", "CSAT score per session stored in DB"],
    ]
    ot = Table(obs_data, colWidths=[TW*0.18, TW*0.30, TW*0.52])
    ot.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HexColor("#4e342e")),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, HexColor("#efebe9")]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
        ('VALIGN',     (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(ot)
    story.append(Paragraph("Table 10 – Observability Stack", S['caption']))
    story.append(PageBreak())

    # ────────────────────────────────────────────────────────────────────────────
    # 5. DATA FLOW DIAGRAM
    # ────────────────────────────────────────────────────────────────────────────
    story.append(SectionBanner("5. Data Flow Diagram (Request Sequence)", TW, C_TEAL))
    story.append(Spacer(1, 8))
    story.append(DataFlowDiagram(TW, 200))
    story.append(Paragraph(
        "Figure 2 – Sequence diagram for a single user turn with one tool call. "
        "Dashed arrows represent responses/returns; solid arrows are requests. "
        "Steps 6–7 (red) show the tool round-trip; steps in blue show SSE streaming back to the user.",
        S['caption']
    ))
    story.append(Spacer(1, 8))

    story.append(Paragraph("<b>Agentic Loop Detail</b>", S['h3']))
    story.append(Paragraph(
        "The orchestrator implements the LLM agentic loop:", S['body']
    ))
    loop_code = (
        "WHILE True:\n"
        "    response = llm.messages.create(model, tools, messages, stream=True)\n"
        "    relay_tokens_to_client(response)\n"
        "    IF response.stop_reason == 'end_turn': BREAK\n"
        "    IF response.stop_reason == 'tool_use':\n"
        "        tool_results = []\n"
        "        FOR tool_call IN response.tool_use_blocks:\n"
        "            result = tool_layer.execute(tool_call.name, tool_call.input)\n"
        "            tool_results.append({id: tool_call.id, content: result})\n"
        "        messages.append({role:'assistant', content: response.content})\n"
        "        messages.append({role:'user', content: tool_results})\n"
        "    IF response.stop_reason == 'max_tokens': raise MaxTokensError\n"
        "persist_turn(messages[-1])\n"
        "embed_async(messages[-1])"
    )
    story.append(Paragraph(loop_code.replace("\n", "<br/>"), S['code']))
    story.append(PageBreak())

    # ────────────────────────────────────────────────────────────────────────────
    # 6. API DESIGN & SCHEMAS
    # ────────────────────────────────────────────────────────────────────────────
    story.append(SectionBanner("6. API Design & Data Schemas", TW, C_MID_BLUE))
    story.append(Spacer(1, 6))

    story.append(Paragraph("<b>REST Endpoints</b>", S['h3']))
    endpoints = [
        ["Method", "Path", "Description", "Auth"],
        ["POST", "/v1/chat", "Send a user message, receive SSE stream", "Bearer JWT"],
        ["GET",  "/v1/sessions", "List user's sessions", "Bearer JWT"],
        ["GET",  "/v1/sessions/{id}/messages", "Paginated message history", "Bearer JWT"],
        ["DELETE","/v1/sessions/{id}", "Archive session", "Bearer JWT"],
        ["POST", "/v1/sessions/{id}/feedback", "Submit thumbs up/down", "Bearer JWT"],
        ["GET",  "/v1/health", "Liveness probe", "None"],
        ["GET",  "/v1/metrics", "Prometheus scrape endpoint", "Internal only"],
    ]
    et = Table(endpoints, colWidths=[TW*0.10, TW*0.36, TW*0.40, TW*0.14])
    et.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_MID_BLUE),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, C_LIGHT_BLUE]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
    ]))
    story.append(et)
    story.append(Paragraph("Table 11 – Public API Endpoints", S['caption']))

    story.append(Paragraph("<b>POST /v1/chat – Request Body</b>", S['h3']))
    req_body = (
        '{\n'
        '  "session_id":  "uuid",          // optional; omit to create new session\n'
        '  "message":     "string",         // user utterance (required)\n'
        '  "persona":     "assistant",      // optional persona override\n'
        '  "locale":      "en-US",          // for localisation\n'
        '  "attachments": [                 // optional file_ids from Files API\n'
        '    {"type":"document","file_id":"file_..."}\n'
        '  ],\n'
        '  "stream":      true              // always recommended\n'
        '}'
    )
    story.append(Paragraph(req_body.replace("\n", "<br/>"), S['code']))

    story.append(Paragraph("<b>SSE Stream Events</b>", S['h3']))
    sse_data = [
        ["Event Name", "Payload", "Meaning"],
        ["session_created", '{"session_id": "uuid"}', "New session initialised"],
        ["content_start", '{"type": "text"}', "LLM text block starting"],
        ["content_delta", '{"text": "..."}', "Token chunk from LLM"],
        ["tool_start", '{"name": "web_search", "input": {...}}', "Tool call initiated"],
        ["tool_result", '{"tool_id": "...", "content": "..."}', "Tool result returned"],
        ["content_stop", '{}', "Content block finished"],
        ["message_stop", '{"usage": {"input_tokens": N, "output_tokens": M}}', "Full response done"],
        ["error", '{"code": "rate_limit", "message": "..."}', "Error occurred"],
    ]
    sset = Table(sse_data, colWidths=[TW*0.22, TW*0.44, TW*0.34])
    sset.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_TEAL),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 7.5),
        ('FONTNAME',   (0,1), (-1,-1), 'Courier'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, HexColor("#e0f2f1")]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING',(0,0),(-1,-1), 3),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
        ('VALIGN',     (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(sset)
    story.append(Paragraph("Table 12 – SSE Event Types", S['caption']))
    story.append(PageBreak())

    # ────────────────────────────────────────────────────────────────────────────
    # 7. NON-FUNCTIONAL REQUIREMENTS
    # ────────────────────────────────────────────────────────────────────────────
    story.append(SectionBanner("7. Non-Functional Requirements", TW, C_PURPLE))
    story.append(Spacer(1, 6))
    nfr = [
        ["NFR", "Requirement", "Implementation Strategy"],
        ["Performance", "p50 TTFB < 2 s; p99 < 8 s",
         "Streaming LLM, edge caching of system prompt, pre-warmed containers"],
        ["Scalability", "10 K concurrent sessions; 100 K/day messages",
         "Horizontal pod autoscaling (HPA), stateless orchestrator, read replicas"],
        ["Availability", "99.9 % SLA (≤ 43 min downtime/month)",
         "Multi-AZ deployment, health checks, circuit breakers, graceful degradation"],
        ["Reliability", "Exactly-once message persistence",
         "Idempotent writes with upsert-by-ID, outbox pattern for async events"],
        ["Security", "OWASP Top-10 compliance, SOC 2 Type II",
         "Input validation, parameterised queries, secrets in AWS Secrets Manager, pen-testing"],
        ["Privacy", "GDPR / CCPA compliant",
         "PII redaction before storage, data deletion endpoints, audit logs"],
        ["Cost", "< $0.05 per conversation (avg 10 turns)",
         "Prompt caching, Haiku for classification, batching for embeddings"],
        ["Fault Tolerance", "Tool failures do not crash the chat",
         "try/catch per tool call; LLM receives error message and adapts"],
        ["Maintainability", "CI/CD, feature flags, A/B model testing",
         "GitHub Actions, LaunchDarkly, LangSmith eval comparisons"],
    ]
    nfrt = Table(nfr, colWidths=[TW*0.18, TW*0.30, TW*0.52])
    nfrt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_PURPLE),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, HexColor("#f3e5f5")]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
        ('VALIGN',     (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(nfrt)
    story.append(Paragraph("Table 13 – Non-Functional Requirements", S['caption']))
    story.append(PageBreak())

    # ────────────────────────────────────────────────────────────────────────────
    # 8. SCALING & COST ESTIMATES
    # ────────────────────────────────────────────────────────────────────────────
    story.append(SectionBanner("8. Scaling & Cost Estimates", TW, C_GREEN))
    story.append(Spacer(1, 6))

    story.append(Paragraph("<b>Horizontal Scaling Model</b>", S['h3']))
    story.append(Paragraph(
        "All services are containerised (Docker) and orchestrated by <b>Kubernetes</b>. "
        "The orchestration service is the primary scaling unit. Each pod handles ~100 concurrent "
        "streaming SSE connections. A single 4-vCPU / 8 GB pod costs ~$0.05/hour on AWS EKS.", S['body']
    ))

    scale_data = [
        ["Tier", "Concurrent Sessions", "Orchestrator Pods", "LLM Spend/Month", "Infra Spend/Month"],
        ["Starter", "< 200", "2 pods (HA)", "~$500", "~$300"],
        ["Growth", "1 000", "5–10 pods", "~$2 500", "~$800"],
        ["Scale", "10 000", "30–50 pods + HPA", "~$25 000", "~$5 000"],
        ["Enterprise", "100 000+", "Auto-scale + multi-region", "~$200 000+", "~$40 000+"],
    ]
    sct = Table(scale_data, colWidths=[TW*0.15, TW*0.24, TW*0.25, TW*0.18, TW*0.18])
    sct.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_GREEN),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 8),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, HexColor("#e8f5e9")]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
    ]))
    story.append(sct)
    story.append(Paragraph(
        "Table 14 – Scaling Tiers (estimates based on Claude Opus 4.6 pricing, avg 1 000 tokens/turn, 10 turns/session)",
        S['caption']
    ))

    story.append(Paragraph("<b>Cost Optimisation Levers</b>", S['h3']))
    cost_opts = [
        "• <b>Prompt caching</b> – system prompt cached → ~90% cheaper for cached tokens (Anthropic ephemeral cache)",
        "• <b>Model routing</b> – use claude-haiku-4-5 ($1/$5/1M) for intent classification, FAQ lookups",
        "• <b>Batch embeddings</b> – embed async in SQS workers, not in request path",
        "• <b>Vector search</b> – reduce retrieved chunks to only what fits token budget (avoid over-retrieval)",
        "• <b>Conversation compaction</b> – Anthropic compact-2026-01-12 beta: server-side context summarisation",
        "• <b>Spot instances</b> – use spot/preemptible nodes for background workers (60–70% cost saving)",
    ]
    for c in cost_opts:
        story.append(Paragraph(c, S['bullet']))
    story.append(PageBreak())

    # ────────────────────────────────────────────────────────────────────────────
    # 9. REAL-WORLD EXAMPLES
    # ────────────────────────────────────────────────────────────────────────────
    story.append(SectionBanner("9. Real-World Examples & Analogues", TW, C_ACCENT))
    story.append(Spacer(1, 6))
    examples = [
        ["Product", "Company", "Key Architecture Choices", "Lesson Learned"],
        ["Claude.ai", "Anthropic", "Claude Opus 4.6, Projects (long-term memory), Artifacts (code renderer), "
         "MCP tool marketplace", "Adaptive thinking + streaming = best perceived quality"],
        ["ChatGPT Plus", "OpenAI", "GPT-4o, DALL-E integration, browsing via Bing, Code Interpreter "
         "(sandboxed Python), plugin ecosystem",
         "Plugins → GPTs: curated is safer than open marketplace"],
        ["GitHub Copilot", "GitHub / MS", "Codex/GPT-4o, LSP integration, context from open files + repo, "
         "real-time streaming completions",
         "Low-latency first: p50 < 500 ms; use Haiku-class model for completions"],
        ["Intercom Fin", "Intercom", "Fine-tuned LLM + RAG over help-centre docs, human escalation "
         "when confidence < threshold",
         "Escalation is a feature, not a failure; track escalation rate as KPI"],
        ["Cursor IDE", "Cursor", "Claude Sonnet/Opus for heavy reasoning, Haiku for quick edits, "
         "Apply/Reject diff UI, codebase-wide RAG",
         "Multi-model routing: match model to task complexity"],
        ["Glean", "Glean", "Enterprise search + chat over private data, connector ecosystem "
         "(Slack, GDrive, Jira, Salesforce), RBAC-aware retrieval",
         "Access control must be enforced at retrieval, not just at display"],
        ["Perplexity AI", "Perplexity", "LLM + real-time web search, inline citations, follow-up questions, "
         "Sonar model family",
         "Citations dramatically increase user trust; always show sources"],
    ]
    ext = Table(examples, colWidths=[TW*0.16, TW*0.14, TW*0.42, TW*0.28])
    ext.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_ACCENT),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 7.5),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, HexColor("#fff8e1")]),
        ('GRID',       (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 5),
        ('VALIGN',     (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(ext)
    story.append(Paragraph("Table 15 – Real-World Chat Agent Products", S['caption']))
    story.append(PageBreak())

    # ────────────────────────────────────────────────────────────────────────────
    # 10. REFERENCES
    # ────────────────────────────────────────────────────────────────────────────
    story.append(SectionBanner("10. References", TW, C_DEEP_BLUE))
    story.append(Spacer(1, 6))

    refs = [
        ("Documentation & Papers", [
            "[1] Anthropic. Claude Opus 4.6 Model Card & API Docs. platform.claude.com/docs",
            "[2] Anthropic. Claude Agent SDK Python – GitHub. github.com/anthropics/claude-agent-sdk-python",
            "[3] Anthropic. Tool Use & MCP Overview. platform.claude.com/docs/en/agents-and-tools",
            "[4] OpenAI. GPT-4 Technical Report (2023). arxiv.org/abs/2303.08774",
            "[5] Lewis et al. Retrieval-Augmented Generation (RAG). NeurIPS 2020. arxiv.org/abs/2005.11401",
            "[6] Yao et al. ReAct: Synergizing Reasoning and Acting in Language Models. ICLR 2023. arxiv.org/abs/2210.03629",
            "[7] Anthropic. Prompt Caching Guide. platform.claude.com/docs/en/build-with-claude/prompt-caching",
            "[8] Model Context Protocol (MCP) Specification. modelcontextprotocol.io",
        ]),
        ("Blog Posts & Articles", [
            "[9] Lilian Weng. LLM Powered Autonomous Agents. lilianweng.github.io/posts/2023-06-23-agent",
            "[10] Chip Huyen. Building LLM Applications for Production. huyenchip.com/2023/04/11/llm-engineering.html",
            "[11] Eugene Yan. Patterns for Building LLM-Based Systems & Products. eugeneyan.com/writing/llm-patterns",
            "[12] Anthropic. Introducing Claude's Tool Use. anthropic.com/news/tool-use-ga",
            "[13] Hamel Husain. Your AI Product Needs Evals. hamel.dev/blog/posts/evals",
            "[14] Simon Willison. LLMs as Operating Systems. simonwillison.net",
        ]),
        ("YouTube Videos", [
            "[15] Andrej Karpathy. Intro to Large Language Models. youtube.com/watch?v=zjkBMFhNj_g",
            "[16] AI Engineer Summit 2024 – Agents & Tool Use track. youtube.com/@aiDotEngineer",
            "[17] Anthropic. Building with Claude: Tool Use Deep Dive. (Anthropic YouTube channel)",
            "[18] LangChain. Building Production-Grade Agents. youtube.com/@LangChain",
            "[19] Full Stack Deep Learning. LLM Bootcamp 2023. youtube.com/playlist?list=PL1T8fO7ArWleyIqOy37OVXsP4hFXymdOZ",
        ]),
        ("Tools & Frameworks", [
            "[20] LangChain – framework for building LLM apps. langchain.com",
            "[21] LangSmith – LLM observability and evals. smith.langchain.com",
            "[22] Pinecone – managed vector database. pinecone.io",
            "[23] Weaviate – open-source vector DB. weaviate.io",
            "[24] Kong – API Gateway. konghq.com",
            "[25] OpenTelemetry – observability framework. opentelemetry.io",
            "[26] Braintrust – LLM evaluation platform. braintrustdata.com",
        ]),
    ]

    for (section, items) in refs:
        story.append(Paragraph(section, S['h3']))
        for item in items:
            story.append(Paragraph(f"• {item}", S['bullet']))
        story.append(Spacer(1, 4))

    # ── FOOTER NOTE ──────────────────────────────────────────────────────────────
    story.append(HRFlowable(width=TW, color=C_BORDER, thickness=0.5))
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        "Generated by Claude Opus 4.6 · Anthropic Agent SDK · March 2026  |  "
        "For questions: system-design@company.ai",
        ParagraphStyle('footer', fontSize=7, textColor=grey, alignment=TA_CENTER, fontName='Helvetica')
    ))

    doc.build(story)
    print(f"✅  PDF written to: {out_path}")


if __name__ == "__main__":
    out = os.path.join(os.path.dirname(__file__), "interactive_chat_agent_system_design.pdf")
    build_pdf(out)
