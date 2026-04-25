#!/usr/bin/env python3
"""
Distributed Key-Value Store System Design — PDF Generator
Uses ReportLab to produce a professional, diagram-rich design document.
"""

import math
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, cm, inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether, Image
)
from reportlab.graphics.shapes import (
    Drawing, Circle, Rect, Line, String, Polygon, Path,
    PolyLine, Ellipse, Group
)
from reportlab.graphics import renderPDF
from reportlab.graphics.charts.textlabels import Label
from reportlab.platypus.flowables import Flowable

# ─────────────────────────────────────────────
# COLOUR PALETTE
# ─────────────────────────────────────────────
C_NAVY      = colors.HexColor("#0D1B2A")
C_BLUE      = colors.HexColor("#1A73E8")
C_LIGHT_BLUE= colors.HexColor("#4A9EFF")
C_TEAL      = colors.HexColor("#00B4D8")
C_GREEN     = colors.HexColor("#06D6A0")
C_ORANGE    = colors.HexColor("#FF9F1C")
C_RED       = colors.HexColor("#EF476F")
C_PURPLE    = colors.HexColor("#7B2FBE")
C_YELLOW    = colors.HexColor("#FFD166")
C_GRAY      = colors.HexColor("#6B7280")
C_LIGHT_GRAY= colors.HexColor("#F3F4F6")
C_MID_GRAY  = colors.HexColor("#D1D5DB")
C_WHITE     = colors.white
C_BG        = colors.HexColor("#F8FAFC")

W, H = A4

# ─────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────
styles = getSampleStyleSheet()

def make_style(name, **kwargs):
    return ParagraphStyle(name, **kwargs)

S_TITLE = make_style("S_TITLE",
    fontSize=28, textColor=C_WHITE, alignment=TA_CENTER,
    fontName="Helvetica-Bold", spaceAfter=6, leading=34)

S_SUBTITLE = make_style("S_SUBTITLE",
    fontSize=14, textColor=C_LIGHT_BLUE, alignment=TA_CENTER,
    fontName="Helvetica", spaceAfter=4, leading=18)

S_H1 = make_style("S_H1",
    fontSize=18, textColor=C_NAVY, fontName="Helvetica-Bold",
    spaceBefore=14, spaceAfter=6, leading=22)

S_H2 = make_style("S_H2",
    fontSize=14, textColor=C_BLUE, fontName="Helvetica-Bold",
    spaceBefore=10, spaceAfter=4, leading=18)

S_H3 = make_style("S_H3",
    fontSize=11, textColor=C_NAVY, fontName="Helvetica-Bold",
    spaceBefore=8, spaceAfter=3, leading=15)

S_BODY = make_style("S_BODY",
    fontSize=9.5, textColor=C_NAVY, fontName="Helvetica",
    spaceBefore=3, spaceAfter=3, leading=14, alignment=TA_JUSTIFY)

S_BULLET = make_style("S_BULLET",
    fontSize=9.5, textColor=C_NAVY, fontName="Helvetica",
    spaceBefore=2, spaceAfter=2, leading=14,
    leftIndent=14, bulletIndent=4)

S_CODE = make_style("S_CODE",
    fontSize=8.5, textColor=colors.HexColor("#1E293B"),
    fontName="Courier", backColor=colors.HexColor("#F1F5F9"),
    spaceBefore=4, spaceAfter=4, leading=13,
    leftIndent=8, rightIndent=8)

S_CAPTION = make_style("S_CAPTION",
    fontSize=8, textColor=C_GRAY, fontName="Helvetica-Oblique",
    alignment=TA_CENTER, spaceBefore=2, spaceAfter=6)

S_CALLOUT = make_style("S_CALLOUT",
    fontSize=9, textColor=C_NAVY, fontName="Helvetica",
    spaceBefore=4, spaceAfter=4, leading=14,
    leftIndent=12, rightIndent=12)

S_QA_Q = make_style("S_QA_Q",
    fontSize=10, textColor=C_BLUE, fontName="Helvetica-Bold",
    spaceBefore=8, spaceAfter=2, leading=14)

S_QA_A = make_style("S_QA_A",
    fontSize=9.5, textColor=C_NAVY, fontName="Helvetica",
    spaceBefore=2, spaceAfter=4, leading=14,
    leftIndent=12, alignment=TA_JUSTIFY)

S_REF = make_style("S_REF",
    fontSize=9, textColor=C_BLUE, fontName="Helvetica",
    spaceBefore=2, spaceAfter=2, leading=13)

def b(t): return f"<b>{t}</b>"
def it(t): return f"<i>{t}</i>"
def code(t): return f'<font name="Courier" size="9" color="#1E293B">{t}</font>'
def col(t, c): return f'<font color="{c}">{t}</font>'

# ─────────────────────────────────────────────
# CUSTOM FLOWABLES
# ─────────────────────────────────────────────

class SectionHeader(Flowable):
    """Coloured full-width section banner."""
    def __init__(self, title, subtitle="", bg=C_NAVY, height=38):
        super().__init__()
        self.title    = title
        self.subtitle = subtitle
        self.bg       = bg
        self.height   = height
        self.width    = W - 2*cm

    def draw(self):
        w, h = self.width, self.height
        self.canv.setFillColor(self.bg)
        self.canv.roundRect(0, 0, w, h, 6, fill=1, stroke=0)
        # accent bar
        self.canv.setFillColor(C_TEAL)
        self.canv.rect(0, 0, 5, h, fill=1, stroke=0)
        # title
        self.canv.setFont("Helvetica-Bold", 14)
        self.canv.setFillColor(C_WHITE)
        self.canv.drawString(16, h - 22 if self.subtitle else h/2 - 7, self.title)
        if self.subtitle:
            self.canv.setFont("Helvetica", 9)
            self.canv.setFillColor(C_LIGHT_BLUE)
            self.canv.drawString(16, 8, self.subtitle)

    def wrap(self, aw, ah):
        return self.width, self.height + 6


class CalloutBox(Flowable):
    """Highlighted info/note/warning box."""
    ICONS = {"info": "ℹ", "warn": "⚠", "tip": "★", "key": "🔑"}
    BG    = {"info": colors.HexColor("#EFF6FF"),
             "warn": colors.HexColor("#FFFBEB"),
             "tip":  colors.HexColor("#F0FDF4"),
             "key":  colors.HexColor("#FAF5FF")}
    BORDER={"info": C_BLUE, "warn": C_ORANGE, "tip": C_GREEN, "key": C_PURPLE}

    def __init__(self, text, kind="info", avail_width=None):
        super().__init__()
        self.text  = text
        self.kind  = kind
        self._aw   = avail_width or (W - 2*cm)

    def draw(self):
        icon  = self.ICONS.get(self.kind, "")
        bg    = self.BG.get(self.kind, C_LIGHT_GRAY)
        bdr   = self.BORDER.get(self.kind, C_GRAY)
        w, h  = self._aw, self._h

        self.canv.setFillColor(bg)
        self.canv.roundRect(0, 0, w, h, 6, fill=1, stroke=0)
        self.canv.setStrokeColor(bdr)
        self.canv.setLineWidth(1.5)
        self.canv.roundRect(0, 0, w, h, 6, fill=0, stroke=1)
        # left accent
        self.canv.setFillColor(bdr)
        self.canv.rect(0, 0, 5, h, fill=1, stroke=0)

        self.canv.setFillColor(C_NAVY)
        self.canv.setFont("Helvetica", 9)
        # wrap text manually
        lines = self.text.split("\n")
        y = h - 14
        for line in lines:
            self.canv.drawString(14, y, line)
            y -= 13

    def wrap(self, aw, ah):
        self._aw = aw
        lines = self.text.split("\n")
        self._h = len(lines) * 13 + 12
        return aw, self._h


# ─────────────────────────────────────────────
# DIAGRAM BUILDERS
# ─────────────────────────────────────────────

def draw_consistent_hashing_ring():
    """Circular ring with 5 physical nodes and virtual node tokens."""
    W, H = 440, 340
    d = Drawing(W, H)
    cx, cy, r = W/2, H/2 - 10, 120

    # Ring background
    d.add(Circle(cx, cy, r+18, fillColor=colors.HexColor("#EFF6FF"),
                 strokeColor=C_LIGHT_BLUE, strokeWidth=1))
    d.add(Circle(cx, cy, r+18, fillColor=None,
                 strokeColor=C_LIGHT_BLUE, strokeWidth=36,
                 fillOpacity=0.0))

    # Ring circle
    d.add(Circle(cx, cy, r, fillColor=None,
                 strokeColor=C_BLUE, strokeWidth=2.5))

    # Physical node positions (5 nodes evenly spaced)
    nodes = ["N1", "N2", "N3", "N4", "N5"]
    node_colors = [C_BLUE, C_GREEN, C_ORANGE, C_RED, C_PURPLE]
    angles = [90, 162, 234, 306, 18]  # degrees from top, evenly spaced

    for i, (label, color, angle) in enumerate(zip(nodes, node_colors, angles)):
        rad = math.radians(angle)
        nx = cx + r * math.cos(rad)
        ny = cy + r * math.sin(rad)

        # Node circle
        d.add(Circle(nx, ny, 16, fillColor=color, strokeColor=C_WHITE,
                     strokeWidth=2))
        d.add(String(nx, ny - 4, label,
                     fontName="Helvetica-Bold", fontSize=9,
                     fillColor=C_WHITE, textAnchor="middle"))

        # Virtual node tokens (3 per physical node)
        for j, offset in enumerate([-22, 0, 22]):
            vrad = math.radians(angle + offset)
            vx = cx + (r + 24) * math.cos(vrad)
            vy = cy + (r + 24) * math.sin(vrad)
            d.add(Circle(vx, vy, 5, fillColor=color,
                         strokeColor=C_WHITE, strokeWidth=1, fillOpacity=0.6))

    # Centre label
    d.add(Circle(cx, cy, 28, fillColor=C_NAVY, strokeColor=C_TEAL,
                 strokeWidth=2))
    d.add(String(cx, cy + 3, "Hash", fontName="Helvetica-Bold", fontSize=8,
                 fillColor=C_WHITE, textAnchor="middle"))
    d.add(String(cx, cy - 8, "Ring", fontName="Helvetica-Bold", fontSize=8,
                 fillColor=C_WHITE, textAnchor="middle"))

    # Key mapping arrow example
    d.add(String(cx - 65, 18, "● Key K → hashed → lands between N3–N4 → owned by N4",
                 fontName="Helvetica", fontSize=8, fillColor=C_GRAY))

    # Legend
    lx, ly = 8, H - 20
    d.add(Circle(lx + 6, ly, 5, fillColor=C_BLUE))
    d.add(String(lx + 16, ly - 4, "Physical Node",
                 fontName="Helvetica", fontSize=8, fillColor=C_NAVY))
    lx2 = lx + 110
    d.add(Circle(lx2 + 6, ly, 5, fillColor=C_BLUE, fillOpacity=0.6))
    d.add(String(lx2 + 16, ly - 4, "Virtual Node (token)",
                 fontName="Helvetica", fontSize=8, fillColor=C_NAVY))

    # Title
    d.add(String(W/2, H - 12, "Consistent Hashing Ring  (5 nodes × 3 virtual tokens each = 15 tokens)",
                 fontName="Helvetica-Bold", fontSize=9,
                 fillColor=C_NAVY, textAnchor="middle"))
    return d


def draw_write_path():
    """Write path: Client → Coordinator → N replicas."""
    W, H = 480, 220
    d = Drawing(W, H)

    def box(x, y, w, h, label, sublabel="", fill=C_BLUE):
        d.add(Rect(x, y, w, h, rx=6, ry=6,
                   fillColor=fill, strokeColor=C_WHITE, strokeWidth=1.5))
        d.add(String(x + w/2, y + h/2 + (5 if sublabel else 0), label,
                     fontName="Helvetica-Bold", fontSize=9,
                     fillColor=C_WHITE, textAnchor="middle"))
        if sublabel:
            d.add(String(x + w/2, y + h/2 - 8, sublabel,
                         fontName="Helvetica", fontSize=7.5,
                         fillColor=colors.HexColor("#BEE3F8"), textAnchor="middle"))

    def arrow(x1, y1, x2, y2, label="", color=C_GRAY):
        d.add(Line(x1, y1, x2, y2, strokeColor=color, strokeWidth=1.5))
        # arrowhead
        dx, dy = x2 - x1, y2 - y1
        length = math.sqrt(dx*dx + dy*dy)
        ux, uy = dx/length, dy/length
        px, py = -uy, ux
        size = 7
        pts = [x2 - ux*size + px*size*0.4,
               y2 - uy*size + py*size*0.4,
               x2, y2,
               x2 - ux*size - px*size*0.4,
               y2 - uy*size - py*size*0.4]
        d.add(Polygon(pts, fillColor=color, strokeColor=color, strokeWidth=1))
        if label:
            mx, my = (x1+x2)/2, (y1+y2)/2
            d.add(String(mx + 4, my + 3, label,
                         fontName="Helvetica", fontSize=7.5,
                         fillColor=C_GRAY))

    # Client
    box(10, 90, 70, 40, "Client")

    # Coordinator
    box(130, 85, 90, 50, "Coordinator", "Node N1", C_TEAL)

    # 3 Replica nodes
    rep_y = [170, 90, 10]
    rep_labels = ["Replica A", "Replica B", "Replica C"]
    rep_sub    = ["(N=1)", "(N=2)", "(N=3)"]
    rep_colors = [C_GREEN, C_ORANGE, C_RED]

    for i, (ry, rl, rs, rc) in enumerate(zip(rep_y, rep_labels, rep_sub, rep_colors)):
        box(320, ry, 100, 40, rl, rs, rc)

    # Client → Coordinator
    arrow(80, 110, 130, 110, "PUT /kv/key", C_BLUE)

    # Coordinator → Replicas
    arrow(410, 175, 320, 175)
    arrow(220, 115, 320, 105, "", C_TEAL)
    arrow(220, 105, 320, 30, "", C_TEAL)
    arrow(220, 115, 320, 190, "", C_TEAL)

    # ACK arrows back
    for ry2 in [190, 105, 30]:
        d.add(Line(410, ry2, 440, ry2,
                   strokeColor=C_GREEN, strokeWidth=1, strokeDashArray=[3,2]))

    # W=2 ack note
    d.add(Rect(330, 130, 110, 22, rx=4, ry=4,
               fillColor=colors.HexColor("#F0FDF4"),
               strokeColor=C_GREEN, strokeWidth=1))
    d.add(String(385, 138, "Wait for W=2 ACKs ✓",
                 fontName="Helvetica-Bold", fontSize=8,
                 fillColor=C_GREEN, textAnchor="middle"))

    # Step labels
    steps = [
        (95, 122, "①"),
        (185, 122, "②"),
        (455, 145, "③"),
        (95, 75,  "④ Return success to client"),
    ]
    for sx, sy, sl in steps:
        d.add(String(sx, sy, sl, fontName="Helvetica-Bold", fontSize=9,
                     fillColor=C_NAVY, textAnchor="middle"))

    # Title
    d.add(String(W/2, H - 10, "Write Path: Client → Coordinator → N Replicas  (N=3, W=2)",
                 fontName="Helvetica-Bold", fontSize=9,
                 fillColor=C_NAVY, textAnchor="middle"))
    return d


def draw_vector_clock():
    """Vector clock / conflict resolution diagram."""
    W, H = 460, 240
    d = Drawing(W, H)

    # Time axis
    d.add(Line(80, 30, 80, H - 30, strokeColor=C_MID_GRAY, strokeWidth=1))

    nodes = ["Node A", "Node B", "Node C"]
    xs    = [120, 240, 360]
    node_cs = [C_BLUE, C_GREEN, C_PURPLE]

    for i, (label, x, c) in enumerate(zip(nodes, xs, node_cs)):
        # vertical timeline
        d.add(Line(x, 30, x, H-30, strokeColor=c, strokeWidth=1.5,
                   strokeDashArray=[4, 2]))
        d.add(String(x, H - 18, label, fontName="Helvetica-Bold", fontSize=9,
                     fillColor=c, textAnchor="middle"))

    def event(x, y, label, vc_str, fill=C_BLUE):
        d.add(Circle(x, y, 8, fillColor=fill, strokeColor=C_WHITE, strokeWidth=1.5))
        d.add(String(x + 14, y - 3, label,
                     fontName="Helvetica", fontSize=8, fillColor=C_NAVY))
        d.add(String(x + 14, y - 13,
                     vc_str, fontName="Courier", fontSize=7.5,
                     fillColor=C_GRAY))

    def msg_arrow(x1, y1, x2, y2, label=""):
        d.add(Line(x1, y1, x2, y2, strokeColor=C_ORANGE, strokeWidth=1.5))
        dx, dy = x2-x1, y2-y1
        length = math.sqrt(dx*dx+dy*dy) or 1
        ux, uy = dx/length, dy/length
        px, py = -uy, ux
        size = 6
        pts = [x2-ux*size+px*size*0.4, y2-uy*size+py*size*0.4,
               x2, y2,
               x2-ux*size-px*size*0.4, y2-uy*size-py*size*0.4]
        d.add(Polygon(pts, fillColor=C_ORANGE, strokeColor=C_ORANGE))
        if label:
            d.add(String((x1+x2)/2+4, (y1+y2)/2+4, label,
                         fontName="Helvetica", fontSize=7.5, fillColor=C_ORANGE))

    # Events
    event(120, 180, "write x=1", "[A:1,B:0,C:0]", C_BLUE)
    event(120, 130, "write x=5", "[A:2,B:0,C:0]", C_BLUE)
    event(240, 165, "write x=7", "[A:0,B:1,C:0]", C_GREEN)
    event(120,  80, "merge",     "[A:3,B:1,C:0]", C_TEAL)
    event(240, 115, "merge",     "[A:2,B:2,C:0]", C_TEAL)

    # Conflict box
    d.add(Rect(300, 95, 150, 60, rx=6, ry=6,
               fillColor=colors.HexColor("#FFF7ED"),
               strokeColor=C_RED, strokeWidth=1.5))
    d.add(String(375, 145, "⚡ Concurrent Writes",
                 fontName="Helvetica-Bold", fontSize=9,
                 fillColor=C_RED, textAnchor="middle"))
    d.add(String(375, 132, "A:2 ∥ B:1  (no causal order)",
                 fontName="Courier", fontSize=8,
                 fillColor=C_NAVY, textAnchor="middle"))
    d.add(String(375, 118, "→ App resolves (LWW or merge)",
                 fontName="Helvetica", fontSize=8,
                 fillColor=C_GRAY, textAnchor="middle"))

    # Message arrows
    msg_arrow(120, 130, 240, 165, "sync")
    msg_arrow(240, 165, 120, 80, "sync")

    # Title
    d.add(String(W/2, H-10, "Vector Clocks: Detecting Concurrent Writes",
                 fontName="Helvetica-Bold", fontSize=9,
                 fillColor=C_NAVY, textAnchor="middle"))
    # Time label
    d.add(String(52, H/2, "Time", fontName="Helvetica-Bold", fontSize=8,
                 fillColor=C_GRAY, textAnchor="middle"))

    return d


def draw_merkle_tree():
    """Merkle tree for anti-entropy reconciliation."""
    W, H = 460, 230
    d = Drawing(W, H)

    def node_box(x, y, label, sublabel="", fill=C_BLUE, w=80, h=30):
        d.add(Rect(x - w/2, y - h/2, w, h, rx=5, ry=5,
                   fillColor=fill, strokeColor=C_WHITE, strokeWidth=1.5))
        d.add(String(x, y + (5 if sublabel else 0), label,
                     fontName="Helvetica-Bold", fontSize=8.5,
                     fillColor=C_WHITE, textAnchor="middle"))
        if sublabel:
            d.add(String(x, y - 7, sublabel,
                         fontName="Courier", fontSize=7.5,
                         fillColor=colors.HexColor("#BEE3F8"), textAnchor="middle"))

    def edge(x1, y1, x2, y2, mismatch=False):
        c = C_RED if mismatch else C_MID_GRAY
        w = 2 if mismatch else 1.5
        d.add(Line(x1, y1, x2, y2, strokeColor=c, strokeWidth=w,
                   strokeDashArray=([3, 2] if mismatch else None)))

    # Root
    node_box(230, 195, "Root Hash", "h(L∥R)", C_NAVY, 100, 30)

    # Level 1
    node_box(130, 150, "Hash L", "h(LL∥LR)", C_BLUE, 90, 28)
    node_box(330, 150, "Hash R", "h(RL∥RR)", C_RED, 90, 28)  # mismatch

    edge(230, 180, 130, 164)
    edge(230, 180, 330, 164, mismatch=True)

    # Level 2 (leaves)
    leaves = [
        (80,  105, "Shard 0-3",   "hash ok",  C_GREEN),
        (180, 105, "Shard 4-7",   "hash ok",  C_GREEN),
        (280, 105, "Shard 8-11",  "MISMATCH", C_RED),
        (380, 105, "Shard 12-15", "hash ok",  C_GREEN),
    ]
    for (lx, ly, ll, ls, lc) in leaves:
        node_box(lx, ly, ll, ls, lc, 85, 30)

    edge(130, 136, 80,  120)
    edge(130, 136, 180, 120)
    edge(330, 136, 280, 120, mismatch=True)
    edge(330, 136, 380, 120)

    # Data rows under Shard 8-11
    for j, (key, status) in enumerate([("key_99", "STALE"), ("key_42", "MISSING")]):
        lx, ly2 = 280, 75 - j*20
        d.add(Rect(lx - 40, ly2 - 8, 80, 16, rx=3, ry=3,
                   fillColor=colors.HexColor("#FEF2F2"),
                   strokeColor=C_RED, strokeWidth=1))
        d.add(String(lx, ly2, f"{key} → {status}",
                     fontName="Courier", fontSize=7,
                     fillColor=C_RED, textAnchor="middle"))

    # Annotation
    d.add(String(W/2, 12, "Anti-Entropy: only divergent subtree (Shard 8-11) needs syncing  →  O(log n) bandwidth",
                 fontName="Helvetica", fontSize=8,
                 fillColor=C_GRAY, textAnchor="middle"))

    # Title
    d.add(String(W/2, H-12, "Merkle Tree for Anti-Entropy Synchronization",
                 fontName="Helvetica-Bold", fontSize=9,
                 fillColor=C_NAVY, textAnchor="middle"))
    return d


def draw_hinted_handoff():
    """Hinted handoff diagram: N3 down, N4 stores hint."""
    W, H = 460, 200
    d = Drawing(W, H)

    def node_box(x, y, label, sublabel="", fill=C_BLUE, w=85, h=36, dead=False):
        d.add(Rect(x - w/2, y - h/2, w, h, rx=5, ry=5,
                   fillColor=fill if not dead else C_LIGHT_GRAY,
                   strokeColor=C_RED if dead else C_WHITE, strokeWidth=2))
        d.add(String(x, y + (6 if sublabel else 0), label,
                     fontName="Helvetica-Bold", fontSize=9,
                     fillColor=C_GRAY if dead else C_WHITE, textAnchor="middle"))
        if sublabel:
            d.add(String(x, y - 8, sublabel,
                         fontName="Helvetica", fontSize=7.5,
                         fillColor=C_RED if dead else colors.HexColor("#BEE3F8"),
                         textAnchor="middle"))
        if dead:
            d.add(String(x + 38, y + 15, "✗",
                         fontName="Helvetica-Bold", fontSize=12,
                         fillColor=C_RED))

    def arrow(x1, y1, x2, y2, label="", dashed=False, color=C_TEAL):
        da = [4, 3] if dashed else None
        d.add(Line(x1, y1, x2, y2, strokeColor=color, strokeWidth=1.5,
                   strokeDashArray=da))
        dx, dy = x2-x1, y2-y1
        length = math.sqrt(dx*dx+dy*dy) or 1
        ux, uy = dx/length, dy/length
        px, py = -uy, ux
        s = 7
        pts = [x2-ux*s+px*s*0.4, y2-uy*s+py*s*0.4, x2, y2,
               x2-ux*s-px*s*0.4, y2-uy*s-py*s*0.4]
        d.add(Polygon(pts, fillColor=color, strokeColor=color))
        if label:
            d.add(String((x1+x2)/2, (y1+y2)/2 + 5, label,
                         fontName="Helvetica", fontSize=8, fillColor=color))

    # Client
    node_box(60,  100, "Client", "", C_NAVY, 60, 30)
    # Coordinator
    node_box(175, 100, "N1", "Coordinator", C_TEAL, 70, 36)
    # N2 (alive)
    node_box(295, 145, "N2", "Replica", C_GREEN, 70, 36)
    # N3 (dead)
    node_box(295,  55, "N3", "DOWN", C_RED, 70, 36, dead=True)
    # N4 (hint store)
    node_box(400, 100, "N4", "Hint Store", C_ORANGE, 80, 36)

    arrow(90, 100, 140, 100, "write", color=C_BLUE)
    arrow(210, 115, 260, 140, "replicate")
    arrow(210,  85, 260,  60, "", dashed=True, color=C_RED)  # fails
    d.add(String(245, 72, "FAIL", fontName="Helvetica-Bold", fontSize=8,
                 fillColor=C_RED))
    arrow(210, 100, 360, 100, "hint for N3", color=C_ORANGE)

    # Hint box on N4
    d.add(Rect(358, 60, 94, 26, rx=4, ry=4,
               fillColor=colors.HexColor("#FFF7ED"), strokeColor=C_ORANGE, strokeWidth=1))
    d.add(String(405, 70, "hint: {N3, key, val}",
                 fontName="Courier", fontSize=7.5, fillColor=C_NAVY, textAnchor="middle"))

    # Recovery arrow
    arrow(400, 82, 330, 65, "N3 recovers →\nN4 delivers hint", dashed=True, color=C_GREEN)

    d.add(String(W/2, 12, "Hinted Handoff: write availability preserved even when N3 is unreachable",
                 fontName="Helvetica", fontSize=8, fillColor=C_GRAY, textAnchor="middle"))
    d.add(String(W/2, H-12, "Hinted Handoff Flow",
                 fontName="Helvetica-Bold", fontSize=9, fillColor=C_NAVY, textAnchor="middle"))
    return d


def draw_gossip_protocol():
    """Gossip protocol membership propagation."""
    W, H = 420, 220
    d = Drawing(W, H)

    cx, cy = W/2, H/2 + 10
    positions = []
    n = 6
    for i in range(n):
        angle = math.radians(90 + i * 360/n)
        x = cx + 110 * math.cos(angle)
        y = cy + 90 * math.sin(angle)
        positions.append((x, y))

    colors_nodes = [C_BLUE, C_GREEN, C_TEAL, C_ORANGE, C_PURPLE, C_RED]
    labels = [f"N{i+1}" for i in range(n)]
    states = ["knows: {N2,N3}", "knows: {N1}", "knows: {N1,N2}", "knows: all", "knows: {N4}", "knows: {N5}"]

    # Gossip edges (partial, as in real gossip)
    gossip_pairs = [(0,1),(1,2),(2,3),(3,4),(4,5),(0,3),(1,4)]
    for (i, j) in gossip_pairs:
        x1, y1 = positions[i]
        x2, y2 = positions[j]
        d.add(Line(x1, y1, x2, y2,
                   strokeColor=C_LIGHT_BLUE, strokeWidth=1, strokeDashArray=[3,2]))

    for i, ((x, y), c, lbl) in enumerate(zip(positions, colors_nodes, labels)):
        d.add(Circle(x, y, 18, fillColor=c, strokeColor=C_WHITE, strokeWidth=2))
        d.add(String(x, y - 5, lbl,
                     fontName="Helvetica-Bold", fontSize=10,
                     fillColor=C_WHITE, textAnchor="middle"))

    # Highlight one gossip message
    x1, y1 = positions[0]
    x2, y2 = positions[1]
    d.add(Line(x1, y1, x2, y2, strokeColor=C_ORANGE, strokeWidth=2.5))
    dx, dy = x2-x1, y2-y1
    length = math.sqrt(dx*dx+dy*dy) or 1
    ux, uy = dx/length, dy/length
    px, py = -uy, ux
    s = 8
    pts = [x2-ux*s+px*s*0.4, y2-uy*s+py*s*0.4, x2, y2,
           x2-ux*s-px*s*0.4, y2-uy*s-py*s*0.4]
    d.add(Polygon(pts, fillColor=C_ORANGE, strokeColor=C_ORANGE))
    d.add(String((x1+x2)/2 + 5, (y1+y2)/2 + 8, "gossip msg",
                 fontName="Helvetica", fontSize=7.5, fillColor=C_ORANGE))

    d.add(String(W/2, 14, "Each node periodically selects k random peers and exchanges membership state",
                 fontName="Helvetica", fontSize=8, fillColor=C_GRAY, textAnchor="middle"))
    d.add(String(W/2, H-10, "Gossip Protocol — Membership & Failure Detection",
                 fontName="Helvetica-Bold", fontSize=9, fillColor=C_NAVY, textAnchor="middle"))
    return d


def draw_architecture_layers():
    """Layered architecture overview."""
    W, H = 460, 290
    d = Drawing(W, H)

    layers = [
        ("Client SDK / REST API",              C_NAVY,    (20, 250, 420, 28)),
        ("Request Routing & Load Balancer",    C_BLUE,    (20, 212, 420, 28)),
        ("Coordinator Layer  (any node)",      C_TEAL,    (20, 174, 420, 28)),
        ("Replication Manager  (N=3)",         C_GREEN,   (20, 136, 420, 28)),
        ("Consistency Engine  (Quorum W+R>N)", C_ORANGE,  (20,  98, 420, 28)),
        ("Storage Engine  (LSM-Tree / SSTables)",C_PURPLE,(20,  60, 420, 28)),
        ("Physical Disk / SSD",                C_GRAY,    (20,  22, 420, 28)),
    ]

    for (label, color, (x, y, w, h)) in layers:
        d.add(Rect(x, y, w, h, rx=4, ry=4,
                   fillColor=color, strokeColor=C_WHITE, strokeWidth=1.5))
        d.add(String(x + w/2, y + h/2 - 4, label,
                     fontName="Helvetica-Bold", fontSize=9,
                     fillColor=C_WHITE, textAnchor="middle"))

    # Arrows between layers
    for i in range(len(layers) - 1):
        _, _, (_, y_top, _, h_top) = layers[i]
        _, _, (_, y_bot, _, _)     = layers[i+1]
        mid_y_top = y_top
        mid_y_bot = y_bot + layers[i+1][2][3]
        d.add(Line(W/2, mid_y_top, W/2, mid_y_bot,
                   strokeColor=C_MID_GRAY, strokeWidth=1))

    # Side labels
    d.add(String(W - 8, 155, "Data Flow",
                 fontName="Helvetica-Bold", fontSize=8,
                 fillColor=C_GRAY, textAnchor="end"))

    d.add(String(W/2, H - 10, "Distributed KV Store — Layered Architecture",
                 fontName="Helvetica-Bold", fontSize=9,
                 fillColor=C_NAVY, textAnchor="middle"))
    return d


def draw_quorum_diagram():
    """N=3, W=2, R=2 quorum visualisation."""
    W, H = 420, 180
    d = Drawing(W, H)

    scenarios = [
        ("W=2, R=2, N=3\nW+R > N ✓", C_GREEN, 80, 90, True),
        ("W=1, R=1, N=3\nW+R ≤ N ✗", C_RED,  230, 90, False),
        ("W=3, R=1, N=3\nW+R > N ✓", C_ORANGE,380, 90, True),
    ]

    for (label, c, cx, cy, valid) in scenarios:
        # Outer ring = N=3
        for i in range(3):
            angle = math.radians(90 + i * 120)
            nx = cx + 42 * math.cos(angle)
            ny = cy + 35 * math.sin(angle)
            filled = (i < (2 if "W=2" in label else (3 if "W=3" in label else 1)))
            fill = c if filled else C_LIGHT_GRAY
            d.add(Circle(nx, ny, 14, fillColor=fill,
                         strokeColor=C_MID_GRAY, strokeWidth=1.5))
            d.add(String(nx, ny - 4, "N" + str(i+1),
                         fontName="Helvetica-Bold", fontSize=8,
                         fillColor=C_WHITE if filled else C_GRAY,
                         textAnchor="middle"))

        lines = label.split("\n")
        d.add(String(cx, cy - 60, lines[0],
                     fontName="Helvetica-Bold", fontSize=8.5,
                     fillColor=C_NAVY, textAnchor="middle"))
        d.add(String(cx, cy - 72, lines[1],
                     fontName="Helvetica-Bold", fontSize=9,
                     fillColor=c, textAnchor="middle"))

    d.add(String(W/2, H - 10,
                 "Quorum Configurations  —  shaded = participating nodes",
                 fontName="Helvetica-Bold", fontSize=9,
                 fillColor=C_NAVY, textAnchor="middle"))
    d.add(String(W/2, 10,
                 "Rule: W + R > N guarantees at least one node overlaps between read and write sets",
                 fontName="Helvetica", fontSize=7.5,
                 fillColor=C_GRAY, textAnchor="middle"))
    return d


# ─────────────────────────────────────────────
# TABLE HELPERS
# ─────────────────────────────────────────────

def make_table(headers, rows, col_widths=None, header_bg=C_NAVY):
    data = [[Paragraph(b(h), make_style("th", fontName="Helvetica-Bold",
                          fontSize=8.5, textColor=C_WHITE, leading=12))
             for h in headers]]
    for row in rows:
        data.append([Paragraph(str(c), make_style("td", fontName="Helvetica",
                               fontSize=8.5, textColor=C_NAVY, leading=12))
                     for c in row])

    style = TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  header_bg),
        ("TEXTCOLOR",     (0,0), (-1,0),  C_WHITE),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [C_WHITE, C_LIGHT_GRAY]),
        ("GRID",          (0,0), (-1,-1), 0.5, C_MID_GRAY),
        ("TOPPADDING",    (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING",   (0,0), (-1,-1), 7),
        ("RIGHTPADDING",  (0,0), (-1,-1), 7),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
    ])

    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(style)
    return tbl


# ─────────────────────────────────────────────
# COVER PAGE
# ─────────────────────────────────────────────

def cover_page(elements):
    class CoverBG(Flowable):
        def draw(self):
            c = self.canv
            pw = W - 2*cm
            c.setFillColor(C_NAVY)
            c.roundRect(0, 0, pw, 200, 10, fill=1, stroke=0)
            # gradient-like stripes
            for i, col in enumerate([C_BLUE, C_TEAL, C_GREEN]):
                c.setFillColor(col)
                c.setFillAlpha(0.15)
                c.rect(i * pw/3, 0, pw/3, 200, fill=1, stroke=0)
            c.setFillAlpha(1)
            # accent line
            c.setFillColor(C_TEAL)
            c.rect(0, 0, pw, 5, fill=1, stroke=0)

        def wrap(self, aw, ah):
            return aw, 200

    elements.append(Spacer(1, 1.5*cm))
    elements.append(CoverBG())

    class CoverText(Flowable):
        def draw(self):
            c = self.canv
            pw = W - 2*cm
            c.setFillColor(C_WHITE)
            c.setFont("Helvetica-Bold", 26)
            c.drawCentredString(pw/2, 155, "Distributed Key-Value Store")
            c.setFont("Helvetica-Bold", 20)
            c.setFillColor(C_TEAL)
            c.drawCentredString(pw/2, 125, "System Design Blueprint")
            c.setFont("Helvetica", 11)
            c.setFillColor(C_LIGHT_BLUE)
            c.drawCentredString(pw/2, 98,
                "A deep dive into Amazon Dynamo-style architecture")
            c.setFont("Helvetica", 9)
            c.setFillColor(colors.HexColor("#64748B"))
            c.drawCentredString(pw/2, 72,
                "Partitioning · Replication · Consistency · Conflict Resolution · Failure Handling")
            c.setFont("Helvetica", 9)
            c.setFillColor(C_MID_GRAY)
            c.drawCentredString(pw/2, 30, "System Design Series  ·  2025")

        def wrap(self, aw, ah):
            return aw, 200

    elements.append(CoverText())
    elements.append(Spacer(1, 0.8*cm))

    # Tag pills
    class TagRow(Flowable):
        TAGS = ["Dynamo", "Cassandra", "DynamoDB", "Redis", "etcd",
                "CAP Theorem", "Consistent Hashing", "CRDT", "Gossip", "LSM-Tree"]
        def draw(self):
            c  = self.canv
            pw = W - 2*cm
            x  = 10
            y  = 14
            tag_colors = [C_BLUE, C_GREEN, C_TEAL, C_ORANGE, C_PURPLE,
                          C_RED, C_BLUE, C_GREEN, C_TEAL, C_ORANGE]
            for tag, tc in zip(self.TAGS, tag_colors):
                tw = len(tag) * 6 + 14
                c.setFillColor(tc)
                c.setFillAlpha(0.15)
                c.roundRect(x, y - 8, tw, 18, 4, fill=1, stroke=0)
                c.setFillAlpha(1)
                c.setStrokeColor(tc)
                c.setLineWidth(1)
                c.roundRect(x, y - 8, tw, 18, 4, fill=0, stroke=1)
                c.setFont("Helvetica-Bold", 8)
                c.setFillColor(tc)
                c.drawString(x + 7, y - 1, tag)
                x += tw + 8
                if x > pw - 80:
                    x = 10
                    y -= 24

        def wrap(self, aw, ah):
            return aw, 50

    elements.append(TagRow())
    elements.append(Spacer(1, 0.5*cm))
    elements.append(HRFlowable(width="100%", color=C_MID_GRAY))


# ─────────────────────────────────────────────
# DOCUMENT SECTIONS
# ─────────────────────────────────────────────

def section_problem(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("1. The Problem",
        "Or: how to build the world's most reliable sticky-note board"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph(
        "Imagine you run the world's biggest vending machine empire. "
        "Millions of customers, every second, slide coins in and demand candy — "
        "<b>instantly, reliably, even when half your machines are on fire</b>. "
        "That is the humble key-value store at planet scale.", S_BODY))
    elements.append(Spacer(1, 4))
    elements.append(Paragraph(
        "A <b>distributed key-value store</b> is deceptively simple: "
        "you <b>put</b> a value under a key, you <b>get</b> it back. "
        "But when you must do this for <i>millions of requests per second</i>, "
        "across <i>hundreds of servers</i> spread across <i>multiple data centres</i>, "
        "while <i>tolerating hardware failures, network splits, and the occasional "
        "raccoon chewing through a fibre cable</i> — the design becomes a masterpiece "
        "of distributed systems engineering.", S_BODY))
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("The Core Challenge", S_H2))
    challenge_rows = [
        ("Scale",        "Millions of QPS, petabytes of data, thousands of nodes"),
        ("Availability", "Always writable — even during partial failures"),
        ("Consistency",  "Reads should reflect recent writes (tunable)"),
        ("Durability",   "Survive node crashes, disk failures, and datacenter outages"),
        ("Operability",  "Self-healing, auto-rebalancing, zero-downtime deploys"),
    ]
    t = make_table(["Requirement", "What It Means"], challenge_rows,
                   col_widths=[100, 340])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(CalloutBox(
        "The Foundational Paper: Amazon's Dynamo (SOSP 2007)\n"
        "Dynamo pioneered: consistent hashing, vector clocks, sloppy quorum,\n"
        "hinted handoff, and gossip-based membership. It inspired Cassandra,\n"
        "DynamoDB, Riak, and Voldemort.", kind="key"))


def section_cap(elements):
    elements.append(Spacer(1, 8))
    elements.append(SectionHeader("2. Technical Landscape",
        "CAP Theorem, consistency models, and the physics of distribution"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph("The CAP Theorem — Pick Two (But Mostly CP or AP)", S_H2))
    elements.append(Paragraph(
        "CAP states that in the presence of a <b>network partition</b>, "
        "you must choose between <b>Consistency</b> and <b>Availability</b>. "
        "A Dynamo-style store chooses <b>AP</b>: always accept writes, "
        "resolve conflicts lazily.", S_BODY))

    cap_rows = [
        ("CP (e.g. ZooKeeper, etcd)",  "Consistent", "Partition-tolerant", "May refuse writes during partition"),
        ("AP (e.g. Cassandra, Dynamo)", "Available", "Partition-tolerant", "May serve stale reads; eventual consistency"),
        ("CA (single node DB)",         "Consistent", "Available",          "Cannot survive network partitions"),
    ]
    t = make_table(["System Type", "Consistent", "Available", "Trade-off"], cap_rows,
                   col_widths=[120, 70, 90, 160])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Consistency Spectrum", S_H2))
    consistency_rows = [
        ("Strong Consistency",    "Every read sees the most recent write",              "High latency, low availability",     "Spanner, CockroachDB"),
        ("Linearisability",       "Operations appear instantaneous & globally ordered", "Very high latency",                  "etcd, ZooKeeper"),
        ("Sequential Consistency","All nodes see same order, not real-time",            "Moderate latency",                   "Cassandra w/ QUORUM"),
        ("Eventual Consistency",  "Given no new updates, all replicas converge",        "Lowest latency, highest availability","Cassandra ONE, Dynamo"),
        ("Read-your-writes",      "Client always reads its own writes",                 "Sticky sessions required",           "DynamoDB sessions"),
    ]
    t = make_table(["Model", "Guarantee", "Trade-off", "Example"],
                   consistency_rows,
                   col_widths=[110, 135, 115, 90])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("The PACELC Extension", S_H2))
    elements.append(Paragraph(
        "PACELC adds: even when there is <i>no</i> partition, "
        "you must still trade <b>Latency</b> against <b>Consistency</b>. "
        "Dynamo-style stores optimise for low latency by writing locally first "
        "and reconciling later.", S_BODY))


def section_requirements(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("3. Requirements & API",
        "What we're building and how clients talk to it"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph("Functional Requirements", S_H2))
    func_rows = [
        ("put(key, value)", "Store or update a value. Returns a context (vector clock)."),
        ("get(key)",        "Retrieve a value (or conflicting versions) for reconciliation."),
        ("delete(key)",     "Tombstone the key (logical delete; GC'd later)."),
        ("Tunable reads",   "Caller specifies consistency level: ONE / QUORUM / ALL."),
        ("Tunable writes",  "Caller specifies write durability: ONE / QUORUM / ALL."),
    ]
    t = make_table(["Operation", "Description"], func_rows,
                   col_widths=[120, 320])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Non-Functional Requirements", S_H2))
    nfr_rows = [
        ("Throughput",    "> 1M QPS aggregate",          "Horizontal sharding + replication"),
        ("Latency",       "p99 < 10 ms (read/write)",    "Async replication, local writes"),
        ("Availability",  "99.999% (five nines)",        "Sloppy quorum, hinted handoff"),
        ("Durability",    "No data loss on node failure", "N=3 replicas, WAL on each"),
        ("Scalability",   "Linear scale-out",            "Consistent hashing ring"),
        ("Key size",      "≤ 256 bytes",                 "Fixed-size hash regardless"),
        ("Value size",    "≤ 1 MB",                      "Chunked storage for large vals"),
    ]
    t = make_table(["Property", "Target", "Mechanism"],
                   nfr_rows, col_widths=[90, 130, 220])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("REST API Design", S_H2))
    api_code = [
        "PUT  /kv/{key}",
        "     Body:    { \"value\": <bytes>, \"context\": <vector_clock_token> }",
        "     Returns: 200 OK | 503 Unavailable",
        "",
        "GET  /kv/{key}?consistency=quorum",
        "     Returns: { \"value\": <bytes>, \"context\": <vclock> }",
        "              OR { \"values\": [v1,v2], \"context\": <vclock> }  // conflict",
        "",
        "DELETE /kv/{key}",
        "     Returns: 200 OK (tombstone written)",
        "",
        "GET  /kv/{key}/metadata",
        "     Returns: { \"replicas\": [n1,n2,n3], \"version\": <vclock>, \"ttl\": <ms> }",
    ]
    for line in api_code:
        elements.append(Paragraph(line if line else " ", S_CODE))

    elements.append(Spacer(1, 6))
    elements.append(Paragraph("Data Model", S_H2))
    dm_rows = [
        ("key",          "bytes[256]",  "Hash-distributed across ring"),
        ("value",        "bytes[1MB]",  "Opaque blob; app-defined schema"),
        ("vector_clock", "map[node→int]","Causality tracker; grows with writes"),
        ("timestamp",    "int64 (ms)",  "Wall-clock for LWW fallback"),
        ("checksum",     "uint32 CRC32","Detect bit-rot on read"),
        ("ttl",          "int32 (s)",   "Optional expiry; 0 = forever"),
        ("is_deleted",   "bool",        "Tombstone flag"),
    ]
    t = make_table(["Field", "Type", "Purpose"], dm_rows,
                   col_widths=[90, 100, 250])
    elements.append(t)


def section_architecture(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("4. High-Level Architecture",
        "The 10,000-foot view before we dive into the engine room"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph(
        "Every node in a Dynamo-style cluster is <b>equal</b> (no master). "
        "Any node can serve as <b>coordinator</b> for any request. "
        "Data is spread across nodes via a <b>consistent hash ring</b>, "
        "and each key is replicated to <b>N successor nodes</b> on the ring.", S_BODY))
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Layered Architecture", S_H2))
    d = draw_architecture_layers()
    elements.append(d)
    elements.append(Paragraph("Figure 1 — Logical layers of a distributed KV node", S_CAPTION))
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Consistent Hashing Ring", S_H2))
    d2 = draw_consistent_hashing_ring()
    elements.append(d2)
    elements.append(Paragraph(
        "Figure 2 — 5-node ring with virtual nodes. Each physical node owns multiple "
        "token ranges, balancing load even with heterogeneous hardware.", S_CAPTION))

    elements.append(Spacer(1, 6))
    elements.append(CalloutBox(
        "Why virtual nodes?\n"
        "Without them, adding a new physical node would only relieve one neighbour.\n"
        "With 150+ virtual nodes per physical node, load is redistributed across\n"
        "the entire ring proportionally — like shuffling a deck of cards.", kind="tip"))


def section_partitioning(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("5. Partitioning — Slicing the Pie",
        "Consistent hashing, virtual nodes, and token management"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph("Consistent Hashing Deep Dive", S_H2))
    elements.append(Paragraph(
        "The hash space is a ring from <b>0</b> to <b>2³²−1</b>. "
        "Each key is hashed (MD5 or MurmurHash3) to a position on this ring. "
        "The key is <i>owned</i> by the first node encountered clockwise from "
        "that position — the <b>coordinator node</b>.", S_BODY))
    elements.append(Spacer(1, 4))

    part_rows = [
        ("Hash function",    "MurmurHash3 (128-bit)",     "Fast, low collision, uniform distribution"),
        ("Ring size",        "2³² positions (4 billion)", "Ample space for fine-grained tokens"),
        ("Virtual nodes",    "150–256 per physical node", "Balances load; isolates hot spots"),
        ("Token assignment", "Random or evenly spaced",   "Evenly spaced = more predictable load"),
        ("Rebalancing",      "Incremental (streaming)",   "New node streams data from successor"),
        ("Metadata store",   "Gossip-propagated ring map","All nodes hold full ring state"),
    ]
    t = make_table(["Parameter", "Value / Range", "Rationale"],
                   part_rows, col_widths=[110, 130, 200])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Node Addition Algorithm", S_H2))
    steps = [
        "1. New node N_new assigned k random token positions on the ring.",
        "2. For each token, N_new becomes responsible for keys between "
        "   its predecessor token and itself.",
        "3. Existing nodes stream their data ranges to N_new in parallel.",
        "4. N_new announces itself via gossip; ring map updated cluster-wide.",
        "5. Once N_new acknowledges ownership, predecessors delete transferred ranges.",
    ]
    for step in steps:
        elements.append(Paragraph(step, S_BULLET))
    elements.append(Spacer(1, 4))

    elements.append(CalloutBox(
        "Hot Key Problem: if one key receives disproportionate traffic (celebrity post,\n"
        "trending item), a single coordinator gets hammered. Mitigations:\n"
        "  • Key sharding: append random suffix, fan-out reads\n"
        "  • Caching layer (Redis) in front of the KV store\n"
        "  • Adaptive load balancing via token reassignment", kind="warn"))


def section_replication(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("6. Replication — Never Lose Your Candy",
        "N replicas, preference lists, and the write path end-to-end"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph(
        "Each key is replicated across <b>N nodes</b> (typically N=3). "
        "The <b>preference list</b> for a key is the ordered list of nodes "
        "responsible for storing it — the coordinator plus its N−1 successors "
        "on the ring.", S_BODY))
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Write Path (Step by Step)", S_H2))
    d = draw_write_path()
    elements.append(d)
    elements.append(Paragraph("Figure 3 — Write path with N=3, W=2", S_CAPTION))
    elements.append(Spacer(1, 4))

    write_steps = [
        ("①  Client sends PUT",         "Client hashes key → finds coordinator from local ring map"),
        ("②  Coordinator receives write", "Writes to local storage + WAL"),
        ("③  Forwards to N−1 replicas",  "Parallel async replication to successor nodes"),
        ("④  Wait for W ACKs",           "W=2 → success after 2/3 nodes confirm (including itself)"),
        ("⑤  Return to client",          "Success response with vector clock context"),
        ("⑥  Background replication",    "Remaining replica eventually receives update"),
    ]
    t = make_table(["Step", "Description"], write_steps,
                   col_widths=[110, 330])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Read Path", S_H2))
    read_steps = [
        ("①  Client sends GET",          "Client contacts coordinator for the key"),
        ("②  Coordinator queries R nodes","Sends read to R nodes in preference list in parallel"),
        ("③  Collect R responses",        "Wait for R=2 responses (fastest wins)"),
        ("④  Conflict check",             "Compare vector clocks; if divergent → return all versions"),
        ("⑤  Read repair",                "Stale replicas updated asynchronously in background"),
        ("⑥  Return to client",           "Latest value (by vector clock) or conflict set"),
    ]
    t = make_table(["Step", "Description"], read_steps,
                   col_widths=[110, 330])
    elements.append(t)


def section_consistency(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("7. Consistency — Tuning the Dial",
        "Quorum, vector clocks, and the art of eventually agreeing"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph("Quorum Consistency (W + R > N)", S_H2))
    elements.append(Paragraph(
        "The magic inequality <b>W + R > N</b> guarantees that every read "
        "intersects with every write — at least one node in the read set must "
        "have the latest write. This is the heart of tunable consistency.", S_BODY))
    elements.append(Spacer(1, 4))

    d = draw_quorum_diagram()
    elements.append(d)
    elements.append(Paragraph("Figure 4 — Quorum configurations and their trade-offs", S_CAPTION))
    elements.append(Spacer(1, 4))

    quorum_rows = [
        ("W=1, R=1", "Lowest",   "Highest", "Fastest",    "Shopping cart (loss ok)"),
        ("W=2, R=2", "Medium",   "High",    "Medium",     "User profile, inventory"),
        ("W=3, R=1", "High",     "Medium",  "Write slow", "Financial audit log"),
        ("W=1, R=3", "High",     "Medium",  "Read slow",  "Config/feature flags"),
        ("W=3, R=3", "Strongest","Lowest",  "Slowest",    "Critical settings"),
    ]
    t = make_table(
        ["Config (N=3)", "Consistency", "Availability", "Latency", "Use Case"],
        quorum_rows, col_widths=[75, 80, 80, 80, 125])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Vector Clocks & Conflict Resolution", S_H2))
    elements.append(Paragraph(
        "A <b>vector clock</b> is a map of <code>{node_id → counter}</code> "
        "attached to every value. It tracks causality: if clock A "
        "dominates clock B (every counter ≥), A happened after B. "
        "If neither dominates → <b>concurrent writes</b> → conflict.", S_BODY))
    elements.append(Spacer(1, 4))

    d2 = draw_vector_clock()
    elements.append(d2)
    elements.append(Paragraph("Figure 5 — Vector clock divergence and conflict detection", S_CAPTION))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph("Conflict Resolution Strategies", S_H3))
    conflict_rows = [
        ("Last-Write-Wins (LWW)",    "Highest wall-clock timestamp wins",        "Simple; loses data on concurrent writes", "Cassandra default"),
        ("Vector clock + app merge", "Return all versions; app picks winner",     "No data loss; app complexity",            "Dynamo, Riak"),
        ("CRDT",                     "Mathematically mergeable data types",       "No conflict; only supports CRDT ops",     "Redis, Riak sets"),
        ("Multi-version (MVCC)",     "Keep all versions, reader picks",           "Storage cost; read complexity",           "Spanner, Postgres"),
    ]
    t = make_table(["Strategy", "How", "Trade-off", "Example"],
                   conflict_rows, col_widths=[110, 120, 120, 90])
    elements.append(t)


def section_failure(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("8. Failure Handling — When Things Go Wrong",
        "Hinted handoff, Merkle trees, and gossip — the holy trinity of resilience"))
    elements.append(Spacer(1, 4))

    # ── Hinted Handoff ──
    elements.append(Paragraph("8.1  Hinted Handoff", S_H2))
    elements.append(Paragraph(
        "When a replica node is temporarily unreachable, the coordinator "
        "does <i>not</i> fail the write. Instead, it sends the write to the "
        "next available node in the preference list with a <b>hint</b>: "
        '"deliver this to Node X once it recovers." '
        "This preserves <b>write availability</b> at the cost of potentially "
        "serving stale reads until recovery.", S_BODY))
    elements.append(Spacer(1, 4))
    d = draw_hinted_handoff()
    elements.append(d)
    elements.append(Paragraph("Figure 6 — Hinted handoff: N3 is down; N4 stores hint and delivers on recovery", S_CAPTION))
    elements.append(Spacer(1, 4))

    elements.append(CalloutBox(
        "Sloppy Quorum vs. Strict Quorum\n"
        "Strict: only the true preference-list nodes count toward W/R.\n"
        "Sloppy: any available node counts (Dynamo default).\n"
        "Sloppy = higher availability; Strict = stronger consistency guarantee.", kind="info"))
    elements.append(Spacer(1, 6))

    # ── Merkle Trees ──
    elements.append(Paragraph("8.2  Merkle Trees for Anti-Entropy", S_H2))
    elements.append(Paragraph(
        "After a node recovers, how do we efficiently identify which keys "
        "are stale without comparing every single one? "
        "<b>Merkle trees</b> to the rescue: each node maintains a hash tree "
        "over its key ranges. Comparing root hashes identifies divergent "
        "subtrees in <b>O(log n)</b> operations, narrowing the repair "
        "to only the affected key ranges.", S_BODY))
    elements.append(Spacer(1, 4))
    d2 = draw_merkle_tree()
    elements.append(d2)
    elements.append(Paragraph("Figure 7 — Merkle tree: only the red subtree needs repair", S_CAPTION))
    elements.append(Spacer(1, 4))

    merkle_rows = [
        ("Full key scan",  "O(n)",      "Compare every key/hash",    "Tiny datasets only"),
        ("Merkle tree",    "O(log n)",  "Compare hashes top-down",   "Dynamo, Cassandra"),
        ("Bloom filter",   "O(1) space","Probabilistic set membership","Presence checks, HBase"),
    ]
    t = make_table(["Method", "Complexity", "How", "Used In"],
                   merkle_rows, col_widths=[90, 80, 170, 100])
    elements.append(t)
    elements.append(Spacer(1, 6))

    # ── Gossip ──
    elements.append(Paragraph("8.3  Gossip Protocol for Membership", S_H2))
    elements.append(Paragraph(
        "There is no central registry. Nodes learn about each other and "
        "detect failures via a <b>gossip protocol</b>: every T seconds, "
        "each node picks k random peers and exchanges its view of cluster "
        "membership (which nodes are alive, their token assignments). "
        "Failures propagate in <b>O(log n)</b> rounds.", S_BODY))
    elements.append(Spacer(1, 4))
    d3 = draw_gossip_protocol()
    elements.append(d3)
    elements.append(Paragraph("Figure 8 — Gossip: information floods the cluster in O(log n) rounds", S_CAPTION))
    elements.append(Spacer(1, 4))

    gossip_rows = [
        ("Heartbeat interval",     "1–2 s",         "Balance detection speed vs. chatter"),
        ("Fan-out (k)",            "3 random peers", "Higher k = faster propagation, more bandwidth"),
        ("Failure threshold",      "Phi accrual (φ=8)", "Adaptive; adjusts to network jitter"),
        ("Ring state convergence", "O(log n) rounds","Logarithmic in cluster size"),
        ("State per message",      "< 1 KB",         "Only deltas / compressed ring diffs"),
    ]
    t = make_table(["Parameter", "Value", "Notes"],
                   gossip_rows, col_widths=[130, 120, 190])
    elements.append(t)


def section_storage(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("9. Storage Engine",
        "LSM-Trees, SSTables, and the write-optimised path to disk"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph(
        "Most distributed KV stores use a <b>Log-Structured Merge (LSM) tree</b> "
        "as the local storage engine. It converts random writes into sequential "
        "I/O — perfect for SSDs and HDDs alike.", S_BODY))
    elements.append(Spacer(1, 4))

    lsm_rows = [
        ("MemTable",      "In-memory sorted map",         "Recent writes; O(log n) lookup"),
        ("WAL",           "Append-only write-ahead log",  "Crash recovery; written before MemTable"),
        ("SSTable",       "Sorted, immutable disk files", "Flushed from MemTable"),
        ("Bloom Filter",  "Per-SSTable probabilistic set","Avoid disk reads for missing keys"),
        ("Block Cache",   "OS page cache / LRU cache",    "Hot-data acceleration"),
        ("Compaction",    "Merge & GC old SSTables",      "Reclaim space, merge tombstones"),
    ]
    t = make_table(["Component", "What It Is", "Purpose"],
                   lsm_rows, col_widths=[100, 170, 170])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("LSM Write & Read Path", S_H3))
    lsm_write = [
        "WRITE:  client → WAL (sync) → MemTable (async) → return OK",
        "        [background] MemTable full → flush to SSTable",
        "        [background] compaction merges SSTables (Leveled or Tiered)",
        "",
        "READ:   MemTable → L0 SSTables (newest first) → L1 → L2 ...",
        "        Bloom filter at each level to skip levels with no match",
        "        Block cache shortcut for hot keys",
    ]
    for line in lsm_write:
        elements.append(Paragraph(line if line else " ", S_CODE))
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Storage Configuration Reference", S_H2))
    cfg_rows = [
        ("MemTable size",          "64 MB",        "Larger = fewer flushes, more RAM"),
        ("SSTable block size",     "4 KB (SSD) / 64 KB (HDD)", "Match to I/O block size"),
        ("Bloom filter FP rate",   "1%",           "1% false positive rate typical"),
        ("L0 → L1 compaction",     "4 SSTables",   "Triggers compaction to prevent read amplification"),
        ("Level multiplier",       "10×",          "Each level 10× size of previous"),
        ("Compression",            "Snappy / Zstd","Block-level; 3–5× compression typical"),
        ("Compaction strategy",    "Leveled",       "Better read perf; Tiered = better write perf"),
        ("WAL sync policy",        "fsync per write", "Full durability; async for higher throughput"),
    ]
    t = make_table(["Parameter", "Typical Value", "Notes"],
                   cfg_rows, col_widths=[130, 130, 180])
    elements.append(t)

    elements.append(Spacer(1, 6))
    elements.append(CalloutBox(
        "Real engines: RocksDB (Facebook/Meta, used in TiKV, CockroachDB, Kafka),\n"
        "LevelDB (Google), WiredTiger (MongoDB), Bitcask (Riak — hash index + log).\n"
        "RocksDB is the de-facto standard for embedded KV storage today.", kind="key"))


def section_real_world(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("10. Real-World Systems",
        "How the giants built their key-value stores"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph(
        "Theory is fun, but production systems reveal the real trade-offs. "
        "Here is how the industry's most battle-tested KV stores compare.", S_BODY))
    elements.append(Spacer(1, 6))

    real_rows = [
        ("Amazon DynamoDB",    "Managed, serverless",  "Paxos (strong) / eventual", "Partition + sort key","Millions/s at Amazon", "E-commerce, sessions"),
        ("Apache Cassandra",   "Masterless ring (Dynamo)", "Tunable (ONE–ALL)",     "Composite partition key","Discord: 1T messages","Messaging, IoT, logs"),
        ("Redis",              "Single/Cluster",        "Strong (single) / AP (cluster)","Hash slots","1M+ QPS single node","Cache, sessions, pub/sub"),
        ("etcd",               "Raft-based leader",     "Strong (linearisable)",   "Flat key (+ watch)","Kubernetes control plane","Config, service discovery"),
        ("Riak KV",            "Dynamo-style ring",     "Tunable quorum",           "Bucket + key","Used by Basho, GitHub","Healthcare, user data"),
        ("TiKV",               "Raft groups (regions)", "Strong via Raft",          "Raw bytes",   "TiDB storage layer","HTAP, financial DBs"),
    ]
    t = make_table(
        ["System", "Architecture", "Consistency", "Key Model", "Scale Claim", "Use Cases"],
        real_rows, col_widths=[72, 82, 80, 70, 72, 68])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Discord's Cassandra Journey — A Case Study", S_H2))
    elements.append(Paragraph(
        "Discord stored <b>1 trillion messages</b> in Cassandra by 2023, then "
        "migrated hot data to <b>ScyllaDB</b> (C++ rewrite of Cassandra) "
        "to reduce tail latency. Key lessons: partition keys matter enormously "
        "(they routed by user ID → hot users = hot partitions), "
        "time-bucketed keys solved the hot partition problem, "
        "and compaction strategy tuning cut p99 latency by 3×.", S_BODY))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph("Amazon's Dynamo (2007) — The Blueprint", S_H2))
    elements.append(Paragraph(
        "Dynamo was built for Amazon's shopping cart service — a system that "
        "<b>must never lose a cart item</b>. The original paper reported "
        "sub-10ms p99.9 read latencies at millions of requests/second. "
        "Key insight: 'always writable' is a business requirement, not just "
        "an engineering preference. Dynamo inspired an entire generation of "
        "NoSQL systems.", S_BODY))


def section_performance(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("11. Performance & Capacity Planning",
        "Sizing your cluster, benchmarks, and tuning knobs"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph("Throughput Sizing Model", S_H2))
    sizing_rows = [
        ("Single SSD node write",   "~100K ops/s (4KB values, async)"),
        ("Single SSD node read",    "~200K ops/s (cached)"),
        ("Replication overhead",    "N× write bandwidth (N=3 → 3× write IOPS)"),
        ("Gossip bandwidth",        "~50–200 KB/s per node (scales with cluster size)"),
        ("Merkle sync bandwidth",   "Proportional to divergence; O(log n) comparison"),
        ("Compaction I/O",          "10–40% of write bandwidth in steady state"),
    ]
    t = make_table(["Metric", "Rule of Thumb"],
                   sizing_rows, col_widths=[190, 250])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Cluster Size Calculator", S_H2))
    calc_rows = [
        ("Target write QPS",     "1,000,000",  "Given"),
        ("Writes per node",      "50,000/s",   "Assuming SSD nodes"),
        ("Raw nodes needed",     "20",          "1M / 50K"),
        ("With N=3 replication", "60",          "20 × 3 (each write touches 3 nodes)"),
        ("With 20% headroom",    "72",          "60 × 1.2 (spare capacity)"),
        ("Virtual nodes/node",   "256",         "For load balance"),
        ("Total ring tokens",    "18,432",      "72 × 256"),
    ]
    t = make_table(["Parameter", "Value", "Notes"],
                   calc_rows, col_widths=[160, 100, 180])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Latency Budget (p99 target: 10 ms)", S_H2))
    lat_rows = [
        ("Network RTT (same AZ)",   "0.5 ms",  "Intra-AZ typical"),
        ("Serialisation overhead",  "0.1 ms",  "Protocol Buffers encoding"),
        ("Coordinator processing",  "0.5 ms",  "Hashing, routing"),
        ("MemTable write",          "0.2 ms",  "In-memory sorted insert"),
        ("WAL fsync",               "1–5 ms",  "Dominant cost; use battery-backed cache"),
        ("Waiting for W=2 ACKs",    "1–3 ms",  "Parallel, limited by slowest of 2"),
        ("Total write p99",         "≈ 5–8 ms","Within 10 ms budget"),
    ]
    t = make_table(["Operation", "Latency", "Notes"],
                   lat_rows, col_widths=[160, 80, 200])
    elements.append(t)


def section_secondary_indexes(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("12. Advanced Topic: Secondary Indexes",
        "Looking up by value, not just by key — and why it's hard"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph(
        "A pure KV store only supports <b>point lookups by key</b>. "
        "Adding secondary indexes (query by value, range scans) "
        "is architecturally tricky because data is partitioned by primary key.", S_BODY))
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Approaches", S_H2))
    idx_rows = [
        ("Local secondary index\n(LSI)",
         "Index maintained per partition node",
         "Scatter-gather: fan out query to all nodes, merge results",
         "Fast writes; O(n) reads",
         "Cassandra CQL"),
        ("Global secondary index\n(GSI)",
         "Separate hash ring partitioned by index key",
         "Write to main table + async update index table",
         "Consistent reads; async write lag",
         "DynamoDB GSI"),
        ("Inverted index",
         "External search engine (Elasticsearch)",
         "Dual-write or CDC-based sync",
         "Full-text search; operational complexity",
         "OpenSearch + DynamoDB"),
        ("Materialised view",
         "Pre-computed query result stored as separate KV table",
         "Application or DB trigger on primary write",
         "Instant reads; write amplification",
         "Cassandra MV"),
    ]
    t = make_table(
        ["Approach", "How", "Query Mechanism", "Trade-off", "Example"],
        idx_rows, col_widths=[80, 90, 100, 95, 75])
    elements.append(t)
    elements.append(Spacer(1, 6))

    elements.append(CalloutBox(
        "Interview insight: when asked to add secondary indexes, always discuss:\n"
        "  • Consistency between primary and index (async lag vs. 2PC overhead)\n"
        "  • Write amplification (each write may update multiple index tables)\n"
        "  • Scatter-gather cost for local indexes (fan-out to all shards)\n"
        "  • Operational complexity of maintaining index consistency at scale", kind="tip"))


def section_interview_qa(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("13. Interview Q&A",
        "Five questions every system design interviewer loves to ask"))
    elements.append(Spacer(1, 4))

    qa = [
        (
            "Q1. Walk through a complete write path in a Dynamo-style KV store.",
            [
                "① Client hashes the key (MurmurHash3) to find the coordinator node "
                "   from its local ring map.",
                "② Coordinator writes to its local storage (WAL → MemTable) for durability.",
                "③ Coordinator forwards the write in parallel to N−1 replicas (N=3 → 2 more nodes).",
                "④ Each replica writes to WAL + MemTable and sends an ACK.",
                "⑤ Coordinator waits for W=2 total ACKs (including itself). Returns success to client.",
                "⑥ Background: third replica eventually receives the write (async). "
                "   Read repair corrects any lag on the next GET.",
                "Key points to mention: vector clock attached, sloppy quorum, hinted handoff "
                "if a replica is down.",
            ]
        ),
        (
            "Q2. How does the system handle a network partition isolating 2 of 5 nodes?",
            [
                "Partition scenario: nodes {N4, N5} are isolated; {N1, N2, N3} form the majority.",
                "• N1–N3 form a sloppy quorum and continue accepting writes (W=2 from 3 nodes).",
                "• If a key's preference list spans both partitions, the coordinator on the "
                "  majority side uses the next available nodes (hinted handoff for N4/N5).",
                "• N4, N5 can still serve reads/writes among themselves but risk divergence.",
                "• On partition heal: anti-entropy kicks in; Merkle trees identify diverged keys; "
                "  conflict resolution merges or LWW resolves them.",
                "• Availability is preserved for the majority side; eventual consistency "
                "  handles divergence.",
            ]
        ),
        (
            "Q3. Explain how Merkle trees enable efficient anti-entropy synchronisation.",
            [
                "Each node builds a Merkle tree over its key range: leaf nodes hash individual "
                "key-value pairs; internal nodes hash their children.",
                "Synchronisation protocol between nodes A and B:",
                "  1. Exchange root hashes → equal means in sync (done!).",
                "  2. If different, recursively compare left/right child hashes.",
                "  3. Recurse only into divergent subtrees → O(log n) comparisons.",
                "  4. Sync only the leaf-level keys that differ → minimal bandwidth.",
                "Without Merkle trees: must compare every key hash — O(n) bandwidth.",
                "With Merkle trees: O(log n) steps to pinpoint differences — "
                "critical for millions of keys.",
            ]
        ),
        (
            "Q4. What is hinted handoff and why is it important for write availability?",
            [
                "Hinted handoff allows writes to proceed even when target replicas are down.",
                "Mechanism: if node N3 (intended replica) is unreachable, the coordinator "
                "routes the write to the next node on the ring (N4) and attaches a 'hint': "
                "{'intended_for': 'N3', 'key': ..., 'value': ...}.",
                "N4 stores this hint in a special hints database (not its normal keyspace).",
                "When N3 recovers, N4 delivers all queued hints and N3 re-integrates.",
                "Why critical: without hinted handoff, a write to N=3 replicas would fail "
                "as soon as any one replica is down → write availability collapses. "
                "With it, the system remains 'always writable'.",
                "Caveat: if N4 also fails before delivering the hint, the data may be lost "
                "→ durability depends on W and the number of simultaneous failures.",
            ]
        ),
        (
            "Q5. How would you add support for secondary indexes?",
            [
                "Step 1: Decide on index type based on access patterns:",
                "  • Local index (partition-local): fast writes, requires scatter-gather reads.",
                "  • Global index (separate partition): consistent reads, async write lag.",
                "Step 2: Choose a consistency model for the index:",
                "  • Synchronous (2PC or paxos): strong consistency, high latency.",
                "  • Asynchronous (CDC/event): low write latency, eventual index consistency.",
                "Step 3: Handle write amplification:",
                "  • Each PUT to primary key triggers a PUT to index table(s).",
                "  • Use batch writes or a transactional log to keep them in sync.",
                "Step 4: Handle index reads:",
                "  • Range scan on index returns a set of primary keys.",
                "  • Secondary lookup fetches values by primary key.",
                "Real-world: DynamoDB GSI uses async propagation with ~seconds lag; "
                "Cassandra materialised views replicate synchronously within a partition.",
            ]
        ),
    ]

    for i, (question, answers) in enumerate(qa):
        elements.append(KeepTogether([
            Paragraph(question, S_QA_Q),
            *[Paragraph(ans, S_QA_A) for ans in answers],
            Spacer(1, 6),
        ]))


def section_references(elements):
    elements.append(PageBreak())
    elements.append(SectionHeader("14. References & Further Reading",
        "Papers, blog posts, and talks to go deeper"))
    elements.append(Spacer(1, 4))

    elements.append(Paragraph("Foundational Papers", S_H2))
    papers = [
        ("Amazon Dynamo",
         "DeCandia et al. Dynamo: Amazon's Highly Available Key-value Store. SOSP 2007.",
         "The original blueprint. Must-read."),
        ("Cassandra",
         "Lakshman & Malik. Cassandra: A Decentralized Structured Storage System. SIGOPS 2010.",
         "Combines Dynamo's distribution with Bigtable's data model."),
        ("Google Bigtable",
         "Chang et al. Bigtable: A Distributed Storage System for Structured Data. OSDI 2006.",
         "Foundational paper for sorted KV stores and LSM."),
        ("Spanner",
         "Corbett et al. Spanner: Google's Globally Distributed Database. OSDI 2012.",
         "Strong consistency at global scale using TrueTime."),
        ("Raft",
         "Ongaro & Ousterhout. In Search of an Understandable Consensus Algorithm. USENIX 2014.",
         "Easier-to-understand alternative to Paxos."),
    ]
    for (title, cite, note) in papers:
        elements.append(Paragraph(f"• <b>{title}</b>: {cite} <i>({note})</i>", S_REF))
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Essential Blog Posts & Guides", S_H2))
    blogs = [
        "Martin Kleppmann — Designing Data-Intensive Applications (O'Reilly) — Chapter 5–7",
        "High Scalability — blog.highscalability.com — case studies on Cassandra, Redis, DynamoDB",
        "Discord Engineering Blog — 'How Discord Stores Billions of Messages' (2017, 2022, 2023)",
        "AWS DynamoDB Developer Guide — docs.aws.amazon.com/amazondynamodb",
        "Datastax — 'How Apache Cassandra Balances Consistency, Availability, and Performance'",
        "Facebook Engineering — 'RocksDB: A Persistent Key-Value Store for Fast Storage'",
        "Martin Fowler — 'CQRS' and 'Event Sourcing' patterns that build on KV stores",
        "Werner Vogels — 'Eventually Consistent' (ACM Queue, 2008) — written by Amazon's CTO",
    ]
    for b in blogs:
        elements.append(Paragraph(f"• {b}", S_REF))
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("YouTube / Conference Talks", S_H2))
    talks = [
        "MIT 6.824 Distributed Systems — lectures on Dynamo, Raft, Zookeeper (YouTube)",
        "GOTO 2017 — 'Designing for Failure in a Microservices World' — Adrian Cockcroft",
        "StrangeLoop 2014 — 'Consistency without Consensus' — Peter Bailis",
        "QCon London — 'Inside Cassandra's Storage Engine' — Benedict Elliott Smith",
        "AWS re:Invent — 'DynamoDB Under the Hood: How We Built a Hyper-Scale Database'",
        "CMU Database Group — 'Advanced Database Systems: Key-Value Stores' (15-721)",
    ]
    for t in talks:
        elements.append(Paragraph(f"• {t}", S_REF))
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Open-Source Code to Study", S_H2))
    oss = [
        ("etcd",      "github.com/etcd-io/etcd",        "Raft-based KV; read the raft/ package"),
        ("RocksDB",   "github.com/facebook/rocksdb",     "Gold standard LSM storage engine"),
        ("Riak KV",   "github.com/basho/riak_kv",        "Full Dynamo implementation in Erlang"),
        ("TiKV",      "github.com/tikv/tikv",            "Production Raft KV in Rust"),
        ("Bitcask",   "github.com/basho/bitcask",        "Simple hash-index log KV; great to read first"),
    ]
    t = make_table(["Project", "Repository", "Why Read It"],
                   oss, col_widths=[70, 160, 210])
    elements.append(t)


# ─────────────────────────────────────────────
# PAGE DECORATIONS
# ─────────────────────────────────────────────

def on_first_page(canvas, doc):
    pass


def on_later_pages(canvas, doc):
    canvas.saveState()
    w = W
    # Footer line
    canvas.setStrokeColor(C_MID_GRAY)
    canvas.setLineWidth(0.5)
    canvas.line(1*cm, 1.5*cm, w - 1*cm, 1.5*cm)
    # Footer text
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(C_GRAY)
    canvas.drawString(1*cm, 1*cm, "Distributed Key-Value Store — System Design Blueprint")
    canvas.drawRightString(w - 1*cm, 1*cm, f"Page {doc.page}")
    # Top rule
    canvas.setStrokeColor(C_TEAL)
    canvas.setLineWidth(2)
    canvas.line(1*cm, H - 1.2*cm, w - 1*cm, H - 1.2*cm)
    canvas.restoreState()


# ─────────────────────────────────────────────
# MAIN BUILD
# ─────────────────────────────────────────────

def build_pdf(output_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=1*cm,
        rightMargin=1*cm,
        topMargin=1.8*cm,
        bottomMargin=2*cm,
        title="Distributed Key-Value Store — System Design Blueprint",
        author="System Design Series",
        subject="KV Store Architecture",
    )

    elements = []
    cover_page(elements)
    section_problem(elements)
    section_cap(elements)
    section_requirements(elements)
    section_architecture(elements)
    section_partitioning(elements)
    section_replication(elements)
    section_consistency(elements)
    section_failure(elements)
    section_storage(elements)
    section_real_world(elements)
    section_performance(elements)
    section_secondary_indexes(elements)
    section_interview_qa(elements)
    section_references(elements)

    doc.build(elements,
              onFirstPage=on_first_page,
              onLaterPages=on_later_pages)
    print(f"✅  PDF written to: {output_path}")


if __name__ == "__main__":
    import os
    out = os.path.join(os.path.dirname(__file__), "kv_store_system_design.pdf")
    build_pdf(out)
