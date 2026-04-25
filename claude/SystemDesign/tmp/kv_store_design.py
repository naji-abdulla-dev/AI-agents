"""
Distributed Key-Value Store System Design - PDF Generator
"""

import math
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import (
    HexColor, black, white, grey, lightgrey, darkgrey,
    Color
)
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether, Flowable
)
from reportlab.graphics.shapes import (
    Drawing, Circle, Rect, Line, String,
    Polygon, Ellipse, Path, Group
)
from reportlab.graphics import renderPDF

# ─── Colour Palette ──────────────────────────────────────────────────────────
C_NAVY    = HexColor('#0D1B2A')   # deep navy – headings
C_BLUE    = HexColor('#1565C0')   # primary blue
C_LTBLUE  = HexColor('#42A5F5')   # accent
C_TEAL    = HexColor('#00838F')   # secondary accent
C_GREEN   = HexColor('#2E7D32')
C_AMBER   = HexColor('#E65100')
C_RED     = HexColor('#C62828')
C_BG      = HexColor('#F8FAFC')   # page background tint (tables)
C_STRIPE  = HexColor('#E3F2FD')   # table alternating row
C_BORDER  = HexColor('#90CAF9')
C_GREY    = HexColor('#546E7A')
C_LGREY   = HexColor('#ECEFF1')

# ─── Styles ──────────────────────────────────────────────────────────────────
styles = getSampleStyleSheet()

def S(name, **kw):
    return ParagraphStyle(name, **kw)

TITLE_S = S('Title',
    fontName='Helvetica-Bold', fontSize=28, textColor=C_NAVY,
    spaceAfter=6, alignment=TA_CENTER, leading=34)

SUBTITLE_S = S('SubTitle',
    fontName='Helvetica', fontSize=14, textColor=C_GREY,
    spaceAfter=4, alignment=TA_CENTER)

H1 = S('H1',
    fontName='Helvetica-Bold', fontSize=18, textColor=C_NAVY,
    spaceBefore=18, spaceAfter=8, leading=22,
    borderPad=4)

H2 = S('H2',
    fontName='Helvetica-Bold', fontSize=13, textColor=C_BLUE,
    spaceBefore=12, spaceAfter=5, leading=17)

H3 = S('H3',
    fontName='Helvetica-BoldOblique', fontSize=11, textColor=C_TEAL,
    spaceBefore=8, spaceAfter=4, leading=14)

BODY = S('Body',
    fontName='Helvetica', fontSize=10, textColor=HexColor('#1A1A2E'),
    spaceBefore=3, spaceAfter=4, leading=15, alignment=TA_JUSTIFY)

BULLET = S('Bullet',
    fontName='Helvetica', fontSize=10, textColor=HexColor('#1A1A2E'),
    spaceBefore=2, spaceAfter=2, leading=14,
    leftIndent=16, bulletIndent=6)

CODE = S('Code',
    fontName='Courier', fontSize=9, textColor=HexColor('#1A1A2E'),
    spaceBefore=4, spaceAfter=4, leading=13,
    leftIndent=12, backColor=C_LGREY)

CAPTION = S('Caption',
    fontName='Helvetica-Oblique', fontSize=9, textColor=C_GREY,
    alignment=TA_CENTER, spaceBefore=2, spaceAfter=8)

TABLE_HDR = S('TableHdr',
    fontName='Helvetica-Bold', fontSize=9, textColor=white,
    alignment=TA_CENTER)

TABLE_CELL = S('TableCell',
    fontName='Helvetica', fontSize=9, textColor=HexColor('#1A1A2E'),
    alignment=TA_LEFT, leading=12)

ANSWER = S('Answer',
    fontName='Helvetica', fontSize=10, textColor=HexColor('#0D3B66'),
    spaceBefore=2, spaceAfter=4, leading=14,
    leftIndent=12, alignment=TA_JUSTIFY)

NOTE = S('Note',
    fontName='Helvetica-Oblique', fontSize=9, textColor=C_GREY,
    spaceBefore=2, spaceAfter=4, leading=13, leftIndent=8)


# ─── Helper Flowables ────────────────────────────────────────────────────────

class SectionRule(Flowable):
    """Coloured rule used as section divider."""
    def __init__(self, color=C_BLUE, thickness=2, width=None):
        super().__init__()
        self.color = color
        self.thickness = thickness
        self._width = width

    def wrap(self, availW, availH):
        self._avail = availW
        return (self._width or availW), self.thickness + 4

    def draw(self):
        w = self._width or self._avail
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 2, w, self.thickness, fill=1, stroke=0)


class CalloutBox(Flowable):
    """Highlighted callout / info box."""
    def __init__(self, text, bg=C_STRIPE, border=C_BORDER, width=None):
        super().__init__()
        self.text = text
        self.bg = bg
        self.border = border
        self._width = width
        self._style = S('cb', fontName='Helvetica', fontSize=10,
                        textColor=C_NAVY, leading=14,
                        leftIndent=8, rightIndent=8)

    def wrap(self, aW, aH):
        self._avail = aW
        w = self._width or aW
        p = Paragraph(self.text, self._style)
        _, h = p.wrap(w - 24, aH)
        self._h = h + 16
        self._w = w
        return w, self._h

    def draw(self):
        c = self.canv
        c.setFillColor(self.bg)
        c.setStrokeColor(self.border)
        c.setLineWidth(1.2)
        c.roundRect(0, 0, self._w, self._h, 6, fill=1, stroke=1)
        p = Paragraph(self.text, self._style)
        p.wrap(self._w - 24, self._h)
        p.drawOn(c, 12, 8)


def p(text, style=BODY): return Paragraph(text, style)
def h1(text): return p(text, H1)
def h2(text): return p(text, H2)
def h3(text): return p(text, H3)
def sp(n=6): return Spacer(1, n)
def rule(color=C_BLUE): return SectionRule(color)
def pb(): return PageBreak()


def bullet_list(items, style=BULLET):
    return [Paragraph(f'&#8226;  {item}', style) for item in items]


def make_table(headers, rows, col_widths=None, stripe=True):
    """Create a styled table with header and optional striped rows."""
    hdr_cells = [Paragraph(h, TABLE_HDR) for h in headers]
    data = [hdr_cells]
    for i, row in enumerate(rows):
        cells = [Paragraph(str(c), TABLE_CELL) for c in row]
        data.append(cells)

    style_cmds = [
        ('BACKGROUND',  (0, 0), (-1, 0),  C_NAVY),
        ('TEXTCOLOR',   (0, 0), (-1, 0),  white),
        ('FONTNAME',    (0, 0), (-1, 0),  'Helvetica-Bold'),
        ('FONTSIZE',    (0, 0), (-1, 0),  9),
        ('ALIGN',       (0, 0), (-1, 0),  'CENTER'),
        ('VALIGN',      (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1),
            [C_BG, C_STRIPE] if stripe else [C_BG]),
        ('GRID',        (0, 0), (-1, -1), 0.5, C_BORDER),
        ('FONTNAME',    (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE',    (0, 1), (-1, -1), 9),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING',(0, 0), (-1, -1), 6),
        ('TOPPADDING',  (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING',(0,0), (-1, -1), 5),
    ]
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle(style_cmds))
    return t


# ─── Diagrams ────────────────────────────────────────────────────────────────

def draw_hash_ring(w=480, h=340):
    """Consistent hash ring with 5 physical nodes + virtual nodes."""
    d = Drawing(w, h)
    cx, cy, R = w / 2, h / 2 - 10, 130

    # Background circle (ring track)
    d.add(Circle(cx, cy, R + 12, fillColor=C_LGREY, strokeColor=C_BORDER, strokeWidth=1))
    d.add(Circle(cx, cy, R - 12, fillColor=white, strokeColor=C_BORDER, strokeWidth=1))

    # Title
    d.add(String(cx, h - 18, 'Consistent Hash Ring  (N=5 physical nodes, ~150 vnodes each)',
                 fontSize=10, textAnchor='middle', fillColor=C_NAVY,
                 fontName='Helvetica-Bold'))

    # Physical nodes
    node_colors = [C_BLUE, C_TEAL, C_GREEN, C_AMBER, C_RED]
    node_labels = ['A', 'B', 'C', 'D', 'E']
    angles_deg  = [90, 162, 234, 306, 18]  # evenly spaced, start at top

    for i, (ang, label, col) in enumerate(zip(angles_deg, node_labels, node_colors)):
        rad = math.radians(ang)
        nx = cx + R * math.cos(rad)
        ny = cy + R * math.sin(rad)

        # vnode tick marks (show 3 representative vnodes per physical)
        for dv in [-25, 0, 25]:
            vrad = math.radians(ang + dv)
            vx = cx + R * math.cos(vrad)
            vy = cy + R * math.sin(vrad)
            ix = cx + (R - 12) * math.cos(vrad)
            iy = cy + (R - 12) * math.sin(vrad)
            ox = cx + (R + 12) * math.cos(vrad)
            oy = cy + (R + 12) * math.sin(vrad)
            d.add(Line(ix, iy, ox, oy, strokeColor=col, strokeWidth=1.2))

        # Node circle
        d.add(Circle(nx, ny, 16, fillColor=col, strokeColor=white, strokeWidth=2))
        d.add(String(nx, ny - 5, label, fontSize=11, textAnchor='middle',
                     fillColor=white, fontName='Helvetica-Bold'))

        # Label offset outward
        lx = cx + (R + 36) * math.cos(rad)
        ly = cy + (R + 36) * math.sin(rad)
        d.add(String(lx, ly - 4, f'Node {label}', fontSize=8.5,
                     textAnchor='middle', fillColor=col, fontName='Helvetica-Bold'))

    # Key example on ring
    key_ang = math.radians(50)
    kx = cx + R * math.cos(key_ang)
    ky = cy + R * math.sin(key_ang)
    d.add(Circle(kx, ky, 6, fillColor=C_AMBER, strokeColor=C_NAVY, strokeWidth=1.5))
    d.add(String(kx + 12, ky + 2, '"user:42"', fontSize=7.5,
                 textAnchor='start', fillColor=C_AMBER, fontName='Helvetica-Bold'))

    # Legend
    lx0, ly0 = 14, 14
    for i, (label, col) in enumerate(zip(node_labels, node_colors)):
        d.add(Circle(lx0 + i * 62, ly0, 6, fillColor=col, strokeColor=white, strokeWidth=1))
        d.add(String(lx0 + i * 62 + 10, ly0 - 4, f'Node {label}', fontSize=8,
                     textAnchor='start', fillColor=col, fontName='Helvetica'))

    d.add(Circle(lx0 + 5 * 62, ly0, 6, fillColor=C_AMBER, strokeColor=C_NAVY, strokeWidth=1.5))
    d.add(String(lx0 + 5 * 62 + 10, ly0 - 4, 'Key', fontSize=8,
                 textAnchor='start', fillColor=C_AMBER, fontName='Helvetica'))

    return d


def draw_replication(w=480, h=240):
    """Show a key hashing to coordinator, then replicated to N-1 successors."""
    d = Drawing(w, h)

    d.add(String(w / 2, h - 15, 'Replication: Key -> Coordinator -> N Successors  (N=3)',
                 fontSize=10, textAnchor='middle', fillColor=C_NAVY,
                 fontName='Helvetica-Bold'))

    node_x = [60, 160, 260, 360, 460]
    labels  = ['Node A', 'Node B', 'Node C', 'Node D', 'Node E']
    colors  = [C_BLUE, C_TEAL, C_GREEN, C_GREY, C_GREY]
    roles   = ['Coordinator\n(N1)', 'Replica\n(N2)', 'Replica\n(N3)', '', '']

    y_node  = h / 2 - 10

    # Key arrow
    d.add(String(node_x[0], y_node + 70, '"product:101"', fontSize=9,
                 textAnchor='middle', fillColor=C_AMBER, fontName='Helvetica-Bold'))
    d.add(Line(node_x[0], y_node + 58, node_x[0], y_node + 38,
               strokeColor=C_AMBER, strokeWidth=1.5))
    # arrowhead
    d.add(Polygon([node_x[0], y_node + 38, node_x[0] - 5, y_node + 48,
                   node_x[0] + 5, y_node + 48],
                  fillColor=C_AMBER, strokeColor=C_AMBER, strokeWidth=0))

    for i in range(3):
        col = colors[i]
        x   = node_x[i]
        # node box
        d.add(Rect(x - 30, y_node - 22, 60, 44,
                   fillColor=C_STRIPE if i > 0 else HexColor('#BBDEFB'),
                   strokeColor=col, strokeWidth=1.5, rx=5, ry=5))
        d.add(String(x, y_node + 4, labels[i], fontSize=8.5,
                     textAnchor='middle', fillColor=col, fontName='Helvetica-Bold'))
        # role label below
        role_lines = roles[i].split('\n')
        for j, rl in enumerate(role_lines):
            d.add(String(x, y_node - 32 - j * 11, rl, fontSize=8,
                         textAnchor='middle', fillColor=col,
                         fontName='Helvetica-BoldOblique'))

        # replication arrow
        if i < 2:
            d.add(Line(x + 32, y_node, node_x[i + 1] - 32, y_node,
                       strokeColor=C_TEAL, strokeWidth=2))
            mx = (x + node_x[i + 1]) / 2
            d.add(String(mx, y_node + 8, 'replicate', fontSize=7.5,
                         textAnchor='middle', fillColor=C_TEAL,
                         fontName='Helvetica-Oblique'))
            # arrowhead
            ax = node_x[i + 1] - 32
            d.add(Polygon([ax, y_node, ax - 8, y_node + 5, ax - 8, y_node - 5],
                          fillColor=C_TEAL, strokeColor=C_TEAL, strokeWidth=0))

    # Inactive nodes
    for i in range(3, 5):
        x = node_x[i]
        d.add(Rect(x - 30, y_node - 22, 60, 44,
                   fillColor=C_LGREY, strokeColor=C_GREY, strokeWidth=1, rx=5, ry=5))
        d.add(String(x, y_node + 4, labels[i], fontSize=8.5,
                     textAnchor='middle', fillColor=C_GREY, fontName='Helvetica'))
        d.add(String(x, y_node - 12, '(not in', fontSize=7,
                     textAnchor='middle', fillColor=C_GREY, fontName='Helvetica'))
        d.add(String(x, y_node - 22, 'replica set)', fontSize=7,
                     textAnchor='middle', fillColor=C_GREY, fontName='Helvetica'))

    # Ring line beneath
    d.add(Line(20, y_node - 38, w - 20, y_node - 38,
               strokeColor=C_BORDER, strokeWidth=1.5, strokeDashArray=[4, 4]))
    d.add(String(w / 2, y_node - 52, '<-- Hash ring token space -->',
                 fontSize=8, textAnchor='middle', fillColor=C_GREY,
                 fontName='Helvetica-Oblique'))

    return d


def draw_write_path(w=500, h=300):
    """Write path flowchart: client -> coordinator -> replicas -> quorum ack."""
    d = Drawing(w, h)

    d.add(String(w / 2, h - 15, 'Write Path  (W=2, N=3)',
                 fontSize=10, textAnchor='middle', fillColor=C_NAVY,
                 fontName='Helvetica-Bold'))

    boxes = [
        (60,  'Client',        C_AMBER, 'PUT key=v'),
        (160, 'Coordinator',   C_BLUE,  'Hash -> ring\nAdd VClock'),
        (300, 'Replica N2',    C_TEAL,  'Persist\nAck W'),
        (420, 'Replica N3',    C_GREEN, 'Persist\n(async)'),
    ]

    y = h / 2 - 20

    for bx, label, col, sub in boxes:
        d.add(Rect(bx - 38, y - 26, 76, 52,
                   fillColor=HexColor('#E8F4FD') if col != C_AMBER else HexColor('#FFF3E0'),
                   strokeColor=col, strokeWidth=1.8, rx=5, ry=5))
        d.add(String(bx, y + 10, label, fontSize=8.5,
                     textAnchor='middle', fillColor=col, fontName='Helvetica-Bold'))
        for j, sl in enumerate(sub.split('\n')):
            d.add(String(bx, y - 4 - j * 11, sl, fontSize=7.5,
                         textAnchor='middle', fillColor=C_GREY, fontName='Helvetica'))

    # Arrows
    arrows = [
        (60 + 38, 160 - 38, y, '1. PUT'),
        (160 + 38, 300 - 38, y, '2. forward'),
        (160 + 38, 420 - 38, y - 20, '2. forward'),
        (300 - 38, 160 + 38, y - 40, '3. ACK'),
    ]
    for x1, x2, ay, lbl in arrows:
        d.add(Line(x1, ay, x2, ay, strokeColor=C_TEAL, strokeWidth=1.5))
        d.add(String((x1 + x2) / 2, ay + 6, lbl, fontSize=7.5,
                     textAnchor='middle', fillColor=C_TEAL, fontName='Helvetica-Oblique'))

    # Quorum box
    d.add(Rect(130, y - 85, 120, 34,
               fillColor=HexColor('#E8F5E9'), strokeColor=C_GREEN, strokeWidth=1.5, rx=4, ry=4))
    d.add(String(190, y - 60, 'Quorum Met (W=2)', fontSize=8.5,
                 textAnchor='middle', fillColor=C_GREEN, fontName='Helvetica-Bold'))
    d.add(String(190, y - 73, 'Return success to client', fontSize=8,
                 textAnchor='middle', fillColor=C_GREEN, fontName='Helvetica'))
    d.add(Line(190, y - 86, 190, y - 54,  strokeColor=C_GREEN, strokeWidth=1))

    # Hinted handoff note
    d.add(Rect(330, y + 50, 140, 30,
               fillColor=HexColor('#FFF8E1'), strokeColor=C_AMBER, strokeWidth=1, rx=3, ry=3))
    d.add(String(400, y + 72, 'N3 down? -> Hinted Handoff', fontSize=7.5,
                 textAnchor='middle', fillColor=C_AMBER, fontName='Helvetica-BoldOblique'))
    d.add(String(400, y + 60, 'Store hint, deliver later', fontSize=7.5,
                 textAnchor='middle', fillColor=C_AMBER, fontName='Helvetica'))

    return d


def draw_read_path(w=500, h=260):
    """Read path: coordinator reads from R replicas, returns latest."""
    d = Drawing(w, h)

    d.add(String(w / 2, h - 15, 'Read Path  (R=2, N=3)',
                 fontSize=10, textAnchor='middle', fillColor=C_NAVY,
                 fontName='Helvetica-Bold'))

    y = h / 2 - 10

    boxes = [
        (60,  'Client',       C_AMBER),
        (180, 'Coordinator',  C_BLUE),
        (320, 'Replica N2',   C_TEAL),
        (440, 'Replica N3',   C_GREEN),
    ]

    for bx, label, col in boxes:
        d.add(Rect(bx - 38, y - 22, 76, 44,
                   fillColor=HexColor('#E8F4FD') if col != C_AMBER else HexColor('#FFF3E0'),
                   strokeColor=col, strokeWidth=1.8, rx=5, ry=5))
        d.add(String(bx, y + 4, label, fontSize=8.5,
                     textAnchor='middle', fillColor=col, fontName='Helvetica-Bold'))

    # Arrows
    fwd = [
        (60 + 38, 180 - 38, y + 8,  '1. GET'),
        (180 + 38, 320 - 38, y + 8, '2. read'),
        (180 + 38, 440 - 38, y + 8, '2. read'),
    ]
    for x1, x2, ay, lbl in fwd:
        d.add(Line(x1, ay, x2, ay, strokeColor=C_BLUE, strokeWidth=1.5))
        d.add(String((x1 + x2) / 2, ay + 6, lbl, fontSize=7.5,
                     textAnchor='middle', fillColor=C_BLUE, fontName='Helvetica-Oblique'))

    ack = [
        (320 - 38, 180 + 38, y - 8, 'v@vc[3]'),
        (440 - 38, 180 + 38, y - 20, 'v@vc[2] (stale)'),
    ]
    for x1, x2, ay, lbl in ack:
        d.add(Line(x1, ay, x2, ay, strokeColor=C_GREEN, strokeWidth=1.5))
        d.add(String((x1 + x2) / 2, ay - 8, lbl, fontSize=7.5,
                     textAnchor='middle', fillColor=C_GREEN, fontName='Helvetica-Oblique'))

    # Read repair
    d.add(Rect(290, y - 80, 170, 32,
               fillColor=HexColor('#FCE4EC'), strokeColor=C_RED, strokeWidth=1, rx=3, ry=3))
    d.add(String(375, y - 58, 'Read Repair Triggered', fontSize=8.5,
                 textAnchor='middle', fillColor=C_RED, fontName='Helvetica-Bold'))
    d.add(String(375, y - 71, 'Coordinator pushes latest -> stale replica', fontSize=7.5,
                 textAnchor='middle', fillColor=C_RED, fontName='Helvetica'))

    d.add(Line(440, y - 22, 440, y - 48, strokeColor=C_RED,
               strokeWidth=1.2, strokeDashArray=[3, 3]))
    d.add(String(448, y - 36, 'repair', fontSize=7.5, textAnchor='start',
                 fillColor=C_RED, fontName='Helvetica-BoldOblique'))

    return d


def draw_vector_clock(w=500, h=200):
    """Vector clock example showing concurrent writes and conflict."""
    d = Drawing(w, h)

    d.add(String(w / 2, h - 15, 'Vector Clocks - Detecting Concurrent Writes',
                 fontSize=10, textAnchor='middle', fillColor=C_NAVY,
                 fontName='Helvetica-Bold'))

    nodes = ['Client 1', 'Node A', 'Node B']
    x_pos = [80, 240, 420]
    colors = [C_AMBER, C_BLUE, C_TEAL]

    y_top = h - 40
    y_bot = 30

    for x, label, col in zip(x_pos, nodes, colors):
        d.add(Line(x, y_top, x, y_bot, strokeColor=col, strokeWidth=1.5,
                   strokeDashArray=[3, 3]))
        d.add(String(x, y_top + 4, label, fontSize=9, textAnchor='middle',
                     fillColor=col, fontName='Helvetica-Bold'))

    events = [
        # (x_from, x_to, y, label, col)
        (80,  240, 140, 'write "x=1"  vc:{A:1}',    C_BLUE),
        (240, 420,  115, 'replicate   vc:{A:1}',      C_TEAL),
        (80,  240,  90, 'write "x=2"  vc:{A:2}',    C_BLUE),   # concurrent!
        (420, 240,  70, 'write "x=3"  vc:{B:1}',    C_AMBER),  # concurrent from B
    ]

    for x1, x2, ey, lbl, col in events:
        d.add(Line(x1, ey, x2, ey, strokeColor=col, strokeWidth=1.5))
        d.add(Circle(x2, ey, 5, fillColor=col, strokeColor=white, strokeWidth=1))
        mx = (x1 + x2) / 2
        d.add(String(mx, ey + 7, lbl, fontSize=7.5, textAnchor='middle',
                     fillColor=col, fontName='Helvetica'))

    # Conflict box
    d.add(Rect(145, 30, 190, 30,
               fillColor=HexColor('#FCE4EC'), strokeColor=C_RED, strokeWidth=1.5, rx=3, ry=3))
    d.add(String(240, 52, 'CONFLICT: vc:{A:2} || vc:{B:1}', fontSize=8,
                 textAnchor='middle', fillColor=C_RED, fontName='Helvetica-Bold'))
    d.add(String(240, 40, 'Application resolves (LWW or merge)', fontSize=7.5,
                 textAnchor='middle', fillColor=C_RED, fontName='Helvetica'))

    return d


def draw_merkle_tree(w=480, h=260):
    """Merkle tree for anti-entropy."""
    d = Drawing(w, h)

    d.add(String(w / 2, h - 15, 'Merkle Tree - Anti-Entropy Sync Between Replicas',
                 fontSize=10, textAnchor='middle', fillColor=C_NAVY,
                 fontName='Helvetica-Bold'))

    bw, bh = 64, 28

    def box(x, y, label, sub, col, highlight=False):
        fc = HexColor('#FFF9C4') if highlight else HexColor('#E3F2FD')
        sc = C_RED if highlight else col
        d.add(Rect(x - bw // 2, y - bh // 2, bw, bh,
                   fillColor=fc, strokeColor=sc,
                   strokeWidth=2 if highlight else 1.2, rx=4, ry=4))
        d.add(String(x, y + 4, label, fontSize=7.5, textAnchor='middle',
                     fillColor=sc, fontName='Helvetica-Bold'))
        d.add(String(x, y - 7, sub, fontSize=6.5, textAnchor='middle',
                     fillColor=C_GREY, fontName='Helvetica'))

    def edge(x1, y1, x2, y2):
        d.add(Line(x1, y1, x2, y2, strokeColor=C_BORDER, strokeWidth=1))

    # Level 0 - Root
    edge(w // 2, h - 40 - bh // 2, w // 2 - 120, h - 80 + bh // 2)
    edge(w // 2, h - 40 - bh // 2, w // 2 + 120, h - 80 + bh // 2)
    box(w // 2, h - 40, 'Root Hash', 'a4f2 != 9b1c', C_RED, highlight=True)

    # Level 1
    edge(w // 2 - 120, h - 80 - bh // 2, w // 2 - 180, h - 130 + bh // 2)
    edge(w // 2 - 120, h - 80 - bh // 2, w // 2 - 60,  h - 130 + bh // 2)
    edge(w // 2 + 120, h - 80 - bh // 2, w // 2 + 60,  h - 130 + bh // 2)
    edge(w // 2 + 120, h - 80 - bh // 2, w // 2 + 180, h - 130 + bh // 2)
    box(w // 2 - 120, h - 80, 'Node L', 'keys 0-127', C_RED, highlight=True)
    box(w // 2 + 120, h - 80, 'Node R', 'keys 128-255', C_GREEN)

    # Level 2
    y2 = h - 130
    for i, (xc, lbl, sub, hl) in enumerate([
        (w // 2 - 180, 'L0-63',  'c3d1...',  True),
        (w // 2 - 60,  'L64-127','ab34...',  False),
        (w // 2 + 60,  'R128-191','ef89...', False),
        (w // 2 + 180, 'R192-255','12cd...', False),
    ]):
        col = C_RED if hl else C_TEAL
        box(xc, y2, lbl, sub, col, highlight=hl)

    # Level 3 (leaves for the differing segment)
    y3 = h - 185
    edge(w // 2 - 180, y2 - bh // 2, w // 2 - 210, y3 + bh // 2)
    edge(w // 2 - 180, y2 - bh // 2, w // 2 - 150, y3 + bh // 2)
    box(w // 2 - 210, y3, 'key 001', '7f3a...', C_RED, highlight=True)
    box(w // 2 - 150, y3, 'key 002', 'ab21...', C_TEAL)

    # Legend
    d.add(Rect(14, 14, 14, 10, fillColor=HexColor('#FFF9C4'), strokeColor=C_RED, strokeWidth=1.5, rx=2, ry=2))
    d.add(String(32, 18, 'Hash mismatch (sync needed)', fontSize=7.5, textAnchor='start',
                 fillColor=C_RED, fontName='Helvetica'))
    d.add(Rect(200, 14, 14, 10, fillColor=HexColor('#E3F2FD'), strokeColor=C_TEAL, strokeWidth=1, rx=2, ry=2))
    d.add(String(218, 18, 'Hashes match (in sync)', fontSize=7.5, textAnchor='start',
                 fillColor=C_TEAL, fontName='Helvetica'))

    return d


def draw_gossip(w=460, h=240):
    """Gossip protocol: fan-out membership update."""
    d = Drawing(w, h)

    d.add(String(w / 2, h - 15, 'Gossip Protocol - Membership & Failure Detection',
                 fontSize=10, textAnchor='middle', fillColor=C_NAVY,
                 fontName='Helvetica-Bold'))

    positions = {
        'A': (w // 2,     h - 50),
        'B': (w // 2 - 140, h - 110),
        'C': (w // 2 + 140, h - 110),
        'D': (w // 2 - 80,  50),
        'E': (w // 2 + 80,  50),
    }
    colors = {'A': C_BLUE, 'B': C_TEAL, 'C': C_GREEN,
              'D': C_AMBER, 'E': C_GREY}
    failed = {'E'}

    # Gossip edges
    gossip_edges = [('A','B'), ('A','C'), ('B','D'), ('C','D'), ('B','C')]
    for n1, n2 in gossip_edges:
        x1, y1 = positions[n1]
        x2, y2 = positions[n2]
        d.add(Line(x1, y1, x2, y2, strokeColor=C_LTBLUE,
                   strokeWidth=1.5, strokeDashArray=[5, 3]))
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        d.add(String(mx + 4, my, '[msg]', fontSize=7.5, textAnchor='middle',
                     fillColor=C_LTBLUE, fontName='Helvetica'))

    for node, (nx, ny) in positions.items():
        col = C_RED if node in failed else colors[node]
        fc  = HexColor('#FCE4EC') if node in failed else HexColor('#E3F2FD')
        d.add(Circle(nx, ny, 22, fillColor=fc, strokeColor=col, strokeWidth=2))
        d.add(String(nx, ny + 3, node, fontSize=12, textAnchor='middle',
                     fillColor=col, fontName='Helvetica-Bold'))
        if node in failed:
            d.add(String(nx, ny - 13, 'DOWN', fontSize=6.5, textAnchor='middle',
                         fillColor=C_RED, fontName='Helvetica-Bold'))

    # Legend
    d.add(Line(14, 18, 44, 18, strokeColor=C_LTBLUE, strokeWidth=1.5,
               strokeDashArray=[5, 3]))
    d.add(String(50, 14, 'Gossip heartbeat', fontSize=8, textAnchor='start',
                 fillColor=C_LTBLUE, fontName='Helvetica'))
    d.add(Circle(200, 18, 7, fillColor=HexColor('#FCE4EC'), strokeColor=C_RED, strokeWidth=1.5))
    d.add(String(212, 14, 'Failed node', fontSize=8, textAnchor='start',
                 fillColor=C_RED, fontName='Helvetica'))

    return d


def draw_architecture(w=540, h=320):
    """High-level architecture: clients, load balancer, cluster, storage."""
    d = Drawing(w, h)

    d.add(String(w / 2, h - 15, 'High-Level Architecture',
                 fontSize=11, textAnchor='middle', fillColor=C_NAVY,
                 fontName='Helvetica-Bold'))

    # Layer definitions
    layers = [
        (h - 50,  [130, 270, 410],   ['Client A', 'Client B', 'Client C'],
         C_AMBER,  HexColor('#FFF3E0'), 40, 22),
        (h - 110, [w / 2],            ['Load Balancer / Router'],
         C_BLUE,   HexColor('#BBDEFB'), 120, 22),
        (h - 175, [80, 200, 320, 440], ['Node A\n(Coord)', 'Node B', 'Node C', 'Node D'],
         C_TEAL,   HexColor('#E0F2F1'), 55, 28),
        (50,       [80, 200, 320, 440], ['RocksDB', 'RocksDB', 'RocksDB', 'RocksDB'],
         C_GREEN,  HexColor('#E8F5E9'), 55, 22),
    ]

    prev_layer_y   = None
    prev_layer_xs  = None

    for (ly, xs, labels, col, fc, bw, bh) in layers:
        # Boxes
        for x, lbl in zip(xs, labels):
            d.add(Rect(x - bw // 2, ly - bh // 2, bw, bh,
                       fillColor=fc, strokeColor=col, strokeWidth=1.5, rx=4, ry=4))
            for j, line in enumerate(lbl.split('\n')):
                offset = 4 if '\n' not in lbl else (7 - j * 11)
                d.add(String(x, ly + offset, line, fontSize=8,
                             textAnchor='middle', fillColor=col,
                             fontName='Helvetica-Bold'))

        # Connections to previous layer
        if prev_layer_xs is not None:
            for px in prev_layer_xs:
                for cx in xs:
                    d.add(Line(px, prev_layer_y - bh // 2 - 2,
                               cx, ly + bh // 2 + 2,
                               strokeColor=C_BORDER, strokeWidth=0.8))

        prev_layer_y  = ly
        prev_layer_xs = xs

    # Layer labels on the right
    label_data = [
        (h - 50,  'Client Tier',   C_AMBER),
        (h - 110, 'Routing',       C_BLUE),
        (h - 175, 'KV Node Layer', C_TEAL),
        (50,      'Storage Layer', C_GREEN),
    ]
    for yl, txt, col in label_data:
        d.add(String(w - 8, yl - 4, txt, fontSize=8, textAnchor='end',
                     fillColor=col, fontName='Helvetica-BoldOblique'))

    # Gossip horizontal arrows between KV nodes
    node_xs = [80, 200, 320, 440]
    for i in range(len(node_xs) - 1):
        x1, x2 = node_xs[i] + 28, node_xs[i + 1] - 28
        d.add(Line(x1, h - 175, x2, h - 175,
                   strokeColor=C_LTBLUE, strokeWidth=1,
                   strokeDashArray=[3, 2]))

    d.add(String(260, h - 155, '<-- gossip / replication -->', fontSize=7.5,
                 textAnchor='middle', fillColor=C_LTBLUE,
                 fontName='Helvetica-Oblique'))

    return d


# ─── Document Builder ────────────────────────────────────────────────────────

def build_pdf(path):
    doc = SimpleDocTemplate(
        path,
        pagesize=letter,
        rightMargin=0.75 * inch,
        leftMargin=0.75 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
        title='Distributed Key-Value Store - System Design',
        author='System Design Deep Dive',
    )

    W = letter[0] - 1.5 * inch   # usable content width

    story = []

    # TITLE PAGE
    story += [
        sp(30),
        p('Distributed Key-Value Store', TITLE_S),
        p('System Design Deep Dive', SUBTITLE_S),
        sp(4),
        rule(C_TEAL),
        sp(10),
        p('Inspired by Amazon Dynamo  |  Apache Cassandra  |  Redis Cluster',
          S('sub2', fontName='Helvetica-Oblique', fontSize=11, textColor=C_GREY,
            alignment=TA_CENTER, leading=16)),
        sp(6),
        p('Partitioning  |  Replication  |  Consistency  |  Failure Handling',
          S('sub3', fontName='Helvetica', fontSize=10, textColor=C_LTBLUE,
            alignment=TA_CENTER, leading=14)),
        sp(40),
        draw_architecture(540, 300),
        p('Figure 1 - Overall cluster architecture', CAPTION),
        pb(),
    ]

    # 1. PROBLEM STATEMENT
    story += [
        h1('1.  The Problem - "A Galaxy-Scale Filing Cabinet"'),
        rule(C_BLUE),
        sp(6),
        p('''Imagine you run the world\'s busiest shopping website - Black Friday, every single
        day. Millions of customers are hammering your servers simultaneously: reading product
        listings, updating shopping carts, writing session tokens. A single relational database
        keels over in seconds. You need something that can eat a million requests per second for
        breakfast, survive disk failures, network partitions, and entire data-centre outages -
        and still hand back the right value when you ask for it.'''),
        sp(4),
        p('''Enter the <b>distributed key-value store</b>: the simplest possible interface
        (<code>put(k,v)</code> and <code>get(k)</code>) wrapped around a fearsomely clever
        distributed system that spans hundreds of commodity servers across multiple availability
        zones. Amazon built Dynamo for exactly this reason in 2007, and the ideas in that paper
        now underpin Cassandra, DynamoDB, Riak, and Voldemort.'''),
        sp(6),
        CalloutBox(
            '<b>One-liner:</b>  A distributed KV store is a hash table that lives across '
            'hundreds of machines, never loses data, and is always open for writes - '
            'even when parts of the cluster are on fire.',
            bg=HexColor('#E8F5E9'), border=C_GREEN
        ),
        sp(10),
    ]

    # 2. REQUIREMENTS & SCOPE
    story += [
        h1('2.  Requirements & Scope'),
        rule(C_BLUE),
        sp(6),
        h2('2.1  Functional Requirements'),
    ]
    story += bullet_list([
        '<b>put(key, value)</b> - store/update a value; key <= 256 B, value <= 1 MB',
        '<b>get(key)</b> - retrieve value (or a set of conflicting values with context)',
        '<b>delete(key)</b> - tombstone deletion',
        '<b>Context / versioning</b> - vector-clock context returned for conflict detection',
    ])
    story += [sp(8), h2('2.2  Non-Functional Requirements')]
    story += bullet_list([
        'Throughput: sustain <b>1 M+ QPS</b> across the cluster (reads + writes)',
        'Latency: p99 read/write < <b>10 ms</b> at peak load',
        'Availability: <b>always writable</b> even during network partitions (AP, not CP)',
        'Durability: <b>zero data loss</b> - every written key survives node failures',
        'Scalability: linear horizontal scale; adding nodes does not require downtime',
        'Tunable consistency: operators choose N, R, W trade-offs per use-case',
    ])
    story += [
        sp(10),
        make_table(
            ['Parameter', 'Typical Value', 'Description'],
            [
                ['N (replication factor)', '3',       'Number of replicas per key'],
                ['W (write quorum)',        '2',       'Min ACKs before write succeeds'],
                ['R (read quorum)',         '2',       'Min responses before read returns'],
                ['Virtual nodes (vnodes)', '150/node', 'Tokens per physical node on ring'],
                ['Max key size',           '256 B',   'Byte limit on key string'],
                ['Max value size',         '1 MB',    'Byte limit on value blob'],
                ['Gossip interval',        '1 s',     'Heartbeat frequency'],
                ['Hinted handoff TTL',     '1 hour',  'Max time to buffer hints'],
            ],
            col_widths=[W * 0.32, W * 0.22, W * 0.46],
        ),
        p('Table 1 - Core system parameters', CAPTION),
        sp(8),
    ]

    # 3. API DESIGN
    story += [
        h1('3.  API Design'),
        rule(C_BLUE),
        sp(6),
        h2('3.1  External HTTP/gRPC API'),
        make_table(
            ['Verb', 'Path', 'Request Body', 'Response', 'Notes'],
            [
                ['PUT',    '/kv/{key}',  '{ value, context }', '200 OK + new_context',
                 'Upsert; include old context to detect conflicts'],
                ['GET',    '/kv/{key}',  '-',                  '200 {value, context} or 300 {[values]}',
                 '300 = conflicting versions; client resolves'],
                ['DELETE', '/kv/{key}',  '{ context }',        '200 OK',
                 'Writes a tombstone; physically removed later'],
            ],
            col_widths=[W * 0.07, W * 0.17, W * 0.19, W * 0.24, W * 0.33],
        ),
        p('Table 2 - External REST API surface', CAPTION),
        sp(6),
        h2('3.2  Internal Node RPC'),
    ]
    story += bullet_list([
        '<b>Coordinate(key, value, context, W)</b> - coordinator initiates write',
        '<b>Replicate(key, value, vc)</b> - peer-to-peer replication',
        '<b>ForwardRead(key, R)</b> - coordinator reads from replicas',
        '<b>Gossip(memberList, hintedKeys)</b> - membership & failure state',
        '<b>AntiEntropy(token_range)</b> - exchange Merkle hashes for a ring segment',
    ])
    story += [sp(10)]

    # 4. DATA MODEL
    story += [
        h1('4.  Data Model'),
        rule(C_BLUE),
        sp(6),
        h2('4.1  Key-Value Record'),
        p('''Each stored record is more than a raw key/value pair - it carries the metadata
        needed for conflict detection and replication:'''),
        sp(4),
        p('Record schema (internal):', H3),
        p('''key:           bytes (<=256 B)
value:         bytes (<=1 MB)
vector_clock:  map[node_id -> logical_clock]   # e.g. {A:3, B:1}
timestamp:     int64 (Unix microseconds)
checksum:      uint32 (CRC32C)
ttl:           int64 (optional expiry)
is_tombstone:  bool''', CODE),
        sp(6),
        h2('4.2  Ring Metadata (per node)'),
        make_table(
            ['Field', 'Type', 'Description'],
            [
                ['ring_members',     'map[node_id -> endpoint]', 'All live/suspect nodes'],
                ['token_assignments','map[token -> node_id]',    'Which node owns which ring position'],
                ['node_status',      'enum {UP, SUSPECT, DOWN}', 'Local view of each peer'],
                ['hint_store',       'map[node_id -> [record]]', 'Buffered writes for downed nodes'],
            ],
            col_widths=[W * 0.25, W * 0.30, W * 0.45],
        ),
        p('Table 3 - Ring metadata schema', CAPTION),
        sp(10),
    ]

    # 5. HIGH-LEVEL ARCHITECTURE
    story += [
        h1('5.  High-Level Architecture'),
        rule(C_BLUE),
        sp(6),
        p('''Every node in the cluster is a <b>peer</b> - there is no master. A client can
        contact <i>any</i> node; that node becomes the <b>coordinator</b> for the request.
        The coordinator uses consistent hashing to locate the responsible replicas and fans
        out reads/writes. Nodes continuously gossip membership and failure information so every
        node has an eventually-consistent view of the ring.'''),
        sp(10),
        draw_architecture(540, 310),
        p('Figure 2 - Four-tier view: clients -> routing -> KV nodes -> storage engines',
          CAPTION),
        sp(6),
    ]
    story += bullet_list([
        '<b>Client Tier</b> - any SDK (Java, Python, Go) that speaks HTTP/gRPC',
        '<b>Routing Layer</b> - stateless load balancer; any node can route any request',
        '<b>KV Node Layer</b> - stateful peers; each owns a set of ring tokens + an LSM store',
        '<b>Storage Layer</b> - per-node LSM-tree engine (RocksDB / custom LSM)',
    ])
    story += [sp(10), pb()]

    # 6. CONSISTENT HASHING & PARTITIONING
    story += [
        h1('6.  Partitioning - Consistent Hashing'),
        rule(C_BLUE),
        sp(6),
        p('''Naive modulo hashing breaks catastrophically when you add or remove a node -
        almost every key remaps. <b>Consistent hashing</b> arranges both keys and nodes on a
        circular ring (0 to 2^32, wrap-around). A key is owned by the first node clockwise
        from its hash position. When a node joins or leaves, only the keys adjacent to that
        position move - typically just K/N keys out of K total.'''),
        sp(6),
        p('''<b>Virtual nodes (vnodes)</b> give each physical node 150+ positions on the ring,
        so the token space is divided far more evenly. This also means a failed node\'s load
        is spread across <i>all</i> surviving nodes rather than dumped on one neighbour.'''),
        sp(10),
        draw_hash_ring(540, 350),
        p('Figure 3 - Consistent hash ring: 5 physical nodes, each with virtual-node tokens',
          CAPTION),
        sp(8),
        make_table(
            ['Approach', 'Node join/leave', 'Load balance', 'Hotspot risk'],
            [
                ['Modulo hashing',    'Remaps ~all keys',    'Perfect',  'Medium'],
                ['Consistent hashing','Remaps ~K/N keys',   'Uneven',   'High (without vnodes)'],
                ['Consistent + vnodes','Remaps ~K/N keys',  'Very even','Low'],
            ],
            col_widths=[W * 0.28, W * 0.25, W * 0.22, W * 0.25],
        ),
        p('Table 4 - Partitioning scheme comparison', CAPTION),
        sp(10),
    ]

    # 7. REPLICATION
    story += [
        h1('7.  Replication'),
        rule(C_BLUE),
        sp(6),
        p('''After the coordinator places a key on the ring, it copies the record to the
        next <b>N-1</b> clockwise successor nodes. With N=3, every key has three identical
        copies on three different physical machines. Replicas are chosen from distinct
        failure domains (racks / AZs) where possible.'''),
        sp(10),
        draw_replication(540, 240),
        p('Figure 4 - A key is placed on N1 (coordinator) and replicated to N2, N3',
          CAPTION),
        sp(6),
        make_table(
            ['N', 'W', 'R', 'Consistency', 'Availability', 'Use case'],
            [
                ['3', '1', '1', 'Eventual', 'Highest',  'Async best-effort (metrics, logs)'],
                ['3', '2', '2', 'Strong*',  'High',     'General purpose (shopping cart)'],
                ['3', '3', '1', 'Strong',   'Lower',    'Metadata, leader election'],
                ['3', '1', '3', 'Strong',   'Lower',    'Read-heavy, stale writes OK'],
            ],
            col_widths=[W*0.05, W*0.05, W*0.05, W*0.16, W*0.16, W*0.53],
        ),
        p('Table 5 - Common N/W/R trade-off profiles  (* eventual strong; no linearisability)',
          CAPTION),
        sp(10),
    ]

    # 8. CONSISTENCY
    story += [
        h1('8.  Consistency - Quorum & Vector Clocks'),
        rule(C_BLUE),
        sp(6),
        h2('8.1  Quorum Writes & Reads'),
        p('''The rule <b>W + R &gt; N</b> guarantees that a read quorum and a write quorum
        always share at least one node, so the most recent write is always included in any
        read response. With N=3, W=2, R=2: any two-node read set must overlap with any
        two-node write set (pigeonhole principle).'''),
        sp(8),
        draw_write_path(540, 300),
        p('Figure 5 - Write path: coordinator fans out to N replicas, waits for W ACKs',
          CAPTION),
        sp(8),
        draw_read_path(540, 260),
        p('Figure 6 - Read path: coordinator reads R replicas, repairs stale copies',
          CAPTION),
        sp(8),
        h2('8.2  Vector Clocks'),
        p('''Lamport clocks assign a single integer but cannot distinguish "A happened before B"
        from "A and B happened concurrently". <b>Vector clocks</b> track one counter per node.
        If vc(A) is dominated by vc(B) element-wise, A happened before B. Otherwise the writes
        are <i>concurrent</i> - a conflict the application must resolve (or LWW wins).'''),
        sp(8),
        draw_vector_clock(540, 200),
        p('Figure 7 - Two concurrent writes produce a conflict; application or LWW resolves',
          CAPTION),
        sp(8),
        CalloutBox(
            '<b>Last-Write-Wins (LWW)</b> uses the wall-clock timestamp to break ties. '
            'Simple but risky: clock skew can silently discard valid writes. '
            'Prefer LWW only for immutable or append-only workloads.',
            bg=HexColor('#FFF8E1'), border=C_AMBER
        ),
        sp(10), pb(),
    ]

    # 9. FAILURE HANDLING
    story += [
        h1('9.  Failure Handling'),
        rule(C_BLUE),
        sp(6),
        h2('9.1  Hinted Handoff'),
        p('''If a replica node (say Node C) is unreachable during a write, the coordinator
        instead stores the write on a <i>different</i> live node along with a <b>hint</b>
        indicating the intended destination. When Node C recovers and gossip marks it UP,
        the hint-holder delivers the buffered writes. This preserves write availability
        without violating the quorum contract.'''),
        sp(6),
        CalloutBox(
            '<b>Trade-off:</b> Hinted handoff is bounded - hints are dropped after a TTL '
            '(default 1 hour). If a node is down longer, anti-entropy must catch it up.',
            bg=HexColor('#E3F2FD'), border=C_BLUE
        ),
        sp(10),
        h2('9.2  Merkle Tree Anti-Entropy'),
        p('''Even with quorums and hinted handoff, replicas can drift (network hiccups, bugs).
        Each node continuously builds a <b>Merkle tree</b> over its key range - leaf nodes
        are hashes of individual keys, and internal nodes hash their children. Two replicas
        exchange root hashes; a mismatch means they diverge <i>somewhere</i>. Binary search
        down the tree locates the diverging leaf without transferring all data.'''),
        sp(8),
        draw_merkle_tree(540, 270),
        p('Figure 8 - Merkle tree comparison: mismatch at root -> drill down to individual key',
          CAPTION),
        sp(6),
        p('''Complexity: O(log N) messages to locate a single diverging key vs O(N) for a
        full scan. Cassandra uses this for its "repair" command.'''),
        sp(10),
        h2('9.3  Gossip Protocol'),
        p('''Every second, each node picks two random peers and exchanges:
        (a) its membership list, (b) heartbeat counters, (c) pending hint metadata.
        If a node\'s heartbeat stops incrementing, peers mark it SUSPECT; after a timeout
        they mark it DOWN. False positives are handled by the Phi Accrual Failure Detector,
        which adapts its threshold to network conditions.'''),
        sp(8),
        draw_gossip(540, 240),
        p('Figure 9 - Gossip fan-out: each node spreads state to random peers each second',
          CAPTION),
        sp(6),
        make_table(
            ['Mechanism', 'Handles', 'Bounded by', 'Real example'],
            [
                ['Hinted Handoff',  'Short-lived node failures during writes',
                 'Hint TTL (1 hr)',        'Dynamo, Cassandra'],
                ['Merkle Anti-Entropy', 'Long-term replica drift, silent corruption',
                 'Background repair schedule', 'Cassandra "nodetool repair"'],
                ['Gossip',          'Membership, failure detection',
                 'Network bandwidth (low)', 'Cassandra, Riak, Serf'],
                ['Read Repair',     'Stale reads caught in-flight',
                 'Read path overhead',     'Dynamo, Voldemort'],
            ],
            col_widths=[W * 0.22, W * 0.30, W * 0.24, W * 0.24],
        ),
        p('Table 6 - Failure-handling mechanisms summary', CAPTION),
        sp(10), pb(),
    ]

    # 10. STORAGE ENGINE
    story += [
        h1('10. Storage Engine'),
        rule(C_BLUE),
        sp(6),
        p('''Each node uses an <b>LSM-tree (Log-Structured Merge-tree)</b> for local storage.
        Writes go to an in-memory buffer (MemTable) and a Write-Ahead Log (WAL). When the
        MemTable fills, it is flushed to an immutable SSTable on disk. Compaction merges and
        garbage-collects SSTables in the background.'''),
        sp(6),
        make_table(
            ['Component', 'Purpose', 'Typical size', 'Notes'],
            [
                ['WAL (Write-Ahead Log)', 'Crash recovery', '64-256 MB per segment',
                 'Append-only; rotated on flush'],
                ['MemTable', 'In-memory sorted buffer', '64-512 MB',
                 'Red-black tree or skip list'],
                ['SSTable', 'Immutable sorted file', '32 MB - 2 GB',
                 'Bloom filter per file; block index'],
                ['Bloom Filter', 'Avoid unnecessary disk reads', '~10 bits/key',
                 '1% false positive rate typical'],
                ['Block Cache', 'Hot block caching', '512 MB - 64 GB RAM',
                 'LRU; reduces SSTable I/O by 90%+'],
                ['Compaction', 'Merge SSTables, purge tombstones', 'Background I/O',
                 'Leveled or STCS strategy'],
            ],
            col_widths=[W * 0.24, W * 0.26, W * 0.22, W * 0.28],
        ),
        p('Table 7 - LSM-tree storage engine components', CAPTION),
        sp(6),
        CalloutBox(
            '<b>Why LSM over B-Tree?</b>  B-Trees update in place (random I/O). '
            'LSM writes are always sequential (append to log, flush sorted file). '
            'On spinning disks this is 10-100x faster. On NVMe SSDs the gap narrows '
            'but LSM still wins on write-heavy workloads common in KV stores.',
            bg=HexColor('#F3E5F5'), border=HexColor('#7B1FA2')
        ),
        sp(10),
    ]

    # 11. NETWORK PARTITION HANDLING
    story += [
        h1('11. Network Partitions & the CAP Trade-off'),
        rule(C_BLUE),
        sp(6),
        p('''A Dynamo-style KV store is <b>AP</b>: during a partition it chooses availability
        over consistency. Writes succeed as long as W nodes are reachable; reads succeed as long
        as R nodes are reachable. Nodes on the minority side of a partition accept writes, which
        diverge from the majority side until the partition heals and anti-entropy reconciles them.'''),
        sp(6),
        p('''<b>Scenario: 2-of-5 nodes isolated</b><br/>
        Nodes D and E are partitioned from A, B, C. With W=2, R=2:'''),
        sp(4),
    ]
    story += bullet_list([
        'Majority partition (A, B, C): can serve reads AND writes (quorum met)',
        'Minority partition (D, E): <i>cannot</i> form quorum - writes are <b>rejected</b>',
        'Reads from (D, E): may return stale data until partition heals',
        'On healing: gossip detects reconnection; Merkle anti-entropy syncs diverged keys',
    ])
    story += [
        sp(6),
        make_table(
            ['Consistency Model', 'Guarantee', 'Allows during partition', 'Example'],
            [
                ['Linearisability (CP)', 'Reads see latest write',
                 'May reject writes', 'Zookeeper, etcd'],
                ['Sequential consistency', 'All ops in program order',
                 'Limited', 'Redis (single node)'],
                ['Eventual consistency (AP)', 'Replicas converge eventually',
                 'Always writable', 'Dynamo, Cassandra'],
                ['Causal consistency', 'Causally related ops ordered',
                 'Yes (CRDT-based)', 'Riak (CRDT mode)'],
            ],
            col_widths=[W * 0.25, W * 0.28, W * 0.22, W * 0.25],
        ),
        p('Table 8 - Consistency model spectrum', CAPTION),
        sp(10), pb(),
    ]

    # 12. SECONDARY INDEXES
    story += [
        h1('12. Secondary Indexes'),
        rule(C_BLUE),
        sp(6),
        p('''The primary access path is <code>get(key)</code>. Secondary indexes let clients
        query by attribute value (e.g., <i>"all users in city=London"</i>). Two common
        approaches:'''),
        sp(6),
        h2('12.1  Local (Scatter-Gather) Index'),
    ]
    story += bullet_list([
        'Each node indexes only the keys it owns.',
        'A query fans out to <i>all</i> nodes ("scatter"), each returns matches, coordinator merges ("gather").',
        'Simple to implement; no cross-node write coordination.',
        'Expensive reads: O(N) network hops per query.',
    ])
    story += [
        sp(4),
        h2('12.2  Global Index'),
    ]
    story += bullet_list([
        'A separate index partition maps <i>attribute to [primary keys]</i>.',
        'Index partitioned independently (e.g., by first letter of attribute value).',
        'Reads are O(1) - hit one index partition, then fetch primary keys.',
        'Writes are expensive: a single KV write may update index partitions on different nodes (distributed transaction or async update).',
    ])
    story += [
        sp(6),
        make_table(
            ['Index Type', 'Read complexity', 'Write complexity', 'Consistency', 'Examples'],
            [
                ['Local / scatter-gather', 'O(N) fanout',     'O(1) local',  'Strong (local)',     'Cassandra secondary idx'],
                ['Global partitioned',     'O(1) + fetch',   'O(index nodes)','Eventual (async)', 'DynamoDB GSI'],
                ['External search engine','O(1) search',     'Dual write',  'Eventual',           'Elasticsearch + KV'],
            ],
            col_widths=[W * 0.24, W * 0.18, W * 0.18, W * 0.18, W * 0.22],
        ),
        p('Table 9 - Secondary index approaches', CAPTION),
        sp(10),
    ]

    # 13. INTERVIEW Q&A
    story += [
        h1('13. Interview Deep-Dives'),
        rule(C_BLUE),
        sp(6),
    ]

    qas = [
        (
            'Q1 - Walk through a complete write path in a Dynamo-style KV store.',
            [
                '1. Client issues PUT /kv/user:42 with value + context to any node.',
                '2. <b>Coordinator</b> hashes the key (MD5/SHA1 mod 2^32) -> position on ring.',
                '3. Coordinator identifies the <b>N=3 successor nodes</b> clockwise from that position.',
                '4. Coordinator adds/increments its entry in the <b>vector clock</b>.',
                '5. Coordinator sends Replicate(key, value, vc) to all N replicas in parallel.',
                '6. Each replica persists to WAL + MemTable, returns ACK.',
                '7. Coordinator waits for <b>W=2 ACKs</b>; returns 200 OK to client.',
                '8. Remaining replica(s) may ACK later (async); if a replica is DOWN, hints are stored.',
                '9. Client receives new vector-clock context for future conditional writes.',
            ]
        ),
        (
            'Q2 - How does the system handle a network partition isolating 2 of 5 nodes?',
            [
                '<b>Gossip detects</b> the partition within a few heartbeat cycles (seconds).',
                '<b>Majority partition (A,B,C)</b>: can form W=2 and R=2 quorums - fully operational.',
                '<b>Minority partition (D,E)</b>: only 2 nodes; writes to keys whose replicas span '
                'both partitions may fail quorum.',
                '<b>On healing</b>: gossip exchanges ring state; Merkle anti-entropy syncs diverged '
                'ranges; hinted handoffs are delivered.',
                '<b>Design choice</b>: to remain writable during partition, the system may allow sloppy '
                'quorum (write to any W live nodes, not necessarily the "true" N).',
            ]
        ),
        (
            'Q3 - Explain how Merkle trees enable efficient anti-entropy.',
            [
                'Each node builds a Merkle tree over its key range: leaf = hash(key+value), '
                'parent = hash(children).',
                'Two replicas exchange only their <b>root hashes</b> (8 bytes each).',
                'If roots match -> fully in sync, done.',
                'If mismatch -> exchange child hashes; recursively descend until '
                'the diverging leaf (= diverging key) is found.',
                'Complexity: <b>O(log K)</b> hash exchanges to locate 1 diverging key vs '
                'O(K) for a full key scan.',
                'After locating the key, the node with the higher vector clock pushes the '
                'correct value to the stale replica.',
            ]
        ),
        (
            'Q4 - What is hinted handoff and why is it important?',
            [
                'If replica Node C is unreachable, the coordinator stores the write on a '
                '<b>surrogate node</b> (e.g., Node D) with a "hint" tag: '
                '"this write belongs to C".',
                '<b>Why it matters</b>: without hinted handoff, a write to a key whose '
                'primary replica is down would fail the quorum check and return an error '
                'to the client - breaking the "always writable" guarantee.',
                'When C comes back up, Node D detects it via gossip and delivers all '
                'buffered hints; C is now caught up.',
                '<b>Limits</b>: hints are bounded by TTL and disk space. Long outages '
                'must be healed by full anti-entropy (Merkle) repair.',
            ]
        ),
        (
            'Q5 - How would you add secondary indexes?',
            [
                '<b>Step 1 - Choose strategy</b>: local index for simplicity; global index '
                'for low-latency reads.',
                '<b>Step 2 - Dual-write</b>: on every PUT, write to both the '
                'primary KV store <i>and</i> the index partition. Accept eventual consistency '
                'between them (or use a saga/outbox for stronger guarantees).',
                '<b>Step 3 - Index schema</b>: index key = '
                '"{attribute}:{value}:{primary_key}" -> empty value. This keeps '
                'the index in the same KV store with no new infrastructure.',
                '<b>Step 4 - Query</b>: '
                'GET /kv/city:London:* via range scan on index partition, '
                'then bulk-fetch primary keys.',
                '<b>Trade-off</b>: index writes become a bottleneck for high-cardinality '
                'attributes; consider async index updates with a change-data-capture stream.',
            ]
        ),
    ]

    for q, answers in qas:
        story.append(KeepTogether([
            p(q, H3),
        ] + [p(f'&#8226;  {a}', ANSWER) for a in answers] + [sp(8)]))

    story += [sp(6), pb()]

    # 14. REAL-WORLD EXAMPLES
    story += [
        h1('14. Real-World Examples'),
        rule(C_BLUE),
        sp(6),
        make_table(
            ['System', 'Org', 'Partitioning', 'Consistency', 'Storage', 'Notable Feature'],
            [
                ['Dynamo',      'Amazon',    'Consistent hash + vnodes', 'AP / quorum',
                 'BDB / custom', 'Original paper; pioneered hinted handoff + vector clocks'],
                ['DynamoDB',    'AWS',       'Hash + range partitions', 'AP (default) / strong opt-in',
                 'Proprietary LSM', 'Managed; global tables; DAX cache'],
                ['Cassandra',   'Apache/Meta','Consistent hash + vnodes', 'AP / tunable',
                 'Custom LSM',  'Wide rows; CQL; multi-DC replication'],
                ['Riak KV',     'Basho',     'Consistent hash',         'AP / CRDTs',
                 'Bitcask/LevelDB','CRDT-native conflict resolution'],
                ['Voldemort',   'LinkedIn',  'Consistent hash',         'AP / eventual',
                 'BDB/MySQL',   'Used for members + job recommendations'],
                ['Redis Cluster','Redis Ltd', 'Hash slots (16384)',      'CP (primary)',
                 'In-memory + AOF', 'Sub-ms latency; rich data structures'],
                ['TiKV',        'PingCAP',   'Range partitions',        'CP / Raft',
                 'RocksDB',     'Transactional; used by TiDB; Raft consensus'],
            ],
            col_widths=[W*0.13, W*0.11, W*0.20, W*0.15, W*0.15, W*0.26],
        ),
        p('Table 10 - Distributed KV stores in production', CAPTION),
        sp(10),
    ]

    # 15. CAPACITY PLANNING
    story += [
        h1('15. Capacity Planning'),
        rule(C_BLUE),
        sp(6),
        p('For a target of <b>1 M QPS</b> with 60/40 read/write split and key size 64 B, '
          'value size 256 B average:'),
        sp(6),
        make_table(
            ['Metric', 'Calculation', 'Result'],
            [
                ['Write throughput', '400K writes/s x 320 B/write', '~128 MB/s incoming'],
                ['Replication overhead', 'x N=3 replicas',           '~384 MB/s total network write'],
                ['Read throughput', '600K reads/s x 320 B/read',     '~192 MB/s outgoing'],
                ['Storage per year (1 KB avg)', '400K w/s x 3 replicas x 1 KB x 31.5M s',
                 '~37 TB raw / year'],
                ['Node count (10 Gbps NIC, 50% util)', '384 MB/s / (625 MB/s x 0.5)',
                 '~2 nodes for write bandwidth'],
                ['Node count (disk I/O, 500 MB/s SSD)', '384 MB/s / 500 MB/s',
                 '~1 node; typically CPU/memory bound'],
                ['Recommended cluster size', 'Storage + redundancy + headroom',
                 '8-12 nodes (practical)'],
            ],
            col_widths=[W * 0.32, W * 0.38, W * 0.30],
        ),
        p('Table 11 - Back-of-envelope capacity for 1M QPS workload', CAPTION),
        sp(10), pb(),
    ]

    # 16. REFERENCES
    story += [
        h1('16. References & Further Reading'),
        rule(C_BLUE),
        sp(6),
        h2('Foundational Papers'),
    ]
    papers = [
        ("Amazon Dynamo: Amazon's Highly Available Key-value Store",
         'DeCandia et al., SOSP 2007',
         'dl.acm.org/doi/10.1145/1294261.1294281'),
        ('Bigtable: A Distributed Storage System for Structured Data',
         'Chang et al., OSDI 2006',
         'static.googleusercontent.com/media/research.google.com/en//archive/bigtable-osdi06.pdf'),
        ('Cassandra: A Decentralized Structured Storage System',
         'Lakshman & Malik, LADIS 2009',
         'dl.acm.org/doi/10.1145/1773912.1773922'),
        ('Chord: A Scalable Peer-to-peer Lookup Service for Internet Applications',
         'Stoica et al., SIGCOMM 2001',
         'dl.acm.org/doi/10.1145/383059.383071'),
    ]
    for title, author, url in papers:
        story.append(p(f'<b>{title}</b><br/>{author}<br/>'
                       f'<font color="#1565C0"><i>{url}</i></font>', BULLET))
        story.append(sp(4))

    story += [sp(6), h2('Blog Posts & Tutorials')]
    blogs = [
        ('Designing Data-Intensive Applications - Martin Kleppmann (book)',
         'dataintensive.net'),
        ('System Design Primer - Donne Martin',
         'github.com/donnemartin/system-design-primer'),
        ('How Cassandra Stores Data - DataStax blog',
         'datastax.com/blog/storage-engine-internals'),
        ('Consistent Hashing and Random Trees - Karger et al.',
         'cs.princeton.edu/courses/archive/fall09/cos518/papers/chash.pdf'),
        ('DynamoDB Under the Hood - AWS re:Invent 2018 (video)',
         'youtube.com/watch?v=yvBR71D0nAQ'),
        ('LSM-Trees and RocksDB Internals - Facebook Engineering',
         'rocksdb.org/blog/2021/05/26/rocksdb-intro.html'),
    ]
    for title, url in blogs:
        story.append(p(f'<b>{title}</b><br/>'
                       f'<font color="#1565C0"><i>{url}</i></font>', BULLET))
        story.append(sp(3))

    story += [
        sp(6),
        h2('YouTube Videos'),
    ]
    videos = [
        ("Dynamo: Amazon's Highly Available Key-Value Store - MIT 6.824",
         'youtube.com/watch?v=Q3YZTk8fli8'),
        ('Consistent Hashing - System Design Interview (Gaurav Sen)',
         'youtube.com/watch?v=K0Ta65OqQkY'),
        ('Design a Key-Value Store - ByteByteGo',
         'youtube.com/watch?v=rnZmdmlR-2M'),
        ('How Cassandra Works - DataStax Academy',
         'youtube.com/watch?v=d7o6a75sfY0'),
    ]
    for title, url in videos:
        story.append(p(f'<b>{title}</b><br/>'
                       f'<font color="#1565C0"><i>{url}</i></font>', BULLET))
        story.append(sp(3))

    # Footer note
    story += [
        sp(20),
        rule(C_GREY),
        sp(4),
        p('Generated by System Design Deep Dive  |  Distributed Key-Value Store  |  2025',
          S('foot', fontName='Helvetica-Oblique', fontSize=8, textColor=C_GREY,
            alignment=TA_CENTER)),
    ]

    # Build
    doc.build(story)
    print(f'PDF written to {path}')


if __name__ == '__main__':
    build_pdf('/Users/naji/WORK/github.com/AI/claude/Agent/kv_store_design.pdf')
