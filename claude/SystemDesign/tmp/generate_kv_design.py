#!/usr/bin/env python3
"""
Distributed Key-Value Store System Design Document Generator
Generates a comprehensive PDF using reportlab
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.platypus.flowables import Flowable
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
import os

# Color palette
DARK_BLUE = HexColor('#1a365d')
ACCENT_BLUE = HexColor('#2b6cb0')
LIGHT_BLUE = HexColor('#ebf8ff')
MEDIUM_BLUE = HexColor('#bee3f8')
DARK_GRAY = HexColor('#2d3748')
MEDIUM_GRAY = HexColor('#718096')
LIGHT_GRAY = HexColor('#f7fafc')
TABLE_HEADER_BG = HexColor('#2b6cb0')
TABLE_ALT_ROW = HexColor('#f0f4f8')
WHITE = colors.white
BLACK = colors.black
GREEN = HexColor('#276749')
RED = HexColor('#9b2c2c')

OUTPUT_PATH = '/Users/naji/WORK/github.com/AI/cluade/Agent/distributed_kv_store_design.pdf'


class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []
        self._page_count = 0

    def showPage(self):
        self._page_count += 1
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for i, state in enumerate(self._saved_page_states):
            self.__dict__.update(state)
            page_num = i + 1
            if page_num > 2:  # Skip title and TOC pages
                self.setFont('Helvetica', 9)
                self.setFillColor(MEDIUM_GRAY)
                self.drawCentredString(
                    letter[0] / 2, 0.4 * inch,
                    f"Distributed Key-Value Store: System Design  \u2014  Page {page_num} of {num_pages}"
                )
                # Footer line
                self.setStrokeColor(MEDIUM_BLUE)
                self.setLineWidth(0.5)
                self.line(0.75 * inch, 0.55 * inch, letter[0] - 0.75 * inch, 0.55 * inch)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)


def build_styles():
    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        name='DocTitle',
        fontSize=32,
        textColor=WHITE,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold',
        leading=40,
        spaceAfter=12,
    ))
    styles.add(ParagraphStyle(
        name='DocSubtitle',
        fontSize=18,
        textColor=MEDIUM_BLUE,
        alignment=TA_CENTER,
        fontName='Helvetica',
        leading=24,
        spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        name='DocMeta',
        fontSize=11,
        textColor=LIGHT_BLUE,
        alignment=TA_CENTER,
        fontName='Helvetica',
        leading=16,
    ))
    styles.add(ParagraphStyle(
        name='SectionHeader',
        fontSize=16,
        textColor=DARK_BLUE,
        fontName='Helvetica-Bold',
        spaceBefore=18,
        spaceAfter=8,
        leading=20,
        borderPad=4,
    ))
    styles.add(ParagraphStyle(
        name='SubHeader',
        fontSize=13,
        textColor=ACCENT_BLUE,
        fontName='Helvetica-Bold',
        spaceBefore=12,
        spaceAfter=6,
        leading=16,
    ))
    styles.add(ParagraphStyle(
        name='SubSubHeader',
        fontSize=11,
        textColor=DARK_GRAY,
        fontName='Helvetica-Bold',
        spaceBefore=8,
        spaceAfter=4,
        leading=14,
    ))
    styles['BodyText'].fontSize = 10
    styles['BodyText'].textColor = DARK_GRAY
    styles['BodyText'].fontName = 'Helvetica'
    styles['BodyText'].leading = 15
    styles['BodyText'].spaceAfter = 6
    styles['BodyText'].alignment = TA_JUSTIFY
    styles.add(ParagraphStyle(
        name='BulletItem',
        fontSize=10,
        textColor=DARK_GRAY,
        fontName='Helvetica',
        leading=14,
        spaceAfter=3,
        leftIndent=16,
        bulletIndent=4,
    ))
    styles.add(ParagraphStyle(
        name='NumberedItem',
        fontSize=10,
        textColor=DARK_GRAY,
        fontName='Helvetica',
        leading=14,
        spaceAfter=4,
        leftIndent=20,
    ))
    styles.add(ParagraphStyle(
        name='Monospace',
        fontSize=8.5,
        textColor=DARK_GRAY,
        fontName='Courier',
        leading=12,
        leftIndent=0,
        spaceAfter=2,
    ))
    styles.add(ParagraphStyle(
        name='MonospaceBox',
        fontSize=8,
        textColor=HexColor('#1a202c'),
        fontName='Courier',
        leading=11,
        leftIndent=8,
        rightIndent=8,
        spaceBefore=2,
        spaceAfter=2,
        backColor=HexColor('#edf2f7'),
    ))
    styles.add(ParagraphStyle(
        name='TOCEntry',
        fontSize=10,
        textColor=DARK_GRAY,
        fontName='Helvetica',
        leading=16,
        leftIndent=0,
    ))
    styles.add(ParagraphStyle(
        name='TOCEntryIndent',
        fontSize=9,
        textColor=MEDIUM_GRAY,
        fontName='Helvetica',
        leading=14,
        leftIndent=16,
    ))
    styles.add(ParagraphStyle(
        name='Caption',
        fontSize=9,
        textColor=MEDIUM_GRAY,
        fontName='Helvetica-Oblique',
        alignment=TA_CENTER,
        spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        name='Note',
        fontSize=9,
        textColor=HexColor('#744210'),
        fontName='Helvetica-Oblique',
        leading=13,
        leftIndent=8,
        backColor=HexColor('#fefcbf'),
        spaceAfter=6,
    ))
    return styles


def hr(width=1.0):
    return HRFlowable(
        width=f"{int(width*100)}%",
        thickness=1,
        color=MEDIUM_BLUE,
        spaceAfter=8,
        spaceBefore=4,
    )


def section_hr():
    return HRFlowable(
        width="100%",
        thickness=2,
        color=DARK_BLUE,
        spaceAfter=10,
        spaceBefore=6,
    )


_TBL_HDR_STYLE = ParagraphStyle(
    'TblHdr',
    fontSize=9, fontName='Helvetica-Bold', leading=12,
    textColor=WHITE, alignment=TA_CENTER,
)
_TBL_BODY_STYLE = ParagraphStyle(
    'TblBody',
    fontSize=9, fontName='Helvetica', leading=13,
    textColor=DARK_GRAY,
)


def _cell(value, is_header=False):
    """Wrap a cell value in a Paragraph for proper word-wrap and HTML safety."""
    if not isinstance(value, str):
        return value
    # Escape XML special chars so Paragraph parser doesn't choke on JSON/math content
    safe = value.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return Paragraph(safe, _TBL_HDR_STYLE if is_header else _TBL_BODY_STYLE)


def make_table(data, col_widths, header_bg=TABLE_HEADER_BG, alternate=True):
    para_data = [
        [_cell(c, row_idx == 0) for c in row]
        for row_idx, row in enumerate(data)
    ]
    style = [
        ('BACKGROUND', (0, 0), (-1, 0), header_bg),
        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
        ('VALIGN', (0, 1), (-1, -1), 'TOP'),
        ('ROWBACKGROUND', (0, 1), (-1, -1), [WHITE, TABLE_ALT_ROW] if alternate else [WHITE]),
        ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#cbd5e0')),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]
    t = Table(para_data, colWidths=col_widths)
    t.setStyle(TableStyle(style))
    return t


def ascii_box(lines, styles_obj):
    """Render ASCII art inside a styled box."""
    items = []
    box_style = ParagraphStyle(
        name='AsciiBox',
        fontSize=8,
        fontName='Courier',
        leading=11,
        leftIndent=4,
        rightIndent=4,
        textColor=HexColor('#1a202c'),
        backColor=HexColor('#edf2f7'),
        spaceAfter=2,
        spaceBefore=2,
        borderColor=ACCENT_BLUE,
        borderWidth=1,
        borderPad=6,
    )
    all_lines = '<br/>'.join(
        line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace(' ', '&nbsp;')
        for line in lines
    )
    return Paragraph(all_lines, box_style)


def build_document():
    styles = build_styles()
    story = []

    # =========================================================================
    # TITLE PAGE
    # =========================================================================
    # Blue background title block using a table
    title_data = [[
        Paragraph("Distributed Key-Value Store", styles['DocTitle']),
    ]]
    subtitle_data = [[
        Paragraph("System Design Document", styles['DocTitle']),
    ]]
    title_table = Table(
        [
            [Paragraph("Distributed Key-Value Store", ParagraphStyle(
                'T1', fontSize=28, textColor=WHITE, fontName='Helvetica-Bold',
                alignment=TA_CENTER, leading=34
            ))],
            [Paragraph("System Design Document", ParagraphStyle(
                'T2', fontSize=22, textColor=WHITE, fontName='Helvetica-Bold',
                alignment=TA_CENTER, leading=28, spaceAfter=6
            ))],
            [Paragraph("Dynamo-Style Architecture", ParagraphStyle(
                'T3', fontSize=16, textColor=MEDIUM_BLUE, fontName='Helvetica',
                alignment=TA_CENTER, leading=22
            ))],
        ],
        colWidths=[6.5 * inch],
    )
    title_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), DARK_BLUE),
        ('TOPPADDING', (0, 0), (-1, -1), 16),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 16),
        ('LEFTPADDING', (0, 0), (-1, -1), 24),
        ('RIGHTPADDING', (0, 0), (-1, -1), 24),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ROUNDEDCORNERS', [8, 8, 8, 8]),
    ]))

    story.append(Spacer(1, 1.2 * inch))
    story.append(title_table)
    story.append(Spacer(1, 0.5 * inch))

    meta_table = Table([
        [Paragraph("Version 1.0", styles['DocMeta'])],
        [Paragraph("March 2026", styles['DocMeta'])],
        [Paragraph("System Design Reference", styles['DocMeta'])],
        [Spacer(1, 0.1 * inch)],
        [Paragraph("Topics: Consistent Hashing · Vector Clocks · Quorum Consensus", styles['DocMeta'])],
        [Paragraph("Merkle Trees · Gossip Protocol · LSM-Tree Storage", styles['DocMeta'])],
    ], colWidths=[6.5 * inch])
    meta_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), ACCENT_BLUE),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ROUNDEDCORNERS', [6, 6, 6, 6]),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 0.6 * inch))

    # Key stats boxes
    stats = [
        ("10M+\nQPS", "Target\nThroughput"),
        ("<10ms\np99", "Read/Write\nLatency"),
        ("99.9%\nUptime", "Availability\nSLA"),
        ("Petabyte\nScale", "Data\nCapacity"),
    ]
    stats_data = [[
        Paragraph(f'<b><font size="14" color="#2b6cb0">{s[0]}</font></b><br/><font size="8" color="#718096">{s[1]}</font>', ParagraphStyle('SC', alignment=TA_CENTER, leading=16))
        for s in stats
    ]]
    stats_table = Table(stats_data, colWidths=[1.6 * inch] * 4)
    stats_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), LIGHT_GRAY),
        ('GRID', (0, 0), (-1, -1), 1, MEDIUM_BLUE),
        ('TOPPADDING', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(stats_table)
    story.append(PageBreak())

    # =========================================================================
    # TABLE OF CONTENTS
    # =========================================================================
    story.append(Paragraph("Table of Contents", styles['SectionHeader']))
    story.append(section_hr())
    story.append(Spacer(1, 0.1 * inch))

    toc_entries = [
        ("1.", "Requirements & Scope", False),
        ("", "Functional Requirements", True),
        ("", "Non-Functional Requirements", True),
        ("", "Scale Assumptions", True),
        ("2.", "API Design", False),
        ("", "REST API Endpoints", True),
        ("", "Client SDK Interface", True),
        ("", "Request / Response Schemas", True),
        ("3.", "High-Level Architecture", False),
        ("4.", "Data Model", False),
        ("5.", "Consistent Hashing & Partitioning", False),
        ("", "Virtual Nodes (vnodes)", True),
        ("", "Token Ring", True),
        ("6.", "Replication Strategy", False),
        ("", "Preference List Construction", True),
        ("", "Rack / AZ Awareness", True),
        ("7.", "Consistency Model", False),
        ("", "Quorum Reads & Writes", True),
        ("", "CAP Theorem Positioning", True),
        ("", "Vector Clocks & Conflict Resolution", True),
        ("8.", "Write Path", False),
        ("", "Step-by-Step Walkthrough", True),
        ("", "Hinted Handoff", True),
        ("", "Latency Budget", True),
        ("9.", "Read Path", False),
        ("", "Step-by-Step Walkthrough", True),
        ("", "Read Repair", True),
        ("10.", "Failure Handling", False),
        ("", "Hinted Handoff", True),
        ("", "Anti-Entropy with Merkle Trees", True),
        ("", "Gossip Protocol", True),
        ("11.", "Network Partition Handling", False),
        ("12.", "Storage Engine (LSM-Tree)", False),
        ("13.", "Secondary Indexes", False),
        ("14.", "Capacity Planning", False),
        ("15.", "Real-World Implementations", False),
        ("16.", "Interview Q&A", False),
        ("17.", "References", False),
    ]

    for num, title, indent in toc_entries:
        style = styles['TOCEntryIndent'] if indent else styles['TOCEntry']
        prefix = f"&nbsp;&nbsp;&nbsp;&nbsp;{title}" if indent else f"<b>{num}</b>&nbsp;&nbsp;{title}"
        story.append(Paragraph(prefix, style))

    story.append(PageBreak())

    # =========================================================================
    # SECTION 1: REQUIREMENTS & SCOPE
    # =========================================================================
    story.append(Paragraph("1. Requirements & Scope", styles['SectionHeader']))
    story.append(section_hr())

    # Origin story callout
    origin_table = Table([[
        Paragraph(
            '<b>The Origin Story</b><br/>'
            'In 2007, Amazon engineers published the Dynamo paper after building a key-value store '
            'that had to <i>always</i> be writable — even during hardware failures and network partitions. '
            'Their shopping cart service could never return "write failed" to a customer adding an item. '
            'The result was a system that chose availability over consistency, resolved conflicts later, '
            'and inspired a generation of distributed databases: Cassandra, Riak, Voldemort, and DynamoDB itself. '
            'This document designs that exact class of system.',
            ParagraphStyle('OriginText', fontSize=9.5, fontName='Helvetica', leading=14,
                           textColor=HexColor('#1a365d'))
        )
    ]], colWidths=[6.5 * inch])
    origin_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), HexColor('#ebf8ff')),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('BOX', (0, 0), (-1, -1), 1.5, HexColor('#2b6cb0')),
        ('ROUNDEDCORNERS', [6, 6, 6, 6]),
    ]))
    story.append(origin_table)
    story.append(Spacer(1, 0.15 * inch))

    story.append(Paragraph("1.1 Functional Requirements", styles['SubHeader']))
    story.append(Paragraph(
        "The system must expose a simple key-value interface that hides all distributed "
        "complexity from the caller. Three core operations are required:", styles['BodyText']
    ))

    func_req_data = [
        ["Operation", "Signature", "Description"],
        ["PUT", "put(key, value, context=None)", "Store or update a value. Optional context carries the vector clock from a prior GET for optimistic concurrency."],
        ["GET", "get(key) → (value, context)", "Retrieve the latest value. Returns value plus a context object for subsequent writes."],
        ["DELETE", "delete(key, context=None)", "Logically delete a key by writing a tombstone record. Physical removal happens during compaction."],
        ["BATCH PUT", "batch_put([(key, value), ...])", "Atomic within a shard; best-effort across shards. Reduces round-trip overhead."],
    ]
    story.append(make_table(func_req_data,
        [1.0*inch, 2.1*inch, 3.2*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph(
        "<b>Tunable Consistency:</b> Callers may override the cluster-wide default quorum "
        "by passing a ConsistencyLevel hint (ONE, QUORUM, ALL). This allows read-your-writes "
        "semantics for latency-sensitive paths while tolerating eventual consistency elsewhere.",
        styles['BodyText']
    ))

    story.append(Paragraph("1.2 Non-Functional Requirements", styles['SubHeader']))
    nfr_data = [
        ["Requirement", "Target", "Notes"],
        ["Throughput", "10M reads/sec, 1M writes/sec", "Aggregate across cluster; scales linearly with nodes"],
        ["Read Latency", "< 5ms p50, < 10ms p99", "Measured at client, includes network RTT"],
        ["Write Latency", "< 8ms p50, < 15ms p99", "Durable write (W=2 default)"],
        ["Availability", "99.9% (8.7 hrs downtime/year)", "AP system; prefers availability over consistency"],
        ["Durability", "99.999999999% (11 nines)", "Via replication + checksums + WAL"],
        ["Fault Tolerance", "Survive loss of N-1 replicas", "Automatic failover, no manual intervention"],
        ["Scalability", "Horizontal, linear scale-out", "Add nodes with automatic rebalancing"],
        ["Recovery Time", "< 30 seconds for node failure detection", "Via gossip + heartbeat"],
    ]
    story.append(make_table(nfr_data, [1.5*inch, 1.9*inch, 2.9*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph("1.3 Scale Assumptions", styles['SubHeader']))
    scale_data = [
        ["Parameter", "Value", "Rationale"],
        ["Key size", "1 – 256 bytes", "Short keys improve cache efficiency"],
        ["Value size", "1 byte – 1 MB", "Capped to avoid head-of-line blocking"],
        ["Dataset size", "Petabyte-scale", "~1B keys × 10KB avg value × 3 replicas"],
        ["Hot key distribution", "Zipfian (top 1% = 80% traffic)", "Caching layer required for hotspots"],
        ["Replication factor (N)", "3", "Industry standard; balances cost vs durability"],
        ["Write quorum (W)", "2 (default)", "Majority write; tunable per-request"],
        ["Read quorum (R)", "2 (default)", "Strong consistency when W+R > N"],
        ["Nodes per cluster", "20–200", "Starting at 20; each node 2TB NVMe SSD"],
    ]
    story.append(make_table(scale_data, [1.6*inch, 1.8*inch, 2.9*inch]))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 2: API DESIGN
    # =========================================================================
    story.append(Paragraph("2. API Design", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph("2.1 REST API Endpoints", styles['SubHeader']))
    story.append(Paragraph(
        "All endpoints are exposed over HTTPS. Authentication uses HMAC-SHA256 request signing "
        "(similar to AWS Signature v4). TLS 1.3 is mandatory.", styles['BodyText']
    ))

    api_data = [
        ["Method", "Endpoint", "Request Body", "Response", "Status Codes"],
        ["PUT", "/kv/{key}", '{"value": "<base64>", "ttl": 3600}', '{"version": "v1:node3:5"}', "200 OK, 400 Bad Request, 503 Unavailable"],
        ["GET", "/kv/{key}", "—", '{"value": "<base64>", "version": "..."}', "200 OK, 404 Not Found, 503 Unavailable"],
        ["DELETE", "/kv/{key}", '{"version": "v1:node3:5"}', '{"deleted": true}', "200 OK, 404 Not Found, 409 Conflict"],
        ["GET", "/kv/{key}/versions", "—", '{"versions": [...]}', "200 OK, 404 Not Found"],
        ["POST", "/kv/batch", '{"ops": [{"type":"put","key":"k","value":"v"}]}', '{"results": [...]}', "207 Multi-Status"],
    ]
    story.append(make_table(api_data, [0.5*inch, 1.2*inch, 1.4*inch, 1.5*inch, 1.7*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph("2.2 Client SDK Interface", styles['SubHeader']))
    story.append(Paragraph(
        "The SDK wraps the REST API with automatic retry, jitter, and vector clock management:", styles['BodyText']
    ))
    story.append(ascii_box([
        "# Python SDK — simplified interface",
        "",
        "class KVClient:",
        "    def __init__(self, endpoints: List[str], consistency: str = 'QUORUM'):",
        "        self.lb = LoadBalancer(endpoints)    # round-robin with health checks",
        "        self.consistency = consistency",
        "",
        "    def get(self, key: str) -> Tuple[bytes, VectorClock]:",
        "        resp = self._request('GET', f'/kv/{key}')",
        "        return base64.decode(resp['value']), VectorClock.parse(resp['version'])",
        "",
        "    def put(self, key: str, value: bytes,",
        "             clock: VectorClock = None) -> VectorClock:",
        "        body = {'value': base64.encode(value),",
        "                'version': str(clock) if clock else None}",
        "        resp = self._request('PUT', f'/kv/{key}', body)",
        "        return VectorClock.parse(resp['version'])",
        "",
        "    def delete(self, key: str, clock: VectorClock) -> bool:",
        "        body = {'version': str(clock)}",
        "        resp = self._request('DELETE', f'/kv/{key}', body)",
        "        return resp['deleted']",
    ], styles))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph("2.3 Vector Clock Context for Versioning", styles['SubHeader']))
    story.append(Paragraph(
        "A <b>vector clock</b> is a list of (node_id, counter) pairs that tracks causality "
        "across distributed writes. When a client reads a value, it receives the current vector clock. "
        "When it writes back, it passes that clock, allowing the system to detect concurrent writes "
        "vs. causally ordered updates.", styles['BodyText']
    ))

    vc_data = [
        ["Field", "Type", "Example", "Description"],
        ["node_id", "string", '"node-3"', "Logical node identifier"],
        ["counter", "uint64", "42", "Monotonically increasing write counter per node"],
        ["timestamp", "uint64", "1711234567890", "Wall clock (ms) — used only for LWW tie-breaking"],
        ["Full clock", "string", '"[node1:5,node3:2]"', "Serialized as JSON array in HTTP header X-KV-Version"],
    ]
    story.append(make_table(vc_data, [1.0*inch, 0.8*inch, 1.4*inch, 3.1*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph("2.4 Request / Response Schemas", styles['SubHeader']))

    schema_data = [
        ["Schema", "Field", "Type", "Required", "Description"],
        ["PUT Request", "value", "string (base64)", "Yes", "Value to store"],
        ["PUT Request", "ttl", "integer (seconds)", "No", "Auto-expiry; 0 = no expiry"],
        ["PUT Request", "consistency", "enum: ONE/QUORUM/ALL", "No", "Override cluster default"],
        ["PUT Request", "X-KV-Version header", "string", "No", "Vector clock from prior GET"],
        ["GET Response", "value", "string (base64)", "Yes", "Retrieved value"],
        ["GET Response", "X-KV-Version header", "string", "Yes", "Current vector clock"],
        ["GET Response", "last_modified", "ISO-8601 timestamp", "Yes", "Server-side write time"],
        ["GET Response", "checksum", "string (hex SHA-256)", "Yes", "For client-side integrity check"],
        ["Error Response", "error_code", "string", "Yes", "Machine-readable code"],
        ["Error Response", "message", "string", "Yes", "Human-readable description"],
        ["Error Response", "retry_after", "integer (ms)", "No", "Present on 429/503 responses"],
    ]
    story.append(make_table(schema_data, [1.0*inch, 1.3*inch, 1.4*inch, 0.65*inch, 1.9*inch]))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 3: HIGH-LEVEL ARCHITECTURE
    # =========================================================================
    story.append(Paragraph("3. High-Level Architecture", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph(
        "The system follows a fully decentralized, peer-to-peer architecture inspired by Amazon "
        "Dynamo. Any node can act as a coordinator for any request—there is no single master. "
        "A load balancer distributes incoming requests across all healthy nodes.", styles['BodyText']
    ))

    story.append(ascii_box([
        "                      CLIENTS",
        "          +-------+  +-------+  +-------+",
        "          | App A |  | App B |  | App C |",
        "          +---+---+  +---+---+  +---+---+",
        "              |          |          |",
        "              +----------+----------+",
        "                         |",
        "                  +------+------+",
        "                  | Load Balancer|  (L4/L7, health-checked)",
        "                  +------+------+",
        "                         |",
        "         +---------------+----------------+",
        "         |               |                |",
        "    +----+----+    +----+----+     +----+----+",
        "    |  Node 1  |    |  Node 2  |     |  Node 3  |",
        "    |Coordinator|   |Coordinator|    |Coordinator|",
        "    +----+----+    +----+----+     +----+----+",
        "         |               |                |",
        "         +-------+-------+--------+-------+",
        "                 |                |",
        "           +-----+-----+    +-----+-----+",
        "           |   Node 4   |    |   Node 5   |",
        "           +-----------+    +-----------+",
        "",
        "  CONSISTENT HASH RING (5 physical nodes, 150 vnodes each = 750 tokens)",
        "",
        "     [N1-t0]--[N3-t1]--[N5-t2]--[N2-t3]--[N4-t4]--[N1-t5]-- ...",
        "        0        85      170      255      340      425     511",
        "",
        "  KEY REPLICATION: key -> hash -> coordinator -> replicate to next 2 nodes",
        "  Example: key 'user:42' -> token 180 -> N5 (primary) -> N2, N4 (replicas)",
    ], styles))
    story.append(Paragraph("Figure 1: High-level system architecture with 5 nodes and consistent hash ring.", styles['Caption']))

    story.append(Paragraph("Component Responsibilities", styles['SubHeader']))
    comp_data = [
        ["Component", "Responsibility", "Key Design Choice"],
        ["Load Balancer", "Route requests to healthy nodes", "L4 TCP passthrough; no sticky sessions needed"],
        ["Coordinator Node", "Hash key, build preference list, manage quorum", "Any node can coordinate; no single point of failure"],
        ["Storage Node", "Store key-value data in LSM-tree", "RocksDB or custom LSM; WAL for durability"],
        ["Gossip Daemon", "Membership, failure detection, metadata sync", "O(log N) messages per round; eventual convergence"],
        ["Compaction Worker", "Merge SSTables, garbage-collect tombstones", "Background thread with I/O rate limiting"],
        ["Anti-Entropy Worker", "Sync diverged replicas via Merkle trees", "Runs during low-traffic windows"],
    ]
    story.append(make_table(comp_data, [1.3*inch, 2.2*inch, 2.8*inch]))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 4: DATA MODEL
    # =========================================================================
    story.append(Paragraph("4. Data Model", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph("4.1 Key-Value Record Schema", styles['SubHeader']))
    story.append(Paragraph(
        "Each stored entry is a versioned record containing the value, conflict-resolution "
        "metadata, and integrity information:", styles['BodyText']
    ))

    record_data = [
        ["Field", "Type", "Size", "Description"],
        ["key", "bytes", "1–256 B", "Raw key bytes; no encoding enforced"],
        ["value", "bytes", "0–1 MB", "Opaque payload; application interprets encoding"],
        ["vector_clock", "[(node_id, counter)]", "~50 B typical", "Causality tracker; grows with concurrent writers"],
        ["timestamp", "uint64 (epoch ms)", "8 B", "Server-side write time; LWW tie-breaker only"],
        ["ttl", "uint32 (seconds)", "4 B", "0 = no expiry; triggers lazy deletion at read time"],
        ["tombstone", "bool", "1 B", "True if logically deleted; removed during compaction"],
        ["checksum", "bytes (SHA-256 truncated)", "8 B", "Detects silent data corruption (bit rot)"],
        ["flags", "uint8 bitmask", "1 B", "Compression codec, encryption flag, etc."],
    ]
    story.append(make_table(record_data, [1.2*inch, 1.7*inch, 1.1*inch, 2.3*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph("4.2 Metadata Store", styles['SubHeader']))
    story.append(Paragraph(
        "Cluster metadata is stored separately from user data, replicated via gossip to all nodes "
        "and persisted locally. It contains:", styles['BodyText']
    ))

    meta_data = [
        ["Metadata", "Storage Location", "Update Frequency", "Purpose"],
        ["Ring membership (node list)", "Gossip state + local disk", "On node join/leave", "Locate data owners"],
        ["Token assignments (vnode map)", "Gossip state + local disk", "On rebalance", "Map hash ranges to nodes"],
        ["Node health status", "In-memory gossip table", "Every heartbeat (1s)", "Coordinator selects healthy replicas"],
        ["Schema version", "Metadata namespace", "On schema change", "Backwards-compatibility enforcement"],
        ["Replication factor (N)", "Cluster config file", "Manual update", "Drives preference list size"],
        ["Hints queue", "Local disk per node", "On write to downed replica", "Hinted handoff recovery"],
    ]
    story.append(make_table(meta_data, [1.5*inch, 1.5*inch, 1.2*inch, 2.1*inch]))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 5: CONSISTENT HASHING & PARTITIONING
    # =========================================================================
    story.append(Paragraph("5. Consistent Hashing & Partitioning", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph(
        "Consistent hashing maps both keys and nodes onto a circular integer space (the \"ring\"). "
        "Each key is owned by the first node clockwise from its hash position. This minimizes "
        "key remapping when nodes join or leave—only keys on the affected arc need to move, "
        "roughly K/N keys on average.", styles['BodyText']
    ))

    # Fun analogy box
    analogy_table = Table([[
        Paragraph(
            '<b>Analogy: The Ice Cream Stand Ring</b> — Imagine 5 ice cream stands placed '
            'around a circular park. Customers (keys) walk clockwise until they find the nearest stand. '
            'If one stand closes (node failure), its customers just walk a bit further to the next one. '
            'Only the customers "between" the closed stand and the previous one are affected — not everyone. '
            'Virtual nodes are like each stand having multiple pop-up carts scattered around the park, '
            'ensuring no single area is ever too crowded.',
            ParagraphStyle('AnalogyText', fontSize=9, fontName='Helvetica-Oblique', leading=13,
                           textColor=HexColor('#276749'))
        )
    ]], colWidths=[6.5 * inch])
    analogy_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), HexColor('#f0fff4')),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('BOX', (0, 0), (-1, -1), 1, HexColor('#276749')),
        ('ROUNDEDCORNERS', [4, 4, 4, 4]),
    ]))
    story.append(analogy_table)
    story.append(Spacer(1, 0.12 * inch))

    story.append(Paragraph("5.1 Virtual Nodes (vnodes)", styles['SubHeader']))
    story.append(Paragraph(
        "Instead of assigning a single token range to each physical node, each node is assigned "
        "<b>150+ virtual tokens</b> distributed evenly around the ring. This provides three benefits:", styles['BodyText']
    ))

    for item in [
        "<b>Load balancing:</b> With random token placement, 150 vnodes per node gives ~1% standard deviation in load distribution, vs. ~40% for single tokens.",
        "<b>Heterogeneous hardware:</b> Nodes with more RAM/CPU can be assigned proportionally more vnodes.",
        "<b>Faster rebalancing:</b> When a node is added, it takes a small fraction of tokens from many existing nodes instead of a large fraction from one, allowing parallel data movement.",
    ]:
        story.append(Paragraph(f"• {item}", styles['BulletItem']))

    story.append(Spacer(1, 0.1*inch))
    vnode_data = [
        ["Nodes", "vnodes/node", "Total tokens", "Load std dev (theoretical)"],
        ["5", "1 (naive)", "5", "~40%"],
        ["5", "150", "750", "~1.3%"],
        ["20", "256 (Cassandra default)", "5120", "~0.5%"],
        ["100", "256", "25600", "~0.2%"],
    ]
    story.append(make_table(vnode_data, [0.8*inch, 1.2*inch, 1.2*inch, 3.1*inch]))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(
        "<b>Real-world:</b> Apache Cassandra uses 256 vnodes per node by default (configurable via "
        "num_tokens). Amazon DynamoDB uses consistent hashing internally with automatic token reassignment.",
        styles['Note']
    ))

    story.append(Paragraph("5.2 Token Ring Diagram", styles['SubHeader']))
    story.append(ascii_box([
        "  TOKEN RING — 5 physical nodes, 4 vnodes each shown (150 in production)",
        "",
        "                         0 / 512",
        "                    N1-t0 * * N3-t0",
        "               N2-t3 *             * N1-t3",
        "              /                           \\",
        "        N5-t2 *   CONSISTENT HASH RING     * N2-t0",
        "             |         (SHA-256 mod 2^32)  |",
        "        N4-t2 *                            * N3-t1",
        "              \\                           /",
        "               N1-t1 *             * N4-t1",
        "                    N5-t0 * * N4-t0",
        "                         256",
        "",
        "  KEY LOOKUP EXAMPLE:",
        "  hash('user:alice') = 73  -->  first node clockwise: N3-t0 (token 85)",
        "  Coordinator: N3 (physical node owning token 85)",
        "  Replicas:    N5 (next token 170), N2 (next token 255)",
        "",
        "  Token assignment (sorted):",
        "  Token:  0   17  34  51  68  85 102 119 136 153 170 ...",
        "  Owner: N1  N2  N3  N4  N5  N3  N1  N4  N2  N5  N5 ...",
    ], styles))
    story.append(Paragraph("Figure 2: Consistent hash ring with virtual nodes.", styles['Caption']))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 6: REPLICATION STRATEGY
    # =========================================================================
    story.append(Paragraph("6. Replication Strategy", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph("6.1 Preference List Construction", styles['SubHeader']))
    story.append(Paragraph(
        "For a given key K with replication factor N=3, the <b>preference list</b> is the ordered "
        "list of nodes responsible for storing K. It is constructed as follows:", styles['BodyText']
    ))
    for i, step in enumerate([
        "Hash key K to position P on the ring.",
        "Walk clockwise from P, collecting the first N <i>distinct physical nodes</i> (skipping duplicate vnodes from the same physical node).",
        "The first node in the list becomes the <b>coordinator</b> for this write.",
        "The coordinator replicates to the remaining N-1 nodes in the preference list.",
        "The coordinator is the primary; the others are secondary replicas.",
    ], 1):
        story.append(Paragraph(f"{i}. {step}", styles['NumberedItem']))

    story.append(Spacer(1, 0.1*inch))
    story.append(ascii_box([
        "  PREFERENCE LIST EXAMPLE (N=3):",
        "",
        "  Key: 'order:9987'  ->  hash -> token 290",
        "",
        "  Ring scan clockwise from 290:",
        "    Token 340 -> N4  (physical node A)  --> COORDINATOR + Primary replica",
        "    Token 425 -> N1  (physical node B)  --> Secondary replica",
        "    Token 470 -> N3  (physical node C)  --> Tertiary replica",
        "    Token 510 -> N4  (skip - same physical as A)",
        "    Token  85 -> N3  (skip - same physical as C)",
        "",
        "  Preference list: [N4, N1, N3]",
        "  Stored on: AZ-1 (N4), AZ-2 (N1), AZ-3 (N3)  <- Rack-aware",
    ], styles))

    story.append(Paragraph("6.2 Rack / AZ Awareness", styles['SubHeader']))
    story.append(Paragraph(
        "To survive an entire Availability Zone (AZ) outage, the preference list must span "
        "at least two AZs. The ring-walk algorithm is augmented to skip nodes in already-selected "
        "AZs until all N replicas span distinct AZs:", styles['BodyText']
    ))

    az_data = [
        ["Scenario", "AZ Layout", "Outcome"],
        ["3 replicas, 3 AZs", "N4(AZ-1), N1(AZ-2), N3(AZ-3)", "Survives single AZ failure — full availability"],
        ["3 replicas, 2 AZs", "N4(AZ-1), N1(AZ-2), N3(AZ-2)", "Survives 1 AZ failure if AZ-1 goes down"],
        ["3 replicas, 1 AZ", "N4(AZ-1), N1(AZ-1), N3(AZ-1)", "No AZ redundancy — avoid in production"],
        ["5 replicas, 3 AZs", "2+2+1 distribution", "Survives 2-node failure in any single AZ"],
    ]
    story.append(make_table(az_data, [1.5*inch, 2.0*inch, 2.8*inch]))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 7: CONSISTENCY MODEL
    # =========================================================================
    story.append(Paragraph("7. Consistency Model", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph("7.1 Quorum Reads & Writes", styles['SubHeader']))
    story.append(Paragraph(
        "Quorum-based consistency uses the relationship <b>W + R > N</b> to guarantee that any "
        "read sees at least one node that participated in the most recent write. W is the number "
        "of nodes that must acknowledge a write; R is the number of nodes queried on a read.", styles['BodyText']
    ))

    quorum_data = [
        ["N", "W", "R", "W+R", "Consistency", "Use Case", "Trade-off"],
        ["3", "1", "1", "2", "Eventual", "High-throughput cache", "Stale reads possible"],
        ["3", "2", "1", "3", "Read-your-writes", "Session-consistent app", "1 slow replica blocks read"],
        ["3", "1", "2", "3", "Read-heavy strong", "Read-heavy workload", "1 slow replica blocks read"],
        ["3", "2", "2", "4", "Strong (default)", "General purpose", "Balanced latency/consistency"],
        ["3", "3", "1", "4", "Strong write-heavy", "Audit logs, ledgers", "Write blocks if any replica slow"],
        ["3", "3", "3", "6", "Full sync", "Critical financial data", "High latency, low availability"],
    ]
    story.append(make_table(quorum_data, [0.35*inch, 0.35*inch, 0.35*inch, 0.45*inch, 1.2*inch, 1.4*inch, 2.2*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph("7.2 CAP Theorem Positioning", styles['SubHeader']))
    story.append(Paragraph(
        "This system is an <b>AP (Available + Partition-tolerant)</b> system. During a network "
        "partition, the system chooses to remain available (accept reads/writes) at the cost of "
        "potential inconsistency. Consistency is restored after the partition heals via anti-entropy.", styles['BodyText']
    ))

    cap_data = [
        ["Property", "Choice", "Implication"],
        ["Consistency (C)", "Sacrificed during partitions", "Stale reads possible; eventual consistency guaranteed"],
        ["Availability (A)", "Preserved", "Cluster continues serving with degraded replica count"],
        ["Partition Tolerance (P)", "Always required", "Network partitions are inevitable in distributed systems"],
    ]
    story.append(make_table(cap_data, [1.5*inch, 1.8*inch, 3.0*inch]))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(
        "<b>PACELC extension:</b> During normal operation (no partition), the system trades "
        "<b>Latency for Consistency</b> — lower W/R values reduce latency at the cost of "
        "potential stale reads.", styles['Note']
    ))

    story.append(Paragraph("7.3 Vector Clocks & Conflict Detection", styles['SubHeader']))
    story.append(Paragraph(
        "Vector clocks track causality. When comparing two versions V1 and V2:", styles['BodyText']
    ))

    vc_compare_data = [
        ["Relationship", "Condition", "Action"],
        ["V1 causally before V2", "All V1[i] <= V2[i] and at least one <", "Discard V1; V2 is the winner"],
        ["V2 causally before V1", "All V2[i] <= V1[i] and at least one <", "Discard V2; V1 is the winner"],
        ["Concurrent (conflict)", "Neither dominates the other", "Keep both versions; application resolves or LWW"],
        ["Identical", "All V1[i] == V2[i]", "Deduplicate; return either"],
    ]
    story.append(make_table(vc_compare_data, [1.5*inch, 2.2*inch, 2.6*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph(
        "<b>Last-Write-Wins (LWW):</b> A simpler alternative that uses wall-clock timestamp as "
        "the tiebreaker. LWW is lossy (concurrent writes cause data loss) but is appropriate for "
        "use cases where the most recent value is always the correct one (e.g., sensor readings, "
        "user presence status).", styles['BodyText']
    ))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 8: WRITE PATH
    # =========================================================================
    story.append(Paragraph("8. Write Path (Detailed)", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph("8.1 Step-by-Step Walkthrough", styles['SubHeader']))
    story.append(Paragraph(
        "The following describes a PUT operation with N=3, W=2 (default quorum write):", styles['BodyText']
    ))

    write_steps = [
        ("1", "Client → Load Balancer → Coordinator", "0–1ms",
         "Client sends PUT /kv/{key} to any available node via the load balancer. "
         "The receiving node becomes the coordinator."),
        ("2", "Key hashing & preference list", "~0.01ms",
         "Coordinator computes SHA-256(key) mod 2^32 to find the token position. "
         "Walks the ring clockwise to build the preference list of N=3 distinct physical nodes."),
        ("3", "Local write to storage engine", "1–3ms",
         "Coordinator writes to WAL (Write-Ahead Log) first for durability, then to MemTable "
         "(in-memory). Returns immediately after WAL fsync."),
        ("4", "Parallel replication to N-1 replicas", "0ms (parallel with step 5)",
         "Coordinator sends the write to the other 2 preference list nodes in parallel, "
         "without waiting for a response before initiating both requests."),
        ("5", "Wait for W acknowledgments", "2–8ms",
         "Coordinator waits for W=2 total ACKs (including its own local write). "
         "Uses a deadline timer; if timeout, returns error to client."),
        ("6", "Return success to client", "0.1ms",
         "Once W=2 ACKs received, coordinator responds 200 OK with the new vector clock "
         "in the X-KV-Version response header."),
        ("7", "Background: async replication", "async",
         "If any replica was slow or down, the coordinator either waits for it in the "
         "background (if W=3) or uses hinted handoff to queue the write for later delivery."),
    ]

    write_data = [["Step", "Action", "Latency", "Details"]]
    for step, action, lat, detail in write_steps:
        write_data.append([step, action, lat, detail])

    story.append(make_table(write_data, [0.4*inch, 1.5*inch, 0.9*inch, 3.5*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(ascii_box([
        "  WRITE PATH TIMELINE (N=3, W=2):",
        "",
        "  Client      Load        Coordinator    Replica-2    Replica-3",
        "    |          Balancer       (N4)           (N1)         (N3)",
        "    |----PUT----->|             |              |            |",
        "    |             |---PUT------>|              |            |",
        "    |             |             |--WAL write-->|            |",
        "    |             |             | (local, 2ms) |            |",
        "    |             |             |----replicate----------->  |",
        "    |             |             |----replicate----->        |",
        "    |             |             |<---ACK (3ms)---           |",
        "    |             |             | W=2 met!                  |",
        "    |<---200 OK---|             |<--------ACK (5ms)---------|",
        "    |  (total: ~5ms)            |  (background, W already met)",
        "",
        "  Total client-observed latency: ~5ms (dominated by slowest of first 2 ACKs)",
    ], styles))
    story.append(Paragraph("Figure 3: Write path timing diagram for W=2 quorum write.", styles['Caption']))

    story.append(Paragraph("8.2 Hinted Handoff", styles['SubHeader']))
    story.append(Paragraph(
        "When a preferred replica (e.g., N1) is unreachable, the coordinator does not fail the "
        "write (provided W can still be satisfied). Instead, it writes the data to a healthy "
        "node (e.g., N2) tagged with a <b>hint</b> indicating the intended destination:", styles['BodyText']
    ))

    for point in [
        "The hint is stored in a separate 'hints queue' on N2's disk.",
        "N2 periodically checks if N1 has recovered (via gossip membership table).",
        "Once N1 is back, N2 replays all queued hints to N1, restoring full replication.",
        "Hints are purged after successful delivery or after a configurable TTL (e.g., 1 hour).",
        "This ensures <b>W writes succeed</b> even when preferred nodes are temporarily down.",
    ]:
        story.append(Paragraph(f"• {point}", styles['BulletItem']))

    story.append(Paragraph("8.3 Latency Budget", styles['SubHeader']))
    latency_data = [
        ["Component", "p50", "p99", "Notes"],
        ["Client → LB → Coordinator (network)", "0.5ms", "2ms", "Same datacenter"],
        ["Key hash + preference list lookup", "0.01ms", "0.1ms", "In-memory ring map"],
        ["WAL fsync (NVMe SSD)", "0.3ms", "1ms", "Group commit batching"],
        ["MemTable insert", "0.01ms", "0.05ms", "Lock-free skip list"],
        ["Network: coordinator → replica", "0.5ms", "2ms", "Same datacenter"],
        ["Replica WAL fsync", "0.3ms", "1ms", "Parallel with other replica"],
        ["Coordinator → client response", "0.1ms", "0.5ms", "After W ACKs received"],
        ["Total (W=2 quorum write)", "3ms", "10ms", "Dominated by replica WAL + network"],
    ]
    story.append(make_table(latency_data, [2.2*inch, 0.7*inch, 0.7*inch, 2.7*inch]))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 9: READ PATH
    # =========================================================================
    story.append(Paragraph("9. Read Path (Detailed)", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph("9.1 Step-by-Step Walkthrough", styles['SubHeader']))
    story.append(Paragraph(
        "The following describes a GET operation with N=3, R=2 (default quorum read):", styles['BodyText']
    ))

    read_steps = [
        ("1", "Client → Coordinator", "0–1ms",
         "Client sends GET /kv/{key}. Coordinator is any healthy node (same as write path)."),
        ("2", "Preference list construction", "~0.01ms",
         "Coordinator builds the same preference list as for writes. Selects R=2 nodes to query."),
        ("3", "Parallel queries to R replicas", "0ms (parallel)",
         "Coordinator sends read requests to all N=3 replicas simultaneously (or R=2 + hedged request). "
         "Uses the first R responses."),
        ("4", "Wait for R responses", "1–5ms",
         "Coordinator waits for R=2 responses. If a replica is slow, a hedged request may be "
         "sent to the third replica after a short timeout (e.g., 2ms)."),
        ("5", "Vector clock comparison", "~0.01ms",
         "Compare vector clocks of all received versions. Return the causally latest version "
         "to the client. If concurrent versions exist, return all of them for client resolution."),
        ("6", "Trigger read repair (async)", "async",
         "If any replica returned a stale version, coordinator asynchronously sends the latest "
         "version to the stale replica to bring it up to date."),
    ]

    read_data = [["Step", "Action", "Latency", "Details"]]
    for step, action, lat, detail in read_steps:
        read_data.append([step, action, lat, detail])
    story.append(make_table(read_data, [0.4*inch, 1.5*inch, 0.9*inch, 3.5*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(ascii_box([
        "  READ PATH TIMELINE (N=3, R=2):",
        "",
        "  Client     Coordinator    Replica-2    Replica-3",
        "    |            (N4)           (N1)         (N3)",
        "    |---GET------>|              |            |",
        "    |             |----read-------------->    |",
        "    |             |----read----->             |",
        "    |             |<---v:[N4:5](2ms)----      |",
        "    |             |<---v:[N4:5](3ms)----------| R=2 met!",
        "    |             | compare vectors -> same   |",
        "    |<--200 OK----|              |            |",
        "    | (total ~4ms)|              |            |",
        "    |             |  [if stale: async repair] |",
        "",
        "  STALE REPLICA SCENARIO:",
        "    Replica-3 returns v:[N4:4] (older version)",
        "    Coordinator detects stale via vector clock comparison",
        "    Returns latest v:[N4:5] to client immediately",
        "    Asynchronously sends PUT v:[N4:5] to Replica-3 (read repair)",
    ], styles))
    story.append(Paragraph("Figure 4: Read path timeline with read repair.", styles['Caption']))

    story.append(Paragraph("9.2 Read Repair Mechanism", styles['SubHeader']))
    story.append(Paragraph(
        "Read repair is a lazy consistency mechanism that heals stale replicas during normal "
        "read operations, without requiring a dedicated background process:", styles['BodyText']
    ))
    repair_data = [
        ["Aspect", "Detail"],
        ["Trigger", "Coordinator detects version mismatch after receiving R responses"],
        ["Scope", "Only the specific key that was just read"],
        ["Delivery", "Async write to stale replica; does not block client response"],
        ["Limitation", "Only repairs keys that are actually read — cold keys may stay stale"],
        ["Complement", "Anti-entropy with Merkle trees handles keys that are never read"],
        ["Rate limiting", "Repair writes are throttled to avoid overloading replicas"],
    ]
    story.append(make_table(repair_data, [1.5*inch, 4.8*inch]))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 10: FAILURE HANDLING
    # =========================================================================
    story.append(Paragraph("10. Failure Handling", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph("10.1 Hinted Handoff", styles['SubHeader']))
    story.append(Paragraph(
        "Hinted handoff maintains write availability when a replica is temporarily unavailable. "
        "The coordinator stores hints — pending writes with metadata about the intended recipient — "
        "and delivers them when the target recovers:", styles['BodyText']
    ))

    hh_data = [
        ["Phase", "Action", "Storage Location"],
        ["Write time (replica down)", "Write value + hint metadata to healthy substitute node", "Hints queue on substitute node disk"],
        ["Recovery detection", "Gossip confirms target node is live again", "In-memory gossip state"],
        ["Hint replay", "Substitute streams all queued hints to recovered node", "Removed from hints queue after ACK"],
        ["Cleanup", "If hint TTL expires before node recovers, discard hint", "Repair via anti-entropy instead"],
    ]
    story.append(make_table(hh_data, [1.4*inch, 2.6*inch, 2.3*inch]))

    story.append(Paragraph("10.2 Anti-Entropy with Merkle Trees", styles['SubHeader']))
    story.append(Paragraph(
        "Anti-entropy is a background process that ensures all replicas eventually converge "
        "to the same state. It uses <b>Merkle trees</b> to efficiently identify diverged key ranges "
        "without transferring the full dataset:", styles['BodyText']
    ))

    for point in [
        "Each replica builds a Merkle tree over its assigned key range. Leaf nodes contain hashes of individual key-value pairs; internal nodes contain hashes of their children.",
        "Two replicas compare their root hashes. If roots match, the entire key range is in sync.",
        "If roots differ, they recurse into subtrees, identifying only the diverged leaf nodes.",
        "Communication complexity is <b>O(log N) messages</b> to identify differences instead of transferring all keys.",
        "Only the diverged keys are synchronized, minimizing network overhead.",
    ]:
        story.append(Paragraph(f"• {point}", styles['BulletItem']))
    story.append(Spacer(1, 0.1*inch))

    story.append(ascii_box([
        "  MERKLE TREE STRUCTURE (key range: tokens 0-511, depth 3):",
        "",
        "                  [ROOT: h(h01, h23)]",
        "                  /                  \\",
        "          [h01: h(h0,h1)]          [h23: h(h2,h3)]",
        "          /           \\            /             \\",
        "    [h0: hash(       [h1: hash(  [h2: hash(    [h3: hash(",
        "     keys 0-127)]    keys128-255)] keys256-383)] keys384-511)]",
        "          |               |              |             |",
        "    [individual key hashes at leaf level]",
        "",
        "  SYNC PROTOCOL between Replica-A and Replica-B:",
        "",
        "    1. A sends root hash to B",
        "    2. B compares: if equal -> DONE (no sync needed)",
        "    3. If different: B requests A's children hashes",
        "    4. Recurse until diverged leaves identified",
        "    5. A sends only diverged key-value pairs to B",
        "",
        "  EXAMPLE: Only keys in range 256-383 differ",
        "    Messages: 3 hash comparisons + 1 data transfer",
        "    vs. naive full scan: 1B key comparisons",
        "    Savings: O(log N) vs O(N)  -- massive for large datasets",
    ], styles))
    story.append(Paragraph("Figure 5: Merkle tree structure and anti-entropy sync protocol.", styles['Caption']))

    story.append(Paragraph("10.3 Gossip Protocol", styles['SubHeader']))
    story.append(Paragraph(
        "The gossip protocol provides decentralized failure detection and membership management. "
        "It requires no central coordinator and scales to hundreds of nodes:", styles['BodyText']
    ))

    gossip_data = [
        ["Aspect", "Detail"],
        ["Message type", "Each node maintains a membership list: {node_id -> (heartbeat_counter, timestamp)}"],
        ["Gossip frequency", "Each node selects 1–3 random peers every 1 second and sends its membership list"],
        ["Heartbeat update", "On receiving a gossip message, a node updates counters if received > local"],
        ["Failure detection", "If a node's heartbeat_counter has not increased in T_fail = 10 seconds, mark SUSPECT"],
        ["Confirmation", "If heartbeat stays stale for T_remove = 60 seconds, mark REMOVED from ring"],
        ["Convergence time", "O(log N) rounds to propagate info to all nodes; ~7 rounds for 100-node cluster"],
        ["Bandwidth", "Each gossip message: ~1KB for 100-node cluster; ~1KB/sec per node"],
    ]
    story.append(make_table(gossip_data, [1.6*inch, 4.7*inch]))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(
        "<b>Real-world:</b> Apache Cassandra uses this exact gossip approach. Each node sends "
        "a GossipDigestSyn message every second to 3 random peers. Failure detection uses "
        "the Phi Accrual Failure Detector, which outputs a suspicion level (phi) rather than "
        "a binary alive/dead decision.",
        styles['Note']
    ))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 11: NETWORK PARTITION HANDLING
    # =========================================================================
    story.append(Paragraph("11. Network Partition Handling", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph(
        "Network partitions are inevitable in distributed systems. This section describes how "
        "the system behaves when 2 of 5 nodes become isolated from the remaining 3.", styles['BodyText']
    ))

    story.append(ascii_box([
        "  PARTITION SCENARIO: 5-node cluster, 2 nodes isolated",
        "",
        "  Before partition:                After partition:",
        "  +----+  +----+  +----+           MAJORITY PARTITION     MINORITY",
        "  | N1 |--| N2 |--| N3 |           +----+  +----+  +----+ | +----+  +----+",
        "  +----+  +----+  +----+           | N1 |--| N2 |--| N3 | | | N4 |--| N5 |",
        "    |       |       |              +----+  +----+  +----+ | +----+  +----+",
        "  +----+  +----+                    ↑ quorum possible ↑    ↑ no quorum ↑",
        "  | N4 |--| N5 |                                           (W=2 fails)",
        "  +----+  +----+",
        "",
        "  WITH W=2, R=2 (default):                                              ",
        "    Majority side (N1,N2,N3): CAN serve reads and writes (3 nodes >= W=2)",
        "    Minority side (N4,N5):    CANNOT complete writes (only 2 nodes; W=2  ",
        "                              satisfied internally, but preference list   ",
        "                              nodes N1-N3 are unreachable)               ",
        "",
        "  SLOPPY QUORUM:",
        "    With sloppy quorum enabled, N4 and N5 CAN still accept writes using  ",
        "    any 2 available nodes, storing with hints for the actual owners.     ",
        "    After partition heals, N4/N5 replay hints to N1/N2/N3.              ",
    ], styles))
    story.append(Paragraph("Figure 6: Network partition with 2/5 nodes isolated.", styles['Caption']))

    partition_data = [
        ["Phase", "Action", "W=2 behavior", "Sloppy Quorum behavior"],
        ["During partition", "N4, N5 isolated from N1, N2, N3", "N4/N5 reject writes (can't reach quorum owners)", "N4/N5 accept writes using each other as substitute replicas"],
        ["Majority side", "N1, N2, N3 operational", "Full read/write service continues", "No change; quorum always met"],
        ["Partition heals", "All 5 nodes reconnect", "Anti-entropy syncs any diverged keys", "Hints replayed; Merkle tree reconciliation"],
        ["Conflict resolution", "Concurrent writes to both sides", "Vector clocks detect conflicts", "LWW or application-level resolution"],
    ]
    story.append(make_table(partition_data, [1.0*inch, 1.4*inch, 1.5*inch, 2.4*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph(
        "<b>Key insight:</b> A system with N=5, W=2 can survive a 2-node partition while "
        "the majority side (3 nodes) remains fully operational. The minority side loses "
        "availability unless sloppy quorum is enabled. This is the fundamental AP trade-off.",
        styles['Note']
    ))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 12: STORAGE ENGINE
    # =========================================================================
    story.append(Paragraph("12. Storage Engine (LSM-Tree)", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph(
        "The storage layer uses a <b>Log-Structured Merge Tree (LSM-Tree)</b>, which provides "
        "superior write throughput compared to B-trees by converting random writes into "
        "sequential I/O. This is critical for a write-heavy key-value store.", styles['BodyText']
    ))

    lsm_data = [
        ["Component", "Location", "Size", "Purpose"],
        ["MemTable", "RAM (skip list)", "64–256 MB", "Active write buffer; supports O(log N) point lookups"],
        ["Write-Ahead Log (WAL)", "Disk (sequential)", "Unbounded (rotated)", "Crash recovery; replayed on restart"],
        ["L0 SSTables", "Disk (NVMe)", "64 MB each, ~10 files", "Flushed MemTables; may overlap key ranges"],
        ["L1 SSTables", "Disk (NVMe)", "256 MB total", "Compacted L0; sorted, no overlap"],
        ["L2+ SSTables", "Disk (SATA/HDD)", "10x growth per level", "Cold data; infrequently accessed"],
        ["Bloom Filters", "RAM (per SSTable)", "~10 bits/key", "Probabilistic check: key definitely absent → skip SSTable"],
        ["Block Cache", "RAM (LRU)", "Configurable (1–64 GB)", "Cache hot SSTable blocks; dramatically reduces read I/O"],
        ["Compaction", "Background CPU/IO", "N/A", "Merge SSTables, garbage-collect deleted/overwritten keys"],
    ]
    story.append(make_table(lsm_data, [1.3*inch, 1.2*inch, 1.3*inch, 2.5*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(ascii_box([
        "  LSM-TREE WRITE PATH:",
        "",
        "  Client Write",
        "      |",
        "      v",
        "  +---+-----+  1. Append to WAL (sequential, fast)",
        "  |   WAL   |  2. Insert into MemTable (in-memory skip list)",
        "  +---------+",
        "      |",
        "      v",
        "  +-----------+  When MemTable full (256MB):",
        "  | MemTable  |  3. Flush to L0 SSTable (sorted, immutable)",
        "  +-----------+",
        "      |  (flush)",
        "      v",
        "  +--+--+--+--+  L0: 0–10 SSTables (may overlap)",
        "  |S0|S1|S2|S3|",
        "  +--+--+--+--+",
        "      |  (compaction: merge + sort)",
        "      v",
        "  +-------------+  L1: Non-overlapping, 256MB total",
        "  |     L1      |",
        "  +-------------+",
        "      |  (compaction)",
        "      v",
        "  +-------------+  L2+: Each level 10x larger",
        "  |   L2, L3... |",
        "  +-------------+",
        "",
        "  LSM-TREE READ PATH (worst case, with bloom filters):",
        "  1. Check MemTable (O(log N) in skip list)        -> found? return",
        "  2. Check Block Cache (O(1) hash lookup)          -> found? return",
        "  3. Check L0 SSTables newest to oldest            -> bloom filter first",
        "     (bloom says 'maybe present' -> binary search SSTable)",
        "  4. Check L1 SSTable (only one overlaps key)      -> bloom filter",
        "  5. Check L2+ ... (one per level)                 -> bloom filter",
        "  Without bloom filters: read amplification ~10x",
        "  With bloom filters:    ~1.5 I/Os for 99% of reads",
    ], styles))
    story.append(Paragraph("Figure 7: LSM-Tree write and read paths.", styles['Caption']))

    story.append(Paragraph(
        "<b>Real-world:</b> RocksDB (Facebook's LSM-tree implementation) is used as the "
        "storage engine by Apache Cassandra (via JNA), TiKV (Rust bindings), and CockroachDB. "
        "It achieves 400K+ write IOPS on NVMe SSDs with compaction I/O rate limiting.",
        styles['Note']
    ))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 13: SECONDARY INDEXES
    # =========================================================================
    story.append(Paragraph("13. Secondary Indexes", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph(
        "The primary index maps key → value via consistent hashing. Secondary indexes allow "
        "lookups by non-key attributes (e.g., find all users by email). Two approaches exist:",
        styles['BodyText']
    ))

    story.append(Paragraph("13.1 Local (Node-Scoped) Indexes", styles['SubHeader']))
    story.append(Paragraph(
        "Each node maintains an index only for the keys it stores. Queries require a "
        "<b>scatter-gather</b> pattern: send the query to all N nodes, collect and merge results.",
        styles['BodyText']
    ))

    story.append(Paragraph("13.2 Global Indexes", styles['SubHeader']))
    story.append(Paragraph(
        "A dedicated set of index nodes maintains the full index across all partitions. "
        "A write to the primary data must also update the global index — this requires "
        "a distributed transaction or eventual consistency.", styles['BodyText']
    ))

    idx_data = [
        ["Aspect", "Local Index", "Global Index"],
        ["Write overhead", "Low — local update only", "High — must update remote index node"],
        ["Read pattern", "Scatter-gather to all N nodes", "Single index node lookup"],
        ["Read latency", "High (fan-out)", "Low (direct lookup)"],
        ["Consistency", "Eventual (index may lag)", "Tunable (can be strong with 2PC)"],
        ["Fan-out", "O(nodes) per query", "O(1) per query"],
        ["Failure impact", "Partial results if nodes down", "Index node becomes SPOF without replication"],
        ["Use case", "Low-cardinality attributes", "High-cardinality, latency-sensitive lookups"],
        ["Example", "Cassandra secondary indexes", "DynamoDB Global Secondary Indexes (GSI)"],
    ]
    story.append(make_table(idx_data, [1.3*inch, 2.5*inch, 2.5*inch]))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 14: CAPACITY PLANNING
    # =========================================================================
    story.append(Paragraph("14. Capacity Planning", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph("14.1 Data Model Assumptions", styles['SubHeader']))
    assump_data = [
        ["Parameter", "Value", "Notes"],
        ["Total keys", "1 billion", "1 × 10^9"],
        ["Average key size", "100 bytes", "~100 character string keys"],
        ["Average value size", "10 KB", "JSON objects, serialized protobufs"],
        ["Replication factor (N)", "3", "Cross-AZ replication"],
        ["Metadata overhead", "~5% of data size", "Vector clocks, bloom filters, SSTable indexes"],
        ["Compression ratio", "~2:1", "LZ4 compression on SSTables"],
        ["Write amplification (LSM)", "~10x", "Compaction magnifies disk writes"],
        ["Read amplification (LSM)", "~2x", "With bloom filters enabled"],
    ]
    story.append(make_table(assump_data, [2.0*inch, 1.5*inch, 2.8*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph("14.2 Storage Calculation", styles['SubHeader']))
    story.append(ascii_box([
        "  STORAGE MATH:",
        "",
        "  Raw data:    1B keys × (100B key + 10KB value + ~50B metadata)",
        "             = 1B × 10.15 KB ≈ 10.15 TB uncompressed",
        "",
        "  Compressed:  10.15 TB / 2 (LZ4) ≈ 5 TB compressed",
        "",
        "  Replicated:  5 TB × 3 replicas = 15 TB total cluster storage",
        "",
        "  With 20 nodes: 15 TB / 20 = 750 GB per node",
        "  Node sizing:   2 TB NVMe SSD per node -> 62% utilization (good headroom)",
        "",
        "  WRITE AMPLIFICATION IMPACT:",
        "  Write amplification factor: ~10x (typical for leveled compaction)",
        "  Effective disk writes: 1M writes/sec × 10KB × 10x = 100 GB/sec disk write across cluster",
        "  Per node: 100 GB/sec / 20 nodes = 5 GB/sec -> NVMe handles 3-7 GB/sec sequentially",
        "  Action: Rate-limit compaction I/O; use io_uring for async I/O",
    ], styles))

    story.append(Paragraph("14.3 QPS and Node Sizing", styles['SubHeader']))
    sizing_data = [
        ["Metric", "Target", "Per Node (20 nodes)", "Headroom (3x)"],
        ["Write QPS", "1M writes/sec", "50K writes/sec", "150K/sec node capacity"],
        ["Read QPS", "10M reads/sec", "500K reads/sec", "1.5M/sec with cache hits"],
        ["Storage", "15 TB total", "750 GB/node", "2 TB SSD/node (37% used)"],
        ["Network ingress", "1M × 10KB = 10 GB/sec", "500 MB/sec/node", "10 GbE = 1.25 GB/sec"],
        ["Network egress", "10M × 10KB = 100 GB/sec", "5 GB/sec/node", "25 GbE = 3.1 GB/sec"],
        ["CPU", "Hashing, compression, Merkle tree builds", "~8 cores @ 60%", "16-core nodes recommended"],
        ["RAM", "MemTable + Block Cache", "8 GB MemTable + 32 GB cache", "64 GB RAM per node"],
    ]
    story.append(make_table(sizing_data, [1.7*inch, 1.6*inch, 1.6*inch, 2.1*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph("14.4 Throughput Per Node Summary", styles['SubHeader']))
    throughput_data = [
        ["Node Spec", "Write IOPS", "Read IOPS", "Disk Bandwidth", "Network"],
        ["16-core, 64GB RAM, 2TB NVMe", "150K/sec", "1.5M/sec (with cache)", "3 GB/sec (NVMe)", "25 GbE"],
        ["32-core, 128GB RAM, 4TB NVMe", "300K/sec", "3M/sec", "6 GB/sec (NVMe RAID)", "25 GbE"],
        ["8-core, 32GB RAM, 2TB SATA SSD", "50K/sec", "500K/sec", "500 MB/sec", "10 GbE"],
    ]
    story.append(make_table(throughput_data, [2.2*inch, 1.0*inch, 1.5*inch, 1.5*inch, 0.8*inch]))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 15: REAL-WORLD IMPLEMENTATIONS
    # =========================================================================
    story.append(Paragraph("15. Real-World Implementations", styles['SectionHeader']))
    story.append(section_hr())

    # Battle-tested stats banner
    stats2 = [
        ("DynamoDB", "Handles Amazon\nPrime Day: 105M\nrequests/sec peak"),
        ("Cassandra", "Apple runs\n75,000+ nodes\nacross multiple DCs"),
        ("Redis", "Twitter uses\n~1,000 Redis\ncluster instances"),
        ("etcd", "Powers every\nKubernetes cluster\nin production"),
    ]
    stats2_data = [[
        Paragraph(
            f'<b><font size="9" color="#2b6cb0">{s[0]}</font></b><br/>'
            f'<font size="8" color="#4a5568">{s[1]}</font>',
            ParagraphStyle('StatCell2', alignment=TA_CENTER, leading=12)
        )
        for s in stats2
    ]]
    stats2_table = Table(stats2_data, colWidths=[1.625 * inch] * 4)
    stats2_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), HexColor('#ebf8ff')),
        ('GRID', (0, 0), (-1, -1), 1, HexColor('#bee3f8')),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(stats2_table)
    story.append(Spacer(1, 0.12 * inch))

    story.append(Paragraph(
        "The following table compares production distributed key-value stores across key "
        "architectural dimensions:", styles['BodyText']
    ))

    rw_data = [
        ["System", "Consistency", "Replication", "Storage Engine", "Special Feature / Notes"],
        ["Amazon DynamoDB", "Tunable (eventual → strong)", "Multi-AZ (3 replicas)", "B-tree (custom)", "Fully managed; auto-scaling; single-digit ms latency SLA; DynamoDB Streams"],
        ["Apache Cassandra", "Tunable quorum (ONE/QUORUM/ALL)", "vnodes, rack-aware", "LSM-Tree (SSTables)", "CQL query language; 256 vnodes default; wide column model; Netflix, Apple scale"],
        ["Riak KV", "Eventual (AP)", "Consistent hash ring", "Bitcask / LevelDB", "CRDT support (counters, sets); automatic conflict merge; Erlang-based"],
        ["Redis Cluster", "Eventual (async replication)", "Primary-replica, 16384 hash slots", "In-memory (RDB/AOF)", "Pub/Sub; streams; 100K+ QPS/node; optional persistence; used for caching"],
        ["etcd", "Strong (Raft consensus)", "Raft log replication", "BoltDB / bbolt", "Linearizable reads; used in Kubernetes; watch API; ~10K writes/sec limit"],
        ["TiKV", "Strong (Raft per region)", "Multi-Raft groups", "RocksDB", "Used by TiDB; 2PC for transactions; 96MB region splits; CNCF graduated"],
        ["Aerospike", "Tunable (AP/CP modes)", "Cross-datacenter", "Hybrid (RAM+SSD)", "Direct SSD access; no OS page cache; sub-millisecond reads; financial services"],
        ["Voldemort (LinkedIn)", "Eventual", "Consistent hashing", "BDB / MySQL", "Original Dynamo implementation; inspired LinkedIn's feed infrastructure"],
    ]
    story.append(make_table(rw_data, [1.1*inch, 1.2*inch, 1.1*inch, 1.0*inch, 2.6*inch]))
    story.append(Spacer(1, 0.1*inch))

    story.append(Paragraph(
        "<b>Key insight:</b> Strongly consistent systems (etcd, TiKV) sacrifice throughput "
        "and partition availability for linearizability. AP systems (Cassandra, Riak, DynamoDB "
        "eventual mode) achieve much higher throughput by relaxing consistency guarantees. "
        "Most production systems choose tunable consistency to serve both use cases.",
        styles['Note']
    ))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 16: INTERVIEW Q&A
    # =========================================================================
    story.append(Paragraph("16. Interview Q&A", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph(
        "The following Q&A covers the most frequently asked interview questions about "
        "distributed key-value store design:", styles['BodyText']
    ))

    qa_pairs = [
        (
            "Q1: Walk me through the write path when a client calls put('user:42', data).",
            [
                "The client sends a PUT request to any node via the load balancer — that node becomes the coordinator.",
                "The coordinator hashes the key using SHA-256 mod 2^32 to find the token position on the ring.",
                "It walks clockwise to build the preference list of N=3 distinct physical nodes.",
                "The coordinator writes to its local WAL (append-only, sequential, fast) then inserts into MemTable.",
                "In parallel, it forwards the write to the other 2 replicas in the preference list.",
                "It waits for W=2 total ACKs (including its own). Once met, it returns 200 OK to the client with the new vector clock.",
                "If a replica was down, the coordinator stores a hint and delivers it when the replica recovers.",
                "Total latency: ~5ms p50, ~10ms p99 for a same-datacenter W=2 write.",
            ]
        ),
        (
            "Q2: 2 of your 5 nodes become unreachable. What happens?",
            [
                "The ring now has 3 live nodes (N1, N2, N3) and 2 isolated nodes (N4, N5).",
                "With W=2, R=2: the majority side (3 nodes) can still satisfy all quorum operations. Service continues normally for keys whose preference lists include at least 2 of N1/N2/N3.",
                "The minority side (N4, N5) cannot complete quorum writes to keys whose preference list includes N1/N2/N3. With sloppy quorum enabled, N4/N5 accept writes using each other and store hints.",
                "After the partition heals, the gossip protocol detects node recovery within ~10 seconds. Hinted handoff replays queued writes. Merkle tree anti-entropy reconciles any remaining divergence.",
                "The key trade-off: with sloppy quorum, we favor availability (AP). Without it, the minority side is unavailable for writes (CP behavior).",
            ]
        ),
        (
            "Q3: How do Merkle trees help with anti-entropy? Why not just compare all keys?",
            [
                "A brute-force comparison requires transferring all key hashes between replicas: O(N) messages for N keys. With 1 billion keys, that's gigabytes of hash data.",
                "A Merkle tree hashes key ranges hierarchically. The root hash summarizes the entire dataset. If two replicas have the same root hash, their data is identical — no further comparison needed.",
                "If root hashes differ, we recurse into subtrees: compare left child hashes, then right child hashes. At each level, we only recurse into differing subtrees.",
                "This requires only O(log N) messages to identify the diverged leaf nodes, even for billion-key datasets.",
                "Example: 1B keys in a depth-30 tree. If only 1% of keys differ, we'd identify the diverged ranges in ~30 comparisons instead of 1 billion.",
                "Cassandra uses exactly this approach: it builds Merkle trees over each VNode's token range and runs anti-entropy repair periodically.",
            ]
        ),
        (
            "Q4: Why is hinted handoff critical for availability?",
            [
                "Without hinted handoff, any node in the preference list being temporarily unavailable would cause writes to fail (if W cannot be met with the remaining nodes).",
                "With hinted handoff, the coordinator finds a healthy substitute node, writes the data there with metadata indicating the intended recipient, and returns success to the client.",
                "This means a single node restart, network blip, or rolling upgrade does not impact write availability — as long as W healthy nodes exist somewhere.",
                "The hints are durable (stored on disk) and survive coordinator crashes. When the target node recovers, hints are replayed in order.",
                "The trade-off: hints consume disk space on the substitute node. A hint TTL (e.g., 1 hour) prevents unbounded growth. If the node doesn't recover within the TTL, anti-entropy handles the divergence.",
                "Real-world: DynamoDB, Cassandra, and Riak all use hinted handoff. It is a foundational mechanism for AP distributed stores.",
            ]
        ),
        (
            "Q5: How would you implement secondary indexes on this system?",
            [
                "Option 1 — Local indexes: Each node indexes only its own data. A query fans out to all N nodes (scatter-gather), each returns matching keys, the coordinator merges and deduplicates results. Simple to implement, but O(N) read fan-out makes it expensive for large clusters.",
                "Option 2 — Global indexes: Dedicate a separate set of nodes (or a separate KV namespace) to store the inverted index (attribute_value → [primary_keys]). Writes must update both the primary store and the global index — either via 2PC (strong consistency, expensive) or async (eventual consistency, simpler).",
                "Hybrid approach: Use local indexes for low-cardinality attributes (e.g., status: active/inactive) and global indexes for high-cardinality attributes (e.g., email).",
                "DynamoDB's GSI (Global Secondary Index) maintains a separate partition of the data sorted by the indexed attribute. It uses asynchronous replication from the primary table, so GSI reads may be slightly stale.",
                "The key interview insight: there is no free lunch. Secondary indexes always trade write overhead (maintaining the index) for read performance (avoiding full scans).",
            ]
        ),
    ]

    for q, answers in qa_pairs:
        story.append(Paragraph(q, styles['SubHeader']))
        for i, ans in enumerate(answers, 1):
            story.append(Paragraph(f"{i}. {ans}", styles['NumberedItem']))
        story.append(Spacer(1, 0.1*inch))
        story.append(hr())

    story.append(PageBreak())

    # =========================================================================
    # SECTION 17: REFERENCES
    # =========================================================================
    story.append(Paragraph("17. References", styles['SectionHeader']))
    story.append(section_hr())

    story.append(Paragraph("Academic Papers", styles['SubHeader']))
    papers = [
        ("Decandia et al. (2007)", "Dynamo: Amazon's Highly Available Key-value Store", "ACM SIGOPS, 2007. The foundational paper describing consistent hashing, vector clocks, quorum consensus, hinted handoff, and Merkle tree anti-entropy as used in Amazon's production system."),
        ("Lakshman & Malik (2010)", "Cassandra: A Decentralized Structured Storage System", "ACM SIGOPS Operating Systems Review. Describes Cassandra's architecture: gossip, vnodes, tunable consistency, and SSTable storage."),
        ("Chang et al. (2006)", "Bigtable: A Distributed Storage System for Structured Data", "Google OSDI 2006. Describes the tablet-based architecture that influenced HBase, Cassandra's data model, and LSM-tree adoption."),
    ]
    for authors, title, desc in papers:
        story.append(Paragraph(f"<b>{authors}</b> — <i>{title}</i>", styles['BulletItem']))
        story.append(Paragraph(desc, ParagraphStyle('RefDesc', parent=styles['BodyText'], leftIndent=24, fontSize=9)))
        story.append(Spacer(1, 0.05*inch))

    story.append(Paragraph("Books", styles['SubHeader']))
    books = [
        ("Martin Kleppmann", "Designing Data-Intensive Applications (O'Reilly, 2017)", "Chapter 5 (Replication), Chapter 6 (Partitioning), Chapter 9 (Consistency & Consensus). Essential reading for distributed systems engineers."),
        ("Werner Vogels", "Eventually Consistent (ACM Queue, 2008)", "Werner Vogels (Amazon CTO) explains the practical implications of eventual consistency in production distributed systems."),
    ]
    for authors, title, desc in books:
        story.append(Paragraph(f"<b>{authors}</b> — <i>{title}</i>", styles['BulletItem']))
        story.append(Paragraph(desc, ParagraphStyle('RefDesc2', parent=styles['BodyText'], leftIndent=24, fontSize=9)))
        story.append(Spacer(1, 0.05*inch))

    story.append(Paragraph("Videos", styles['SubHeader']))
    videos = [
        ("Werner Vogels (AWS re:Invent)", "DynamoDB Internals — How we built a planet-scale database", "youtube.com — DynamoDB architecture deep dive from the CTO of Amazon."),
        ("DataStax Academy", "Cassandra Architecture — Understanding the Distributed Database", "academy.datastax.com — Covers ring topology, vnodes, gossip, and compaction with live demos."),
        ("MIT 6.824 (Raft)", "Distributed Systems Lectures — Raft Consensus Algorithm", "pdos.csail.mit.edu/6.824 — Comprehensive lecture series covering Raft, Zookeeper, and distributed transactions."),
    ]
    for author, title, url in videos:
        story.append(Paragraph(f"<b>{author}</b> — <i>{title}</i>", styles['BulletItem']))
        story.append(Paragraph(url, ParagraphStyle('RefURL', parent=styles['BodyText'], leftIndent=24, fontSize=9, textColor=ACCENT_BLUE)))
        story.append(Spacer(1, 0.05*inch))

    story.append(Paragraph("Blog Posts & Documentation", styles['SubHeader']))
    blogs = [
        ("AWS Blog", "DynamoDB under the hood — AWS re:Invent deep dive", "aws.amazon.com/blogs/database — Covers DynamoDB's internal partition management and adaptive capacity."),
        ("DataStax Blog", "How Cassandra handles writes — WAL, MemTable, and SSTables", "datastax.com/blog — Detailed walkthrough of Cassandra's write path with performance benchmarks."),
        ("Facebook Engineering", "RocksDB: A persistent key-value store for flash and RAM storage", "rocksdb.org — RocksDB documentation covering compaction strategies, bloom filters, and tuning."),
        ("Basho Technologies", "Riak Technical Overview — Consistent Hashing & Vector Clocks", "docs.riak.com — Explains Riak's CRDT support and eventual consistency model."),
    ]
    for author, title, url in blogs:
        story.append(Paragraph(f"<b>{author}</b> — <i>{title}</i>", styles['BulletItem']))
        story.append(Paragraph(url, ParagraphStyle('RefURL2', parent=styles['BodyText'], leftIndent=24, fontSize=9, textColor=ACCENT_BLUE)))
        story.append(Spacer(1, 0.05*inch))

    story.append(Spacer(1, 0.3*inch))
    story.append(section_hr())
    story.append(Paragraph(
        "This document was generated programmatically as a system design reference. "
        "All architectural decisions reflect industry best practices as of March 2026. "
        "For production deployments, consult the official documentation of the chosen "
        "implementation (Cassandra, DynamoDB, TiKV, etc.).",
        ParagraphStyle('Footer', parent=styles['BodyText'], textColor=MEDIUM_GRAY, fontSize=9, alignment=TA_CENTER)
    ))

    return story


def main():
    doc = SimpleDocTemplate(
        OUTPUT_PATH,
        pagesize=letter,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
        title="Distributed Key-Value Store: System Design",
        author="System Design Reference",
        subject="Dynamo-Style Distributed KV Store Architecture",
    )

    story = build_document()

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"PDF generated successfully: {OUTPUT_PATH}")
    size_bytes = os.path.getsize(OUTPUT_PATH)
    print(f"File size: {size_bytes / 1024:.1f} KB ({size_bytes / 1024 / 1024:.2f} MB)")


if __name__ == '__main__':
    main()
