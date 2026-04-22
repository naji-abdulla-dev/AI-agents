from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.platypus import Flowable
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import math


# ── Custom drawing flowables ──────────────────────────────────────────────────

class BinaryTreeDiagram(Flowable):
    """Draws a simple binary tree (7 nodes)."""
    W, H = 320, 180

    def wrap(self, *_):
        return self.W, self.H

    def draw(self):
        c = self.canv
        # node positions  (x, y)
        nodes = {
            1: (160, 155),
            2: (90,  105),
            3: (230, 105),
            4: (55,   55),
            5: (125,  55),
            6: (195,  55),
            7: (265,  55),
        }
        edges = [(1,2),(1,3),(2,4),(2,5),(3,6),(3,7)]
        r = 18

        # edges
        c.setStrokeColor(colors.HexColor("#4a6fa5"))
        c.setLineWidth(1.5)
        for p, ch in edges:
            x1,y1 = nodes[p]
            x2,y2 = nodes[ch]
            c.line(x1, y1-r, x2, y2+r)

        # nodes
        labels = {1:"A",2:"B",3:"C",4:"D",5:"E",6:"F",7:"G"}
        for nid,(x,y) in nodes.items():
            c.setFillColor(colors.HexColor("#4a6fa5"))
            c.setStrokeColor(colors.white)
            c.circle(x, y, r, fill=1)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 11)
            c.drawCentredString(x, y-4, labels[nid])


class BSTDiagram(Flowable):
    """Draws a BST with numeric keys."""
    W, H = 320, 180

    def wrap(self, *_):
        return self.W, self.H

    def draw(self):
        c = self.canv
        nodes = {
            50: (160, 155),
            30: (90,  105),
            70: (230, 105),
            20: (55,   55),
            40: (125,  55),
            60: (195,  55),
            80: (265,  55),
        }
        edges = [(50,30),(50,70),(30,20),(30,40),(70,60),(70,80)]
        r = 18

        c.setStrokeColor(colors.HexColor("#2e7d32"))
        c.setLineWidth(1.5)
        for p, ch in edges:
            x1,y1 = nodes[p]
            x2,y2 = nodes[ch]
            c.line(x1, y1-r, x2, y2+r)

        for key,(x,y) in nodes.items():
            c.setFillColor(colors.HexColor("#2e7d32"))
            c.setStrokeColor(colors.white)
            c.circle(x, y, r, fill=1)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 10)
            c.drawCentredString(x, y-4, str(key))

        # annotation arrows
        c.setFillColor(colors.HexColor("#555555"))
        c.setFont("Helvetica-Oblique", 8)
        c.drawString(6, 55, "< 30")
        c.drawString(130, 55, "> 30")


class BTreeDiagram(Flowable):
    """Draws a 3-level B-tree (order 3)."""
    W, H = 380, 200

    def _node(self, c, x, y, keys, fill_color):
        r = 14
        pad = 8
        w = len(keys) * (2*r + pad) + pad
        h = 32

        c.setFillColor(fill_color)
        c.setStrokeColor(colors.HexColor("#555555"))
        c.setLineWidth(1)
        c.roundRect(x - w/2, y - h/2, w, h, 6, fill=1, stroke=1)

        # dividers
        cell_w = w / len(keys)
        for i, k in enumerate(keys):
            cx = x - w/2 + cell_w*(i+0.5)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 10)
            c.drawCentredString(cx, y-4, str(k))
            if i > 0:
                c.setStrokeColor(colors.HexColor("#888888"))
                c.line(x - w/2 + cell_w*i, y-h/2+3, x - w/2 + cell_w*i, y+h/2-3)
        return x - w/2, y, w, h

    def wrap(self, *_):
        return self.W, self.H

    def draw(self):
        c = self.canv
        root_color    = colors.HexColor("#7b1fa2")
        level1_color  = colors.HexColor("#1565c0")
        leaf_color    = colors.HexColor("#00695c")

        # root
        self._node(c, 190, 175, [30, 60], root_color)

        # level 1
        self._node(c, 80,  115, [10, 20], level1_color)
        self._node(c, 190, 115, [40, 50], level1_color)
        self._node(c, 310, 115, [70, 80], level1_color)

        # edges root -> l1
        c.setStrokeColor(colors.HexColor("#555555"))
        c.setLineWidth(1.2)
        for lx in [80, 190, 310]:
            c.line(190, 159, lx, 131)

        # leaves
        leaves = [(45, 55), (115, 55), (165, 55), (215, 55), (275, 55), (345, 55)]
        leaf_keys = [[5,8], [15,18], [35,38], [45,48], [65,68], [75,78]]
        parents_x = [80, 80, 190, 190, 310, 310]

        for i,(lx,ly) in enumerate(leaves):
            self._node(c, lx, ly, leaf_keys[i], leaf_color)
            c.setStrokeColor(colors.HexColor("#555555"))
            c.line(parents_x[i], 99, lx, 71)

        # legend
        c.setFont("Helvetica", 7)
        c.setFillColor(colors.HexColor("#333333"))
        c.drawString(4, 8, "Each node holds multiple keys; children = keys+1")


# ── Document ──────────────────────────────────────────────────────────────────

def build():
    doc = SimpleDocTemplate(
        "tree_data_structures.pdf",
        pagesize=A4,
        rightMargin=2*cm, leftMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("title", parent=styles["Title"],
                                 fontSize=22, spaceAfter=6,
                                 textColor=colors.HexColor("#1a237e"))
    subtitle_style = ParagraphStyle("sub", parent=styles["Normal"],
                                    fontSize=11, textColor=colors.HexColor("#555555"),
                                    spaceAfter=18, alignment=TA_CENTER)
    h2 = ParagraphStyle("h2", parent=styles["Heading2"],
                         fontSize=14, textColor=colors.HexColor("#1a237e"),
                         spaceBefore=14, spaceAfter=6)
    h3 = ParagraphStyle("h3", parent=styles["Heading3"],
                         fontSize=11, textColor=colors.HexColor("#37474f"),
                         spaceBefore=8, spaceAfter=4)
    body = ParagraphStyle("body", parent=styles["Normal"],
                          fontSize=10, leading=14, spaceAfter=4)
    caption = ParagraphStyle("cap", parent=styles["Normal"],
                              fontSize=9, textColor=colors.HexColor("#666666"),
                              alignment=TA_CENTER, spaceAfter=10)
    code_style = ParagraphStyle("code", parent=styles["Code"],
                                fontSize=8.5, backColor=colors.HexColor("#f5f5f5"),
                                textColor=colors.HexColor("#212121"),
                                leftIndent=10, rightIndent=10,
                                borderPadding=6, spaceBefore=4, spaceAfter=8)

    story = []

    # ── Title ─────────────────────────────────────────────────────────────────
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("Tree Data Structures", title_style))
    story.append(Paragraph("Binary Tree · Binary Search Tree · B-Tree", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=1.5,
                             color=colors.HexColor("#1a237e")))
    story.append(Spacer(1, 0.5*cm))

    # ══════════════════════════════════════════════════════════════════════════
    # 1. Binary Tree
    # ══════════════════════════════════════════════════════════════════════════
    story.append(Paragraph("1. Binary Tree", h2))
    story.append(Paragraph(
        "A <b>binary tree</b> is a hierarchical structure where each node has "
        "<b>at most two children</b>, called the <i>left</i> and <i>right</i> child. "
        "There are no ordering constraints — values can be arranged arbitrarily.",
        body))

    story.append(Spacer(1, 0.3*cm))
    story.append(BinaryTreeDiagram())
    story.append(Paragraph("Figure 1 — A generic binary tree with 7 nodes", caption))

    story.append(Paragraph("Key Characteristics", h3))
    chars = [
        ["Property", "Value"],
        ["Max children per node", "2 (left, right)"],
        ["Ordering constraint", "None"],
        ["Height (balanced)", "O(log n)"],
        ["Height (worst case)", "O(n) — degenerate / skewed"],
        ["Search time", "O(n) — must traverse all nodes"],
        ["Insert / Delete", "O(1) at known position, O(n) to find position"],
    ]
    t = Table(chars, colWidths=[8*cm, 7*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), colors.HexColor("#1a237e")),
        ("TEXTCOLOR",   (0,0), (-1,0), colors.white),
        ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.HexColor("#f0f4ff"), colors.white]),
        ("GRID",        (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
        ("TOPPADDING",  (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
    ]))
    story.append(t)

    story.append(Paragraph("Common Uses", h3))
    story.append(Paragraph(
        "• Expression trees (compilers, calculators)<br/>"
        "• Hierarchical data (file systems, DOM)<br/>"
        "• Foundation for more specialised trees (BST, heaps, tries)",
        body))

    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#cccccc")))

    # ══════════════════════════════════════════════════════════════════════════
    # 2. Binary Search Tree
    # ══════════════════════════════════════════════════════════════════════════
    story.append(Paragraph("2. Binary Search Tree (BST)", h2))
    story.append(Paragraph(
        "A BST is a binary tree with an <b>ordering invariant</b>: for every node N, "
        "all keys in the <i>left</i> subtree are <b>less than</b> N's key, and all keys "
        "in the <i>right</i> subtree are <b>greater than</b> N's key.",
        body))

    story.append(Spacer(1, 0.3*cm))
    story.append(BSTDiagram())
    story.append(Paragraph(
        "Figure 2 — BST with root 50. Left subtree ≤ 50, right subtree ≥ 50 (recursively)",
        caption))

    story.append(Paragraph("Key Characteristics", h3))
    chars2 = [
        ["Property", "Value"],
        ["Max children per node", "2"],
        ["Ordering constraint", "left < node < right"],
        ["Search / Insert / Delete (balanced)", "O(log n)"],
        ["Search / Insert / Delete (worst case)", "O(n) — sorted input"],
        ["In-order traversal", "Yields sorted sequence"],
        ["Self-balancing variants", "AVL tree, Red-Black tree"],
    ]
    t2 = Table(chars2, colWidths=[8*cm, 7*cm])
    t2.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), colors.HexColor("#1b5e20")),
        ("TEXTCOLOR",   (0,0), (-1,0), colors.white),
        ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.HexColor("#f1f8e9"), colors.white]),
        ("GRID",        (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
        ("TOPPADDING",  (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
    ]))
    story.append(t2)

    story.append(Paragraph("Common Uses", h3))
    story.append(Paragraph(
        "• Dynamic sorted sets and maps (std::map in C++, TreeMap in Java)<br/>"
        "• Range queries<br/>"
        "• Auto-complete / ordered dictionary lookups",
        body))

    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#cccccc")))

    # ══════════════════════════════════════════════════════════════════════════
    # 3. B-Tree
    # ══════════════════════════════════════════════════════════════════════════
    story.append(Paragraph("3. B-Tree", h2))
    story.append(Paragraph(
        "A <b>B-tree of order m</b> is a self-balancing search tree in which each node "
        "can hold <b>up to m−1 keys</b> and have <b>up to m children</b>. All leaves "
        "are at the same depth. Designed to minimise disk I/O by maximising branching factor.",
        body))

    story.append(Spacer(1, 0.3*cm))
    story.append(BTreeDiagram())
    story.append(Paragraph(
        "Figure 3 — B-tree (order 3). Each internal node holds 2 keys and 3 child pointers",
        caption))

    story.append(Paragraph("Key Characteristics", h3))
    chars3 = [
        ["Property", "Value"],
        ["Children per node", "⌈m/2⌉ … m  (root: 2 … m)"],
        ["Keys per node", "⌈m/2⌉−1 … m−1"],
        ["Ordering constraint", "Generalised BST ordering across keys"],
        ["Height", "O(log_m n) — very shallow"],
        ["Search / Insert / Delete", "O(log n)"],
        ["All leaves", "Same depth (perfectly balanced)"],
        ["Disk pages", "One node ≈ one disk block → minimal I/O"],
    ]
    t3 = Table(chars3, colWidths=[8*cm, 7*cm])
    t3.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), colors.HexColor("#4a148c")),
        ("TEXTCOLOR",   (0,0), (-1,0), colors.white),
        ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.HexColor("#f3e5f5"), colors.white]),
        ("GRID",        (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
        ("TOPPADDING",  (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
    ]))
    story.append(t3)

    story.append(Paragraph("Common Uses", h3))
    story.append(Paragraph(
        "• Database indexes (MySQL InnoDB, PostgreSQL)<br/>"
        "• File systems (NTFS, HFS+, ext4 with htree)<br/>"
        "• Key-value stores (RocksDB, LevelDB)",
        body))

    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=1.5,
                             color=colors.HexColor("#1a237e")))

    # ══════════════════════════════════════════════════════════════════════════
    # Comparison Table
    # ══════════════════════════════════════════════════════════════════════════
    story.append(Paragraph("Side-by-Side Comparison", h2))

    cmp = [
        ["Feature", "Binary Tree", "BST", "B-Tree"],
        ["Node children", "≤ 2", "≤ 2", "⌈m/2⌉ … m"],
        ["Keys per node", "1", "1", "⌈m/2⌉−1 … m−1"],
        ["Order invariant", "None", "left<node<right", "Generalised BST"],
        ["Always balanced?", "No", "No", "Yes"],
        ["Search", "O(n)", "O(log n) avg", "O(log n)"],
        ["Insert", "O(1)*", "O(log n) avg", "O(log n)"],
        ["Delete", "O(n)", "O(log n) avg", "O(log n)"],
        ["Memory layout", "Pointer-based", "Pointer-based", "Block/page-friendly"],
        ["Primary use", "General hierarchy", "In-memory sets/maps", "Disk-based indexes"],
        ["Variants", "Heap, Trie, …", "AVL, Red-Black", "B+ tree, B* tree"],
    ]
    tc = Table(cmp, colWidths=[4.2*cm, 3.8*cm, 3.8*cm, 3.8*cm])
    tc.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0), colors.HexColor("#1a237e")),
        ("TEXTCOLOR",    (0,0), (-1,0), colors.white),
        ("FONTNAME",     (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",     (0,0), (-1,-1), 8.5),
        ("FONTNAME",     (0,1), (0,-1), "Helvetica-Bold"),
        ("TEXTCOLOR",    (0,1), (0,-1), colors.HexColor("#1a237e")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1),
            [colors.HexColor("#f0f4ff"), colors.white]),
        ("GRID",         (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
        ("TOPPADDING",   (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("ALIGN",        (0,0), (-1,-1), "CENTER"),
        ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
    ]))
    story.append(tc)
    story.append(Paragraph("* O(1) if insertion point is already known", caption))

    # ══════════════════════════════════════════════════════════════════════════
    # When to use what
    # ══════════════════════════════════════════════════════════════════════════
    story.append(Paragraph("When to Use Which?", h2))
    when = [
        ["Scenario", "Recommended"],
        ["Represent any hierarchy (no ordering needed)", "Binary Tree"],
        ["In-memory sorted collection, small dataset", "BST (or AVL/RB variant)"],
        ["In-memory sorted collection, need guarantees", "AVL / Red-Black Tree"],
        ["Database or file-system index", "B-Tree / B+ Tree"],
        ["Large dataset that doesn't fit in RAM", "B-Tree"],
        ["Prefix/string lookups", "Trie (specialised tree)"],
    ]
    tw = Table(when, colWidths=[10*cm, 5.6*cm])
    tw.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0), colors.HexColor("#37474f")),
        ("TEXTCOLOR",    (0,0), (-1,0), colors.white),
        ("FONTNAME",     (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",     (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS", (0,1), (-1,-1),
            [colors.HexColor("#fafafa"), colors.white]),
        ("GRID",         (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
        ("TOPPADDING",   (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("FONTNAME",     (1,1), (1,-1), "Helvetica-Bold"),
        ("TEXTCOLOR",    (1,1), (1,-1), colors.HexColor("#1b5e20")),
    ]))
    story.append(tw)

    story.append(Spacer(1, 1*cm))
    story.append(Paragraph(
        "Generated 2026-03-23 · Tree Data Structures Reference",
        ParagraphStyle("footer", parent=styles["Normal"],
                       fontSize=8, textColor=colors.HexColor("#aaaaaa"),
                       alignment=TA_CENTER)))

    doc.build(story)
    print("PDF written → tree_data_structures.pdf")


if __name__ == "__main__":
    build()
