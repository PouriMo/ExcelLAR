"""Generates the updated ExcelLAR workflow diagram as SVG (+ HTML wrapper for rendering)."""
import html
import re
import sys

OUT = sys.argv[1]          # output path without extension
SUBS = "abs"               # matrix subscript used in the paper ("abs" as in the old figure, or "ass" as in the code)

W, H = 1860, 2440
SANS = "Arial, 'Liberation Sans', Helvetica, sans-serif"
MONO = "Consolas, 'Liberation Mono', 'Courier New', monospace"
parts = []


def markup(s, fs):
    """Escape, then: _{x} -> subscript, **w** -> bold."""
    s = html.escape(s, quote=False)
    d = round(0.32 * fs)
    s = re.sub(r"_\{([^}]*)\}",
               lambda m: f'<tspan dy="{d}" font-size="{round(fs * 0.72)}">{m.group(1)}</tspan>'
                         f'<tspan dy="-{d}">&#8203;</tspan>', s)
    s = re.sub(r"\*\*([^*]+)\*\*", r'<tspan font-weight="bold">\1</tspan>', s)
    return s


def text(x, y, s, fs=29, bold=False, anchor="start", font=SANS, fill="#000"):
    fw = "bold" if bold else "normal"
    parts.append(f'<text x="{x}" y="{y}" font-family="{font}" font-size="{fs}" font-weight="{fw}" '
                 f'text-anchor="{anchor}" fill="{fill}" xml:space="preserve" style="white-space:pre">'
                 f'{markup(s, fs)}</text>')


def box(x, y, w, h, rx=30, dash=False, sw=3):
    da = ' stroke-dasharray="16 12"' if dash else ""
    parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="#fff" '
                 f'stroke="#000" stroke-width="{sw}"{da}/>')


def line(pts, dashed=False, head=True, both=False):
    d = "M" + " L".join(f"{x},{y}" for x, y in pts)
    col = "#555" if dashed else "#000"
    st = f'stroke="{col}" stroke-width="3" fill="none"'
    if dashed:
        st += ' stroke-dasharray="14 10"'
    mk = "ahg" if dashed else "ah"
    if head:
        st += f' marker-end="url(#{mk})"'
    if both:
        st += f' marker-start="url(#{mk}s)"'
    parts.append(f'<path d="{d}" {st}/>')


def module(x, y, w, h, title, bullets, tfs=30, bfs=29, gap=58):
    box(x, y, w, h)
    text(x + 40, y + 45, title, tfs, bold=True)
    for k, b in enumerate(bullets):
        yy = y + 45 + 62 + k * gap
        text(x + 45, yy, "►", 22)
        text(x + 80, yy, b, bfs)


def algorithm(x, y, w, title, code, lh=30, fs=26):
    h = 66 + 40 + len(code) * lh + 30
    box(x, y, w, h)
    parts.append(f'<rect x="{x + 30}" y="{y + 16}" width="{w - 60}" height="66" fill="#f2f2f2"/>')
    text(x + 47, y + 60, title, 30, bold=True)
    for k, c in enumerate(code):
        text(x + 47, y + 66 + 16 + 46 + k * lh, c, fs, font=MONO)
    return h


# ---------------------------------------------------------------- defs
parts.append(f'''<defs>
<marker id="ah" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#000"/></marker>
<marker id="ahs" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#000"/></marker>
<marker id="ahg" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#555"/></marker>
<linearGradient id="acad" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#ff3d5f"/><stop offset="1" stop-color="#c8102e"/></linearGradient>
</defs>
<rect width="{W}" height="{H}" fill="#fff"/>''')

# ---------------------------------------------------------------- header (AutoCAD | CSV)
parts.append('<rect x="18" y="28" width="686" height="110" fill="#fff" stroke="#000" stroke-width="3"/>')
parts.append('<line x1="366" y1="28" x2="366" y2="138" stroke="#000" stroke-width="3"/>')
# AutoCAD-like icon
parts.append('<path d="M44,44 L104,44 Q110,44 110,50 L110,120 L44,120 Q38,120 38,114 L38,50 Q38,44 44,44 z" fill="url(#acad)"/>')
parts.append('<path d="M38,52 L52,44 L52,112 L38,120 z" fill="#e0204a" opacity="0.8"/>')
parts.append('<rect x="44" y="99" width="66" height="22" fill="#7a0a1e"/>')
text(78, 90, "A", 44, bold=True, anchor="middle", fill="#fff")
text(77, 116, "CAD", 16, bold=True, anchor="middle", fill="#fff")
text(143, 77, "Modeling:", 30, bold=True)
text(146, 110, "AutoCAD", 30)
# Excel-like icon
parts.append('<rect x="404" y="42" width="68" height="80" rx="6" fill="#21a366"/>')
parts.append('<path d="M410,42 L466,42 Q472,42 472,48 L472,62 L404,62 L404,48 Q404,42 410,42 z" fill="#33c481"/>')
parts.append('<path d="M404,102 L472,102 L472,116 Q472,122 466,122 L410,122 Q404,122 404,116 z" fill="#185c37"/>')
parts.append('<rect x="438" y="62" width="34" height="40" fill="#107c41"/>')
parts.append('<rect x="384" y="58" width="48" height="48" rx="5" fill="#107c41"/>')
text(408, 95, "X", 34, bold=True, anchor="middle", fill="#fff")
text(486, 74, "Model output:", 30, bold=True)
text(490, 106, "CSV Data File", 30)

CX = 366   # main column centre

# ---------------------------------------------------------------- main column
line([(CX, 138), (CX, 183)])
module(10, 185, 712, 205, "USER INTERFACE: ExcelLAR Form",
       ["Import CSV into the workbook",
        "Validate interface data & live load"])
line([(CX, 390), (CX, 433)])
module(10, 435, 712, 210, "MODULE 1 & 2: Pre-processing",
       ["Read block geometry & locate interfaces",
        f"Assemble G_{{{SUBS}_in}} and G_{{{SUBS}}} matrices"])
line([(CX, 645), (CX, 693)])
module(30, 695, 672, 272, "MODULE 3 & 4: Variable Reduction",
       [f"Assemble load vectors F_{{{SUBS},W}} and F_{{{SUBS},λ}}",
        "Gauss elimination with partial pivoting",
        "Compute reduced constraints A_{in} and b_{in}"])
line([(CX, 967), (CX, 1015)])
module(70, 1017, 590, 150, "MODULE 5: Optimization Setup",
       ["Apply Row Normalization Algorithm"])
line([(CX, 1167), (CX, 1215)])

# Module 6 group
box(90, 1217, 552, 610, rx=0, dash=True)
text(CX, 1275, "MODULE 6: Solver", 38, bold=True, anchor="middle")
box(236, 1310, 260, 102, rx=28)
text(CX, 1352, "Lower Bound", 30, bold=True, anchor="middle")
text(CX, 1388, "Limit Analysis", 30, bold=True, anchor="middle")
line([(CX, 1415), (CX, 1477)], both=True)
box(190, 1480, 352, 102, rx=28)
text(CX, 1522, "LP Solver:", 30, bold=True, anchor="middle")
text(CX, 1558, "OpenSolver (max λ)", 29, anchor="middle")
line([(CX, 1582), (CX, 1650)])
box(120, 1652, 492, 140, rx=28)
text(CX, 1696, "Back-substitution:", 30, bold=True, anchor="middle")
text(CX, 1734, "N, V, M at all interfaces", 29, anchor="middle")
text(CX, 1772, "e_{i} = M_{i} / N_{i}", 29, anchor="middle")
line([(CX, 1827), (CX, 1873)])

module(30, 1875, 672, 280, "MODULE 9: Collapse Mechanism",
       ["Hinge candidates where |e_{i}| ≈ t_{i}/2",
        "Test every 4-hinge set kinematically",
        "Keep smallest λ_{kin}; check against λ"])

# ---------------------------------------------------------------- decision + outputs
DX, DY = 1000, 2015
line([(702, DY), (DX - 192, DY)])
parts.append(f'<polygon points="{DX - 190},{DY} {DX},{DY - 88} {DX + 190},{DY} {DX},{DY + 88}" '
             f'fill="#fff" stroke="#000" stroke-width="3"/>')
text(DX, DY - 6, "Admissible", 30, bold=True, anchor="middle")
text(DX, DY + 28, "mechanism?", 30, bold=True, anchor="middle")

# False -> warning
line([(DX, DY - 88), (DX, 1832)])
text(DX - 18, 1895, "False", 34, bold=True, anchor="end")
box(850, 1680, 300, 150, rx=28)
text(DX, 1722, "Warning logged", 29, bold=True, anchor="middle")
text(DX, 1760, "Hinges = 0", 28, anchor="middle")
text(DX, 1797, "(no deformed shape)", 26, anchor="middle")

# True -> visualisation
line([(DX + 190, DY), (1238, DY)])
text(DX + 150, DY - 62, "True", 34, bold=True, anchor="middle")
module(1240, 1875, 610, 345, "MODULE 7, 8 & 10: Visualization",
       ["Plot thrust line & hinges (M7)",
        "Rotate the three rigid links (M8)",
        "Plot deformed collapse shape (M8)",
        "Equal-aspect chart axes (M10)"], gap=57)

# ---------------------------------------------------------------- algorithm panels (right column)
AX, AW = 868, 886
algorithm(AX, 210, AW, "Algorithm: Row Normalization (Module 5)", [
    "**for** i = 1 **to** nRows",
    "  rowMax = |b(i)|",
    "  **for** j = 1 **to** nCols          (N, V, M, λ)",
    "    **if** |A(i,j)| > rowMax **then** rowMax = |A(i,j)|",
    "  **end**",
    "  **if** rowMax > 1e-12 **then**",
    "    **for** j = 1 **to** nCols",
    "      A(i,j) = A(i,j) / rowMax",
    "    **end**",
    "    b(i) = b(i) / rowMax",
    "  **end**",
    "**end**",
])
line([(660, 1092), (748, 1092), (748, 445), (866, 445)], dashed=True)

algorithm(AX, 740, AW, "Algorithm: Collapse Mechanism (Module 9)", [
    "cand = {i : |e_{i}|/(t_{i}/2) ≥ 0.98}  (≤3 per run)",
    "     ∪ local peaks with ratio ≥ 0.5 (≤30)",
    "λ_{best} = ∞",
    "**for each** {h_{1}<h_{2}<h_{3}<h_{4}} **in** cand",
    "  solve closure for link rotations ω_{1..3}",
    "  φ_{k} = relative rotation at hinge k",
    "  **if** sign(s_{k}·φ_{k}) differs **then** skip",
    "  λ_{kin} = −W_{dead} / W_{live}",
    "  **if** λ_{kin} < λ_{best} **then** keep set",
    "**end**",
    "**if** none found **then** repeat with",
    "  the 40 largest ratios",
    "check |λ_{kin} − λ| ≤ 1% · λ",
])
line([(702, 1920), (800, 1920), (800, 1110), (866, 1110)], dashed=True)

# ---------------------------------------------------------------- supporting modules strip
box(10, 2262, 1840, 160, rx=24, dash=True)
text(40, 2306, "SUPPORTING MODULES", 28, bold=True)
items = [
    ("LinearSolver", "Gauss elimination (M4, M6)"),
    ("ExcelLAR.frm","user form, Run All, export"),
    ("Module_Launcher", "opens the form"),
    ("ModuleImage", "PNG export & preview panels"),
    ("Module 0", "Solver add-in loader"),
    ("Module 9 utilities", "CSV reader, edge geometry"),
]
for k, (n, d) in enumerate(items):
    col, row = k % 3, k // 3
    x, y = 40 + col * 610, 2352 + row * 44
    text(x, y, "►", 20)
    text(x + 30, y, f"**{n}**: {d}", 25)

svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">\n'
       + "\n".join(parts) + "\n</svg>\n")
with open(OUT + ".svg", "w", encoding="utf-8") as f:
    f.write(svg)
with open(OUT + ".html", "w", encoding="utf-8") as f:
    f.write('<!doctype html><html><head><meta charset="utf-8"><style>html,body{margin:0;padding:0;'
            'background:#fff;overflow:hidden}</style></head><body>' + svg + "</body></html>")
print("ok")
