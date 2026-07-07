#!/usr/bin/env python3
"""Generate the RIA v1 pitch deck (refined concept) as a .pptx file."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------------------------------------------------------------- palette
NAVY = RGBColor(0x0F, 0x24, 0x42)
NAVY2 = RGBColor(0x1B, 0x3A, 0x5F)
ACCENT = RGBColor(0x2E, 0x86, 0xAB)
ACCENT_LT = RGBColor(0xDD, 0xEE, 0xF4)
INK = RGBColor(0x1F, 0x29, 0x37)
GRAY = RGBColor(0x5B, 0x66, 0x76)
LIGHT = RGBColor(0xF2, 0xF5, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOOD = RGBColor(0x1E, 0x7A, 0x4C)
WARN = RGBColor(0xA8, 0x54, 0x1C)
CARD_LINE = RGBColor(0xD4, 0xDC, 0xE4)

FONT = "Calibri"
SW, SH = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.6)

prs = Presentation()
prs.slide_width = SW
prs.slide_height = SH
BLANK = prs.slide_layouts[6]


# ---------------------------------------------------------------- helpers
def slide():
    return prs.slides.add_slide(BLANK)


def rect(s, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE,
         shadow=False):
    sp = s.shapes.add_shape(shape, x, y, w, h)
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(1)
    sp.shadow.inherit = False
    return sp


def text_in(sp, items, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
            inset=0.08):
    """items: list of (text, size, bold, color, space_after) tuples."""
    tf = sp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, Inches(inset))
    first = True
    for (t, size, bold, color, after) in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(after)
        r = p.add_run()
        r.text = t
        f = r.font
        f.name = FONT
        f.size = Pt(size)
        f.bold = bold
        f.color.rgb = color
    return sp


def tbox(s, x, y, w, h, items, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    sp = s.shapes.add_textbox(x, y, w, h)
    return text_in(sp, items, align=align, anchor=anchor, inset=0.02)


def chrome(s, num, title, kicker=None):
    """Standard content-slide header + footer."""
    rect(s, 0, 0, SW, SH, fill=WHITE)
    y = Inches(0.42)
    if kicker:
        tbox(s, MARGIN, Inches(0.30), Inches(9), Inches(0.3),
             [(kicker.upper(), 11, True, ACCENT, 0)])
        y = Inches(0.60)
    tbox(s, MARGIN, y, Inches(12.13), Inches(0.6),
         [(title, 26, True, NAVY, 0)])
    rect(s, MARGIN, y + Inches(0.58), Inches(1.5), Inches(0.045), fill=ACCENT)
    tbox(s, SW - Inches(1.4), SH - Inches(0.42), Inches(0.9), Inches(0.3),
         [(f"{num:02d}", 10, False, GRAY, 0)], align=PP_ALIGN.RIGHT)
    tbox(s, MARGIN, SH - Inches(0.42), Inches(6), Inches(0.3),
         [("RIA — Reconfigurable Intelligence Architecture", 9, False,
           GRAY, 0)])


def bullets(s, x, y, w, h, rows, size=14, gap=10, lead_color=INK):
    """rows: list of (lead, rest) or plain string."""
    sp = s.shapes.add_textbox(x, y, w, h)
    tf = sp.text_frame
    tf.word_wrap = True
    first = True
    for row in rows:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(gap)
        b = p.add_run()
        b.text = "▪  "
        b.font.name = FONT
        b.font.size = Pt(size)
        b.font.color.rgb = ACCENT
        if isinstance(row, tuple):
            lead, rest = row
            r1 = p.add_run()
            r1.text = lead + " — " if rest else lead
            r1.font.name = FONT
            r1.font.size = Pt(size)
            r1.font.bold = True
            r1.font.color.rgb = lead_color
            if rest:
                r2 = p.add_run()
                r2.text = rest
                r2.font.name = FONT
                r2.font.size = Pt(size)
                r2.font.color.rgb = GRAY
        else:
            r1 = p.add_run()
            r1.text = row
            r1.font.name = FONT
            r1.font.size = Pt(size)
            r1.font.color.rgb = INK
    return sp


def card(s, x, y, w, h, head, body, head_color=NAVY, fill=LIGHT,
         head_size=15, body_size=12.5):
    c = rect(s, x, y, w, h, fill=fill, line=CARD_LINE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    c.adjustments[0] = 0.055
    items = [(head, head_size, True, head_color, 6)]
    items.append((" ".join(body), body_size, False, GRAY, 0))
    text_in(c, items, inset=0.16)
    return c


def table(s, x, y, w, col_w, rows, header_fill=NAVY, size=11.5,
          row_h=Inches(0.34), header_h=Inches(0.38)):
    """rows[0] is the header. col_w: list of column width fractions."""
    n_r, n_c = len(rows), len(rows[0])
    g = s.shapes.add_table(n_r, n_c, x, y, w, Inches(0.3)).table
    for i, frac in enumerate(col_w):
        g.columns[i].width = int(w * frac)
    g.rows[0].height = header_h
    for i in range(1, n_r):
        g.rows[i].height = row_h
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = g.cell(ri, ci)
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.06)
            cell.margin_top = Inches(0.03)
            cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if ri == 0:
                cell.fill.fore_color.rgb = header_fill
            else:
                cell.fill.fore_color.rgb = WHITE if ri % 2 else LIGHT
            p = cell.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = str(val)
            f = r.font
            f.name = FONT
            f.size = Pt(size + (0.5 if ri == 0 else 0))
            f.bold = ri == 0 or ci == 0
            f.color.rgb = WHITE if ri == 0 else (NAVY if ci == 0 else INK)
    return g


# ============================================================ S1 — TITLE
s = slide()
rect(s, 0, 0, SW, SH, fill=NAVY)
rect(s, 0, SH - Inches(1.9), SW, Inches(0.06), fill=ACCENT)
tbox(s, MARGIN, Inches(2.1), Inches(12.1), Inches(2.2), [
    ("RIA", 66, True, WHITE, 2),
    ("Reconfigurable Intelligence Architecture", 30, False, ACCENT_LT, 14),
    ("One tile. Any model. Compiler-shaped silicon.", 20, True, ACCENT, 0),
])
tbox(s, MARGIN, Inches(5.9), Inches(12.1), Inches(1.0), [
    ("Licensable edge-AI accelerator IP  •  FPGA-proven reference  •  "
     "ASIC-ready RTL", 14, False, ACCENT_LT, 6),
    ("Concept v1  •  July 2026", 12, False, RGBColor(0x8F, 0xA6, 0xBD), 0),
])

# ==================================================== S2 — THE PROBLEM
s = slide()
chrome(s, 2, "Edge AI silicon is stuck between three bad choices",
       "The problem")
cy, ch, cw, gap = Inches(1.55), Inches(2.5), Inches(3.94), Inches(0.2)
card(s, MARGIN, cy, cw, ch, "Fixed-function NPUs",
     ["Peak efficiency — on yesterday's models.",
      "A CNN accelerator taped out in 2022 cannot serve",
      "a transformer workload in 2025.",
      "Silicon is frozen; model architectures are not."])
card(s, MARGIN + cw + gap, cy, cw, ch, "GPUs",
     ["Fully programmable, but the power budget",
      "of the edge (mW–single-digit W) rules them out.",
      "Utilization on small batch-1 inference is poor."])
card(s, MARGIN + 2 * (cw + gap), cy, cw, ch, "FPGAs",
     ["Flexible, but every model change means hours of",
      "synthesis and place-and-route in third-party tools,",
      "and demands hardware expertise most",
      "AI teams do not have."])
b = rect(s, MARGIN, Inches(4.35), Inches(12.13), Inches(1.15),
         fill=ACCENT_LT, line=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
b.adjustments[0] = 0.12
text_in(b, [
    ("Underneath all three: memory bandwidth dominates energy, and "
     "utilization collapses the moment the workload drifts from the "
     "silicon's design-target model.", 14, True, NAVY, 4),
    ("Model architectures turn over every ~18 months. Silicon takes 24+ "
     "months to ship. Programmability must live above the silicon.",
     13, False, GRAY, 0)], inset=0.18)
tbox(s, MARGIN, Inches(5.75), Inches(12.1), Inches(0.5),
     [("The gap: a licensable accelerator that stays efficient as models "
       "evolve — without re-synthesis, and without a hardware team on the "
       "customer side.", 14, True, ACCENT, 0)])

# ============================================ S3 — DESIGN SPACE (journey)
s = slide()
chrome(s, 3, "We explored both extremes — and rejected them",
       "Why this shape")
rows = [
    ["Approach", "What it gives", "Why it fails", "Verdict"],
    ["15 specialized tiles\n(heterogeneous NPU)",
     "Each operator gets ideal hardware",
     "Dark silicon: every model idles the tiles it doesn't use; "
     "15x design & verification effort",
     "Rejected"],
    ["Fine-grained fabric\n(FPGA-style blocks)",
     "Perfect per-model specialization",
     "Hours of P&R per model in tools we don't own; fine-grained "
     "reconfigurability doesn't survive ASIC hardening (eFPGA: 10–20x "
     "area/power)",
     "Rejected"],
    ["One programmable tile, replicated\n(RIA)",
     "High utilization on every model family; one tile to verify; "
     "compiles in seconds",
     "Efficiency gap vs. fixed function must be engineered down — "
     "measured, not assumed",
     "Chosen"],
]
g = table(s, MARGIN, Inches(1.6), Inches(12.13),
          [0.24, 0.27, 0.37, 0.12], rows, size=11.5,
          row_h=Inches(1.05), header_h=Inches(0.4))
for ri, color in ((1, WARN), (2, WARN), (3, GOOD)):
    cell = g.cell(ri, 3)
    cell.text_frame.paragraphs[0].runs[0].font.color.rgb = color
    cell.text_frame.paragraphs[0].runs[0].font.bold = True
tbox(s, MARGIN, Inches(5.6), Inches(12.1), Inches(1.2), [
    ("The durable middle: hardened, identical tiles whose function is "
     "assigned by software.", 16, True, NAVY, 6),
    ("Reconfiguration is instructions and configuration registers — not "
     "rewiring. It survives the FPGA-to-ASIC transition intact.",
     13.5, False, GRAY, 0)])

# ================================================== S4 — WHAT IS RIA
s = slide()
chrome(s, 4, "What is RIA?", "The concept")
tbox(s, MARGIN, Inches(1.6), Inches(12.1), Inches(1.5), [
    ("RIA is a homogeneous array of programmable compute tiles, a "
     "microcoded scheduler, and an MLIR-based compiler that together "
     "execute any ML workload on one hardware platform.",
     17, True, INK, 8),
    ("The compiler — not the silicon — decides how the hardware is shaped: "
     "the tile mix at customer-integration time, and each tile's function "
     "at runtime, layer by layer.", 14.5, False, GRAY, 0)])
cy, cw, ch, gap = Inches(3.3), Inches(3.94), Inches(2.25), Inches(0.2)
card(s, MARGIN, cy, cw, ch, "One tile, replicated N times",
     ["A single verified compute tile is stamped",
      "across the die. More performance = more tiles.",
      "One design, one verification effort,",
      "natural yield redundancy."])
card(s, MARGIN + cw + gap, cy, cw, ch, "Personality by instruction",
     ["A tile is a matrix engine this layer and a",
      "softmax engine the next — selected by microcode,",
      "swapped per layer with near-zero overhead.",
      "No synthesis. No rewiring."])
card(s, MARGIN + 2 * (cw + gap), cy, cw, ch, "Fabless IP business",
     ["FPGA implementation = silicon-proven reference.",
      "Product = ASIC-ready RTL + compiler, licensed",
      "to SoC companies as configurable IP."])
tbox(s, MARGIN, Inches(5.95), Inches(12.1), Inches(0.5),
     [("Positioning: value-edge NPU IP — MCU-class to mid-range edge SoCs.",
       14, True, ACCENT, 0)])

# ================================================== S5 — PHILOSOPHY
s = slide()
chrome(s, 5, "Build hardware for what all models share — not for models",
       "Philosophy")
b = rect(s, MARGIN, Inches(1.6), Inches(12.13), Inches(1.0),
         fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
b.adjustments[0] = 0.14
text_in(b, [("Roughly 90% of the compute in every modern model — CNN, "
             "transformer, LLM, audio, signal — is multiply-accumulate. "
             "The rest is thin vector math.", 15.5, True, WHITE, 0)],
        anchor=MSO_ANCHOR.MIDDLE, inset=0.2)
bullets(s, MARGIN, Inches(3.0), Inches(6.6), Inches(3.2), [
    ("What varies between models is not the arithmetic",
     "it is dataflow, scheduling, and operator mix. Those are software "
     "decisions — so RIA puts them in software."),
    ("Don't build hardware for models",
     "fixed-model accelerators become e-waste on the next architecture "
     "shift."),
    ("Don't build hardware for 15 operator categories either",
     "our own operator-mapping analysis showed 3 units do nearly all the "
     "work; the rest is dark silicon."),
    ("Build one tile that covers MAC + vector math",
     "and make every model-specific decision changeable after tape-out."),
], size=13.5, gap=14)
c = rect(s, Inches(7.6), Inches(3.0), Inches(5.1), Inches(3.35),
         fill=LIGHT, line=CARD_LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
c.adjustments[0] = 0.05
text_in(c, [
    ("Everything below is a vector program,", 13, True, NAVY, 2),
    ("not separate hardware:", 13, True, NAVY, 8),
    ("Softmax  •  GELU / ReLU / SiLU  •  Sigmoid / Tanh", 12.5, False,
     GRAY, 5),
    ("LayerNorm / RMSNorm / BatchNorm", 12.5, False, GRAY, 5),
    ("Pooling  •  Reductions (sum / mean / max)", 12.5, False, GRAY, 5),
    ("Quantize / Dequantize  •  Elementwise ops", 12.5, False, GRAY, 10),
    ("And everything below is the same MAC array,", 13, True, NAVY, 2),
    ("fed in a different order:", 13, True, NAVY, 8),
    ("GEMM  •  Conv (1D/2D/depthwise)  •  Attention QKᵀ",
     12.5, False, GRAY, 0),
], inset=0.2)

# ================================================== S6 — THE TILE
s = slide()
chrome(s, 6, "The RIA tile — one microarchitecture, every personality",
       "Architecture")
cy, ch = Inches(1.6), Inches(2.15)
cw, gap = Inches(5.96), Inches(0.2)
card(s, MARGIN, cy, cw, ch, "MAC array",
     ["Dense multiply-accumulate grid. Executes GEMM,",
      "convolution, and attention score computation —",
      "the same multipliers; microcode selects the",
      "data-feeding order (dataflow mode)."],
     fill=ACCENT_LT)
card(s, MARGIN + cw + gap, cy, cw, ch, "Vector unit + SFU",
     ["Vector lanes with special-function units",
      "(exp, reciprocal, rsqrt). Runs softmax, GELU/ReLU,",
      "LayerNorm/RMSNorm, pooling, reductions,",
      "quantize/dequantize — as programs."],
     fill=ACCENT_LT)
cy2 = cy + ch + Inches(0.2)
card(s, MARGIN, cy2, cw, ch, "Local SRAM (double-buffered)",
     ["Holds weights, activations, and microcode.",
      "The next layer's weights and program load while",
      "the current layer computes — reconfiguration",
      "overhead target: < 1% of runtime."])
card(s, MARGIN + cw + gap, cy2, cw, ch, "Programmable sequencer",
     ["Runs the microcode that gives the tile its",
      "current personality, and coordinates MAC array,",
      "vector unit, and SRAM without host",
      "intervention."])
b = rect(s, MARGIN, Inches(6.15), Inches(12.13), Inches(0.75),
         fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
b.adjustments[0] = 0.18
text_in(b, [("A tile is a softmax tile this layer and a matrix tile the "
             "next — by instruction, not rewiring.",
             15, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE,
        align=PP_ALIGN.CENTER, inset=0.1)

# ============================================ S7 — SYSTEM BLOCK DIAGRAM
s = slide()
chrome(s, 7, "System architecture", "Architecture")


def block(x, y, w, h, label, sub=None, fill=LIGHT, tcolor=NAVY, size=12):
    c = rect(s, x, y, w, h, fill=fill, line=CARD_LINE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    c.adjustments[0] = min(0.5, 0.09 * Inches(1) / h if h else 0.1)
    items = [(label, size, True, tcolor, 2 if sub else 0)]
    if sub:
        items.append((sub, size - 2.5, False, GRAY, 0))
    text_in(c, items, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            inset=0.04)
    return c


# control subsystem (top)
block(Inches(2.5), Inches(1.5), Inches(8.3), Inches(0.72),
      "Control & Scheduler Subsystem",
      "instruction fetch  •  dependency manager  •  tile configurator  •  "
      "performance monitor", fill=NAVY, tcolor=WHITE, size=13)
# host + ext mem (sides)
block(Inches(0.6), Inches(2.6), Inches(1.7), Inches(3.0),
      "External\nMemory", "DDR/LPDDR\nDMA engine\nprefetch +\nwrite-back")
block(Inches(11.05), Inches(2.6), Inches(1.7), Inches(3.0),
      "Host I/F +\nConfig", "AXI interface\nconfig & status\npower mgmt\n"
      "debug / trace")
# tile grid 4x2 identical
tx, ty = Inches(2.5), Inches(2.42)
tw, th = Inches(1.94), Inches(1.1)
gp = Inches(0.18)
for r in range(2):
    for c_ in range(4):
        block(tx + c_ * (tw + gp), ty + r * (th + gp), tw, th,
              "RIA Tile", "MAC • vector\nSRAM • sequencer",
              fill=ACCENT_LT, size=12.5)
tbox(s, Inches(2.5), Inches(4.86), Inches(8.3), Inches(0.3),
     [("…identical tiles, replicated to the customer's performance point",
       10.5, False, GRAY, 0)], align=PP_ALIGN.CENTER)
# scratchpad + NoC
block(Inches(2.5), Inches(5.2), Inches(8.3), Inches(0.55),
      "Shared Scratchpad  (weights • activations • intermediates)",
      fill=LIGHT, size=12)
block(Inches(2.5), Inches(5.9), Inches(8.3), Inches(0.5),
      "NoC / Crossbar Interconnect", fill=NAVY2, tcolor=WHITE, size=12)
tbox(s, MARGIN, Inches(6.6), Inches(12.1), Inches(0.6),
     [("Dataflow strategy (weight- vs. output-stationary) is chosen per "
       "layer by the compiler to minimize external memory traffic — the "
       "dominant energy cost.", 12.5, False, GRAY, 0)])

# ====================================== S8 — TWO-LEVEL CONFIGURATION
s = slide()
chrome(s, 8, "The compiler shapes the hardware — twice",
       "Key mechanism")
cw = Inches(5.96)
c = rect(s, MARGIN, Inches(1.6), cw, Inches(3.9), fill=LIGHT,
         line=CARD_LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
c.adjustments[0] = 0.04
text_in(c, [
    ("LEVEL 1 — DESIGN TIME", 12, True, ACCENT, 4),
    ("Per customer, at IP integration", 16, True, NAVY, 10),
    ("The compiler profiles the customer's model family and emits the IP "
     "configuration:", 12.5, False, GRAY, 8),
    ("•  tile count and MAC array dimensions", 12.5, False, INK, 4),
    ("•  vector width and SFU set", 12.5, False, INK, 4),
    ("•  SRAM sizing and NoC topology", 12.5, False, INK, 10),
    ("A vision customer gets a MAC-heavy build; an audio/LLM customer a "
     "different ratio. Hardware requirements are derived from the model — "
     "not guessed.", 12.5, False, GRAY, 0),
], inset=0.22)
c = rect(s, MARGIN + cw + Inches(0.2), Inches(1.6), cw, Inches(3.9),
         fill=LIGHT, line=CARD_LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
c.adjustments[0] = 0.04
text_in(c, [
    ("LEVEL 2 — RUNTIME", 12, True, ACCENT, 4),
    ("Per model, per layer", 16, True, NAVY, 10),
    ("ONNX graph in → microcode + scheduler instruction stream out:",
     12.5, False, GRAY, 8),
    ("•  compiles in seconds — no synthesis, no P&R", 12.5, False, INK, 4),
    ("•  tiles switch personality layer-by-layer", 12.5, False, INK, 4),
    ("•  identical flow on FPGA prototype and ASIC", 12.5, False, INK, 10),
    ("Ship a new model to silicon already in the field with a software "
     "update.", 12.5, False, GRAY, 0),
], inset=0.22)
b = rect(s, MARGIN, Inches(5.75), Inches(12.13), Inches(0.95),
         fill=ACCENT_LT, line=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
b.adjustments[0] = 0.14
text_in(b, [("This two-level scheme captures most of the efficiency of "
             "per-model hardware while keeping a seconds-long compile loop "
             "and a clean ASIC path — the property neither a fixed NPU nor "
             "an FPGA flow can offer.", 13.5, True, NAVY, 0)],
        anchor=MSO_ANCHOR.MIDDLE, inset=0.18)

# ================================================== S9 — COMPILER
s = slide()
chrome(s, 9, "The compiler is the product", "Software stack")
stages = [
    ("ONNX ingest", "PyTorch / TF /\nJAX via ONNX"),
    ("Operator\nlowering", "graph ops →\ntile primitives"),
    ("Spatial\npartitioning", "which tiles run\nwhich operator"),
    ("Temporal\nscheduling", "when personalities\nswitch, per layer"),
    ("Memory\nplanning", "scratchpad alloc,\nprefetch, dataflow"),
    ("Microcode\nemission", "tile programs +\nscheduler stream"),
]
n = len(stages)
bw, bh = Inches(1.78), Inches(1.25)
gap_x = (Inches(12.13) - bw * n) / (n - 1)
y0 = Inches(1.7)
for i, (h, sub) in enumerate(stages):
    x = MARGIN + i * (bw + gap_x)
    block(x, y0, bw, bh, h, sub, fill=ACCENT_LT if i in (2, 3, 4) else LIGHT,
          size=12)
    if i < n - 1:
        a = rect(s, x + bw + Inches(0.01), y0 + bh / 2 - Inches(0.08),
                 gap_x - Inches(0.02), Inches(0.16), fill=ACCENT,
                 shape=MSO_SHAPE.RIGHT_ARROW)
tbox(s, MARGIN, Inches(3.15), Inches(12.1), Inches(0.4),
     [("Highlighted stages are the hard, differentiating middle — a "
       "mini place-and-route problem solved in software, in seconds.",
       12, False, GRAY, 0)])
bullets(s, MARGIN, Inches(3.75), Inches(12.1), Inches(2.6), [
    ("Built on MLIR", "we stand on LLVM-community infrastructure instead "
     "of hand-rolling a compiler from scratch — lower risk, faster "
     "operator coverage."),
    ("Why software-heavy is deliberate", "a compiler keeps improving "
     "after tape-out; silicon doesn't. Every scheduling win ships to all "
     "deployed chips."),
    ("Where the moat is", "history is unambiguous — AI hardware ventures "
     "fail on software, not RTL. We budget the majority of engineering "
     "here, and say so."),
    ("Operator coverage plan", "MAC + vector primitives cover the ONNX "
     "core opset; long-tail ops fall back to the host CPU until "
     "profiled as worth accelerating."),
], size=13.5, gap=12)

# ============================================ S10 — MODEL MAPPING
s = slide()
chrome(s, 10, "How model families map onto the same silicon",
       "Model coverage")
rows = [
    ["Model family", "Dominant tile personalities",
     "Approx. compute mix", "Note"],
    ["CNNs  (MobileNet, ResNet)", "Conv-mode MAC + vector "
     "(pool, norm, activation)", "~85% MAC / 15% vector",
     "conv = MAC array in conv dataflow"],
    ["Transformers  (BERT, ViT)", "GEMM-mode MAC + vector "
     "(softmax, LayerNorm, GELU)", "~75% MAC / 25% vector",
     "attention = GEMM + vector softmax"],
    ["LLMs  (Llama-class, decode)", "GEMM-mode MAC + vector; "
     "bandwidth-bound", "~70% MAC / 30% vector",
     "prefetch + quantization critical"],
    ["Audio  (Whisper, KWS)", "Conv + GEMM modes + vector",
     "~80% MAC / 20% vector", "FFT frontend runs as vector program"],
    ["Signal / Health  (ECG, radar)", "Conv-mode MAC + vector reductions",
     "~85% MAC / 15% vector", "small models — few tiles, low power"],
]
table(s, MARGIN, Inches(1.65), Inches(12.13), [0.24, 0.32, 0.20, 0.24],
      rows, size=11.5, row_h=Inches(0.62))
b = rect(s, MARGIN, Inches(5.6), Inches(12.13), Inches(1.05),
         fill=ACCENT_LT, line=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
b.adjustments[0] = 0.12
text_in(b, [
    ("Every row uses the same two hardware resources — MAC array and "
     "vector unit — in different proportions over time.", 14, True,
     NAVY, 4),
    ("Because any tile can serve either role, utilization stays high "
     "across all families. No family leaves silicon dark.",
     13, False, GRAY, 0)], inset=0.18)

# ============================================ S11 — FLAGSHIP PROOF
s = slide()
chrome(s, 11, "The flagship demo — the thesis in one experiment",
       "Proof plan")
b = rect(s, MARGIN, Inches(1.6), Inches(12.13), Inches(1.15),
         fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
b.adjustments[0] = 0.10
text_in(b, [
    ("MobileNetV2 and a small transformer, both compiled from ONNX, both "
     "running on the same FPGA bitstream — zero RTL changes between them.",
     16, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE, inset=0.2)
cy, cw, ch, gap = Inches(3.1), Inches(2.9), Inches(1.9), Inches(0.18)
metrics = [
    ("> 70%", "tile utilization held on both model families "
     "(fixed-function baseline: high on one, collapses on the other)"),
    ("< 1%", "reconfiguration overhead — personality swaps hidden under "
     "double-buffered execution"),
    ("seconds", "model-to-microcode compile time, versus hours for "
     "per-model FPGA synthesis"),
    ("inf/s/W", "reported against a fixed-function baseline and a CPU/GPU "
     "reference on the same board"),
]
for i, (big, small) in enumerate(metrics):
    c = rect(s, MARGIN + i * (cw + gap), cy, cw, ch, fill=LIGHT,
             line=CARD_LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    c.adjustments[0] = 0.08
    text_in(c, [(big, 24, True, ACCENT, 4), (small, 10.5, False, GRAY, 0)],
            inset=0.15)
bullets(s, MARGIN, Inches(5.35), Inches(12.1), Inches(1.5), [
    ("Utilization across model families is the headline number",
     "it is the one claim fixed-function NPUs cannot match and FPGA flows "
     "cannot match at our compile speed."),
    ("Everything is measured, nothing asserted",
     "the same report doubles as the evaluation section of a publication "
     "and the benchmark annex of a licensing data sheet."),
], size=13, gap=10)

# ============================================ S12 — COMPETITION
s = slide()
chrome(s, 12, "Competitive landscape — the honest version", "Market")
rows = [
    ["Player", "What they ship", "RIA's differentiation"],
    ["ARM Ethos-U/N, Ceva NeuPro,\nSynopsys ARC NPX, Cadence "
     "Tensilica", "Incumbent licensable NPU IP; configurable at design "
     "time only",
     "Compiler-derived configuration + runtime personality switching: "
     "same silicon tracks model evolution"],
    ["AMD Versal AI Engines",
     "Shipping adaptive-tile silicon (buy a chip)",
     "RIA is licensable IP for customers' own SoCs, at edge-class "
     "power/cost"],
    ["Tenstorrent, SambaNova",
     "Identical-programmable-tile architectures at datacenter scale",
     "Validation that the principle works; RIA claims the edge-IP niche "
     "they don't serve"],
    ["FINN / hls4ml flows",
     "Per-model FPGA generation for small frozen models",
     "RIA compiles in seconds, needs no P&R, and has an ASIC path"],
]
table(s, MARGIN, Inches(1.6), Inches(12.13), [0.28, 0.30, 0.42], rows,
      size=11, row_h=Inches(0.85))
b = rect(s, MARGIN, Inches(5.5), Inches(12.13), Inches(1.15),
         fill=ACCENT_LT, line=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
b.adjustments[0] = 0.12
text_in(b, [
    ("The Indian RISC-V SoC ecosystem (Mindgrove, InCore, and peers) is "
     "our customer base, not our competition —", 13.5, True, NAVY, 4),
    ("they need exactly this block, and the India Semiconductor Mission / "
     "DLI scheme is the tailwind for domestic IP.", 13, False, GRAY, 0)],
    inset=0.18)

# ============================================ S13 — BUSINESS MODEL
s = slide()
chrome(s, 13, "Business model — fabless IP licensing", "Business")
cy, cw, ch, gap = Inches(1.6), Inches(3.94), Inches(2.5), Inches(0.2)
card(s, MARGIN, cy, cw, ch, "The product",
     ["ASIC-ready RTL of the tile + fabric,",
      "delivered as configurable IP, plus the",
      "compiler/runtime stack. The FPGA build is the",
      "silicon-proven reference and each customer's",
      "design-space exploration vehicle."])
card(s, MARGIN + cw + gap, cy, cw, ch, "Revenue",
     ["License fee per design win,",
      "per-unit royalty on shipped SoCs,",
      "annual compiler & toolchain subscription,",
      "NRE for customer-specific tile-mix",
      "configuration."])
card(s, MARGIN + 2 * (cw + gap), cy, cw, ch, "Target segment",
     ["Value-edge SoCs: smart cameras, wearables,",
      "hearables, industrial sensing, automotive",
      "sub-systems, health devices —",
      "below Tenstorrent's market,",
      "above frozen-model FPGA flows."])
bullets(s, MARGIN, Inches(4.5), Inches(12.1), Inches(2.2), [
    ("Capital-efficient by construction", "no tape-out on our books; "
     "customers carry silicon cost. FPGA proof is achievable on a "
     "startup budget."),
    ("The compiler subscription compounds", "scheduling improvements ship "
     "to every licensee — recurring revenue tied to real, measurable "
     "gains."),
    ("Ecosystem tailwind", "India Semiconductor Mission and the DLI "
     "scheme actively fund domestic IP; domestic SoC startups need an "
     "NPU block they can license without ARM-scale fees."),
], size=13.5, gap=12)

# ============================================ S14 — ROADMAP
s = slide()
chrome(s, 14, "Roadmap — milestones, not checkmarks", "Execution")
phases = [
    ("Phase 1  •  months 0–6", "Foundations", [
        "Tile RTL + ISA specification frozen",
        "Single-tile FPGA validation (GEMM, conv, vector ops)",
        "Compiler MVP: ONNX → microcode for core opset"]),
    ("Phase 2  •  months 6–12", "Fabric", [
        "Multi-tile array + NoC + scheduler subsystem",
        "MobileNetV2 end-to-end on FPGA",
        "First utilization + power measurements"]),
    ("Phase 3  •  months 12–18", "Proof", [
        "Transformer support in compiler",
        "Flagship demo: two model families, one bitstream",
        "Publication + benchmark report"]),
    ("Phase 4  •  months 18–24", "Product", [
        "ASIC hardening + PPA characterization",
        "Design-time configurator (tile-mix tool)",
        "First evaluation licenses signed"]),
]
cw, gap = Inches(2.93), Inches(0.14)
for i, (when, name, items) in enumerate(phases):
    x = MARGIN + i * (cw + gap)
    c = rect(s, x, Inches(1.7), cw, Inches(3.6), fill=LIGHT, line=CARD_LINE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    c.adjustments[0] = 0.045
    body = [(when, 10.5, True, ACCENT, 3), (name, 17, True, NAVY, 10)]
    for it in items:
        body.append(("•  " + it, 11.5, False, GRAY, 7))
    text_in(c, body, inset=0.18)
    rect(s, x, Inches(1.7), cw, Inches(0.07), fill=ACCENT)
tbox(s, MARGIN, Inches(5.6), Inches(12.1), Inches(0.5),
     [("Exit criterion for each phase is a measured number on hardware — "
       "not a document.", 13.5, True, NAVY, 0)])

# ============================================ S15 — RISKS
s = slide()
chrome(s, 15, "Risks — stated, owned, and instrumented", "Honesty")
rows = [
    ["Risk", "Reality", "Mitigation & metric"],
    ["Efficiency gap vs.\nfixed-function",
     "Programmability always costs silicon; the gap must be measured, "
     "never assumed",
     "Target: within 2–3x of a fixed NPU per-op while holding >70% "
     "utilization across model families — net wins on real workloads"],
    ["Compiler execution",
     "The schedule risk lives in software, not RTL — this is where AI "
     "hardware ventures die",
     "MLIR base, majority of engineering budget, compiler milestones "
     "gate every phase"],
    ["Memory bandwidth",
     "Dominates edge energy; a scratchpad alone is not an answer",
     "Per-layer dataflow selection, prefetch/double-buffering; DRAM "
     "traffic per inference reported in every benchmark"],
    ["Incumbent IP vendors",
     "ARM/Ceva/Synopsys have distribution and trust",
     "Win on configuration agility, model-evolution insurance, cost, "
     "and domestic-ecosystem access — not on raw TOPS"],
]
table(s, MARGIN, Inches(1.65), Inches(12.13), [0.20, 0.36, 0.44], rows,
      size=11, row_h=Inches(1.0))
tbox(s, MARGIN, Inches(6.15), Inches(12.1), Inches(0.5),
     [("We publish these numbers whether they flatter us or not — "
       "credibility is the asset an IP vendor sells.", 13, True, ACCENT,
       0)])

# ============================================ S16 — CLOSE
s = slide()
rect(s, 0, 0, SW, SH, fill=NAVY)
rect(s, 0, Inches(1.35), SW, Inches(0.05), fill=ACCENT)
tbox(s, MARGIN, Inches(0.65), Inches(12), Inches(0.6),
     [("HOW WE GOT HERE", 13, True, ACCENT, 0)])
tbox(s, MARGIN, Inches(1.7), Inches(12.1), Inches(3.2), [
    ("v0 proposed 15 specialized tiles — honest about primitives, but "
     "really an NPU with 15 fixed accelerators: dark silicon for every "
     "model that skipped one.", 15, False, ACCENT_LT, 10),
    ("A fully FPGA-like fabric fixes utilization but costs hours-long "
     "compiles and loses the ASIC path.", 15, False, ACCENT_LT, 10),
    ("RIA v1 is the durable middle: one hardened programmable tile, "
     "replicated — the compiler shapes the array at integration time and "
     "reprograms it per layer at runtime.", 15, True, WHITE, 0),
])
tbox(s, MARGIN, Inches(4.9), Inches(12.1), Inches(1.4), [
    ("One architecture. Every model. Any customer.", 30, True, WHITE, 8),
    ("High utilization on every model family  •  seconds to compile  •  "
     "a clean FPGA-to-ASIC-IP path", 15, False, ACCENT, 0),
])
tbox(s, MARGIN, SH - Inches(0.85), Inches(12.1), Inches(0.5),
     [("RIA — Reconfigurable Intelligence Architecture  •  Concept v1  •  "
       "July 2026", 12, False, RGBColor(0x8F, 0xA6, 0xBD), 0)])

OUT = "RIA_Deck_v1.pptx"
prs.save(OUT)
print(f"saved {OUT} with {len(prs.slides.__iter__.__self__._sldIdLst)} slides")
