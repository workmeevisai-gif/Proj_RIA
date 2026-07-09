#!/usr/bin/env python3
"""Generate the SASTRA-TCS TIC project proposal (RIA proof-of-concept)
as a .docx — Times New Roman 12, within the 12-page limit."""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1B, 0x3A, 0x5F)
BLACK = RGBColor(0, 0, 0)

doc = Document()

# ---- page + base style -------------------------------------------------
for sec in doc.sections:
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.2)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)

st = doc.styles["Normal"]
st.font.name = "Times New Roman"
st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
st.paragraph_format.space_after = Pt(6)
st.paragraph_format.line_spacing = 1.15


def para(text, bold=False, italic=False, align=None, size=12, color=BLACK,
         space_after=6):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    r.font.color.rgb = color
    return p


def heading(num, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(f"{num}. {text}" if num else text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.bold = True
    r.font.color.rgb = NAVY
    return p


def bullet(lead, rest=""):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(lead)
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.bold = bool(rest)
    if rest:
        r2 = p.add_run(" " + rest)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(12)
    return p


def table(rows, widths=None, header=True, size=11):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci)
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(str(val))
            r.font.name = "Times New Roman"
            r.font.size = Pt(size)
            if header and ri == 0:
                r.bold = True
                shd = OxmlElement("w:shd")
                shd.set(qn("w:fill"), "1B3A5F")
                cell._tc.get_or_add_tcPr().append(shd)
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            if widths:
                cell.width = Cm(widths[ci])
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


# ======================================================== TITLE BLOCK
para("Project Proposal", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
     size=14, color=NAVY, space_after=2)
para("SASTRA–TCS Technology Innovation Centre (TIC) — Call for Proposal",
     align=WD_ALIGN_PARAGRAPH.CENTER, size=11, space_after=14)

para("RIA-PoC: A Reconfigurable Tile-Based Edge-AI Accelerator on FPGA",
     bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=16, color=NAVY,
     space_after=14)

table([
    ["Domain", "Artificial Intelligence (edge-AI hardware)"],
    ["Duration", "18 months"],
    ["Funding requested", "Rs. 50.00 Lakh (capital + revenue)"],
    ["Current TRL", "TRL 2–3 (concept formulated; architecture defined "
     "on paper)"],
    ["Target TRL", "TRL 4–5 (technology validated in laboratory on "
     "representative workloads)"],
    ["Institution", "<Name of the Institution>"],
    ["Principal Investigator", "<PI name, designation, department>"],
    ["Team", "1 PI, 1 Co-PI (VLSI / embedded systems), 2 project "
     "associates, 4 UG/PG students"],
], widths=[4.5, 11.5], header=False, size=11)

# ======================================================== 1. SUMMARY
heading(1, "Executive Summary")
para("Edge devices increasingly need to run machine-learning models "
     "locally, but the hardware options available to Indian product "
     "companies are unsatisfactory: fixed-function AI accelerators become "
     "obsolete when model architectures change, general-purpose processors "
     "are too power-hungry, and raw FPGAs demand hardware expertise that "
     "most product teams do not have.")
para("This project will build and measure a proof-of-concept of the "
     "Reconfigurable Intelligence Architecture (RIA): a small array of "
     "four identical programmable compute tiles on an FPGA, where each "
     "tile can act as a matrix-multiply engine, a convolution engine, or "
     "a vector/activation engine depending on the instructions loaded "
     "into it. The same hardware, without re-synthesis, will be "
     "demonstrated on two representative edge workloads: a compact "
     "vision CNN and a 1-D ECG classification model. Success is defined "
     "by measured numbers — tile utilisation, energy per inference "
     "against the on-board embedded CPU, and model-swap time — not by "
     "claims.")

# ======================================================== 2. PROBLEM
heading(2, "Problem Statement and Motivation")
bullet("Fixed accelerators age badly.", "NPUs hard-wired for one family "
       "of models (e.g., CNNs) waste silicon or fail outright when the "
       "deployed model changes; edge products live 5–10 years while "
       "model architectures turn over in 1–2.")
bullet("The edge power budget rules out GPUs.", "Battery-powered and "
       "fanless devices operate in the milliwatt-to-watt range.")
bullet("FPGAs are flexible but inaccessible.", "Every model change "
       "requires hours of synthesis and place-and-route plus digital "
       "design expertise — impractical for AI/product teams.")
para("The gap this project addresses: an accelerator whose function can "
     "be changed in software, in seconds, while remaining far more "
     "energy-efficient than an embedded CPU. The insight that makes this "
     "tractable is that the bulk of the computation in practically every "
     "ML model is multiply-accumulate work plus a thin layer of vector "
     "operations (activations, normalisation, pooling); what differs "
     "between models is the order and proportion in which these "
     "primitives are used — and that can be expressed as instructions "
     "rather than as circuit structure.")

# ======================================================== 3. OBJECTIVES
heading(3, "Objectives")
para("The project deliberately proposes a narrow, measurable scope:",
     space_after=4)
bullet("O1.", "Design and verify the RTL of one programmable compute "
       "tile: a MAC (multiply-accumulate) array, a small vector unit "
       "with activation support, local buffer memory, and an instruction "
       "sequencer, operating on INT8/INT16 data.")
bullet("O2.", "Integrate a 2×2 array of four identical tiles with a "
       "shared on-chip buffer, a simple interconnect, and a "
       "microprogrammed scheduler on a Xilinx Zynq UltraScale+ FPGA.")
bullet("O3.", "Develop a basic mapping utility that converts a trained "
       "model (ONNX, restricted to a documented set of supported "
       "layers) into tile instructions and a schedule. Hand-tuning of "
       "the generated schedule is acceptable at this stage; a fully "
       "automatic optimising compiler is explicitly out of scope.")
bullet("O4.", "Demonstrate two edge-relevant workloads on the same "
       "bitstream, switching between them in software only: (a) a "
       "compact image-classification CNN (MobileNet-class), and (b) a "
       "1-D CNN for ECG arrhythmia classification.")
bullet("O5.", "Publish a measured benchmark report: tile utilisation, "
       "latency, energy per inference versus the on-board ARM "
       "Cortex-A53 as baseline, and reconfiguration time.")

# ======================================================== 4. CONCEPT
heading(4, "Conceptualisation and Technical Approach")
para("Architecture. Each RIA tile contains four elements: (i) an 8×8 "
     "INT8 MAC array for matrix and convolution arithmetic; (ii) a "
     "16-lane vector unit implementing activations (ReLU family), "
     "pooling, and normalisation as short programs; (iii) dual-banked "
     "local SRAM so that the next layer's weights and instructions load "
     "while the current layer computes; and (iv) a sequencer that "
     "executes the loaded microprogram. A tile therefore behaves as a "
     "convolution engine in one layer and as an activation/pooling "
     "engine in the next, purely by instruction change.")
para("System. Four tiles, a shared scratchpad buffer, DMA to external "
     "DDR memory, and a scheduler that dispatches per-layer work to "
     "tiles are integrated on the programmable logic of a Zynq "
     "UltraScale+ device. The on-chip ARM core acts as host and as the "
     "measurement baseline. The array size (2×2) is chosen to be large "
     "enough to demonstrate multi-tile scheduling and small enough to "
     "be completed and verified within 18 months by a small team.")
para("Software (deliberately modest). The mapping utility parses an "
     "ONNX graph, checks it against the supported layer set "
     "(convolution, dense, ReLU-family activation, pooling, batch "
     "normalisation folding, quantisation to INT8), allocates layers to "
     "tiles, and emits instruction streams. It is a translation tool, "
     "not an optimising compiler; where automatic mapping is "
     "inefficient, schedules will be hand-adjusted and the adjustment "
     "documented. This honesty keeps the software effort proportionate "
     "to a proof-of-concept.")
para("What is intentionally excluded: transformer/attention support, "
     "multi-FPGA scaling, automatic design-space exploration, and any "
     "ASIC work. These are follow-on activities, not commitments of "
     "this project.")

# ======================================================== 5. WORK PLAN
heading(5, "Work Plan, Design and Delivery Milestones (18 months)")
table([
    ["Phase", "Months", "Work", "Milestone (verifiable)"],
    ["P1", "1–3", "Tile micro-architecture specification; instruction "
     "set definition; testbench setup",
     "M1: Frozen tile spec + ISA document"],
    ["P2", "4–7", "Single-tile RTL design and verification; synthesis "
     "on target FPGA",
     "M2: One tile passing GEMM, convolution and vector test vectors "
     "on hardware"],
    ["P3", "8–11", "2×2 array integration: shared buffer, DMA, "
     "scheduler; multi-tile test cases",
     "M3: Four-tile array executing a multi-layer synthetic network"],
    ["P4", "10–14", "Mapping utility for the supported ONNX layer set; "
     "INT8 quantisation flow (overlaps P3)",
     "M4: ECG model running end-to-end from ONNX input"],
    ["P5", "14–17", "Second workload (MobileNet-class CNN); performance "
     "and power measurement campaign",
     "M5: Both models on one bitstream; measured benchmark tables"],
    ["P6", "17–18", "Benchmark report, documentation, publication and "
     "IP filing; TIC demonstration",
     "M6: Public demo + final report to TIC"],
], widths=[1.4, 1.8, 7.0, 5.8], size=10.5)
para("Milestones M2, M3 and M5 are hardware demonstrations and can be "
     "witnessed by TIC reviewers on the bench.", italic=True)

# ======================================================== 6. DELIVERABLES
heading(6, "Deliverables and Success Metrics")
bullet("D1.", "Verified RTL of the RIA tile and 4-tile system "
       "(institution-owned IP).")
bullet("D2.", "Working FPGA prototype demonstrating both workloads "
       "with software-only switching.")
bullet("D3.", "Mapping utility and user documentation for the "
       "supported layer set.")
bullet("D4.", "Measured benchmark report; at least one publication in "
       "a peer-reviewed VLSI/embedded-systems venue; one Indian patent "
       "application on the tile/scheduler mechanism.")
bullet("D5.", "Four UG/PG students trained in accelerator design and "
       "two project associates skilled in RTL + ML deployment.")
para("Quantitative success criteria:", bold=True, space_after=4)
table([
    ["Metric", "Target", "How measured"],
    ["Average tile utilisation", "≥ 60% on both workloads",
     "Hardware performance counters"],
    ["Energy per inference", "≥ 5× better than ARM Cortex-A53 baseline "
     "on the same board", "On-board power rails (PMBus) measurement"],
    ["Model switch time", "< 1 minute, software only (no re-synthesis)",
     "Wall-clock, scripted"],
    ["Reconfiguration overhead", "< 5% of inference runtime",
     "Cycle counters"],
    ["Accuracy after INT8 quantisation", "Within 1.5% of FP32 reference",
     "Standard test sets (e.g., MIT-BIH for ECG)"],
], widths=[4.6, 6.0, 5.4], size=10.5)

# ======================================================== 7. GTM
heading(7, "Prototype to Go-To-Market Strategy")
para("This project delivers a laboratory-validated prototype (TRL 4–5). "
     "The commercialisation path beyond it is staged and realistic:")
bullet("Near term (during project).", "The ECG demonstration is chosen "
       "deliberately: single-model, low-power medical and industrial "
       "monitoring devices are a market where FPGA-based deployment is "
       "already accepted, so the prototype itself is close to a usable "
       "reference design.")
bullet("After the project (12–24 months).", "Package the design as an "
       "evaluation kit (board + mapping utility + documentation) for "
       "Indian device makers; incubate the activity through the "
       "institution's TBI with SASTRA–TCS TIC guidance; apply to the "
       "Design-Linked Incentive (DLI) scheme, for which a working "
       "FPGA prototype with measured results is the natural entry "
       "evidence.")
bullet("Long term.", "Scale the tile array, extend the supported layer "
       "set, and license the design as semiconductor IP to Indian SoC "
       "companies. This is the ambition, but no deliverable of the "
       "present proposal depends on it.")

# ======================================================== 8. TEAM
heading(8, "Team Composition")
para("<One-paragraph profile of the team lead / PI to be inserted here, "
     "per the call requirement: name, designation, relevant experience "
     "in VLSI/FPGA/embedded ML, key publications or projects.>",
     italic=True)
table([
    ["Role", "Count", "Responsibility"],
    ["Principal Investigator (faculty)", "1",
     "Architecture, technical direction, reporting to TIC"],
    ["Co-PI (faculty, VLSI/embedded)", "1",
     "Verification methodology, measurement campaign"],
    ["Project Associates (full-time)", "2",
     "One RTL-focused, one software/ML-deployment-focused"],
    ["UG/PG student members", "4",
     "Testbenches, model preparation, benchmarking, documentation"],
], widths=[6.4, 1.6, 8.0], size=10.5)

# ======================================================== 9. BUDGET
heading(9, "Budget (Total: Rs. 50.00 Lakh)")
para("Capital expenditure:", bold=True, space_after=4)
table([
    ["#", "Item", "Amount (Rs. Lakh)"],
    ["1", "2 × AMD-Xilinx Zynq UltraScale+ FPGA development boards "
     "(ZCU104 class), incl. import duty/GST", "6.00"],
    ["2", "2 × development workstations (simulation and synthesis)",
     "4.00"],
    ["3", "1 × large-format 4K monitor (43\") for waveform/layout "
     "debugging", "0.80"],
    ["4", "Lab accessories: bench power supplies, JTAG programmers, "
     "current-measurement instrumentation, cabling, storage", "1.20"],
    ["", "Sub-total (capital)", "12.00"],
], widths=[0.9, 11.5, 3.6], size=10.5)
para("Software licences:", bold=True, space_after=4)
table([
    ["#", "Item", "Amount (Rs. Lakh)"],
    ["5", "2 × AMD-Xilinx Vivado/Vitis design suite licences "
     "(18 months, node-locked seats — required for the target device)",
     "7.50"],
    ["", "Sub-total (software)", "7.50"],
], widths=[0.9, 11.5, 3.6], size=10.5)
para("Manpower (revenue):", bold=True, space_after=4)
table([
    ["#", "Item", "Amount (Rs. Lakh)"],
    ["6", "2 × Project Associates @ Rs. 45,000/month × 18 months",
     "16.20"],
    ["7", "UG/PG student stipends and internship support", "1.80"],
    ["", "Sub-total (manpower)", "18.00"],
], widths=[0.9, 11.5, 3.6], size=10.5)
para("Operations and other heads:", bold=True, space_after=4)
table([
    ["#", "Item", "Amount (Rs. Lakh)"],
    ["8", "Cloud compute for regression simulation and CI", "1.00"],
    ["9", "Conference travel and registration (two national venues)",
     "1.50"],
    ["10", "Publication charges and one Indian patent filing", "1.50"],
    ["11", "Consumables (SD cards, adapters, spare components)", "0.50"],
    ["12", "Contingency", "3.00"],
    ["13", "Institutional overhead (10%)", "5.00"],
    ["", "Sub-total (operations & other)", "12.50"],
], widths=[0.9, 11.5, 3.6], size=10.5)
table([
    ["Grand total", "Rs. 50.00 Lakh"],
], widths=[11.5, 4.5], header=True, size=11)

# ======================================================== 10. RISKS
heading(10, "Risks and Mitigation")
table([
    ["Risk", "Mitigation"],
    ["Tile RTL verification takes longer than planned",
     "Array fixed at 2×2 and layer set frozen at M1; scope is cut "
     "(fewer supported layers), never extended, to protect milestones"],
    ["Mapping utility proves harder than expected",
     "Hand-written schedules are an accepted fallback for both demos; "
     "the utility then targets only the ECG model"],
    ["Energy target (≥5× vs CPU) not met initially",
     "Double-buffering and DMA batching are the known levers; the "
     "measurement campaign in P5 reserves time for two optimisation "
     "iterations"],
    ["Staff attrition (project associates)",
     "Overlapping student training from month 1 ensures continuity; "
     "documentation is a graded deliverable of every phase"],
], widths=[6.8, 9.2], size=10.5)

# ======================================================== 11. FACILITIES
heading(11, "Institutional Facilities and TIC Alignment")
para("The institution will provide laboratory space within campus (per "
     "the call requirement), existing computing infrastructure, and "
     "faculty time. Access to SASTRA facilities offered under the call "
     "will be used for measurement instrumentation and for the "
     "demonstration reviews at milestones M2, M3 and M5. The project "
     "matches the TIC objective of proof-of-concept to real-world "
     "impact: it ends with working hardware, measured numbers, a "
     "medical-adjacent demonstration, and a documented path into "
     "incubation — not with a report alone.")

doc.save("RIA_TIC_Proposal.docx")
print("saved RIA_TIC_Proposal.docx")
