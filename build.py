"""Build index.html from the proposal deck.

Usage:  python build.py ["deck.pptx"]
Thumbnails: run export_thumbs.ps1 first (needs PowerPoint), they go to thumbs/<n>.jpg.
When the deck changes, update SECTIONS / QUESTIONS below if slide numbers moved.
"""
import html, json, re, sys, zipfile
from pathlib import Path

DECK = sys.argv[1] if len(sys.argv) > 1 else "Robotaxi - Proposal Presentation (1).pptx"

# (first slide, section name, is_appendix) - a section runs until the next one starts
SECTIONS = [
    (1, "Intro: Background, Problem, Objectives, Scope", False),
    (13, "Review of Existing Knowledge & Gaps", False),
    (23, "System Overview & Hardware", False),
    (29, "System Architecture & Dashboard", False),
    (41, "Roles & Teamwork", False),
    (45, "Roadmap & Q&A", False),
    (49, "Appendices (cover)", True),
    (50, "Safety Metrics", True),
    (60, "Datasets & Feasibility", True),
    (85, "AV Regulations by Country (SG, CN, US, EU)", True),
    (110, "Misc / Extra", True),
    (117, "Thailand Regulatory Proposal (TDRI)", True),
    (127, "Regulation Matrix & Requirement-Feature Mapping", True),
    (154, "AI for Transportation", True),
    (160, "Agentic Frameworks (Netradyne, ChatSUMO, Multi-agent)", True),
    (180, "Agentic AI Adoption (MLflow, Incident, RCA, NL Query)", True),
    (202, "Model Evaluation (BFCL tool calling)", True),
    (220, "Incident Report Agent Evaluation", True),
    (228, "Stress & Performance Testing", True),
    (239, "AI Model Cost", True),
    (244, "Architecture Decision Records (ADR)", True),
    (250, "Timeline", True),
    (258, "Data Size", True),
    (260, "Architecture (detailed diagrams)", True),
]

# Likely committee questions -> slides. Edit freely.
QUESTIONS = [
    ("Why this project? Motivation / why Thailand needs it", "5-6,22,111"),
    ("What exactly is the problem?", "7-8"),
    ("Objectives, scope, limitations, benefits", "9-12"),
    ("What is new vs existing systems / research gap?", "18-22,92-96,161-167"),
    ("Which safety metrics and why these?", "51-59,247"),
    ("How do you compute TTC / PET? Is it traceable?", "52-55,80-81,246"),
    ("Why nuScenes? What other datasets did you consider?", "16,61-66,72-79,245"),
    ("Thai data? Does it work in Thailand context?", "67-71,82-84"),
    ("What do regulators abroad require?", "15,86-109,128-144"),
    ("Thailand regulation / TDRI context", "22,117-126"),
    ("How do features map to regulatory requirements?", "135,145-153"),
    ("System architecture overview", "24-25,30-40,260-276"),
    ("Hardware / compute / constraints", "26-28"),
    ("What does the dashboard show?", "32-35"),
    ("Why AI agents? What is agentic here?", "17,21,155-159,169-179,201"),
    ("Incident analysis / report agent", "188-193,220"),
    ("Root cause analysis", "194-197"),
    ("Natural language query", "34,198-200,219"),
    ("How do you evaluate the AI? Hallucination?", "21,203-227"),
    ("Is tool calling reliable? (BFCL results)", "203-218"),
    ("Model latency / throughput / load test", "212,229-238"),
    ("How much does the LLM cost? Self-host vs API", "239-243"),
    ("Why Ollama / MLflow / self-host?", "37,182-187"),
    ("Privacy / PDPA / security / access control", "121-122,139,248"),
    ("Data recording, retention, storage size", "119-120,136,143,249,258-259"),
    ("Real-time vs replay?", "110,123,245,249"),
    ("Timeline: can you finish on time?", "46-47,251-257"),
    ("Who does what?", "42-44"),
]

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def paras(xml):
    """Text per paragraph; runs inside one paragraph are joined without spaces."""
    out = []
    for p in re.findall(r"<a:p>(.*?)</a:p>", xml, re.S):
        t = "".join(html.unescape(x) for x in re.findall(r"<a:t>([^<]*)</a:t>", p)).strip()
        if t:
            out.append(t)
    return out


def title_of(xml, ps):
    for sp in re.findall(r"<p:sp>(.*?)</p:sp>", xml, re.S):
        if re.search(r'<p:ph[^>]*type="(title|ctrTitle)"', sp):
            t = " ".join(paras(sp))
            if t:
                return t
    for t in ps:  # first text that isn't a page number / decoration
        if not re.fullmatch(r"[\d\s+|]+|Computer Engineering|Chulalongkorn University", t):
            return t
    return ""


def ranges(spec):
    out = []
    for part in spec.split(","):
        a, _, b = part.partition("-")
        out += range(int(a), int(b or a) + 1)
    return out


def build():
    z = zipfile.ZipFile(DECK)
    pres = z.read("ppt/presentation.xml").decode("utf8")
    rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf8")
    rid2file = {}
    for rel in re.findall(r"<Relationship [^>]*/>", rels):
        rid, tgt = re.search(r'Id="([^"]+)"', rel)[1], re.search(r'Target="([^"]+)"', rel)[1]
        rid2file[rid] = tgt
    order = [rid2file[r] for r in re.findall(r'<p:sldId [^>]*r:id="([^"]+)"', pres)]

    ocr_file = Path(__file__).with_name("ocr.json")  # text for image-only slides
    ocr = json.loads(ocr_file.read_text(encoding="utf8")) if ocr_file.exists() else {}
    starts = [s[0] for s in SECTIONS]
    slides = []
    for n, f in enumerate(order, 1):
        xml = z.read("ppt/" + f).decode("utf8")
        ps = paras(xml)
        notes = ""
        rel_path = "ppt/slides/_rels/" + f.split("/")[-1] + ".rels"
        if rel_path in z.namelist():
            m = re.search(r'Target="\.\./notesSlides/([^"]+)"', z.read(rel_path).decode("utf8"))
            if m:
                notes = " ".join(paras(z.read("ppt/notesSlides/" + m[1]).decode("utf8")))
                notes = re.sub(r"\s*\b\d+\s*$", "", notes)  # drop trailing slide-number placeholder
        sec = max(i for i, s in enumerate(starts) if s <= n)
        slides.append({
            "n": n,
            "t": title_of(xml, ps) or f"Slide {n}",
            "x": " ".join(ps + [ocr.get(str(n), "")]).strip(),
            "notes": notes,
            "s": sec,
            "hidden": ' show="0"' in xml[:600],
        })

    total = len(slides)
    sections = []
    for i, (start, name, app) in enumerate(SECTIONS):
        end = SECTIONS[i + 1][0] - 1 if i + 1 < len(SECTIONS) else total
        sections.append({"name": name, "from": start, "to": end, "app": app})
    questions = [{"q": q, "slides": [n for n in ranges(spec) if n <= total]} for q, spec in QUESTIONS]

    data = json.dumps({"slides": slides, "sections": sections, "questions": questions}, ensure_ascii=False)
    tpl = Path(__file__).with_name("template.html").read_text(encoding="utf8")
    Path(__file__).with_name("index.html").write_text(tpl.replace("/*DATA*/null", data), encoding="utf8")
    print(f"index.html: {total} slides, {len(sections)} sections, {len(questions)} questions")


if __name__ == "__main__":
    build()
