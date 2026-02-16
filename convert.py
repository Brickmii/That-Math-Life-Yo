"""
Convert That_Math_Life_Yo.docx into a static HTML book in docs/.
"""

import os
import re
import shutil
from html import escape
from docx import Document

DOCX_PATH = os.path.join(os.path.dirname(__file__), "That_Math_Life_Yo.docx")
DOCS_DIR = os.path.join(os.path.dirname(__file__), "docs")
BASE_URL = "/That-Math-Life-Yo"


# ── Section definitions ──────────────────────────────────────────────

SECTIONS = [
    ("preface",         "1. Preface (about me and math)",           "Preface",                    "about me and math"),
    ("glossary",        "3. Term glossary",                         "Term Glossary",              "a cheat sheet for the casino"),
    ("intro",           "4. Intro (about math and cool shit)",      "Intro",                      "about math and cool shit"),
    ("arithmetic",      "5. Arithmetic (about function of values)", "Arithmetic",                 "about function of values"),
    ("algebra",         "6. Algebra (about value of functions)",    "Algebra",                    "about value of functions"),
    ("geometry",        "7. Geometry (about math and shapes)",      "Geometry",                   "about math and shapes"),
    ("set-theory",      "Set Theory (about sets and elements)",     "Set Theory",                 "about sets and elements"),
    ("trigonometry",    "9. Trigonometry (the fuckin triangle yo)",  "Trigonometry",               "the fuckin triangle yo"),
    ("calculus",        "10. Calculus (don\u2019t be afraid of change)", "Calculus",                   "don't be afraid of change"),
    ("statistics",      "11. Statistics & Probabilities (un+certainty)", "Statistics & Probabilities", "un+certainty"),
    ("linear-algebra",  "12. Linear Algebra (Matrix Re-lative)",    "Linear Algebra",             "Matrix Re-lative"),
    ("boolean-algebra", "13. Boolean Algebra (Not this and that)",  "Boolean Algebra",            "Not this and that"),
    ("topology",        "14. Topology (definitely knot this)",      "Topology",                   "definitely knot this"),
    ("game-theory",     "15. Game Theory (we go2 casino)",          "Game Theory",                "we go2 casino"),
]

NAV_ORDER = [s[0] for s in SECTIONS]
NAV_LABELS = {s[0]: s[2] for s in SECTIONS}


# ── Paragraph helpers ────────────────────────────────────────────────

def runs_to_html(paragraph):
    """Convert a paragraph's runs into HTML, preserving bold/italic."""
    parts = []
    for run in paragraph.runs:
        text = escape(run.text)
        if not text:
            continue
        if run.bold and run.italic:
            text = f"<strong><em>{text}</em></strong>"
        elif run.bold:
            text = f"<strong>{text}</strong>"
        elif run.italic:
            text = f"<em>{text}</em>"
        parts.append(text)
    return "".join(parts) or escape(paragraph.text)


def is_section_title(paragraph, section_titles):
    """Check if a paragraph matches a known section title."""
    text = paragraph.text.strip()
    return text in section_titles


def is_subsection_heading(paragraph):
    """Detect sub-section headings (all-bold, short, not a numbered list item)."""
    text = paragraph.text.strip()
    if not text or len(text) > 100:
        return False
    if not paragraph.runs:
        return False
    has_bold_text = False
    for r in paragraph.runs:
        if r.text.strip():
            if not r.bold:
                return False
            has_bold_text = True
    if not has_bold_text:
        return False
    # Exclude numbered list items like "1. There are..."
    if re.match(r"^\d+\.\s+[A-Z].*\w{10,}", text):
        return False
    return True


# ── Document parsing ─────────────────────────────────────────────────

def normalize_text(text):
    """Normalize smart quotes and dashes for matching."""
    return (text
            .replace("\u2018", "'").replace("\u2019", "'")
            .replace("\u201c", '"').replace("\u201d", '"')
            .replace("\u2013", "-").replace("\u2014", "-"))


def split_into_sections(doc):
    """Split document paragraphs into named sections."""
    # Build lookup: title text -> section slug
    title_to_slug = {}
    for s in SECTIONS:
        title_to_slug[s[1]] = s[0]
        normed = normalize_text(s[1])
        if normed != s[1]:
            title_to_slug[normed] = s[0]

    # Add TOC as a boundary marker (will be discarded later)
    title_to_slug["2. Table of contents"] = "_toc"

    paragraphs = list(doc.paragraphs)

    # Pass 1: find the paragraph index for each section title.
    # For duplicates, keep the LAST occurrence (real content, not TOC).
    slug_to_index = {}
    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        normed = normalize_text(text)
        slug = title_to_slug.get(text) or title_to_slug.get(normed)
        if slug:
            slug_to_index[slug] = i  # last occurrence wins

    # Sort by paragraph index
    ordered = sorted(slug_to_index.items(), key=lambda x: x[1])

    # Pass 2: slice paragraphs between section starts
    sections = {}
    for idx, (slug, start_i) in enumerate(ordered):
        if slug == "_toc":
            continue  # skip TOC boundary
        if idx + 1 < len(ordered):
            end_i = ordered[idx + 1][1]
        else:
            end_i = len(paragraphs)
        sections[slug] = paragraphs[start_i:end_i]

    return sections


def section_to_html(paragraphs, skip_title=True):
    """Convert a list of paragraphs into HTML body content."""
    html_parts = []
    skipped = 0
    skip_count = 1 if skip_title else 0

    # Also skip placeholder lines between calculus and statistics
    skip_lines = {
        "Statistics & Probabilities (un+certainty)",
        "Linear Algebra (Matrix Re-lative)",
        "Boolean Algebra (Not this and that)",
        "Topology (definitely knot this)",
        "Game Theory (we go2 casino)",
        "Statistics & Probabilies (un+certainty)",
        "Linear Algebra (Relativity Nativity)",
    }

    for para in paragraphs:
        text = para.text.strip()
        if not text:
            continue

        if skipped < skip_count:
            skipped += 1
            continue

        # Skip placeholder TOC-like lines that appear between chapters
        if text in skip_lines:
            continue

        # Sub-section heading
        if is_subsection_heading(para):
            html_parts.append(f"<h2>{escape(text)}</h2>")
            continue

        # Regular paragraph
        content = runs_to_html(para)
        if text.startswith("\u2022") or text.startswith("- "):
            html_parts.append(f'<p class="list-item">{content}</p>')
        else:
            html_parts.append(f"<p>{content}</p>")

    return "\n".join(html_parts)


# ── HTML Templates ───────────────────────────────────────────────────

def page_html(title, body_content, page_id, subtitle=None):
    """Wrap body content in a full HTML page."""
    idx = NAV_ORDER.index(page_id) if page_id in NAV_ORDER else -1
    prev_link = ""
    next_link = ""
    if idx > 0:
        prev_id = NAV_ORDER[idx - 1]
        prev_link = f'<a href="{BASE_URL}/{prev_id}.html">&larr; {NAV_LABELS[prev_id]}</a>'
    if 0 <= idx < len(NAV_ORDER) - 1:
        next_id = NAV_ORDER[idx + 1]
        next_link = f'<a href="{BASE_URL}/{next_id}.html">{NAV_LABELS[next_id]} &rarr;</a>'

    subtitle_html = f'<p class="subtitle">{escape(subtitle)}</p>' if subtitle else ""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{escape(title)} — That Math Life Yo</title>
<link rel="stylesheet" href="{BASE_URL}/style.css">
</head>
<body>
<header>
  <nav class="top-nav">
    <a href="{BASE_URL}/" class="nav-title">That Math Life Yo</a>
    <div class="nav-links">
      <a href="{BASE_URL}/">Contents</a>
      <a href="{BASE_URL}/That_Math_Life_Yo.docx" class="download-link">Download .docx</a>
    </div>
  </nav>
</header>
<main>
  <h1 class="page-title">{escape(title)}</h1>
  {subtitle_html}
  {body_content}
</main>
<footer>
  <nav class="chapter-nav">
    <div class="nav-prev">{prev_link}</div>
    <div class="nav-toc"><a href="{BASE_URL}/">Table of Contents</a></div>
    <div class="nav-next">{next_link}</div>
  </nav>
</footer>
</body>
</html>
"""


def index_html():
    """Generate the cover/index page."""
    toc_items = []
    for slug, _, title, subtitle in SECTIONS:
        toc_items.append(
            f'<li><a href="{BASE_URL}/{slug}.html">'
            f"{title} <span class=\"chapter-subtitle\">({subtitle})</span></a></li>"
        )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>That Math Life Yo</title>
<link rel="stylesheet" href="{BASE_URL}/style.css">
</head>
<body>
<main class="cover">
  <h1 class="book-title">That Math Life Yo</h1>
  <p class="book-subtitle">Math is fucking dope.</p>
  <div class="download-section">
    <a href="{BASE_URL}/That_Math_Life_Yo.docx" class="btn-download">Download Original (.docx)</a>
  </div>
  <nav class="toc">
    <h2>Contents</h2>
    <ol>
      {"".join(toc_items)}
    </ol>
  </nav>
</main>
</body>
</html>
"""


# ── Main ──────────────────────────────────────────────────────────────

def main():
    os.makedirs(DOCS_DIR, exist_ok=True)

    print("Reading .docx ...")
    doc = Document(DOCX_PATH)
    sections = split_into_sections(doc)

    # Index
    print("Writing index.html")
    with open(os.path.join(DOCS_DIR, "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html())

    # Sections
    for slug, _, title, subtitle in SECTIONS:
        print(f"Writing {slug}.html")
        if slug not in sections:
            print(f"  WARNING: {slug} not found in document!")
            continue

        body = section_to_html(sections[slug], skip_title=True)
        with open(os.path.join(DOCS_DIR, f"{slug}.html"), "w", encoding="utf-8") as f:
            f.write(page_html(title, body, slug, subtitle=subtitle))

    # Copy .docx
    dest = os.path.join(DOCS_DIR, "That_Math_Life_Yo.docx")
    shutil.copy2(DOCX_PATH, dest)
    print(f"Copied .docx to {dest}")
    print("Done!")


if __name__ == "__main__":
    main()