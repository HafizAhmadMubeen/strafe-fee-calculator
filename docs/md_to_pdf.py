r"""
md_to_pdf.py  -  tiny Markdown -> PDF converter (no third-party packages).

Handles the subset used by the project docs: headings, bold, inline code,
bullet / numbered lists, GFM pipe tables, blockquotes, horizontal rules, links.
Renders to HTML, then prints to PDF with headless Microsoft Edge.

    .venv\Scripts\python docs\md_to_pdf.py "docs\User Guide.md"
        -> writes  docs\User Guide.pdf
"""
from __future__ import annotations

import html
import os
import re
import subprocess
import sys
import tempfile

CSS = """
@page { size: A4; margin: 18mm 16mm; }
* { box-sizing: border-box; }
body { font: 11pt/1.5 "Segoe UI", "Helvetica Neue", Arial, sans-serif;
       color: #1f2933; margin: 0; }
h1 { font-size: 21pt; margin: 0 0 4pt; color: #16324f; }
h2 { font-size: 15pt; margin: 20pt 0 6pt; padding-bottom: 3pt;
     border-bottom: 1px solid #d8dee4; color: #16324f; }
h3 { font-size: 12.5pt; margin: 14pt 0 4pt; color: #24313d; }
p  { margin: 6pt 0; }
ul, ol { margin: 6pt 0 6pt 0; padding-left: 22pt; }
li { margin: 3pt 0; }
code { font-family: Consolas, "Courier New", monospace; font-size: 9.5pt;
       background: #f2f4f6; padding: 1px 4px; border-radius: 3px; }
blockquote { margin: 8pt 0; padding: 4pt 12pt; color: #4b5563;
             border-left: 3px solid #c9d2da; background: #f7f9fa; }
hr { border: 0; border-top: 1px solid #d8dee4; margin: 16pt 0; }
table { border-collapse: collapse; width: 100%; margin: 8pt 0; font-size: 10pt; }
th, td { border: 1px solid #d0d7de; padding: 5pt 8pt; text-align: left;
         vertical-align: top; }
th { background: #eef2f5; }
tr, table, h1, h2, h3 { page-break-inside: avoid; }
h1, h2, h3 { page-break-after: avoid; }
a { color: #1d4ed8; text-decoration: none; }
"""


def _inline(text: str) -> str:
    text = html.escape(text, quote=False)
    text = re.sub(r"`([^`]+)`", lambda m: f"<code>{m.group(1)}</code>", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"(?<!\*)\*([^*\n]+)\*(?!\*)", r"<em>\1</em>", text)
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r'<a href="\2">\1</a>', text)
    return text


def _split_row(line: str):
    line = line.strip().strip("|")
    return [c.strip() for c in line.split("|")]


_H  = re.compile(r"^(#{1,6})\s+(.*)$")
_HR = re.compile(r"^(-{3,}|\*{3,}|_{3,})$")
_UL = re.compile(r"^(\s*)([-*+])\s+(.*)$")
_OL = re.compile(r"^(\s*)(\d+)[.)]\s+(.*)$")
_TSEP = re.compile(r"^\s*\|?[\s:|-]+\|[\s:|-]*$")


def _is_block_start(line: str, nxt: str | None) -> bool:
    s = line.strip()
    if not s:
        return True
    if _H.match(s) or _HR.match(s) or s.startswith(">"):
        return True
    if _UL.match(line) or _OL.match(line):
        return True
    if "|" in line and nxt is not None and _TSEP.match(nxt):
        return True
    return False


def md_to_html_body(md: str) -> str:
    lines = md.replace("\r\n", "\n").split("\n")
    out: list[str] = []
    i, n = 0, len(lines)
    list_stack: list[str] = []

    def close_lists():
        while list_stack:
            out.append(f"</{list_stack.pop()}>")

    def cont(first_text: str, start: int) -> tuple[str, int]:
        """first_text + any soft-wrapped continuation lines, joined with spaces."""
        parts = [first_text]
        j = start
        while j < n:
            nxt = lines[j + 1] if j + 1 < n else None
            if not lines[j].strip() or _is_block_start(lines[j], nxt):
                break
            parts.append(lines[j].strip())
            j += 1
        return " ".join(parts), j

    while i < n:
        line = lines[i]
        stripped = line.strip()

        if "|" in line and i + 1 < n and _TSEP.match(lines[i + 1]):
            close_lists()
            headers = _split_row(line)
            out.append("<table><thead><tr>"
                       + "".join(f"<th>{_inline(h)}</th>" for h in headers)
                       + "</tr></thead><tbody>")
            i += 2
            while i < n and "|" in lines[i] and lines[i].strip():
                cells = _split_row(lines[i])
                cells += [""] * (len(headers) - len(cells))
                out.append("<tr>" + "".join(f"<td>{_inline(c)}</td>"
                                            for c in cells[:len(headers)]) + "</tr>")
                i += 1
            out.append("</tbody></table>")
            continue

        if not stripped:
            close_lists()
            i += 1
            continue

        m = _H.match(stripped)
        if m:
            close_lists()
            lvl = len(m.group(1))
            out.append(f"<h{lvl}>{_inline(m.group(2))}</h{lvl}>")
            i += 1
            continue

        if _HR.match(stripped):
            close_lists()
            out.append("<hr>")
            i += 1
            continue

        if stripped.startswith(">"):
            close_lists()
            buf = []
            while i < n and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip().lstrip(">").strip())
                i += 1
            out.append(f"<blockquote>{_inline(' '.join(buf))}</blockquote>")
            continue

        m = _UL.match(line)
        if m:
            if not list_stack or list_stack[-1] != "ul":
                close_lists()
                list_stack.append("ul")
                out.append("<ul>")
            text, i = cont(m.group(3), i + 1)
            out.append(f"<li>{_inline(text)}</li>")
            continue

        m = _OL.match(line)
        if m:
            if not list_stack or list_stack[-1] != "ol":
                close_lists()
                list_stack.append("ol")
                out.append("<ol>")
            text, i = cont(m.group(3), i + 1)
            out.append(f"<li>{_inline(text)}</li>")
            continue

        close_lists()
        text, i = cont(stripped, i + 1)
        out.append(f"<p>{_inline(text)}</p>")

    close_lists()
    return "\n".join(out)


def find_edge() -> str:
    for p in (
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%LocalAppData%\Microsoft\Edge\Application\msedge.exe"),
    ):
        if os.path.isfile(p):
            return p
    raise SystemExit("Could not find msedge.exe - Microsoft Edge is required.")


def convert(md_path: str, pdf_path: str | None = None) -> str:
    md_path = os.path.abspath(md_path)
    if pdf_path is None:
        pdf_path = os.path.splitext(md_path)[0] + ".pdf"
    pdf_path = os.path.abspath(pdf_path)

    with open(md_path, "r", encoding="utf-8") as f:
        md = f.read()
    title = html.escape(os.path.splitext(os.path.basename(md_path))[0])
    doc = (f"<!doctype html><html><head><meta charset='utf-8'><title>{title}</title>"
           f"<style>{CSS}</style></head><body>{md_to_html_body(md)}</body></html>")

    tmp_dir = tempfile.mkdtemp(prefix="md2pdf_")
    html_path = os.path.join(tmp_dir, "doc.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(doc)

    url = "file:///" + html_path.replace("\\", "/")
    edge = find_edge()
    cmd = [edge, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
           f"--print-to-pdf={pdf_path}", url]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
    if not os.path.isfile(pdf_path):
        # older Edge builds use a different flag name
        cmd[3] = "--print-to-pdf-no-header"
        subprocess.run(cmd, capture_output=True, text=True, timeout=90)
    if not os.path.isfile(pdf_path):
        raise SystemExit(f"Edge did not produce the PDF.\n{r.stdout}\n{r.stderr}")
    return pdf_path


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(__file__), "User Guide.md")
    dst = sys.argv[2] if len(sys.argv) > 2 else None
    out = convert(src, dst)
    print("wrote", out)
