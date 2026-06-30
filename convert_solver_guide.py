#!/usr/bin/env python3
import markdown
from weasyprint import HTML, CSS

with open("Solver_Formulas_Guide.md", "r") as f:
    md_text = f.read()

html_body = markdown.markdown(md_text, extensions=["tables", "fenced_code"])

css = CSS(string="""
@page {
    size: A4;
    margin: 2cm 2.2cm 2cm 2.2cm;
}
body {
    font-family: Calibri, Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.65;
    color: #1a1a1a;
}
h1 {
    background: #1F4E79;
    color: white;
    padding: 14px 18px;
    font-size: 17pt;
    margin: 0 0 22px 0;
    border-radius: 4px;
}
h2 {
    color: #1F4E79;
    border-bottom: 2.5px solid #2E75B6;
    padding-bottom: 5px;
    font-size: 13pt;
    margin-top: 28px;
    margin-bottom: 10px;
}
h3 {
    color: #2E75B6;
    font-size: 11pt;
    margin-top: 18px;
    margin-bottom: 6px;
}
p { margin: 6px 0 10px 0; }
table {
    border-collapse: collapse;
    width: 100%;
    margin: 12px 0 16px 0;
    font-size: 10pt;
}
th {
    background: #2E75B6;
    color: white;
    padding: 8px 12px;
    text-align: left;
    font-weight: bold;
}
td {
    padding: 6px 12px;
    border: 1px solid #BDD7EE;
    vertical-align: top;
}
tr:nth-child(even) td { background: #EBF3FA; }
code {
    background: #F0F4F8;
    border: 1px solid #C8D8E8;
    padding: 2px 5px;
    font-family: Consolas, "Courier New", monospace;
    font-size: 9.5pt;
    border-radius: 3px;
    color: #1F4E79;
}
pre {
    background: #F0F4F8;
    border-left: 5px solid #2E75B6;
    padding: 12px 16px;
    margin: 12px 0;
    font-size: 9.5pt;
    border-radius: 0 4px 4px 0;
    page-break-inside: avoid;
}
pre code {
    background: none;
    border: none;
    padding: 0;
    color: #1a1a1a;
    font-size: 9.5pt;
}
blockquote {
    border-left: 4px solid #ED7D31;
    margin: 10px 0;
    padding: 8px 14px;
    background: #FFF9F0;
    color: #555;
    font-style: italic;
}
hr {
    border: none;
    border-top: 1.5px solid #BDD7EE;
    margin: 22px 0;
}
ul, ol { margin: 8px 0; padding-left: 22px; }
li { margin-bottom: 4px; }
strong { color: #1F4E79; }
""")

html_full = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body>{html_body}</body></html>"""

HTML(string=html_full).write_pdf("Solver_Formulas_Guide.pdf", stylesheets=[css])
print("Done: Solver_Formulas_Guide.pdf")
