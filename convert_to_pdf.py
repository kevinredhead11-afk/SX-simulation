#!/usr/bin/env python3
import markdown
from weasyprint import HTML, CSS

with open("SX_Model_Guide.md", "r") as f:
    md_text = f.read()

html_body = markdown.markdown(md_text, extensions=["tables", "fenced_code"])

css = CSS(string="""
@page {
    size: A4;
    margin: 2cm 2.5cm 2cm 2.5cm;
}
body {
    font-family: Calibri, Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.6;
    color: #1a1a1a;
}
h1 {
    background: #1F4E79;
    color: white;
    padding: 12px 16px;
    font-size: 18pt;
    margin: 0 0 20px 0;
    border-radius: 4px;
}
h2 {
    color: #1F4E79;
    border-bottom: 2px solid #2E75B6;
    padding-bottom: 4px;
    font-size: 13pt;
    margin-top: 24px;
}
h3 {
    color: #2E75B6;
    font-size: 11pt;
    margin-top: 16px;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin: 12px 0;
    font-size: 10pt;
}
th {
    background: #2E75B6;
    color: white;
    padding: 7px 10px;
    text-align: left;
}
td {
    padding: 6px 10px;
    border: 1px solid #BDD7EE;
}
tr:nth-child(even) td {
    background: #EBF3FA;
}
code {
    background: #F2F2F2;
    border: 1px solid #CCCCCC;
    padding: 1px 4px;
    font-family: Consolas, Courier New, monospace;
    font-size: 9.5pt;
    border-radius: 3px;
}
pre {
    background: #F2F2F2;
    border-left: 4px solid #2E75B6;
    padding: 10px 14px;
    margin: 10px 0;
    overflow-x: auto;
    font-size: 9pt;
}
pre code {
    background: none;
    border: none;
    padding: 0;
}
blockquote {
    border-left: 4px solid #ED7D31;
    margin: 10px 0;
    padding: 6px 14px;
    background: #FFF9F0;
    color: #555;
}
hr {
    border: none;
    border-top: 1px solid #BDD7EE;
    margin: 20px 0;
}
""")

html_full = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body>{html_body}</body></html>"""

HTML(string=html_full).write_pdf("SX_Model_Guide.pdf", stylesheets=[css])
print("Done: SX_Model_Guide.pdf")
