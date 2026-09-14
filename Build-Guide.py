"""
Rigenera ExchangeHealthCheck-Guide.pdf a partire da GUIDE.md.

Dipendenze (pure Python, nessun binario esterno richiesto):
    python -m pip install markdown xhtml2pdf

Uso:
    python Build-Guide.py

Va rilanciato ogni volta che GUIDE.md cambia: il PDF non si aggiorna da solo.

Nota per chi modifica GUIDE.md: <pre> di xhtml2pdf non implementa
word-wrap/overflow-wrap e una riga di codice piu larga della pagina viene
TAGLIATA in silenzio, non va a capo. Tenere le righe nei blocchi ```
sotto le ~80-85 colonne (vedi il commento sulla regola CSS "pre" qui sotto);
per un comando lungo, spezzarlo su piu righe con la pipeline PowerShell o
l'operatore backtick di continuazione.
"""

import io
import os

import markdown
from xhtml2pdf import pisa

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE_DIR, "GUIDE.md")
OUT = os.path.join(BASE_DIR, "ExchangeHealthCheck-Guide.pdf")

with io.open(SRC, "r", encoding="utf-8") as f:
    md_text = f.read()

body_html = markdown.markdown(
    md_text,
    extensions=["tables", "fenced_code", "toc", "sane_lists"],
)

CSS = """
@page {
    size: A4;
    margin: 2.2cm 2cm 2.4cm 2cm;
    @frame footer_frame {
        -pdf-frame-content: footer_content;
        bottom: 1cm; margin-left: 2cm; margin-right: 2cm; height: 1cm;
    }
}
body {
    font-family: "Helvetica", "Arial", sans-serif;
    font-size: 9.5pt;
    line-height: 1.45;
    color: #1c2530;
}
h1 {
    font-size: 19pt;
    color: #14324d;
    border-bottom: 2px solid #2f6690;
    padding-bottom: 6px;
    margin-top: 0;
    margin-bottom: 14px;
}
h2 {
    font-size: 13.5pt;
    color: #14324d;
    margin-top: 22px;
    margin-bottom: 8px;
    border-bottom: 0.75px solid #c7d2db;
    padding-bottom: 3px;
}
h3 {
    font-size: 11pt;
    color: #1c4b6e;
    margin-top: 14px;
    margin-bottom: 5px;
}
p { margin: 5px 0 9px 0; text-align: left; }
strong { color: #0f2438; }
ul, ol { margin: 4px 0 10px 18px; padding-left: 8px; }
li { margin-bottom: 3px; }
hr { border: none; border-top: 1px solid #c7d2db; margin: 18px 0; }
a { color: #2f6690; }

code {
    font-family: "Courier New", monospace;
    font-size: 8.7pt;
    background-color: #eef2f5;
    padding: 1px 4px;
    border-radius: 2px;
    color: #0f2438;
}
pre {
    font-family: "Courier New", monospace;
    font-size: 8.5pt;
    background-color: #f3f6f8;
    border: 0.75px solid #d5dee5;
    border-left: 3px solid #2f6690;
    padding: 7px 10px;
    margin: 6px 0 12px 0;
    line-height: 1.35;
    /* <pre> di default non va a capo (white-space: pre): una riga piu lunga
       della larghezza stampabile viene tagliata da xhtml2pdf invece di
       avvolgersi, perdendo silenziosamente il resto del comando. Con
       pre-wrap + break-word va a capo restando comunque preformattato. */
    white-space: pre-wrap;
    word-wrap: break-word;
    overflow-wrap: break-word;
}
pre code { background-color: transparent; padding: 0; }

table {
    border-collapse: collapse;
    width: 100%;
    margin: 8px 0 14px 0;
    font-size: 8.7pt;
}
th {
    background-color: #14324d;
    color: #ffffff;
    text-align: left;
    padding: 5px 7px;
    border: 0.75px solid #14324d;
}
td {
    padding: 5px 7px;
    border: 0.75px solid #d5dee5;
    vertical-align: top;
}
tr:nth-child(even) td { background-color: #f6f9fa; }

#footer_content {
    font-size: 7.5pt;
    color: #7c8a96;
    text-align: center;
    border-top: 0.5px solid #d5dee5;
    padding-top: 4px;
}
"""

html = """<html><head><meta charset="utf-8"/><style>{css}</style></head>
<body>
{body}
<div id="footer_content">Exchange Health Check &mdash; Guida utente &middot; pagina <pdf:pagenumber/> di <pdf:pagecount/></div>
</body></html>""".format(css=CSS, body=body_html)

with io.open(OUT, "wb") as out_file:
    result = pisa.CreatePDF(src=html, dest=out_file, encoding="utf-8")

print("Errori xhtml2pdf:", result.err)
print("PDF scritto in:", OUT)
