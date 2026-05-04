#!/usr/bin/env python3
"""Inject PDF download button into all HTML material files in Formulario Portafolio."""
import os
import glob

MATERIALES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "static", "recursos", "materiales")

# ── CSS to inject before </style> ──
PDF_CSS = """
/* PDF Download */
.pdf-btn-wrap{position:fixed;bottom:20px;right:20px;z-index:9999}
.pdf-btn{background:linear-gradient(135deg,#1565c0,#1976d2);color:#fff;border:none;padding:12px 22px;border-radius:30px;font-size:0.95em;font-weight:600;cursor:pointer;box-shadow:0 4px 15px rgba(21,101,192,.4);transition:all .3s;font-family:'Inter',sans-serif;display:flex;align-items:center;gap:8px}
.pdf-btn:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(21,101,192,.5)}
.pdf-btn:disabled{opacity:.7;cursor:wait;transform:none}
@media print{.pdf-btn-wrap{display:none!important}}
.objective-card,.concept-box,.oral-box,.reading-box,.grammar-box,.safety-box,.eval-box,.ticket-box,.quiz-item,.vocab-card,.score-bar{page-break-inside:avoid}
.section-header{page-break-after:avoid}
"""

# ── HTML + JS to inject before </body> ──
PDF_BODY = r"""<div class="pdf-btn-wrap"><button onclick="downloadPDF()" class="pdf-btn" id="pdfBtn">&#x1F4E5; Descargar PDF</button></div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
<script>
function downloadPDF(){
  var b=document.getElementById('pdfBtn'),t=b.innerHTML;
  b.innerHTML='&#9203; Generando PDF...';b.disabled=true;
  var w=document.createElement('div');
  w.appendChild(document.querySelector('.header').cloneNode(true));
  w.appendChild(document.querySelector('.container').cloneNode(true));
  w.querySelectorAll('.vocab-card').forEach(function(c){c.classList.add('revealed')});
  var rm=w.querySelector('.pdf-btn-wrap');if(rm)rm.remove();
  w.querySelectorAll('input.ticket-input,input[type="text"]').forEach(function(i){
    var d=document.createElement('div');
    d.style.cssText='border-bottom:1.5px solid #bbb;padding:8px 4px;margin-top:8px;min-height:26px;color:#888;font-size:0.88em;';
    d.textContent=i.placeholder||'';i.parentNode.replaceChild(d,i)});
  w.querySelectorAll('textarea').forEach(function(a){
    var d=document.createElement('div');
    d.style.cssText='border:1.5px solid #ccc;border-radius:8px;padding:10px;margin-top:8px;min-height:60px;color:#888;font-size:0.88em;';
    d.textContent=a.placeholder||'';a.parentNode.replaceChild(d,a)});
  html2pdf().set({
    margin:[10,10,12,10],
    filename:document.title+'.pdf',
    image:{type:'jpeg',quality:0.95},
    html2canvas:{scale:2,useCORS:true,logging:false},
    jsPDF:{unit:'mm',format:'letter',orientation:'portrait'},
    pagebreak:{mode:['avoid-all','css'],avoid:'.vocab-card,.quiz-item,.objective-card,.concept-box,.oral-box,.ticket-box,.score-bar,.reading-box,.grammar-box,.safety-box,.eval-box'}
  }).from(w).save().then(function(){b.innerHTML=t;b.disabled=false}).catch(function(e){console.error(e);b.innerHTML=t;b.disabled=false});
}
</script>
"""

def main():
    count = 0
    skipped = 0
    for fpath in sorted(glob.glob(os.path.join(MATERIALES_DIR, "**", "*.html"), recursive=True)):
        with open(fpath, "r", encoding="utf-8") as f:
            html = f.read()
        if "pdf-btn" in html:
            skipped += 1
            continue
        if "</style>" not in html or "</body>" not in html:
            print(f"SKIP (no style/body): {os.path.relpath(fpath, MATERIALES_DIR)}")
            skipped += 1
            continue
        html = html.replace("</style>", PDF_CSS + "\n</style>", 1)
        html = html.replace("</body>", PDF_BODY + "\n</body>", 1)
        with open(fpath, "w", encoding="utf-8") as f:
            f.write(html)
        count += 1
        print(f"OK: {os.path.relpath(fpath, MATERIALES_DIR)}")
    print(f"\nDone: {count} updated, {skipped} skipped")

if __name__ == "__main__":
    main()
