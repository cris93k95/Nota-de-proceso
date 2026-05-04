#!/usr/bin/env python3
"""Fix the PDF download function in all HTML material files."""
import os
import glob
import re

MATERIALES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "static", "recursos", "materiales")

OLD_SCRIPT = r"""<script>
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
</script>"""

NEW_SCRIPT = r"""<script>
function downloadPDF(){
  var b=document.getElementById('pdfBtn'),t=b.innerHTML;
  b.innerHTML='&#9203; Generando PDF...';b.disabled=true;
  var w=document.createElement('div');
  w.style.cssText='width:900px;background:#fff;position:absolute;left:-9999px;top:0;';
  var hdr=document.querySelector('.header').cloneNode(true);
  hdr.style.margin='0';
  w.appendChild(hdr);
  var ctn=document.querySelector('.container').cloneNode(true);
  ctn.style.cssText='max-width:900px;margin:0;padding:12px 15px;';
  w.appendChild(ctn);
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
  document.body.appendChild(w);
  html2pdf().set({
    margin:[8,8,10,8],
    filename:document.title+'.pdf',
    image:{type:'jpeg',quality:0.95},
    html2canvas:{scale:2,useCORS:true,logging:false,scrollY:0,windowWidth:960},
    jsPDF:{unit:'mm',format:'letter',orientation:'portrait'},
    pagebreak:{mode:['avoid-all','css'],avoid:'.vocab-card,.quiz-item,.objective-card,.concept-box,.oral-box,.ticket-box,.score-bar,.reading-box,.grammar-box,.safety-box,.eval-box,.section'}
  }).from(w).save().then(function(){document.body.removeChild(w);b.innerHTML=t;b.disabled=false}).catch(function(e){console.error(e);if(w.parentNode)document.body.removeChild(w);b.innerHTML=t;b.disabled=false});
}
</script>"""


def main():
    count = 0
    skipped = 0
    for fpath in sorted(glob.glob(os.path.join(MATERIALES_DIR, "**", "*.html"), recursive=True)):
        with open(fpath, "r", encoding="utf-8") as f:
            html = f.read()
        if OLD_SCRIPT not in html:
            skipped += 1
            continue
        html = html.replace(OLD_SCRIPT, NEW_SCRIPT, 1)
        with open(fpath, "w", encoding="utf-8") as f:
            f.write(html)
        count += 1
        print(f"OK: {os.path.relpath(fpath, MATERIALES_DIR)}")
    print(f"\nDone: {count} fixed, {skipped} skipped")


if __name__ == "__main__":
    main()
