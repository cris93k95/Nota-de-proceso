import os
from pathlib import Path

import PyPDF2

BASE_DIR = Path(__file__).resolve().parent
SOURCE_DIR = BASE_DIR / "Portafolio Docente 2026"
OUTPUT_DIR = BASE_DIR / "pdf_extracts"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Priority files to extract
priority_files = [
    "Manual 7° y 8° Ed. Básica y Educación Media 2025.pdf",
    "Manual Ed. Básica Asignaturas 2025.pdf",
    "Manual Ed. Básica Generalista 2025.pdf",
    "Manual Ed. Media Técnico Profesional 2025.pdf",
    "Manual Ed. Personas Jóvenes y Adultas 2025.pdf",
    "Manual Educación Especial Escuela Especial (NEEP) 2025.pdf",
    "Manual Educación Especial Escuela Regular 2025.pdf",
    "Manual Educación Parvularia 2025.pdf",
    "Rúbricas 7 Básico a 4 Medio Asignaturas 2025.pdf",
    "Rubricas Educacion Basic a 1° a 6° BÁSICO 2025.pdf",
    "Rúbricas Generalista 2025.pdf",
    "Rubricas Educacion Media T écnico Profesional 2025.pdf",
    "Rubricas Educacion de Person as Jóvenes y Adultas 2025.pdf",
    "Rubricas Educacion Especial Escu ela Regular versión corregida 2025.pdf",
    "Rúbricas Educación Especial Escuela Especial 2025.pdf",
    "Rúbricas Educación Parvularia_2025.pdf",
    "MBE-2021.pdf",
]

for filename in priority_files:
    filepath = SOURCE_DIR / filename
    if not filepath.exists():
        print(f"NOT FOUND: {filename}")
        continue
    
    try:
        reader = PyPDF2.PdfReader(str(filepath))
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
            text += "\n---PAGE---\n"
        
        out_name = filename.replace(".pdf", ".txt")
        out_path = OUTPUT_DIR / out_name
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"OK: {filename} -> {len(text)} chars, {len(reader.pages)} pages")
    except Exception as e:
        print(f"ERROR: {filename} -> {e}")

print("\nDone!")
