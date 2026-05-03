#!/usr/bin/env python3
"""
Genera la Prueba de Comprensión Lectora - Unidad 1 - 1ro Medio
Formato HTML para impresión (tamaño carta), con exportación PDF y Word.
"""

import base64
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Institutional header used in the assessment start section.
LOGO_PATH = next(
    (
        path
        for path in (
            os.path.join(SCRIPT_DIR, '_logo_header_resized.png'),
            os.path.join(SCRIPT_DIR, 'Logo Colegio.png'),
        )
        if os.path.exists(path)
    ),
    None,
)
LOGO_B64 = ""
if LOGO_PATH:
    with open(LOGO_PATH, 'rb') as f:
        LOGO_B64 = base64.b64encode(f.read()).decode('ascii')

TOTAL_POINTS = 40
PASSING_PERCENTAGE = 60
TEST_TITLE = "Prueba de Comprensión Lectora — Inglés — Unidad 1 — 1° Medio"
TEST_OBJECTIVE = (
    "Evaluar la comprensión lectora y el uso de Present Perfect, Used to y Passive Voice "
    "en contextos técnicos vinculados a las especialidades de 1° medio."
)
TEST_SKILLS = "Comprensión lectora — vocabulario en contexto — gramática aplicada"
HEADER_SUMMARY_ITEMS = [
    ("Asignatura", "Idioma Extranjero: Inglés"),
    ("Curso", "1° Medio TP"),
    ("Evaluación", "Nota 2 — U1"),
    ("Puntaje total", f"{TOTAL_POINTS} puntos"),
    ("Exigencia", f"{PASSING_PERCENTAGE}%"),
]
GENERAL_INSTRUCTIONS = [
    "Lee cada texto con atención antes de responder.",
    "Todas las respuestas deben basarse en la información entregada en los textos.",
    "Marca solo una alternativa por pregunta, usando lápiz pasta azul o negro.",
    f"La prueba tiene {TOTAL_POINTS} puntos en total y una exigencia de {PASSING_PERCENTAGE}%.",
    "Tienes 60 minutos para completar la evaluación.",
    "No se permite el uso de diccionario ni dispositivos electrónicos.",
]

# ============================================================
# TEST CONTENT
# ============================================================

# PART I: Reading Comprehension - 5 texts (1 per specialty), 3 questions each, 5 options
# The texts explicitly include Present Perfect, Used to and Passive Voice forms.
READING_TEXTS = [
    {
        "title": "Text 1 — Automotive Mechanics",
        "text": (
            "My name is Felipe and I study Automotive Mechanics in Rancagua. "
            "I am in first year and I enjoy the workshop classes. "
            "Every Monday, the tools are prepared by the teacher before the lesson starts. "
            "A small car is checked and the oil level is reviewed with the group. "
            "I have worked with brake parts since March, and I have used the scanner for six weeks. "
            "When I was younger, I used to watch my father repair engines at home. "
            "Now I read simple service notes and answer questions about the cars."
        ),
        "questions": [
            {
                "num": 1,
                "q": "What is prepared by the teacher before the lesson starts?",
                "options": [
                    "A) The brake parts",
                    "B) The tools",
                    "C) The service notes",
                    "D) The student uniforms",
                    "E) The engine belts"
                ],
                "answer": "B"
            },
            {
                "num": 2,
                "q": "How long has Felipe used the scanner?",
                "options": [
                    "A) Since March",
                    "B) For six weeks",
                    "C) For one year",
                    "D) Since primary school",
                    "E) For two days"
                ],
                "answer": "B"
            },
            {
                "num": 3,
                "q": "What did Felipe use to do when he was younger?",
                "options": [
                    "A) He used to write service notes",
                    "B) He used to check the oil level alone",
                    "C) He used to watch his father repair engines",
                    "D) He used to drive the small car",
                    "E) He used to clean the workshop every day"
                ],
                "answer": "C"
            }
        ]
    },
    {
        "title": "Text 2 — Industrial Mechanics",
        "text": (
            "Camila studies Industrial Mechanics in Concepción and she likes practical tasks. "
            "In the workshop, metal pieces are cleaned before they are welded. "
            "A welding helmet and safety boots are worn in every class. "
            "She has practiced basic welding since April, and she has joined the workshop team for two months. "
            "Before this school, she used to study in a regular classroom. "
            "She also used to feel nervous around large machines. "
            "Now she follows the process card and writes short reports after each task."
        ),
        "questions": [
            {
                "num": 4,
                "q": "What happens to the metal pieces before welding?",
                "options": [
                    "A) They are painted",
                    "B) They are measured",
                    "C) They are cleaned",
                    "D) They are cut in half",
                    "E) They are stored outside"
                ],
                "answer": "C"
            },
            {
                "num": 5,
                "q": "Since when has Camila practiced basic welding?",
                "options": [
                    "A) Since April",
                    "B) For two months",
                    "C) Since last week",
                    "D) For one semester",
                    "E) Since January"
                ],
                "answer": "A"
            },
            {
                "num": 6,
                "q": "How did Camila use to feel around large machines?",
                "options": [
                    "A) Proud",
                    "B) Nervous",
                    "C) Angry",
                    "D) Bored",
                    "E) Relaxed"
                ],
                "answer": "B"
            }
        ]
    },
    {
        "title": "Text 3 — Electricity",
        "text": (
            "Diego studies Electricity in Santiago and he likes installation work. "
            "In each lesson, the main switch is turned off before the circuits are tested. "
            "The cables are labeled with small tags, so the students can work safely. "
            "Diego has read circuit diagrams since May, and he has installed simple lights for three weeks. "
            "When he was a child, he used to help his uncle after school. "
            "His uncle used to show him how houses were wired step by step. "
            "Now Diego wants to work in building installation."
        ),
        "questions": [
            {
                "num": 7,
                "q": "What is turned off before the circuits are tested?",
                "options": [
                    "A) The main switch",
                    "B) The classroom lights",
                    "C) The wall fan",
                    "D) The soldering iron",
                    "E) The computer screen"
                ],
                "answer": "A"
            },
            {
                "num": 8,
                "q": "Since when has Diego read circuit diagrams?",
                "options": [
                    "A) Since March",
                    "B) Since May",
                    "C) For three weeks",
                    "D) Since childhood",
                    "E) Since Monday"
                ],
                "answer": "B"
            },
            {
                "num": 9,
                "q": "What did Diego's uncle use to show him?",
                "options": [
                    "A) How school machines were repaired",
                    "B) How a motor was cleaned",
                    "C) How houses were wired step by step",
                    "D) How posters were printed",
                    "E) How reports were written"
                ],
                "answer": "C"
            }
        ]
    },
    {
        "title": "Text 4 — Electronics",
        "text": (
            "Valentina studies Electronics in Valparaíso and she enjoys the lab. "
            "In the room, small parts are placed in labeled boxes and each circuit is tested after it is assembled. "
            "She has built five alarm circuits since the semester started, and she has used a soldering iron for two months. "
            "Before entering the technical school, she used to spend most afternoons playing video games. "
            "She also used to think electronics was too difficult for her. "
            "Now she programs simple devices and checks each result on the screen."
        ),
        "questions": [
            {
                "num": 10,
                "q": "What is tested after it is assembled?",
                "options": [
                    "A) Each circuit",
                    "B) Each notebook",
                    "C) Each safety boot",
                    "D) Each report",
                    "E) Each table"
                ],
                "answer": "A"
            },
            {
                "num": 11,
                "q": "How many alarm circuits has Valentina built since the semester started?",
                "options": [
                    "A) Three",
                    "B) Four",
                    "C) Five",
                    "D) Six",
                    "E) Seven"
                ],
                "answer": "C"
            },
            {
                "num": 12,
                "q": "What did Valentina use to do most afternoons before entering the technical school?",
                "options": [
                    "A) She used to read circuit diagrams",
                    "B) She used to play video games",
                    "C) She used to write service notes",
                    "D) She used to clean metal pieces",
                    "E) She used to print posters"
                ],
                "answer": "B"
            }
        ]
    },
    {
        "title": "Text 5 — Graphic Design",
        "text": (
            "Matías studies Graphic Design in Temuco and he likes visual projects. "
            "In the computer lab, posters are designed for school events and the final versions are printed on thick paper. "
            "The colors are reviewed by the group before the work is presented. "
            "Matías has created ten posters since March, and he has worked on the school campaign for four weeks. "
            "In primary school, he used to draw in every notebook and he used to copy comic characters at home. "
            "Now he likes making simple designs for real clients."
        ),
        "questions": [
            {
                "num": 13,
                "q": "What is reviewed by the group before the work is presented?",
                "options": [
                    "A) The safety rules",
                    "B) The colors",
                    "C) The brake parts",
                    "D) The building plan",
                    "E) The labels"
                ],
                "answer": "B"
            },
            {
                "num": 14,
                "q": "Since when has Matías created ten posters?",
                "options": [
                    "A) Since January",
                    "B) For four weeks",
                    "C) Since March",
                    "D) Since last Friday",
                    "E) For two months"
                ],
                "answer": "C"
            },
            {
                "num": 15,
                "q": "What did Matías use to copy at home in primary school?",
                "options": [
                    "A) Comic characters",
                    "B) Circuit diagrams",
                    "C) Service notes",
                    "D) Safety labels",
                    "E) Factory reports"
                ],
                "answer": "A"
            }
        ]
    }
]

# PART II: Vocabulary Matching (terminos from the texts)
VOCAB_MATCHING = {
    "instructions": (
        "Questions 16–25<br/>"
        "Match each English term from the texts (16–25) with its correct Spanish definition (A–L).<br/>"
        "Write the correct letter on the line. There are TWO extra definitions you do not need."
    ),
    "terms": [
        {"num": 16, "term": "scanner"},
        {"num": 17, "term": "brake parts"},
        {"num": 18, "term": "welding helmet"},
        {"num": 19, "term": "process card"},
        {"num": 20, "term": "tag"},
        {"num": 21, "term": "main switch"},
        {"num": 22, "term": "soldering iron"},
        {"num": 23, "term": "labeled box"},
        {"num": 24, "term": "thick paper"},
        {"num": 25, "term": "service notes"},
    ],
    "definitions": [
        {"letter": "A", "definition": "Registro breve con información de mantenimiento o reparación"},
        {"letter": "B", "definition": "Interruptor principal que corta la energía"},
        {"letter": "C", "definition": "Etiqueta pequeña con información"},
        {"letter": "D", "definition": "Herramienta caliente para unir componentes electrónicos"},
        {"letter": "E", "definition": "Piezas del sistema de freno"},
        {"letter": "F", "definition": "Protección que se usa para soldar"},
        {"letter": "G", "definition": "Caja marcada con nombre o categoría"},
        {"letter": "H", "definition": "Aparato para revisar datos del vehículo"},
        {"letter": "I", "definition": "Hoja guía con los pasos de un trabajo"},
        {"letter": "J", "definition": "Papel más resistente que el papel común"},
        {"letter": "K", "definition": "Panel donde se registran las notas del curso"},
        {"letter": "L", "definition": "Herramienta para medir voltaje eléctrico"},
    ],
    "answers": {
        16: "H",
        17: "E",
        18: "F",
        19: "I",
        20: "C",
        21: "B",
        22: "D",
        23: "G",
        24: "J",
        25: "A",
    }
}

# PART III: Grammar Activities
# Activity 1: Present Perfect in context - 5 questions
# Activity 2: Used to in context - 5 questions
# Activity 3: Passive Voice in context - 5 questions

GRAMMAR_ACTIVITIES = [
    {
        "title": "Activity A — Present Perfect in Context",
        "instructions": (
            "Questions 26–30<br/>"
            "Use the information from the texts to choose the correct Present Perfect form."
        ),
        "questions": [
            {
                "num": 26,
                "q": "Felipe _______ with brake parts since March.",
                "options": [
                    "A) has worked",
                    "B) have worked",
                    "C) has work",
                    "D) worked",
                    "E) is worked"
                ],
                "answer": "A"
            },
            {
                "num": 27,
                "q": "Camila _______ basic welding since April.",
                "options": [
                    "A) has practiced",
                    "B) have practiced",
                    "C) practiced",
                    "D) has practice",
                    "E) is practiced"
                ],
                "answer": "A"
            },
            {
                "num": 28,
                "q": "Diego _______ simple lights for three weeks.",
                "options": [
                    "A) has install",
                    "B) have installed",
                    "C) has installed",
                    "D) installed",
                    "E) is installing"
                ],
                "answer": "C"
            },
            {
                "num": 29,
                "q": "Valentina _______ five alarm circuits since the semester started.",
                "options": [
                    "A) has build",
                    "B) built",
                    "C) have built",
                    "D) has built",
                    "E) is built"
                ],
                "answer": "D"
            },
            {
                "num": 30,
                "q": "Matías _______ on the school campaign for four weeks.",
                "options": [
                    "A) has worked",
                    "B) worked",
                    "C) have worked",
                    "D) has work",
                    "E) is worked"
                ],
                "answer": "A"
            }
        ]
    },
    {
        "title": "Activity B — Used to in Context",
        "instructions": (
            "Questions 31–35<br/>"
            "Use the information from the texts to choose the correct form with used to."
        ),
        "questions": [
            {
                "num": 31,
                "q": "Felipe _______ his father repair engines at home.",
                "options": [
                    "A) used to watch",
                    "B) use to watch",
                    "C) used watching",
                    "D) used to watched",
                    "E) uses to watch"
                ],
                "answer": "A"
            },
            {
                "num": 32,
                "q": "Camila _______ nervous around large machines.",
                "options": [
                    "A) used feeling",
                    "B) used to feel",
                    "C) use to feel",
                    "D) used to felt",
                    "E) uses to feel"
                ],
                "answer": "B"
            },
            {
                "num": 33,
                "q": "Diego _______ his uncle after school.",
                "options": [
                    "A) used helping",
                    "B) use to help",
                    "C) used to help",
                    "D) used to helped",
                    "E) uses to help"
                ],
                "answer": "C"
            },
            {
                "num": 34,
                "q": "Valentina _______ most afternoons playing video games.",
                "options": [
                    "A) used to spend",
                    "B) use to spend",
                    "C) used spending",
                    "D) uses to spend",
                    "E) used to spent"
                ],
                "answer": "A"
            },
            {
                "num": 35,
                "q": "Matías _______ comic characters at home.",
                "options": [
                    "A) used copying",
                    "B) used to copied",
                    "C) uses to copy",
                    "D) used to copy",
                    "E) use to copy"
                ],
                "answer": "D"
            }
        ]
    },
    {
        "title": "Activity C — Passive Voice in Context",
        "instructions": (
            "Questions 36–40<br/>"
            "Use the information from the texts to choose the correct Passive Voice form."
        ),
        "questions": [
            {
                "num": 36,
                "q": "Before Felipe's lesson starts, a small car _______.",
                "options": [
                    "A) are checked",
                    "B) checks",
                    "C) is checked",
                    "D) is check",
                    "E) checked"
                ],
                "answer": "C"
            },
            {
                "num": 37,
                "q": "In Camila's workshop, metal pieces _______ before they are welded.",
                "options": [
                    "A) cleans",
                    "B) is cleaned",
                    "C) cleaned",
                    "D) are cleaned",
                    "E) are clean"
                ],
                "answer": "D"
            },
            {
                "num": 38,
                "q": "In Diego's class, the main switch _______ before the circuits are tested.",
                "options": [
                    "A) turn off",
                    "B) is turning off",
                    "C) are turned off",
                    "D) turned off",
                    "E) is turned off"
                ],
                "answer": "E"
            },
            {
                "num": 39,
                "q": "In Valentina's lab, each circuit _______ after it is assembled.",
                "options": [
                    "A) are tested",
                    "B) tests",
                    "C) tested",
                    "D) is test",
                    "E) is tested"
                ],
                "answer": "E"
            },
            {
                "num": 40,
                "q": "In Matías's class, the final versions _______ on thick paper.",
                "options": [
                    "A) is printed",
                    "B) print",
                    "C) are printed",
                    "D) printed",
                    "E) are print"
                ],
                "answer": "C"
            }
        ]
    }
]

# ============================================================
# HTML GENERATION
# ============================================================

def generate_html():
    # Collect all answers for answer key
    all_answers = {}
    for text in READING_TEXTS:
        for q in text["questions"]:
            all_answers[q["num"]] = q["answer"]
    for num, ans in VOCAB_MATCHING["answers"].items():
        all_answers[num] = ans
    for activity in GRAMMAR_ACTIVITIES:
        for q in activity["questions"]:
            all_answers[q["num"]] = q["answer"]

    # Build questions HTML for reading
    reading_html = ""
    for text in READING_TEXTS:
        reading_html += f'<div class="reading-text"><span class="text-title">{text["title"]}</span><br/>{text["text"]}</div>\n'
        for q in text["questions"]:
            opts = "".join(f'<div class="option">{o}</div>' for o in q["options"])
            reading_html += f'<div class="question"><p class="q-text">{q["num"]}. {q["q"]}</p>{opts}</div>\n'

    # Build vocabulary matching HTML
    terms_rows = ""
    for t in VOCAB_MATCHING["terms"]:
        terms_rows += f'<tr><td style="font-weight:bold; width:35px; text-align:center;">{t["num"]}.</td><td style="width:180px;">{t["term"]}</td><td style="width:40px; border-bottom:1px solid #999;"></td></tr>\n'

    defs_rows = ""
    for d in VOCAB_MATCHING["definitions"]:
        defs_rows += f'<tr><td style="font-weight:bold; width:25px;">{d["letter"]})</td><td>{d["definition"]}</td></tr>\n'

    vocab_html = f"""
<div style="display:flex; gap:18px; align-items:flex-start;">
<div>
<table style="border-collapse:collapse; font-size:8.5pt;">
<tr><th colspan="3" style="text-align:left; padding-bottom:4px; font-size:8pt; color:#333;">English Terms</th></tr>
{terms_rows}
</table>
</div>
<div>
<table style="border-collapse:collapse; font-size:8.5pt;">
<tr><th colspan="2" style="text-align:left; padding-bottom:4px; font-size:8pt; color:#333;">Spanish Definitions</th></tr>
{defs_rows}
</table>
</div>
</div>
"""

    # Build grammar HTML
    grammar_html = ""
    for activity in GRAMMAR_ACTIVITIES:
        grammar_html += f'<div class="part-header" style="margin-top:6px;">{activity["title"]}</div>\n'
        grammar_html += f'<p class="instructions">{activity["instructions"]}</p>\n'
        for q in activity["questions"]:
            opts = "".join(f'<div class="option">{o}</div>' for o in q["options"])
            grammar_html += f'<div class="question"><p class="q-text">{q["num"]}. {q["q"]}</p>{opts}</div>\n'

    # Build answer key
    answer_items = ""
    for num in sorted(all_answers.keys()):
        answer_items += f'<div class="answer-item"><span class="ans-num">{num}.</span> {all_answers[num]}</div>\n'

    header_summary = "".join(
        f'<span><strong>{label}:</strong> {value}</span>' for label, value in HEADER_SUMMARY_ITEMS
    )
    instruction_items = "".join(f'<li>{item}</li>' for item in GENERAL_INSTRUCTIONS)

    total_points = TOTAL_POINTS

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8"/>
<meta content="width=device-width, initial-scale=1.0" name="viewport"/>
<title>{TEST_TITLE}</title>
<style>
    @page {{
        size: 21.59cm 27.94cm;
        margin: 1.5cm 1.8cm 1.5cm 1.8cm;
    }}

    * {{ margin: 0; padding: 0; box-sizing: border-box; }}

    body {{
        font-family: Arial, Helvetica, sans-serif;
        font-size: 9pt;
        line-height: 1.3;
        color: #1a1a1a;
        background: #f0f0f0;
        padding: 0;
    }}

    .page {{
        background: white;
        max-width: 21.59cm;
        margin: 1cm auto;
        padding: 0.5cm 1.8cm 1.5cm 1.8cm;
        box-shadow: 0 2px 12px rgba(0,0,0,0.12);
    }}

    .logo-container {{
        text-align: center;
        margin: 0;
        padding: 0;
    }}
    .logo-container img {{
        width: 100%;
        display: block;
        margin-bottom: -4px;
    }}

    .test-title {{
        text-align: center;
        font-size: 11pt;
        font-weight: bold;
        margin: 6px 0 4px 0;
        padding: 0;
    }}

    .header-panel {{
        margin-bottom: 8px;
    }}
    .meta-summary {{
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 4px 16px;
        padding: 5px 0;
        border-top: 1px solid #0f4c81;
        border-bottom: 1px solid #0f4c81;
        font-size: 8.2pt;
    }}
    .student-row {{
        display: flex;
        flex-wrap: wrap;
        justify-content: space-between;
        gap: 4px 18px;
        padding: 5px 0 4px 0;
        font-size: 8.2pt;
        border-bottom: 1px solid #d9d9d9;
    }}
    .meta-line {{
        margin-top: 4px;
        font-size: 8pt;
        line-height: 1.35;
    }}
    .instructions-block {{
        margin-top: 4px;
        padding-bottom: 4px;
        border-bottom: 1px solid #d9d9d9;
    }}
    .instructions-title {{
        font-size: 8pt;
        font-weight: bold;
        margin-bottom: 2px;
    }}
    .instructions-list {{
        margin: 0;
        padding-left: 16px;
        font-size: 7.8pt;
        line-height: 1.35;
    }}
    .instructions-list li {{
        margin-bottom: 1px;
    }}

    .part {{
        margin-bottom: 8px;
    }}
    .part-intro {{
        page-break-inside: avoid;
    }}
    .part-header {{
        background: #2962FF;
        color: white;
        padding: 3px 8px;
        font-weight: bold;
        font-size: 9pt;
        margin-bottom: 3px;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }}
    .instructions {{
        font-style: italic;
        font-size: 8pt;
        color: #333;
        margin-bottom: 4px;
        line-height: 1.35;
    }}

    .reading-text {{
        background: #f5f5f5;
        border: 1px solid #ddd;
        padding: 6px 8px;
        margin-bottom: 6px;
        font-size: 8.5pt;
        line-height: 1.35;
        page-break-inside: avoid;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }}
    .text-title {{
        font-weight: bold;
        font-size: 9pt;
        margin-bottom: 2px;
    }}

    .question {{
        margin-bottom: 4px;
        page-break-inside: avoid;
    }}
    .q-text {{
        font-weight: bold;
        font-size: 8.5pt;
        margin-bottom: 1px;
    }}
    .option {{
        padding-left: 14px;
        font-size: 8pt;
        line-height: 1.35;
    }}
    .options-inline {{
        padding-left: 14px;
        font-size: 8pt;
        color: #333;
    }}

    .test-footer {{
        text-align: center;
        font-style: italic;
        font-size: 7.5pt;
        color: #888;
        margin-top: 10px;
        padding-top: 6px;
        border-top: 1px solid #ddd;
    }}

    .answer-key {{
        page-break-before: always;
        margin-top: 0;
        padding-top: 10px;
        page-break-inside: avoid;
    }}
    .answer-key-title {{
        text-align: center;
        font-size: 10pt;
        font-weight: bold;
        margin-bottom: 6px;
    }}
    .answer-key-subtitle {{
        text-align: center;
        font-size: 8.5pt;
        color: #555;
        margin-bottom: 8px;
    }}
    .answer-grid {{
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 2px 12px;
        font-size: 8pt;
        margin: 0 auto;
    }}
    .answer-item {{
        padding: 1px 0;
    }}
    .answer-item .ans-num {{
        font-weight: bold;
        display: inline-block;
        min-width: 22px;
    }}

    .score-table {{
        width: 100%;
        border-collapse: collapse;
        margin-top: 12px;
        font-size: 8pt;
    }}
    .score-table td, .score-table th {{
        border: 1px solid #000;
        padding: 3px 6px;
        text-align: center;
    }}
    .score-table th {{
        background: #e8e8e8;
        font-weight: bold;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }}

    @media print {{
        body {{
            background: white;
            padding: 0;
            margin: 0;
        }}
        .page {{
            max-width: none;
            margin: 0;
            padding: 0;
            box-shadow: none;
        }}
        .part {{
            page-break-inside: auto;
        }}
        .part-intro {{
            page-break-inside: avoid;
        }}
        .reading-text {{
            page-break-inside: avoid;
        }}
        .question {{
            page-break-inside: avoid;
        }}
        .part-header {{
            page-break-after: avoid;
            background: #2962FF !important;
            color: white !important;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }}
        .reading-text {{
            background: #f5f5f5 !important;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }}
    }}
</style>
</head>
<body>
<div class="page">
<div class="header-panel">
<div class="logo-container">
<img src="data:image/png;base64,{LOGO_B64}" alt="Encabezado institucional"/>
</div>
<div class="test-title">{TEST_TITLE}</div>
<div class="meta-summary">{header_summary}</div>
<div class="student-row">
<span><strong>Nombre:</strong> ________________________________</span>
<span><strong>Curso:</strong> __________</span>
<span><strong>Fecha:</strong> ___/___/2026</span>
</div>
<p class="meta-line"><strong>Objetivo:</strong> {TEST_OBJECTIVE}</p>
<p class="meta-line"><strong>Habilidades:</strong> {TEST_SKILLS}</p>
<div class="instructions-block">
<p class="instructions-title">Instrucciones generales</p>
<ul class="instructions-list">{instruction_items}</ul>
</div>
</div>

<!-- ==================== PART I: READING COMPREHENSION ==================== -->
<div class="part">
<div class="part-intro">
<div class="part-header">PART I — READING COMPREHENSION (15 points)</div>
<p class="instructions">Questions 1–15<br/>
Read each text carefully and choose the correct answer for each question.<br/>
Circle the letter of your answer (A, B, C, D or E).</p>
</div>
{reading_html}
</div>

<!-- ==================== PART II: VOCABULARY MATCHING ==================== -->
<div class="part">
<div class="part-intro">
<div class="part-header">PART II — VOCABULARY: Matching / Términos Pareados (10 points)</div>
<p class="instructions">{VOCAB_MATCHING["instructions"]}</p>
</div>
{vocab_html}
</div>

<!-- ==================== PART III: GRAMMAR ==================== -->
<div class="part">
<div class="part-intro">
<div class="part-header">PART III — GRAMMAR (15 points)</div>
<p class="instructions">Questions 26–40<br/>
Use the information from Part I to complete each sentence correctly.<br/>
Circle the letter of your answer (A, B, C, D or E).</p>
</div>
{grammar_html}
</div>

<!-- FOOTER -->
<div class="test-footer">— Fin de la evaluación — Revisa tus respuestas antes de entregar. —</div>
</div>

<!-- ==================== ANSWER KEY ==================== -->
<div class="page answer-key">
<div class="answer-key-title">Pauta de Respuestas Correctas</div>
<div class="answer-key-subtitle">Prueba de Comprensión Lectora — Unidad 1 — 1° Medio</div>
<div class="answer-grid">
{answer_items}
</div>

<!-- SCORING TABLE -->
<table class="score-table" style="margin-top:20px;">
<tr>
<th colspan="6">Tabla de Conversión de Puntaje a Nota (60% de exigencia)</th>
</tr>
<tr>
<th>Puntaje</th><th>Nota</th><th>Puntaje</th><th>Nota</th><th>Puntaje</th><th>Nota</th>
</tr>
"""

    # Generate score table with 60% exigency
    # passing score = 60% of 40 = 24
    passing = total_points * 0.6
    scores = []
    for pts in range(0, total_points + 1):
        if pts < passing:
            nota = 1.0 + 3.0 * (pts / passing)
        else:
            nota = 4.0 + 3.0 * ((pts - passing) / (total_points - passing))
        nota = min(nota, 7.0)
        nota = round(nota, 1)
        scores.append((pts, nota))

    # Build score table rows (3 columns of pairs)
    rows_per_col = (len(scores) + 2) // 3
    for i in range(rows_per_col):
        html += "<tr>"
        for col in range(3):
            idx = i + col * rows_per_col
            if idx < len(scores):
                pts, nota = scores[idx]
                style = ' style="background:#e8f5e9; -webkit-print-color-adjust:exact; print-color-adjust:exact;"' if nota >= 4.0 else ''
                html += f'<td>{pts}</td><td{style}>{nota}</td>'
            else:
                html += '<td></td><td></td>'
        html += "</tr>\n"

    html += """</table>
</div>
</body>
</html>"""

    return html


if __name__ == "__main__":
    output_dir = os.path.join(SCRIPT_DIR, "PLANIFICACIONES_2026_LISTO_IMPRESION")
    os.makedirs(output_dir, exist_ok=True)

    html_content = generate_html()
    output_path = os.path.join(output_dir, "PRUEBA_COMPRENSION_LECTORA_U1_1RO_MEDIO.html")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"HTML generado: {output_path}")
    print(f"Tamaño: {len(html_content):,} bytes")
    print("Abre el archivo en el navegador y usa Ctrl+P para imprimir a PDF.")
