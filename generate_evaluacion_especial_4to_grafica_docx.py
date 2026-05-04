#!/usr/bin/env python3
"""
Genera una evaluacion especial de comprension lectora para 4to Medio - Grafica.

La prueba incluye 2 textos tecnicos adaptados a nivel B1:
- proceso de impresion
- diferencias entre formatos CMYK y RGB

Cada texto incluye 6 preguntas de seleccion multiple con 5 alternativas,
equilibrando preguntas explicitas e implicitas.
"""

from pathlib import Path

from generate_evaluacion_especial_1ro_semestre_docx import (
    BASE_DIR,
    add_body_paragraph,
    add_box_heading,
    add_bullets,
    add_info_table,
    add_logo,
    add_title,
    create_document,
    save_document,
)


OUTPUT_NAME = "EVALUACION_ESPECIAL_COMPRENSION_LECTORA_4TO_MEDIO_GRAFICA.docx"
OUTPUT_PATHS = [
    BASE_DIR / "PLANIFICACIONES_2026_LISTO_IMPRESION" / OUTPUT_NAME,
    BASE_DIR / "tranquiprofe.cl" / "static" / "recursos" / "materiales" / "4to-medio" / "instrumentos" / OUTPUT_NAME,
]


TEXTS = [
    {
        "heading": "Texto 1 - The Printing Process",
        "title": "From File to Finished Brochure",
        "paragraphs": [
            "Daniela works in a print shop that produces brochures for local companies. When a new job arrives, she begins in the prepress area. First, she opens the client file and checks the size, image resolution, and color profile. If a photo is too small or a font is missing, she fixes the problem before printing starts.",
            "After that, Daniela prepares a digital proof. This proof shows how the brochure should look on paper. The client reviews it and approves the design. Once the file is approved, the team creates the printing plates and installs them on the offset press. They choose offset printing because the order is large and the color quality must stay consistent in every brochure.",
            "During production, Daniela checks the ink density and paper position. She stops the press if she sees lines, stains, or incorrect color balance. When the printed sheets are dry, they move to postpress. There, the brochures are trimmed, folded, and stapled. Before delivery, Daniela makes a final inspection to confirm that the order is clean, complete, and ready for the client.",
        ],
        "questions": [
            {
                "number": 1,
                "type": "explicita",
                "prompt": "What does Daniela check first when a new job arrives?",
                "options": [
                    "A) The delivery time, truck number, and storage shelf.",
                    "B) The paper cost, ink amount, and machine oil level.",
                    "C) The file size, image resolution, and color profile.",
                    "D) The stapling settings, folding marks, and blade pressure.",
                    "E) The client invoice, order code, and package label.",
                ],
            },
            {
                "number": 2,
                "type": "explicita",
                "prompt": "Why does the team use offset printing for this brochure order?",
                "options": [
                    "A) Because that press is the only large machine available.",
                    "B) Because the brochures do not need strong color quality.",
                    "C) Because prepress cannot correct digital files before printing.",
                    "D) Because the client asked for plastic instead of paper.",
                    "E) Because the order is large and needs consistent color quality.",
                ],
            },
            {
                "number": 3,
                "type": "explicita",
                "prompt": "What happens to the brochures after the printed sheets are dry?",
                "options": [
                    "A) They return to prepress for another proof and review.",
                    "B) They are trimmed, folded, and stapled in postpress.",
                    "C) They are scanned again to improve image resolution.",
                    "D) They are mixed with fresh paper for another order.",
                    "E) They go straight to the client without inspection.",
                ],
            },
            {
                "number": 4,
                "type": "explicita",
                "prompt": "Why does Daniela prepare a digital proof before production?",
                "options": [
                    "A) To reduce later corrections during trimming and folding.",
                    "B) To confirm paper size and sheet weight before printing.",
                    "C) To align plates and rollers before the press run begins.",
                    "D) To show the brochure on paper before the press run starts.",
                    "E) To prepare an online version before the print job starts.",
                ],
            },
            {
                "number": 5,
                "type": "implicita",
                "prompt": "What can be inferred about the final inspection?",
                "options": [
                    "A) It lowers the risk of sending mistakes to the client.",
                    "B) It focuses mainly on black ink and other dark areas.",
                    "C) It replaces the review done earlier in prepress.",
                    "D) It happens only if the client checks the order.",
                    "E) It matters less than drying because errors can wait.",
                ],
            },
            {
                "number": 6,
                "type": "implicita",
                "prompt": "Which idea best describes Daniela's role in the print shop?",
                "options": [
                    "A) She mainly speaks with clients and tracks each order.",
                    "B) She mostly organizes deliveries and transport problems.",
                    "C) She checks many stages and notices details in production.",
                    "D) She spends most of the day folding finished brochures.",
                    "E) She prefers design tasks and leaves control to others.",
                ],
            },
        ],
        "answer_key": ["C", "E", "B", "D", "A", "C"],
    },
    {
        "heading": "Texto 2 - Color Systems in Graphic Arts",
        "title": "Why Designers Use RGB and CMYK",
        "paragraphs": [
            "Lucas is a graphic arts technician who often explains color systems to new students. He says that RGB and CMYK are both important, but they are used for different purposes. RGB stands for red, green, and blue. It is the color system used by screens such as phones, tablets, and computers. These devices create color with light, so RGB usually looks bright and vivid.",
            "CMYK stands for cyan, magenta, yellow, and black. This system is used in printing because printers place ink on paper instead of using light. When a designer makes a poster for social media, RGB is usually the best option. When the same poster must be printed, the file should be converted to CMYK before production begins.",
            "Lucas warns that some colors seen on a screen cannot be reproduced exactly on paper. Very bright blue, green, or neon tones may look softer after printing. For that reason, graphic technicians prepare a proof and compare the printed result with the original design. They adjust the color profile if necessary. Lucas believes that understanding both systems is essential because good design is not only about creativity, but also about knowing how the final product will look in real life.",
        ],
        "questions": [
            {
                "number": 7,
                "type": "explicita",
                "prompt": "Which color system is used by screens?",
                "options": [
                    "A) PMS, a system for matching special spot colors.",
                    "B) Grayscale, a mode for black and white images.",
                    "C) Offset, a process used with plates and paper.",
                    "D) RGB, the color system used by screens.",
                    "E) CMYK-Plus, an extended setting for some printers.",
                ],
            },
            {
                "number": 8,
                "type": "explicita",
                "prompt": "What does the letter K represent in CMYK?",
                "options": [
                    "A) Keyline, the guide used around a printed shape.",
                    "B) Black, the dark ink used in four-color printing.",
                    "C) Brown, a warm tone used in some paper samples.",
                    "D) Brightness, the screen effect added to digital images.",
                    "E) Contrast, the difference between light and dark areas.",
                ],
            },
            {
                "number": 9,
                "type": "explicita",
                "prompt": "What may happen if a file designed for screen use is printed without proper color adjustment?",
                "options": [
                    "A) The paper may absorb more ink than the team planned for.",
                    "B) The printer may change some file settings during the job.",
                    "C) The poster may print faster than expected in production.",
                    "D) The final poster may come out larger than the planned size.",
                    "E) The printed colors may not match the colors seen on screen.",
                ],
            },
            {
                "number": 10,
                "type": "explicita",
                "prompt": "Why do graphic technicians prepare a proof?",
                "options": [
                    "A) To compare the printed sample with the original design.",
                    "B) To choose a clearer name for the shop and its products.",
                    "C) To remove the color profile stored in the working file.",
                    "D) To skip the file conversion before print production.",
                    "E) To add brighter tones to the final printed poster.",
                ],
            },
            {
                "number": 11,
                "type": "implicita",
                "prompt": "What can be inferred about neon colors in printed work?",
                "options": [
                    "A) They may appear stronger on screens than on printed paper.",
                    "B) They may fade when black ink is added in the print job.",
                    "C) They may need changes because print cannot match neon tones.",
                    "D) They may work better in animation than in printed posters.",
                    "E) They may be used in any project prepared for CMYK printing.",
                ],
            },
            {
                "number": 12,
                "type": "implicita",
                "prompt": "What is Lucas's main message to new students?",
                "options": [
                    "A) Good print design needs less creativity than digital design.",
                    "B) Modern designers should use RGB instead of CMYK most times.",
                    "C) Color systems matter only after workers gain more experience.",
                    "D) Good designers should focus more on screens than on print.",
                    "E) Good technicians understand both screen and print systems.",
                ],
            },
        ],
        "answer_key": ["D", "B", "E", "A", "C", "E"],
    },
]


def add_question(doc, number, prompt, options):
    add_body_paragraph(doc, f"{number}. {prompt}", bold=True, size=8.5)
    for option in options:
        add_body_paragraph(doc, option, left_indent=0.5, size=8)


def add_answer_key(doc):
    doc.add_page_break()
    add_title(
        doc,
        "Pauta de Correccion",
        "Evaluacion Especial de Comprension Lectora - 4to Medio Grafica",
    )
    add_box_heading(doc, "1. Criterio de correccion", "1A237E")
    add_bullets(
        doc,
        [
            "Cada respuesta correcta vale 2 puntos.",
            "Puntaje maximo: 24 puntos.",
            "Exigencia institucional: 60%.",
            "No se consideran justificaciones del estudiante porque el instrumento es de seleccion multiple.",
        ],
    )

    for index, text in enumerate(TEXTS, start=1):
        add_box_heading(doc, f"{index + 1}. Respuestas correctas - {text['heading']}", "1A237E")
        for question, answer in zip(text["questions"], text["answer_key"]):
            add_body_paragraph(
                doc,
                f"{question['number']}. {answer} ({question['type']})",
                size=8.5,
            )


def build_document():
    doc = create_document()
    add_logo(doc)
    add_title(
        doc,
        "Evaluacion Especial de Comprension Lectora - Ingles - 4to Medio Grafica",
        "Especialidad: Grafica | Nivel adaptado B1 | Enfoque tecnico-profesional",
    )
    add_info_table(
        doc,
        evaluation_name="Evaluacion especial de comprension lectora",
        due_date="____ / ____ / 2026",
        total_points="24 puntos",
        products="2 textos / 12 preguntas de seleccion multiple",
        objective="Evaluar la comprension lectora en ingles de textos tecnicos vinculados a procesos de impresion y gestion del color en la especialidad de Grafica, mediante preguntas explicitas e implicitas adaptadas a nivel B1.",
        modality="Instrumento especial enfocado en la especialidad",
    )

    add_box_heading(doc, "1. Instrucciones generales", "1A237E")
    add_bullets(
        doc,
        [
            "Lee cada texto con atencion antes de responder.",
            "Cada pregunta tiene 5 alternativas. Marca solo una respuesta en cada caso.",
            "Las preguntas combinan informacion explicita e implicita.",
            "Puedes subrayar ideas importantes, palabras tecnicas y conectores mientras lees.",
            "No uses traductor automatico para responder. Puedes apoyarte en el vocabulario tecnico que ya conoces de la especialidad.",
        ],
    )

    add_box_heading(doc, "2. Habilidades evaluadas", "1A237E")
    add_bullets(
        doc,
        [
            "Comprension de informacion explicita en textos tecnicos.",
            "Inferencia de ideas y conclusiones a partir del contexto.",
            "Reconocimiento de vocabulario propio de la especialidad de Grafica.",
            "Analisis de procesos de impresion y gestion del color en situaciones reales de trabajo.",
        ],
    )

    for text in TEXTS:
        add_box_heading(doc, f"3. {text['heading']}" if text is TEXTS[0] else f"4. {text['heading']}", "1A237E")
        add_body_paragraph(doc, text["title"], bold=True, size=9)
        for paragraph in text["paragraphs"]:
            add_body_paragraph(doc, paragraph, size=8.5)
        add_body_paragraph(doc, "Questions", bold=True, size=8.5)
        for question in text["questions"]:
            add_question(doc, question["number"], question["prompt"], question["options"])

    add_box_heading(doc, "5. Escala de puntaje", "1A237E")
    add_bullets(
        doc,
        [
            "12 preguntas en total.",
            "2 puntos por respuesta correcta.",
            "Puntaje maximo: 24 puntos.",
            "Exigencia institucional: 60%.",
        ],
    )
    add_answer_key(doc)
    return doc


def main():
    document = build_document()
    save_document(document, OUTPUT_PATHS)


if __name__ == "__main__":
    main()