"""
Generate Academic Project Report — ML_Report.docx
===================================================
This script generates a clean, professionally formatted 15-page academic
report for the NLP Based Resume Parser and Matcher project using python-docx.

Run: python generate_report.py
"""

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement
import os


def set_cell_shading(cell, color):
    """Set background color for a table cell."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def add_page_break(doc):
    """Add a page break."""
    doc.add_page_break()


def set_paragraph_spacing(paragraph, before=0, after=0, line_spacing=1.15):
    """Set paragraph spacing."""
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line_spacing


def add_page_number_to_footer(section):
    """Add centered page number to a section's footer."""
    footer = section.footer
    footer.is_linked_to_previous = False
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Add PAGE field
    run = paragraph.add_run()
    run.font.size = Pt(10)
    run.font.name = 'Times New Roman'
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)

    run2 = paragraph.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' PAGE '
    run2._r.append(instrText)

    run3 = paragraph.add_run()
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run3._r.append(fldChar2)


def remove_footer(section):
    """Ensure a section has no footer content and is not linked to previous."""
    footer = section.footer
    footer.is_linked_to_previous = False
    for p in footer.paragraphs:
        for r in p.runs:
            r.text = ''


def add_heading_styled(doc, text, level=1):
    """Add a styled heading."""
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0, 51, 102)
    set_paragraph_spacing(heading, before=6, after=6)
    return heading


def add_body_paragraph(doc, text, bold=False, italic=False, font_size=11, alignment=None, space_after=6, space_before=0):
    """Add a body paragraph with consistent formatting."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(font_size)
    run.font.name = 'Times New Roman'
    run.bold = bold
    run.italic = italic
    if alignment:
        p.alignment = alignment
    set_paragraph_spacing(p, before=space_before, after=space_after, line_spacing=1.15)
    return p


def add_body_text(doc, text, space_after=6, space_before=0):
    """Add a standard body text paragraph."""
    return add_body_paragraph(doc, text, font_size=11, space_after=space_after, space_before=space_before)


def add_bullet(doc, text, level=0):
    """Add a bullet point."""
    p = doc.add_paragraph(text, style='List Bullet')
    for run in p.runs:
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
    if level > 0:
        p.paragraph_format.left_indent = Inches(0.5 * (level + 1))
    set_paragraph_spacing(p, before=2, after=2, line_spacing=1.15)
    return p


def create_simple_table(doc, headers, rows, col_widths=None):
    """Create a formatted table."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        run = p.add_run(header)
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = 'Times New Roman'
        run.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, "003366")

    # Data rows
    for r_idx, row_data in enumerate(rows):
        for c_idx, cell_text in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            run = p.add_run(str(cell_text))
            run.font.size = Pt(10)
            run.font.name = 'Times New Roman'
            if r_idx % 2 == 1:
                set_cell_shading(cell, "F2F2F2")

    if col_widths:
        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Inches(width)

    return table


def generate_report():
    """Generate the complete 15-page academic report."""
    doc = Document()

    # ================================================================
    # GLOBAL STYLE SETUP
    # ================================================================
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.space_after = Pt(6)

    # Set margins for all sections
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    # ================================================================
    # PAGE 1: TITLE PAGE
    # ================================================================

    # Add vertical spacing to center content
    for _ in range(4):
        p = doc.add_paragraph()
        set_paragraph_spacing(p, before=0, after=0)

    # Project Title
    p = doc.add_paragraph()
    run = p.add_run("NLP Based Resume Parser and Matcher")
    run.bold = True
    run.font.size = Pt(26)
    run.font.name = 'Times New Roman'
    run.font.color.rgb = RGBColor(0, 51, 102)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, before=0, after=6)

    # Subtitle
    p = doc.add_paragraph()
    run = p.add_run("A From Scratch Implementation Using TF IDF and Cosine Similarity")
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'
    run.font.color.rgb = RGBColor(80, 80, 80)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, before=0, after=24)

    # Horizontal line
    p = doc.add_paragraph()
    run = p.add_run("_" * 60)
    run.font.color.rgb = RGBColor(0, 51, 102)
    run.font.size = Pt(12)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, before=0, after=24)

    # Report type
    p = doc.add_paragraph()
    run = p.add_run("Project Report")
    run.bold = True
    run.font.size = Pt(16)
    run.font.name = 'Times New Roman'
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, before=0, after=6)

    # Course info
    p = doc.add_paragraph()
    run = p.add_run("Course: Machine Learning (CS/IT Elective)")
    run.font.size = Pt(13)
    run.font.name = 'Times New Roman'
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, before=0, after=24)

    # Institute Name
    p = doc.add_paragraph()
    run = p.add_run("Institute of Engineering and Technology")
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, before=0, after=2)

    p = doc.add_paragraph()
    run = p.add_run("Dr. A.P.J. Abdul Kalam Technical University, Lucknow")
    run.font.size = Pt(12)
    run.font.name = 'Times New Roman'
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, before=0, after=24)

    # Submitted By header
    p = doc.add_paragraph()
    run = p.add_run("Submitted By")
    run.bold = True
    run.font.size = Pt(13)
    run.font.name = 'Times New Roman'
    run.font.color.rgb = RGBColor(0, 51, 102)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, before=12, after=12)

    # Member Details Table
    member_table = doc.add_table(rows=5, cols=3)
    member_table.style = 'Table Grid'
    member_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers_row = member_table.rows[0]
    for i, h in enumerate(["S.No.", "Student Name", "Roll Number"]):
        cell = headers_row.cells[i]
        cell.text = ''
        run = cell.paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
        run.font.color.rgb = RGBColor(255, 255, 255)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, "003366")

    members = [
        ("1", "Shivam Kumar", "2100270120135"),
        ("2", "Priya Sharma", "2100270120098"),
        ("3", "Amit Verma", "2100270120015"),
        ("4", "Neha Singh", "2100270120078"),
    ]
    for r_idx, (sno, name, roll) in enumerate(members):
        for c_idx, val in enumerate([sno, name, roll]):
            cell = member_table.rows[r_idx + 1].cells[c_idx]
            cell.text = ''
            run = cell.paragraphs[0].add_run(val)
            run.font.size = Pt(11)
            run.font.name = 'Times New Roman'
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if r_idx % 2 == 1:
                set_cell_shading(cell, "F2F2F2")

    for row in member_table.rows:
        row.cells[0].width = Inches(0.8)
        row.cells[1].width = Inches(2.8)
        row.cells[2].width = Inches(2.4)

    # Session
    for _ in range(2):
        doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run("Academic Session: 2024 - 2025")
    run.font.size = Pt(12)
    run.font.name = 'Times New Roman'
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ================================================================
    # PAGE 2: TABLE OF CONTENTS
    # ================================================================
    add_page_break(doc)

    p = doc.add_paragraph()
    run = p.add_run("Table of Contents")
    run.bold = True
    run.font.size = Pt(20)
    run.font.name = 'Times New Roman'
    run.font.color.rgb = RGBColor(0, 51, 102)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, before=12, after=12)

    # Horizontal line
    p = doc.add_paragraph()
    run = p.add_run("_" * 72)
    run.font.color.rgb = RGBColor(0, 51, 102)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, before=0, after=12)

    toc_entries = [
        ("1.", "Introduction", "1"),
        ("2.", "Problem Statement", "2"),
        ("3.", "Objectives", "2"),
        ("4.", "Literature Review", "3"),
        ("5.", "System Analysis and Architecture", "4"),
        ("6.", "Proposed Methodology", "5"),
        ("7.", "Implementation Details", "7"),
        ("8.", "Results and Discussion", "10"),
        ("9.", "Conclusion", "11"),
        ("10.", "Future Scope", "12"),
        ("11.", "References", "13"),
    ]

    toc_table = doc.add_table(rows=len(toc_entries), cols=3)
    toc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, (num, title, page) in enumerate(toc_entries):
        # Number
        cell0 = toc_table.rows[r_idx].cells[0]
        cell0.text = ''
        run = cell0.paragraphs[0].add_run(num)
        run.bold = True
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'

        # Title
        cell1 = toc_table.rows[r_idx].cells[1]
        cell1.text = ''
        run = cell1.paragraphs[0].add_run(title)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'

        # Page number
        cell2 = toc_table.rows[r_idx].cells[2]
        cell2.text = ''
        run = cell2.paragraphs[0].add_run(page)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        cell2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # Row height
        for cell in toc_table.rows[r_idx].cells:
            set_paragraph_spacing(cell.paragraphs[0], before=4, after=4)

    toc_table.columns[0].width = Inches(0.6)
    toc_table.columns[1].width = Inches(4.5)
    toc_table.columns[2].width = Inches(0.8)

    # ================================================================
    # SECTION BREAK: Start new section for content pages with page numbers
    # ================================================================
    # Add a section break (new page) so content pages are in a separate section
    new_section = doc.add_section(WD_SECTION.NEW_PAGE)
    new_section.top_margin = Cm(2.54)
    new_section.bottom_margin = Cm(2.54)
    new_section.left_margin = Cm(2.54)
    new_section.right_margin = Cm(2.54)

    # Remove footer from the first section (title + TOC)
    remove_footer(doc.sections[0])

    # Add page numbers to the content section, starting from 1
    new_section.start_type = WD_SECTION.NEW_PAGE
    # Set page number restart to 1
    sectPr = new_section._sectPr
    pgNumType = OxmlElement('w:pgNumType')
    pgNumType.set(qn('w:start'), '1')
    sectPr.append(pgNumType)

    # Add page number in footer of content section
    add_page_number_to_footer(new_section)

    # ================================================================
    # PAGE 1 (content): INTRODUCTION
    # ================================================================
    add_heading_styled(doc, "1. Introduction", level=1)

    add_body_text(doc,
        "The process of hiring suitable candidates has always been one of the most resource intensive "
        "activities for any organization. With the rapid growth of online job portals and digital "
        "applications, the number of resumes received for a single job opening has increased dramatically. "
        "Human resource professionals and recruiters are often overwhelmed by the volume of applications "
        "and must spend considerable time manually reviewing each resume to determine its relevance to "
        "a specific job description. This manual process is not only slow but also prone to human error "
        "and subjective bias."
    )

    add_body_text(doc,
        "Natural Language Processing, a subfield of Artificial Intelligence, provides powerful techniques "
        "for understanding and analyzing human language in a computational manner. By applying NLP methods "
        "to the recruitment domain, it becomes possible to automate the extraction of relevant information "
        "from resumes and compare it systematically against job requirements. This project leverages NLP "
        "concepts to build a Resume Parser and Matcher system that can intelligently rank candidates "
        "based on the textual similarity between their resumes and a given job description."
    )

    add_body_text(doc,
        "The system presented in this report implements the entire NLP and matching pipeline from scratch, "
        "without relying on external machine learning libraries such as scikit learn, spaCy, or NLTK for "
        "the core algorithms. The fundamental techniques used include tokenization, stop word removal, "
        "Porter stemming, Term Frequency Inverse Document Frequency (TF IDF) vectorization, and cosine "
        "similarity computation. Each of these components has been coded manually using only basic Python "
        "constructs and standard mathematical libraries. This approach ensures a thorough understanding of "
        "the underlying algorithms and demonstrates the ability to implement them independently."
    )

    add_body_text(doc,
        "The project also includes a fully interactive web based dashboard built using Streamlit, which "
        "allows recruiters to upload resumes in PDF, DOCX, or TXT format, paste a job description, and "
        "instantly receive a ranked list of candidates along with detailed skill gap analysis, TF IDF "
        "vector breakdowns, and interactive visualizations. The combination of algorithmic depth and "
        "practical usability makes this project both an educational exercise and a functional prototype "
        "for real world resume screening."
    )

    # ================================================================
    # PAGE 4: PROBLEM STATEMENT + OBJECTIVES
    # ================================================================
    add_page_break(doc)
    add_heading_styled(doc, "2. Problem Statement", level=1)

    add_body_text(doc,
        "In the current job market, organizations receive hundreds or even thousands of applications "
        "for each open position. Manually screening these resumes is an extremely time consuming task "
        "that often leads to delays in the hiring cycle. Furthermore, human reviewers may overlook "
        "qualified candidates due to fatigue, bias, or inconsistent evaluation criteria. There is a "
        "pressing need for automated systems that can objectively assess the relevance of a candidate's "
        "resume against a specific job description."
    )

    add_body_text(doc,
        "While several Applicant Tracking Systems (ATS) exist in the market, most of them rely on "
        "proprietary algorithms or simple keyword matching that fails to capture the semantic relationships "
        "between terms. Additionally, these systems are typically expensive and opaque, offering little "
        "transparency into how candidates are scored and ranked."
    )

    add_body_text(doc,
        "The problem this project addresses is the development of a transparent, from scratch NLP based "
        "system that can parse resumes, extract meaningful features, and rank candidates by computing the "
        "mathematical similarity between their resumes and a given job description. The system must be "
        "explainable, meaning that recruiters should be able to understand exactly why a particular "
        "candidate received a specific score."
    )

    add_heading_styled(doc, "3. Objectives", level=1)

    add_body_text(doc,
        "The primary objectives of this project are outlined below. Each objective has been designed to "
        "ensure that the system is technically sound, practically useful, and academically rigorous."
    )

    objectives = [
        "To design and implement a complete text preprocessing pipeline from scratch, including tokenization, stop word removal, and stemming, without using any external NLP library.",
        "To implement the TF IDF vectorization algorithm from scratch using only basic Python constructs and the standard math library, enabling the conversion of textual documents into numerical vectors.",
        "To implement the cosine similarity metric from scratch for measuring the degree of similarity between two document vectors.",
        "To build a skill extraction and gap analysis module that identifies matched, missing, and additional skills by comparing resume content against the job description.",
        "To develop an interactive web based dashboard using Streamlit that supports resume upload in multiple formats (PDF, DOCX, TXT) and presents ranked results with detailed analysis.",
        "To validate the correctness of the from scratch implementation by comparing its output against the results produced by established libraries such as scikit learn.",
        "To provide full mathematical transparency by displaying intermediate computation steps including TF values, IDF values, TF IDF vectors, dot products, and per term contributions to the similarity score.",
    ]

    for obj in objectives:
        add_bullet(doc, obj)

    # ================================================================
    # PAGE 5: LITERATURE REVIEW
    # ================================================================
    add_page_break(doc)
    add_heading_styled(doc, "4. Literature Review", level=1)

    add_body_text(doc,
        "Resume parsing and matching is a well established area of research that lies at the intersection "
        "of Natural Language Processing, Information Retrieval, and Human Resource Analytics. This section "
        "reviews the key techniques and prior work that form the theoretical foundation of this project."
    )

    add_heading_styled(doc, "4.1 TF IDF in Information Retrieval", level=2)

    add_body_text(doc,
        "The Term Frequency Inverse Document Frequency (TF IDF) weighting scheme was first proposed by "
        "Karen Sparck Jones in 1972 and has since become one of the most widely used techniques in "
        "information retrieval and text mining. The core idea behind TF IDF is that the importance of "
        "a word in a document should be proportional to how frequently it appears in that document (Term "
        "Frequency) but inversely proportional to how commonly it appears across all documents in the "
        "corpus (Inverse Document Frequency). Words that are frequent within a specific document but rare "
        "across the collection are considered highly discriminative and receive higher TF IDF scores. This "
        "makes TF IDF particularly effective for identifying the most relevant terms in a resume or job "
        "description."
    )

    add_heading_styled(doc, "4.2 Cosine Similarity for Document Comparison", level=2)

    add_body_text(doc,
        "Cosine similarity is a metric used to measure how similar two documents are, regardless of their "
        "size. It calculates the cosine of the angle between two vectors in a multi dimensional space. "
        "When applied to TF IDF vectors, cosine similarity effectively captures the directional alignment "
        "of two documents rather than their magnitude. This property makes it ideal for comparing resumes "
        "of varying lengths against a standard job description. A cosine similarity of 1.0 indicates that "
        "the two documents are identical in terms of their word distribution, while a value of 0.0 "
        "indicates no overlap."
    )

    add_heading_styled(doc, "4.3 Porter Stemming Algorithm", level=2)

    add_body_text(doc,
        "The Porter Stemming Algorithm, published by Martin Porter in 1980, is one of the most widely "
        "used stemming algorithms for the English language. It reduces words to their root or base form "
        "by applying a series of suffix stripping rules organized into multiple steps. For example, the "
        "words \"running\", \"runner\", and \"runs\" are all reduced to the stem \"run\". This normalization "
        "is essential in NLP pipelines because it allows the system to treat morphological variants of "
        "the same word as a single entity, thereby improving the accuracy of document comparison."
    )

    add_heading_styled(doc, "4.4 Related Work in Resume Matching", level=2)

    add_body_text(doc,
        "Several prior systems have been developed for resume matching using various approaches. Some "
        "systems use pre trained word embeddings such as Word2Vec or BERT to capture semantic meaning, "
        "while others rely on simpler bag of words models with TF IDF weighting. Commercial Applicant "
        "Tracking Systems often use a combination of keyword matching, rule based parsing, and machine "
        "learning classifiers. However, most of these systems treat their algorithms as black boxes, "
        "providing little insight into how scores are computed. The present project differentiates itself "
        "by implementing all core algorithms from scratch and providing complete mathematical transparency "
        "at every stage of the pipeline."
    )

    # ================================================================
    # PAGE 6: SYSTEM ANALYSIS AND ARCHITECTURE
    # ================================================================
    add_page_break(doc)
    add_heading_styled(doc, "5. System Analysis and Architecture", level=1)

    add_heading_styled(doc, "5.1 System Overview", level=2)

    add_body_text(doc,
        "The NLP Based Resume Parser and Matcher is designed as a modular, layered system where each "
        "component performs a specific function in the overall pipeline. The architecture follows a "
        "clear separation of concerns, with distinct modules for text preprocessing, feature extraction, "
        "similarity computation, and user interface rendering. This modular design makes the system easy "
        "to understand, test, and extend."
    )

    add_heading_styled(doc, "5.2 System Architecture", level=2)

    add_body_text(doc,
        "The system architecture consists of four primary layers, each responsible for a distinct phase "
        "of the matching process. The flow of data through these layers is described below."
    )

    add_body_text(doc,
        "Input Layer: This layer handles the ingestion of raw data from the user. The recruiter provides "
        "a job description as plain text and uploads one or more resumes in PDF, DOCX, or TXT format. "
        "The file parsers in this layer extract raw text from the uploaded documents using utility "
        "libraries such as pdfplumber and python docx.",
        space_after=4
    )

    add_body_text(doc,
        "Preprocessing Layer: The extracted raw text passes through three sequential preprocessing stages. "
        "First, the tokenizer splits the text into individual word tokens after removing URLs, email "
        "addresses, phone numbers, and special characters. Second, the stop word filter removes commonly "
        "occurring words that carry little semantic value using a hand crafted list of over 120 words, "
        "including domain specific resume stop words. Third, the Porter stemmer reduces each remaining "
        "token to its root form by applying 25 suffix stripping rules.",
        space_after=4
    )

    add_body_text(doc,
        "Vectorization and Similarity Layer: The preprocessed tokens are converted into numerical TF IDF "
        "vectors. Term Frequency is computed for each document individually, while Inverse Document "
        "Frequency is computed across the entire corpus (all resumes plus the job description). The "
        "TF IDF score for each term is obtained by multiplying its TF and IDF values. Cosine similarity "
        "is then computed between the job description vector and each resume vector to produce a "
        "numerical similarity score.",
        space_after=4
    )

    add_body_text(doc,
        "Presentation Layer: The Streamlit dashboard receives the ranked results and presents them through "
        "four interactive tabs: Candidate Rankings, Skills Analysis, How It Works (mathematical "
        "transparency), and Visualizations. Recruiters can explore matched and missing skills, view "
        "TF IDF term breakdowns, and export results as CSV files."
    )

    add_heading_styled(doc, "5.3 Project Directory Structure", level=2)

    add_body_text(doc,
        "The project follows a well organized directory structure that separates source code, data, "
        "tests, and documentation into distinct folders. The main components are described below."
    )

    dir_rows = [
        ["app.py", "Streamlit dashboard and main entry point for the web application"],
        ["src/preprocessing/", "Tokenizer, stop words filter, and Porter stemmer modules"],
        ["src/vectorization/", "Term Frequency, Inverse Document Frequency, and TF IDF modules"],
        ["src/similarity/", "Cosine similarity and resume ranking pipeline modules"],
        ["src/parsers/", "PDF and DOCX text extraction utilities"],
        ["data/", "Sample resumes, job descriptions, and skills database (JSON)"],
        ["tests/", "Unit tests and validation tests against scikit learn"],
        ["notebooks/", "Jupyter notebook with step by step mathematical walkthrough"],
    ]
    create_simple_table(doc, ["Component", "Description"], dir_rows, [1.8, 4.5])

    # ================================================================
    # PAGES 7-8: PROPOSED METHODOLOGY
    # ================================================================
    add_page_break(doc)
    add_heading_styled(doc, "6. Proposed Methodology", level=1)

    add_body_text(doc,
        "The proposed methodology for the NLP Based Resume Parser and Matcher follows a systematic "
        "pipeline that transforms raw textual input into ranked similarity scores. Each stage of the "
        "pipeline performs a mathematically well defined operation, and the stages are connected "
        "sequentially so that the output of one stage becomes the input of the next. This section "
        "describes the methodology for each stage in detail."
    )

    add_heading_styled(doc, "6.1 Text Preprocessing Pipeline", level=2)

    add_body_text(doc,
        "Text preprocessing is the foundational stage that prepares raw text for numerical analysis. "
        "The preprocessing pipeline consists of three core operations applied in sequence."
    )

    add_body_paragraph(doc, "Tokenization", bold=True, font_size=11, space_after=4)

    add_body_text(doc,
        "Tokenization is the process of splitting raw text into individual word tokens. The tokenizer "
        "implemented in this project performs the following operations in order: conversion of all "
        "characters to lowercase for case normalization, removal of URLs and email addresses using "
        "regular expressions, removal of phone numbers in various formats, removal of special characters "
        "while retaining only letters, digits, and spaces, collapsing of multiple spaces into a single "
        "space, and filtering out single character tokens except for meaningful ones such as 'C' and 'R' "
        "which represent programming languages. The tokenizer also includes a phrase detection function "
        "that identifies common multi word technical phrases such as \"machine learning\" and \"deep "
        "learning\" to preserve their meaning as single units."
    )

    add_body_paragraph(doc, "Stop Word Removal", bold=True, font_size=11, space_after=4)

    add_body_text(doc,
        "Stop words are common words that appear frequently in text but carry little semantic meaning "
        "for the purpose of document comparison. Removing stop words allows the TF IDF algorithm to "
        "focus on words that are truly distinctive and informative. The stop word list used in this "
        "project is hand crafted and contains over 120 words organized into categories including "
        "articles, pronouns, prepositions, conjunctions, auxiliary verbs, determiners, and adverbs. "
        "In addition, the list includes domain specific resume stop words such as \"resume\", "
        "\"responsibilities\", \"qualifications\", \"experienced\", and \"professional\". These words appear "
        "in nearly every resume and job description and therefore do not help in distinguishing one "
        "document from another."
    )

    add_body_paragraph(doc, "Stemming", bold=True, font_size=11, space_after=4)

    add_body_text(doc,
        "Stemming reduces words to their root or base form so that morphological variants of the same "
        "word are treated as identical during comparison. The project implements a simplified version of "
        "the Porter Stemming Algorithm with approximately 25 suffix stripping rules organized across six "
        "steps. Step 1 handles plurals and past participles by removing suffixes such as \"sses\", \"ies\", "
        "and \"s\". Step 2 handles verb endings such as \"ed\" and \"ing\" with post processing rules for "
        "cases like \"hopping\" becoming \"hop\". Step 3 converts trailing \"y\" to \"i\" when preceded by "
        "a vowel. Steps 4 and 5 handle longer suffixes such as \"ational\" to \"ate\", \"ization\" to "
        "\"ize\", and \"fulness\" to \"ful\". Step 6 performs final cleanup by removing trailing \"e\" when "
        "the stem measure exceeds a threshold. The measure function counts the number of vowel consonant "
        "sequences in the stem and is used to ensure that suffixes are only removed when the remaining "
        "stem is long enough to be meaningful."
    )

    add_heading_styled(doc, "6.2 TF IDF Vectorization", level=2)

    add_body_text(doc,
        "After preprocessing, the cleaned tokens are converted into numerical vectors using the TF IDF "
        "weighting scheme. This conversion enables mathematical comparison between documents."
    )

    add_body_paragraph(doc, "Term Frequency (TF)", bold=True, font_size=11, space_after=4)

    add_body_text(doc,
        "Term Frequency measures how often a specific word appears in a single document, normalized by "
        "the total number of words in that document. The formula used is: TF(t, d) = count(t in d) / "
        "total_terms(d). Normalization by document length ensures that longer documents do not receive "
        "artificially higher scores simply because they contain more words. The implementation uses "
        "Python's collections.Counter to count word occurrences efficiently."
    )

    add_body_paragraph(doc, "Inverse Document Frequency (IDF)", bold=True, font_size=11, space_after=4)

    add_body_text(doc,
        "Inverse Document Frequency measures how rare or unique a word is across all documents in the "
        "corpus. The formula used is: IDF(t) = log(N / (1 + df(t))), where N is the total number of "
        "documents and df(t) is the number of documents containing term t. The addition of 1 in the "
        "denominator serves as a smoothing factor to prevent division by zero. The natural logarithm "
        "compresses the scale of the values. Words that appear in many documents receive low IDF scores "
        "because they are not useful for distinguishing between documents, while words that appear in "
        "few documents receive high IDF scores because they are highly discriminative."
    )

    add_body_paragraph(doc, "Combined TF IDF Score", bold=True, font_size=11, space_after=4)

    add_body_text(doc,
        "The final TF IDF score for each word in a document is obtained by multiplying its Term Frequency "
        "with its Inverse Document Frequency: TF IDF(t, d) = TF(t, d) x IDF(t). A word receives a high "
        "TF IDF score when it appears frequently in the current document but rarely across other documents. "
        "Conversely, a word that appears in most documents will have a TF IDF score close to zero regardless "
        "of how often it appears in any individual document. The TF IDF vector for a document is essentially "
        "a dictionary that maps each word to its TF IDF score, representing the document as a point in a "
        "high dimensional vector space."
    )

    add_heading_styled(doc, "6.3 Cosine Similarity and Ranking", level=2)

    add_body_text(doc,
        "Once both the job description and each resume have been converted into TF IDF vectors, the "
        "similarity between them is computed using cosine similarity. The formula is: cos(A, B) = "
        "(A dot B) / (|A| x |B|), where A dot B represents the dot product of the two vectors (the "
        "sum of the products of corresponding components), and |A| and |B| represent the magnitudes "
        "(Euclidean norms) of the vectors. The result is a value between 0 and 1, where 0 indicates "
        "no similarity and 1 indicates identical term distributions. The system also provides a detailed "
        "variant that tracks each term's individual contribution to the dot product, enabling the "
        "\"How It Works\" tab to display exactly which words drove the match score. After computing the "
        "cosine similarity for all resumes, the results are sorted in descending order to produce the "
        "final ranking."
    )

    add_heading_styled(doc, "6.4 Skill Extraction and Gap Analysis", level=2)

    add_body_text(doc,
        "In addition to the TF IDF based similarity score, the system performs skill extraction using a "
        "curated skills database containing over 200 skills organized into eight categories: programming "
        "languages, machine learning and AI, data science, web development, cloud and DevOps, databases, "
        "tools and platforms, and soft skills. The system scans both the job description and each resume "
        "for mentions of these skills using keyword matching. It then computes the skill gap by comparing "
        "the skills found in the job description against those found in each resume. The output includes "
        "three lists: matched skills (present in both), missing skills (required by the job but absent "
        "from the resume), and extra skills (present in the resume but not required by the job). A skill "
        "match rate is also calculated as the percentage of job description skills that the candidate "
        "possesses."
    )

    # ================================================================
    # PAGES 9-11: IMPLEMENTATION DETAILS
    # ================================================================
    add_page_break(doc)
    add_heading_styled(doc, "7. Implementation Details", level=1)

    add_body_text(doc,
        "This section provides a detailed description of the technical implementation of each module in "
        "the system. All core NLP and machine learning algorithms are implemented from scratch using only "
        "Python's built in libraries and the standard math module. External libraries are used exclusively "
        "for utility purposes such as file parsing, user interface rendering, and data visualization."
    )

    add_heading_styled(doc, "7.1 Technology Stack", level=2)

    tech_rows = [
        ["Programming Language", "Python 3.8+"],
        ["Web Framework", "Streamlit 1.28+"],
        ["PDF Parsing", "pdfplumber, PyPDF2 (fallback)"],
        ["DOCX Parsing", "python docx"],
        ["Data Visualization", "Plotly, Matplotlib"],
        ["Data Handling", "Pandas"],
        ["Notebook", "Jupyter"],
        ["Testing", "pytest"],
        ["Validation", "scikit learn (comparison only, not used at runtime)"],
    ]
    create_simple_table(doc, ["Layer", "Technology"], tech_rows, [2.2, 4.1])

    doc.add_paragraph()

    add_heading_styled(doc, "7.2 Preprocessing Module Implementation", level=2)

    add_body_text(doc,
        "The preprocessing module is located in the src/preprocessing/ directory and consists of three "
        "files: tokenizer.py, stopwords.py, and stemmer.py. Each file implements a single, well defined "
        "step in the preprocessing pipeline."
    )

    add_body_text(doc,
        "The tokenizer module (tokenizer.py) exposes two primary functions. The tokenize() function "
        "accepts raw text as input and returns a list of clean, lowercase word tokens. It uses Python's "
        "built in re module for all regular expression operations. The function first converts all "
        "characters to lowercase, then sequentially removes URLs, email addresses, and phone numbers "
        "using carefully crafted regex patterns. After removing special characters, it collapses multiple "
        "spaces, splits the text into tokens, and filters out single character tokens except for "
        "meaningful exceptions such as the programming languages C and R. The tokenize_preserve_phrases() "
        "function additionally detects 27 common multi word technical phrases such as \"machine learning\", "
        "\"natural language processing\", and \"cloud computing\"."
    )

    add_body_text(doc,
        "The stop words module (stopwords.py) defines a comprehensive set of over 120 stop words organized "
        "into clearly labeled categories: articles, pronouns, prepositions, conjunctions, auxiliary verbs, "
        "determiners, adverbs, and miscellaneous words. The module also defines a set of domain specific "
        "resume stop words including terms like \"resume\", \"qualifications\", \"responsibilities\", and "
        "\"proficient\" that appear in virtually every resume and job description. All categories are merged "
        "into a single master set using Python's set union operation. The remove_stopwords() function "
        "filters an input token list against this set using a list comprehension."
    )

    add_body_text(doc,
        "The stemmer module (stemmer.py) implements a simplified Porter Stemmer with 25 suffix stripping "
        "rules distributed across six processing steps. The module includes helper functions for computing "
        "the word measure (the number of vowel consonant sequences), checking for vowel presence, detecting "
        "double consonant endings, and identifying the consonant vowel consonant pattern at the end of a "
        "word. The stem() function processes each word through all six steps in order, applying only the "
        "first matching rule within each step. The stem_tokens() function applies the stemmer to an entire "
        "list of tokens and returns the list of stemmed forms."
    )

    add_heading_styled(doc, "7.3 Vectorization Module Implementation", level=2)

    add_body_text(doc,
        "The vectorization module is located in the src/vectorization/ directory and consists of three "
        "files: tf.py, idf.py, and tfidf.py. These modules implement the TF IDF vectorization algorithm "
        "from scratch."
    )

    add_body_text(doc,
        "The Term Frequency module (tf.py) provides the compute_tf() function, which accepts a list of "
        "preprocessed tokens and returns a dictionary mapping each unique word to its normalized frequency. "
        "The implementation uses collections.Counter to count word occurrences and divides each count by "
        "the total number of tokens in the document. A compute_raw_tf() function is also provided for "
        "debugging and educational purposes, returning raw (unnormalized) word counts."
    )

    add_body_text(doc,
        "The Inverse Document Frequency module (idf.py) provides the compute_idf() function, which "
        "accepts a list of documents (each represented as a list of tokens) and returns a dictionary "
        "mapping each word to its IDF score. The implementation first computes the document frequency "
        "for each word by iterating over documents and using set deduplication to count each word only "
        "once per document. It then applies the formula IDF(t) = log(N / (1 + df(t))) using Python's "
        "math.log function. A compute_document_frequency() function is also provided for inspection."
    )

    add_body_text(doc,
        "The TF IDF module (tfidf.py) combines the TF and IDF modules to produce final TF IDF vectors. "
        "The compute_tfidf() function computes the TF IDF vector for a single document by multiplying "
        "each word's TF score with its corresponding IDF score. The build_tfidf_matrix() function "
        "orchestrates the entire process for multiple documents by first computing IDF across the full "
        "corpus and then computing TF IDF vectors for each individual document. The get_top_terms() "
        "function returns the top N most important terms from a TF IDF vector, sorted by score in "
        "descending order, which is used in the dashboard to display the most distinctive words for "
        "each resume."
    )

    add_heading_styled(doc, "7.4 Similarity and Ranking Module Implementation", level=2)

    add_body_text(doc,
        "The similarity module is located in the src/similarity/ directory and consists of two files: "
        "cosine.py and ranker.py."
    )

    add_body_text(doc,
        "The cosine similarity module (cosine.py) implements two functions. The cosine_similarity() "
        "function computes the standard cosine similarity between two TF IDF vectors represented as "
        "dictionaries. It finds the union of all words across both vectors, computes the dot product "
        "by summing the products of corresponding values (treating missing words as zero), computes "
        "the magnitude of each vector as the square root of the sum of squared values, and returns "
        "the dot product divided by the product of magnitudes. The cosine_similarity_detailed() function "
        "performs the same computation but additionally tracks and returns the dot product value, both "
        "magnitudes, the list of common terms, and a dictionary of per term contributions to the dot "
        "product. This detailed information is displayed in the \"How It Works\" tab of the dashboard."
    )

    add_body_text(doc,
        "The ranker module (ranker.py) serves as the central orchestrator that connects all other modules "
        "into a complete matching pipeline. The rank_resumes() function is the main entry point. It "
        "accepts a raw job description string and a dictionary of resume names mapped to their raw text. "
        "The function preprocesses all documents using the tokenize, remove stopwords, and stem tokens "
        "pipeline. It then builds the corpus by combining the job description tokens with all resume "
        "tokens, computes IDF across the entire corpus, and generates TF IDF vectors for each document. "
        "Cosine similarity is computed between the job description vector and each resume vector. "
        "Additionally, the function extracts skills from both the job description and each resume using "
        "a curated JSON database of over 200 skills across eight categories, computes the skill gap "
        "(matched, missing, and extra skills), and retrieves the top 10 TF IDF terms for each resume. "
        "The results are sorted by cosine similarity score in descending order and returned as a list "
        "of dictionaries containing all computed metrics."
    )

    add_heading_styled(doc, "7.5 File Parser Module Implementation", level=2)

    add_body_text(doc,
        "The file parser module is located in the src/parsers/ directory and provides utilities for "
        "extracting raw text from resume files in different formats. These parsers are classified as "
        "utility components because they do not involve any NLP or machine learning logic."
    )

    add_body_text(doc,
        "The PDF parser (pdf_parser.py) implements a dual strategy approach. It first attempts to extract "
        "text using pdfplumber, which provides superior accuracy for most PDF files. If pdfplumber fails "
        "or produces empty output, the parser falls back to PyPDF2 as a secondary extraction engine. "
        "Both strategies support file path inputs as well as file like objects from Streamlit's file "
        "uploader. The parser iterates over all pages of the PDF, extracts text from each page, and "
        "joins the results with newline characters."
    )

    add_body_text(doc,
        "The DOCX parser (docx_parser.py) uses the python docx library to extract text from Microsoft "
        "Word documents. It processes both paragraph elements and table elements within the document, "
        "joining table cell contents with pipe separators. This ensures that tabular data commonly found "
        "in resumes, such as skill matrices or work experience tables, is properly captured."
    )

    add_heading_styled(doc, "7.6 User Interface Implementation", level=2)

    add_body_text(doc,
        "The web based dashboard is implemented in app.py using the Streamlit framework. The interface "
        "features a premium dark theme with custom CSS styling including gradient backgrounds, animated "
        "score cards, and color coded skill tags. The dashboard is organized into a sidebar for system "
        "information and a main content area with four tabs."
    )

    add_body_text(doc,
        "The Candidate Rankings tab displays a leaderboard of resumes sorted by cosine similarity, with "
        "expandable detail cards showing matched, missing, and extra skills along with top TF IDF terms. "
        "The Skills Analysis tab provides a comprehensive skill gap analysis with per candidate breakdowns. "
        "The How It Works tab offers full mathematical transparency by displaying the step by step "
        "computation for any selected resume, including raw tokens, filtered tokens, stemming examples, "
        "TF IDF vectors, dot product breakdowns, and per term contributions. The Visualizations tab "
        "presents interactive Plotly charts including a match score bar chart, a skill gap grouped bar "
        "chart, and a cosine similarity versus skill match rate scatter plot."
    )

    add_heading_styled(doc, "7.7 Testing and Validation", level=2)

    add_body_text(doc,
        "The project includes a comprehensive test suite located in the tests/ directory. The tests are "
        "organized into four test files covering different aspects of the system."
    )

    test_rows = [
        ["test_preprocessing.py", "Validates tokenization, stop word removal, and stemming correctness"],
        ["test_tfidf.py", "Validates TF, IDF, and TF IDF computations with known expected values"],
        ["test_cosine.py", "Validates cosine similarity with manual calculations and edge cases"],
        ["test_validate_sklearn.py", "Compares from scratch output against scikit learn to prove correctness"],
    ]
    create_simple_table(doc, ["Test File", "Purpose"], test_rows, [2.2, 4.1])

    doc.add_paragraph()

    add_body_text(doc,
        "The most significant test file is test_validate_sklearn.py, which runs the same input data "
        "through both the from scratch implementation and scikit learn's TfidfVectorizer and "
        "cosine_similarity functions. The test asserts that the cosine similarity values match within "
        "floating point precision (difference less than 0.000001) and that the ranking order of candidates "
        "is identical. This validation proves that the from scratch implementation is mathematically "
        "correct and produces results equivalent to established, production grade libraries."
    )

    # ================================================================
    # PAGES 12-13: RESULTS AND DISCUSSION
    # ================================================================
    add_page_break(doc)
    add_heading_styled(doc, "8. Results and Discussion", level=1)

    add_body_text(doc,
        "This section presents the results obtained from testing the system with sample data and discusses "
        "the observations, strengths, and limitations of the implemented approach."
    )

    add_heading_styled(doc, "8.1 Test Data and Experimental Setup", level=2)

    add_body_text(doc,
        "The system was tested using five sample resumes representing different professional profiles: "
        "a Data Scientist, a Web Developer, a DevOps Engineer, a Marketing Analyst, and a Fresh Graduate. "
        "Three sample job descriptions were used for matching: a Data Scientist position, a Full Stack "
        "Developer position, and a Cloud Engineer position. Each resume and job description was stored "
        "as a plain text file in the project's data directory."
    )

    add_heading_styled(doc, "8.2 Matching Results", level=2)

    add_body_text(doc,
        "When matching against the Data Scientist job description, the system correctly ranked the Data "
        "Scientist resume as the top candidate with the highest cosine similarity score. The Web Developer "
        "and DevOps Engineer resumes ranked lower due to limited overlap with the required machine learning "
        "and data science terminology. The Marketing Analyst resume showed moderate similarity owing to "
        "shared data analysis skills, while the Fresh Graduate resume ranked based on foundational skills "
        "mentioned in the resume. These results demonstrate that the TF IDF and cosine similarity approach "
        "effectively captures the relevance of resumes to job descriptions based on term overlap and "
        "term importance."
    )

    add_heading_styled(doc, "8.3 Skill Gap Analysis Results", level=2)

    add_body_text(doc,
        "The skill extraction module successfully identified technical skills from both job descriptions "
        "and resumes by matching against the curated database of over 200 skills. For the Data Scientist "
        "job description, the system identified skills such as Python, TensorFlow, machine learning, "
        "deep learning, pandas, and numpy. The skill gap analysis accurately highlighted which required "
        "skills were present in each resume and which were missing, providing actionable insights for "
        "recruiters. The skill match rate metric offered a complementary perspective to the cosine "
        "similarity score, particularly useful for cases where a candidate has the right skills but uses "
        "different terminology."
    )

    add_heading_styled(doc, "8.4 Validation Against scikit learn", level=2)

    add_body_text(doc,
        "The validation tests confirmed that the from scratch cosine similarity computation matches "
        "scikit learn's output within floating point precision (absolute difference less than 0.000001). "
        "Furthermore, the ranking order produced by the from scratch implementation was identical to "
        "scikit learn's ranking for the same set of documents. This validation provides strong evidence "
        "that the manual implementation is mathematically correct and reliable."
    )

    add_heading_styled(doc, "8.5 Dashboard Performance", level=2)

    add_body_text(doc,
        "The Streamlit dashboard performed efficiently during testing, processing five resumes against a "
        "job description in under two seconds on a standard laptop. The interactive visualizations "
        "rendered smoothly, and the mathematical transparency features in the \"How It Works\" tab "
        "correctly displayed all intermediate computation steps including token counts, stemming "
        "examples, TF IDF vectors, dot product values, vector magnitudes, and per term contributions."
    )

    add_heading_styled(doc, "8.6 Strengths of the Approach", level=2)

    strengths = [
        "Complete transparency: Every step of the computation is visible and verifiable by the user.",
        "No dependency on black box libraries for core NLP and ML algorithms.",
        "Modular architecture allows individual components to be tested and replaced independently.",
        "The skill gap analysis provides actionable insights beyond a single similarity score.",
        "The system supports multiple file formats (PDF, DOCX, TXT) for practical usability.",
        "Validation against scikit learn proves the correctness of the implementation.",
    ]
    for s in strengths:
        add_bullet(doc, s)

    add_heading_styled(doc, "8.7 Limitations", level=2)

    limitations = [
        "The TF IDF approach treats words independently and does not capture semantic relationships between terms. For example, \"Python\" and \"programming\" are treated as unrelated words.",
        "The Porter Stemmer can produce non intuitive stems. For instance, \"organization\" may be stemmed to \"organ\", which could affect matching accuracy.",
        "The skill extraction relies on keyword matching against a predefined database and may miss skills not listed in the database or skills expressed using unconventional terminology.",
        "The system currently processes only text based PDF files and does not support scanned or image based documents that would require Optical Character Recognition.",
    ]
    for l in limitations:
        add_bullet(doc, l)

    # ================================================================
    # PAGE 13-14: CONCLUSION
    # ================================================================
    add_page_break(doc)
    add_heading_styled(doc, "9. Conclusion", level=1)

    add_body_text(doc,
        "This project successfully demonstrates the design and implementation of an NLP Based Resume "
        "Parser and Matcher built entirely from scratch. The system applies fundamental Natural Language "
        "Processing techniques including tokenization, stop word removal, and Porter stemming to "
        "preprocess textual data from resumes and job descriptions. The preprocessed text is converted "
        "into numerical representations using the Term Frequency Inverse Document Frequency weighting "
        "scheme, and the similarity between documents is measured using cosine similarity. All of these "
        "algorithms have been implemented manually using only Python's standard library and the math "
        "module, without relying on any external NLP or machine learning frameworks."
    )

    add_body_text(doc,
        "The project goes beyond basic functionality by incorporating a comprehensive skill extraction "
        "and gap analysis module that compares candidate skills against job requirements using a curated "
        "database of over 200 skills across eight professional categories. This feature provides "
        "recruiters with detailed, actionable information about each candidate's strengths and gaps."
    )

    add_body_text(doc,
        "The interactive Streamlit dashboard serves as the user facing interface, offering a professional "
        "and visually appealing experience with features such as ranked candidate leaderboards, skill "
        "gap visualizations, mathematical transparency tools, and CSV export capabilities. The dashboard "
        "supports resume uploads in PDF, DOCX, and TXT formats, making it practical for real world use."
    )

    add_body_text(doc,
        "The correctness of the from scratch implementation has been rigorously validated through unit "
        "tests and comparison against scikit learn's production grade implementations. The validation "
        "tests confirm that the cosine similarity values and ranking orders produced by the from scratch "
        "code match those of scikit learn within floating point precision. This validation demonstrates "
        "that it is entirely feasible to implement these algorithms manually while maintaining the same "
        "level of accuracy as established libraries."
    )

    add_body_text(doc,
        "Through this project, the team has gained a deep understanding of the mathematical foundations "
        "of text analysis, information retrieval, and document similarity. The from scratch approach "
        "provided invaluable learning that would not have been possible by simply calling library "
        "functions. The project stands as both an educational exercise in algorithm implementation and "
        "a functional prototype for automated resume screening."
    )

    # ================================================================
    # PAGE 14: FUTURE SCOPE
    # ================================================================
    add_page_break(doc)
    add_heading_styled(doc, "10. Future Scope", level=1)

    add_body_text(doc,
        "While the current system effectively demonstrates the application of TF IDF and cosine similarity "
        "for resume matching, there are several avenues for future improvement and extension that would "
        "enhance the system's capabilities and practical applicability."
    )

    add_heading_styled(doc, "10.1 Semantic Understanding with Word Embeddings", level=2)

    add_body_text(doc,
        "The current TF IDF approach treats each word independently and does not capture semantic "
        "relationships between terms. Integrating word embedding models such as Word2Vec, GloVe, or "
        "contextual embeddings from transformer models like BERT would allow the system to understand "
        "that terms such as \"Python\" and \"programming\" or \"TensorFlow\" and \"deep learning\" are "
        "semantically related, even if they do not share identical surface forms. This would significantly "
        "improve matching accuracy for resumes that use different terminology to describe the same skills."
    )

    add_heading_styled(doc, "10.2 Named Entity Recognition for Structured Parsing", level=2)

    add_body_text(doc,
        "Implementing Named Entity Recognition (NER) would enable the system to automatically extract "
        "structured information from resumes, such as candidate names, educational institutions, degree "
        "types, company names, job titles, and dates of employment. This structured data could be used "
        "for more refined matching criteria, such as filtering candidates by years of experience or "
        "educational qualifications."
    )

    add_heading_styled(doc, "10.3 Support for Scanned and Image Based Resumes", level=2)

    add_body_text(doc,
        "Currently, the system can only process text based PDF documents. Many resumes are submitted as "
        "scanned images or PDFs with embedded images. Integrating Optical Character Recognition (OCR) "
        "technology, such as Tesseract, would allow the system to extract text from image based documents "
        "and significantly expand the range of resumes it can process."
    )

    add_heading_styled(doc, "10.4 Multi Language Support", level=2)

    add_body_text(doc,
        "The current implementation is designed for English language documents. Extending the system to "
        "support multiple languages would make it applicable in international recruitment contexts. This "
        "would require language specific tokenizers, stop word lists, and stemming algorithms, as well "
        "as potentially multilingual embedding models."
    )

    add_heading_styled(doc, "10.5 Machine Learning Based Ranking", level=2)

    add_body_text(doc,
        "A supervised machine learning model could be trained on historical hiring data to learn the "
        "relationship between resume features and hiring outcomes. This would allow the system to go "
        "beyond simple text similarity and incorporate factors such as career progression, industry "
        "relevance, and cultural fit. Techniques such as gradient boosted trees or neural ranking models "
        "could be explored for this purpose."
    )

    add_heading_styled(doc, "10.6 Real Time Job Portal Integration", level=2)

    add_body_text(doc,
        "The system could be integrated with existing job portals and human resource management systems "
        "through APIs, allowing it to automatically screen incoming applications in real time. This would "
        "transform the prototype into a production ready tool that fits seamlessly into existing "
        "recruitment workflows."
    )

    add_heading_styled(doc, "10.7 Bias Detection and Fairness", level=2)

    add_body_text(doc,
        "Future work could incorporate bias detection mechanisms to ensure that the system does not "
        "inadvertently discriminate against candidates based on gender, ethnicity, age, or other "
        "protected characteristics. This could involve analyzing the sensitivity of the matching "
        "algorithm to demographic indicators and implementing fairness constraints in the ranking process."
    )

    # ================================================================
    # PAGE 15: REFERENCES
    # ================================================================
    add_page_break(doc)
    add_heading_styled(doc, "11. References", level=1)

    references = [
        "Porter, M. F. (1980). \"An Algorithm for Suffix Stripping.\" Program, Vol. 14, No. 3, pp. 130\u2013137.",
        "Sparck Jones, K. (1972). \"A Statistical Interpretation of Term Specificity and Its Application in Retrieval.\" Journal of Documentation, Vol. 28, No. 1, pp. 11\u201321.",
        "Salton, G., and Buckley, C. (1988). \"Term Weighting Approaches in Automatic Text Retrieval.\" Information Processing and Management, Vol. 24, No. 5, pp. 513\u2013523.",
        "Manning, C. D., Raghavan, P., and Schutze, H. (2008). Introduction to Information Retrieval. Cambridge University Press.",
        "Baeza Yates, R., and Ribeiro Neto, B. (2011). Modern Information Retrieval: The Concepts and Technology behind Search. Addison Wesley.",
        "Jurafsky, D., and Martin, J. H. (2023). Speech and Language Processing. 3rd Edition (Draft). Pearson.",
        "Bird, S., Klein, E., and Loper, E. (2009). Natural Language Processing with Python. O'Reilly Media.",
        "Pedregosa, F. et al. (2011). \"Scikit learn: Machine Learning in Python.\" Journal of Machine Learning Research, Vol. 12, pp. 2825\u20132830.",
        "Streamlit Documentation. Available at: https://docs.streamlit.io/. Accessed: June 2025.",
        "Python Software Foundation. \"Python 3 Documentation.\" Available at: https://docs.python.org/3/. Accessed: June 2025.",
        "pdfplumber Documentation. Available at: https://github.com/jsvine/pdfplumber. Accessed: June 2025.",
        "python docx Documentation. Available at: https://python-docx.readthedocs.io/. Accessed: June 2025.",
    ]

    for i, ref in enumerate(references, 1):
        p = doc.add_paragraph()
        run = p.add_run(f"[{i}]  ")
        run.bold = True
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
        run2 = p.add_run(ref)
        run2.font.size = Pt(11)
        run2.font.name = 'Times New Roman'
        set_paragraph_spacing(p, before=2, after=6, line_spacing=1.15)
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.first_line_indent = Inches(-0.4)

    # ================================================================
    # SAVE THE DOCUMENT
    # ================================================================
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ML_Report.docx")
    doc.save(output_path)
    print(f"\n{'='*60}")
    print(f"  Report generated successfully!")
    print(f"  Output: {output_path}")
    print(f"{'='*60}")


if __name__ == "__main__":
    generate_report()
