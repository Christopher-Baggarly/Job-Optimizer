# Job-Optimizer
AI Resume Automation and Regional Job Search
import os
import sys
import json
import re
import zipfile
import xml.etree.ElementTree as ET
import jinja2
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxtpl import DocxTemplate
from openai import OpenAI

# =====================================================================
# ⚙️ CORES, UTILITIES & DATE FORMATTERS
# =====================================================================
def clean_folder_date(date_str):
    """Normalizes raw system strings (e.g., 'Mar-27', '3/1/2027') into 'Month Year'."""
    if not date_str:
        return "March 2027"
    
    date_clean = date_str.strip().replace("-", " ").replace("/", " ")
    
    months_map = {
        "1": "January", "2": "February", "3": "March", "4": "April", "5": "May", "6": "June",
        "7": "July", "8": "August", "9": "September", "10": "October", "11": "November", "12": "December",
        "01": "January", "02": "February", "03": "March", "04": "April", "05": "May", "06": "June",
        "07": "July", "08": "August", "09": "September", "10": "October", "11": "November", "12": "December"
    }
    short_months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
    full_months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    
    lower_str = date_clean.lower()
    matched_month = None
    for i, m in enumerate(short_months):
        if m in lower_str:
            matched_month = full_months[i]
            break
            
    years = re.findall(r'\b(\d{2}|\d{4})\b', date_clean)
    matched_year = "2027"
    if years:
        for y in years:
            if len(y) == 4:
                matched_year = y
            elif len(y) == 2:
                matched_year = f"20{y}"
                
    if not matched_month:
        digits = re.findall(r'\b\d{1,2}\b', date_clean)
        if digits and digits[0] in months_map:
            matched_month = months_map[digits[0]]
            
    if not matched_month:
        matched_month = "March"
        
    return f"{matched_month} {matched_year}"

def scrub(text):
    if not isinstance(text, str): 
        return ""
    filler = ["dates not provided", "undisclosed", "not provided", "n/a", "unknown", "none"]
    return "" if text.lower().strip() in filler else text.strip()

def clean_bullet(text):
    if not isinstance(text, str): 
        return ""
    cleaned = re.sub(r'\s+', ' ', text)
    return cleaned.strip()

def simplify(text):
    return re.sub(r'[^a-zA-Z0-9]', '', text).lower()

# Initialize OpenAI Client Safely via Environment Variable
api_key = ("OPENAI_API_KEY")
if not api_key:
    print("⚠️ Warning: OPENAI_API_KEY environment variable not found. Check local system context.")
client = OpenAI(api_key=api_key)

# =====================================================================
# 🎮 LIVE VBA INTERFACE ROUTER
# =====================================================================
if len(sys.argv) >= 8:
    TEMPLATE_NAME = sys.argv[1]
    STUDENT_NAME = sys.argv[2]
    TARGET_TITLE = sys.argv[3]
    PROGRAM_FOLDER = sys.argv[4]
    RAW_DATE_INPUT = sys.argv[5]  
    RAW_RESUME_FILENAME = sys.argv[6]
    NEED_OPTION = sys.argv[7]
    HAS_EXPERIENCE = sys.argv[8].lower() == "true"
else:
    print("⚠️ No VBA arguments detected. Falling back to manual test parameters...")
    TEMPLATE_NAME = "Template_Modern_Side"
    STUDENT_NAME = "Student Name"
    TARGET_TITLE = "Medical Assistant"
    PROGRAM_FOLDER = "MSMA"
    RAW_DATE_INPUT = "Mar-27"  
    RAW_RESUME_FILENAME = "Student Name Resume 1.docx"
    NEED_OPTION = "Improve"
    HAS_EXPERIENCE = True

CLASS_DATE_FOLDER = clean_folder_date(RAW_DATE_INPUT)

BASE_DIR = r"C:\Resume_System"
TEMPLATES_FOLDER = os.path.join(BASE_DIR, "Templates")
RECORDS_DIR = os.path.join(BASE_DIR, "Student_Records")

if os.path.exists(BASE_DIR):
    os.chdir(BASE_DIR)

tokens = STUDENT_NAME.strip().split()
last_name = tokens[-1] if len(tokens) > 1 else "Unknown"
first_name = " ".join(tokens[:-1]) if len(tokens) > 1 else tokens[0]

output_dir = os.path.join(RECORDS_DIR, PROGRAM_FOLDER, CLASS_DATE_FOLDER, f"{last_name}, {first_name}")
os.makedirs(output_dir, exist_ok=True)

class_year = 2026
year_match = re.search(r'\b(20\d{2})\b', CLASS_DATE_FOLDER)
if year_match:
    class_year = int(year_match.group(1))

# =====================================================================
# STAGE 0: TEMPLATE ANALYZER
# =====================================================================
template_path = os.path.join(TEMPLATES_FOLDER, f"{TEMPLATE_NAME}.docx")
if not os.path.exists(template_path):
    print(f"❌ Error: Cannot find target template at {template_path}")
    sys.exit(1)

namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
detected_tags = set()

try:
    with zipfile.ZipFile(template_path) as docx_zip:
        xml_content = docx_zip.read('word/document.xml')
        root = ET.fromstring(xml_content)
        text_nodes = [node.text for node in root.findall('.//w:t', namespaces) if node.text]
        combined_xml_stream = "".join(text_nodes)
        tags = re.findall(r'\{\{[\s\S]*?\}\}|\{%[\s\S]*?%\}', combined_xml_stream)
        for tag in tags:
            detected_tags.add(tag.strip())
except Exception as e:
    print(f"❌ Stage 0 Analysis Halted: {str(e)}")
    sys.exit(1)

requires_horizontal_skills = any('horizontal_skills' in t for t in detected_tags)
requires_horizontal_certs = any('horizontal_certs' in t for t in detected_tags)
requires_vertical_skills = any('student_skills' in t for t in detected_tags)
requires_vertical_certs = any('student_certs' in t for t in detected_tags)

# =====================================================================
# STAGE 1: UNCONDITIONAL MASTER INGESTION & DEEP EARLY-OPTIMIZATION
# =====================================================================
try:
    if not RAW_RESUME_FILENAME.lower().endswith('.docx'):
        RAW_RESUME_FILENAME += ".docx"

    SEARCH_PATHS = [
        os.path.join(BASE_DIR, RAW_RESUME_FILENAME),
        os.path.join(RECORDS_DIR, RAW_RESUME_FILENAME),
        os.path.join(output_dir, RAW_RESUME_FILENAME),
        os.path.join(RECORDS_DIR, PROGRAM_FOLDER, RAW_DATE_INPUT, RAW_RESUME_FILENAME)
    ]

    raw_source_path = next((p for p in SEARCH_PATHS if os.path.exists(p)), None)
    if not raw_source_path:
        print(f"❌ Error: Source '{RAW_RESUME_FILENAME}' missing from system paths.")
        sys.exit(1)

    doc = Document(raw_source_path)
    all_text = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell and cell.text.strip():
                    all_text.append(cell.text.strip())

    full_content = "\n".join(all_text)

    # Dynamic Formatting & Prompt Routing based entirely on Experience boolean
    if HAS_EXPERIENCE:
        strategy_instruction = f"""
        - FORMATTING MODE: EXPANDED DETAILED PORTFOLIO.
        - Write an elite, high-impact 3-sentence EXECUTIVE/PROFESSIONAL SUMMARY.
        - Lead immediately with the dynamic target title: "{TARGET_TITLE}".
        - Focus heavily on cross-functional metrics, multi-site operational scope, and deep-dive achievements.
        - BULLET CONFIGURATION: Generate an uncapped, comprehensive array of elite STAR-method achievements per role. Extract ALL metric-driven details, budgets, and leadership milestones. Do not artificially truncate or limit the data.
        """
        system_bullet_example = '"Elite executive achievement milestone detailing multi-site operational impact, metrics, and leadership scope."'
    else:
        strategy_instruction = f"""
        - FORMATTING MODE: GENERAL STRICT 3-BULLET RULE.
        - Write a highly tailored, deeply personalized 3-sentence CAREER OBJECTIVE for the role of "{TARGET_TITLE}".
        - BAN ALL BANAL BUZZWORDS: Absolutely do not use generic filler like 'results-driven professional' or 'seeking an opportunity'.
        - RE-INJECT HUMAN PERSONALITY: Weave in concrete context clues from their actual background text.
        - Sentence 1: State immediate intent to transition into a "{TARGET_TITLE}" role, explicitly framed by an operational hallmark of their background.
        - Sentence 2: Connect a real, concrete operational strength from their past or studies directly to the immediate problems they will solve in the target position.
        - Sentence 3: Close with a punchy, personalized statement of intent that focuses on professional execution.
        - BULLET CONFIGURATION: Generate EXACTLY 3 elite STAR-method bullets per role. Keep descriptions uniform, concise, and punchy. Even if the original role was unrelated, rewrite the duties to emphasize transferability to "{TARGET_TITLE}".
        """
        system_bullet_example = '"Elite statement 1", "Elite statement 2", "Elite statement 3"'

    prompt = f"""
    You are an elite Senior Resume Architect.
    Your mission is to completely extract and early-optimize ALL raw historical data for the Target Title: "{TARGET_TITLE}".
    
    STRATEGY AND CONFIGURATION RULES:
    {strategy_instruction}
    
    CRITICAL ARCHITECTURAL RULES (UNCONDITIONAL HARVESTING & DEEP POLISHING):
    1. EXHAUSTIVE EXTRACTION FIRST: Look at every single line in the raw text. You MUST extract EVERY SINGLE historical job entry found. Do not omit any role because it seems old, short, or unaligned.
    
    2. SIMULTANEOUS PRODUCTION-READY POLISHING: You are NOT returning raw extractions. Every single object populated into the `experience_pool` array must be immediately clean, professional, and optimized as follows:
       - role & company: Clean up formatting, capitalization, and naming standards (e.g., split clean metrics or titles if joined by punctuation).
       - ai_relevance_score: 1-5 Scale evaluated strictly against the target intent: "{TARGET_TITLE}".
         * Score 5: Direct match (roles sharing the exact title, core duties, or industry setting as "{TARGET_TITLE}").
         * Score 4: Strong Transferable (roles in different fields that heavily feature core soft skills, operational oversight, client/customer/patient interactions, compliance, or administrative workflows critical to being a successful "{TARGET_TITLE}").
         * Score 3: Moderate Transferable (roles with general professional skills, team coordination, or basic operational tasks).
         * Score 1-2: Low/Unaligned (completely unrelated task profiles or passive helper roles).
       - resume_bullets: Adhere strictly to the rules of the selected FORMATTING MODE outlined above.
       - appendix_paragraph: Write EXACTLY 1-2 sentences summarizing the core functional footprint and operational reliability of this role for a master appendix. This must be written in a sophisticated narrative style (no bullet formats).
       
    3. STRICT DATA SEGREGATION:
       - "skills": Core technical competencies or professional practices. Generate these if missing from raw data, ensuring they map dynamically to the Target Title: "{TARGET_TITLE}".
       - "certifications": ONLY formal credentials/licensures.
       - DEDUPLICATION: No skill can appear in the certifications array.
       
    RAW RESUME OBJECTS TO HARVEST AND REWRITE:
    {full_content}
    """

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {
                "role": "system",
                "content": (
                    "Return a fully populated JSON dictionary using this exact schema structure:\n"
                    "{\n"
                    "  \"summary\": \"Confident 3-sentence profile.\",\n"
                    "  \"experience_pool\": [\n"
                    "       {\n"
                    "         \"company\": \"string\",\n"
                    "         \"role\": \"string\",\n"
                    "         \"dates\": \"string\",\n"
                    "         \"ai_relevance_score\": 5,\n"
                    f"         \"resume_bullets\": [{system_bullet_example}],\n"
                    "         \"appendix_paragraph\": \"Punchy polished narrative description sentence 1.\"\n"
                    "       }\n"
                    "  ],\n"
                    "  \"education_pool\": [{\"institution\": \"str\", \"degree\": \"str\", \"dates\": \"str\"}],\n"
                    "  \"skills\": [\"str\"],\n"
                    "  \"certifications\": [\"str\"],\n"
                    "  \"references\": [\"str\"]\n"
                    "}\n"
                    "Do not enclose your output in markdown syntax code-blocks."
                )
            },
            {"role": "user", "content": prompt}
        ],
        temperature=0.1
    )

    raw_response = response.choices[0].message.content.strip()
    
    backtick_3 = "\x60\x60\x60"
    backtick_json = "\x60\x60\x60json"
    
    if raw_response.startswith(backtick_3):
        if raw_response.startswith(backtick_json):
            raw_response = raw_response.removeprefix(backtick_json)
        else:
            raw_response = raw_response.removeprefix(backtick_3)
        raw_response = raw_response.removesuffix(backtick_3)
        
    raw_response = raw_response.strip()
    ai_data = json.loads(raw_response)

    contact_email = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', full_content)
    contact_phone = re.search(r'(\(\d{3}\)\s*\d{3}-\d{4}|\d{3}-\d{3}-\d{4})', full_content)
    ai_data["email"] = contact_email.group(0) if contact_email else ""
    ai_data["phone"] = contact_phone.group(0) if contact_phone else ""
    ai_data["address"] = ""

except Exception as e:
    print(f"❌ Stage 1 Ingestion Failure: {str(e)}")
    sys.exit(1)

# =====================================================================
# STAGE 2: PYTHONIC RELEVANCE STRATIFICATION & SORTING MATRIX
# =====================================================================
try:
    def parse_end_year(date_str, default_val):
        if not date_str:
            return 0
        clean_d = str(date_str).lower()
        if "present" in clean_d or "current" in clean_d:
            return default_val
        found_years = [int(y) for y in re.findall(r'\b(20\d{2}|19\d{2})\b', str(date_str))]
        return max(found_years) if found_years else 0

    master_experience_pool = ai_data.get("experience_pool", [])
    for job in master_experience_pool:
        job["end_year"] = parse_end_year(job.get("dates", ""), class_year)
        if "ai_relevance_score" not in job:
            job["ai_relevance_score"] = 1

    appendix_jobs_sorted = sorted(master_experience_pool, key=lambda x: x["end_year"], reverse=True)
    
    sorted_by_relevance_matrix = sorted(
        master_experience_pool,
        key=lambda x: (x.get("ai_relevance_score", 1), x.get("end_year", 0)),
        reverse=True
    )
    
    # Conditional Layout Routing Matrix
    if HAS_EXPERIENCE:
        # EXPANDED MODE: Pass all roles directly to primary layout; bypass appendix
        primary_jobs = sorted_by_relevance_matrix
        create_appendix = False
        experience_note = ""
    else:
        # GENERAL MODE: Restrict to top 3 roles; route overflow to appendix
        primary_jobs = sorted_by_relevance_matrix[:3]
        create_appendix = len(appendix_jobs_sorted) > 3
        experience_note = "Gaps in work history are due to education or prioritizing experience relevant to the position. Full work history attached for review." if create_appendix else ""

    primary_jobs = sorted(primary_jobs, key=lambda x: x.get("end_year", 0), reverse=True)

    raw_refs = ai_data.get("references", [])
    final_references = [{"name": "References available upon request.", "phone": ""}] if not raw_refs else [{"name": str(r), "phone": ""} for r in raw_refs]

    raw_certs = [c.strip() for c in ai_data.get("certifications", []) if c.strip()]
    raw_skills = [s.strip() for s in ai_data.get("skills", []) if s.strip()]
    
    normalized_certs = {simplify(c) for c in raw_certs}
    filtered_skills = [skill for skill in raw_skills if simplify(skill) not in normalized_certs]

    def build_primary_experience_payload(job_list):
        payload = []
        for j in job_list:
            payload.append({
                "title": j.get("role", ""),
                "company": j.get("company", ""),
                "date": scrub(j.get("dates", "")),
                "bullets": [clean_bullet(b) for b in j.get("resume_bullets", [])]
            })
        return payload

    section_title = "Professional Summary" if HAS_EXPERIENCE else "Career Objective"

    final_payload = {
        "title": TARGET_TITLE,
        "summary_title": section_title,
        "summary": scrub(ai_data.get("summary", "")),
        "email": ai_data.get("email", ""),
        "phone": ai_data.get("phone", ""),
        "address": ai_data.get("address", ""),
        "first": {"name": first_name}, "last": {"name": last_name},
        "First": {"name": first_name}, "Last": {"name": last_name},
        "experience_note": experience_note,
        "student": {
            "education": [{"institution": e.get("institution", ""), "degree": e.get("degree", ""), "date": scrub(e.get("dates", ""))} for e in ai_data.get("education_pool", [])],
            "experience": build_primary_experience_payload(primary_jobs),
            "references": final_references
        }
    }

    if requires_horizontal_skills:
        final_payload["horizontal_skills"] = "  •  ".join(filtered_skills)
    if requires_vertical_skills:
        final_payload["student_skills"] = filtered_skills

    if requires_horizontal_certs:
        final_payload["horizontal_certs"] = "  |  ".join(raw_certs)
    if requires_vertical_certs:
        final_payload["student_certs"] = raw_certs

except Exception as e:
    print(f"❌ Stage 2 Logic Failure: {str(e)}")
    sys.exit(1)

# =====================================================================
# STAGE 3: RECURSIVE STRUCTURAL DEEP-SCANNER GENERATION PIPELINE
# =====================================================================
output_filename = f"{last_name}_{first_name}_{TEMPLATE_NAME}_Final.docx"
output_docx_path = os.path.join(output_dir, output_filename)

try:
    jinja_env = jinja2.Environment(trim_blocks=True, lstrip_blocks=True)
    
    doc = DocxTemplate(template_path)
    doc.render(final_payload, jinja_env)

    # 🛠️ RECURSIVE LAYOUT ENGINE UNROLLER
    def get_all_paragraphs_recursively(element, paragraph_list=None):
        if paragraph_list is None:
            paragraph_list = []
            
        # Extract plain root-level paragraphs
        for p in element.paragraphs:
            paragraph_list.append(p)
            
        # Crawl structural grids and unroll internal tables recursively
        for table in element.tables:
            for row in table.rows:
                for cell in row.cells:
                    get_all_paragraphs_recursively(cell, paragraph_list)
        return paragraph_list

    # Generate complete list of searchable paragraphs across all layout layers
    all_searchable_paras = get_all_paragraphs_recursively(doc)

    # Footnote Anchor Coupling Engine
    if final_payload.get("experience_note"):
        footnote_target = None
        if final_payload["student"]["experience"]:
            last_p_job = final_payload["student"]["experience"][-1]
            bullets = last_p_job.get("bullets", [])
            anchor_text = clean_bullet(bullets[-1]) if bullets else ""
            
            if anchor_text:
                anchor_simplified = simplify(anchor_text)
                for p in all_searchable_paras:
                    if p.text and anchor_simplified in simplify(p.text):
                        footnote_target = p
                        break
            
            if footnote_target is None:
                anchor_text = last_p_job.get("company", "")
                if anchor_text:
                    anchor_simplified = simplify(anchor_text)
                    for p in all_searchable_paras:
                        if p.text and anchor_simplified in simplify(p.text):
                            footnote_target = p
                            break

        if footnote_target is not None:
            # Inject directly into the explicit structural parent container of the anchor paragraph
            parent_container = footnote_target._element.getparent()
            
            p_note = footnote_target._parent.add_paragraph()
            target_index = parent_container.index(footnote_target._element) + 1
            parent_container.insert(target_index, p_note._element)
            
            p_note.paragraph_format.space_before = Pt(10)
            p_note.paragraph_format.space_after = Pt(10)
            p_note.paragraph_format.line_spacing = 1.15
            
            run = p_note.add_run(final_payload["experience_note"])
            run.font.italic = True
            run.font.size = Pt(9.5)
            run.font.name = 'Calibri'
            print("📎 Verbatim chronological structural footnote successfully coupled inside the layout structure.")
        else:
            print("⚠️ Warning: Footnote anchor text target could not be resolved across any nested matrix layout layer.")

    # Added EXECUTIVE PROFILE to target keep_with_next optimization
    TARGET_HEADERS = ["SKILLS", "EXPERIENCE", "EDUCATION", "REFERENCES", "PROFESSIONAL SUMMARY", "CAREER OBJECTIVE", "EXECUTIVE PROFILE"]
    for p in all_searchable_paras:
        if p.text:
            txt_upper = p.text.strip().upper()
            if txt_upper in TARGET_HEADERS:
                p.paragraph_format.keep_with_next = True

    doc.save(output_docx_path)
    print(f"✅ Primary Resume Compiled Successfully ➔ {output_filename}")

    # --- PART 2: COMPREHENSIVE UN-STYLED PLAIN TEXT APPENDIX ---
    if create_appendix:
        appendix_filename = f"{last_name}_{first_name}_Work_History_Appendix.docx"
        appendix_path = os.path.join(output_dir, appendix_filename)
        
        app_doc = Document()
        for section in app_doc.sections:
            section.top_margin = Pt(72)
            section.bottom_margin = Pt(72)
            section.left_margin = Pt(72)
            section.right_margin = Pt(72)

        title_p = app_doc.add_paragraph()
        title_p.paragraph_format.space_after = Pt(18)
        title_run = title_p.add_run(f"Comprehensive Work History Appendix — {first_name} {last_name}")
        title_run.font.name = 'Calibri'
        title_run.font.size = Pt(14)
        title_run.font.bold = True
        
        for idx, j in enumerate(appendix_jobs_sorted):
            p_header = app_doc.add_paragraph()
            p_header.paragraph_format.space_before = Pt(12) if idx > 0 else Pt(0)
            p_header.paragraph_format.space_after = Pt(1)
            p_header.paragraph_format.keep_with_next = True
            
            job_company_str = f"{j.get('role', 'Associate')} | {j.get('company', 'Organization')}"
            run_head = p_header.add_run(job_company_str)
            run_head.font.name = 'Calibri'
            run_head.font.size = Pt(11)
            run_head.font.bold = True
            
            p_years = app_doc.add_paragraph()
            p_years.paragraph_format.space_after = Pt(3)
            p_years.paragraph_format.keep_with_next = True
            
            run_years = p_years.add_run(scrub(j.get("dates", "Timeline Not Provided")))
            run_years.font.name = 'Calibri'
            run_years.font.size = Pt(9.5)
            run_years.font.italic = True
            
            p_desc = app_doc.add_paragraph()
            p_desc.paragraph_format.space_after = Pt(8)
            p_desc.paragraph_format.line_spacing = 1.15
            
            raw_paragraph_narrative = j.get("appendix_paragraph", "").strip()
            if not raw_paragraph_narrative:
                raw_paragraph_narrative = "Executed operational procedures and task workflows in alignment with system performance metrics."
                
            run_desc = p_desc.add_run(raw_paragraph_narrative)
            run_desc.font.name = 'Calibri'
            run_desc.font.size = Pt(10.5)

        app_doc.save(appendix_path)
        print(f"📎 Appendix Triggered (>3 jobs total). Compiled successfully ➔ {appendix_filename}")
    else:
        print("ℹ️ Total jobs <= 3 or Expanded Profile used. Appendix file creation bypassed completely.")

    print(f"📂 Output Folder Setup Completed: {output_dir}")
    print("\n🏁 PIPELINE EXECUTION COMPLETED WITHOUT ERRORS\n")

except Exception as e:
    print(f"❌ Stage 3 Rendering Pipeline Aborted: {str(e)}")
    sys.exit(1)

    JOB HUNT
import os
import sys
import re
import requests
import json
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# =====================================================================
# 🛠️ SETTINGS & UTILITIES
# =====================================================================
# IMPORTANT: Insert your actual Serper API Key below
API_KEY = "SERPER_API_KEY"

def add_active_hyperlink(paragraph, url, text, color="0000FF", underline=True):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    if color:
        c = OxmlElement('w:color')
        c.set(qn('w:val'), color)
        rPr.append(c)
    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)
    new_run.append(rPr)
    text_node = OxmlElement('w:t')
    text_node.text = text
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def extract_text_from_docx(file_path):
    if not os.path.exists(file_path): return ""
    doc = Document(file_path)
    text_tokens = []
    for p in doc.paragraphs:
        if p.text.strip(): text_tokens.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text.strip(): text_tokens.append(p.text.strip())
    return " ".join(text_tokens)

# =====================================================================
# 📂 INITIALIZATION & PATH RESOLVER
# =====================================================================
if len(sys.argv) >= 5:
    RAW_FILE_NAME = sys.argv[1]
    HAS_APPENDIX_FLAG = sys.argv[2].strip()  
    TARGET_LOCALITY = sys.argv[3].strip()    
    TARGET_TITLE = sys.argv[4].strip()       
else:
    sys.exit(1)

tokens = RAW_FILE_NAME.replace(".docx", "").split("_")
last_name, first_name = tokens[0], tokens[1]
TARGET_FOLDER_NAME = f"{last_name}, {first_name}".lower().strip()
BASE_DIR = r"C:\Resume_System\Student_Records"
OUTPUT_DIR = None
MAIN_RESUME_PATH = None

for root, dirs, files in os.walk(BASE_DIR):
    if os.path.basename(root).lower().strip() == TARGET_FOLDER_NAME:
        OUTPUT_DIR = root
        potential_resume = os.path.join(root, RAW_FILE_NAME)
        MAIN_RESUME_PATH = potential_resume if os.path.exists(potential_resume) else None
        break

if not OUTPUT_DIR:
    sys.exit(1)

# =====================================================================
# 🧠 ANALYSIS & API FETCHING
# =====================================================================
full_resume_text = extract_text_from_docx(MAIN_RESUME_PATH) if MAIN_RESUME_PATH else f"{TARGET_TITLE} {TARGET_LOCALITY}"

# Call Serper API for structured data
url = "https://google.serper.dev/search"
payload = json.dumps({"q": f"{TARGET_TITLE} jobs in {TARGET_LOCALITY}"})
headers = {'X-API-KEY': API_KEY, 'Content-Type': 'application/json'}
response = requests.request("POST", url, headers=headers, data=payload)
results = response.json().get('organic', [])

# =====================================================================
# 📄 GENERATE DOCUMENT WITH TABLE
# =====================================================================
report_doc = Document()
report_doc.add_heading("Targeted Placement Opportunities", 0)
report_doc.add_paragraph(f"Prepared For: {first_name} {last_name} | Role: {TARGET_TITLE}")

# Create Table
table = report_doc.add_table(rows=1, cols=3)
table.style = 'Table Grid'
hdr = table.rows[0].cells
hdr[0].text = 'Position'
hdr[1].text = 'Source/Snippet'
hdr[2].text = 'Action'

for job in results[:8]:
    row = table.add_row().cells
    row[0].text = job.get('title', 'N/A')
    row[1].text = job.get('snippet', 'N/A')
    add_active_hyperlink(row[2].add_paragraph(), job.get('link', '#'), "Apply Now")

report_doc.save(os.path.join(OUTPUT_DIR, f"{last_name}_{first_name}_Matched_Local_Postings.docx"))
print("Success")
