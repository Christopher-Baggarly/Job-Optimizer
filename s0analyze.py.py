import os
import json
import re
import logging
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==============================================================================
# LOGGING SETUP
# ==============================================================================
logging.basicConfig(
    filename="blueprint_inspector.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

INPUT_FOLDER = r"./templates"
OUTPUT_FOLDER = os.path.join(INPUT_FOLDER, "blueprints")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ==============================================================================
# STYLISTIC HELPER FUNCTIONS (XML DEEP DIG)
# ==============================================================================
def get_style_attribs(run_element):
    styles = {
        "font_size": None,
        "color": None,
        "bold": False,
        "italic": False,
        "style_name": None
    }
    rPr = run_element.find(qn('w:rPr'))
    if rPr is not None:
        sz = rPr.find(qn('w:sz'))
        if sz is not None:
            styles["font_size"] = int(sz.get(qn('w:val'))) / 2
            
        color = rPr.find(qn('w:color'))
        if color is not None:
            styles["color"] = color.get(qn('w:val'))
            
        if rPr.find(qn('w:b')) is not None: styles["bold"] = True
        if rPr.find(qn('w:i')) is not None: styles["italic"] = True
        
        rStyle = rPr.find(qn('w:rStyle'))
        if rStyle is not None:
            styles["style_name"] = rStyle.get(qn('w:val'))
            
    return styles

# ==============================================================================
# CORE XML STRUCTURAL WALKER
# ==============================================================================
class DocumentXmlWalker:
    def __init__(self, doc):
        self.doc = doc
        self.blocks = []
        self.has_textboxes = False
        
    def walk_element(self, element, current_location="body"):
        for child in element.iterchildren():
            tag = child.tag
            
            if tag.endswith('tbl'):
                self.walk_table(child)
                continue
                
            elif tag.endswith('txbxContent'):
                self.has_textboxes = True
                self.walk_element(child, current_location="floating_textbox")
                continue
                
            elif tag.endswith('p'):
                p_text = []
                p_styles = []
                
                pPr = child.find(qn('w:pPr'))
                p_style_name = "Normal"
                if pPr is not None:
                    pStyle = pPr.find(qn('w:pStyle'))
                    if pStyle is not None:
                        p_style_name = pStyle.get(qn('w:val'))
                
                for run in child.findall(qn('w:r')):
                    t_node = run.find(qn('w:t'))
                    if t_node is not None and t_node.text:
                        p_text.append(t_node.text)
                        p_styles.append(get_style_attribs(run))
                            
                full_p_text = "".join(p_text).strip()
                if full_p_text:
                    self.blocks.append({
                        "text": full_p_text,
                        "location_context": current_location,
                        "paragraph_style": p_style_name,
                        "run_formatting": p_styles
                    })
                    
            elif tag.endswith('drawing') or tag.endswith('wsp') or tag.endswith('fallback'):
                self.walk_element(child, current_location="shape_canvas")
            else:
                if len(child) > 0:
                    self.walk_element(child, current_location)

    def walk_table(self, tbl_element):
        for r_idx, row in enumerate(tbl_element.findall(qn('w:tr'))):
            for c_idx, cell in enumerate(row.findall(qn('w:tc'))):
                tcPr = cell.find(qn('w:tcPr'))
                cell_bg = None
                if tcPr is not None:
                    shd = tcPr.find(qn('w:shd'))
                    if shd is not None:
                        cell_bg = shd.get(qn('w:val')) or shd.get(qn('w:fill'))
                        
                context_str = f"table_row_{r_idx}_col_{c_idx}"
                if cell_bg and cell_bg != 'auto':
                    context_str += f"_bg_{cell_bg}"
                    
                self.walk_element(cell, current_location=context_str)


# ==============================================================================
# INSPECTION RULES ENGINE
# ==============================================================================
def inspect_and_evaluate(blocks, has_textboxes):
    raw_joined = " ".join([b["text"] for b in blocks])
    flat_compiled_text = re.sub(r'\s+', ' ', raw_joined).lower()
    
    loops = {
        "experience": any(x in flat_compiled_text for x in ["for job in student.experience", "student.experience"]),
        "education": any(x in flat_compiled_text for x in ["for edu in student.education", "student.education"]),
        "skills": any(x in flat_compiled_text for x in ["student.skills", "skills and expertise", "technical skills"]),
        "summary": any(x in flat_compiled_text for x in ["student.summary", "career summary", "professional summary", "intentions statement"])
    }
    
    all_fields = set()
    for b in blocks:
        matches = re.findall(r"\{\{\s*(.*?)\s*\}\}", b["text"])
        for m in matches:
            all_fields.add(m.strip())

    passed_inspection = True
    failure_reasons = []
    
    for key, exists in loops.items():
        if not exists:
            passed_inspection = False
            failure_reasons.append(f"Missing logical block structure for: {key.upper()}")
            
    return {
        "passed_inspection": passed_inspection,
        "inspection_report": failure_reasons if failure_reasons else ["VALIDATED FOR PRODUCTION INSERTION"],
        "loops_found": loops,
        "detected_fields": list(all_fields),
        "layout_indicators": {
            "contains_floating_textboxes": has_textboxes,
            "total_mapped_elements": len(blocks)
        }
    }


# ==============================================================================
# PIPELINE APPLICATION RUNNER
# ==============================================================================
def execute_system_blueprint():
    for file in os.listdir(INPUT_FOLDER):
        if not file.endswith(".docx") or file.startswith("~$"):
            continue

        file_path = os.path.join(INPUT_FOLDER, file)
        print(f"Parsing raw XML blueprint architecture for: {file}")

        try:
            doc = Document(file_path)
            walker = DocumentXmlWalker(doc)
            walker.walk_element(doc.element.body)
            
            diagnostics = inspect_and_evaluate(walker.blocks, walker.has_textboxes)
            
            blueprint = {
                "template_name": file,
                "structural_integrity": diagnostics,
                "mapped_layout_tree": walker.blocks
            }
            
            output_name = file.replace(".docx", ".json")
            output_path = os.path.join(OUTPUT_FOLDER, output_name)
            
            with open(output_path, "w", encoding="utf-8") as out_f:
                json.dump(blueprint, out_f, indent=4)
                
            if diagnostics["passed_inspection"]:
                print(f"  └── [PASS] Ready for Database Production Integration.")
                logging.info(f"{file} passed mapping inspection.")
            else:
                print(f"  └── [FAIL] Diagnostics logged: {diagnostics['inspection_report']}")
                logging.warning(f"{file} failed inspection -> {diagnostics['inspection_report']}")

        except Exception as e:
            print(f"  └── [CRITICAL EXCEPTION] Could not parse OpenXML elements for {file}: {str(e)}")
            logging.error(f"Critical mapping crash on file {file}: {str(e)}", exc_info=True)


if __name__ == "__main__":
    execute_system_blueprint()