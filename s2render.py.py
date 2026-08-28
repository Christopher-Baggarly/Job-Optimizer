import os
import json
from docxtpl import DocxTemplate, RichText
import jinja2

TEMPLATES_FOLDER = r"./templates"
DATA_JSON_PATH = r"./output/normalized_data.json"
OUTPUT_FOLDER = r"./output"

def mass_compile_resumes():
    print("[STAGE 2] Loading audited JSON student payload...")
    if not os.path.exists(DATA_JSON_PATH):
        print(f"Error: Cannot find data file at {DATA_JSON_PATH}")
        return
        
    with open(DATA_JSON_PATH, 'r', encoding='utf-8') as f:
        payload = json.load(f)
    
    # Clean structural mappings for name keys
    first_name = payload.get("First", {}).get("name", "")
    last_name = payload.get("Last", {}).get("name", "")
    payload["name"] = f"{first_name} {last_name}".strip()
    payload["title"] = "HVAC/R Technician"
    payload["first"] = {"name": first_name}
    payload["last"] = {"name": last_name}
    
    # Horizontal layout preparation (pulling from correct nested student dict)
    student_data = payload.get("student", {})
    
    raw_skills = student_data.get("skills", [])
    payload["horizontal_skills"] = "  •  ".join(raw_skills)
    
    raw_certs = student_data.get("certifications", [])
    if raw_certs and isinstance(raw_certs[0], dict):
        cert_names = [c.get("name", "") for c in raw_certs if c.get("name")]
        payload["horizontal_certs"] = "  |  ".join(cert_names)
    else:
        payload["horizontal_certs"] = "  |  ".join(raw_certs)

    # Force compact whitespace environment
    jinja_env = jinja2.Environment(
        trim_blocks=True, 
        lstrip_blocks=True,
        keep_trailing_newline=False
    )
    
    print("[STAGE 2] Commencing parallel document generation pipeline...")
    
    for file in os.listdir(TEMPLATES_FOLDER):
        if not file.endswith(".docx") or file.startswith("~$") or "blueprints" in file:
            continue
            
        template_path = os.path.join(TEMPLATES_FOLDER, file)
        
        # Determine clean output naming structure
        base_name = file.replace(".docx", "")
        output_filename = f"Sample_{base_name}_2.docx" if not base_name.endswith("_2") else f"Sample_{base_name}.docx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        try:
            doc = DocxTemplate(template_path)
            doc.render(payload, jinja_env)
            doc.save(output_path)
            print(f"  ├── [RENDERED SUCCESSFULLY] -> {output_filename}")
        except Exception as e:
            print(f"  ├── [CRITICAL LAYOUT CRASH] Failed rendering {file}: {e}")
            
    print(f"\nClean run complete! Production assets updated inside:\n{OUTPUT_FOLDER}")

if __name__ == "__main__":
    mass_compile_resumes()