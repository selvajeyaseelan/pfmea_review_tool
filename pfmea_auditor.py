import os
import re
import copy
import time
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv
from openai import OpenAI
from xlsx2html import xlsx2html

load_dotenv()

# Initialize OpenAI Client
api_key = os.getenv("API_KEY_S30") or os.getenv("API_KEY_Pro") or "euri-8f710b297ee82784619c32dc5bcec4eb9eed10ea602322ce468846219c34e8a9"
client = OpenAI(
    api_key=api_key,
    base_url="https://api.euron.one/api/v1/euri",
)

# AIAG FMEA Standards (Embedded System Prompts - No Excel Extraction)
AIAG_SEVERITY_STANDARDS = """AIAG SEVERITY DEFINITIONS (1-10):
10: Product: Potential failure mode affects safe operation and/or involves noncompliance with regulations without warning. Process: May endanger operator, machine or assembly without warning.
9:  Product: Potential failure mode affects safe operation and/or involves noncompliance with regulations with warning. Process: May endanger operator, machine or assembly with warning.
8:  Product: Loss of primary function (product inoperable, does not affect safe operation). Process: Major disruption: 100% of product may have to be scrapped. Line shutdown or stop ship.
7:  Product: Degradation of primary function (product operable, but at a reduced level of performance). Process: Significant disruption: A portion of the production run may have to be scrapped. Deviation from primary process; decreased line speed or added manpower.
6:  Product: Loss of secondary function (product operable but service life greatly reduced, convenience item(s) inoperable, customer dissatisfied). Process: Moderate disruption: 100% of production run may have to be reworked off line and accepted.
5:  Product: Degradation of secondary function (product operable but appearance affected, convenience item(s) operable at a reduced level, customer dissatisfied). Process: A proportion of the production run may have to be reworked off line and accepted.
4:  Product: Appearance, fit and finish type items do not conform, defect noticed by most of the customers (>75%). Process: Moderate disruption: 100% of production run may have to be reworked in station before it is processed.
3:  Product: Appearance, fit and finish type items do not conform, defect noticed by about half of the customers (50%). Process: A proportion of the production run may have to be reworked in station before it is processed.
2:  Product: Appearance, fit and finish type items do not conform, defect noticed by discriminating customers (<25%). Process: Minor disruption: Slight inconvenience to process, operation or operator.
1:  Product: No discernible effect. Process: No discernible effect."""

AIAG_OCCURRENCE_STANDARDS = """AIAG OCCURRENCE DEFINITIONS (1-10):
10: Very High - Inevitable failure: >= 100 per thousand (>= 1 in 10) | PPM: 500,000 | >1 per shift (100% of production).
9:  High - Repeated failures: 50 per thousand (1 in 20) | PPM: 50,000 | >1 per day (50% of production).
8:  High - Relatively high failure rate: 20 per thousand (1 in 50) | PPM: 20,000 | >1 per 2-3 days (20% of production).
7:  Moderate - Occasional failure rate: 10 per thousand (1 in 100) | PPM: 10,000 | >1 per week (10% of production).
6:  Moderate - Moderate failure rate: 2 per thousand (1 in 500) | PPM: 5,000 | >1 per 2 weeks / 1 per month (5% of production).
5:  Moderate - Relatively low failure rate: 0.5 per thousand (1 in 2,000) | PPM: 1,000 | >1 per quarter / 2 per year (0.5% of production).
4:  Moderately Low - Infrequent failure rate: 0.1 per thousand (1 in 10,000) | PPM: 100 | >1 per half year / 1 per year (0.1% of production).
3:  Low - Isolated failures: 0.01 per thousand (1 in 100,000) | PPM: 10 | >1 per year / 1 per 5 years (0.05% of production).
2:  Low - Extremely few failures: < 0.001 per thousand (1 in 1,000,000) | PPM: 1 | <1 per year / 1 per 10 years (0.01% of production).
1:  Very Low - Failure eliminated through preventive controls: Almost never | PPM: 0 | Less than 0.01% of production."""

AIAG_DETECTION_STANDARDS = """AIAG DETECTION DEFINITIONS (1-10):
- 10 (Absolute Uncertainty): No current process control; Cannot detect or compliance analysis is not performed. 100% Human Inspection.
- 9 (Difficult to Detect): Defect (Failure Mode) and/or Error (Cause) is not easily detected (e.g. Random audits).
- 8 (Defect Detection Post Processing): Defect (Failure Mode) detection post-processing by operator through visual/tactile/audible means with no boundary samples.
- 7 (Defect Detection at Source): Defect (Failure Mode) detection in-station by operator through visual/tactile/audible means or post-processing through use of attribute gauging (go/no-go, manual torque check/clicker wrench, etc.) with no boundary samples. Manual gauging used on every part.
- 6 (Defect Detection Post Processing): Defect (Failure Mode) detection post-processing by operator through use of variable gauging or in-station by operator through use of attribute gauging (go/no-go, manual torque check/clicker wrench, etc.) with boundary samples.
- 5 (Defect Detection at Source): Defect (Failure Mode) or Error (Cause) detection in-station by operator through use of variable gauging or by automated controls that will detect discrepant part and notify operator (light, buzzer, etc.). Automatic gauging or controls; setup and first-piece check.
- 4 (Defect Detection Post Processing): Defect (Failure Mode) detection post-processing by automated controls that will detect discrepant part and lock part to prevent further processing (controls in place for mistake proofing the assembly).
- 3 (Defect Detection at Source): Defect (Failure Mode) detection in-station by automated controls that will detect discrepant part and automatically lock part in station to prevent further processing.
- 2 (Error Detection and/or Defect Prevention): Error (Cause) detection in-station by automated controls that will detect error and prevent discrepant part from being made.
- 1 (Detection not applicable): Error (Cause) prevention as a result of fixture design, machine design or part design (Poka-Yoke / mistake proofing)."""

def extract_number(val):
    if val is None: return None
    if isinstance(val, (int, float)): return int(val)
    matches = re.findall(r'\d+', str(val))
    return int(matches[0]) if matches else None

def get_det_zone(sev, det):
    if sev is None or det is None: return ""
    try: s, d = int(sev), int(det)
    except ValueError: return ""
    if not (1 <= s <= 10 and 1 <= d <= 10): return ""
    matrix = {
        10: {1:3, 2:2, 3:1, 4:1, 5:1, 6:1, 7:1, 8:1, 9:1, 10:1},
        9:  {1:3, 2:2, 3:1, 4:1, 5:1, 6:1, 7:1, 8:1, 9:1, 10:1},
        8:  {1:3, 2:2, 3:2, 4:2, 5:2, 6:2, 7:1, 8:1, 9:1, 10:1},
        7:  {1:3, 2:3, 3:3, 4:2, 5:2, 6:2, 7:2, 8:1, 9:1, 10:1},
        6:  {1:3, 2:3, 3:3, 4:3, 5:3, 6:2, 7:2, 8:1, 9:1, 10:1},
        5:  {1:3, 2:3, 3:3, 4:3, 5:3, 6:3, 7:3, 8:2, 9:2, 10:2},
        4:  {1:3, 2:3, 3:3, 4:3, 5:3, 6:3, 7:3, 8:2, 9:2, 10:2},
        3:  {1:3, 2:3, 3:3, 4:3, 5:3, 6:3, 7:3, 8:3, 9:3, 10:3},
        2:  {1:3, 2:3, 3:3, 4:3, 5:3, 6:3, 7:3, 8:3, 9:3, 10:3},
        1:  {1:3, 2:3, 3:3, 4:3, 5:3, 6:3, 7:3, 8:3, 9:3, 10:3}
    }
    return matrix[d][s]

def get_sev_zone(sev, occ):
    if sev is None or occ is None: return ""
    try: s, o = int(sev), int(occ)
    except ValueError: return ""
    if not (1 <= s <= 10 and 1 <= o <= 10): return ""
    matrix = {
        10: {1:3, 2:1, 3:1, 4:1, 5:1, 6:1, 7:1, 8:1, 9:1, 10:1},
        9:  {1:3, 2:1, 3:1, 4:1, 5:1, 6:1, 7:1, 8:1, 9:1, 10:1},
        8:  {1:3, 2:2, 3:1, 4:1, 5:1, 6:1, 7:1, 8:1, 9:1, 10:1},
        7:  {1:3, 2:2, 3:2, 4:2, 5:1, 6:1, 7:1, 8:1, 9:1, 10:1},
        6:  {1:3, 2:2, 3:2, 4:2, 5:1, 6:1, 7:1, 8:1, 9:1, 10:1},
        5:  {1:3, 2:3, 3:2, 4:2, 5:2, 6:2, 7:1, 8:1, 9:1, 10:1},
        4:  {1:3, 2:3, 3:3, 4:3, 5:2, 6:2, 7:1, 8:1, 9:1, 10:1},
        3:  {1:3, 2:3, 3:3, 4:3, 5:3, 6:3, 7:2, 8:2, 9:1, 10:1},
        2:  {1:3, 2:3, 3:3, 4:3, 5:3, 6:3, 7:2, 8:2, 9:1, 10:1},
        1:  {1:3, 2:3, 3:3, 4:3, 5:3, 6:3, 7:3, 8:3, 9:3, 10:3}
    }
    return matrix[o][s]

def get_priority_level(sev_zone, det_zone):
    if sev_zone is None or det_zone is None: return ""
    try: sz, dz = int(sev_zone), int(det_zone)
    except ValueError: return ""
    if not (1 <= sz <= 3 and 1 <= dz <= 3): return ""
    matrix = {
        3: {1: 2, 2: 2, 3: 3},
        2: {1: 1, 2: 2, 3: 3},
        1: {1: 1, 2: 1, 3: 2}
    }
    return matrix[dz][sz]

def infer_aiag_severity_from_text(effect_text):
    """Return the highest severity explicitly supported by the effect wording."""
    normalized = re.sub(r"\s+", " ", str(effect_text).lower())
    # Scores in the source cell are annotations, not semantic evidence.
    normalized = re.sub(r"\([^)]*\)", " ", normalized)

    rules = [
        (10, r"(?:safe operation|noncompliance|regulat).{0,100}without warning|endanger.{0,100}without warning"),
        (9, r"(?:safe operation|noncompliance|regulat).{0,100}with warning|endanger.{0,100}with warning"),
        (8, r"loss of primary function|line shutdown|stop ship|100\s*%[^.]{0,100}(?:scrap|scrapped)"),
        (7, r"degradation of primary function|\b(?:a )?(?:portion|proportion) of (?:the )?production run[^.]{0,100}scrap"),
        (6, r"loss of secondary function|100\s*%[^.]{0,100}rework[^.]{0,50}off[- ]line"),
        (5, r"degradation of secondary function|\b(?:a )?(?:portion|proportion) of (?:the )?production run[^.]{0,100}rework[^.]{0,50}off[- ]line"),
        (4, r"100\s*%[^.]{0,100}rework[^.]{0,50}in[- ]station|noticed by most|more than 75\s*%|>\s*75\s*%"),
        (3, r"(?:a )?(?:portion|proportion) of (?:the )?production run[^.]{0,100}rework[^.]{0,50}in[- ]station|noticed by about half|50\s*%"),
        (2, r"minor disruption|slight inconvenience|discriminating customers|less than 25\s*%|<\s*25\s*%"),
        (1, r"no discernible effect"),
    ]
    matches = [severity for severity, pattern in rules if re.search(pattern, normalized)]
    return max(matches) if matches else None

def evaluate_severity_with_llm(effect_text, original_sev):
    if not effect_text or str(effect_text).strip().lower() == 'nan': return "", "", 0, 0
    try: clean_original_sev = int(float(original_sev))
    except (ValueError, TypeError): clean_original_sev = str(original_sev).strip()

    deterministic_severity = infer_aiag_severity_from_text(effect_text)
    if deterministic_severity is not None:
        severity = str(deterministic_severity)
        reason = "severity is correct" if clean_original_sev == deterministic_severity else (
            f"The effect text maps to AIAG Severity {deterministic_severity}, "
            f"but the stated severity is {clean_original_sev}. The text-derived severity "
            "is the highest severity supported by the listed effects."
        )
        return severity, reason, 0, 0

    system_prompt = f"""
    You are an expert AIAG PFMEA auditor. Audit the failure-effect text in two independent passes.

    {AIAG_SEVERITY_STANDARDS}

    REQUIRED TWO-PASS CHECK:
    1. Read EVERY labelled effect separately (for example, S and OEM). Ignore the score in
       parentheses after a label while classifying that effect's words.
    2. Assign each effect the AIAG severity whose definition is actually described by its
       words. Then select the HIGHEST of those text-derived severities.
    3. Independently compare that highest text-derived severity with the provided Stated Severity.
       The result is a mismatch if either the text-derived severity differs from Stated Severity
       OR Stated Severity is not the highest severity supported by the effect texts.

        IMPORTANT AIAG DISTINCTION:
        - "A portion of the production run may have to be scrapped" is Severity 7.
        - Severity 8 requires a major disruption such as "100% of the product/production run may
            have to be scrapped", line shutdown, or stop ship.
        - Do not upgrade Severity 7 to 8 merely because the word "scrapped" appears.

    Do not use a number in parentheses as semantic evidence. For example, in
    "S: ... scrapped (5) OEM: ... scrapped (7)", classify both texts, choose 7, and compare
    that result with Stated Severity.

    4. Output your response EXACTLY in this format on two lines:
    Severity: <integer 1-10>
    Reason: <write "severity is correct" only when both checks pass; otherwise explain the
    text-derived severity, the stated severity, and the exact failed check.>

    LEARNING EXAMPLES:
    Text: "S: Portion of production run may have to be scrapped (5) \\n OEM: Portion of production run may have to be scrapped (7)"
    Stated Severity: 7
    Output:
    Severity: 7
    Reason: severity is correct

    Text: "MismatchEU: Loss of primary function (product inoperable, does not affect safe operation) (4)"
    Stated Severity: 4
    Output:
    Severity: 8
    Reason: 'Loss of primary function' maps to AIAG 8. The stated (4) is an under-rating.

    IMPORTANT FOR VAGUE TEXTS: 
    Never default to Severity 1 for vague failure texts like "Reduced performance" or "Poor life". Vague texts indicating reduced performance or degradation must map closest to Severity 7 (Degradation of primary function). Do not predict 1 unless the text explicitly states "No discernible effect".
    """
    try:
        response = client.chat.completions.create(
            model="gpt-5.4-mini",
            messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": f"Stated Severity: {clean_original_sev}\\nText: {effect_text}\\nOutput:"}],
            temperature=0.0
        )
        output_text = response.choices[0].message.content.strip()
        severity, reason = "", ""
        for line in output_text.split('\n'):
            if line.startswith("Severity:"): severity = line.replace("Severity:", "").strip()
            elif line.startswith("Reason:"): reason = line.replace("Reason:", "").strip()

        return severity, reason, response.usage.prompt_tokens, response.usage.completion_tokens
    except Exception as e:
        return "", f"Error: {str(e)}", 0, 0

def extract_strongest_detection(det_text):
    if not det_text: return None
    lowest_val = 999
    matches = re.findall(r"\((\d+)\)", str(det_text)) or re.findall(r"\b(\d+)\b", str(det_text))
    if matches:
        for m in matches:
            if int(m) < lowest_val: lowest_val = int(m)
    return lowest_val if lowest_val != 999 else None

def evaluate_detection_with_llm(full_det_text, original_det, pred_det, det_definitions=None):
    if not full_det_text: return "", 0, 0
    try: clean_original_det = int(float(original_det))
    except (ValueError, TypeError): clean_original_det = str(original_det).strip()
    
    definitions = det_definitions if det_definitions else AIAG_DETECTION_STANDARDS
    system_prompt = f"""
    You are an expert AIAG PFMEA auditor. Your job is to strictly verify the consistency of the 'Current Process Detection Controls' column against the assigned Detection rankings.

    {definitions}

    INSTRUCTIONS FOR AUDIT:
    I will provide you with:
    1. Original Stated Detection: (Variable A)
    2. Strongest Detection Extracted: (Variable B)
    3. The Full Detection Controls Text

    You must perform two checks:
    CHECK 1: Does Variable A exactly match Variable B?
    CHECK 2: SEMANTIC VERIFICATION (CRITICAL). Read the actual words in "Full Detection Controls Text". Do they describe the AIAG criteria for the lowest number written next to them? 
    - If the text describes a weak control (e.g., "visual inspection", "random audit") but claims a strong number (e.g., 2, 3, 4), Check 2 FAILS (Over-rating).
    - If the text describes a strong control (e.g., "automated lock", "error proofing") but claims a weak number (e.g., 7, 8), Check 2 FAILS (Under-rating).

    CRITICAL OUTPUT CONSTRAINTS:
    You must format your response EXACTLY like one of the following two templates.
    TEMPLATE 1 (Pass): "detection controls text matches AIAG detection criteria"
    TEMPLATE 2 (Fail): "detection controls text keywords do not match AIAG detection criteria. [State exactly what failed]. Recommended Detection: [Insert the CORRECT AIAG number]"
    """
    try:
        response = client.chat.completions.create(
            model="gpt-5.4-mini", 
            messages=[
                {"role": "system", "content": system_prompt}, 
                {"role": "user", "content": f"Original Stated Detection: {clean_original_det}\nStrongest Detection Extracted: {pred_det}\nFull Detection Controls Text:\n{full_det_text}"}
            ],
            temperature=0.0
        )
        return response.choices[0].message.content.strip(), response.usage.prompt_tokens, response.usage.completion_tokens
    except Exception as e:
        return f"Error contacting LLM: {str(e)}", 0, 0

def inject_modern_css(html_path):
    with open(html_path, 'r', encoding='utf-8') as f: html_content = f.read()
    parts = html_content.split('<tr')
    target_idx = next((i for i, part in enumerate(parts) if "PROCESS FAILURE MODE AND EFFECTS ANALYSIS" in part.upper()), -1)
    if target_idx > 1:
        for i in range(1, target_idx): parts[i] = ' class="hidden-row" ' + parts[i]
    html_content = '<tr'.join(parts)
    html_content = re.sub(r'(<td[^>]*>)(.*?PROCESS FAILURE MODE AND EFFECTS ANALYSIS.*?)(</td>)', r'\1<div class="pfmea-main-header">\2</div>\3', html_content, flags=re.IGNORECASE | re.DOTALL)
    modern_css = """<style>@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap'); body { font-family: 'Inter', sans-serif; background-color: #f4f7f9; color: #334155; margin: 40px; } table { border-collapse: collapse; width: 100%; background-color: #ffffff; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); font-size: 14px; } table, th, td { border: 1px solid #e2e8f0 !important; } th, td { padding: 12px 16px !important; line-height: 1.5; } .hidden-row { display: none !important; } .pfmea-main-header { font-size: 24px !important; font-weight: 700 !important; background-color: #1e3a8a !important; color: #ffffff !important; padding: 16px !important; text-align: center !important; border-radius: 6px; text-transform: uppercase; } tr:nth-child(-n+11) td { background-color: #ffffff !important; border: none !important; font-size: 13px; color: #64748b !important; } tr:nth-child(12) td, tr:nth-child(13) td { background-color: #0f172a !important; color: #ffffff !important; font-weight: 600 !important; text-transform: uppercase; font-size: 12px; } tr:nth-child(n+14):nth-child(even) td { background-color: #f8fafc !important; } tr:nth-child(n+14):hover td { background-color: #f1f5f9 !important; transition: background-color 0.2s ease; } td[style*="color: #CC0000"], td[style*="color: CC0000"] { color: #b91c1c !important; background-color: #fef2f2 !important; border: 1px solid #fca5a5 !important; font-size: 14px !important; font-weight: bold !important; } td[style*="background-color: #FFFF00"], td[style*="background-color: #ffff00"], td[style*="background-color: FFFF00"] { background-color: #FFFF00 !important; color: #FF0000 !important; font-weight: bold !important; border: 1px solid #eab308 !important; }</style>"""
    html_content = html_content.replace('</head>', f'{modern_css}\n</head>') if '</head>' in html_content else modern_css + html_content
    with open(html_path, 'w', encoding='utf-8') as f: f.write(html_content)

def audit_pfmea_generator(input_file, output_file="pfmea_audited.xlsx", output_html="audited_pfmea.html", target_priority=1, det_threshold=4):
    """
    Generator that processes and audits the PFMEA workbook row-by-row,
    yielding real-time status and structured row objects for web streaming.
    """
    try:
        target_priority = int(target_priority)
    except (ValueError, TypeError):
        target_priority = 1

    try:
        det_threshold = float(det_threshold)
    except (ValueError, TypeError):
        det_threshold = 4.0

    yield {
        "type": "log",
        "message": f"Loading workbook '{os.path.basename(input_file)}' and setting up 'PFMEA_Audited' sheet..."
    }

    wb = openpyxl.load_workbook(input_file)
    if "PFMEA" not in wb.sheetnames:
        yield {"type": "error", "message": "Input workbook does not contain required 'PFMEA' sheet."}
        return

    ws = wb["PFMEA"]

    if "PFMEA_Audited" in wb.sheetnames:
        del wb["PFMEA_Audited"]
    out_ws = wb.copy_worksheet(ws)
    out_ws.title = "PFMEA_Audited"

    for rng in list(out_ws.merged_cells.ranges):
        if rng.min_row >= 12:
            out_ws.unmerge_cells(str(rng))

    total_input_tokens, total_output_tokens = 0, 0

    # 1. Map original columns
    col_map = {}
    for col in range(1, ws.max_column + 1):
        val11 = str(ws.cell(11, col).value or "").strip().lower()
        val12 = str(ws.cell(12, col).value or "").strip().lower()
        val = val11 + " " + val12
        if not val.strip(): continue
        if "process" in val and "prevention" not in val and "detection" not in val: col_map["process"] = col
        elif "potential failure mode" in val: col_map["potential failure mode"] = col
        elif "effect" in val: col_map["potential effects of failure"] = col
        elif "sev" == val12 or "sev" == val11: col_map["sev"] = col
        elif "cls" == val12 or "cls" == val11: col_map["cls"] = col
        elif "cause" in val or "mechanism" in val: col_map["potential causes / mechanisms of failure"] = col
        elif "prevention" in val: col_map["current process prevention controls"] = col
        elif "occ" == val12 or "occ" == val11: col_map["occ"] = col
        elif "detection controls" in val: col_map["current process detection controls"] = col
        elif "det" == val12 or "det" == val11: col_map["det"] = col 
        elif "rpn" == val12 or "rpn" == val11: col_map["rpn"] = col 
        elif "recommended action" in val or "rec act" in val: col_map["recommended actions"] = col
        elif "severity zone" in val or "sev zone" in val: col_map["severity zone"] = col
        elif "detection zone" in val or "det zone" in val: col_map["detection zone"] = col
        elif "priority level" in val: col_map["priority level"] = col

    # 2. Comprehensive Target Columns Order
    targets = [
        ("Process", col_map.get("process")),
        ("Potential Failure Mode", col_map.get("potential failure mode")),
        ("Potential Effects of Failure", col_map.get("potential effects of failure")),
        ("Sev", col_map.get("sev")),
        ("Pred Sev", col_map.get("sev")), 
        ("Sev Reasoning", col_map.get("sev")),
        ("Cls", col_map.get("cls")),
        ("Potential Causes / Mechanisms of Failure", col_map.get("potential causes / mechanisms of failure")),
        ("Current Process Prevention Controls", col_map.get("current process prevention controls")),
        ("Occ", col_map.get("occ")),
        ("Current Process Detection Controls", col_map.get("current process detection controls")),
        ("Det", col_map.get("det")),
        ("Pred Det", col_map.get("det")),
        ("Det Reasoning", col_map.get("det")),
        ("RPN", col_map.get("rpn")),
        ("D-P gap", col_map.get("rpn")),
        ("New RPN", col_map.get("rpn")),
        ("Recommended Actions", col_map.get("recommended actions")),
        ("Severity Zone", col_map.get("severity zone")),
        ("Pred Sev Zone", col_map.get("severity zone") or col_map.get("rpn")),
        ("Detection Zone", col_map.get("detection zone")),
        ("Pred Det Zone", col_map.get("detection zone") or col_map.get("rpn")),
        ("Priority Level", col_map.get("priority level")),
        ("Pred Priority Level", col_map.get("priority level") or col_map.get("rpn")),
    ]

    # Setup headers in out_ws
    for idx, (title, ref_col) in enumerate(targets, 1):
        cell12 = out_ws.cell(12, idx)
        cell13 = out_ws.cell(13, idx)
        cell12.value = title
        out_ws.merge_cells(start_row=12, start_column=idx, end_row=13, end_column=idx)
        if ref_col:
            cell12._style = copy.copy(ws.cell(12, ref_col)._style)
            cell13._style = copy.copy(ws.cell(13, ref_col)._style)
            width = ws.column_dimensions[get_column_letter(ref_col)].width
            out_ws.column_dimensions[get_column_letter(idx)].width = width or 15
        else:
            out_ws.column_dimensions[get_column_letter(idx)].width = 15
        if "Reasoning" in title or title in ["Potential Effects of Failure", "Current Process Prevention Controls"]:
            out_ws.column_dimensions[get_column_letter(idx)].width = 50
        cell12.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if out_ws.max_column > len(targets):
        out_ws.delete_cols(len(targets) + 1, out_ws.max_column - len(targets))

    total_rows = max(0, ws.max_row - 13)
    yield {
        "type": "init",
        "total_rows": total_rows,
        "columns": [t[0] for t in targets]
    }

    violations_count = 0
    sev_discrepancies_count = 0
    det_discrepancies_count = 0

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    red_font = Font(color="FF0000", bold=True)
    mismatch_font = Font(color="CC0000")

    # Evaluate each row
    for row_idx, row in enumerate(range(14, ws.max_row + 1), 1):
        o_s = extract_number(ws.cell(row, col_map.get("sev", 0)).value if col_map.get("sev") else None)
        o_o = extract_number(ws.cell(row, col_map.get("occ", 0)).value if col_map.get("occ") else None)
        o_d = extract_number(ws.cell(row, col_map.get("det", 0)).value if col_map.get("det") else None)
        
        proc_val = str(ws.cell(row, col_map.get("process", 0)).value or "") if col_map.get("process") else ""
        fm_val = str(ws.cell(row, col_map.get("potential failure mode", 0)).value or "") if col_map.get("potential failure mode") else ""
        effect_text = ws.cell(row, col_map.get("potential effects of failure", 0)).value if col_map.get("potential effects of failure") else None
        det_text = ws.cell(row, col_map.get("current process detection controls", 0)).value if col_map.get("current process detection controls") else None

        yield {
            "type": "log",
            "message": f"Row {row_idx}/{total_rows} (Excel Row {row}): Auditing '{proc_val[:30]}' - '{fm_val[:30]}'"
        }

        # 1. Evaluate Severity
        pred_sev, sev_reasoning = None, ""
        if effect_text and o_s is not None:
            pred_sev_str, sev_reasoning, s_in, s_out = evaluate_severity_with_llm(effect_text, o_s)
            pred_sev = extract_number(pred_sev_str)
            total_input_tokens += s_in
            total_output_tokens += s_out

        # 2. Evaluate Detection
        pred_det, det_reasoning = None, ""
        if det_text and o_d is not None:
            ext_det = extract_strongest_detection(det_text)
            det_reasoning, d_in, d_out = evaluate_detection_with_llm(det_text, o_d, ext_det)
            match = re.search(r"Recommended Detection:\s*(\d+)", det_reasoning, re.IGNORECASE)
            pred_det = int(match.group(1)) if match else ext_det
            total_input_tokens += d_in
            total_output_tokens += d_out

        p_s = pred_sev if pred_sev is not None else o_s
        p_d = pred_det if pred_det is not None else o_d

        # 3. D-P Gap logic
        dp_gap_val, dp_gap_flag = "", False
        if p_s in [9, 10]:
            if p_d is not None and p_d <= 3:
                dp_gap_val = "no gap"
            else:
                dp_gap_val = "Gap identified. Det must be <3"
                dp_gap_flag = True
        elif p_s is not None and p_s > 0:
            dp_gap_val = "no gap"

        # Check severity / detection discrepancies
        has_sev_mismatch = (sev_reasoning and "severity is correct" not in sev_reasoning.lower())
        has_det_mismatch = ("keywords do not match" in det_reasoning.lower())
        if has_sev_mismatch: sev_discrepancies_count += 1
        if has_det_mismatch: det_discrepancies_count += 1

        # Calculate zones and priority
        sz = get_sev_zone(p_s, o_o)
        dz = get_det_zone(p_s, p_d)
        prio = get_priority_level(sz, dz)

        orig_sz = extract_number(ws.cell(row, col_map.get("severity zone", 0)).value if col_map.get("severity zone") else None)
        orig_dz = extract_number(ws.cell(row, col_map.get("detection zone", 0)).value if col_map.get("detection zone") else None)
        orig_prio = extract_number(ws.cell(row, col_map.get("priority level", 0)).value if col_map.get("priority level") else None)

        # Audit check: Highlight violation if Pred Priority Level == target_priority but Pred Det > det_threshold (or missing/invalid)
        is_det_violation = False
        try:
            if p_d is None or float(p_d) > det_threshold:
                is_det_violation = True
        except (ValueError, TypeError):
            is_det_violation = True

        is_prio_violation = (prio == target_priority and is_det_violation)
        if is_prio_violation:
            violations_count += 1

        row_data = {}

        # Populate out_ws
        for idx, (title, ref_col) in enumerate(targets, 1):
            cell = out_ws.cell(row, idx)
            if ref_col:
                cell._style = copy.copy(ws.cell(row, ref_col)._style)
            
            if "Reasoning" in title:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            val_to_store = None
            if title == "Pred Sev":
                val_to_store = pred_sev if pred_sev is not None else ""
                cell.value = val_to_store
                if has_sev_mismatch: cell.font = mismatch_font
            elif title == "Sev Reasoning":
                val_to_store = sev_reasoning
                cell.value = val_to_store
                if has_sev_mismatch: cell.font = mismatch_font
            elif title == "Pred Det":
                val_to_store = pred_det if pred_det is not None else ""
                cell.value = val_to_store
                if has_det_mismatch: cell.font = mismatch_font
            elif title == "Det Reasoning":
                val_to_store = det_reasoning
                cell.value = val_to_store
                if has_det_mismatch: cell.font = mismatch_font
            elif title == "D-P gap":
                val_to_store = dp_gap_val
                cell.value = val_to_store
                if dp_gap_flag: cell.font = mismatch_font
            elif title == "RPN":
                val_to_store = (o_s * o_o * o_d) if (o_s is not None and o_o is not None and o_d is not None) else ""
                cell.value = val_to_store
            elif title == "New RPN":
                val_to_store = (p_s * o_o * p_d) if (p_s is not None and o_o is not None and p_d is not None) else ""
                cell.value = val_to_store
            elif title == "Pred Sev Zone":
                val_to_store = sz
                cell.value = val_to_store
                if orig_sz is not None and orig_sz != sz: cell.font = mismatch_font
            elif title == "Pred Det Zone":
                val_to_store = dz
                cell.value = val_to_store
                if orig_dz is not None and orig_dz != dz: cell.font = mismatch_font
            elif title == "Pred Priority Level":
                val_to_store = prio
                cell.value = val_to_store
                if is_prio_violation:
                    cell.fill = yellow_fill
                    cell.font = red_font
                elif orig_prio is not None and orig_prio != prio:
                    cell.font = mismatch_font
            else:
                val_to_store = ws.cell(row, ref_col).value if ref_col else ""
                cell.value = val_to_store

            row_data[title] = val_to_store

        # Stream row object to browser
        yield {
            "type": "row",
            "row_index": row_idx,
            "total_rows": total_rows,
            "progress_pct": round((row_idx / total_rows) * 100, 1),
            "data": row_data,
            "flags": {
                "is_prio_violation": is_prio_violation,
                "is_prio1_violation": is_prio_violation,
                "has_sev_mismatch": has_sev_mismatch,
                "has_det_mismatch": has_det_mismatch,
                "dp_gap_flag": dp_gap_flag,
                "target_priority": target_priority,
                "det_threshold": det_threshold
            },
            "stats": {
                "violations_count": violations_count,
                "sev_discrepancies": sev_discrepancies_count,
                "det_discrepancies": det_discrepancies_count,
                "input_tokens": total_input_tokens,
                "output_tokens": total_output_tokens
            }
        }

    # Save finalized output files
    wb.save(output_file)
    wb.close()
    
    yield {"type": "log", "message": f"Saved audited Excel file to '{output_file}'"}

    try:
        xlsx2html(output_file, output_html, sheet="PFMEA_Audited")
        inject_modern_css(output_html)
        yield {"type": "log", "message": f"Generated interactive HTML report '{output_html}'"}
    except Exception as e:
        yield {"type": "log", "message": f"Note on HTML generation: {str(e)}"}

    yield {
        "type": "complete",
        "message": "Auditing completed successfully!",
        "summary": {
            "total_rows": total_rows,
            "violations_count": violations_count,
            "sev_discrepancies": sev_discrepancies_count,
            "det_discrepancies": det_discrepancies_count,
            "input_tokens": total_input_tokens,
            "output_tokens": total_output_tokens,
            "target_priority": target_priority,
            "det_threshold": det_threshold,
            "excel_filename": os.path.basename(output_file),
            "html_filename": os.path.basename(output_html)
        }
    }
