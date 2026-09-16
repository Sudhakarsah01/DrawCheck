"""
extractor.py — DrawCheck AI
Extracts ALL real QA FreeText annotations from PDF.
Uses content-based title block detection instead of fixed position cutoff.
"""
import pymupdf as fitz
import re
from collections import Counter

HVAC_PREFIXES = [
    "FCU-","AHU-","FD-","VCD-","MVCD-","EAF-","OAF-","AP-","BMS-",
    "HUM-","SP-","MSD-","CHWP-","HWP-","BT-","ERAC-","MSSB-","IFD-",
    "DAF-","EAD-","SAD-","NRD-","ORC-","TRC-","ERC-","CU-","CDU-",
    "RCU-","BS-","SS-","RL-","SL-","TL-","EL-","DPM-","SFD-","EFD-",
    "ACU-","PAU-","MAU-","ERV-","HRV-","RTU-","SAF-","RAF-","BC-",
    "DG-","MVCD-","IFD-","BC-",
]




# ── Content that is NEVER a QA comment ───────────────────
NOISE_CONTENT = {
    "f/a","t/a","f/b","s/a","r/a","e/a","o/a","c/w",
    "n/a","t.b.c","tbc","tbd","n.t.s","nts",
}



NOISE_PATTERNS = [
    r'^\d+[\d,]*\.?\d*\s*(mm|m|sq\.?m|l/s|ls|kg|kpa|pa|%)?$',
    r'^\d+[xX]\d+([xX]\d+)?$',
    r'^[øØoO]?\d+[øØ]?\s*[xX]?\s*\d*$',
    r'^([øØ]?\d+[øØ]?\s*){2,}$',
    r'^\(?\d+\.?\d*\s*[lL]/[sS]\)?$',
    r'^\d+\s*[lL]/[sS]\s+[A-Z]{2,5}-',
    r'^[A-Z]{2,6}-[A-Z0-9]+-[A-Z0-9]+-[A-Z0-9]+-\d+$',
    r'^([A-Z]{2,5}-[A-Z0-9]+-[A-Z0-9]+-\d+\s*){2,}$',
    r'^S\d+-[A-Z]-\d+(\s+\d+)?$',
    r'^\d+\s*-\s*[A-Z]$',
    r'^\d+[xX]\d+\s+\d+$',
    r'^[Cc]/[Ww]\s+\d+[xX]\d+\s+\d+[xX]\d+',
    r'^[A-Z]{1,2}\d*$',
    r'^P\d{2,3}$',
    r'^\d{1,2}/\d{2}/\d{2,4}$',
    r'^[A-Z]{3}\s*[\']\d{2}$',
    r'^\d{1,2}:\d{2}(:\d{2})?\s*[AP]M$',
    r'^[A-Z0-9]{5,}-[A-Z0-9]{3,}-[A-Z0-9]{3,}-',
    r'^\d{1,2}\s*/\s*\d{1,2}$',
    r'^@\d+%',
    r'^[A-Z]{1,3}\d+-[A-Z]{1,2}-\d+',
    r'^\d+\.\s+ALL\s+',
    r'^(Level\s*)?\d+$',
    r'^(GF|B\d+|L\d+|RF|UG|P\d)$',
    r'^[\d\s\.\,\/\(\)mmMM]+$',
]

LEGEND_NOISE = [
    "uninsulated metal ductwork","internally insulated ductwork",
    "externally insulated ductwork","insulated flexible ductwork",
    "uninsulated flexible ductwork","fire rated ductwork",
    "supply air ductwork","exhaust air ductwork","return air ductwork",
    "outside air ductwork","transfer air ductwork","toilet exhaust",
    "kitchen exhaust ductwork","ceiling access panel","duct access panel",
    "temperature sensor","fire damper","volume control damper",
    "acoustic flexible ductwork","natural ventilation",
    "digital touchscreen panel","wall mounted thermostat",
    "room controller","co2 sensor","all duct sizes are metal",
    "all ducts are flange joint","all louvres sizes are neck",
    "security clips to be documented","all openings contain ember",
    "all return air grilles to be complete","model used",
    "checked by auditor","for construction","preliminary issue",
    "cc3 issue","for information","project no","project number",
    "drawing no","drawing number","drawing title","revision",
    "description","issued by","client","contractor","builder",
    "key plan","links","services",
]

# ── QA keywords — must have at least one ─────────────────
QA_KEYWORDS = [
    # Action words
    "missing","wrong","check","confirm","add","remove","write",
    "change","fix","provide","show","move","replace","correct",
    "update","ensure","verify","blanked","spelling","incorrect",
    "error","not required","instead","coordinate","clash","overlap",
    "not able","not sufficient","non compliant","noncompliant",
    "fail","incomplete","hyphen not","not matching","never",
    "confirm if","make sure","double check","louvre required",
    "bulkhead required","penetration required","use","place",
    "put","refer","install","route","relocate","resize","flip",
    "align","avoid","section","view","3d","confirm","shift",
    "rotate","reduce","increase","lower","raise","extend",
    "hide","layout","tag","label","note","detail",
    # Question words — reviewer asking something
    "please","why","what","where","how","maybe","might","could",
    "would","want","is there","isn't","aren't","don't","doesn't",
    "shouldn't","can't","cannot","will not","won't","should not",
    "is not","are not","is this","isn't this","how will",
    "can we","can't these","is it","does this","why this",
    "which one","what is","where is","how many","are these",
    "is there","isn't there","should this","should these",
    # HVAC specific
    "tag missing","tag not","tag clash","grille tag","duct tag",
    "louvre tag","ap missing","access panel","duct ap","ceiling ap",
    "internal lining","lining missing","full height","pipe connection",
    "spigot","open ended","close ended","return air path","filter",
    "insulation","sensor missing","co sensor","mvcd symbol",
    "nrd direction","bou","bot","cl ","setdown","set down",
    "return air","supply air","exhaust air","outside air",
    "fire rated","fire rating","smoke","compartment","egress",
    "clash with","clashing with","pass through","run through",
    "minimum","maximum","required","gap required","distance",
    "interval","radius","fall towards","toward","minimum gap",
    "small letter","capital letter","capital 's'","small 'x'",
    "spelling mistake","spelling error","wrong spelling",
    "arrow missing","tag & arrow","ap & tag","symbol not correct",
    "make up","doorway","open doorway","makeup via",
    "incomplete","incomplete pipe","rotate","180 degree",
    "shift","shift bulkhead","shift it","move it",
    "hide layout","too sharp","sharp","smooth radius",
    "properly","draw flex","draw duct","draw pipe",
    "overlapping","text overlap","tag overlap","overlap with",
    "too close","min.","max.","off ",
]


def _get_title_block_rect(page):
    """
    Detect actual title block region by finding the revision table,
    drawing number, and standard title block fields.
    Returns the y-coordinate above which content is the drawing area.
    Uses 95% as absolute maximum to avoid cutting real QA comments.
    """
    ph    = page.rect.height
    pw    = page.rect.width
    tb_y  = ph  # default: no exclusion

    # Look for title block markers in page text
    blocks = page.get_text("blocks")
    for block in blocks:
        btext = block[4].lower().strip()
        by0   = block[1]
        # Standard title block markers
        if any(marker in btext for marker in [
            "drawing no","drawing number","project no",
            "checked by","issued by","for construction",
        ]):
            # Only treat as title block if in bottom 20% AND right 40%
            bx0 = block[0]
            if by0 > ph * 0.75 and bx0 > pw * 0.50:
                if by0 < tb_y:
                    tb_y = by0

    # Never exclude more than bottom 20% (80% cutoff)
    # This ensures annotations near/below the drawing area are still captured
    return max(tb_y, ph * 0.80)


def _is_noise(text):
    t  = text.strip()
    tl = t.lower().strip()

    if len(t) < 3:
        return True

    if tl in NOISE_CONTENT:
        return True

    if any(leg in tl for leg in LEGEND_NOISE):
        return True

    for pat in NOISE_PATTERNS:
        if re.match(pat, t, re.IGNORECASE):
            return True

    return False


def _is_qa_comment(text):
    """Must contain at least one QA keyword."""
    tl = text.lower()
    return any(kw in tl for kw in QA_KEYWORDS)


def _nearest_tag(cx, cy, words_pos, radius=600):
    best_tag, best_dist = "", radius
    for wx, wy, word in words_pos:
        wup = word.upper().rstrip(".,;)([]")
        if any(wup.startswith(p.upper()) for p in HVAC_PREFIXES):
            dist = ((wx-cx)**2 + (wy-cy)**2)**0.5
            if dist < best_dist:
                best_dist = dist
                best_tag  = word.rstrip(".,;)([]")
    return best_tag


def _location(cx, cy, pw, ph):
    nx, ny = cx/pw, cy/ph
    v = "Top"    if ny < 0.20 else \
        "Upper"  if ny < 0.40 else \
        "Middle" if ny < 0.60 else \
        "Lower"  if ny < 0.80 else "Bottom"
    h = "Left"      if nx < 0.20 else \
        "Left-Ctr"  if nx < 0.40 else \
        "Centre"    if nx < 0.60 else \
        "Right-Ctr" if nx < 0.80 else "Right"
    return f"{v}-{h}"


def _norm(text):
    return re.sub(r'\s+', ' ', text.lower().strip())[:70]


def _drawing_info(page):
    ph   = page.rect.height
    info = {"drawing_number": "", "revision": ""}
    full = ""
    for block in page.get_text("blocks"):
        if block[1] > ph * 0.80:
            full += " " + block[4].strip()
    m = re.search(r'[A-Z]{2,6}-[A-Z0-9]+-[A-Z0-9]+-[A-Z]+-[A-Z0-9]+', full)
    if m: info["drawing_number"] = m.group(0)
    m = re.search(r'\bRev\s*([A-Z0-9]{1,3})\b', full, re.I)
    if m: info["revision"] = m.group(1)
    return info


HIGH_KW   = [
    "missing","wrong","incorrect","error","must","critical","remove",
    "clash","clashing","not required","full height","cannot","not able",
    "non compliant","noncompliant","fail","incomplete","hyphen not",
    "not matching","louvre required","bulkhead required","penetration required",
    "symbol not correct","spelling mistake","spelling error","incomplete pipe",
    "too sharp","too close","don't allow",
]
MEDIUM_KW = [
    "check","confirm","provide","show","move","update","change","add",
    "ensure","verify","review","consider","please","should","recommend",
    "suggest","clarify","query","avoid","need to","why","better to",
    "confirm if","make sure","double check","rotate","shift","reduce",
    "increase","hide","layout","is there","can we","how will",
    "isn't","isn't this","which one","minimum","maximum",
]
CATEGORY_RULES = [
    (["tag missing","tag not","label missing","no tag","missing tag",
      "grille tag","duct tag","louvre tag","hyphen not","tag clash",
      "small letter","capital letter","small 'x'","capital 's'",
      "tag & arrow","ap & tag missing","tag missing","arrow missing",
      "symbol not correct","mvcd symbol","nrd direction","wrong nrd",
      "spelling mistake","spelling error","wrong spelling"],             "Tag Issue"),
    (["ap missing","access panel","ap clashing","duct ap","ap not",
      "ceiling ap","duct mounted ap","flip ap","600x600 ap",
      "not able to access","ap placed","ap for","add ap","add 1 ap",
      "ap & tag","bc"],                                                  "Access Panel Issue"),
    (["insulation missing","internal lining","lining missing",
      "not fire rated","insulation lines","dashed","insulated plenum",
      "insulation missing","insulated","fire rated","fire rating"],      "Insulation Issue"),
    (["clash","clashing","conflict","interference","full height",
      "opening to pass","wall opening","opening required","egress path",
      "fire egress","pass through","passing through","run through",
      "don't allow","clash with slab"],                                  "Clash Issue"),
    (["duct","flex","spigot","fitting","plenum","transition",
      "pipe","oval","setdown","set down","bou","bot","top of",
      "offset","route","reroute","reducer","radius","fall towards",
      "smooth radius","draw flex","draw duct","sharp","too sharp",
      "200 dia","200 high duct"],                                        "Duct Issue"),
    (["fcu","ahu","fan coil","equipment","coil","grille","louvre",
      "nrd","sensor","co sensor","sensor missing","co2","mvcd",
      "ifd","door grille","dg missing","make up","doorway",
      "rotate cu","shift cu"],                                           "Equipment Issue"),
    (["dimension","level","height","distance","cl ","centre line",
      "minimum","maximum","min.","max.","gap required","interval",
      "1.","2.","3.","not full height"],                                 "Dimension Issue"),
    (["fire","fd","fire damper","compartment","smoke","fire rated",
      "not fire rated","fire egress","egress"],                          "Fire Issue"),
    (["section","view","3d","detail","legend","note","label",
      "ga plan","room tag","drawing","hide layout","text overlap",
      "text overlapping","overlapping","incomplete drawing"],            "Drawing Issue"),
    (["riser","vertical","shaft","straight from","upper levels",
      "can't these risers"],                                             "Riser Issue"),
    (["schedule","model","type","product","not matching",
      "open ended plenum","close ended plenum","plenum box",
      "airflow shows","which one is","confirm which"],                   "Schedule Issue"),
    (["return air","r/a path","return air path","ra path"],              "Return Air Issue"),
    (["louvre required","bulkhead required","penetration required",
      "make up","open doorway","makeup via","natural ventilation",
      "filter missing","filter in o/a","filter in oa"],                 "Ventilation Issue"),
    (["incomplete pipe","pipe connection","spigot access",
      "add 1 ap for this spigot","spigot & ap"],                        "Pipe Issue"),
]


def _classify(text):
    if not text: return "General Issue", "Low"
    t = text.lower()
    sev = "Low"
    if any(k in t for k in HIGH_KW):    sev = "High"
    elif any(k in t for k in MEDIUM_KW): sev = "Medium"
    cat = "General Issue"
    for kws, category in CATEGORY_RULES:
        if any(k in t for k in kws): cat = category; break
    return cat, sev


def extract_qa_markups(pdf_bytes):
    """
    Extracts ALL real QA FreeText annotations.
    
    Key fix: uses content-based title block detection
    instead of fixed position — captures annotations
    at 88%, 95%, even 109% page height that are real QA comments.
    Only excludes text that is ACTUALLY in the title block
    (right side of page, contains standard TB field names).
    """
    doc     = fitz.open(stream=pdf_bytes, filetype="pdf")
    results = {}

    for page_num in range(len(doc)):
        page   = doc[page_num]
        pw, ph = page.rect.width, page.rect.height

        # Smart title block detection — content based, not position based
        tb_y = _get_title_block_rect(page)

        # Word index
        words_pos = []
        for w in page.get_text("words"):
            words_pos.append(((w[0]+w[2])/2, (w[1]+w[3])/2, w[4]))

        dwg_info = _drawing_info(page)
        markups  = []
        seen     = set()
        idx      = 1

        for annot in page.annots():
            atype = annot.type[1]

            # Only FreeText and Text (sticky note) annotations
            if atype not in ("FreeText", "Text"):
                continue

            info    = annot.info
            content = " ".join(info.get("content","").split()).strip()
            subject = " ".join(info.get("subject","").split()).strip()
            title   = info.get("title","").strip()
            text    = content or subject

            if not text or len(text) < 3:
                continue

            # Get position
            rect   = annot.rect
            cx, cy = (rect.x0+rect.x1)/2, (rect.y0+rect.y1)/2

            # Smart title block exclusion:
            # Only exclude if BOTH: below tb_y AND in right 40% of page
            # (title block is always bottom-right)
            in_tb = (cy > tb_y) and (cx > pw * 0.55)
            if in_tb:
                continue

            # Filter noise
            if _is_noise(text):
                continue

            # Must be a real QA comment
            if not _is_qa_comment(text):
                continue

            # Dedup
            norm = _norm(text)
            if norm in seen:
                continue
            # Truncated duplicate check
            if any(norm[:50] in existing for existing in seen):
                continue
            seen.add(norm)

            cat, sev = _classify(text)
            tag       = _nearest_tag(cx, cy, words_pos)
            loc       = _location(cx, cy, pw, ph)

            markups.append({
                "id":             idx,
                "page":           page_num + 1,
                "drawing_number": dwg_info["drawing_number"],
                "location":       loc,
                "comment":        text,
                "affected_tag":   tag,
                "category":       cat,
                "severity":       sev,
                "reviewer":       title,
            })
            idx += 1

        results[page_num + 1] = {
            "markups":      markups,
            "drawing_info": dwg_info,
            "total":        len(markups),
        }

    doc.close()
    return results





def get_stats(results):
    all_m = [m for r in results.values() for m in r["markups"]]
    return {
        "total":       len(all_m),
        "high":        sum(1 for m in all_m if m["severity"]=="High"),
        "medium":      sum(1 for m in all_m if m["severity"]=="Medium"),
        "low":         sum(1 for m in all_m if m["severity"]=="Low"),
        "sheets":      len([p for p,r in results.items() if r["total"]>0]),
        "by_cat":      Counter(m["category"] for m in all_m),
        "by_sheet":    {pg:r["total"] for pg,r in results.items() if r["total"]>0},
        "all_markups": all_m,
    }