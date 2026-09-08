import streamlit as st
import os
import base64
import pandas as pd
from datetime import datetime
from PIL import Image
import pymupdf as fitz
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
from collections import Counter
from io import BytesIO

st.set_page_config(
    page_title="DrawCheck",
    #page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── Brand colors ──────────────────────────────────────────
PRIMARY   = "#3D4BA0"
PRIMARY_D = "#2C3680"
PRIMARY_L = "#E8EAF6"
ACCENT    = "#5C6BC0"
WHITE     = "#FFFFFF"
GRAY_L    = "#F0F2F8"
GRAY_M    = "#E0E4F0"
TEXT_D    = "#1A1A2E"
TEXT_M    = "#4A4A6A"
RED       = "#D32F2F"
AMBER     = "#F57C00"
GREEN     = "#388E3C"



# ── CSS ──────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Poppins:wght@600;700&display=swap');

html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
    color: {TEXT_D};
}}

/* Hide default streamlit header */
#MainMenu {{visibility: hidden;}}
header {{visibility: hidden;}}
footer {{visibility: hidden;}}

/* App background */
.stApp {{
    background: {WHITE};
}}

/* ── Header bar ── */
.header-bar {{
    background: linear-gradient(135deg, {PRIMARY_D} 0%, {PRIMARY} 60%, {ACCENT} 100%);
    padding: 0;
    margin: -1rem -1rem 2rem -1rem;
    border-radius: 0 0 16px 16px;
    box-shadow: 0 4px 20px rgba(61,75,160,0.25);
}}
.header-inner {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 18px 36px;
}}
.header-logo img {{
    height: 52px;
    filter: brightness(0) invert(1);
}}
.header-title {{
    text-align: center;
    flex: 1;
    padding: 0 24px;
}}
.header-title h1 {{
    font-family: 'Poppins', sans-serif;
    font-size: 26px;
    font-weight: 700;
    color: {WHITE};
    margin: 0;
    letter-spacing: 0.5px;
}}
.header-title p {{
    font-size: 13px;
    color: rgba(255,255,255,0.80);
    margin: 3px 0 0 0;
    font-weight: 300;
}}
.header-badge {{
    background: rgba(255,255,255,0.15);
    border: 1px solid rgba(255,255,255,0.30);
    border-radius: 20px;
    padding: 6px 16px;
    color: {WHITE};
    font-size: 12px;
    font-weight: 500;
    white-space: nowrap;
}}

/* ── Upload zone ── */
.upload-zone {{
    border: 2px dashed {ACCENT};
    border-radius: 16px;
    background: {PRIMARY_L};
    padding: 40px 24px;
    text-align: center;
    margin: 12px 0;
    transition: all 0.2s;
}}
.upload-zone:hover {{
    border-color: {PRIMARY};
    background: {GRAY_M};
}}
.upload-icon {{
    font-size: 48px;
    margin-bottom: 12px;
}}
.upload-title {{
    font-family: 'Poppins', sans-serif;
    font-size: 18px;
    font-weight: 600;
    color: {PRIMARY};
    margin-bottom: 6px;
}}
.upload-sub {{
    font-size: 13px;
    color: {TEXT_M};
}}

/* ── Metric cards ── */
.metric-row {{
    display: flex;
    gap: 14px;
    margin: 16px 0;
    flex-wrap: wrap;
}}
.metric-card {{
    flex: 1;
    min-width: 120px;
    background: {WHITE};
    border: 1px solid {GRAY_M};
    border-radius: 14px;
    padding: 18px 16px;
    text-align: center;
    box-shadow: 0 2px 8px rgba(61,75,160,0.08);
    transition: transform 0.15s, box-shadow 0.15s;
}}
.metric-card:hover {{
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(61,75,160,0.14);
}}
.metric-card.blue  {{ border-top: 4px solid {PRIMARY}; }}
.metric-card.red   {{ border-top: 4px solid {RED}; }}
.metric-card.amber {{ border-top: 4px solid {AMBER}; }}
.metric-card.green {{ border-top: 4px solid {GREEN}; }}
.metric-card.gray  {{ border-top: 4px solid #9E9E9E; }}
.metric-val {{
    font-family: 'Poppins', sans-serif;
    font-size: 32px;
    font-weight: 700;
    color: {TEXT_D};
    line-height: 1;
    margin-bottom: 6px;
}}
.metric-card.red   .metric-val {{ color: {RED}; }}
.metric-card.amber .metric-val {{ color: {AMBER}; }}
.metric-card.green .metric-val {{ color: {GREEN}; }}
.metric-label {{
    font-size: 12px;
    color: {TEXT_M};
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}}

/* ── Section headings ── */
.section-heading {{
    font-family: 'Poppins', sans-serif;
    font-size: 16px;
    font-weight: 600;
    color: {PRIMARY};
    border-left: 4px solid {PRIMARY};
    padding-left: 12px;
    margin: 24px 0 12px 0;
}}

/* ── Sheet nav panel ── */
.sheet-panel {{
    background: {GRAY_L};
    border-radius: 14px;
    padding: 18px;
    border: 1px solid {GRAY_M};
}}
.sheet-panel h4 {{
    font-family: 'Poppins', sans-serif;
    color: {PRIMARY};
    font-size: 14px;
    font-weight: 600;
    margin: 0 0 12px 0;
}}

/* ── Severity badges in table ── */
.badge-high   {{ background:#FFEBEE; color:{RED};   padding:2px 10px; border-radius:10px; font-size:11px; font-weight:600; }}
.badge-medium {{ background:#FFF3E0; color:{AMBER}; padding:2px 10px; border-radius:10px; font-size:11px; font-weight:600; }}
.badge-low    {{ background:#E8F5E9; color:{GREEN};  padding:2px 10px; border-radius:10px; font-size:11px; font-weight:600; }}

/* ── Buttons ── */
.stButton > button {{
    background: linear-gradient(135deg, {PRIMARY} 0%, {ACCENT} 100%);
    color: {WHITE};
    border: none;
    border-radius: 10px;
    font-weight: 600;
    font-size: 14px;
    padding: 10px 20px;
    transition: all 0.2s;
    box-shadow: 0 3px 10px rgba(61,75,160,0.25);
}}
.stButton > button:hover {{
    transform: translateY(-1px);
    box-shadow: 0 6px 18px rgba(61,75,160,0.35);
}}
.stDownloadButton > button {{
    background: linear-gradient(135deg, {GREEN} 0%, #43A047 100%) !important;
    color: {WHITE} !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 15px !important;
    padding: 12px 24px !important;
    box-shadow: 0 4px 14px rgba(56,142,60,0.30) !important;
}}

/* ── Dataframe ── */
.stDataFrame {{ border-radius: 12px; overflow: hidden; border: 1px solid {GRAY_M}; }}

/* ── Divider ── */
hr {{ border: none; border-top: 1px solid {GRAY_M}; margin: 20px 0; }}

/* ── Info boxes ── */
.feature-card {{
    background: {GRAY_L};
    border-radius: 14px;
    padding: 20px;
    text-align: center;
    border: 1px solid {GRAY_M};
    height: 100%;
}}
.feature-icon {{ font-size: 32px; margin-bottom: 10px; }}
.feature-title {{
    font-family: 'Poppins', sans-serif;
    font-size: 14px;
    font-weight: 600;
    color: {PRIMARY};
    margin-bottom: 6px;
}}
.feature-desc {{ font-size: 12px; color: {TEXT_M}; line-height: 1.5; }}

/* ── Progress bar ── */
.stProgress > div > div > div {{ background: linear-gradient(90deg, {PRIMARY} 0%, {ACCENT} 100%); border-radius: 4px; }}

/* ── Slider ── */
.stSlider > div > div > div {{ color: {PRIMARY}; }}
</style>
""", unsafe_allow_html=True)



# ── Logo as base64 ────────────────────────────────────────
def img_to_b64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return ""

logo_b64 = img_to_b64("/home/sgdn/Downloads/hvac_issue_checker/logo_transparent.png")
if not logo_b64:
    logo_b64 = img_to_b64("logo_transparent.png")

logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height:50px"/>' if logo_b64 else \
            '<span style="color:white;font-weight:700;font-size:18px">SG Design Nepal</span>'

# ── Header ────────────────────────────────────────────────
st.markdown(f"""
<div class="header-bar">
  <div class="header-inner">
    <div class="header-logo">{logo_html}</div>
    <div class="header-title">
      <h1>DrawCheck AI</h1>
      <p>HVAC Drawing QA Annotation Extractor</p>
    </div>
    <div class="header-badge">v1.0</div>
  </div>
</div>
""", unsafe_allow_html=True)




# ── Core extraction logic ─────────────────────────────────
HVAC_PREFIXES = [
    "FCU-","AHU-","FD-","VCD-","MVCD-","EAF-","OAF-","AP-","BMS-",
    "HUM-","SP-","MSD-","CHWP-","HWP-","BT-","ERAC-","MSSB-","IFD-",
    "DAF-","EAD-","SAD-","NRD-","ORC-","TRC-","ERC-","CU-","CDU-",
    "RCU-","BS-","SS-","RL-","SL-","TL-","EL-","DPM-","SFD-","EFD-",
    "ACU-","PAU-","MAU-","ERV-","HRV-","RTU-",
]
ATYPE_LABEL = {
    "FreeText":"Comment","Text":"Sticky Note","Ink":"Freehand",
    "Polygon":"Cloud/Shape","Line":"Line/Arrow","PolyLine":"Polyline",
    "Square":"Rectangle","Circle":"Circle",
}
HIGH_KW   = ["missing","wrong","incorrect","error","must","critical","remove",
             "clash","not required","full height","cannot","not able",
             "non compliant","fail","incomplete","hyphen not","not matching"]
MEDIUM_KW = ["check","confirm","provide","show","move","update","change","add",
             "ensure","verify","review","consider","please","should","recommend",
             "suggest","clarify","query","avoid","need to","why","better to",
             "confirm if","make sure","double check"]
CATEGORY_RULES = [
    (["tag missing","label missing","no tag","missing tag","grille tag",
      "duct tag","louvre tag","hyphen not"],                                  "Tag Issue"),
    (["ap missing","access panel","ap clashing","duct ap","ap not",
      "ceiling ap","duct mounted ap","flip ap","600x600 ap"],                "Access Panel Issue"),
    (["insulation missing","internal lining","lining missing",
      "not fire rated","insulated","fire rated"],                             "Insulation Issue"),
    (["clash","clashing","conflict","interference","full height"],            "Clash Issue"),
    (["duct","flex","spigot","fitting","plenum","transition",
      "pipe","oval","setdown","set down","bou","bot","top of"],              "Duct Issue"),
    (["fcu","ahu","fan coil","equipment","coil","grille",
      "louvre","symbol incorrect"],                                           "Equipment Issue"),
    (["dimension","level","offset","height","distance",
      "cl ","centre line","0.",".m ","m "],                                  "Dimension Issue"),
    (["fire","fd","fire damper","compartment","smoke"],                       "Fire Issue"),
    (["section","view","3d","detail","drawing","legend",
      "note","label","text","spelling"],                                      "Drawing Issue"),
    (["riser","vertical","shaft"],                                            "Riser Issue"),
    (["schedule","model","type","product","not matching"],                    "Schedule Issue"),
]

def classify(text):
    if not text: return "General Issue","Low"
    t = text.lower()
    sev = "Low"
    if any(k in t for k in HIGH_KW):    sev = "High"
    elif any(k in t for k in MEDIUM_KW): sev = "Medium"
    cat = "General Issue"
    for kws, category in CATEGORY_RULES:
        if any(k in t for k in kws): cat = category; break
    return cat, sev

def nearest_tag(cx, cy, words_pos, radius=500):
    best_tag, best_dist = "", radius
    for wx, wy, word in words_pos:
        wup = word.upper().rstrip(".,;)")
        if any(wup.startswith(p.upper()) for p in HVAC_PREFIXES):
            dist = ((wx-cx)**2 + (wy-cy)**2)**0.5
            if dist < best_dist: best_dist, best_tag = dist, word.rstrip(".,;)")
    return best_tag

def location_label(cx, cy, pw, ph):
    nx, ny = cx/pw, cy/ph
    v = "Top" if ny<0.20 else "Upper" if ny<0.40 else "Middle" if ny<0.60 else "Lower" if ny<0.80 else "Bottom"
    h = "Left" if nx<0.20 else "Left-Ctr" if nx<0.40 else "Centre" if nx<0.60 else "Right-Ctr" if nx<0.80 else "Right"
    return f"{v}-{h}"

def normalize(text):
    return re.sub(r'\s+',' ', text.lower().strip())[:65]

def get_drawing_info(page):
    ph   = page.rect.height
    tb_y = ph * 0.87
    info = {"drawing_number":"", "revision":""}
    full = ""
    for block in page.get_text("blocks"):
        if block[1] >= tb_y: full += " " + block[4].strip()
    m = re.search(r'[A-Z]{2,6}-[A-Z0-9]+-[A-Z0-9]+-[A-Z]+-[A-Z0-9]+', full)
    if m: info["drawing_number"] = m.group(0)
    m = re.search(r'\bRev\s*([A-Z0-9]{1,3})\b', full, re.I)
    if m: info["revision"] = m.group(1)
    return info

def extract_all_markups(pdf_bytes):
    doc     = fitz.open(stream=pdf_bytes, filetype="pdf")
    results = {}

    for page_num in range(len(doc)):
        page   = doc[page_num]
        pw, ph = page.rect.width, page.rect.height
        tb_y   = ph * 0.87

        words_pos = []
        for w in page.get_text("words"):
            if (w[1]+w[3])/2 < tb_y:
                words_pos.append(((w[0]+w[2])/2, (w[1]+w[3])/2, w[4]))

        # FreeText spatial index for pairing with shapes
        ft_pool = {}
        for annot in page.annots():
            if annot.type[1] == "FreeText":
                ct = " ".join(annot.info.get("content","").split()).strip()
                if ct and len(ct) > 3:
                    r = annot.rect
                    ft_pool[id(annot)] = {
                        "cx": (r.x0+r.x1)/2, "cy": (r.y0+r.y1)/2, "content": ct
                    }

        drawing_info = get_drawing_info(page)
        markups = []
        seen    = set()
        idx     = 1

        for annot in page.annots():
            atype   = annot.type[1]
            info    = annot.info
            content = " ".join(info.get("content","").split()).strip()
            title   = info.get("title","").strip()
            subject = info.get("subject","").strip()
            rect    = annot.rect
            cx, cy  = (rect.x0+rect.x1)/2, (rect.y0+rect.y1)/2

            if cy > tb_y: continue

            ann_label = ATYPE_LABEL.get(atype, atype)
            comment   = ""

            if atype == "FreeText":
                comment = content
                if not comment: continue

            elif atype == "Text":
                comment = content or subject
                if not comment: continue

            elif atype in ("Polygon","PolyLine"):
                if content and len(content) > 3:
                    comment = content
                else:
                    bt, bd = "", 300
                    for ft in ft_pool.values():
                        dist = ((ft["cx"]-cx)**2+(ft["cy"]-cy)**2)**0.5
                        if dist < bd: bd=dist; bt=ft["content"]
                    comment = f"[Cloud near: {bt[:60]}]" if bt else f"[Cloud at {location_label(cx,cy,pw,ph)}]"

            elif atype == "Ink":
                if content and len(content) > 3:
                    comment = content
                else:
                    bt, bd = "", 250
                    for ft in ft_pool.values():
                        dist = ((ft["cx"]-cx)**2+(ft["cy"]-cy)**2)**0.5
                        if dist < bd: bd=dist; bt=ft["content"]
                    comment = f"[Circled: {bt[:60]}]" if bt else f"[Freehand mark at {location_label(cx,cy,pw,ph)}]"

            elif atype == "Line":
                if content and len(content) > 1:
                    comment = f"[Measurement: {content}]"
                else:
                    bt, bd = "", 200
                    for ft in ft_pool.values():
                        dist = ((ft["cx"]-cx)**2+(ft["cy"]-cy)**2)**0.5
                        if dist < bd: bd=dist; bt=ft["content"]
                    comment = f"[Arrow → {bt[:60]}]" if bt else None
                    if not comment: continue

            elif atype in ("Square","Circle"):
                if content and len(content) > 3:
                    comment = content
                else:
                    bt, bd = "", 200
                    for ft in ft_pool.values():
                        dist = ((ft["cx"]-cx)**2+(ft["cy"]-cy)**2)**0.5
                        if dist < bd: bd=dist; bt=ft["content"]
                    comment = f"[Box: {bt[:60]}]" if bt else None
                    if not comment: continue

            elif atype in ("Highlight","Underline","StrikeOut"):
                try:
                    hl_text = page.get_textbox(fitz.Rect(rect)).strip()
                    if hl_text: comment = f"[{ann_label}: '{hl_text[:60]}']"
                    elif content: comment = f"[{ann_label}: {content}]"
                    else: continue
                except: continue

            elif atype == "Stamp":
                continue  # Reviewer identity stamp — skip

            else:
                if content and len(content) > 3: comment = content
                else: continue

            if not comment: continue
            norm = normalize(comment)
            if norm in seen: continue
            seen.add(norm)

            cat, sev = classify(comment)
            tag       = nearest_tag(cx, cy, words_pos)
            loc       = location_label(cx, cy, pw, ph)

            markups.append({
                "id":             idx,
                "page":           page_num + 1,
                "drawing_number": drawing_info["drawing_number"],
                "location":       loc,
                "comment":        comment,
                "affected_tag":   tag,
                "category":       cat,
                "severity":       sev,
                "reviewer":       title,
                "type":           ann_label,
            })
            idx += 1

        results[page_num + 1] = {
            "markups":      markups,
            "drawing_info": drawing_info,
            "total":        len(markups),
        }

    doc.close()
    return results


# ── Excel builder ─────────────────────────────────────────
C_NAVY,C_BLUE,C_LTBLUE = "1F3864","3D4BA0","E8EAF6"
C_WHITE = "FFFFFF"
C_R_BG,C_R_FG = "FFEBEE","C62828"
C_Y_BG,C_Y_FG = "FFF3E0","E65100"
C_G_BG,C_G_FG = "E8F5E9","2E7D32"
C_ALT          = "F5F6FC"

def _f(h): return PatternFill("solid", fgColor=h)
def _ft(bold=False,color="000000",size=10):
    return Font(bold=bold,color=color,size=size,name="Calibri")
def _al(h="left",wrap=False):
    return Alignment(horizontal=h,vertical="center",wrap_text=wrap)
def _bd():
    s=Side(style="thin",color="C5CAE9")
    return Border(left=s,right=s,top=s,bottom=s)

def cell(ws,r,c,v,bold=False,fg="000000",bg=None,align="left",
         wrap=False,size=10,merge_to=None):
    if merge_to:
        ws.merge_cells(f"{get_column_letter(c)}{r}:{get_column_letter(merge_to)}{r}")
    x = ws.cell(row=r,column=c,value=v)
    x.font=_ft(bold,fg,size); x.alignment=_al(align,wrap); x.border=_bd()
    if bg: x.fill=_f(bg)
    return x

def sev_cell(ws,r,c,sev):
    s={"High":(C_R_BG,C_R_FG),"Medium":(C_Y_BG,C_Y_FG),"Low":(C_G_BG,C_G_FG)}
    bg,fg=s.get(sev,("FFFFFF","000000"))
    x=ws.cell(row=r,column=c,value=sev)
    x.font=_ft(bold=True,color=fg); x.fill=_f(bg)
    x.alignment=_al("center"); x.border=_bd()

def build_summary(wb,all_results,filename):
    ws=wb.active; ws.title="SUMMARY"
    for col,w in zip("ABCDEFGH",[6,28,44,10,10,10,14,12]):
        ws.column_dimensions[col].width=w

    ws.merge_cells("A1:H1")
    t=ws["A1"]; t.value="DRAWCHECK AI — COMPLETE QA ANNOTATION REPORT"
    t.font=Font(bold=True,color=C_WHITE,size=15,name="Calibri")
    t.fill=_f(C_NAVY); t.alignment=_al("center")
    ws.row_dimensions[1].height=32

    ws.merge_cells("A2:H2")
    d=ws["A2"]
    d.value=f"File: {filename}   |   Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}   |   SG Design Nepal"
    d.font=Font(color="3D4BA0",size=10,name="Calibri",bold=True)
    d.fill=_f(C_LTBLUE); d.alignment=_al("center")
    ws.row_dimensions[2].height=18

    all_m=[m for r in all_results.values() for m in r["markups"]]
    total=len(all_m)
    high=sum(1 for m in all_m if m["severity"]=="High")
    med=sum(1 for m in all_m if m["severity"]=="Medium")
    lo=sum(1 for m in all_m if m["severity"]=="Low")
    sheets=[p for p,r in all_results.items() if r["total"]>0]

    cell(ws,3,1,f"TOTAL:  {total}",bold=True,fg=C_WHITE,bg=C_BLUE,size=13,align="center",merge_to=2)
    cell(ws,3,3,f"HIGH:  {high}",  bold=True,fg=C_R_FG,bg=C_R_BG,size=12,align="center",merge_to=4)
    cell(ws,3,5,f"MEDIUM:  {med}", bold=True,fg=C_Y_FG,bg=C_Y_BG,size=12,align="center",merge_to=6)
    cell(ws,3,7,f"LOW:  {lo}",     bold=True,fg=C_G_FG,bg=C_G_BG,size=12,align="center")
    cell(ws,3,8,f"SHEETS: {len(sheets)}",bold=True,fg=C_WHITE,bg=C_BLUE,align="center",size=12)
    ws.row_dimensions[3].height=26

    cell(ws,5,1,"PER SHEET BREAKDOWN",bold=True,fg=C_WHITE,bg=C_NAVY,align="center",size=11,merge_to=8)
    ws.row_dimensions[5].height=20
    for ci,h in enumerate(["Sheet","Drawing Number","Total","High","Medium","Low","Reviewer","Tab"],1):
        cell(ws,6,ci,h,bold=True,fg="3D4BA0",bg=C_LTBLUE,align="center")
    ws.row_dimensions[6].height=18

    ri=7
    for pg in sorted(sheets):
        r=all_results[pg]; mkps=r["markups"]
        dwg=r["drawing_info"].get("drawing_number","") or f"Page {pg}"
        h_=sum(1 for m in mkps if m["severity"]=="High")
        me_=sum(1 for m in mkps if m["severity"]=="Medium")
        lo_=sum(1 for m in mkps if m["severity"]=="Low")
        rvr=mkps[0]["reviewer"] if mkps else ""
        alt=C_ALT if ri%2==0 else C_WHITE

        cell(ws,ri,1,pg,align="center",bg=alt,bold=True)
        cell(ws,ri,2,dwg[:38],bg=alt,size=9)
        cell(ws,ri,3,len(mkps),align="center",bg=alt,bold=True)
        for ci_,val,rbg,rfg in [(4,h_,C_R_BG,C_R_FG),(5,me_,C_Y_BG,C_Y_FG),(6,lo_,C_G_BG,C_G_FG)]:
            x=ws.cell(row=ri,column=ci_,value=val)
            x.font=_ft(bold=True,color=rfg); x.fill=_f(rbg if val>0 else alt)
            x.alignment=_al("center"); x.border=_bd()
        cell(ws,ri,7,rvr,bg=alt,size=9)
        cell(ws,ri,8,f"Sheet {pg}",bg=alt,size=9,align="center")
        ws.row_dimensions[ri].height=16
        ri+=1

    ri+=1
    cell(ws,ri,1,"CATEGORY BREAKDOWN",bold=True,fg=C_WHITE,bg=C_NAVY,align="center",size=11,merge_to=8)
    ws.row_dimensions[ri].height=20; ri+=1
    cell(ws,ri,1,"#",bold=True,fg="3D4BA0",bg=C_LTBLUE,align="center")
    cell(ws,ri,2,"Category",bold=True,fg="3D4BA0",bg=C_LTBLUE,merge_to=6)
    cell(ws,ri,7,"Count",bold=True,fg="3D4BA0",bg=C_LTBLUE,align="center",merge_to=8)
    ws.row_dimensions[ri].height=18; ri+=1
    for i,(cat,count) in enumerate(Counter(m["category"] for m in all_m).most_common(),1):
        alt=C_ALT if i%2==0 else C_WHITE
        cell(ws,ri,1,i,align="center",bg=alt)
        cell(ws,ri,2,cat,bg=alt,merge_to=6)
        cell(ws,ri,7,count,align="center",bg=alt,bold=True,merge_to=8)
        ws.row_dimensions[ri].height=16; ri+=1

    ws.freeze_panes="A4"

def build_master_tab(wb,all_results):
    ws=wb.create_sheet("ALL MARKUPS")
    for col,w in zip("ABCDEFGHI",[5,7,13,46,22,22,10,14,12]):
        ws.column_dimensions[col].width=w

    ws.merge_cells("A1:I1")
    t=ws["A1"]; t.value="ALL QA MARKUPS — COMPLETE DRAWING SET"
    t.font=Font(bold=True,color=C_WHITE,size=14,name="Calibri")
    t.fill=_f(C_NAVY); t.alignment=_al("center")
    ws.row_dimensions[1].height=28

    for ci,h in enumerate(["No.","Sheet","Location","QA Markup Comment","Affected Tag",
                            "Category","Sev.","Reviewer","Type"],1):
        cell(ws,2,ci,h,bold=True,fg="3D4BA0",bg=C_LTBLUE,align="center")
    ws.row_dimensions[2].height=18
    ws.freeze_panes="A3"

    gidx=1
    for pg in sorted(all_results.keys()):
        for m in all_results[pg]["markups"]:
            alt=C_ALT if gidx%2==0 else C_WHITE
            cell(ws,gidx+2,1,gidx,    align="center",bg=alt)
            cell(ws,gidx+2,2,m["page"],align="center",bg=alt,bold=True)
            cell(ws,gidx+2,3,m["location"],bg=alt,size=9)
            cell(ws,gidx+2,4,m["comment"],bg=alt,wrap=True,size=9)
            cell(ws,gidx+2,5,m["affected_tag"],align="center",bg=alt,size=9)
            cell(ws,gidx+2,6,m["category"],bg=alt,size=9)
            sev_cell(ws,gidx+2,7,m["severity"])
            cell(ws,gidx+2,8,m["reviewer"],bg=alt,size=9)
            cell(ws,gidx+2,9,m["type"],bg=alt,size=9,align="center")
            ws.row_dimensions[gidx+2].height=40 if len(m["comment"])>80 else 20
            gidx+=1

def build_page_tab(wb,sheet_name,markups,drawing_info):
    ws=wb.create_sheet(sheet_name)
    for col,w in zip("ABCDEFGH",[5,13,48,22,22,10,14,12]):
        ws.column_dimensions[col].width=w

    ws.merge_cells("A1:H1")
    t=ws["A1"]; t.value=f"QA MARKUP ANNOTATIONS — {sheet_name}"
    t.font=Font(bold=True,color=C_WHITE,size=13,name="Calibri")
    t.fill=_f(C_NAVY); t.alignment=_al("center")
    ws.row_dimensions[1].height=26

    dwg=drawing_info.get("drawing_number","") or sheet_name
    rev=drawing_info.get("revision","") or "—"
    h_=sum(1 for m in markups if m["severity"]=="High")
    me_=sum(1 for m in markups if m["severity"]=="Medium")
    lo_=sum(1 for m in markups if m["severity"]=="Low")

    ws.merge_cells("A2:C2")
    cell(ws,2,1,f"Drawing: {dwg}",bold=True,fg=C_WHITE,bg=C_BLUE)
    cell(ws,2,4,f"Rev: {rev}",bold=True,fg=C_WHITE,bg=C_BLUE,align="center")
    cell(ws,2,5,f"H:{h_}  M:{me_}  L:{lo_}",bold=True,fg=C_WHITE,bg=C_BLUE,align="center")
    ws.merge_cells("F2:G2")
    cell(ws,2,6,f"Total: {len(markups)}",bold=True,fg=C_WHITE,bg=C_BLUE,align="center")
    cell(ws,2,8,f"{datetime.now().strftime('%d/%m/%Y')}",fg=C_WHITE,bg=C_BLUE,align="center")
    ws.row_dimensions[2].height=18

    for ci,h in enumerate(["No.","Location","QA Markup Comment","Affected Tag",
                            "Category","Sev.","Reviewer","Type"],1):
        cell(ws,3,ci,h,bold=True,fg="3D4BA0",bg=C_LTBLUE,align="center")
    ws.row_dimensions[3].height=16
    ws.freeze_panes="A4"

    for ri,m in enumerate(markups,4):
        alt=C_ALT if ri%2==0 else C_WHITE
        cell(ws,ri,1,m["id"],align="center",bg=alt)
        cell(ws,ri,2,m["location"],bg=alt,size=9)
        cell(ws,ri,3,m["comment"],bg=alt,wrap=True,size=9)
        cell(ws,ri,4,m["affected_tag"],align="center",bg=alt,size=9)
        cell(ws,ri,5,m["category"],bg=alt,size=9)
        sev_cell(ws,ri,6,m["severity"])
        cell(ws,ri,7,m["reviewer"],bg=alt,size=9)
        cell(ws,ri,8,m["type"],bg=alt,size=9,align="center")
        ws.row_dimensions[ri].height=40 if len(m["comment"])>80 else 20

def build_excel(all_results,filename):
    wb=openpyxl.Workbook()
    build_summary(wb,all_results,filename)
    build_master_tab(wb,all_results)
    for pg in sorted(k for k,v in all_results.items() if v["total"]>0):
        r=all_results[pg]
        build_page_tab(wb,f"Sheet {pg}",r["markups"],r["drawing_info"])
    return wb


# ── UI ────────────────────────────────────────────────────
if not st.session_state.get("pdf_loaded"):
    # Landing page
    st.markdown("""
    <div class="section-heading">Upload QA-Marked Drawing PDF</div>
    """, unsafe_allow_html=True)

pdf_file = st.file_uploader(
    "Upload PDF",
    type=["pdf"],
    label_visibility="collapsed",
    help="Supports Bluebeam and Adobe Acrobat annotated PDFs"
)

if not pdf_file:
    st.markdown("""
    <div class="upload-zone">
        <div class="upload-icon">📄</div>
        <div class="upload-title">Drop your QA-marked drawing PDF here</div>
        <div class="upload-sub">Supports Bluebeam · Adobe Acrobat · Any annotated PDF</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    c1.markdown("""<div class="feature-card">
        <div class="feature-icon">🔍</div>
        <div class="feature-title">Extracts All Types</div>
        <div class="feature-desc">FreeText comments, Ink circles, Polygon clouds, Lines, Arrows — every annotation type</div>
    </div>""", unsafe_allow_html=True)
    c2.markdown("""<div class="feature-card">
        <div class="feature-icon">🏷️</div>
        <div class="feature-title">Auto Tag Detection</div>
        <div class="feature-desc">Automatically finds the nearest HVAC tag (FCU, FD, OAF etc.) for each markup</div>
    </div>""", unsafe_allow_html=True)
    c3.markdown("""<div class="feature-card">
        <div class="feature-icon">📊</div>
        <div class="feature-title">Professional Excel</div>
        <div class="feature-desc">Summary + Master sheet + Per-page tabs with severity color coding</div>
    </div>""", unsafe_allow_html=True)
    st.stop()

# Cache pages
if st.session_state.get("pdf_name") != pdf_file.name:
    with st.spinner("Loading pages..."):
        pdf_bytes = pdf_file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        pages = []
        for page in doc:
            mat = fitz.Matrix(100/72, 100/72)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            pages.append(img)
        doc.close()
        st.session_state.update({
            "pdf_name":  pdf_file.name,
            "pages":     pages,
            "pdf_bytes": pdf_bytes,
            "current":   1,
            "pdf_loaded": True,
        })

pages       = st.session_state["pages"]
total_pages = len(pages)

# Layout
col_img, col_side = st.columns([3, 1])

with col_side:
    st.markdown(f"""
    <div class="sheet-panel">
        <h4>📋 {total_pages} Sheet(s) Loaded</h4>
        <div style="font-size:12px;color:{TEXT_M}">{pdf_file.name}</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    if total_pages > 1:
        current = st.slider("Navigate Sheets", 1, total_pages,
                            st.session_state["current"], format="Sheet %d")
        st.session_state["current"] = current
        b1, b2 = st.columns(2)
        if b1.button("◀ Prev", use_container_width=True):
            st.session_state["current"] = max(1, current-1); st.rerun()
        if b2.button("Next ▶", use_container_width=True):
            st.session_state["current"] = min(total_pages, current+1); st.rerun()
    else:
        current = 1

    st.markdown("<br>", unsafe_allow_html=True)
    st.info("Extracts QA annotations from **all pages** simultaneously.")
    run_btn = st.button(
        "🔍  Extract All QA Annotations",
        use_container_width=True, type="primary"
    )

with col_img:
    st.markdown(f'<div class="section-heading">Drawing Preview — Sheet {current} of {total_pages}</div>',
                unsafe_allow_html=True)
    st.image(pages[current-1], use_container_width=True)

st.markdown("<hr>", unsafe_allow_html=True)

if run_btn:
    prog   = st.progress(0)
    status = st.empty()

    status.info("🔍 Scanning all pages for QA annotations...")
    all_results = extract_all_markups(st.session_state["pdf_bytes"])
    prog.progress(60)

    all_m     = [m for r in all_results.values() for m in r["markups"]]
    total     = len(all_m)
    pages_hit = [p for p,r in all_results.items() if r["total"]>0]

    high  = sum(1 for m in all_m if m["severity"]=="High")
    med   = sum(1 for m in all_m if m["severity"]=="Medium")
    lo    = sum(1 for m in all_m if m["severity"]=="Low")

    status.info("📊 Building professional Excel report...")
    wb  = build_excel(all_results, pdf_file.name)
    buf = BytesIO()
    wb.save(buf); buf.seek(0)
    prog.progress(100)
    status.success(f"✅ {total} QA markups extracted from {len(pages_hit)} sheet(s)!")

    # Metrics
    st.markdown('<div class="section-heading">Extraction Summary</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="metric-row">
        <div class="metric-card blue">
            <div class="metric-val">{total}</div>
            <div class="metric-label">Total Markups</div>
        </div>
        <div class="metric-card red">
            <div class="metric-val">{high}</div>
            <div class="metric-label">🔴 High</div>
        </div>
        <div class="metric-card amber">
            <div class="metric-val">{med}</div>
            <div class="metric-label">🟡 Medium</div>
        </div>
        <div class="metric-card green">
            <div class="metric-val">{lo}</div>
            <div class="metric-label">🟢 Low</div>
        </div>
        <div class="metric-card gray">
            <div class="metric-val">{len(pages_hit)}</div>
            <div class="metric-label">Sheets</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Per-page bar chart
    page_counts = {pg: all_results[pg]["total"] for pg in pages_hit}
    df_pg = pd.DataFrame({
        "Sheet":   [f"S{p}" for p in page_counts.keys()],
        "Markups": list(page_counts.values())
    })
    st.markdown('<div class="section-heading">Markups per Sheet</div>', unsafe_allow_html=True)
    st.bar_chart(df_pg.set_index("Sheet"))

    st.markdown("<hr>", unsafe_allow_html=True)

    # Full table with filters
    st.markdown('<div class="section-heading">All QA Annotations</div>', unsafe_allow_html=True)
    df = pd.DataFrame(all_m).rename(columns={
        "id":"No.","page":"Sheet","drawing_number":"Drawing No",
        "location":"Location","comment":"QA Comment",
        "affected_tag":"Tag","category":"Category",
        "severity":"Severity","reviewer":"Reviewer","type":"Ann. Type"
    })

    fc1,fc2,fc3 = st.columns(3)
    sev_f = fc1.multiselect("Severity",
                             ["High","Medium","Low"],
                             default=["High","Medium","Low"])
    pg_f  = fc2.multiselect("Sheet",
                              sorted(df["Sheet"].unique()),
                              default=sorted(df["Sheet"].unique()))
    cat_f = fc3.multiselect("Category",
                              sorted(df["Category"].unique()),
                              default=sorted(df["Category"].unique()))

    view = df[
        df["Severity"].isin(sev_f) &
        df["Sheet"].isin(pg_f) &
        df["Category"].isin(cat_f)
    ]
    st.caption(f"Showing **{len(view)}** of **{total}** markups")
    st.dataframe(
        view[["No.","Sheet","Location","QA Comment","Tag","Category","Severity","Reviewer","Ann. Type"]],
        use_container_width=True, hide_index=True,
        column_config={
            "No.":       st.column_config.NumberColumn(width="small"),
            "Sheet":     st.column_config.NumberColumn(width="small"),
            "Severity":  st.column_config.TextColumn(width="small"),
            "QA Comment":st.column_config.TextColumn(width="large"),
            "Tag":       st.column_config.TextColumn(width="medium"),
        }
    )

    st.markdown("<hr>", unsafe_allow_html=True)

    fn = f"DrawCheck_QA_{os.path.splitext(pdf_file.name)[0]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    st.download_button(
        "📥  Download Complete QA Report (.xlsx)",
        data      = buf.getvalue(),
        file_name = fn,
        mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type      = "primary"
    )




# Footer
st.markdown(f"""
<hr>
<div style="text-align:center;color:{TEXT_M};font-size:12px;padding:8px 0">
    DrawCheck AI v3.0 &nbsp;|&nbsp; SG Design Nepal &nbsp;|&nbsp;
    Extracts FreeText · Ink · Clouds · Lines · Shapes &nbsp;|&nbsp; No API required
</div>
""", unsafe_allow_html=True)



