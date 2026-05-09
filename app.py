import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
import requests
from datetime import datetime

# PDF રિપોર્ટ માટેની લાઇબ્રેરીઓ
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4

# --- ૧. હેલ્પર ફંક્શન્સ ---
def fmt_date(val):
    if isinstance(val, datetime): return val.strftime("%d-%m-%Y")
    return ""

def is_filled(value):
    return value not in (None, "", "NA")

def format_percent(value):
    try:
        val = float(value)
        return f"{int(val * 100)}%" if val <= 1 else f"{int(val)}%"
    except: return str(value)

def center_cell(cell):
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.number_format = "General"

def check_remarks(percent_set, joint_count):
    percent_ints = []
    for p in percent_set:
        try: percent_ints.append(int(p.replace('%', '')))
        except: continue
    if 20 in percent_ints and joint_count == 5: return "FULL"
    if 10 in percent_ints and joint_count == 10: return "FULL"
    if 5 in percent_ints and joint_count == 20: return "FULL"
    return ""

def get_web_workbook(url):
    try:
        file_id = url.split("spreadsheets/d/")[1].split("/")[0]
        d_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
        response = requests.get(d_url)
        if response.status_code == 200:
            return load_workbook(filename=io.BytesIO(response.content), data_only=True)
    except: return None
    return None

# --- ૨. PDF જનરેશન ફંક્શન ---
def generate_status_pdf(completed, progress, report_title):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    for title, data, bg_color in [("COMPLETED LOTS", completed, colors.lightgreen), 
                                   ("UNDER PROGRESS LOTS", progress, colors.lightgrey)]:
        elements.append(Paragraph(f"{report_title}: {title}", styles["Heading1"]))
        elements.append(Spacer(1, 12))
        if len(data) > 1:
            t = Table(data, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), bg_color),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTSIZE', (0, 0), (-1, -1), 7),
            ]))
            elements.append(t)
        else:
            elements.append(Paragraph("કોઈ ડેટા મળ્યો નથી.", styles["Normal"]))
        elements.append(PageBreak())
    
    doc.build(elements)
    buffer.seek(0)
    return buffer

# --- ૩. લોગિન સિસ્ટમ ---
def check_password():
    if "password_correct" not in st.session_state:
        st.text_input("પાસવર્ડ નાખો", type="password", on_change=lambda: st.session_state.update({"password_correct": st.session_state.pwd == st.secrets["password"]}), key="pwd")
        return False
    return st.session_state["password_correct"]

if not check_password(): st.stop()

# --- ૪. પેજ સેટઅપ ---
st.set_page_config(page_title="MPC Quality Dashboard", layout="wide")
st.title("📊 MPC Quality Dashboard")

# છુપાવેલી લિંક્સ અને ટેમ્પલેટ
lhs_link = "https://docs.google.com/spreadsheets/d/1-ZQcx1OYmIQgijsQkPBdNPCX-LL1Y9pK/edit?usp=drive_link"
template_file_path = "LOT MASTER.xlsx"

tab1, tab2 = st.tabs(["🚀 LOT Status Reports", "📄 Individual LOT Generator"])

# --- ટેબ ૧: સ્ટેટસ રિપોર્ટ્સ (RT & DPT PDF) ---
with tab1:
    st.subheader("RT અને DPT લોટ સ્ટેટસ સમરી (PDF)")
    col_rt, col_dpt = st.columns(2)

    if col_rt.button("📊 Generate RT Status PDF"):
        with st.spinner("RT ડેટા પ્રોસેસ થઈ રહ્યો છે..."):
            wb = get_web_workbook(lhs_link)
            if wb:
                ws = wb["Sheet2"]
                lot_dict = {}
                for r in range(2, ws.max_row + 1):
                    lot = ws.cell(row=r, column=38).value # COL_AL
                    if lot in (None, "", "NA"): continue
                    if lot not in lot_dict: lot_dict[lot] = {"rows":0, "pct":set(), "rep":False, "ndt_ok":0}
                    lot_dict[lot]["rows"] += 1
                    if is_filled(ws.cell(row=r, column=33).value): lot_dict[lot]["pct"].add(format_percent(ws.cell(row=r, column=33).value))
                    if is_filled(ws.cell(row=r, column=34).value): lot_dict[lot]["rep"] = True
                    if is_filled(ws.cell(row=r, column=44).value): lot_dict[lot]["ndt_ok"] += 1
                
                comp, prog = [["LOT", "RT %", "JOINTS", "STATUS", "REMARKS", "NDT"]], [["LOT", "RT %", "JOINTS", "STATUS", "REMARKS", "NDT"]]
                for l in sorted(lot_dict.keys(), key=lambda x: int(x) if str(x).isdigit() else 0):
                    info = lot_dict[l]
                    stat = "CLOSED" if info["rep"] else "OPEN"
                    rem = check_remarks(info["pct"], info["rows"])
                    ndt_stat = "OK" if info["ndt_ok"] == info["rows"] else f"BAL({info['rows']-info['ndt_ok']})"
                    row = [str(l), ", ".join(info["pct"]), str(info["rows"]), stat, rem, ndt_stat]
                    if stat == "CLOSED" and rem == "FULL" and ndt_stat == "OK": comp.append(row)
                    else: prog.append(row)
                
                st.download_button("📥 Download RT PDF", generate_status_pdf(comp, prog, "RT LOT STATUS"), "RT_Status.pdf")

    if col_dpt.button("📊 Generate DPT Status PDF"):
        with st.spinner("DPT ડેટા પ્રોસેસ થઈ રહ્યો છે..."):
            wb = get_web_workbook(lhs_link)
            if wb:
                ws = wb["Sheet2"]
                lot_dict = {}
                for r in range(2, ws.max_row + 1):
                    j_type = str(ws.cell(row=r, column=10).value or "").upper()
                    lot = ws.cell(row=r, column=20).value if j_type == "EB" else ws.cell(row=r, column=32).value
                    if lot in (None, "", "NA"): continue
                    if lot not in lot_dict: lot_dict[lot] = {"rows":0, "pct":set(), "rep":False, "ndt_ok":0}
                    lot_dict[lot]["rows"] += 1
                    if j_type == "EB": lot_dict[lot]["pct"].add("10%")
                    elif is_filled(ws.cell(row=r, column=29).value): lot_dict[lot]["pct"].add(format_percent(ws.cell(row=r, column=29).value))
                    rep_col = 21 if j_type == "EB" else 30
                    if is_filled(ws.cell(row=r, column=rep_col).value): lot_dict[lot]["rep"] = True
                    if is_filled(ws.cell(row=r, column=44).value): lot_dict[lot]["ndt_ok"] += 1

                comp, prog = [["LOT", "DPT %", "JOINTS", "STATUS", "REMARKS", "NDT"]], [["LOT", "DPT %", "JOINTS", "STATUS", "REMARKS", "NDT"]]
                for l in sorted(lot_dict.keys(), key=lambda x: int(x) if str(x).isdigit() else 0):
                    info = lot_dict[l]
                    stat = "CLOSED" if info["rep"] else "OPEN"
                    rem = check_remarks(info["pct"], info["rows"])
                    ndt_stat = "OK" if info["ndt_ok"] == info["rows"] else f"BAL({info['rows']-info['ndt_ok']})"
                    row = [str(l), ", ".join(info["pct"]), str(info["rows"]), stat, rem, ndt_stat]
                    if stat == "CLOSED" and rem == "FULL" and ndt_stat == "OK": comp.append(row)
                    else: prog.append(row)
                
                st.download_button("📥 Download DPT PDF", generate_status_pdf(comp, prog, "DPT LOT STATUS"), "DPT_Status.pdf")

# --- ટેબ ૨: Individual LOT Generator (Excel) ---
with tab2:
    st.subheader("ઇન્ડિવિજ્યુઅલ LOT એક્સેલ જનરેટર")
    with st.sidebar:
        st.header("🎯 Generator Inputs")
        ndt_gen_type = st.radio("NDT પસંદ કરો", ["RT", "DPT"])
        lot_gen_no = st.text_input("લોટ નંબર નાખો (Excel માટે)")
        eb_gen_joint = st.radio("EB Joint છે?", ["N", "Y"]) if ndt_gen_type == "DPT" else "N"

    if st.button("🚀 Generate Individual Excel Report"):
        if not lot_gen_no:
            st.error("⚠️ લોટ નંબર નાખો.")
        else:
            with st.spinner("એક્સેલ તૈયાર થઈ રહી છે..."):
                wb = get_web_workbook(lhs_link)
                if wb:
                    ws = wb["Sheet2"]
                    rows = []
                    # ફિલ્ટરિંગ લોજિક
                    if ndt_gen_type == "RT":
                        for r in range(2, ws.max_row + 1):
                            if str(ws.cell(row=r, column=38).value).strip() == lot_gen_no: rows.append(r)
                        c_rep, c_date = 34, 35 # AH, AI
                    else:
                        search_col = 20 if eb_gen_joint == "Y" else 32
                        for r in range(2, ws.max_row + 1):
                            if str(ws.cell(row=r, column=search_col).value).strip() == lot_gen_no: rows.append(r)
                        c_rep, c_date = (21, 22) if eb_gen_joint == "Y" else (30, 31)

                    if not rows: st.warning("ડેટા મળ્યો નથી.")
                    else:
                        fmt_wb = load_workbook(template_file_path)
                        fmt_ws = fmt_wb.active
                        # હેડર ફિલિંગ (Example Logic)
                        fmt_ws["N4"].value = f"LOT NAME: MPL/AIPL/{ndt_gen_type}/LOT-{lot_gen_no}"
                        fmt_ws["G7"].value = lot_gen_no
                        
                        r_out, sr, l_spool = 7, 1, None
                        for r in rows:
                            fmt_ws[f"A{r_out}"].value = sr
                            fmt_ws[f"B{r_out}"].value = ws.cell(row=r, column=6).value # F
                            fmt_ws[f"F{r_out}"].value = f'{ws.cell(row=r, column=11).value}"/{ws.cell(row=r, column=16).value}'
                            fmt_ws[f"L{r_out}"].value = fmt_date(ws.cell(row=r, column=28).value) # AB
                            
                            # Report No logic
                            rep = str(ws.cell(row=r, column=c_rep).value or "")
                            xr = str(ws.cell(row=r, column=36).value or "") # AJ
                            fmt_ws[f"N{r_out}"].value = f"{rep} ({xr})" if rep and xr else rep
                            fmt_ws[f"O{r_out}"].value = fmt_date(ws.cell(row=r, column=c_date).value)

                            for col in "ABCDEFHIJKLMNOPQ": center_cell(fmt_ws[f"{col}{r_out}"])
                            r_out += 1; sr += 1
                        
                        out = io.BytesIO(); fmt_wb.save(out); out.seek(0)
                        st.download_button("📥 Download Excel", out, f"{ndt_gen_type}_LOT_{lot_gen_no}.xlsx")
