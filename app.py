"""
Academy Receipt Generation System  ·  v2.0
- Staff name: dropdown (Omar, Menna, Mariem) + free input
- Notes moved to end
- Receipt = fill Word template → PDF via LibreOffice (fallback: reportlab)
- Deployment-ready (Streamlit Cloud via packages.txt + GitHub)
"""

import streamlit as st
import pandas as pd
import json
import os
import re
import subprocess
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape
import base64
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image as RLImage, HRFlowable,
)
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
# ── Constants ────────────────────────────────────────────────────────────────
RECEIPTS_DIR    = Path("receipts")
COUNTER_FILE    = Path("receipt_counter.json")
EXCEL_FILE      = Path("tracks.xlsx")
TEMPLATE_FILE   = Path("receipt_template.docx")
STAFF_LIST      = ["Fatma Khaled","Menna Hagag", "Mariem Hisham", "Malak Mahmoud", "Omar Mohamed", "Mohamed Hallawa", "Diana Adel", "Ahmed Fathy"]
ACADEMY_NAME    = "TechTrek"

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TechTrek · Receipt System",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: var(--background-color); }
[data-testid="stSidebar"]           { display: none; }
.section-card {
    background: var(--secondary-background-color, #ffffff);
    border-radius:12px; padding:22px 26px; margin-bottom:18px;
    box-shadow:0 2px 8px rgba(0,0,0,.07);
    border-left:4px solid #dc2626;
}
.section-title {
    font-size:.93rem; font-weight:700; color:#dc2626;
    text-transform:uppercase; letter-spacing:.06em; margin-bottom:14px;
}
.app-header {
    background:linear-gradient(135deg,#b91c1c 0%,#dc2626 100%);
    border-radius:14px; padding:22px 30px; margin-bottom:26px;
    display:flex; align-items:center; gap:18px;
}
.app-header h1 { color:#fff; font-size:1.7rem; margin:0; }
.app-header p  { color:#fecaca; margin:4px 0 0; font-size:.9rem; }
.summary-box {
    background:linear-gradient(135deg,#fef2f2,#fee2e2);
    border:1.5px solid #fca5a5; border-radius:10px; padding:18px 22px;
}
[data-theme="dark"] .summary-box {
    background:linear-gradient(135deg,#2d1515,#1c0f0f);
    border-color:#7f1d1d;
}
.s-row { display:flex; justify-content:space-between; padding:6px 0;
         border-bottom:1px dashed var(--text-color, #bfdbfe); font-size:.93rem; }
.s-row:last-child { border-bottom:none; }
.s-lbl { color:#dc2626; font-weight:600; }
.s-val { color: var(--text-color); }
.s-total { font-size:1.05rem; font-weight:700; color:#dc2626; }
div[data-testid="stButton"]>button {
    background:linear-gradient(135deg,#b91c1c,#dc2626); color:white;
    border:none; border-radius:10px; font-weight:700; font-size:1rem;
    padding:13px 34px; width:100%; transition:opacity .2s;
}
div[data-testid="stButton"]>button:hover { opacity:.88; }
div[data-testid="stDownloadButton"]>button {
    background:linear-gradient(135deg,#991b1b,#b91c1c); color:white;
    border:none; border-radius:10px; font-weight:700; padding:11px 26px; width:100%;
}
div[data-testid="stDownloadButton"]>button:hover { opacity:.88; }
[data-testid="metric-container"] {
    background: var(--secondary-background-color, #ffffff);
    border-radius:10px; border:1px solid var(--text-color, #e2e8f0);
    padding:12px 16px !important;
}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# Receipt counter  ·  TTR-YYYY-XXXX
# ════════════════════════════════════════════════════════════════════════════
def _load_counter():
    if COUNTER_FILE.exists():
        with open(COUNTER_FILE) as f:
            return json.load(f)
    return {"year": datetime.now().year, "seq": 0}

def _save_counter(data):
    with open(COUNTER_FILE, "w") as f:
        json.dump(data, f)

def next_receipt_number():
    data = _load_counter()
    yr = datetime.now().year
    if data["year"] != yr:
        data = {"year": yr, "seq": 0}
    data["seq"] += 1
    _save_counter(data)
    return f"TTR-{data['year']}-{data['seq']:04d}"


# ════════════════════════════════════════════════════════════════════════════
# Excel loader
# ════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load_tracks(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath)
    df.columns = df.columns.str.strip()
    rename = {}
    for col in df.columns:
        cl = col.lower()
        if "university" in cl:  rename[col] = "University"
        elif "track" in cl:     rename[col] = "Track Name"
        elif "hour" in cl:      rename[col] = "Hours"
        elif "price" in cl:     rename[col] = "Price"
    df.rename(columns=rename, inplace=True)
    missing = {"University", "Track Name", "Hours", "Price"} - set(df.columns)
    if missing:
        st.error(f"Missing columns in Excel: {missing}")
        st.stop()
    df["Price"] = pd.to_numeric(df["Price"], errors="coerce").fillna(0)
    df["Hours"] = pd.to_numeric(df["Hours"], errors="coerce").fillna(0)
    return df


# ════════════════════════════════════════════════════════════════════════════
# Template filler  ·  handles split {{ key }} across Word XML runs
# ════════════════════════════════════════════════════════════════════════════


def _fix_split_placeholders(xml: str) -> str:
    """Merge {{ key }} placeholders that Word split across multiple runs."""
    # Merge any {{ key }} span across multiple <w:t> runs. The opening braces
    # may follow ordinary text in the same run (for example, an Arabic label).
    def _merge(m):
        text_only = re.sub(r'<[^>]+>', '', m.group(0))
        key = text_only[2:-2].strip().replace(" ", "")
        return "{{ " + key + " }}" if key else m.group(0)

    xml = re.sub(r'\{\{(?:(?!\}\}).)*\}\}', _merge, xml, flags=re.DOTALL)
    return xml

def _substitute(xml: str, data: dict) -> str:
    """
    Replace all template placeholders with XML-safe values.
    Prevents DOCX corruption from &, <, >, quotes, etc.
    """

    for key, val in data.items():

        # Convert value safely for XML
        safe_val = escape(str(val))

        # Handle all spacing variations
        variants = [
            f"{{{{ {key} }}}}",
            f"{{{{{key}}}}}",
            f"{{{{ {key}}}}}",
            f"{{{{{key} }}}}",
        ]

        for variant in variants:
            xml = xml.replace(variant, safe_val)

    return xml

def _generate_pdf(data: dict) -> bytes:
    RED = "#dc2626"
    DARK = "#1e293b"

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        topMargin=20*mm, bottomMargin=15*mm,
        leftMargin=18*mm, rightMargin=18*mm,
    )

    styles = getSampleStyleSheet()
    s_title = ParagraphStyle("T", parent=styles["Heading1"], fontSize=22, textColor=RED,
                              spaceAfter=2*mm, alignment=TA_CENTER)
    s_sub = ParagraphStyle("S", parent=styles["Normal"], fontSize=10, textColor="#64748b",
                           alignment=TA_CENTER, spaceAfter=6*mm)
    s_label = ParagraphStyle("L", parent=styles["Normal"], fontSize=10, textColor=RED,
                              fontName="Helvetica-Bold")
    s_value = ParagraphStyle("V", parent=styles["Normal"], fontSize=10, textColor=DARK)
    s_section = ParagraphStyle("Sec", parent=styles["Heading2"], fontSize=12, textColor=RED,
                                fontName="Helvetica-Bold", spaceBefore=4*mm, spaceAfter=2*mm)
    s_note = ParagraphStyle("N", parent=styles["Normal"], fontSize=9, textColor=DARK,
                             leading=14)

    elements = []

    # ── Logo + Header ──
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        img = RLImage(logo_path, width=50*mm, height=18*mm)
    else:
        img = Spacer(1, 18*mm)

    header_data = [[img, Paragraph(f"<b>{ACADEMY_NAME}</b><br/><font size=8>{'Payment Receipt'}</font>", s_title)]]
    header_table = Table(header_data, colWidths=[50*mm, 110*mm])
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (1, 0), (1, 0), 6*mm),
        ("ALIGN", (1, 0), (1, 0), "LEFT"),
    ]))
    elements.append(header_table)
    elements.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor(RED)))
    elements.append(Spacer(1, 3*mm))

    # ── Receipt Info ──
    ri_data = [
        [Paragraph("Receipt No.", s_label), Paragraph(data["receipt_number"], s_value),
         Paragraph("Date", s_label), Paragraph(data["date"], s_value)],
    ]
    ri_table = Table(ri_data, colWidths=[28*mm, 55*mm, 20*mm, 55*mm])
    ri_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elements.append(ri_table)
    elements.append(Spacer(1, 5*mm))

    # ── Section: Student ──
    elements.append(Paragraph("Student Information", s_section))
    si_data = [
        [Paragraph("Name", s_label), Paragraph(data["student_name"], s_value),
         Paragraph("ID", s_label), Paragraph(data["student_id"] or "—", s_value)],
        [Paragraph("Phone", s_label), Paragraph(data["phone"], s_value),
         Paragraph("Gmail", s_label), Paragraph(data["student_gmail"] or "—", s_value)],
        [Paragraph("Faculty", s_label), Paragraph(data["faculty"] or "—", s_value),
         Paragraph("University", s_label), Paragraph(data["university"], s_value)],
        [Paragraph("Department", s_label), Paragraph(data["department"] or "—", s_value),
         Paragraph("", s_label), Paragraph("", s_value)],
    ]
    si_table = Table(si_data, colWidths=[24*mm, 50*mm, 24*mm, 50*mm])
    si_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2*mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2*mm),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#fef2f2")),
    ]))
    elements.append(si_table)
    elements.append(Spacer(1, 4*mm))

    # ── Section: Track & Payment ──
    elements.append(Paragraph("Track &amp; Payment", s_section))
    tp_data = [
        [Paragraph("Track", s_label), Paragraph(data["track_name"], s_value),
         Paragraph("Method", s_label), Paragraph(data["payment_method"], s_value)],
        [Paragraph("Credit Hours", s_label), Paragraph(data["credit_hours"], s_value),
         Paragraph("Required", s_label), Paragraph(data["required_amount"], s_value)],
        [Paragraph("Paid", s_label), Paragraph(data["paid_amount"], s_value),
         Paragraph("Remaining", s_label), Paragraph(f'<font color="{RED}"><b>{data["remaining_amount"]}</b></font>', s_value)],
        [Paragraph("Status", s_label), Paragraph(data["status"], s_value),
         Paragraph("", s_label), Paragraph("", s_value)],
    ]
    tp_table = Table(tp_data, colWidths=[24*mm, 50*mm, 24*mm, 50*mm])
    tp_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2*mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2*mm),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#fef2f2")),
    ]))
    elements.append(tp_table)
    elements.append(Spacer(1, 4*mm))

    # ── Notes ──
    if data.get("notes"):
        elements.append(Paragraph("Notes", s_section))
        elements.append(Paragraph(data["notes"], s_note))
        elements.append(Spacer(1, 3*mm))

    # ── Receiver ──
    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#e2e8f0")))
    elements.append(Spacer(1, 2*mm))
    r_data = [[Paragraph("Received by", s_label), Paragraph(data["receiver_name"], s_value)]]
    r_table = Table(r_data, colWidths=[28*mm, 120*mm])
    r_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    elements.append(r_table)

    # ── Footer ──
    elements.append(Spacer(1, 8*mm))
    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#e2e8f0")))
    elements.append(Spacer(1, 2*mm))
    elements.append(Paragraph(f"{ACADEMY_NAME} — Thank you for your trust!", ParagraphStyle(
        "F", parent=styles["Normal"], fontSize=8, textColor="#94a3b8", alignment=TA_CENTER)))

    doc.build(elements)
    return buf.getvalue()


def fill_template(template_path: str, data: dict) -> tuple[bytes, bytes]:
    file_contents: dict[str, bytes] = {}
    with zipfile.ZipFile(template_path) as z:
        for name in z.namelist():
            file_contents[name] = z.read(name)

    xml = file_contents["word/document.xml"].decode("utf-8")
    xml = _fix_split_placeholders(xml)
    xml = _substitute(xml, data)
    file_contents["word/document.xml"] = xml.encode("utf-8")

    with tempfile.TemporaryDirectory() as tmp:
        docx_path = os.path.join(tmp, "receipt.docx")
        with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, content in file_contents.items():
                zout.writestr(name, content)
        docx_bytes = open(docx_path, "rb").read()

        # Convert filled DOCX to PDF — try soffice first
        pdf_path = os.path.join(tmp, "receipt.pdf")
        conversion_errors = []
        for converter in ("soffice", "libreoffice"):
            try:
                conversion = subprocess.run(
                    [converter, "--headless", "--convert-to", "pdf",
                     "--outdir", tmp, docx_path],
                    capture_output=True, text=True, timeout=90,
                )
            except FileNotFoundError:
                conversion_errors.append(f"`{converter}` is not installed")
                continue
            except subprocess.TimeoutExpired:
                conversion_errors.append(f"`{converter}` timed out")
                continue

            if conversion.returncode == 0 and os.path.exists(pdf_path):
                pdf_bytes = open(pdf_path, "rb").read()
                return docx_bytes, pdf_bytes

            details = (conversion.stderr or conversion.stdout).strip()
            conversion_errors.append(
                f"`{converter}` exited with code {conversion.returncode}"
                + (f": {details}" if details else "")
            )

    raise RuntimeError(
        "Template-based PDF conversion failed. " + "; ".join(conversion_errors)
    )


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════
def main():
    # Header with logo
    logo_b64 = ""
    if Path("logo.png").exists():
        with open("logo.png", "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()

    logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height:58px;">' if logo_b64 else ""
    st.markdown(f"""
    <div class="app-header">
        {logo_html}
        <div>
            <h1>{ACADEMY_NAME}</h1>
            <p>Student Enrollment Receipt Generation System</p>
        </div>
    </div>""", unsafe_allow_html=True)

    # Guards
    if not TEMPLATE_FILE.exists():
        st.error("❌ `receipt_template.docx` not found. Place it in the app directory.")
        return
    if not EXCEL_FILE.exists():
        st.error("❌ `tracks.xlsx` not found.")
        with st.expander("📥 Generate sample tracks.xlsx"):
            if st.button("Create sample file"):
                pd.DataFrame({
                    "University": ["Cairo University", "Cairo University",
                                   "Ain Shams University", "Ain Shams University",
                                   "Alexandria University"],
                    "Track Name": ["Data Science", "Web Development",
                                   "Cybersecurity", "AI & Machine Learning",
                                   "Project Management"],
                    "Hours":  [120, 80, 100, 150, 60],
                    "Price":  [8500, 6000, 9000, 12000, 5000],
                }).to_excel(EXCEL_FILE, index=False)
                st.success("✅ Created! Rerun the app.")
        return

    df = load_tracks(str(EXCEL_FILE))

    # ── SECTION 1 · User Info ─────────────────────────────────────────────────
    st.markdown('<div class="section-card"><div class="section-title">👤 User Info</div>', unsafe_allow_html=True)
    staff_choice = st.selectbox(
        "Receiver / Staff Name",
        options=STAFF_LIST + ["Other…"],
        index=0,
    )
    if staff_choice == "Other…":
        receiver_name = st.text_input("Enter name manually", placeholder="Full name…")
    else:
        receiver_name = staff_choice
        st.caption(f"✅ Receiving: **{receiver_name}**")
    st.markdown("</div>", unsafe_allow_html=True)

    # ── SECTION 2 · University ────────────────────────────────────────────────
    st.markdown('<div class="section-card"><div class="section-title">🏫 University Selection</div>', unsafe_allow_html=True)
    universities = sorted(df["University"].dropna().unique().tolist())
    if not universities:
        st.error("No universities found.")
        st.markdown("</div>", unsafe_allow_html=True)
        return
    selected_university = st.selectbox(
        "Select University", ["— Please select —"] + universities, key="univ")
    st.markdown("</div>", unsafe_allow_html=True)

    if selected_university == "— Please select —":
        st.info("👆 Select a university to continue.")
        return

    # ── SECTION 3 · Track ─────────────────────────────────────────────────────
    st.markdown('<div class="section-card"><div class="section-title">📚 Track Selection</div>', unsafe_allow_html=True)
    uni_df = df[df["University"] == selected_university].reset_index(drop=True)
    if uni_df.empty:
        st.error(f"No tracks found for **{selected_university}**.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    selected_track = st.selectbox(
        "Select Track", ["— Please select —"] + uni_df["Track Name"].tolist(), key="track")

    if selected_track != "— Please select —":
        row = uni_df[uni_df["Track Name"] == selected_track].iloc[0]
        track_hours = int(row["Hours"])
        track_price = float(row["Price"])
        c1, c2, c3 = st.columns(3)
        c1.metric("Track",    selected_track)
        c2.metric("⏱ Hours",  f"{track_hours} hrs")
        c3.metric("💰 Price", f"EGP {track_price:,.0f}")

    st.markdown("</div>", unsafe_allow_html=True)

    if selected_track == "— Please select —":
        st.info("👆 Select a track to continue.")
        return

    # ── SECTION 4 · Student Info ──────────────────────────────────────────────
    with st.container():
        st.markdown('<div class="section-title">📋 Student Information</div>', unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            student_name = st.text_input("Student Full Name *", placeholder="e.g. Ahmed Mohamed")
        with c2:
            phone = st.text_input("Phone Number *", placeholder="e.g. 01012345678")

        c1, c2 = st.columns(2)
        with c1:
            student_id = st.text_input("Student ID *", placeholder="e.g. 20211234")
        with c2:
            student_gmail = st.text_input("Student Gmail", placeholder="e.g. student@gmail.com")

        c1, c2 = st.columns(2)
        with c1:
            faculty = st.selectbox(
                "Faculty (optional)",
                ["Engineering", "Computer Science", "Science (SIM)", "Business"],
            )
        with c2:
            department_choice = st.selectbox(
                "Department",
                ["— Please select —", "Data Science", "Cyber Security", "SIM", "Artificial Intelligence", "Other"],
            )

        department_other = ""
        if department_choice == "Other":
            department_other = st.text_input("Enter Department", placeholder="e.g. Software Engineering")
        department = department_other if department_choice == "Other" else department_choice

    # ── SECTION 5 · Payment ───────────────────────────────────────────────────
    st.markdown('<div class="section-card"><div class="section-title">💳 Payment Details</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        paid = st.number_input(
            "Paid Amount (EGP)",
            min_value=0.0, max_value=float(track_price),
            value=float(track_price), step=100.0, format="%.0f",
        )
    with c2:
        payment_method = st.selectbox(
            "Payment Method",
            ["Cash", "Bank Transfer", "Wallet"],
        )
    with c3:
        payment_status = st.selectbox("Payment Status", ["Deposit", "Fully Paid"])
    remaining = track_price - paid
    st.markdown("</div>", unsafe_allow_html=True)

    # ── SECTION 6 · Summary ───────────────────────────────────────────────────
    st.markdown('<div class="section-card"><div class="section-title">🧾 Receipt Summary</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="summary-box">
      <div class="s-row"><span class="s-lbl">University</span><span class="s-val">{selected_university}</span></div>
      <div class="s-row"><span class="s-lbl">Track</span><span class="s-val">{selected_track}</span></div>
      <div class="s-row"><span class="s-lbl">Hours</span><span class="s-val">{track_hours} hrs</span></div>
      <div class="s-row"><span class="s-lbl">Student</span><span class="s-val">{student_name or "—"}</span></div>
      <div class="s-row"><span class="s-lbl">Phone</span><span class="s-val">{phone or "—"}</span></div>
      <div class="s-row"><span class="s-lbl">Gmail</span><span class="s-val">{student_gmail or "—"}</span></div>
      <div class="s-row"><span class="s-lbl">Department</span><span class="s-val">{department if department != "— Please select —" else "—"}</span></div>
      <div class="s-row"><span class="s-lbl">Track Price</span><span class="s-val">EGP {track_price:,.0f}</span></div>
      <div class="s-row"><span class="s-lbl">Paid</span><span class="s-val">EGP {paid:,.0f}</span></div>
      <div class="s-row"><span class="s-lbl">Status</span><span class="s-val">{payment_status}</span></div>
      <div class="s-row s-total"><span class="s-lbl">Remaining</span>
        <span class="s-total">EGP {remaining:,.0f}</span></div>
      <div class="s-row"><span class="s-lbl">Receiver</span><span class="s-val">{receiver_name or "—"}</span></div>
    </div>""", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── SECTION 7 · Notes (end of form) ──────────────────────────────────────
    st.markdown('<div class="section-card"><div class="section-title">📝 Notes</div>', unsafe_allow_html=True)
    notes = st.text_area(
        "Additional Notes (optional)",
        placeholder="e.g. First instalment — remaining due next month…",
        height=90,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # ── SECTION 8 · Generate ──────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)

    if st.button("🖨️  Generate Receipt PDF"):
        errors = []
        if not student_name.strip():   errors.append("Student name is required.")
        if not student_id.strip():     errors.append("Student ID is required.")
        if not phone.strip():          errors.append("Phone number is required.")
        if not receiver_name.strip():  errors.append("Receiver / staff name is required.")

        if errors:
            for e in errors:
                st.error(f"❌ {e}")
        else:
            receipt_num  = next_receipt_number()
            receipt_date = datetime.now().strftime("%d/%m/%Y")

            template_data = {
                "receipt_number":   receipt_num,
                "date":             receipt_date,
                "student_name":     student_name.strip(),
                "student_id":       student_id.strip(),
                "phone":            phone.strip(),
                "student_gmail":    student_gmail.strip(),
                "Gmail":            student_gmail.strip(),
                "faculty":          faculty,
                "department":       "" if department == "— Please select —" else department.strip(),
                "university":       selected_university,
                "track_name":       selected_track,
                "credit_hours":     f"{track_hours} hrs",
                "required_amount":  f"EGP {track_price:,.0f}",
                "paid_amount":      f"EGP {paid:,.0f}",
                "remaining_amount": f"EGP {remaining:,.0f}",
                "payment_method":   payment_method,
                "status":           payment_status,
                "notes":            notes.strip(),
                "receiver_name":    receiver_name.strip(),
            }

            with st.spinner("Filling template and generating PDF…"):
                try:
                    docx_bytes, pdf_bytes = fill_template(
                        str(TEMPLATE_FILE), template_data
                    )
                except Exception as exc:
                    st.error(f"❌ PDF generation failed: {exc}")
                    st.stop()

            # Save copies
            RECEIPTS_DIR.mkdir(exist_ok=True)
            (RECEIPTS_DIR / f"{receipt_num}.docx").write_bytes(docx_bytes)
            (RECEIPTS_DIR / f"{receipt_num}.pdf").write_bytes(pdf_bytes)

            st.success(f"✅ Receipt **{receipt_num}** generated successfully!")

            safe_name = student_name.strip().replace(" ", "_")
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥  Download PDF",
                    data=pdf_bytes,
                    file_name=f"Receipt_{receipt_num}_{safe_name}.pdf",
                    mime="application/pdf",
                    key="dl_pdf",
                )
            with col2:
                st.download_button(
                    label="📄  Download Word (.docx)",
                    data=docx_bytes,
                    file_name=f"Receipt_{receipt_num}_{safe_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_docx",
                )

    st.markdown("</div>", unsafe_allow_html=True)

    # ── Receipt History ───────────────────────────────────────────────────────
    if RECEIPTS_DIR.exists():
        pdfs = sorted(RECEIPTS_DIR.glob("TTR-*.pdf"), reverse=True)
        if pdfs:
            with st.expander(f"📂 Receipt History  ({len(pdfs)} receipts)", expanded=False):
                for pf in pdfs[:30]:
                    c1, c2 = st.columns([4, 1])
                    c1.write(f"📄 `{pf.stem}`")
                    c2.download_button(
                        "⬇️", data=pf.read_bytes(),
                        file_name=pf.name, mime="application/pdf",
                        key=f"hist_{pf.stem}",
                    )


if __name__ == "__main__":
    main()
