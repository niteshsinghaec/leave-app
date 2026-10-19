import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="Leave App", page_icon="📄", layout="centered")

st.title("🛡️ Officers Leave App")
st.write("Fill in your details below and download your single-page formatted leave application in **Word** or **PDF** format.")

with st.form("leave_form"):
    rank_name = st.text_input("Rank and Name", "NITESH SINGH, AE(QA)")
    appointment = st.text_input("Appointment", "सहायक अभियंता (गु.आ) / Assistant Engineer (QA)")
    duty_arr = st.text_input("Duty Arrangement (leave blank if none)", "")
    rem_officers = st.text_input("Officers Remaining (leave blank if none)", "")
    
    leave_type = st.selectbox("Leave Type", ["Casual", "EL", "RH", "Furlough", "Sick", "Half Pay"])
    
    col1, col2 = st.columns(2)
    with col1:
        from_date = st.date_input("From Date")
    with col2:
        to_date = st.date_input("To Date")
        
    no_days = st.text_input("No. of Days", "01 day")
    address = st.text_input("Address during leave", "JAMSHEDPUR")
    cl_availed = st.text_input("Days CL Availed", "0 days")
    
    col3, col4 = st.columns(2)
    with col3:
        place = st.text_input("Place", "Mumbai")
    with col4:
        app_date = st.date_input("Application Date")

    submitted = st.form_submit_button("Generate Leave Application")

if submitted:
    f_date_str = from_date.strftime("%d-%m-%Y")
    t_date_str = to_date.strftime("%d-%m-%Y")
    a_date_str = app_date.strftime("%d.%m.%Y")
    
    duty_text = duty_arr if duty_arr else "...................................................."
    rem_text = rem_officers if rem_officers else "...................................................."

    # ==========================================
    # 1. GENERATE DOCX FILE
    # ==========================================
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0, 0, 0)

    def add_p(text, bold=False, space_after=8, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=11):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if text:
            run = p.add_run(text)
            run.bold = bold
            run.font.size = Pt(font_size)
        return p

    add_p("अधिकारियों की छुट्टी के लिए आवेदनपत्र", bold=True, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=13)
    add_p("OFFICERS APPLICATION FOR LEAVE", bold=True, space_after=14, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=13)

    add_p(f"1. संख्या/रैंक और नाम:\n    No. Rank and Name :- {rank_name}", bold=True, space_after=6)
    add_p("2. यूनिट :\n    Unit :- सामान्य भंडार खंड, व. गु. आ. स्था (सा. भं ), मुंबई -83\n              SQAE(GS) Mumbai-83", bold=True, space_after=6)
    add_p(f"3. नियुक्ति :\n    Appoitment :- {appointment}", bold=True, space_after=6)
    add_p(f"4. ड्यूटी के लिए प्रबन्ध :\n    Arrangement for performance of duty : - {duty_text}", bold=True, space_after=6)
    add_p(f"5. शेष अधिकारियों की संख्या : -\n    No. of Officers remaining :- {rem_text}", bold=True, space_after=6)
    
    add_p(f"6. Leave Details:\n    From : {f_date_str}   To : {t_date_str}   No. of Days : {no_days}\n    Type : Casual / EL / RH / Furlough / Sick / Half Pay [Selected: {leave_type}]", bold=True, space_after=6)

    add_p(f"7. छुट्टी के दौरान पता / Address during leave : {address}", bold=True, space_after=6)
    add_p(f"8. Days of Casual Leave availed :- {cl_availed}\n    Station leave permission may please be granted.", bold=True, space_after=10)

    t9 = doc.add_table(rows=1, cols=2)
    t9.alignment = WD_TABLE_ALIGNMENT.CENTER
    t9.autofit = False
    t9.columns[0].width = Inches(3.5)
    t9.columns[1].width = Inches(3.2)
    
    p9_l = t9.rows[0].cells[0].paragraphs[0]
    p9_l.paragraph_format.space_after = Pt(2)
    p9_l.add_run(f"9. Place : {place}    Date : {a_date_str}").bold = True

    p9_r = t9.rows[0].cells[1].paragraphs[0]
    p9_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p9_r.paragraph_format.space_after = Pt(2)
    p9_r.add_run("Signature of Applicant\n\n").bold = True

    add_p("Recommended / सिफ़ारिश की जाती है", bold=True, space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=11)

    t10 = doc.add_table(rows=1, cols=2)
    t10.alignment = WD_TABLE_ALIGNMENT.CENTER
    t10.autofit = False
    t10.columns[0].width = Inches(3.5)
    t10.columns[1].width = Inches(3.2)

    p10_l = t10.rows[0].cells[0].paragraphs[0]
    p10_l.paragraph_format.space_after = Pt(2)
    p10_l.add_run(f"10. Place : {place}    Date : {a_date_str}").bold = True

    p10_r = t10.rows[0].cells[1].paragraphs[0]
    p10_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p10_r.paragraph_format.space_after = Pt(2)
    p10_r.add_run("Signature & Designation of Recommending Authority\n\n").bold = True

    t11 = doc.add_table(rows=1, cols=2)
    t11.alignment = WD_TABLE_ALIGNMENT.CENTER
    t11.autofit = False
    t11.columns[0].width = Inches(3.2)
    t11.columns[1].width = Inches(3.5)

    p11_l = t11.rows[0].cells[0].paragraphs[0]
    p11_l.paragraph_format.space_after = Pt(2)
    p11_l.add_run(f"11. Place : {place}    Date : {a_date_str}").bold = True

    p11_r = t11.rows[0].cells[1].paragraphs[0]
    p11_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p11_r.paragraph_format.space_after = Pt(2)
    p11_r.add_run("Signature & Designation of Sanctioning Authority\nSanctioned / Not Sanctioned\n\n").bold = True

    docx_bio = io.BytesIO()
    doc.save(docx_bio)
    docx_bio.seek(0)

    # ==========================================
    # 2. GENERATE PDF FILE (Precise Sequential Layout)
    # ==========================================
    pdf_bio = io.BytesIO()
    pdf_doc = SimpleDocTemplate(
        pdf_bio,
        pagesize=A4,
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.4*inch,
        bottomMargin=0.4*inch
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=12, alignment=1, spaceAfter=2)
    title_en_style = ParagraphStyle('TitleEnStyle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=11, alignment=1, spaceAfter=10)
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.5, leading=14, spaceAfter=6)
    center_style = ParagraphStyle('CenterStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, alignment=1, spaceAfter=6)
    right_style = ParagraphStyle('RightStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.5, alignment=2, leading=13)

    pdf_elements = [
        Paragraph("OFFICERS APPLICATION FOR LEAVE", title_en_style),
        Paragraph(f"<b>1. No. Rank and Name :-</b> {rank_name}", body_style),
        Paragraph("<b>2. Unit :-</b> SQAE(GS) Mumbai-83 (Samanya Bhandar Khand)", body_style),
        Paragraph(f"<b>3. Appointment :-</b> {appointment}", body_style),
        Paragraph(f"<b>4. Arrangement for performance of duty :-</b> {duty_text}", body_style),
        Paragraph(f"<b>5. No. of Officers remaining :-</b> {rem_text}", body_style),
        Paragraph(f"<b>6. Leave Details:</b><br/>From : {f_date_str} &nbsp;&nbsp; To : {t_date_str} &nbsp;&nbsp; No. of Days : {no_days}<br/>Type : Casual / EL / RH / Furlough / Sick / Half Pay [Selected: {leave_type}]", body_style),
        Paragraph(f"<b>7. Address during leave :-</b> {address}", body_style),
        Paragraph(f"<b>8. Days of Casual Leave availed :-</b> {cl_availed}<br/>Station leave permission may please be granted.", body_style),
        Spacer(1, 6),
        Table([
            [Paragraph(f"<b>9. Place :</b> {place}<br/><b>Date :</b> {a_date_str}", body_style),
             Paragraph("<b>Signature of Applicant</b><br/><br/>", right_style)]
        ], colWidths=[3.6*inch, 3.4*inch], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]),
        Spacer(1, 6),
        Paragraph("<b>Recommended / सिफ़ारिश की जाती है</b>", center_style),
        Spacer(1, 6),
        Table([
            [Paragraph(f"<b>10. Place :</b> {place}<br/><b>Date :</b> {a_date_str}", body_style),
             Paragraph("<b>Signature & Designation of Recommending Authority</b><br/><br/>", right_style)]
        ], colWidths=[3.6*inch, 3.4*inch], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]),
        Spacer(1, 6),
        Table([
            [Paragraph(f"<b>11. Place :</b> {place}<br/><b>Date :</b> {a_date_str}", body_style),
             Paragraph("<b>Signature & Designation of Sanctioning Authority</b><br/>Sanctioned / Not Sanctioned<br/><br/>", right_style)]
        ], colWidths=[3.6*inch, 3.4*inch], style=[('VALIGN', (0,0), (-1,-1), 'TOP')])
    ]

    pdf_doc.build(pdf_elements)
    pdf_bio.seek(0)

    # ==========================================
    # DISPLAY DOWNLOAD BUTTONS
    # ==========================================
    st.success("✅ Leave application generated successfully!")
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label="📥 Download Word (.docx)",
            data=docx_bio,
            file_name="Filled_Leave_Application.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    with col_dl2:
        st.download_button(
            label="📄 Download PDF (.pdf)",
            data=pdf_bio,
            file_name="Filled_Leave_Application.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        
