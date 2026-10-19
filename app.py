import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import os
import urllib.request

# Imports for PDF Generation with Font Support
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="Leave App", page_icon="📄", layout="centered")

st.title("🛡️ Officers Leave App")
st.write("Fill in your details below and download your single-page formatted leave application in **DOCX** or **PDF** (Bilingual Hindi & English).")

# Form inputs
with st.form("leave_form"):
    rank_name = st.text_input("Rank and Name", "श्री. नितेश सिंह, स. अ (गु . आ ), NITESH SINGH, AE(QA)")
    appointment = st.text_input("Appointment", "सहायक अभियंता (गु.आ)")
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
    
    add_p("6. वार्षिक / अर्जित / फरलों / बीमारी (Half Pay) / आकस्मिक / प्रतिबंधित / अर्धवेतन छुट्टी जिसके लिए आवेदन दिया गया है\n"
          f"    on : {f_date_str}   To {t_date_str}   No. of days : {no_days}\n"
          f"    तक के लिए (दिनों {no_days}) Casual / EL / RH / Furlough / Sick / Half Pay  [Selected: {leave_type}]", bold=True, space_after=6)

    add_p(f"7. छुट्टी के दौरान पता : {address}", bold=True, space_after=6)
    add_p(f"8. कितने दिन की आकस्मिक छुट्टी ली :\n    Days of Casual Leave availed :- {cl_availed}\n    Station leave permission may please be granted.", bold=True, space_after=10)

    # Section 9: Applicant Signature Table
    t9 = doc.add_table(rows=1, cols=2)
    t9.alignment = WD_TABLE_ALIGNMENT.CENTER
    t9.autofit = False
    t9.columns[0].width = Inches(3.5)
    t9.columns[1].width = Inches(3.2)
    
    p9_l = t9.rows[0].cells[0].paragraphs[0]
    p9_l.paragraph_format.space_after = Pt(2)
    p9_l.add_run(f"9. स्थान : {place}    तारीख\n    Place : {place}    Date : {a_date_str}").bold = True

    p9_r = t9.rows[0].cells[1].paragraphs[0]
    p9_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p9_r.paragraph_format.space_after = Pt(2)
    p9_r.add_run("आवेदक के हस्ताक्षर\nSignature of Applicant\n\n").bold = True

    add_p("सिफ़ारिश की जाती है", bold=True, space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=11)

    # Section 10: Recommending Authority Table
    t10 = doc.add_table(rows=1, cols=2)
    t10.alignment = WD_TABLE_ALIGNMENT.CENTER
    t10.autofit = False
    t10.columns[0].width = Inches(3.5)
    t10.columns[1].width = Inches(3.2)

    p10_l = t10.rows[0].cells[0].paragraphs[0]
    p10_l.paragraph_format.space_after = Pt(2)
    p10_l.add_run(f"10. स्थान : {place}            तारीख\n      Place : {place}    Date : {a_date_str}").bold = True

    p10_r = t10.rows[0].cells[1].paragraphs[0]
    p10_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p10_r.paragraph_format.space_after = Pt(2)
    p10_r.add_run("सिफ़ारिश करनेवाले प्राधिकारी के हस्ताक्षर तथा पदनाम\nSignature & Designation of Recommending Authority\n\n").bold = True

    # Section 11: Sanctioning Authority Table
    t11 = doc.add_table(rows=1, cols=2)
    t11.alignment = WD_TABLE_ALIGNMENT.CENTER
    t11.autofit = False
    t11.columns[0].width = Inches(3.2)
    t11.columns[1].width = Inches(3.5)

    p11_l = t11.rows[0].cells[0].paragraphs[0]
    p11_l.paragraph_format.space_after = Pt(2)
    p11_l.add_run(f"11. स्थान : {place}    तारीख:\n      Place : {place}    Date : {a_date_str}").bold = True

    p11_r = t11.rows[0].cells[1].paragraphs[0]
    p11_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p11_r.paragraph_format.space_after = Pt(2)
    p11_r.add_run("मंजूर करनेवाले प्राधिकारी के हस्ताक्षर तथा पदनाम\nSignature & Designation of Sanctioning Authority\nमंजूरी की जाती है / Sanctioned / Not Sanctioned\n\n").bold = True

    docx_bio = io.BytesIO()
    doc.save(docx_bio)
    docx_bio.seek(0)

    # ==========================================
    # 2. GENERATE BILINGUAL PDF FILE
    # ==========================================
    # Download Hindi-supported font automatically if not cached
    font_path = "/tmp/NotoSansDevanagari.ttf"
    if not os.path.exists(font_path):
        try:
            url = "https://github.com/google/fonts/raw/true/ofl/notosansdevanagari/NotoSansDevanagari-Regular.ttf"
            urllib.request.urlretrieve(url, font_path)
        except Exception:
            # Fallback direct link
            try:
                url_alt = "https://github.com/google/fonts/raw/main/ofl/notosansdevanagari/NotoSansDevanagari%5Bwdth%2Cwght%5D.ttf"
                urllib.request.urlretrieve(url_alt, font_path)
            except Exception:
                pass

    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont('HindiFont', font_path))
        font_name = 'HindiFont'
    else:
        font_name = 'Helvetica'

    pdf_bio = io.BytesIO()
    pdf_doc = SimpleDocTemplate(
        pdf_bio,
        pagesize=A4,
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName=font_name, fontSize=12, alignment=1, spaceAfter=2)
    title_en_style = ParagraphStyle('TitleEnStyle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=11, alignment=1, spaceAfter=8)
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontName=font_name, fontSize=9.5, leading=13, spaceAfter=4)
    center_style = ParagraphStyle('CenterStyle', parent=styles['Normal'], fontName=font_name, fontSize=9.5, alignment=1, spaceAfter=4)
    right_style = ParagraphStyle('RightStyle', parent=styles['Normal'], fontName=font_name, fontSize=9.5, alignment=2, leading=12)

    pdf_elements = []

    pdf_elements.append(Paragraph("अधिकारियों की छुट्टी के लिए आवेदनपत्र", title_style))
    pdf_elements.append(Paragraph("OFFICERS APPLICATION FOR LEAVE", title_en_style))

    pdf_elements.append(Paragraph(f"<b>1. संख्या/रैंक और नाम / No. Rank and Name :-</b> {rank_name}", body_style))
    pdf_elements.append(Paragraph("<b>2. यूनिट / Unit :-</b> सामान्य भंडार खंड / SQAE(GS) Mumbai-83", body_style))
    pdf_elements.append(Paragraph(f"<b>3. नियुक्ति / Appointment :-</b> {appointment}", body_style))
    pdf_elements.append(Paragraph(f"<b>4. ड्यूटी के लिए प्रबन्ध / Arrangement for performance of duty :-</b> {duty_text}", body_style))
    pdf_elements.append(Paragraph(f"<b>5. शेष अधिकारियों की संख्या / No. of Officers remaining :-</b> {rem_text}", body_style))
    
    p6_text = (f"<b>6. वार्षिक / अर्जित / फरलों / बीमारी / आकस्मिक छुट्टी के लिए आवेदन</b><br/>"
               f"&nbsp;&nbsp;&nbsp;&nbsp;From : {f_date_str} &nbsp;&nbsp; To : {t_date_str} &nbsp;&nbsp; No. of Days : {no_days}<br/>"
               f"&nbsp;&nbsp;&nbsp;&nbsp;Type : Casual / EL / RH / Furlough / Sick / Half Pay [Selected: {leave_type}]")
    pdf_elements.append(Paragraph(p6_text, body_style))

    pdf_elements.append(Paragraph(f"<b>7. छुट्टी के दौरान पता / Address during leave :-</b> {address}", body_style))
    pdf_elements.append(Paragraph(f"<b>8. कितने दिन की आकस्मिक छुट्टी ली / Days of Casual Leave availed :-</b> {cl_availed}<br/>&nbsp;&nbsp;&nbsp;&nbsp;Station leave permission may please be granted.", body_style))

    pdf_elements.append(Spacer(1, 4))

    # Section 9
    t9_data = [
        [Paragraph(f"<b>9. स्थान / Place :</b> {place}<br/><b>तारीख / Date :</b> {a_date_str}", body_style),
         Paragraph("<b>आवेदक के हस्ताक्षर<br/>Signature of Applicant</b><br/><br/>", right_style)]
    ]
    t9_pdf = Table(t9_data, colWidths=[3.6*inch, 3.4*inch])
    t9_pdf.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    pdf_elements.append(t9_pdf)

    pdf_elements.append(Paragraph("<b>सिफ़ारिश की जाती है / Recommended</b>", center_style))

    # Section 10
    t10_data = [
        [Paragraph(f"<b>10. स्थान / Place :</b> {place}<br/><b>तारीख / Date :</b> {a_date_str}", body_style),
         Paragraph("<b>सिफ़ारिश करनेवाले प्राधिकारी के हस्ताक्षर तथा पदनाम<br/>Signature & Designation of Recommending Authority</b><br/><br/>", right_style)]
    ]
    t10_pdf = Table(t10_data, colWidths=[3.6*inch, 3.4*inch])
    t10_pdf.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    pdf_elements.append(t10_pdf)

    # Section 11
    t11_data = [
        [Paragraph(f"<b>11. स्थान / Place :</b> {place}<br/><b>तारीख / Date :</b> {a_date_str}", body_style),
         Paragraph("<b>मंजूर करनेवाले प्राधिकारी के हस्ताक्षर तथा पदनाम<br/>Signature & Designation of Sanctioning Authority<br/>मंजूरी की जाती है / Sanctioned / Not Sanctioned</b><br/><br/>", right_style)]
    ]
    t11_pdf = Table(t11_data, colWidths=[3.6*inch, 3.4*inch])
    t11_pdf.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    pdf_elements.append(t11_pdf)

    pdf_doc.build(pdf_elements)
    pdf_bio.seek(0)

    # ==========================================
    # DISPLAY DOWNLOAD BUTTONS
    # ==========================================
    st.success("✅ Leave application generated successfully in bilingual DOCX & PDF!")
    
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
        
