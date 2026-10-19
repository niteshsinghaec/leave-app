import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
from xhtml2pdf import pisa

st.set_page_config(page_title="Leave App", page_icon="📄", layout="centered")

st.title("🛡️ Officers Leave App")
st.write("Fill in your details below and download your single-page formatted leave application in **DOCX** or **PDF** (Exact Layout Match).")

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
    # 2. GENERATE PDF FILE (Exact Layout Match via HTML)
    # ==========================================
    html_content = f"""
    <html>
    <head>
    <style>
        @page {{
            size: A4;
            margin: 0.5in;
        }}
        body {{
            font-family: Helvetica, sans-serif;
            font-size: 10.5pt;
            color: #000;
            line-height: 1.3;
        }}
        .title-hi {{
            text-align: center;
            font-size: 13pt;
            font-weight: bold;
            margin-bottom: 2px;
        }}
        .title-en {{
            text-align: center;
            font-size: 11.5pt;
            font-weight: bold;
            margin-bottom: 15px;
        }}
        .field {{
            font-weight: bold;
            margin-bottom: 10px;
        }}
        .subtext {{
            margin-left: 20px;
            font-weight: bold;
        }}
        .center-heading {{
            text-align: center;
            font-weight: bold;
            margin: 12px 0;
        }}
        table {{
            width: 100%;
            margin-top: 8px;
            margin-bottom: 8px;
            border-collapse: collapse;
        }}
        td {{
            vertical-align: top;
            font-weight: bold;
            font-size: 10pt;
        }}
        .right-col {{
            text-align: right;
        }}
    </style>
    </head>
    <body>
        <div class="title-hi">अधिकारियों की छुट्टी के लिए आवेदनपत्र</div>
        <div class="title-en">OFFICERS APPLICATION FOR LEAVE</div>

        <div class="field">1. संख्या/रैंक और नाम:<br/>&nbsp;&nbsp;&nbsp;&nbsp;No. Rank and Name :- {rank_name}</div>
        <div class="field">2. यूनिट :<br/>&nbsp;&nbsp;&nbsp;&nbsp;Unit :- सामान्य भंडार खंड, व. गु. आ. स्था (सा. भं ), मुंबई -83<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SQAE(GS) Mumbai-83</div>
        <div class="field">3. नियुक्ति :<br/>&nbsp;&nbsp;&nbsp;&nbsp;Appointment :- {appointment}</div>
        <div class="field">4. ड्यूटी के लिए प्रबन्ध :<br/>&nbsp;&nbsp;&nbsp;&nbsp;Arrangement for performance of duty : - {duty_text}</div>
        <div class="field">5. शेष अधिकारियों की संख्या : -<br/>&nbsp;&nbsp;&nbsp;&nbsp;No. of Officers remaining :- {rem_text}</div>
        <div class="field">6. वार्षिक / अर्जित / फरलों / बीमारी (Half Pay) / आकस्मिक / प्रतिबंधित / अर्धवेतन छुट्टी जिसके लिए आवेदन दिया गया है<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;on : {f_date_str}   To {t_date_str}   No. of days : {no_days}<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;तक के लिए (दिनों {no_days}) Casual / EL / RH / Furlough / Sick / Half Pay [Selected: {leave_type}]</div>
        <div class="field">7. छुट्टी के दौरान पता : {address}</div>
        <div class="field">8. कितने दिन की आकस्मिक छुट्टी ली :<br/>&nbsp;&nbsp;&nbsp;&nbsp;Days of Casual Leave availed :- {cl_availed}<br/>&nbsp;&nbsp;&nbsp;&nbsp;Station leave permission may please be granted.</div>

        <table>
            <tr>
                <td>9. स्थान : {place} &nbsp;&nbsp;&nbsp;&nbsp;तारीख<br/>Place : {place} &nbsp;&nbsp;&nbsp;&nbsp;Date : {a_date_str}</td>
                <td class="right-col">आवेदक के हस्ताक्षर<br/>Signature of Applicant<br/><br/></td>
            </tr>
        </table>

        <div class="center-heading">सिफ़ारिश की जाती है</div>

        <table>
            <tr>
                <td>10. स्थान : {place} &nbsp;&nbsp;&nbsp;&nbsp;तारीख<br/>Place : {place} &nbsp;&nbsp;&nbsp;&nbsp;Date : {a_date_str}</td>
                <td class="right-col">सिफ़ारिश करनेवाले प्राधिकारी के हस्ताक्षर तथा पदनाम<br/>Signature & Designation of Recommending Authority<br/><br/></td>
            </tr>
        </table>

        <table>
            <tr>
                <td>11. स्थान : {place} &nbsp;&nbsp;&nbsp;&nbsp;तारीख:<br/>Place : {place} &nbsp;&nbsp;&nbsp;&nbsp;Date : {a_date_str}</td>
                <td class="right-col">मंजूर करनेवाले प्राधिकारी के हस्ताक्षर तथा पदनाम<br/>Signature & Designation of Sanctioning Authority<br/>मंजूरी की जाती है / Sanctioned / Not Sanctioned<br/><br/></td>
            </tr>
        </table>
    </body>
    </html>
    """

    pdf_bio = io.BytesIO()
    pisa_status = pisa.CreatePDF(io.BytesIO(html_content.encode("utf-8")), dest=pdf_bio)
    pdf_bio.seek(0)

    # ==========================================
    # DISPLAY DOWNLOAD BUTTONS
    # ==========================================
    if not pisa_status.err:
        st.success("✅ Leave application generated successfully in DOCX & PDF!")
    else:
        st.warning("⚠️ Generated with minor PDF warnings, but file is ready.")

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
        
