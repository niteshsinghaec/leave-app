    # ==========================================
    # 2. GENERATE BILINGUAL PDF FILE (Full-Page Layout)
    # ==========================================
    font_path = "/tmp/NotoSansDevanagari.ttf"
    if not os.path.exists(font_path):
        try:
            url = "https://github.com/google/fonts/raw/true/ofl/notosansdevanagari/NotoSansDevanagari-Regular.ttf"
            urllib.request.urlretrieve(url, font_path)
        except Exception:
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
        topMargin=0.4*inch,
        bottomMargin=0.4*inch
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName=font_name, fontSize=12, alignment=1, spaceAfter=1)
    title_en_style = ParagraphStyle('TitleEnStyle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=11, alignment=1, spaceAfter=10)
    
    # Balanced leading and spacing to fill full single page naturally
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontName=font_name, fontSize=10, leading=14, spaceAfter=8)
    center_style = ParagraphStyle('CenterStyle', parent=styles['Normal'], fontName=font_name, fontSize=10, alignment=1, spaceAfter=8)
    right_style = ParagraphStyle('RightStyle', parent=styles['Normal'], fontName=font_name, fontSize=10, alignment=2, leading=13)

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

    pdf_elements.append(Spacer(1, 10))

    # Section 9
    t9_data = [
        [Paragraph(f"<b>9. स्थान / Place :</b> {place}<br/><b>तारीख / Date :</b> {a_date_str}", body_style),
         Paragraph("<b>आवेदक के हस्ताक्षर<br/>Signature of Applicant</b><br/><br/>", right_style)]
    ]
    t9_pdf = Table(t9_data, colWidths=[3.6*inch, 3.4*inch])
    t9_pdf.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    pdf_elements.append(t9_pdf)

    pdf_elements.append(Spacer(1, 4))
    pdf_elements.append(Paragraph("<b>सिफ़ारिश की जाती है / Recommended</b>", center_style))
    pdf_elements.append(Spacer(1, 4))

    # Section 10
    t10_data = [
        [Paragraph(f"<b>10. स्थान / Place :</b> {place}<br/><b>तारीख / Date :</b> {a_date_str}", body_style),
         Paragraph("<b>सिफ़ारिश करनेवाले प्राधिकारी के हस्ताक्षर तथा पदनाम<br/>Signature & Designation of Recommending Authority</b><br/><br/>", right_style)]
    ]
    t10_pdf = Table(t10_data, colWidths=[3.6*inch, 3.4*inch])
    t10_pdf.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    pdf_elements.append(t10_pdf)

    pdf_elements.append(Spacer(1, 6))

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
