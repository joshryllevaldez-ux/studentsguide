                    "Please generate it again."
                ),
            )
    if len(clean) < count:
        raise HTTPException(status_code=502, detail=f"The AI returned {len(clean)} usable questions instead of {count}. Please try generating again.")
    quiz["questions"] = clean[:count]
    quiz["mode"] = mode
    quiz["lecture_title"] = payload.lecture_title
    quiz["generation_type"] = "AI-generated from uploaded lecture"
    quiz["prompt_version"] = PROMPT_VERSION
    return quiz


def make_pdf(guide: dict[str, Any]) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, rightMargin=19*mm, leftMargin=19*mm,
        topMargin=18*mm, bottomMargin=18*mm,
        title=str(guide.get("title", "StudentsGUIDE Study Guide")),
        author="StudentsGUIDE.com",
    )
    navy = colors.HexColor("#142F69")
    blue = colors.HexColor("#087CEB")
    pale = colors.HexColor("#EAF3FF")
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SGTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=23, leading=28, textColor=navy, alignment=TA_LEFT, spaceAfter=8))
    styles.add(ParagraphStyle(name="SGSub", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=13, textColor=colors.HexColor("#5E6E86"), spaceAfter=10))
    styles.add(ParagraphStyle(name="SGH1", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=15, leading=19, textColor=navy, spaceBefore=14, spaceAfter=7, keepWithNext=True))
    styles.add(ParagraphStyle(name="SGH2", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=10.5, leading=14, textColor=blue, spaceBefore=7, spaceAfter=3, keepWithNext=True))
    styles.add(ParagraphStyle(name="SGBody", parent=styles["BodyText"], fontName="Helvetica", fontSize=9.3, leading=14, textColor=colors.HexColor("#263348"), spaceAfter=6))
    styles.add(ParagraphStyle(name="SGSmall", parent=styles["BodyText"], fontName="Helvetica", fontSize=8.5, leading=12, textColor=colors.HexColor("#34445D"), spaceAfter=4))

    def P(value: Any, style: str = "SGBody") -> Paragraph:
        return Paragraph(html.escape(str(value or "")).replace("\n", "<br/>"), styles[style])

    story: list[Any] = []
    logo_path = os.path.join(ASSET_DIR, "studentsguide-brand-pdf.jpg")
    if os.path.exists(logo_path):
        img = RLImage(logo_path, width=172*mm, height=57.3*mm)
        img.hAlign = "CENTER"
        story.extend([img, Spacer(1, 6)])
    story.append(P(guide.get("title", "Study Guide"), "SGTitle"))
    meta = f"{guide.get('course', 'Course not specified')}  •  {guide.get('source_filename', 'Uploaded lecture')}  •  {guide.get('slide_count', '')} slides"
    story.append(P(meta, "SGSub"))
    story.append(P("STUDY GUIDE OVERVIEW", "SGH2"))
    story.append(P(guide.get("overview", "A structured review of the uploaded lecture.")))

    objectives = guide.get("learning_objectives") or []
    if objectives:
        story.append(P("What you should understand", "SGH2"))
        for item in objectives:
            story.append(P("• " + str(item), "SGSmall"))

    for index, section in enumerate(guide.get("sections") or [], start=1):
        block: list[Any] = [P(f"{index}. {section.get('heading', 'Key topic')}", "SGH1")]
        if section.get("key_points"):
            block.append(P("Key points", "SGH2"))
            for point in section.get("key_points", []):
                block.append(P("• " + str(point), "SGSmall"))
        if section.get("explanation"):
            block.extend([P("Simplified explanation", "SGH2"), P(section["explanation"])])
        example = section.get("example") or {}
        if example:
            block.append(P(example.get("title", "Example"), "SGH2"))
            if example.get("scenario"):
                block.append(P(example["scenario"]))
            for step_no, step in enumerate(example.get("steps") or [], start=1):
                block.append(P(f"{step_no}. {step}", "SGSmall"))
            if example.get("takeaway"):
                block.append(P("Takeaway: " + str(example["takeaway"]), "SGSmall"))
        if section.get("brainhack"):
            hack = Table([[P("BRAIN HACK", "SGH2")], [P(section["brainhack"], "SGSmall")]], colWidths=[doc.width])
            hack.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,-1), pale), ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#C8DDF7")),
                ("LEFTPADDING", (0,0), (-1,-1), 9), ("RIGHTPADDING", (0,0), (-1,-1), 9),
                ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ]))
            block.extend([Spacer(1, 3), hack, Spacer(1, 5)])
        terms = section.get("important_terms") or []
        if terms:
            block.append(P("Important terms", "SGH2"))
            for term in terms:
                block.append(P(f"{term.get('term', '')}: {term.get('meaning', '')}", "SGSmall"))
        story.append(KeepTogether(block))

    for heading, key in [("Quick recap", "quick_recap"), ("Exam tips", "exam_tips")]:
        values = guide.get(key) or []
        if values:
            story.append(P(heading, "SGH1"))
            for value in values:
                story.append(P("• " + str(value), "SGSmall"))
    if guide.get("coverage_note"):
        story.append(P("Content coverage note", "SGH2"))
        story.append(P(guide["coverage_note"], "SGSmall"))
    story.append(Spacer(1, 10))
    story.append(P("Generated by StudentsGUIDE.com • Struggles less, study smart.", "SGSub"))

    def footer(canvas, document):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#D9E5F3"))
        canvas.line(19*mm, 14*mm, A4[0]-19*mm, 14*mm)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#71809A"))
        canvas.drawString(19*mm, 9*mm, "StudentsGUIDE.com")
        canvas.drawRightString(A4[0]-19*mm, 9*mm, f"Page {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buffer.getvalue()


@app.post("/api/study-guide-pdf")
async def study_guide_pdf(request: PdfRequest) -> StreamingResponse:
    content = make_pdf(request.guide)
    filename = re.sub(r"[^A-Za-z0-9_-]+", "_", str(request.guide.get("title", "StudentsGUIDE_Study_Guide"))).strip("_")[:70] or "StudentsGUIDE_Study_Guide"
    return StreamingResponse(
        io.BytesIO(content), media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}.pdf"'}
    )
