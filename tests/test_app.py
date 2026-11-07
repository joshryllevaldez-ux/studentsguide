from io import BytesIO
import base64

from pptx import Presentation

from app import extract_pptx_text, make_pdf, parse_json_text


def make_pptx(*slide_texts):
    presentation = Presentation()
    for slide_text in slide_texts:
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        if slide_text:
            box = slide.shapes.add_textbox(100000, 100000, 5000000, 500000)
            box.text = slide_text
    output = BytesIO()
    presentation.save(output)
    return output.getvalue()


def test_extract_pptx_text_keeps_slide_numbers_and_topics():
    source, slide_count = extract_pptx_text(make_pptx(
        'Chapter 1' + chr(10) + 'Accounting equation and the basic structure of financial reporting for business organizations.',
        'Assets = Liabilities + Equity. This relationship explains how resources are financed by creditors and owners.',
    ))
    assert slide_count == 2
    assert 'SLIDE 1' in source and 'SLIDE 2' in source
    assert 'Accounting equation' in source
    assert 'Assets = Liabilities + Equity' in source


def test_extract_pptx_text_flags_slides_without_selectable_text():
    source, slide_count = extract_pptx_text(make_pptx(
        'This slide has enough selectable lecture content to pass the minimum text check.',
        '',
    ))
    assert slide_count == 2
    assert 'SLIDE 2' in source
    assert 'NO SELECTABLE TEXT FOUND' in source


def test_parse_json_text_accepts_plain_and_fenced_json():
    assert parse_json_text('{"ok": true}') == {'ok': True}
    assert parse_json_text('```json' + chr(10) + '{"count": 3}' + chr(10) + '```') == {'count': 3}


def test_make_pdf_returns_a_pdf_for_topic_content():
    pdf = make_pdf({
        'title': 'Accounting',
        'sections': [{
            'heading': 'The accounting equation',
            'slide_refs': [1],
            'explanation': 'Assets equal liabilities plus equity.',
            'brainhack': 'A = L + E',
            'example': {
                'title': 'Example',
                'scenario': 'A business has assets of 100.',
                'steps': ['Liabilities are 40.', 'Equity is 60.'],
                'takeaway': 'Both sides equal 100.',
            },
        }],
    })
    assert pdf.startswith(b'%PDF')
    assert len(pdf) > 1000


def test_make_pdf_handles_empty_sections():
    pdf = make_pdf({'title': 'Empty lecture', 'sections': []})
    assert pdf.startswith(b'%PDF')


def test_extract_pptx_images_returns_empty_for_text_only_slides():
    from app import extract_pptx_images
    assert extract_pptx_images(make_pptx(
        'This slide contains enough readable educational text for a valid presentation.'
    )) == []


def test_extract_pptx_images_collects_embedded_png():
    from app import extract_pptx_images
    png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/pWQAAAAASUVORK5CYII="
    )
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_picture(BytesIO(png), 0, 0)
    output = BytesIO()
    presentation.save(output)
    images = extract_pptx_images(output.getvalue())
    assert len(images) == 1
    assert images[0]["slide"] == "1"
    assert images[0]["mime_type"] == "image/png"
    assert base64.b64decode(images[0]["data"]) == png
