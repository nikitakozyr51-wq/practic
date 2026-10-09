"""Check the rendered PDF against Word headings, captions and page settings."""
from pathlib import Path
from hashlib import sha256
import json
import re
import sys

import fitz
from docx import Document

analysis = Path(__file__).resolve().parent / 'analysis'
pdf_path = Path(sys.argv[1]) if len(sys.argv) > 1 else analysis / 'report.pdf'
receipt = json.loads((analysis / 'report-build-receipt.json').read_text())
expected = json.loads((analysis / 'report-page-map.json').read_text())
pdf = fitz.open(pdf_path)
word = Document(analysis / 'report.docx')
compact = lambda text: re.sub(r'\W', '', text)
page_text = [page.get_text() for page in pdf]
normalized = [compact(text) for text in page_text]
actual = {}
for heading in receipt['headings']:
    matches = [i + 1 for i, text in enumerate(normalized)
               if i > 1 and compact(heading['text']) in text]
    assert matches, heading
    actual[str(heading['number'])] = matches[0]
for key, title in [('bibliography', 'Список использованной литературы'),
                   ('assignment', 'Приложение А. Задание на производственную практику')]:
    actual[key] = next(i + 1 for i, text in enumerate(normalized)
                       if i > 1 and compact(title) in text)
actual['total_pages'] = len(pdf)
assert actual == expected, (actual, expected)
section5_start=next(i for i,p in enumerate(receipt['paragraph_formats'])
                    if p['text']==receipt['headings'][4]['text'])
section6_start=next(i for i,p in enumerate(receipt['paragraph_formats'])
                    if p['text']==receipt['headings'][5]['text'])
section5_last=receipt['paragraph_formats'][section6_start-1]['text']
section5_last_page=next(i+1 for i,text in enumerate(normalized)
                       if compact(section5_last[-150:]) in text)

# The visible TOC must contain the correct numbers and black text.
toc = page_text[1]
for paragraph in word.paragraphs[16:35]:
    title, number = paragraph.text.rsplit('\t', 1)
    label = re.sub(r'\s+', '', title)
    pattern = r'\s*'.join(re.escape(char) for char in label)
    assert re.search(pattern + r'\s*\.{2,}\s*' + re.escape(number) + r'(?=\s|$)', toc), paragraph.text
for block in pdf[1].get_text('dict')['blocks']:
    for line in block.get('lines', []):
        assert all(span['color'] == 0 for span in line['spans'])

photos = []
center_deviations = []
section = word.sections[0]
content_center = (section.left_margin.pt + section.page_width.pt - section.right_margin.pt) / 2
for photo in receipt['photos']:
    caption = f"Рис. {photo['number']}. {photo['caption']}"
    matches = [i for i, text in enumerate(normalized) if compact(caption) in text]
    assert len(matches) == 1, (caption, matches)
    page = pdf[matches[0]]
    images = page.get_image_info()
    caption_rects = page.search_for(f"Рис. {photo['number']}.")
    assert len(caption_rects)==1,(caption,caption_rects)
    above=[fitz.Rect(item['bbox']) for item in images
           if item['bbox'][3] <= caption_rects[0].y0+1]
    assert above,(caption,images)
    bbox = max(above,key=lambda box:box.y1)
    assert page.rect.contains(bbox), (caption, bbox)
    deviation_mm = abs((bbox.x0 + bbox.x1) / 2 - content_center) * 25.4 / 72
    assert deviation_mm < 1, (caption, bbox)
    center_deviations.append(deviation_mm)
    assert caption_rects and caption_rects[0].y0 >= bbox.y1 - 1, caption
    assert caption_rects[0].y0-bbox.y1<35,(caption,bbox,caption_rects[0])
    photos.append({'number': photo['number'], 'page': matches[0] + 1})

assert not any('•' in text for text in page_text)
assert not any(name in '\n'.join(page_text) for name in ('Шеньшин', 'Максим', 'Еремян'))
for page in pdf:
    for image in page.get_image_info():
        assert page.rect.contains(fitz.Rect(image['bbox']))

result = {'pdf_pages': len(pdf), 'toc_page_numbers_verified_in_pdf_and_docx': True,
          'toc_text_black': True, 'photos': photos, 'photos_centered_in_text_area': True,
          'photo_center_tolerance_mm': 1,
          'photo_center_max_deviation_mm': round(max(center_deviations), 3),
          'no_images_outside_pages': True, 'no_unintended_bullet_markers': True,
          'section5_pages_including_partial_pages': [actual['5'], section5_last_page],
          'section5_page_count': section5_last_page-actual['5']+1,
          'assignment_pages': [actual['assignment'], len(pdf)],
          'font_substitution': 'Liberation Serif for Times New Roman',
          'pdf_sha256': sha256(pdf_path.read_bytes()).hexdigest()}
(analysis / 'render-verification.json').write_text(json.dumps(result, ensure_ascii=False, indent=2))
print(json.dumps(result, ensure_ascii=False))
