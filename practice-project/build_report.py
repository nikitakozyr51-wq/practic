"""Build the report using Nazar's original OOXML formatting and Maxim's photos."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
import re, json, hashlib
from datetime import datetime, timezone
from lxml import etree
from docx.oxml.shape import CT_Inline
from PIL import Image

ROOT = Path(__file__).resolve().parent
ANALYSIS = ROOT / 'analysis'
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
      'v': 'urn:schemas-microsoft-com:vml',
      'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
W = '{' + NS['w'] + '}'
EMU_CM = 360000
SOURCE = ROOT / 'originals' / 'Nazar_Otchet_Praktika.docx'
with ZipFile(SOURCE) as z:
    package = {n: z.read(n) for n in z.namelist()}
tree = etree.fromstring(package['word/document.xml'])
body = tree.find(W + 'body')
original_paragraphs = body.findall(W + 'p')
templates = [deepcopy(p) for p in original_paragraphs]
rels = etree.fromstring(package['word/_rels/document.xml.rels'])

def text_of(p):
    return ''.join(p.xpath('.//w:t/text()', namespaces=NS))

def replace_span(p, old, new):
    nodes = p.xpath('.//w:t', namespaces=NS)
    full = ''.join(n.text or '' for n in nodes)
    pos = full.find(old)
    if pos < 0:
        raise ValueError('Missing replacement: ' + old)
    end = pos + len(old)
    cursor = 0
    inserted = False
    for n in nodes:
        s = n.text or ''
        start, stop = cursor, cursor + len(s)
        if stop > pos and start < end:
            left = s[:max(0, pos - start)]
            right = s[max(0, end - start):] if end < stop else ''
            n.text = left + (new if not inserted else '') + right
            inserted = True
        cursor = stop

replace_span(original_paragraphs[6], 'm/v «PANTHERLIGHT»', 'Военный корабль «Чабаненко»')
replace_span(original_paragraphs[7], 'VALLETTA', 'Североморск')
replace_span(original_paragraphs[12], '03 Мая – 23 Августа 2026 г', '03 июня – 08 июля 2026 г')
for p in body.xpath('.//w:p', namespaces=NS):
    txt = text_of(p)
    if 'Еремяна Назара Ашотовича' in txt:
        replace_span(p, 'Еремяна Назара Ашотовича', 'Козырева Никиты Алексеевича')
    if 'Еремян Н. А.' in txt:
        replace_span(p, 'Еремян Н. А.', 'Козырев Н.А.')

# Keep the complete title page, TOC layout, section, logos and footer parts.
start = list(body).index(original_paragraphs[46])
for child in list(body)[start:]:
    if child.tag != W + 'sectPr':
        body.remove(child)
sect = body.find(W + 'sectPr')

def make_p(template_index, text, drop_numbering=False):
    template = templates[template_index]
    p = deepcopy(template)
    ppr = p.find(W + 'pPr')
    for child in list(p):
        if child is not ppr:
            p.remove(child)
    if drop_numbering and ppr is not None:
        for el in ppr.findall(W + 'numPr'):
            ppr.remove(el)
    candidates = template.xpath('./w:r[w:t]/w:rPr', namespaces=NS)
    r = etree.SubElement(p, W + 'r')
    if candidates:
        r.append(deepcopy(candidates[0]))
    t = etree.SubElement(r, W + 't')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    return p

def append(p):
    body.insert(list(body).index(sect), p)

format_receipts=[]
def format_report_p(p,role):
    """Use the sample's dominant text settings with consistent paragraph layout."""
    ppr=p.find(W+'pPr')
    if ppr is None:
        ppr=etree.Element(W+'pPr');p.insert(0,ppr)
    if role=='body' or (role=='heading' and re.match(r'^\d+\.',text_of(p))):
        # Some source paragraphs are bullet-list items. New continuous prose
        # must not accidentally inherit a bullet before every paragraph.
        for item in ppr.findall(W+'numPr'):ppr.remove(item)
    def set_prop(tag,**attrs):
        item=ppr.find(W+tag)
        if item is None:
            item=etree.SubElement(ppr,W+tag)
        for key,value in attrs.items():item.set(W+key,str(value))
        return item
    ind=set_prop('ind',left=0,right=0,firstLine=709 if role in ('body','heading') else 0)
    ind.attrib.pop(W+'hanging',None)
    set_prop('jc',val='both' if role in ('body','heading') else ('left' if role=='reference' else 'center'))
    set_prop('spacing',before=0,after=120 if role=='reference' else 0,line=240,lineRule='auto')
    set_prop('widowControl',val=1)
    if role in ('heading','appendix','bibliography','image'):
        set_prop('keepNext',val=1)
    if role in ('heading','appendix','bibliography','caption'):
        set_prop('keepLines',val=1)
    for run in p.findall(W+'r'):
        if not run.findall(W+'t'):
            continue
        rpr=run.find(W+'rPr')
        if rpr is None:
            rpr=etree.Element(W+'rPr');run.insert(0,rpr)
        for tag,attrs in [('rFonts',{'ascii':'Times New Roman','hAnsi':'Times New Roman','cs':'Times New Roman'}),
                          ('sz',{'val':'24'}),('szCs',{'val':'24'}),('color',{'val':'000000'})]:
            item=rpr.find(W+tag)
            if item is None:item=etree.SubElement(rpr,W+tag)
            for key,value in attrs.items():item.set(W+key,value)
    format_receipts.append({'role':role,'text':text_of(p)})
    return p

def clean_heading(label):
    label=label.strip().replace('оценкарайона','оценка района')
    return re.sub(r'^(\d+(?:\.\d+)*\.?)(?=[А-ЯЁA-Z])',r'\1 ',label)

main_templates = {1:46, 2:61, 3:72, 4:110, 5:124, 6:199, 7:220,
                  8:255, 9:348, 10:354, 11:374, 12:379, 13:390,
                  14:396, 15:406, 16:416, 17:418}
sub_templates = {'1.1':47,'1.2':52,'1.3':57,'2.1':62,'2.2':66,'2.3':68,
                 '2.4':70,'3.1':73,'3.2':89,'3.3':97,'4.1':111,'4.2':115,
                 '4.3':119,'4.4':123,'5.1':125,'5.2':141,'5.3':155,
                 '5.4':160,'5.5':189,'6.1':200,'6.2':202,'6.3':204,
                 '7.1':221,'7.2':240,'7.3':244,'7.4':253,'8.1':256,
                 '8.2':260,'8.3':272,'8.4':290,'8.5':297,'8.6':319,'8.7':336}
body_templates = {1:48,2:63,3:75,4:112,5:126,6:201,7:222,8:257,
                  9:349,10:355,11:375,12:380,13:391,14:397,15:407,
                  16:417,17:419}

# Remove old cached TOC fields; retain its paragraphs and tab-leader styles.
front_paragraphs = body.findall(W + 'p')
for p in front_paragraphs:
    for node in p.xpath('.//w:fldChar|.//w:instrText', namespaces=NS):
        node.getparent().remove(node)
toc_nodes = front_paragraphs[16:33]
def static_toc_props(original_props):
    props=deepcopy(original_props)
    for item in props.findall(W+'rStyle'):
        props.remove(item)
    # TOC fields render their links black in the source. Static text must carry
    # that appearance explicitly rather than inheriting blue hyperlink styling.
    for tag,val in [('color','000000'),('u','none')]:
        item=props.find(W+tag)
        if item is None:
            item=etree.SubElement(props,W+tag)
        item.set(W+'val',val)
    return props
page_map_path = ANALYSIS / 'report-page-map.json'
page_map = json.loads(page_map_path.read_text()) if page_map_path.exists() else {}
for n, old in enumerate(toc_nodes, 1):
    source_tokens = templates[15+n].xpath('.//w:r/*[self::w:t or self::w:tab]',namespaces=NS)
    original_text = ''.join((node.text or '') if node.tag==W+'t' else '\t' for node in source_tokens)
    label = clean_heading(original_text.rsplit('\t',1)[0])
    new = make_p(15+n, '')
    r = new.find(W + 'r')
    for child in list(r):
        r.remove(child)
    label_props = templates[15+n].xpath('.//w:r[w:t]/w:rPr',namespaces=NS)
    if label_props:
        r.append(static_toc_props(label_props[0]))
    for i,chunk in enumerate(label.split('\t')):
        if i:
            etree.SubElement(r,W+'tab')
        etree.SubElement(r,W+'t').text=chunk
    etree.SubElement(r, W + 'tab')
    page_run = etree.SubElement(new,W+'r')
    last_props = templates[15+n].xpath('.//w:r[w:t]/w:rPr',namespaces=NS)
    if last_props:
        page_run.append(deepcopy(last_props[-1]))
        for hidden in page_run.findall('.//'+W+'webHidden'):
            hidden.getparent().remove(hidden)
    etree.SubElement(page_run, W + 't').text = str(page_map.get(str(n), ''))
    old.getparent().replace(old, new)
bib_toc = make_p(32, '')
r = bib_toc.find(W + 'r'); r.find(W + 't').text = 'Список использованной литературы'
props=templates[32].xpath('.//w:r[w:t]/w:rPr',namespaces=NS)
if props:
    r.insert(0,static_toc_props(props[0]))
etree.SubElement(r, W + 'tab')
etree.SubElement(r, W + 't').text = str(page_map.get('bibliography', ''))
body.insert(list(body).index(body.findall(W + 'p')[32])+1, bib_toc)
appendix_toc = make_p(32, '')
r = appendix_toc.find(W + 'r'); r.find(W + 't').text = 'Приложение А. Задание на практику'
if props:
    r.insert(0,static_toc_props(props[0]))
etree.SubElement(r, W + 'tab')
etree.SubElement(r, W + 't').text = str(page_map.get('assignment', ''))
body.insert(list(body).index(bib_toc)+1, appendix_toc)

# No photographs from Nazar remain in the package. Institution logos stay.
for rel in list(rels):
    if rel.get('Type', '').endswith('/image') and rel.get('Target') != 'media/image1.png':
        rels.remove(rel)
for name in list(package):
    if name.startswith('word/media/') and name != 'word/media/image1.png':
        del package[name]

photos = {
    1: ('image3.jpeg','Оптическая система магнитного компаса',10.2),
    2: ('image4.jpeg','Приёмники ГНСС Transas T-701 и JRC; репитеры скорости и глубины',10.2),
    3: ('image2.jpeg','Радар Sperry Marine VisionMaster FT',10.2),
    4: ('image5.jpeg','Панель управления авторулевым Yokogawa PT500',16.0),
    5: ('image1.jpeg','Штурвал на рулевой стойке Raytheon Anschütz ComPilot 20',10.2),
    6: ('image6.jpeg','Консоль ГМССБ',16.0),
    7: ('image25.jpeg','Рабочее место на ходовом мостике',16.0),
}
photo_receipts = []

def add_photo(number):
    filename, caption, width_cm = photos[number]
    source = ANALYSIS / 'examples' / 'friend' / filename
    data = source.read_bytes()
    im = Image.open(source)
    orientation = im.getexif().get(274, 1)
    raw_w, raw_h = im.size
    display_w, display_h = (raw_h,raw_w) if orientation in (6,8) else (raw_w,raw_h)
    cx = round(width_cm * EMU_CM)
    cy = round(cx * display_h / display_w)
    rid = 'rIdPhoto' + str(number)
    name = 'equipment-figure-' + str(number) + '.jpeg'
    target = 'media/' + name
    relation = etree.SubElement(rels, '{'+NS['rel']+'}Relationship')
    relation.set('Id', rid)
    relation.set('Type', NS['r']+'/image')
    relation.set('Target', target)
    package['word/' + target] = data
    p = make_p(153,'',drop_numbering=True)
    ppr = p.find(W + 'pPr')
    keep = etree.SubElement(ppr,W+'keepNext'); keep.set(W+'val','1')
    # Lay out image orientation in OOXML, retaining the source JPEG bytes.
    shape_cx, shape_cy = (cy,cx) if orientation in (6,8) else (cx,cy)
    inline = CT_Inline.new_pic_inline(100+number, rid, name, shape_cx, shape_cy)
    xf = inline.xpath('.//a:xfrm')[0]
    if orientation in (6,8):
        xf.set('rot', '5400000' if orientation==6 else '16200000')
    elif orientation == 3:
        xf.set('rot','10800000')
    drawing = etree.SubElement(p.find(W+'r'),W+'drawing'); drawing.append(inline)
    append(format_report_p(p,'image'))
    cap = make_p(154,'Рис. '+str(number)+'. '+caption,drop_numbering=True)
    append(format_report_p(cap,'caption'))
    photo_receipts.append({'number':number,'source':filename,'caption':caption,
                           'sha256':hashlib.sha256(data).hexdigest(),'width_cm':width_cm,
                           'orientation_exif':orientation,'bytes_unchanged':True})

lines = (ANALYSIS/'REPORT_CONTENT.md').read_text().splitlines()
active = False
section = 1
pending = None
added_photos = set()
heading_receipts = []

def insert_photos_for_finished(sub):
    mapping = {'5.1':[1], '5.3':[2], '5.4':[3,4,5,6], '6.1':[7]}
    for n in mapping.get(sub,[]):
        if n not in added_photos:
            add_photo(n); added_photos.add(n)

for line in lines:
    if not line.strip():
        continue
    main_match = re.match(r'^## (\d+)\.',line)
    sub_match = re.match(r'^### (\d+\.\d+)\.',line)
    if main_match:
        insert_photos_for_finished(pending); pending = None
        section = int(main_match.group(1)); active = True
        label = clean_heading(text_of(templates[main_templates[section]]))
        p = make_p(main_templates[section],label)
        append(format_report_p(p,'heading'))
        heading_receipts.append({'number':section,'text':label,'template_paragraph':main_templates[section]})
    elif sub_match and active:
        insert_photos_for_finished(pending)
        pending = sub_match.group(1)
        idx = sub_templates[pending]
        # Keep the accepted subheading wording and layout, separate from prose.
        full = text_of(templates[idx])
        known = {
            '4.1':'     4.1 ОМС по Солнцу, планетам и звёздам.',
            '4.2':'     4.2 ОМС при помощи судовых приёмников.',
            '4.4':'     4.4 Выбор методов ОМС в различных условиях плавания.',
            '5.1':'5.1 Приборы курсоуказания (магнитный, гиро-, спутниковый компасы).',
            '5.2':'5.2 Приборы определения скорости и глубины.',
            '5.3':'5.3 Приборы определения места судна.',
            '5.4':'5.4 Приборы управления судном (ЭКНИС, САРП, РЛС, АИС, авторулевой).'}
        label = clean_heading(known.get(pending,full))
        append(format_report_p(make_p(idx,label),'heading'))
    elif line.startswith('## Список использованной литературы') and active:
        insert_photos_for_finished(pending); pending = None
        title=format_report_p(make_p(416,'Список использованной литературы'),'bibliography')
        etree.SubElement(title.find(W+'pPr'),W+'pageBreakBefore').set(W+'val','1')
        append(title)
        section = 18
    elif active:
        index = body_templates.get(section,417)
        p = make_p(index,re.sub(r'\*\*(.*?)\*\*',r'\1',line))
        if section == 18:
            p = make_p(417,line,drop_numbering=True)
        append(format_report_p(p,'reference' if section==18 else 'body'))
        immediate = None
        if line.startswith('Оптическая передача позволяет'):
            immediate = 1
        elif line.startswith('На рисунке 2 показаны'):
            immediate = 2
        elif line.startswith('На рисунке 3 показан'):
            immediate = 3
        elif line.startswith('Панель Yokogawa PT500 показана'):
            immediate = 4
        elif 'Рабочее место мостика показано на рисунке 7.' in line:
            immediate = 7
        elif line.startswith('Рулевое управление обеспечивает'):
            immediate = 5
        elif line.startswith('Радиооборудование связано'):
            immediate = 6
        if immediate is not None and immediate not in added_photos:
            add_photo(immediate); added_photos.add(immediate)

assert added_photos == set(range(1,8)), added_photos
assert len(heading_receipts) == 17
assert 'Еремян' not in text_of(tree)
assert 'PANTHERLIGHT' not in text_of(tree)
assert '62 000 кВт' in text_of(tree)
assert '6200' not in text_of(tree)
# Import the assignment as native Word paragraphs. Flatten its style inheritance
# into direct properties so coinciding style IDs cannot change fonts or layout.
assignment_source=ROOT/'originals'/'Gumrf_ProizvPraktikaSV_Zadanie (2).docx'
with ZipFile(assignment_source) as z:
    assignment_tree=etree.fromstring(z.read('word/document.xml'))
    assignment_styles=etree.fromstring(z.read('word/styles.xml'))
assignment_body=assignment_tree.find(W+'body')
style_by_id={s.get(W+'styleId'):s for s in assignment_styles.findall(W+'style')}

def style_chain(style_id):
    chain=[];seen=set()
    while style_id and style_id not in seen and style_id in style_by_id:
        seen.add(style_id);s=style_by_id[style_id];chain.insert(0,s)
        base=s.find(W+'basedOn');style_id=base.get(W+'val') if base is not None else None
    return chain

def merge_props(tag,*layers):
    props=etree.Element(W+tag)
    for layer in layers:
        if layer is None:continue
        for item in layer:
            old=props.find(item.tag)
            if old is None:
                props.append(deepcopy(item));continue
            if len(item) or len(old):
                props.replace(old,deepcopy(item))
            else:
                old.attrib.update(item.attrib)
    return props

default_p=assignment_styles.find('./'+W+'docDefaults/'+W+'pPrDefault/'+W+'pPr')
default_r=assignment_styles.find('./'+W+'docDefaults/'+W+'rPrDefault/'+W+'rPr')
assignment_elements=[deepcopy(x) for x in assignment_body if x.tag!=W+'sectPr']
assignment_drawing_index=200
for element in assignment_elements:
    for drawing in element.xpath('.//wp:docPr',namespaces=NS):
        drawing.set('id',str(assignment_drawing_index))
        assignment_drawing_index+=1
source_assignment_paras=assignment_body.findall(W+'p')
native_paras=[x for x in assignment_elements if x.tag==W+'p']
assert len(native_paras)==len(source_assignment_paras)==82
for element in assignment_elements:
    paras=[element] if element.tag==W+'p' else element.findall('.//'+W+'p')
    for p in paras:
        direct_p=p.find(W+'pPr')
        style=direct_p.find(W+'pStyle') if direct_p is not None else None
        chain=style_chain(style.get(W+'val') if style is not None else 'a')
        ppr=merge_props('pPr',default_p,*[s.find(W+'pPr') for s in chain],direct_p)
        for item in ppr.findall(W+'pStyle')+ppr.findall(W+'numPr'):
            ppr.remove(item)
        if direct_p is not None:p.remove(direct_p)
        p.insert(0,ppr)
        paragraph_mark=ppr.find(W+'rPr')
        for run in p.findall(W+'r'):
            direct_r=run.find(W+'rPr')
            char_style=direct_r.find(W+'rStyle') if direct_r is not None else None
            chars=style_chain(char_style.get(W+'val')) if char_style is not None else []
            rpr=merge_props('rPr',default_r,*[s.find(W+'rPr') for s in chain],paragraph_mark,
                            *[s.find(W+'rPr') for s in chars],direct_r)
            for item in rpr.findall(W+'rStyle'):rpr.remove(item)
            if direct_r is not None:run.remove(direct_r)
            run.insert(0,rpr)
# These three labels are automatic numbering in the source. Spell out the same
# visible labels without importing its numbering definitions into the report.
for index,prefix in {16:'1. ',17:'1.1 ',18:'1.2 '}.items():
    p=native_paras[index]
    text=p.find('.//'+W+'t');text.text=prefix+(text.text or '')
for original,copy in zip(source_assignment_paras,native_paras):
    assert text_of(original) in text_of(copy)
# Empty paragraphs in the supplied form become excessive vertical gaps when its
# styles are flattened. Keep all wording and fields, with explicit single spacing.
for p in [x for e in assignment_elements for x in ([e] if e.tag==W+'p' else e.findall('.//'+W+'p'))]:
    ppr=p.find(W+'pPr')
    spacing=ppr.find(W+'spacing')
    if spacing is None:spacing=etree.SubElement(ppr,W+'spacing')
    spacing.attrib.clear()
    for key,value in {'before':'0','after':'0','line':'240','lineRule':'auto'}.items():spacing.set(W+key,value)
    etree.SubElement(ppr,W+'widowControl').set(W+'val','1')
# Set explicit right tab stops for the form's signature fields. The original
# multiple default tabs wrap the teacher's surname after importing the form.
def native_form_line(index,left,right=None,before=0,after=0):
    p=native_paras[index];ppr=p.find(W+'pPr')
    font=p.find('./'+W+'r/'+W+'rPr')
    for child in list(p):
        if child is not ppr:p.remove(child)
    for tag in ['ind','tabs','jc']:
        for child in ppr.findall(W+tag):ppr.remove(child)
    ind=etree.SubElement(ppr,W+'ind')
    for key in ['left','right','firstLine']:ind.set(W+key,'0')
    etree.SubElement(ppr,W+'jc').set(W+'val','left')
    tabs=etree.SubElement(ppr,W+'tabs')
    stop=etree.SubElement(tabs,W+'tab');stop.set(W+'val','right');stop.set(W+'pos','9345')
    spacing=ppr.find(W+'spacing');spacing.set(W+'before',str(before));spacing.set(W+'after',str(after))
    run=etree.SubElement(p,W+'r')
    if font is not None:run.append(deepcopy(font))
    etree.SubElement(run,W+'t').text=left
    if right is not None:
        etree.SubElement(run,W+'tab')
        etree.SubElement(run,W+'t').text=right
native_form_line(71,text_of(native_paras[71]).strip(),before=240)
native_form_line(72,'к.т.н., доц., спкм','Г.И. Безбородов',after=120)
native_form_line(75,text_of(native_paras[75]).strip(),before=120)
fields=re.findall(r'_{2,}',text_of(source_assignment_paras[76]))
native_form_line(76,'подготовки',fields[-1])
fields=re.findall(r'_{2,}',text_of(source_assignment_paras[80]))
native_form_line(80,'Задание получил:',fields[0]+' /'+fields[1]+'/',before=240)
native_form_line(81,re.sub(r'\s+',' ',text_of(native_paras[81])).strip(),before=120)
etree.SubElement(native_paras[43].find(W+'pPr'),W+'pageBreakBefore').set(W+'val','1')
assignment_elements=[e for e in assignment_elements if text_of(e).strip()]
for original,copy in zip(source_assignment_paras,native_paras):
    def form_text(p):
        return ''.join('\t' if e.tag==W+'tab' else (e.text or '')
                       for e in p.xpath('.//w:t|.//w:tab',namespaces=NS))
    original_words=re.findall(r'[А-Яа-яЁёA-Za-z0-9]+',form_text(original))
    copied_words=re.findall(r'[А-Яа-яЁёA-Za-z0-9]+',form_text(copy))
    assert ' '.join(original_words) in ' '.join(copied_words),text_of(original)
title=format_report_p(make_p(416,'Приложение А. Задание на производственную практику'),'appendix')
etree.SubElement(title.find(W+'pPr'),W+'pageBreakBefore').set(W+'val','1')
append(title)
for element in assignment_elements:append(element)
assignment_receipts={'source':assignment_source.name,'native_word_text':True,
                     'source_paragraphs':82,
                     'inserted_nonempty_paragraphs':sum(e.tag==W+'p' for e in assignment_elements),
                     'all_source_wording_preserved':True,
                     'blank_spacing_paragraphs_omitted':True,'explicit_single_spacing':True,
                     'source_sha256':hashlib.sha256(assignment_source.read_bytes()).hexdigest()}
certificate_files = sorted(p for p in (ROOT/'originals'/'certificate').glob('*')
                           if p.suffix.lower() in ('.jpg','.jpeg','.png'))
certificate_receipts = []
if certificate_files:
    assert len(certificate_files)==2, 'The supplied certificate has two pages.'
    for i,source in enumerate(certificate_files,1):
        title = format_report_p(make_p(416,'Приложение Б. Справка о плавании' if i==1 else ''),'appendix')
        ppr = title.find(W+'pPr')
        etree.SubElement(ppr,W+'pageBreakBefore').set(W+'val','1')
        append(title)
        data=source.read_bytes(); im=Image.open(source)
        orientation=im.getexif().get(274,1)
        if orientation==1 and im.width>im.height:
            orientation=6
        dw,dh=(im.height,im.width) if orientation in (6,8) else im.size
        cx=round(16.5*EMU_CM);cy=round(cx*dh/dw)
        scx,scy=(cy,cx) if orientation in (6,8) else (cx,cy)
        rid='rIdCertificate'+str(i)
        ext='.png' if im.format=='PNG' else '.jpeg'
        target='media/certificate-page-'+str(i)+ext
        rel=etree.SubElement(rels,'{'+NS['rel']+'}Relationship')
        rel.set('Id',rid);rel.set('Type',NS['r']+'/image');rel.set('Target',target)
        package['word/'+target]=data
        p=make_p(153,'',drop_numbering=True)
        inline=CT_Inline.new_pic_inline(300+i,rid,source.name,scx,scy)
        if orientation in (3,6,8):
            inline.xpath('.//a:xfrm')[0].set('rot',{3:'10800000',6:'5400000',8:'16200000'}[orientation])
        etree.SubElement(p.find(W+'r'),W+'drawing').append(inline)
        append(format_report_p(p,'image'))
        certificate_receipts.append({'page':i,'source':source.name,'sha256':hashlib.sha256(data).hexdigest()})
package['word/document.xml'] = etree.tostring(tree,xml_declaration=True,encoding='UTF-8',standalone=True)
package['word/_rels/document.xml.rels'] = etree.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True)
# Do not inherit the template student's name and old dates in file properties.
core=etree.fromstring(package['docProps/core.xml'])
core_ns={'dc':'http://purl.org/dc/elements/1.1/','cp':'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
         'dt':'http://purl.org/dc/terms/'}
now=datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
for tag,value in [('dc:title','Отчёт о плавательной практике'),('dc:creator','Козырев Никита Алексеевич'),
                  ('cp:lastModifiedBy','Codex'),('cp:revision','1'),('dt:created',now),('dt:modified',now)]:
    node=core.find(tag,core_ns)
    if node is not None:node.text=value
package['docProps/core.xml']=etree.tostring(core,xml_declaration=True,encoding='UTF-8',standalone=True)
app=etree.fromstring(package['docProps/app.xml'])
appns='{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'
for tag,value in [('Pages',page_map.get('total_pages')),('Words',len(text_of(tree).split())),
                  ('Paragraphs',len(body.findall(W+'p')))]:
    node=app.find(appns+tag)
    if node is not None and value is not None:node.text=str(value)
package['docProps/app.xml']=etree.tostring(app,xml_declaration=True,encoding='UTF-8',standalone=True)
# Explicitly add content-type defaults if needed by the inherited package.
types = etree.fromstring(package['[Content_Types].xml'])
ctns = 'http://schemas.openxmlformats.org/package/2006/content-types'
if not types.xpath('*[@Extension="jpeg"]'):
    el = etree.SubElement(types,'{'+ctns+'}Default')
    el.set('Extension','jpeg'); el.set('ContentType','image/jpeg')
package['[Content_Types].xml'] = etree.tostring(types,xml_declaration=True,encoding='UTF-8',standalone=True)
out = ANALYSIS/'report.docx'
with ZipFile(out,'w',ZIP_DEFLATED) as z:
    for name,data in package.items():
        z.writestr(name,data)
with ZipFile(out) as z:
    assert z.testzip() is None
    for name in ['word/styles.xml','word/theme/theme1.xml','word/numbering.xml','word/footer1.xml','word/footer2.xml']:
        with ZipFile(SOURCE) as original:
            assert z.read(name) == original.read(name), name
(ANALYSIS/'report-build-receipt.json').write_text(json.dumps({'source_template':SOURCE.name,
    'template_styles_theme_numbering_footers_identical':True,'headings':heading_receipts,
    'photos':photo_receipts,'assignment':assignment_receipts,'paragraph_formats':format_receipts,
    'certificate_pages':certificate_receipts,
    'page_map':page_map,'output_bytes':out.stat().st_size},ensure_ascii=False,indent=2))
print(json.dumps({'output':str(out),'photos':len(photo_receipts),'bytes':out.stat().st_size,'layout_parts_identical':True}))
