#!/usr/bin/env python3
"""Conservative, plan-driven DOCX format repair; never overwrites input/output.

Uses lxml and standard library. Does not classify semantics, rewrite text, update
fields, move objects, or prove rendered layout. See references/audit-and-repair.md.
"""
import argparse
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}

def tag(name):
    return '{' + W + '}' + name

def child(parent, name):
    node = parent.find(tag(name))
    if node is None:
        node = ET.SubElement(parent, tag(name))
    return node

def attrs(parent, name, **values):
    node = child(parent, name)
    for k, v in values.items():
        node.set(tag(k), str(v))
    return node

# size, Chinese font, alignment, line (twips), lineRule, before/after, firstLineChars, leftChars
ROLES = {
    'body': (12, '宋体', 'both', 400, 'exact', 0, 0, 200, 0),
    'appendix_body': (12, '宋体', 'both', 400, 'exact', 0, 0, 200, 0),
    'abstract_zh': (12, '宋体', None, 400, 'exact', 0, 0, None, None),
    'abstract_en': (12, 'Times New Roman', None, 400, 'exact', 0, 0, None, None),
    'abstract_zh_title': (18, '黑体', 'center', 240, 'auto', 480, 360, 0, 0),
    'abstract_en_title': (18, 'Arial', 'center', 240, 'auto', 480, 360, 0, 0),
    'chapter': (16, '黑体', 'center', 240, 'auto', 480, 360, 0, 0),
    'section': (14, '黑体', 'center', 240, 'auto', 480, 120, 0, 0),
    'item': (13, '黑体', 'left', 240, 'auto', 240, 120, 0, 200),
    'subitem': (12, '黑体', 'left', 240, 'auto', 240, 120, 0, 200),
    'figure_caption': (10.5, '宋体', 'center', 240, 'auto', 120, 240, 0, 0),
    'table_caption': (10.5, '宋体', 'center', 240, 'auto', 120, 120, 0, 0),
    'references': (10.5, '宋体', None, 320, 'exact', 0, 0, None, None),
    'thanks': (12, '仿宋', None, 320, 'exact', 0, 0, None, None),
    'toc_title': (16, '黑体', 'center', 240, 'auto', 480, 360, 0, 0),
}
BOLD = {'abstract_zh_title', 'abstract_en_title', 'chapter', 'section', 'toc_title'}
OUTLINE = {'chapter': 0, 'section': 1, 'item': 2, 'subitem': 3}

def fonts(rpr, size, east, latin, bold=None):
    f = child(rpr, 'rFonts')
    # Themes otherwise override literal font names.
    for k in list(f.attrib):
        if ET.QName(k).localname.endswith('Theme'):
            del f.attrib[k]
    for key, value in [('eastAsia', east), ('ascii', latin), ('hAnsi', latin)]:
        f.set(tag(key), value)
    attrs(rpr, 'sz', val=int(size * 2))
    attrs(rpr, 'szCs', val=int(size * 2))
    if bold is not None:
        color = attrs(rpr, 'color', val='000000')
        for k in list(color.attrib):
            if ET.QName(k).localname.startswith('theme'):
                del color.attrib[k]
        attrs(rpr, 'b', val=int(bold))
        attrs(rpr, 'bCs', val=int(bold))

def text_of(p):
    return ''.join(p.xpath('.//w:t/text()', namespaces=NS))

def repair(source, plan, output):
    source, output = Path(source), Path(output)
    if source.resolve() == output.resolve() or output.exists():
        raise ValueError('Output must be a new file; refusing overwrite')
    raw = source.read_bytes()
    digest = hashlib.sha256(raw).hexdigest()
    if plan.get('source_sha256') != digest:
        raise ValueError('Source SHA-256 mismatch; rebuild the plan for this input')
    with ZipFile(source) as z:
        if z.testzip():
            raise ValueError('Corrupt DOCX ZIP')
        names = z.namelist()
        if len(names) != len(set(names)):
            raise ValueError('Duplicate ZIP members are not supported')
        if any(n.startswith('_xmlsignatures/') for n in names):
            raise ValueError('Digitally signed DOCX requires a signature-aware Word workflow')
        infos = z.infolist()
        package = {i.filename: z.read(i.filename) for i in infos}
    parser = ET.XMLParser(resolve_entities=False, no_network=True)
    doc = ET.fromstring(package['word/document.xml'], parser)
    styles = ET.fromstring(package['word/styles.xml'], parser)
    before_text = doc.xpath('//w:t/text()', namespaces=NS)
    paragraphs = doc.xpath('./w:body/w:p', namespaces=NS)
    planned, seen = [], set()
    for item in plan.get('paragraphs', []):
        index, role = item['index'], item['role']
        if not isinstance(index, int) or isinstance(index, bool) or index < 1 or index > len(paragraphs) or index in seen:
            raise ValueError(f'Invalid/duplicate paragraph index: {index}')
        seen.add(index)
        if role not in ROLES:
            raise ValueError(f'Unknown role: {role}')
        p = paragraphs[index - 1]
        if text_of(p) != item['text']:
            raise ValueError(f'Paragraph {index} text mismatch')
        # Do not flatten mathematical, tracked, structured or field-bearing content.
        protected = {'oMath', 'oMathPara', 'drawing', 'pict', 'object', 'fldChar', 'fldSimple',
                     'instrText', 'ins', 'del', 'moveFrom', 'moveTo', 'sdt', 'pPrChange',
                     'rPrChange', 'sectPr', 'footnoteReference', 'endnoteReference'}
        if any(ET.QName(n).localname in protected for n in p.iter()):
            raise ValueError(f'Paragraph {index} contains protected content; use targeted Word editing')
        mode = item.get('mode', 'two')
        if mode not in ('one', 'two'):
            raise ValueError('mode must be one or two')
        planned.append((index, p, role, mode))
    changes = []
    if plan.get('page_geometry', False):
        for i, sec in enumerate(doc.xpath('//w:sectPr', namespaces=NS), 1):
            page = attrs(sec, 'pgSz', w=11906, h=16838, orient='portrait')
            attrs(sec, 'pgMar', top=2154, bottom=2154, left=1814, right=1814,
                  header=1701, footer=1701, gutter=0)
            changes.append({'section': i, 'change': 'A4; margins 3.8/3.2cm; header/footer 3cm; gutter 0'})
    for index, p, role, mode in planned:
        size, east, align, line, rule, before, after, first, left = ROLES[role]
        latin = 'Arial' if role == 'abstract_en_title' else 'Times New Roman'
        if role in OUTLINE and mode == 'one':
            align, left = 'left', 0
        existing_pr = p.find(tag('pPr'))
        existing_style = existing_pr.find(tag('pStyle')) if existing_pr is not None else None
        base = existing_style.get(tag('val')) if existing_style is not None else 'Normal'
        # Retain inherited emphasis, numbering and other properties not targeted by the plan.
        style_id = f'NKF_{role}_{mode}_' + hashlib.sha256(base.encode()).hexdigest()[:8]
        matching = styles.xpath('./w:style[@w:styleId=$value]', namespaces=NS, value=style_id)
        if matching:
            style = matching[0]
            if style.get(tag('customStyle')) != '1':
                raise ValueError(f'Style name conflict: {style_id}')
            styles.remove(style)
        style = ET.SubElement(styles, tag('style'), {tag('type'): 'paragraph', tag('customStyle'): '1', tag('styleId'): style_id})
        attrs(style, 'name', val=style_id)
        attrs(style, 'basedOn', val=base)
        pf = child(style, 'pPr')
        if align:
            attrs(pf, 'jc', val=align)
        attrs(pf, 'spacing', before=before, after=after, line=line, lineRule=rule)
        if first is not None:
            attrs(pf, 'ind', firstLineChars=first, leftChars=left, right=0)
        if role in OUTLINE:
            attrs(pf, 'outlineLvl', val=OUTLINE[role])
            attrs(pf, 'keepNext', val=1)
        else:
            attrs(pf, 'outlineLvl', val=9)
        if role == 'table_caption' or role.endswith('_title'):
            attrs(pf, 'keepNext', val=1)
        if role == 'chapter':
            attrs(pf, 'pageBreakBefore', val=1)
        bold = True if role in BOLD else (False if role in {'item', 'subitem', 'figure_caption', 'table_caption'} else None)
        fonts(child(style, 'rPr'), size, east, latin, bold)
        direct = p.find(tag('pPr'))
        if direct is None:
            direct = ET.Element(tag('pPr')); p.insert(0, direct)
        # Remove only attributes explicitly governed by this role. Keep tabs, bookmarks, numbering, etc.
        governed = ['pStyle', 'spacing', 'outlineLvl']
        if align:
            governed.append('jc')
        if first is not None:
            governed.append('ind')
        if role in OUTLINE or role == 'table_caption' or role.endswith('_title'):
            governed.append('keepNext')
        if role == 'chapter':
            governed.append('pageBreakBefore')
        for key in governed:
            for node in direct.findall(tag(key)):
                direct.remove(node)
        ps = ET.Element(tag('pStyle'), {tag('val'): style_id}); direct.insert(0, ps)
        for run in p.xpath('.//w:r[w:t]', namespaces=NS):
            rp = run.find(tag('rPr'))
            if rp is None:
                rp = ET.Element(tag('rPr')); run.insert(0, rp)
            fonts(rp, size, east, latin, bold)
        changes.append({'paragraph': index, 'text': text_of(p), 'role': role, 'mode': mode, 'style': style_id})
    # OOXML pPr/rPr/sectPr children have schema ordering requirements.
    orders = {
        'pPr': 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange',
        'rPr': 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange',
        'sectPr': 'headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange',
    }
    for tree in (doc, styles):
        for kind, keys in orders.items():
            ranks = {name: i for i, name in enumerate(keys.split())}
            for node in tree.iter(tag(kind)):
                # Sort only nodes we created/edited, not unrelated document structures.
                if tree is doc and kind != 'sectPr' and not any(node is q or q in node.iterancestors() for _, q, _, _ in planned):
                    continue
                if tree is doc and kind == 'sectPr' and not plan.get('page_geometry'):
                    continue
                if tree is styles and not any(a.tag == tag('style') and a.get(tag('styleId'), '').startswith('NKF_') for a in node.iterancestors()):
                    continue
                node[:] = sorted(node, key=lambda e: ranks.get(ET.QName(e).localname, 999))
    if before_text != doc.xpath('//w:t/text()', namespaces=NS):
        raise ValueError('Unexpected text change')
    if plan.get('page_geometry') or planned:
        package['word/document.xml'] = ET.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
    if planned:
        package['word/styles.xml'] = ET.tostring(styles, xml_declaration=True, encoding='UTF-8', standalone=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open('xb') as f, ZipFile(f, 'w') as z:
        for info in infos:
            z.writestr(info, package[info.filename])
    with ZipFile(output) as z:
        if z.testzip():
            raise ValueError('Output ZIP verification failed')
    return {'source': str(source.resolve()), 'output': str(output.resolve()), 'source_sha256': digest,
            'changes': changes, 'status': '已修改待验证',
            'limitations': ['未更新Word域', '未渲染分页', '仅修改计划中段落及所选全节页面参数']}

def main():
    p = argparse.ArgumentParser(description=__doc__, epilog='Roles: ' + ', '.join(ROLES))
    p.add_argument('input', type=Path)
    p.add_argument('--plan', required=True, type=Path)
    p.add_argument('--out', required=True, type=Path)
    p.add_argument('--log', required=True, type=Path)
    args = p.parse_args()
    if args.log.exists() or args.log.resolve() in {args.input.resolve(), args.out.resolve(), args.plan.resolve()}:
        p.error('Log must be a new path distinct from input, output and plan')
    try:
        report = repair(args.input, json.loads(args.plan.read_text()), args.out)
    except (ValueError, KeyError) as e:
        p.error(str(e))
    args.log.parent.mkdir(parents=True, exist_ok=True)
    with args.log.open('x', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(json.dumps(report, ensure_ascii=False, indent=2))

if __name__ == '__main__':
    main()
