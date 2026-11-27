"""Fill the bundled six-section DOCX template without rebuilding its styles."""
import argparse
from copy import deepcopy
from hashlib import sha256
import json
from pathlib import Path
import re
from zipfile import ZipFile

from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
HEADINGS = {
    1: '一、文献基本信息',
    10: '二、摘要原文及中文翻译',
    15: '三、Introduction总结（450字以内）',
    17: '四、实验方法总结（450字以内）',
    26: '六、结论与展望',
}


def text_of(p):
    return ''.join(p.xpath('.//w:t/text()', namespaces=NS))


def visible_length(value):
    return len(re.sub(r'\s+', '', value))


def validate_content(data):
    if not isinstance(data, dict):
        raise ValueError('Content must be a JSON object.')
    singles = ['title', 'abstract_en', 'abstract_zh', 'introduction', 'methods']
    for key in singles:
        if not isinstance(data.get(key), str) or not data[key].strip():
            raise ValueError(f'{key} must be a nonempty string.')
    for key, count in [('metadata', 8), ('conclusions', 2), ('captions', None)]:
        values = data.get(key)
        if not isinstance(values, list) or (count is not None and len(values) != count):
            raise ValueError(f'{key} must be a list' + (f' of {count} items.' if count else '.'))
        if any(not isinstance(t, str) or not t.strip() for t in values):
            raise ValueError(f'{key} contains an empty or non-string paragraph.')
    values = [data[k] for k in singles] + data['metadata'] + data['conclusions'] + data['captions']
    for value in values:
        if '\n' in value or '\r' in value or any(ord(c) < 32 for c in value):
            raise ValueError('Each string must be one paragraph with no control characters.')
        if '{{' in value or '}}' in value:
            raise ValueError('Replace template field markers with paper content.')
    for key in ['introduction', 'methods']:
        if visible_length(data[key]) > 450:
            raise ValueError(f'{key} exceeds 450 visible characters: {visible_length(data[key])}')
    for index, caption in enumerate(data['captions'], 1):
        match = re.match(r'^图\s*(\d+)(?=\s|[.．：:]|$)', caption)
        if not match or int(match.group(1)) != index:
            raise ValueError(f'Caption {index} must begin with its consecutive figure number.')


def build(content, output, report=None, overwrite=False):
    template = Path(__file__).resolve().parents[1] / 'assets' / 'reading-template.docx'
    output = Path(output).resolve()
    if output == template.resolve() or output == Path(content).resolve():
        raise ValueError('Output cannot overwrite the template or content JSON.')
    if output.exists() and not overwrite:
        raise FileExistsError(f'Output exists; use --overwrite to revise: {output}')
    if report and Path(report).resolve() in {template.resolve(), output, Path(content).resolve()}:
        raise ValueError('Report path must be distinct from inputs and DOCX output.')
    data = json.loads(Path(content).read_text(encoding='utf-8-sig'))
    validate_content(data)
    source_hash = sha256(template.read_bytes()).hexdigest()
    with ZipFile(template) as z:
        infos = z.infolist()
        original = {i.filename: z.read(i.filename) for i in infos}
    root = E.fromstring(original['word/document.xml'])
    body = root.find('w:body', NS)
    paras = body.findall('w:p', NS)
    if len(paras) != 29 or any(text_of(paras[i]) != t for i, t in HEADINGS.items()):
        raise ValueError('Unexpected template structure; only the bundled six-section template is supported.')
    if text_of(paras[0]) != '{{title}}':
        raise ValueError('Bundled template field map does not match.')

    def paragraph(slot, value):
        p = deepcopy(paras[slot])
        texts = p.findall('.//w:t', NS)
        if not texts:
            raise ValueError(f'Template slot {slot} has no text run.')
        texts[0].text = value
        texts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        for node in texts[1:]:
            node.text = ''
        return p

    result = [paragraph(0, data['title']), deepcopy(paras[1])]
    result += [paragraph(i + 2, t) for i, t in enumerate(data['metadata'])]
    result += [deepcopy(paras[10]), deepcopy(paras[11]), paragraph(12, data['abstract_en']),
               deepcopy(paras[13]), paragraph(14, data['abstract_zh']), deepcopy(paras[15]),
               paragraph(16, data['introduction']), deepcopy(paras[17]), paragraph(18, data['methods'])]
    count = len(data['captions'])
    label = f'五、图注完整翻译（Fig. 1–{count}）' if count > 1 else ('五、图注完整翻译（Fig. 1）' if count else '五、图注完整翻译')
    result.append(paragraph(19, label))
    result += [paragraph(20 + min(i, 5), t) for i, t in enumerate(data['captions'])]
    if not count:
        result.append(paragraph(20, '原文没有编号的正文图。'))
    result += [deepcopy(paras[26]), paragraph(27, data['conclusions'][0]), paragraph(28, data['conclusions'][1])]
    for p in paras:
        body.remove(p)
    for i, p in enumerate(result):
        body.insert(i, p)
    patches = {'word/document.xml': E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)}
    core = E.fromstring(original['docProps/core.xml'])
    for node in core:
        name = E.QName(node).localname
        if name == 'title':
            node.text = data['title']
        elif name in {'subject', 'creator', 'lastModifiedBy', 'description', 'keywords'}:
            node.text = ''
    patches['docProps/core.xml'] = E.tostring(core, encoding='UTF-8', xml_declaration=True, standalone=True)
    app = E.fromstring(original['docProps/app.xml'])
    for node in list(app):
        if E.QName(node).localname in {'Pages', 'Words', 'Characters', 'CharactersWithSpaces', 'Lines', 'Paragraphs', 'TotalTime'}:
            app.remove(node)
    patches['docProps/app.xml'] = E.tostring(app, encoding='UTF-8', xml_declaration=True, standalone=True)
    if '{{' in ''.join(text_of(p) for p in result):
        raise ValueError('Unfilled field marker remains in output.')
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output, 'w') as z:
        for item in infos:
            z.writestr(item, patches.get(item.filename, original[item.filename]))
    preserved = [name for name in original if name not in patches]
    with ZipFile(output) as z:
        if not all(z.read(name) == original[name] for name in preserved):
            raise RuntimeError('Preserved package parts changed.')
    if sha256(template.read_bytes()).hexdigest() != source_hash:
        raise RuntimeError('Source template changed during authoring.')
    result_report = {
        'output': str(output), 'paragraphs': len(result), 'figures': count,
        'introduction_chars': visible_length(data['introduction']),
        'methods_chars': visible_length(data['methods']), 'preserved_parts': len(preserved),
        'template_sha256': source_hash, 'output_sha256': sha256(output.read_bytes()).hexdigest(),
        'visual_qa': 'pending: render and inspect every page',
    }
    if report:
        report_path = Path(report)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(result_report, ensure_ascii=False, indent=2), encoding='utf-8')
    return result_report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--content', required=True)
    parser.add_argument('--output', required=True)
    parser.add_argument('--report')
    parser.add_argument('--overwrite', action='store_true')
    args = parser.parse_args()
    print(json.dumps(build(args.content, args.output, args.report, args.overwrite), ensure_ascii=False))


if __name__ == '__main__':
    main()
