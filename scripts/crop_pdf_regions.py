"""Crop explicitly measured regions from Poppler-rendered PDF pages."""
import argparse
from hashlib import sha256
import json
import math
from pathlib import Path
import re
import shutil
import subprocess
import tempfile

from PIL import Image


def crop(args):
    pdf = Path(args.pdf).resolve(strict=True)
    manifest_path = Path(args.manifest).resolve(strict=True)
    manifest = json.loads(manifest_path.read_text(encoding='utf-8-sig'))
    items = manifest.get('items')
    if not isinstance(items, list) or not items:
        raise ValueError('Manifest must contain a nonempty items list.')
    if not 72 <= args.dpi <= 1200:
        raise ValueError('DPI must be between 72 and 1200.')
    out = Path(args.out_dir).resolve()
    names = set()
    numbers = set()
    figures = 0
    for item in items:
        if not isinstance(item, dict):
            raise ValueError('Each item must be an object.')
        name = item.get('name')
        if (not isinstance(name, str) or not name or name != name.strip() or
                re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith('.') or
                re.match(r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)', name, re.I)):
            raise ValueError(f'Invalid output name: {name!r}')
        if name.casefold() in names:
            raise ValueError(f'Duplicate output name: {name}')
        names.add(name.casefold())
        if type(item.get('page')) is not int or item['page'] < 1:
            raise ValueError(f'{name}: page must be a positive integer.')
        box = item.get('bbox')
        if (not isinstance(box, list) or len(box) != 4 or
                any(type(v) not in (int, float) or not math.isfinite(v) for v in box) or
                not (0 <= box[0] < box[2] <= 1 and 0 <= box[1] < box[3] <= 1)):
            raise ValueError(f'{name}: bbox must be [left, top, right, bottom] within 0..1.')
        if item.get('kind') == 'figure':
            number = item.get('number')
            if type(number) is not int or number < 1 or number in numbers:
                raise ValueError(f'{name}: figure number must be a unique positive integer.')
            numbers.add(number)
            figures += 1
        target = out / (name + '.png')
        if target.exists() and not args.overwrite:
            raise FileExistsError(f'Output exists; use --overwrite to revise: {target}')
        if target in {pdf, manifest_path}:
            raise ValueError('Output cannot overwrite an input.')
    expected = manifest.get('expected_figure_count')
    if expected is not None and (type(expected) is not int or expected != figures):
        raise ValueError(f'Expected {expected} figure items, found {figures}.')
    report_path = Path(args.report).resolve() if args.report else None
    if report_path and report_path in {pdf, manifest_path, *[out / (i['name'] + '.png') for i in items]}:
        raise ValueError('Report path must be distinct from inputs and image outputs.')
    executable = shutil.which(args.pdftoppm)
    if executable is None:
        raise FileNotFoundError('pdftoppm not found; supply its executable path with --pdftoppm.')
    work_dir = Path(args.work_dir).resolve()
    work_dir.mkdir(parents=True, exist_ok=True)
    source_hash = sha256(pdf.read_bytes()).hexdigest()
    results = []
    with tempfile.TemporaryDirectory(prefix='pdf-regions-', dir=work_dir) as temp:
        temp_path = Path(temp)
        for page in sorted({i['page'] for i in items}):
            prefix = temp_path / f'page-{page}'
            process = subprocess.run([executable, '-f', str(page), '-l', str(page), '-r', str(args.dpi),
                                      '-png', '-singlefile', str(pdf), str(prefix)],
                                     stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if process.returncode:
                raise RuntimeError(f'Page {page} rendering failed: ' + process.stderr.decode('utf-8', errors='replace'))
            with Image.open(prefix.with_suffix('.png')) as full:
                for item in [i for i in items if i['page'] == page]:
                    x0, y0, x1, y1 = item['bbox']
                    box = (round(x0 * full.width), round(y0 * full.height),
                           round(x1 * full.width), round(y1 * full.height))
                    if box[2] <= box[0] or box[3] <= box[1]:
                        raise ValueError(f'{item["name"]}: bbox rounds to an empty crop.')
                    cropped = full.crop(box)
                    staged = temp_path / (item['name'] + '.png')
                    cropped.save(staged, optimize=True, dpi=(args.dpi, args.dpi))
                    results.append({'name': item['name'], 'page': page, 'bbox': item['bbox'],
                                    'pixels': list(cropped.size), 'file': str(out / staged.name)})
        # Publish only after all requested pages rendered successfully.
        if sha256(pdf.read_bytes()).hexdigest() != source_hash:
            raise RuntimeError('Source PDF changed while rendering.')
        out.mkdir(parents=True, exist_ok=True)
        for item in results:
            shutil.copyfile(temp_path / (item['name'] + '.png'), item['file'])
    report = {'source_pdf': str(pdf), 'source_sha256': source_hash, 'dpi': args.dpi,
              'figures': figures, 'figure_numbers': sorted(numbers), 'regions': results,
              'visual_qa': 'pending: inspect every crop against its original page'}
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--pdf', required=True)
    parser.add_argument('--manifest', required=True)
    parser.add_argument('--out-dir', required=True)
    parser.add_argument('--work-dir', required=True)
    parser.add_argument('--dpi', type=int, default=400)
    parser.add_argument('--pdftoppm', default='pdftoppm')
    parser.add_argument('--report')
    parser.add_argument('--overwrite', action='store_true')
    args = parser.parse_args()
    report = crop(args)
    print(json.dumps({'figures': report['figures'], 'regions': len(report['regions']), 'source_sha256': report['source_sha256']}))


if __name__ == '__main__':
    main()
