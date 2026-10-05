"""Read-only inventory and tag export; not a Siemens SCL compiler."""
import argparse
import csv
import json
import re
from pathlib import Path

import openpyxl

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--repo', type=Path, default=Path('/workspace/AC200L'))
parser.add_argument('--output', type=Path, default=Path('/workspace/ac200l-dev/reports'))
args = parser.parse_args()
repo = args.repo.resolve(strict=True)
output = args.output.resolve()
if output == repo or repo in output.parents:
    parser.error('Keep generated reports outside the repository.')
sources = sorted((repo / 'AC200SourcesMain').glob('*.scl'))
sources += sorted((repo / 'AC200SourcesMain').glob('*.db'))
if not sources:
    raise RuntimeError('No Siemens source files found.')
blocks = []
symbols = set()
for path in sources:
    text = path.read_text(encoding='utf-8-sig')
    declaration = re.search(r'^\s*(FUNCTION_BLOCK|FUNCTION|DATA_BLOCK|TYPE)\s+"([^"]+)"', text)
    if declaration is None:
        raise RuntimeError(f'Cannot inventory declaration: {path.name}')
    blocks.append({'file': str(path.relative_to(repo)), 'kind': declaration[1], 'name': declaration[2]})
    # Remove comments before collecting quoted symbolic references.
    text = re.sub(r'\(\*.*?\*\)', '', text, flags=re.S)
    text = re.sub(r'//[^\n]*', '', text)
    symbols.update(re.findall(r'"([^"]+)"', text))

book = openpyxl.load_workbook(repo / 'PLCTags.xlsx', read_only=True, data_only=False)
tags = []
try:
    rows = book['PLC Tags'].iter_rows(values_only=True)
    header = next(rows)
    for expected in ('Name', 'Data Type', 'Logical Address'):
        if expected not in header:
            raise RuntimeError(f'Missing tag column: {expected}')
    for row in rows:
        if any(cell is not None for cell in row):
            tags.append(dict(zip(header, row)))
finally:
    book.close()
if not tags:
    raise RuntimeError('Tag workbook has no data rows.')
names = {tag['Name'] for tag in tags}
declarations = {block['name'] for block in blocks}
report = {
    'scope': 'Source inventory and Excel parsing only; no compile or PLC execution.',
    'blocks': blocks,
    'tag_count': len(tags),
    'duplicate_tag_names': sorted(name for name in names if sum(tag['Name'] == name for tag in tags) > 1),
    'quoted_symbols_not_in_export': sorted(symbols - names - declarations),
}
output.mkdir(parents=True, exist_ok=True)
(output / 'inventory.json').write_text(json.dumps(report, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
with (output / 'plc-tags.csv').open('w', encoding='utf-8', newline='') as target:
    writer = csv.DictWriter(target, fieldnames=header)
    writer.writeheader()
    writer.writerows(tags)
# Verify the generated export is actually readable and contains every tag.
with (output / 'plc-tags.csv').open(encoding='utf-8', newline='') as target:
    exported = list(csv.DictReader(target))
if len(exported) != len(tags) or {row['Name'] for row in exported} != names:
    raise RuntimeError('Tag export validation failed.')
print(f'Parsed {len(blocks)} block sources and {len(tags)} PLC tags; CSV round trip passed.')
print('Symbols absent from supplied block/tag exports:', ', '.join(report['quoted_symbols_not_in_export']) or 'none')
print('Duplicate tag names:', ', '.join(report['duplicate_tag_names']) or 'none')
print(f'Reports: {output}')
