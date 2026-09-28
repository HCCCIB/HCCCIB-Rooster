#!/usr/bin/env python3
"""
HCCC Rooster Builder
Leest de Excel uit /content/ en injecteert de data in index.html

Gebruik:
    python scripts/build.py

Het script zoekt automatisch het nieuwste .xlsx bestand in de content/ map.
"""

import openpyxl
from datetime import datetime
import json
import re
import sys
import glob
import os

# ── Zoek Excel bestand ────────────────────────────────────────────────────────
script_dir = os.path.dirname(os.path.abspath(__file__))
repo_dir = os.path.dirname(script_dir)
content_dir = os.path.join(repo_dir, 'content')
template_path = os.path.join(repo_dir, 'index.template.html')
output_path = os.path.join(repo_dir, 'index.html')

xlsx_files = glob.glob(os.path.join(content_dir, '*.xlsx'))
if not xlsx_files:
    print("FOUT: geen .xlsx bestand gevonden in content/")
    sys.exit(1)

xlsx_path = sorted(xlsx_files)[-1]  # Nieuwste op naam
print(f"Excel gevonden: {os.path.basename(xlsx_path)}")

# ── Excel inlezen ─────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(xlsx_path)
ws = wb['Rooster 2026']

date_cols = {}
for cell in ws[3]:
    if cell.value and isinstance(cell.value, datetime):
        date_cols[cell.column] = cell.value.strftime('%Y-%m-%d')

print(f"Datumkolommen gevonden: {len(date_cols)}")

# ── Medewerkers en rijen ──────────────────────────────────────────────────────
# Scan rijen om correcte posities te bepalen
name_to_rows = {}
i = 0
rows_found = []
for r in range(7, 60):
    a = ws.cell(row=r, column=1).value
    if a and str(a).strip() and str(a).strip() != '\xa0':
        b = ws.cell(row=r, column=2).value
        full = str(a).strip()
        if b and str(b).strip() and str(b).strip() != '\xa0':
            full = f"{str(a).strip()} {str(b).strip()}"
        rows_found.append((r, full))

print("Rijen gevonden:")
for r, name in rows_found:
    print(f"  rij {r}: {name}")

# Vaste mapping — pas aan als rijen verschuiven
EMPLOYEES = [
    ('Rens Kaal',              8,  9),
    ('Kurt Lemmer',           10, 11),
    ('Jan Peters',            12, 13),
    ('Gerwin Valkenborg',     14, 15),
    ('Hilda Rosenberg',       16, 17),
    ('Dennis van Lingen',     18, 19),
    ('Sander van Roon',       20, 21),
    ('Colette Verbueken',     22, 23),
    ('Mert Kasmer',           24, 25),
    ('Rachel Mens',           37, 38),
    ('Margreet van den Heuvel', 40, 41),
    ('Rowena Um',             43, 44),
    ('Lotte',                 45, 46),
    ('Binbin',                47, 48),
    ('Esther van Santen',     53, 54),
]

NAME_MAP = {
    'Rens Kaal': 'Rens', 'Kurt Lemmer': 'Kurt', 'Jan Peters': 'Jan',
    'Gerwin Valkenborg': 'Gerwin', 'Hilda Rosenberg': 'Hilda',
    'Dennis van Lingen': 'Dennis', 'Sander van Roon': 'Sander',
    'Colette Verbueken': 'Colette', 'Mert Kasmer': 'Mert',
    'Rachel Mens': 'Rachel', 'Margreet van den Heuvel': 'Margreet',
    'Rowena Um': 'Rowena', 'Lotte': 'Lotte', 'Binbin': 'Binbin',
    'Esther van Santen': 'Esther',
}

# ── Hulpfuncties ──────────────────────────────────────────────────────────────
def clean_val(val):
    if val is None: return None
    s = str(val).strip()
    return None if not s or s == '\xa0' else s

def is_cancelled(code): return bool(code and code.endswith('.'))
def strip_dot(code): return code.rstrip('.')

def clean_comment(text):
    if not text: return None
    lines = text.strip().split('\n')
    if lines and lines[0].strip().endswith(':'):
        lines = lines[1:]
    result = ' '.join(
        l.strip() for l in lines
        if '[Threaded comment]' not in l
        and 'Your version of Excel' not in l
        and l.strip()
    )
    return result if result else None

# ── Extractie ─────────────────────────────────────────────────────────────────
NOTES = {}
ROSTER = {}

for full_name, r1, r2 in EMPLOYEES:
    first = NAME_MAP[full_name]
    notes_by_date = {}
    shifts_by_date = {}

    for col, dstr in date_cols.items():
        c1 = ws.cell(row=r1, column=col)
        c2 = ws.cell(row=r2, column=col)
        v1 = clean_val(c1.value)
        v2 = clean_val(c2.value)
        note1 = clean_comment(c1.comment.text if c1.comment else None)
        note2 = clean_comment(c2.comment.text if c2.comment else None)

        active = []
        if v1 and not is_cancelled(v1): active.append(strip_dot(v1).lower())
        if v2 and not is_cancelled(v2): active.append(strip_dot(v2).lower())
        if active:
            shifts_by_date[dstr] = '+'.join(active) if len(active) > 1 else active[0]

        date_notes = {}
        if note1:
            key = strip_dot(v1).lower() if v1 else '_'
            date_notes[key] = note1
        if note2:
            key = strip_dot(v2).lower() if v2 else '_'
            date_notes[key] = note2
        if date_notes:
            notes_by_date[dstr] = date_notes

    ROSTER[first] = shifts_by_date
    if notes_by_date:
        NOTES[first] = notes_by_date

print("\nRooster statistieken:")
for name, shifts in ROSTER.items():
    print(f"  {name}: {len(shifts)} diensten")

# ── JSON valideren ────────────────────────────────────────────────────────────
roster_json = json.dumps(ROSTER, ensure_ascii=False, separators=(',', ':'))
notes_json  = json.dumps(NOTES,  ensure_ascii=False, separators=(',', ':'))
json.loads(roster_json)  # Valideer
json.loads(notes_json)
print("\nJSON validatie: OK")

# ── Injecteren in template ────────────────────────────────────────────────────
with open(template_path, 'r', encoding='utf-8') as f:
    html = f.read()

html = re.sub(
    r'const ROSTER = JSON\.parse\(`.*?`\);',
    f'const ROSTER = JSON.parse(`{roster_json}`);',
    html, flags=re.DOTALL
)
html = re.sub(
    r'const NOTES = JSON\.parse\(`.*?`\);',
    f'const NOTES = JSON.parse(`{notes_json}`);',
    html, flags=re.DOTALL
)

with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"\nKlaar! index.html bijgewerkt ({len(html):,} bytes)")
print(f"Gebouwd op: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
