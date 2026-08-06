#!/usr/bin/env python3
"""Fill the Actions Tracker spreadsheet from MOE-format minutes of meeting (.docx).

Deterministic: same inputs -> same strings. No LLM.

Minutes (per .docx):
  - header table  : عنوان الاجتماع / التاريخ / الموقع ...
  - attendees table : الجهة | الاسم  (used to spot هوّز team members)
  - "المطلوب" table : ملاحظات | الوقت المتوقع/المقترح للانتهاء | المسؤول | المهمة

Tracker ('سجل المهام', body row = B C D E F G H I J K M):
  D المهمة                    <- المهمة
  E المسؤول                   <- المسؤول
  G تاريخ التسليم المستهدف     <- الوقت المتوقع للانتهاء
  H الاجتماع / المصدر          <- "<عنوان الاجتماع> — <التاريخ>"
  I ملاحظات                   <- ملاحظات
  C المسؤول من الفريق          <- best-effort: a هوّز member name in the text,
                                   else an entry from --owners mapping, else blank
  F الحالة                    <- EMPTY unless --status given
  M آخر تحديث                  <- today

Formatting is preserved: the tracker is edited surgically at the XML level
(no openpyxl round-trip), so the Excel Table (banding + filter), the الحالة
dropdown data-validation, and every cell style survive untouched. The table
range and dropdown range are extended to cover the new rows.

Usage:
  python fill_tracker.py --tracker "Actions Tracker.xlsx" \
      --minutes "3.3 Minutes of Meeting" [--out filled.xlsx]
      [--status "لم تبدأ"] [--owners owners.json]
"""

import argparse
import datetime
import json
import os
import re
import shutil
import sys
import zipfile

import docx
from docx.oxml.ns import qn

# Tracker status is left EMPTY by default (user decides later via the dropdown).
DEFAULT_STATUS = ""
OWNERS_FILE = None

# Cell layout of the tracker body rows (fixed template). Column letters and the
# XML style ids (banded: odd rows use STYLE_A, even STYLE_B, as in rows 7/8).
COLUMNS = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "M"]
STYLE_A = {"B": "17", "C": "19", "D": "18", "E": "19", "F": "19",
           "G": "20", "H": "19", "I": "37", "J": "18", "K": "50", "M": "37"}
STYLE_B = {"B": "21", "C": "23", "D": "22", "E": "23", "F": "15",
           "G": "24", "H": "23", "I": "22", "J": "22", "K": "51", "M": "22"}


# --------------------------------------------------------------------------
# minutes .docx parsing
# --------------------------------------------------------------------------
def _blocks(document):
    from docx.table import Table
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:tbl"):
            yield Table(child, document)


def _rows(table):
    return [[c.text.strip() for c in row.cells] for row in table.rows]


def parse_minutes(path, owners=None):
    """Return list of action dicts parsed from one minutes .docx.

    owners: optional dict mapping a responsible-department substring -> is-hawaz
    owner name, used to fill 'المسؤول من الفريق' when no name appears inline.
    """
    doc = docx.Document(path)
    tables = list(_blocks(doc))

    title = date = ""
    if tables:
        for row in _rows(tables[0]):
            if "عنوان الاجتماع" in row:
                title = next((c for c in row if c and c != "عنوان الاجتماع"), "")
            if "التاريخ" in row:
                idx = row.index("التاريخ")
                if idx - 1 >= 0:
                    date = row[idx - 1]

    team_names = []
    for tb in tables:
        data = _rows(tb)
        if len(tb.columns) == 2 and data and "الاسم" in data[0]:
            for ent, name in data[1:]:
                if "هوّز" in ent and name:
                    team_names.append(name)

    actions = []
    for tb in tables:
        data = _rows(tb)
        hdr = next((i for i, r in enumerate(data)
                    if "المهمة" in r and "المسؤول" in r), None)
        if hdr is None:
            continue
        head = data[hdr]
        ti = next((i for i, c in enumerate(head) if "المهمة" in c), None)
        ri = next((i for i, c in enumerate(head) if "المسؤول" in c), None)
        di = next((i for i, c in enumerate(head) if "انتهاء" in c), None)
        ni = next((i for i, c in enumerate(head) if "ملاحظات" in c), None)

        for row in data[hdr + 1:]:
            if not any(row):
                continue
            task = row[ti] if ti is not None and ti < len(row) else ""
            if not task.strip():
                continue
            resp = row[ri] if ri is not None and ri < len(row) else ""
            due = row[di] if di is not None and di < len(row) else ""
            notes = row[ni] if ni is not None and ni < len(row) else ""
            actions.append({
                "task": task,
                "responsible": resp,
                "due": due,
                "notes": notes,
                "team_owner": _match_owner(team_names, owners, task, notes, resp),
                "source": f"{title} — {date}" if (title and date) else title or date,
            })
    return actions


def _match_owner(team_names, owners, task, notes, responsible):
    """Deterministic best-effort attribution of the is-hawaz owner.

    Priority: a team member's first name found inline -> an --owners mapping hit
    (by responsible-department substring) -> "" (blank for a human to fill).
    """
    hay = f"{task} {notes} {responsible}"
    for name in team_names:
        first = name.split()[0] if name.split() else name
        if first and first in hay:
            return name
    if owners:
        for key, owner in owners.items():
            if key and key in responsible:
                return owner
    return ""


# --------------------------------------------------------------------------
# surgical xlsx fill (preserves table + dropdowns + styling)
# --------------------------------------------------------------------------
def _xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def _cell(letter, row, style, value):
    ref = f"{letter}{row}"
    if value in ("", None):
        return f'<c r="{ref}" s="{style}"/>'
    v = _xml_escape(str(value))
    return (f'<c r="{ref}" s="{style}" t="inlineStr">'
            f'<is><t xml:space="preserve">{v}</t></is></c>')


def fill_tracker(tracker_path, actions, out=None, status=DEFAULT_STATUS, owners=None):
    """Append actions rows to the tracker via surgical XML edit.

    Returns (written, dest_path).
    """
    # (owners mapping is applied during parsing; passed here only for the CLI run)
    sheet_path = "xl/worksheets/sheet1.xml"
    table_path = "xl/tables/table1.xml"
    with zipfile.ZipFile(tracker_path) as z:
        sheet = z.read(sheet_path).decode("utf-8")
        table = z.read(table_path).decode("utf-8")

    # existing rows -> ordered list of (row_number, raw_xml); matches both
    # self-closing (<row .../>) and paired (<row ...>...</row>) row elements
    row_re = re.compile(r"<row\b[^>]*?(?:/>|>.*?</row>)", re.S)
    existing = []
    for m in row_re.finditer(sheet):
        num = int(re.search(r'r="(\d+)"', m.group(0)).group(1))
        existing.append((num, m.group(0)))
    existing = dict(existing)  # row_num -> xml (keeps document order)

    last_orig = max(existing) if existing else 6
    first = 7                       # row 6 holds the original sample -> start at 7
    last = first - 1 + len(actions)  # last data row (7 when no actions)

    today = datetime.date.today().strftime("%d/%m/%Y")
    generated = {}
    for i, a in enumerate(actions):
        r = first + i
        style = STYLE_A if r % 2 == 1 else STYLE_B
        values = {
            "B": str(i + 2),          # continue numbering after the sample row (#1)
            "C": a["team_owner"],
            "D": a["task"],
            "E": a["responsible"],
            "F": status,              # EMPTY by default
            "G": a["due"],
            "H": a["source"],
            "I": a["notes"],
            "J": "",
            "K": "",
            "M": today,
        }
        cells = "".join(_cell(col, r, style[col], values[col]) for col in COLUMNS)
        generated[r] = f'<row r="{r}" spans="2:13" ht="16.5">{cells}</row>'

    # reassemble sheetData: keep every original row, replacing in-range rows with
    # generated data, then append any generated rows that exceed the original max
    parts = []
    for num, xml in existing.items():
        if num in generated:
            parts.append(generated[num])
        else:
            parts.append(xml)
    for num in sorted(k for k in generated if k > last_orig):
        parts.append(generated[num])
    sheet_data = "<sheetData>" + "".join(parts) + "</sheetData>"
    sheet = re.sub(r"<sheetData>.*?</sheetData>",
                   lambda _m: sheet_data, sheet, count=1, flags=re.S)

    # extend the sheet dimension, the status-dropdown range, and the table range
    sheet = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="B1:DJ{last}"/>',
                   sheet, count=1)
    sheet = re.sub(r"(<xm:sqref>F\d+:F)\d+(</xm:sqref>)",
                   lambda m: m.group(1) + str(last) + m.group(2), sheet, count=1)
    table = re.sub(r'ref="B5:M\d+"', f'ref="B5:M{last}"', table)

    dest = out or tracker_path
    tmp = dest + ".tmp"
    with zipfile.ZipFile(tracker_path) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_path:
                data = sheet.encode("utf-8")
            elif item.filename == table_path:
                data = table.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, dest)
    return len(actions), dest


# --------------------------------------------------------------------------
def collect_minutes(minutes_arg):
    """Return a list of .docx paths from a folder or a zip archive."""
    files = []
    if os.path.isdir(minutes_arg):
        for root, _, names in os.walk(minutes_arg):
            for n in names:
                if n.lower().endswith(".docx"):
                    files.append(os.path.join(root, n))
    elif zipfile.is_zipfile(minutes_arg):
        import tempfile
        tmp = tempfile.mkdtemp(prefix="minutes_")
        with zipfile.ZipFile(minutes_arg) as z:
            for name in z.namelist():
                if name.lower().endswith(".docx"):
                    z.extract(name, tmp)
        for root, _, names in os.walk(tmp):
            for n in names:
                if n.lower().endswith(".docx"):
                    files.append(os.path.join(root, n))
    else:
        raise SystemExit(f"Not a folder or zip: {minutes_arg}")
    return sorted(files)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tracker", required=True, help="Path to Actions Tracker .xlsx")
    ap.add_argument("--minutes", required=True,
                    help="Folder or .zip containing minutes .docx files")
    ap.add_argument("--out", default=None, help="Output .xlsx (default: overwrite tracker)")
    ap.add_argument("--status", default=DEFAULT_STATUS,
                    help='Status for new rows (default: EMPTY "لم تبدأ" -> no value)')
    ap.add_argument("--owners", default=OWNERS_FILE,
                    help="JSON mapping responsible-department -> is-hawaz owner")
    args = ap.parse_args(argv)

    owners = None
    if args.owners:
        with open(args.owners, encoding="utf-8") as f:
            owners = json.load(f)

    minutes_files = collect_minutes(args.minutes)
    if not minutes_files:
        sys.exit("No .docx minutes found.")

    all_actions = []
    for f in minutes_files:
        all_actions.extend(parse_minutes(f, owners))

    if not all_actions:
        sys.exit("No action items found in any minutes file.")

    written, dest = fill_tracker(args.tracker, all_actions,
                                 out=args.out, status=args.status, owners=owners)
    print(f"Parsed {len(minutes_files)} minutes files; "
          f"wrote {written} action rows -> {dest}")


if __name__ == "__main__":
    main()
