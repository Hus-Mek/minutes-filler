# minutes‑tracker filler

Fills the **Actions Tracker** spreadsheet (`سجل المهام` sheet) from MOE‑format
**Minutes of Meeting** `.docx` files, in a fully **deterministic** way (no LLM,
no randomness — same input → same strings).

## Install

```
pip install -r requirements.txt
```

## Usage

```
python fill_tracker.py --tracker "Actions Tracker.xlsx" \
    --minutes "3.3 Minutes of Meeting" --out filled.xlsx
```

- `--minutes` : folder (recursively scanned) **or** a `.zip` of `.docx` files.
- `--out`     : output path (default: overwrite the tracker in place).
- `--status`  : value for the الحالة column. **Default is empty** — the cell is
  left blank so you pick from the status dropdown (`لم تبدأ / جاري العمل /
  مكتمل / متأخر`) later.
- `--owners`  : optional JSON mapping a responsible-department substring → the
  هوّز owner (`المسؤول من الفريق`). See "Who does what" below.

## How it works

Every minutes file uses the same template. The script reads three tables:

| Table | Fields |
|---|---|
| Header   | meeting title (`عنوان الاجتماع`), date (`التاريخ`) |
| Attendees| `الجهة | الاسم` (used to detect هوّز team members) |
| المطلوب   | `ملاحظات | الوقت المتوقع/المقترح للانتهاء | المسؤول | المهمة` |

Each action row becomes a tracker row. Tracker body columns: `B C D E F G H I J K M`.

| Tracker cell | Source |
|---|---|
| D المهمة       | المهمة |
| E المسؤول      | المسؤول |
| G تاريخ التسليم المستهدف | الوقت المتوقع للانتهاء |
| H الاجتماع / المصدر      | `"<عنوان الاجتماع> — <التاريخ>"` |
| I ملاحظات      | ملاحظات |
| C المسؤول من الفريق | best-effort: a هوّز member name found inline, else `--owners` mapping, else blank |
| F الحالة       | **empty** (unless `--status`) |
| M آخر تحديث    | today |
| B #            | sequential |

### "Who does what" (`المسؤول من الفريق`)

The minutes record the *client* responsible party (e.g. `وزارة الطاقة (إدارة
التقنيات المستجدة)`) in `المسؤول`, **not** the هوّز person who owns the action.
So the هوّز owner is not in the source and cannot be guessed deterministically.

- If a هوّز member's first name appears inside the task/notes text, it is picked up automatically.
- Otherwise pass a mapping so the tool attributes them, e.g.
  `owners.json`:
  ```json
  {
    "إدارة التقنيات المستجدة": "نوف التويجري",
    "إدارة الحوكمة والأداء": "لمياء العبدالعالي"
  }
  ```
  Any `المسؤول` string *containing* the key is assigned that owner.
- With no match, the cell is left blank for a human.

## Formatting is preserved

The tracker is edited surgically at the XML level (zip member `sheet1.xml` +
`table1.xml` only; every other part — styles, images/logo, custom XML, theme —
is copied byte-for-byte). The Excel `Table2` (banded rows + filter), the
`الحالة` dropdown data-validation, and all cell styles survive. Their ranges
are extended to cover the newly added rows.
