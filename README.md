# minutes‑tracker filler

Fills the **Actions Tracker** spreadsheet (`سجل المهام` sheet) from MOE‑format
**Minutes of Meeting** `.docx` files, in a fully **deterministic** way (no LLM,
no randomness — same input → same output).

## How it works

Each minutes file follows a fixed template. The script reads three tables:

| Table | Fields |
|---|---|
| Header  | meeting title (`عنوان الاجتماع`), date (`التاريخ`) |
| Attendees | `الجهة | الاسم` (used to detect هوّز team members) |
| المطلوب (actions) | `ملاحظات | الوقت المتوقع/المقترح للانتهاء | المسؤول | المهمة` |

Each action-table row becomes a tracker row:

```
المهمة                                     <- المهمة
المسؤول                                    <- المسؤول
تاريخ التسليم المستهدف / تاريخ المستهدف وفق المحضر  <- الوقت المتوقع للانتهاء
الاجتماع / المصدر                          <- "<عنوان الاجتماع> — <التاريخ>"
ملاحظات                                    <- ملاحظات
الحالة                                     <- لم تبدأ (default)
آخر تحديث                                  <- today
المسؤول من الفريق                          <- best-effort deterministic match of
                                               a هوّز member name found in the text
                                               (blank when none is found)
```

## Install

```
pip install -r requirements.txt
```

## Usage

```
python fill_tracker.py --tracker "Actions Tracker.xlsx" \
    --minutes "3.3 Minutes of Meeting" \
    --out filled.xlsx
```

`--minutes` accepts either a folder (recursively scanned for `.docx`) or a zip
archive. By default the tracker is overwritten in place; use `--out` to write a
copy. `--status` sets the status written into new rows (default `لم تبدأ`).

## Notes on determinism

- Extraction relies only on table structure + Arabic header labels — stable.
- `المسؤول من الفريق` (the person on the هوّز side owning the action) cannot
  always be derived from the minutes; the script fills it when a team member's
  first name appears in the task/notes text, otherwise leaves it blank for a
  human to complete.
