# Denovo Attendance Report Generator

Processes ZKTeco biometric clocking data into a polished Word document attendance report for Denovo Apparel Ltd.

## Quick Start

```bash
pip install -r requirements.txt
python denovo_attendance.py clockings.dat employees.csv report.docx
```

## Input File Formats

### Clocking data (`.dat`)
Tab-separated, no header row. Only the first two columns are used:

```
1023    2025-06-02 07:45:12    0    1
1023    2025-06-02 13:00:05    0    1
1047    2025-06-02 08:03:44    0    1
```

| Column | Content |
|--------|---------|
| 1 | Employee ID |
| 2 | Timestamp `YYYY-MM-DD HH:MM:SS` |
| 3+ | Ignored |

### Employee CSV
No header row. Columns used: 1 (ID), 3 (first name), 4 (last name).

```
1023,,Jane,Smith,...
1047,,Mohamed,Al-Farsi,...
```

## Usage

```
python denovo_attendance.py <dat_file> <employee_csv> [output.docx]
                            [--company "Company Name"]
                            [--start YYYY-MM-DD]
                            [--end YYYY-MM-DD]
```

| Argument | Default | Description |
|----------|---------|-------------|
| `dat_file` | required | ZKTeco `.dat` clocking file |
| `employee_csv` | required | Employee CSV file |
| `output` | `attendance_report.docx` | Output file path |
| `--company` | `Denovo Apparel Ltd` | Company name in report header |
| `--start` | Earliest date in data | Report period start |
| `--end` | Latest date in data | Report period end |

## Calculation Rules

| Situation | Treatment |
|-----------|-----------|
| Clocking before 08:00 | Capped to 08:00 |
| All clockings | Rounded to nearest 15 minutes |
| Punches within 5 min of each other | Merged (handles double-taps) |
| 2 clockings (in/out only) | Fixed break schedule applied: 10:00–10:15, 13:00–13:30, 16:00–16:15 (only if shift spans that window) |
| 4+ clockings (employee clocked breaks) | Paired into work segments; fixed breaks also deducted from within each segment |
| Odd number of clockings | Best-effort pairing; flagged "Odd clockings - check" |
| Single clocking | Flagged "Single clocking - check"; hours not calculated |

## Report Output

**Title page**
- Company name, period, calculation rules
- Summary table: one row per employee, one column per Mon–Sun week, grand total

**Employee pages** (one per employee)
- Daily table: Date | Day | Clockings | Gross Hrs | Break | Net Hrs | Note
- Weekly subtotal rows between weeks
- Grand total row

**Visual flags**
- Break value in **red bold** → exceeds 60 minutes
- Break value with `*` → calculated from actual clocked breaks (not fixed schedule)
- Orange text → odd number of clockings on that day

## Requirements

- Python 3.9+
- `python-docx >= 1.1.2`
- `lxml >= 5.0.0`

## Known Implementation Notes

- `RGBColor` values accessed via index (`rgb[0]`), not `.r/.g/.b`
- OOXML `tcPr` child element ordering strictly maintained: `tcW → tcBorders → shd → tcMar → vAlign`
- Page size patched directly in XML to avoid stale `w:orient` attribute
- `w:zoom` `percent` attribute defaulted to `100` if missing
- Table column widths set explicitly in DXA and sum to exact content width
