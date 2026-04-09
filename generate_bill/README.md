# Generate Bill — Professional Indian Construction Contractor Bill

A Python project that generates a professional 3-page Indian construction contractor bill in both **PDF** and **DOCX** formats.

## Output Files

| File | Description |
|------|-------------|
| `Bill_Final.pdf` | Professional PDF bill (US Letter, 3 pages) |
| `Bill_Final.docx` | Microsoft Word document bill |

## Bill Contents

The generated bill covers **6 work sections**:

| Section | Description | Sub Total (₹) |
|---------|-------------|--------------|
| A | Civil Work (excavation, PCC, RCC, brick, plaster) | 75,500.00 |
| B | Flooring Work (vitrified tiles, bathroom, kitchen, kota stone) | 30,700.00 |
| C | Granite Stone Work (counter top, window sill, staircase, threshold + skirting) | 24,080.00 |
| D | Painting Work (primer, putty, emulsion, exterior) | 30,000.00 |
| E | Electrical Work (wiring, MCB, earthing) | 15,250.00 |
| F | Plumbing Work (water supply, drainage, sanitary fitting) | 11,440.00 |
| | **GRAND TOTAL** | **₹1,86,970.00** |

### Granite Stone Work (Section C)

This section seamlessly includes both original granite work and the newly completed **Granite Stone Skirting** items:

| Item | Measurement |
|------|-------------|
| Middliya Room Skirting | 72.00 Rft |
| Porch Pillar Skirting (5'-10" × 7 Nos.) | 40.81 Rft |
| VIP Seena | 7.00 Rft |
| Granite Dell | 4.75 Rft |
| Kota Dell (4'0" × 2 Nos.) | 8.00 Rft |
| Fire Seena Skirting Extra | 50.00 Rft |
| **Total Skirting** | **182.00 Rft @ ₹40.00 = ₹7,280.00** |

## Requirements

- Python 3.8+
- `reportlab` — for PDF generation
- `python-docx` — for DOCX generation
- `lxml` — required by python-docx for XML operations

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
cd generate_bill
python generate_bill.py
```

This will create `Bill_Final.pdf` and `Bill_Final.docx` in the same folder.

## Project Structure

```
generate_bill/
├── generate_bill.py   # Main script — generates PDF and DOCX
├── requirements.txt   # Python dependencies
├── README.md          # This file
├── Bill_Final.pdf     # Generated PDF (created after running script)
└── Bill_Final.docx    # Generated DOCX (created after running script)
```

## Notes

- Bill No: 01 | Date: 09/04/2026
- All amounts formatted in Indian number format (e.g., ₹1,86,970.00)
- Grand Total in words: *Rupees One Lakh Eighty-Six Thousand Nine Hundred Seventy Only*
- The document is presented as the original final bill — no "updated" or "revised" markings
