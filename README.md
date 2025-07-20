# SBT UPI Card Generator

This application generates printable payment cards (with QR codes) for sales transactions, based on data from an Excel file. Each card includes shop details, bill information, a UPI QR code, and is formatted for easy printing and cutting. The tool provides a simple graphical interface for selecting the input file and generating a PDF of cards.

## Features
- Reads sales data from an Excel (.xlsx) file
- Generates cards with:
  - Shop name, mobile, bill number, date, and amount
  - UPI QR code for payment (customizable UPI ID)
  - Company logo and branding
  - Receiver signature area
- Outputs a printable PDF with multiple cards per page (6 per A4 sheet)
- Simple GUI for file selection and PDF naming

## Requirements
- Python 3.7+
- The following Python packages:
  - pandas
  - pillow
  - qrcode
  - fpdf
  - tkinter (usually included with Python)

Install dependencies with:
```bash
pip install pandas pillow qrcode fpdf
```

## Files
- `card.py` — Main application script
- `logo.png` — Logo image used on each card (replace with your own if needed)
- `sales list.xlsx` — Example Excel file (structure: columns like Party, Mobile, Bill No, Date, Credit, S.Man, Group)

## Usage
1. Ensure all dependencies are installed and the required files (`card.py`, `logo.png`, and your Excel file) are in the same directory.
2. Run the application:
   ```bash
   python card.py
   ```
3. In the GUI:
   - Enter a name for the output PDF (optional; defaults to `output_cards.pdf`)
   - Click "Select Excel & Generate"
   - Choose your Excel file with sales data
   - The PDF will be generated in the current directory

## Excel File Format
Your Excel file should have columns similar to:
- `Party` (Shop Name)
- `Mobile` (Shop Mobile)
- `Bill No`
- `Date`
- `Credit` (Bill Amount)
- `S.Man` (Salesman)
- `Group` (Area/Group)

## Sample Excel File

Below is a sample of the provided `sales list.xlsx` file. Only the relevant columns are used by the application:

| Sno | Bill No | Date       | Party      | Group   | S.Man   | Credit   | Mobile      |
|-----|---------|------------|------------|---------|---------|----------|-------------|
| 1   | 13614   | 17/07/2025 | (example)  | (group) | (name)  | 1000.00  | 9876543210  |
| 2   | 13615   | 17/07/2025 | (example2) | (group) | (name2) | 2000.00  | 9123456780  |
| 3   | 13616   | 17/07/2025 | (example3) | (group) | (name3) | 1500.00  | 9988776655  |

**Note:** The actual file contains many more columns, but only the above are required for card generation. Make sure your Excel file includes at least these columns: `Party`, `Mobile`, `Bill No`, `Date`, `Credit`, `S.Man`, and `Group`.

## Customization
- **UPI ID:** Change the `