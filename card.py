import sys
import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import qrcode
import qrcode.constants
from fpdf import FPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import tempfile

# === CONFIG ===
CARD_WIDTH = 260
CARD_HEIGHT = 260
CARDS_PER_PAGE = 6
CARDS_PER_ROW = 2
UPI_ID = "SRIBALAJITRADINGCOMPANY612@iob"
FONT_PATH = "arialbd.ttf"  # Use bold font

OUTPUT_DIR = "generated_cards"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def resource_path(relative_path):
    # Get absolute path to resource, works for dev and for PyInstaller
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)

def load_excel(path):
    df = pd.read_excel(path)
    df.columns = df.columns.str.strip()  # Remove extra spaces from column names
    return df

def get_font(size):
    try:
        return ImageFont.truetype(FONT_PATH, size)
    except:
        return ImageFont.load_default()

def generate_card(data, index):
    img = Image.new("RGB", (CARD_WIDTH, CARD_HEIGHT), "white")
    draw = ImageDraw.Draw(img)

    # Add logo to top left corner (portrait, visible size)
    try:
        logo = Image.open(resource_path("logo.png")).convert("RGBA")
        logo_size = (140, 140)
        logo.thumbnail(logo_size, Image.LANCZOS)
        img.paste(logo, (30, 30), logo)
        logo_right = 30 + logo.size[0]
        logo_bottom = 30 + logo.size[1]
    except Exception as e:
        print(f"Logo not added: {e}")
        logo_right = 30
        logo_bottom = 30

    font_huge = get_font(13)   # Title
    font_medium = get_font(8)  # Main fields
    font_small = get_font(6)   # Small text
    spacing = 12               # Space between fields

    # Alignment constants
    label_x = 60
    value_x = 420  # Increased for larger fonts and to avoid overlap
    spacing = 90   # Increased for larger fonts

    # Title (centered after logo, never overlapping logo)
    title = "SBT"
    bbox = draw.textbbox((0, 0), title, font=font_huge)
    title_w = bbox[2] - bbox[0]
    title_x = max(logo_right + 40, (CARD_WIDTH - title_w) // 2)
    title_y = 40
    draw.text((title_x, title_y), title, fill="black", font=font_huge)

    # Salesman name at top right, always within card
    salesman = str(data.get("S.Man", "")).strip()
    if not salesman or salesman.lower() == "nan":
        salesman = "Unknown"
    salesman_text = f"Salesman: {salesman}"
    salesman_bbox = draw.textbbox((0, 0), salesman_text, font=font_small)
    salesman_w = salesman_bbox[2] - salesman_bbox[0]
    salesman_x = min(CARD_WIDTH - salesman_w - 60, CARD_WIDTH - 350)
    draw.text((salesman_x, 60), salesman_text, fill="black", font=font_small)

    # Set starting y below logo/title
    y = max(logo_bottom + 40, title_y + bbox[3] + 40)

    # Bill Amount: use only the 'Credit' column, handle commas and spaces
    bill_amt = 0
    credit_val = data.get('Credit')
    if pd.notna(credit_val):
        try:
            bill_amt = float(str(credit_val).replace(',', '').strip())
        except:
            bill_amt = 0

    shop_name = str(data.get('Party', '')).strip()
    if shop_name.lower() == 'nan':
        shop_name = ''
    shop_mobile = str(data.get('Mobile', '')).strip()
    if shop_mobile.lower() == 'nan':
        shop_mobile = ''

    # Shop Name
    draw.text((label_x, y), "Shop Name     :", fill="black", font=font_medium)
    draw.text((value_x, y), shop_name, fill="black", font=font_medium)

    # Shop Mobile
    draw.text((label_x, y + spacing), "Shop Mobile           :", fill="black", font=font_medium)
    draw.text((value_x, y + spacing), shop_mobile, fill="black", font=font_medium)

    # Bill No
    draw.text((label_x, y + spacing * 2), "Bill No        :", fill="black", font=font_medium)
    draw.text((value_x, y + spacing * 2), str(data.get('Bill No', '')), fill="black", font=font_medium)

    # Date
    draw.text((label_x, y + spacing * 3), "Date               :", fill="black", font=font_medium)
    draw.text((value_x, y + spacing * 3), str(data.get('Date', ''))[:10], fill="black", font=font_medium)

    bill_amt_font = get_font(52)  # Larger for amount
    label_w = draw.textbbox((0, 0), "Received Amount:", font=font_medium)[2]
    bill_amt_label = "Bill Amount   :"
    received_amt_label = "Received Amount:"
    y_bill = y + spacing * 4
    y_received = y + spacing * 5

    # Bill Amount
    draw.text((label_x, y_bill), bill_amt_label, fill="black", font=font_medium)
    draw.text((value_x, y_bill), f"₹{bill_amt:.2f}", fill="#006400", font=bill_amt_font)

    # Add extra space before Received Amount for writing
    y_received = y_bill + spacing + 40  # 40px extra space

    # Received Amount (no line, only text)
    draw.text((label_x, y_received), received_amt_label, fill="black", font=font_medium)
    # (No line after the value)

    signature_y = y + spacing * 7
    draw.text((label_x, signature_y), "Receiver Signature:", fill="black", font=font_small)
    draw.line((value_x, signature_y + 30, value_x + 320, signature_y + 30), fill="black", width=2)

    upi_y = signature_y + 90
    draw.text((label_x, upi_y), f"UPI ID: {UPI_ID}", fill="black", font=font_small)

    # QR code (smaller size for 260x260 card)
    qr_link = f"upi://pay?pa={UPI_ID}&pn=SBT&am={bill_amt:.2f}&cu=INR"
    qr = qrcode.make(qr_link).resize((60, 60))
    qr_x = CARD_WIDTH - 70  # 10px margin from right
    qr_y = CARD_HEIGHT - 70  # 10px margin from bottom
    img.paste(qr, (qr_x, qr_y))

    path = os.path.join(OUTPUT_DIR, f"card_{index}.jpg")
    img.save(path, format="JPEG")
    return path

def generate_pdf_from_text(data_rows, output_pdf="output_cards.pdf"):
    pdf = FPDF("P", "pt", "A4")
    pdf.set_auto_page_break(False)
    page_width, page_height = 595, 842

    card_w, card_h = 260, 260
    cards_per_row = 2
    cards_per_col = 3
    x_margin = (page_width - (cards_per_row * card_w)) // (cards_per_row + 1)
    y_margin = (page_height - (cards_per_col * card_h)) // (cards_per_col + 1)

    for i in range(0, len(data_rows), 6):
        pdf.add_page()
        batch = data_rows[i:i+6]
        for j, row in enumerate(batch):
            col = j % cards_per_row
            r = j // cards_per_row
            x = x_margin + col * (card_w + x_margin)
            y = y_margin + r * (card_h + y_margin)
            # Card border
            pdf.set_draw_color(180, 180, 180)
            pdf.rect(x, y, card_w, card_h)
            # Logo (top left, larger and closer to top, use resource_path)
            logo_path = resource_path("logo.png")
            if os.path.exists(logo_path):
                pdf.image(logo_path, x + 8, y + 2, 70, 70)
            # Area name (from 'Group' column) above SBT, bold
            area_name = str(row.get('Group', ''))
            pdf.set_xy(x, y + 2)
            pdf.set_font("Arial", "B", 12)
            pdf.cell(card_w, 12, area_name, align="C")
            # SBT centered at top
            pdf.set_xy(x, y + 18)
            pdf.set_font("Arial", "B", 18)
            pdf.cell(card_w, 20, "SBT", align="C")
            # Add more vertical space for all fields
            y_text = y + 55
            line_height = 18
            left_margin = x + 15
            text_width = card_w - 20
            pdf.set_xy(left_margin, y_text)
            # Shop Name (left-aligned, label in black bold, value in dark blue bold)
            pdf.set_xy(left_margin, y_text)
            pdf.set_font("Arial", "B", 10)
            label = "Shop Name: "
            value = str(row.get('Party',''))
            label_w = pdf.get_string_width(label)
            pdf.set_text_color(0, 0, 0)  # Black for label
            pdf.cell(label_w, line_height, label, border=0)
            pdf.set_text_color(0, 0, 139)  # Dark blue for value
            pdf.cell(0, line_height, value, ln=1)
            pdf.set_text_color(0, 0, 0)  # Reset to black
            y_text = pdf.get_y() + 8
            # Bill No
            pdf.set_xy(left_margin, y_text)
            pdf.multi_cell(text_width, line_height, f"Bill No: {row.get('Bill No','')}")
            y_text = pdf.get_y() + 6
            # Date
            pdf.set_xy(left_margin, y_text)
            pdf.multi_cell(text_width, line_height, f"Date: {str(row.get('Date',''))[:10]}")
            y_text = pdf.get_y() + 6
            # Bill Amount: label in black, value in green and bold, with minimal gap
            label = "Bill Amount:"
            pdf.set_xy(left_margin, y_text)
            pdf.set_font("Arial", "B", 12)
            label_w = pdf.get_string_width(label)
            gap = 8
            pdf.set_text_color(0, 0, 0)
            pdf.cell(label_w, line_height, label, border=0)
            pdf.set_text_color(0, 128, 0)
            pdf.set_x(left_margin + label_w + gap)
            pdf.cell(0, line_height, f"Rs.{row.get('Credit','')}", ln=1)
            pdf.set_text_color(0, 0, 0)
            y_text = pdf.get_y() + 6
            # Received Amount
            pdf.set_xy(left_margin, y_text)
            pdf.multi_cell(text_width, line_height, "Received Amount:")
            y_text = pdf.get_y() + 6
            # Sign
            pdf.set_xy(left_margin, y_text)
            pdf.multi_cell(text_width, line_height, "Sign: _____________")
            y_text = pdf.get_y() + 6

            # QR code (bottom right, larger size, GPay compatible, high error correction)
            qr_link = f"upi://pay?pa=SRIBALAJITRADINGCOMPANY612@iob&pn=SBT"
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_H,
                box_size=10,
                border=2,
            )
            qr.add_data(qr_link)
            qr.make(fit=True)
            qr_img = qr.make_image(fill_color="black", back_color="white")
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_qr:
                qr_img = qr_img.resize((80, 80))
                qr_img.save(tmp_qr.name)
                qr_y_bottom = y + card_h - 90  # 10px margin from bottom
                pdf.image(tmp_qr.name, x + card_w - 90, qr_y_bottom, 80, 80)
            os.unlink(tmp_qr.name)
        # Draw cutting lines for easier cutting
        for c in range(1, cards_per_row):
            line_x = x_margin + c * (card_w + x_margin) - x_margin // 2
            pdf.set_draw_color(0, 0, 0)
            pdf.set_line_width(1)
            pdf.line(line_x, 0, line_x, page_height)
        for r in range(1, cards_per_col):
            line_y = y_margin + r * (card_h + y_margin) - y_margin // 2
            pdf.set_draw_color(0, 0, 0)
            pdf.set_line_width(1)
            pdf.line(0, line_y, page_width, line_y)


    output_path = os.path.join(os.getcwd(), output_pdf)
    pdf.output(output_path)
    print("✅ PDF saved at:", output_path)

def run_app():
    def select_file():
        filepath = filedialog.askopenfilename(filetypes=[["Excel files", "*.xlsx"]])
        if not filepath:
            return
        try:
            pdf_name = pdf_entry.get().strip()
            if not pdf_name:
                pdf_name = "output_cards.pdf"
            elif not pdf_name.lower().endswith(".pdf"):
                pdf_name += ".pdf"

            df = load_excel(filepath)
            data_rows = df.to_dict(orient="records")
            generate_pdf_from_text(data_rows, pdf_name)
            messagebox.showinfo("Success", f"PDF generated: {pdf_name} with {len(data_rows)} cards.")
        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n{str(e)}")

    root = tk.Tk()
    root.title("SBT Card Generator")
    root.geometry("400x250")

    tk.Label(root, text="SBT UPI Card Generator", font=("Arial", 16)).pack(pady=10)
    tk.Label(root, text="Enter output PDF name:", font=("Arial", 12)).pack()
    pdf_entry = tk.Entry(root, font=("Arial", 12))
    pdf_entry.pack(pady=5)
    pdf_entry.insert(0, "output_cards")

    tk.Button(root, text="Select Excel & Generate", command=select_file, bg="#007BFF", fg="white", height=2).pack(pady=20)

    root.mainloop()

# Ensure Tkinter GUI is used for file selection and PDF generation
if __name__ == "__main__":
    run_app()





