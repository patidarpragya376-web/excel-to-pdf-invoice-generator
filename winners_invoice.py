
import os
import pandas as pd

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm


# ======================================
# 1. PROJECT SETTINGS
# ======================================

EXCEL_FILE = "invoice_data.xlsx"
OUTPUT_FOLDER = "Generated_Invoices"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Business details from your sample.
# Edit these if any details need changing.
SHOP_NAME = "WINNERS CREEK"
ADDRESS = "Capt R. A. Alexander House, Near Care Bear School"
CITY = "Agra Road, Jaipur - 302003 (Rajasthan)"
EMAIL = "winnerscreek@gmail.com"
PHONE = "9314099002"
GSTIN = "08AEQPC9460E2ZA"

BANK_ACCOUNT = "627801010051029"
IFSC = "UBIN0562785"
MICR = "302026501"


# ======================================
# 2. HELPER FUNCTIONS
# ======================================

def clean(value):
    """Convert an Excel value to printable text."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def draw_text(pdf, x, y, text, size=10, bold=False):
    """Draw text using coordinates measured from the top."""
    font = "Helvetica-Bold" if bold else "Helvetica"
    pdf.setFont(font, size)
    pdf.drawString(x * mm, A4[1] - y * mm, str(text))


def draw_line(pdf, x1, y1, x2, y2):
    """Draw a line using coordinates measured from the top."""
    pdf.setStrokeColor(colors.black)
    pdf.setLineWidth(0.7)
    pdf.line(
        x1 * mm, A4[1] - y1 * mm,
        x2 * mm, A4[1] - y2 * mm
    )


def draw_box(pdf, x, y, width, height):
    """Draw a rectangular box."""
    pdf.setStrokeColor(colors.black)
    pdf.setLineWidth(0.7)
    pdf.rect(
        x * mm,
        A4[1] - (y + height) * mm,
        width * mm,
        height * mm,
        fill=0,
        stroke=1
    )


# ======================================
# 3. CREATE ONE INVOICE PDF
# ======================================

def create_invoice(row):

    invoice_no = clean(row["Invoice_No"])
    date = clean(row["Date"])
    customer = clean(row["Customer_Name"])
    customer_details = clean(row["Customer_Details"])
    particulars = clean(row["Particulars"])
    session = clean(row["Session_Year"])
    amount = float(row["Amount"])
    amount_words = clean(row["Amount_In_Words"])

    # Make a safe filename.
    safe_number = "".join(
        ch for ch in invoice_no
        if ch.isalnum() or ch in "-_"
    )

    if not safe_number:
        safe_number = "Unknown"

    filename = os.path.join(
        OUTPUT_FOLDER,
        f"Invoice_{safe_number}.pdf"
    )

    pdf = canvas.Canvas(filename, pagesize=A4)
    pdf.setTitle(f"Invoice {invoice_no}")

    page_width, page_height = A4

    # ==================================
    # 4. BUSINESS HEADER
    # ==================================

    # Logo placeholder
    pdf.setLineWidth(1.5)
    pdf.circle(
        35 * mm,
        page_height - 35 * mm,
        13 * mm
    )

    pdf.setFont("Helvetica-Bold", 15)
    pdf.drawCentredString(
        35 * mm,
        page_height - 37 * mm,
        "WC"
    )

    # Business name and details
    pdf.setFont("Helvetica-Bold", 23)
    pdf.drawCentredString(
        125 * mm,
        page_height - 24 * mm,
        SHOP_NAME
    )

    pdf.setFont("Helvetica", 8.5)
    pdf.drawCentredString(
        125 * mm, page_height - 31 * mm, ADDRESS
    )
    pdf.drawCentredString(
        125 * mm, page_height - 36 * mm, CITY
    )
    pdf.drawCentredString(
        125 * mm, page_height - 41 * mm,
        f"Email: {EMAIL}  |  Phone: {PHONE}"
    )
    pdf.drawCentredString(
        125 * mm, page_height - 46 * mm,
        f"GSTIN: {GSTIN}"
    )

    draw_line(pdf, 12, 54, 198, 54)

    # ==================================
    # 5. CUSTOMER AND INVOICE DETAILS
    # ==================================

    draw_box(pdf, 12, 54, 116, 35)
    draw_box(pdf, 128, 54, 70, 35)

    draw_text(pdf, 16, 61, "To,", 11, True)
    draw_text(pdf, 25, 69, customer, 11, True)

    # Split additional details into two lines.
    detail_lines = customer_details.split(";")
    draw_text(
        pdf, 25, 76,
        detail_lines[0].strip() if detail_lines else "",
        9
    )

    if len(detail_lines) > 1:
        draw_text(pdf, 25, 83, detail_lines[1].strip(), 9)

    draw_text(pdf, 131, 63, "Invoice No.:", 10, True)
    draw_text(pdf, 165, 63, invoice_no, 11, True)

    draw_line(pdf, 128, 71, 198, 71)

    draw_text(pdf, 131, 80, "Date:", 10, True)
    draw_text(pdf, 150, 80, date, 10)

    # ==================================
    # 6. PARTICULARS TABLE
    # ==================================

    draw_box(pdf, 12, 89, 186, 12)

    # Column boundaries
    draw_line(pdf, 32, 89, 32, 222)
    draw_line(pdf, 150, 89, 150, 222)

    pdf.setFillColor(colors.lightgrey)
    pdf.rect(
        12 * mm,
        page_height - 101 * mm,
        186 * mm,
        12 * mm,
        fill=1,
        stroke=0
    )

    # Redraw table header border
    draw_box(pdf, 12, 89, 186, 12)

    pdf.setFillColor(colors.black)
    draw_text(pdf, 16, 97, "S.No.", 10, True)
    draw_text(pdf, 75, 97, "Particulars", 11, True)
    draw_text(pdf, 158, 97, "Amount (Rs.)", 10, True)

    # Body
    draw_box(pdf, 12, 101, 186, 121)

    draw_text(pdf, 18, 119, "1", 10)
    draw_text(pdf, 38, 119, particulars, 12)
    draw_text(pdf, 38, 141, session, 12)

    draw_text(pdf, 157, 119, f"{amount:.2f}", 11, True)

    # ==================================
    # 7. TOTAL AMOUNT
    # ==================================

    draw_box(pdf, 12, 222, 186, 11)
    draw_text(pdf, 112, 229, "TOTAL AMOUNT", 10, True)
    draw_text(pdf, 157, 229, f"{amount:.2f}", 11, True)

    # ==================================
    # 8. DECLARATION
    # ==================================

    draw_box(pdf, 12, 233, 186, 9)
    draw_text(
        pdf, 15, 239,
        "Declaration: Composition Taxable person, "
        "not eligible to collect tax on supplies.",
        8.5, True
    )

    # ==================================
    # 9. AMOUNT IN WORDS
    # ==================================

    draw_box(pdf, 12, 242, 186, 20)
    draw_text(
        pdf, 15, 249,
        "Total Invoice Amount (in words):",
        9, True
    )
    draw_text(pdf, 15, 256, amount_words, 10)

    # ==================================
    # 10. BANK DETAILS AND SIGNATURES
    # ==================================

    draw_box(pdf, 12, 262, 72, 25)
    draw_box(pdf, 84, 262, 55, 25)
    draw_box(pdf, 139, 262, 59, 25)

    draw_text(pdf, 15, 268, "Bank Details:", 9, True)
    draw_text(pdf, 15, 273, f"Bank A/C: {BANK_ACCOUNT}", 6.8)
    draw_text(pdf, 15, 278, f"IFSC: {IFSC}", 7.5)
    draw_text(pdf, 15, 283, f"MICR: {MICR}", 7.5)

    draw_text(pdf, 87, 269, "Prepared by", 9, True)
    draw_line(pdf, 88, 278, 135, 278)

    draw_text(pdf, 142, 268, "Certified correct", 8, True)
    draw_text(pdf, 142, 273, f"for {SHOP_NAME}", 7.5, True)
    draw_line(pdf, 143, 278, 194, 278)
    draw_text(pdf, 143, 284, "Authorised Signatory", 7.5, True)

    # Save the PDF
    pdf.save()

    print(f"Invoice created: {filename}")


# ======================================
# 11. READ EXCEL AND GENERATE INVOICES
# ======================================

def main():

    if not os.path.exists(EXCEL_FILE):
        print(f"Excel file not found: {EXCEL_FILE}")
        print("Please create the Excel file in this project folder.")
        return

    try:
        data = pd.read_excel(EXCEL_FILE)

        required_columns = [
            "Invoice_No",
            "Date",
            "Customer_Name",
            "Customer_Details",
            "Particulars",
            "Session_Year",
            "Amount",
            "Amount_In_Words"
        ]

        missing = [
            column for column in required_columns
            if column not in data.columns
        ]

        if missing:
            print("Missing Excel columns:", missing)
            return

        data = data.dropna(how="all")

        if data.empty:
            print("Please enter invoice details in Excel.")
            return

        for _, row in data.iterrows():
            create_invoice(row)

        print("\nAll invoices have been generated!")
        print("Open the Generated_Invoices folder.")

    except Exception as error:
        print("An error occurred:", error)


if __name__ == "__main__":
    main()