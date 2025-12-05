import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from datetime import datetime
from num2words import num2words
import io

# Page configuration for mobile-friendly design
st.set_page_config(
    page_title="Generator Faktur",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for mobile-friendly design
st.markdown("""
    <style>
        .main > div {
            padding-top: 2rem;
        }
        .stNumberInput > div > div > input {
            font-size: 16px;
        }
        .stSelectbox > div > div > select {
            font-size: 16px;
        }
        .stTextInput > div > div > input {
            font-size: 16px;
        }
        .stDateInput > div > div > input {
            font-size: 16px;
        }
        @media (max-width: 768px) {
            .main > div {
                padding-left: 1rem;
                padding-right: 1rem;
            }
        }
    </style>
""", unsafe_allow_html=True)

# Hardcoded data
SELLER_NAME = "Praktyka Lekarska „Salutem" Iryna Berehova"
SELLER_ADDRESS = "ul. Okręg Wieleński 4a/1, 64-410 Sieraków"
SELLER_NIP = "7882010121"
SELLER_REGON = "388783174"

BUYER_NAME = "SAMODZIELNY PUBLICZNY ZAKŁAD OPIEKI ZDROWOTNEJ W MIĘDZYCHODZIE"
BUYER_ADDRESS = "UL. SZPITALNA 10, 64-400 MIĘDZYCHÓD"
BUYER_NIP = "5951340382"

BANK_ACCOUNT = "76124065531111001128223126"
PAYMENT_METHOD = "przelew"
PAYMENT_TERM = "wg umowy"

NOTES = "SPRZEDAWCA NIE JEST PŁATNIKIEM PODATKU VAT. Usługi zwolnione na podstawie art. 43 ust 1 pkt. 18 Ustawy o podatku od towaru i usług (VAT)."

MONTHS = [
    "Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
    "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"
]

def create_invoice_excel(invoice_no, date, month, year, hours, rate, total, total_words):
    """Create Excel invoice file"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Faktura"
    
    # Define styles
    bold_font = Font(bold=True, size=11)
    regular_font = Font(size=10)
    border_style = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal='center', vertical='center')
    left_alignment = Alignment(horizontal='left', vertical='center')
    right_alignment = Alignment(horizontal='right', vertical='center')
    
    # Set column widths
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 20
    
    row = 1
    
    # Title
    ws.merge_cells(f'A{row}:E{row}')
    ws[f'A{row}'] = "FAKTURA"
    ws[f'A{row}'].font = Font(bold=True, size=16)
    ws[f'A{row}'].alignment = center_alignment
    row += 2
    
    # Invoice number and date
    ws[f'A{row}'] = "Nr faktury:"
    ws[f'A{row}'].font = bold_font
    ws[f'B{row}'] = invoice_no
    ws[f'D{row}'] = "Data wystawienia:"
    ws[f'D{row}'].font = bold_font
    ws[f'E{row}'] = date.strftime("%d.%m.%Y")
    row += 1
    
    # Month and year
    ws[f'A{row}'] = "Miesiąc:"
    ws[f'A{row}'].font = bold_font
    ws[f'B{row}'] = f"{month} {year}"
    row += 2
    
    # Seller section
    ws[f'A{row}'] = "SPRZEDAWCA:"
    ws[f'A{row}'].font = bold_font
    row += 1
    ws[f'A{row}'] = SELLER_NAME
    ws[f'A{row}'].font = regular_font
    row += 1
    ws[f'A{row}'] = SELLER_ADDRESS
    ws[f'A{row}'].font = regular_font
    row += 1
    ws[f'A{row}'] = f"NIP: {SELLER_NIP}"
    ws[f'A{row}'].font = regular_font
    row += 1
    ws[f'A{row}'] = f"REGON: {SELLER_REGON}"
    ws[f'A{row}'].font = regular_font
    row += 2
    
    # Buyer section
    ws[f'A{row}'] = "NABYWCA:"
    ws[f'A{row}'].font = bold_font
    row += 1
    ws[f'A{row}'] = BUYER_NAME
    ws[f'A{row}'].font = regular_font
    row += 1
    ws[f'A{row}'] = BUYER_ADDRESS
    ws[f'A{row}'].font = regular_font
    row += 1
    ws[f'A{row}'] = f"NIP: {BUYER_NIP}"
    ws[f'A{row}'].font = regular_font
    row += 2
    
    # Table header
    ws.merge_cells(f'A{row}:E{row}')
    ws[f'A{row}'] = "POZYCJE FAKTURY"
    ws[f'A{row}'].font = bold_font
    ws[f'A{row}'].alignment = center_alignment
    row += 1
    
    # Table columns
    headers = ["Lp.", "Nazwa usługi", "Ilość (godz.)", "Cena jednostkowa", "Wartość"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_idx)
        cell.value = header
        cell.font = bold_font
        cell.alignment = center_alignment
        cell.border = border_style
    row += 1
    
    # Table data
    ws.cell(row=row, column=1).value = "1"
    ws.cell(row=row, column=1).alignment = center_alignment
    ws.cell(row=row, column=1).border = border_style
    
    ws.cell(row=row, column=2).value = f"Usługi medyczne - {month} {year}"
    ws.cell(row=row, column=2).alignment = left_alignment
    ws.cell(row=row, column=2).border = border_style
    
    ws.cell(row=row, column=3).value = hours
    ws.cell(row=row, column=3).alignment = center_alignment
    ws.cell(row=row, column=3).border = border_style
    
    ws.cell(row=row, column=4).value = f"{rate:.2f} zł"
    ws.cell(row=row, column=4).alignment = right_alignment
    ws.cell(row=row, column=4).border = border_style
    
    ws.cell(row=row, column=5).value = f"{total:.2f} zł"
    ws.cell(row=row, column=5).alignment = right_alignment
    ws.cell(row=row, column=5).border = border_style
    row += 2
    
    # Total section
    ws[f'D{row}'] = "RAZEM:"
    ws[f'D{row}'].font = bold_font
    ws[f'D{row}'].alignment = right_alignment
    ws[f'E{row}'] = f"{total:.2f} zł"
    ws[f'E{row}'].font = bold_font
    ws[f'E{row}'].alignment = right_alignment
    row += 1
    
    # Total in words
    ws.merge_cells(f'A{row}:E{row}')
    ws[f'A{row}'] = f"Słownie: {total_words}"
    ws[f'A{row}'].font = regular_font
    ws[f'A{row}'].alignment = left_alignment
    row += 2
    
    # Payment information
    ws[f'A{row}'] = "Forma płatności:"
    ws[f'A{row}'].font = bold_font
    ws[f'B{row}'] = PAYMENT_METHOD
    ws[f'B{row}'].font = regular_font
    row += 1
    
    ws[f'A{row}'] = "Termin płatności:"
    ws[f'A{row}'].font = bold_font
    ws[f'B{row}'] = PAYMENT_TERM
    ws[f'B{row}'].font = regular_font
    row += 1
    
    ws[f'A{row}'] = "Numer konta:"
    ws[f'A{row}'].font = bold_font
    ws[f'B{row}'] = BANK_ACCOUNT
    ws[f'B{row}'].font = regular_font
    row += 2
    
    # Notes
    ws.merge_cells(f'A{row}:E{row}')
    ws[f'A{row}'] = NOTES
    ws[f'A{row}'].font = Font(size=9, italic=True)
    ws[f'A{row}'].alignment = left_alignment
    
    return wb

def main():
    st.title("📄 Generator Faktur")
    st.markdown("---")
    
    # User inputs in columns for better mobile layout
    col1, col2 = st.columns(2)
    
    with col1:
        invoice_no = st.text_input("Nr faktury", value="10/2025")
        date = st.date_input("Data", value=datetime.now().date())
        month = st.selectbox("Miesiąc", MONTHS, index=0)
    
    with col2:
        year = st.number_input("Rok", min_value=2000, max_value=2100, value=2025, step=1)
        hours = st.number_input("Ilość godzin", min_value=0.0, value=0.0, step=0.5, format="%.1f")
        rate = st.number_input("Stawka (zł)", min_value=0.0, value=170.0, step=1.0, format="%.2f")
    
    st.markdown("---")
    
    # Calculate total
    total = hours * rate
    
    # Display calculation
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Ilość godzin", f"{hours:.1f}")
    with col2:
        st.metric("Stawka", f"{rate:.2f} zł")
    with col3:
        st.metric("RAZEM", f"{total:.2f} zł", delta=None)
    
    # Convert to Polish words
    if total > 0:
        try:
            total_words = num2words(total, lang='pl', to='currency', currency='PLN')
            st.info(f"**Słownie:** {total_words}")
        except Exception as e:
            total_words = f"{total:.2f} zł"
            st.warning(f"Nie udało się przekonwertować na słowa: {e}")
    else:
        total_words = "zero złotych zero groszy"
    
    st.markdown("---")
    
    # Generate Excel button
    if st.button("📥 Generuj Excel", type="primary", use_container_width=True):
        if total <= 0:
            st.error("Proszę wprowadzić prawidłową ilość godzin i stawkę.")
        else:
            try:
                wb = create_invoice_excel(invoice_no, date, month, year, hours, rate, total, total_words)
                
                # Save to bytes
                buffer = io.BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                
                # Download button
                filename = f"Faktura_{invoice_no.replace('/', '_')}.xlsx"
                st.download_button(
                    label="⬇️ Pobierz plik Excel",
                    data=buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.success(f"✅ Faktura wygenerowana pomyślnie: {filename}")
            except Exception as e:
                st.error(f"Błąd podczas generowania faktury: {e}")

if __name__ == "__main__":
    main()

