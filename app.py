import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from datetime import datetime
from num2words import num2words
import io
import os
import tempfile
import pandas as pd

# Page configuration for mobile-friendly design
st.set_page_config(
    page_title="Generator Faktur",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for modern mobile-friendly design
st.markdown("""
    <style>
        /* Основные настройки */
        .main > div {
            padding-top: 1rem;
            max-width: 100%;
        }
        
        /* Крупные поля ввода для мобильных */
        .stNumberInput > div > div > input,
        .stTextInput > div > div > input,
        .stSelectbox > div > div > select {
            font-size: 18px !important;
            padding: 12px !important;
            min-height: 48px !important;
        }
        
        /* Крупные кнопки */
        .stButton > button {
            font-size: 18px !important;
            padding: 14px 24px !important;
            min-height: 52px !important;
            border-radius: 12px !important;
            font-weight: 600 !important;
        }
        
        /* Карточки для клиник */
        .clinic-card {
            padding: 20px;
            border-radius: 16px;
            border: 3px solid #e0e0e0;
            margin: 10px 0;
            cursor: pointer;
            transition: all 0.3s ease;
            background: white;
        }
        .clinic-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        }
        .clinic-card.selected {
            border-color: #1f77b4;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            box-shadow: 0 4px 20px rgba(102, 126, 234, 0.4);
        }
        .clinic-card h3 {
            margin: 0 0 8px 0;
            font-size: 20px;
            font-weight: 700;
        }
        .clinic-card p {
            margin: 0;
            font-size: 14px;
            opacity: 0.9;
        }
        
        /* Секции */
        .section {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 12px;
            margin: 20px 0;
        }
        
        /* Метрики */
        [data-testid="stMetricValue"] {
            font-size: 28px !important;
            font-weight: 700 !important;
        }
        
        /* Мобильная оптимизация */
        @media (max-width: 768px) {
            .main > div {
                padding-left: 0.5rem;
                padding-right: 0.5rem;
            }
            .stButton > button {
                width: 100% !important;
            }
            h1 {
                font-size: 24px !important;
            }
        }
        
        /* Скрыть лишнее */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# Hardcoded data - Seller
SELLER_NAME = 'Praktyka Lekarska „Salutem" Iryna Berehova'
SELLER_ADDRESS = "ul. Okręg Wieleński 4a/1, 64-410 Sieraków"
SELLER_NIP = "7882010121"
SELLER_REGON = "388783174"

BANK_ACCOUNT = "76124065531111001128223126"
PAYMENT_METHOD = "przelew"
PAYMENT_TERM = "wg umowy"

NOTES = "SPRZEDAWCA NIE JEST PŁATNIKIEM PODATKU VAT. Usługi zwolnione na podstawie art. 43 ust 1 pkt. 18 Ustawy o podatku od towaru i usług (VAT)."

# Клиники (Buyers)
CLINICS = {
    "miedzychod": {
        "name": "SAMODZIELNY PUBLICZNY ZAKŁAD OPIEKI ZDROWOTNEJ W MIĘDZYCHODZIE",
        "address_line1": "64-400 MIĘDZYCHÓD",
        "address_line2": "UL. SZPITALNA 10",
        "nip": "5951340382",
        "display_name": "Międzychód"
    },
    "limamed": {
        "name": 'Przychodnia Zespołu Lekarza Rodzinnego „Limamed"',
        "address_line1": "64-316 Kuślin",
        "address_line2": "Ul. Emilii Sczanieckiej 6",
        "nip": "7881731812",
        "display_name": "Limamed"
    }
}

MONTHS = [
    "Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
    "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"
]

def create_invoice_excel(invoice_no, date, month, year, hours, rate, total, total_words, buyer_data):
    """Load .xlsx template and replace only data values"""
    # Выбираем шаблон в зависимости от клиники
    if buyer_data.get("display_name") == "Międzychód":
        template_path = "/Users/teehoo/Documents/med/shablon/FakturaSPZOZ Międzychód 22^25.xlsx"
    else:
        template_path = "/Users/teehoo/Documents/med/shablon/Faktura Limamed 23^2025.xlsx"
    
    # Загружаем .xlsx шаблон напрямую через openpyxl (сохраняет все форматирование)
    wb = load_workbook(template_path)
    ws = wb.active
    
    # Меняем ТОЛЬКО данные в нужных ячейках (сохраняем форматирование из шаблона)
    # D9: номер фактуры
    ws['D9'].value = invoice_no
    
    # D16, D18: даты
    ws['D16'].value = date.strftime("%d.%m.%Y")
    ws['D18'].value = date.strftime("%d.%m.%Y")
    
    # C26: описание услуги
    month_name_lower = month.lower()
    ws['C26'].value = f"usługi medyczne lekarza POZ w miesiącu {month_name_lower} {year}"
    
    # E26, F26, G26: часы, ставка, сумма
    ws['E26'].value = int(hours) if hours == int(hours) else hours
    ws['F26'].value = int(rate) if rate == int(rate) else rate
    ws['G26'].value = int(total) if total == int(total) else total
    
    # G27, D28: итого (всегда заменяем на правильную сумму)
    ws['G27'].value = int(total) if total == int(total) else total
    ws['D28'].value = int(total) if total == int(total) else total
    
    # C31: сумма прописью
    ws['C31'].value = total_words
    
    return wb

def main():
    st.title("📄 Generator Faktur")
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Определяем выбранную клинику
    if 'selected_clinic' not in st.session_state:
        st.session_state.selected_clinic = 'miedzychod'
    
    # Выбор клиники - большие кнопки
    st.markdown("### 🏥 Wybierz klinikę")
    
    col1, col2 = st.columns(2)
    
    with col1:
        clinic1_selected = st.session_state.selected_clinic == 'miedzychod'
        btn_type_1 = "primary" if clinic1_selected else "secondary"
        if st.button(f"🏥\n\n{CLINICS['miedzychod']['display_name']}", key="clinic1", use_container_width=True, type=btn_type_1):
            st.session_state.selected_clinic = 'miedzychod'
            st.rerun()
    
    with col2:
        clinic2_selected = st.session_state.selected_clinic == 'limamed'
        btn_type_2 = "primary" if clinic2_selected else "secondary"
        if st.button(f"🏥\n\n{CLINICS['limamed']['display_name']}", key="clinic2", use_container_width=True, type=btn_type_2):
            st.session_state.selected_clinic = 'limamed'
            st.rerun()
    
    selected_clinic_data = CLINICS[st.session_state.selected_clinic]
    st.success(f"✅ Wybrano: **{selected_clinic_data['display_name']}**")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Секция: Данные фактуры
    st.markdown("### 📝 Dane faktury")
    
    # Одна колонка для мобильных
    month = st.selectbox("📅 Miesiąc", MONTHS, index=0, help="Wybierz miesiąc za który wystawiasz fakturę")
    
    invoice_no = st.text_input("🔢 Nr faktury", value="10/2025", help="Format: XX/YYYY (np. 10/2025)")
    
    # Автоматически определяем год
    current_year = datetime.now().year
    st.info(f"📆 **Rok:** {current_year} (ustawiany automatycznie)")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Секция: Usługi
    st.markdown("### 💼 Usługi")
    
    hours = st.number_input(
        "⏰ Ilość godzin", 
        min_value=0.0, 
        value=0.0, 
        step=0.5, 
        format="%.1f",
        help="Wprowadź liczbę przepracowanych godzin"
    )
    
    rate = st.number_input(
        "💰 Stawka (zł)", 
        min_value=0.0, 
        value=170.0, 
        step=1.0, 
        format="%.2f",
        help="Stawka za godzinę w złotych"
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Автоматически используем сегодняшнюю дату
    date = datetime.now().date()
    year = current_year
    
    # Calculate total
    total = hours * rate
    
    # Секция: Podsumowanie
    st.markdown("### 💵 Podsumowanie")
    
    # Display calculation - крупные метрики
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("⏰ Godziny", f"{hours:.1f}")
    with col2:
        st.metric("💰 Stawka", f"{rate:.2f} zł")
    with col3:
        st.metric("💵 RAZEM", f"{total:.2f} zł", delta=None)
    
    # Convert to Polish words
    if total > 0:
        try:
            total_words = num2words(total, lang='pl', to='currency', currency='PLN')
            st.success(f"📝 **Słownie:** {total_words}")
        except Exception as e:
            total_words = f"{total:.2f} zł"
            st.warning(f"Nie udało się przekonwertować na słowa: {e}")
    else:
        total_words = "zero złotych zero groszy"
        st.info("💡 Wprowadź ilość godzin i stawkę, aby zobaczyć podsumowanie")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Generate Excel button - большая кнопка
    if st.button("📥 Generuj fakturę Excel", type="primary", use_container_width=True):
        if total <= 0:
            st.error("❌ Proszę wprowadzić prawidłową ilość godzin i stawkę.")
        else:
            with st.spinner("⏳ Generowanie faktury..."):
                try:
                    wb = create_invoice_excel(invoice_no, date, month, year, hours, rate, total, total_words, selected_clinic_data)
                    
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
                        use_container_width=True,
                        type="primary"
                    )
                    
                    st.balloons()
                    st.success(f"✅ Faktura wygenerowana pomyślnie: **{filename}**")
                except Exception as e:
                    st.error(f"❌ Błąd podczas generowania faktury: {e}")
                    st.exception(e)

if __name__ == "__main__":
    main()

