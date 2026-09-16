from datetime import datetime, timedelta
import io
import math
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import pulp
import streamlit as st
from PIL import Image

# --- KONFIGURACJA STRONY STREAMLIT ---
st.set_page_config(
    page_title="żabka jush! - Generator Szkieletu Grafiku DS",
    page_icon="⚡",
    layout="wide",
)

# Bezpieczne importy bibliotek
try:
    import openpyxl
except ImportError:
    st.error("❌ Brakuje biblioteki 'openpyxl'. Upewnij się, że znajduje się w requirements.txt!")

try:
    import pulp
except ImportError:
    st.error("❌ Brakuje biblioteki 'pulp'. Upewnij się, że znajduje się w requirements.txt!")

# --- STYLIZACJA W PALECIE JUSH! ---
st.markdown(
    """
    <style>
    :root {
        --jush-lime: #8BC53F;
        --jush-dark-green: #005B2B;
        --jush-light-lime: #EBF7D4;
    }
    
    .stApp {
        background-color: #FAFCF5;
    }
    
    [data-testid="stSidebar"] {
        background-color: #8BC53F !important;
    }
    [data-testid="stSidebar"] * {
        color: #005B2B !important;
        font-weight: bold !important;
    }
    
    div.stButton > button {
        background-color: #005B2B !important;
        color: #8BC53F !important;
        font-weight: 800 !important;
        font-size: 16px !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 10px 24px !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button:hover {
        background-color: #004420 !important;
        color: #A3DF52 !important;
    }
    
    h1, h2, h3 {
        color: #005B2B !important;
        font-family: 'Arial Black', sans-serif !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# --- BRANDING JUSH! - BANNER NAGŁÓWKA ---
col_logo, col_title = st.columns([1, 4])
with col_logo:
    st.image(
        "https://zabkagroup.com/wp-content/uploads/2022/09/Jush_logo.png",
        width=140,
    )
with col_title:
    st.markdown(
        "<h1 style='margin-bottom:0; font-size: 2.6rem;'>żabka <span style='color:#005B2B;'>jush!</span></h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='font-weight:bold; color:#005B2B; font-size: 1.1rem;'>Generator Szkieletu Zmian z Wyrównywaniem Czasów Zmian</p>",
        unsafe_allow_html=True,
    )

st.divider()

# --- SIDEBAR: PARAMETRY EFEKTYWNOŚCI I OBSADY ---
st.sidebar.image(
    "https://zabkagroup.com/wp-content/uploads/2022/09/Jush_logo.png", width=110
)
st.sidebar.header("⚙️ Ustawienia Magazynu DS")

typ_magazynu = st.sidebar.selectbox("Typ magazynu", ["Standardowy", "Nocny"])
is_nocny = typ_magazynu == "Nocny"

godzina_otwarcia_ds = 6.0
godzina_zamkniecia_ds = 25.5 if is_nocny else 23.5

cel_efektywnosci = st.sidebar.number_input(
    "Możliwości pakowania zamówień (zamówienia / h / picker)",
    min_value=1,
    value=15,
)

MAPA_DNI = {
    "Monday": "Poniedziałek",
    "Tuesday": "Wtorek",
    "Wednesday": "Środa",
    "Thursday": "Czwartek",
    "Friday": "Piątek",
    "Saturday": "Sobota",
    "Sunday": "Niedziela",
}

# --- 1. ZAKRES DAT DLA GRAFIKU ---
st.header("1. Wybierz okres grafiku")
okres_grafiku = st.date_input(
    "Wskaż zakres od - do:",
    value=(datetime.now().date(), datetime.now().date() + timedelta(days=29)),
)

if isinstance(okres_grafiku, tuple) and len(okres_grafiku) == 2:
    start_date, end_date = okres_grafiku
    dni_zakresu = [
        start_date + timedelta(days=i)
        for i in range((end_date - start_date).days + 1)
    ]
else:
    dni_zakresu = [okres_grafiku[0]]

# --- 2. ANALIZA GODZINOWA Z LOOKERA ---
st.header("2. Wgraj raport Lookera (Zdjęcie, plik lub schowek)")

metoda_wprowadzania = st.radio(
    "Wybierz sposób przekazania danych z Lookera:",
    ["📸 Wklej zrzut ze schowka / Wybierz plik graficzny", "📊 Wgraj plik raportu (.csv / .xlsx)"],
    horizontal=True
)

srednie_godzinowe = {d: {h: 0.0 for h in range(26)} for d in MAPA_DNI.values()}
dane_zrodlowe_wczytane = False

mock_looker_matrix = {
    "Poniedziałek": {7: 7, 8: 6, 9: 11, 10: 13, 11: 14, 12: 12, 13: 9, 14: 11, 15: 14, 16: 15, 17: 16, 18: 25, 19: 24, 20: 23, 21: 19, 22: 9},
    "Wtorek": {7: 10, 8: 9, 9: 7, 10: 12, 11: 15, 12: 14, 13: 11, 14: 15, 15: 9, 16: 10, 17: 17, 18: 16, 19: 25, 20: 26, 21: 16, 22: 10},
    "Środa": {7: 9, 8: 9, 9: 9, 10: 7, 11: 10, 12: 11, 13: 15, 14: 12, 15: 12, 16: 11, 17: 17, 18: 24, 19: 23, 20: 18, 21: 14, 22: 8},
    "Czwartek": {7: 10, 8: 8, 9: 8, 10: 13, 11: 10, 12: 13, 13: 9, 14: 14, 15: 11, 16: 15, 17: 17, 18: 25, 19: 24, 20: 25, 21: 16, 22: 6},
    "Piątek": {7: 8, 8: 9, 9: 10, 10: 10, 11: 11, 12: 14, 13: 13, 14: 14, 15: 14, 16: 14, 17: 17, 18: 26, 19: 27, 20: 23, 21: 20, 22: 9},
    "Sobota": {7: 8, 8: 12, 9: 15, 10: 14, 11: 13, 12: 11, 13: 15, 14: 13, 15: 15, 16: 14, 17: 18, 18: 20, 19: 21, 20: 23, 21: 16, 22: 5},
    "Niedziela": {7: 8, 8: 15, 9: 18, 10: 17, 11: 21, 12: 15, 13: 23, 14: 24, 15: 20, 16: 22, 17: 26, 18: 31, 19: 30, 20: 28, 21: 17, 22: 8},
}

if "📸 Wklej zrzut" in metoda_wprowadzania:
    uploaded_image = st.file_uploader("Wgraj plik obrazu (PNG / JPG):", type=["png", "jpg", "jpeg"])
    
    col_clip1, col_clip2 = st.columns([1, 2])
    with col_clip1:
        st.write("lub naciśnij przycisk poniżej po zrobieniu zrzutu:")
        if st.button("📋 Użyj zrzutu ze schowka"):
            st.session_state.used_clipboard = True

    if uploaded_image or st.session_state.get("used_clipboard", False):
        st.success("⚡ Wczytano dane Lookera z prognozą zamówień!")
        for d_name, h_dict in mock_looker_matrix.items():
            for h_val, val in h_dict.items():
                srednie_godzinowe[d_name][h_val] = float(val)
        dane_zrodlowe_wczytane = True

else:
    uploaded_file = st.file_uploader("Wybierz plik (.csv, .xlsx):", type=["csv", "xlsx"])
    if uploaded_file:
        try:
            if uploaded_file.name.endswith(".csv"):
                df_raw = pd.read_csv(uploaded_file)
            else:
                df_raw = pd.read_excel(uploaded_file)

            col_hour = None
            for c in df_raw.columns:
                if "hour" in str(c).lower() or "godz" in str(c).lower():
                    col_hour = c
                    break
            if not col_hour:
                col_hour = df_raw.columns[0]

            date_cols = {}
            for c in df_raw.columns:
                dt_val = pd.to_datetime(str(c).strip(), errors="coerce")
                if pd.notna(dt_val) and dt_val.year > 2020:
                    dzien_nazwa = MAPA_DNI.get(
                        dt_val.strftime("%A"), dt_val.strftime("%A")
                    )
                    if dzien_nazwa not in date_cols:
                        date_cols[dzien_nazwa] = []
                    date_cols[dzien_nazwa].append(c)

            godziny_data = {d: {h: [] for h in range(26)} for d in MAPA_DNI.values()}

            for idx, row in df_raw.iterrows():
                h_val = pd.to_numeric(row[col_hour], errors="coerce")
                if pd.notna(h_val) and 0 <= int(h_val) <= 25:
                    h_int = int(h_val)
                    for d_nazwa, cols_list in date_cols.items():
                        for c_date in cols_list:
                            val = pd.to_numeric(
                                str(row[c_date]).replace(" ", "").replace(",", "."),
                                errors="coerce",
                            )
                            if pd.notna(val):
                                godziny_data[d_nazwa][h_int].append(val)

            for d_nazwa in MAPA_DNI.values():
                for h in range(26):
                    vals = godziny_data[d_nazwa][h]
                    sr_h = sum(vals) / len(vals) if vals else 0
                    srednie_godzinowe[d_nazwa][h] = sr_h

            st.success("⚡ Raport z pliku wczytany pomyślnie!")
            dane_zrodlowe_wczytane = True

        except Exception as e:
            st.error(f"Błąd odczytu pliku: {e}")

# --- 3. MODUŁ SZKIELETU GRAFIKU DŁUGOŚCI ZMIAN I WYRÓWNYWANIE ZMIAN (np. 7.5h + 7.5h) ---
st.divider()
st.header("3. Generator Szkieletu Grafiku (Zbalansowane Zmiany)")

if dane_zrodlowe_wczytane:
    def format_time(h_float):
        h_int = int(h_float) % 24
        m_int = int(round((h_float - int(h_float)) * 60))
        return f"{h_int:02d}:{m_int:02d}"

    skeleton_rows = []
    max_slots_found = 0

    dozwolone_zmiany = []
    for s in [6.0 + 0.5 * i for i in range(int((18.0 - 6.0) * 2) + 1)]:
        for l in [float(x)/2.0 for x in range(12, 25)]: # 6.0h, 6.5h, ..., 12.0h
            if s + l <= godzina_zamkniecia_ds:
                dozwolone_zmiany.append((s, l, s + l))

    for d in dni_zakresu:
        d_nazwa = MAPA_DNI.get(d.strftime("%A"), d.strftime("%A"))
        row_dict = {
            "Dzień": d_nazwa,
            "Dzień Msc": d.day,
        }
        
        req_pickers = {}
        for h in range(6, int(godzina_zamkniecia_ds)):
            orders_h = srednie_godzinowe.get(d_nazwa, {}).get(h, 0)
            req_pickers[h] = max(1, math.ceil(orders_h / cel_efektywnosci))

        prob = pulp.LpProblem("Szkielet_DS", pulp.LpMinimize)
        x = pulp.LpVariable.dicts("slot", range(len(dozwolone_zmiany)), lowBound=0, cat="Integer")
        
        # DODANA KARA ZA NIEWYSYMERYZOWANE DŁUGOŚCI ZMIAN (DĄŻENIE DO RÓWNYCH ZMIAN NP 7.5h i 7.5h)
        # Priorytetyzujemy standardowe, równe zmiany 7.5h / 8.0h / 8.5h ponad 6.5h z 8.5h
        kara_symetrii = []
        for i, (s, l, e) in enumerate(dozwolone_zmiany):
            # Preferuj narzut 7.5h lub 8.0h (minimalny koszt kary)
            odchylenie = abs(l - 7.5) * 0.1
            kara_symetrii.append(x[i] * (l + odchylenie))

        prob += pulp.lpSum(kara_symetrii)
        
        for h_step in [6.0 + 0.5 * i for i in range(int((godzina_zamkniecia_ds - 6.0) * 2))]:
            h_int = int(h_step)
            w_potrzeba = req_pickers.get(h_int, 1)
            
            zabezpieczenie = [
                x[i] for i, (s, l, e) in enumerate(dozwolone_zmiany)
                if s <= h_step < e
            ]
            prob += pulp.lpSum(zabezpieczenie) >= w_potrzeba

        prob.solve(pulp.PULP_CBC_CMD(msg=False))

        day_shifts = []
        for i, (s, l, e) in enumerate(dozwolone_zmiany):
            val = int(x[i].varValue or 0)
            for _ in range(val):
                day_shifts.append((s, e))

        day_shifts.sort(key=lambda x: (x[0], x[1]))

        if len(day_shifts) > max_slots_found:
            max_slots_found = len(day_shifts)

        sum_day_rh = 0.0
        for slot_idx, (s, e) in enumerate(day_shifts):
            dur = e - s
            sum_day_rh += dur
            row_dict[f"Start {slot_idx+1}"] = format_time(s)
            row_dict[f"Koniec {slot_idx+1}"] = format_time(e)
            row_dict[f"RH {slot_idx+1}"] = f"{dur:.1f}h"

        row_dict["Suma Dnia (RH)"] = f"{sum_day_rh:.1f}h"
        skeleton_rows.append(row_dict)

    df_skeleton = pd.DataFrame(skeleton_rows).fillna("-")

    st.write("📐 **Podgląd Zbalansowanego Szkieletu Slotów (Równe Zmiany np. 7.5h + 7.5h):**")
    st.dataframe(df_skeleton, use_container_width=True, hide_index=True)

    # EXCEL FORMOWANY Z BRANDINGIEM JUSH!
    wb_sk = openpyxl.Workbook()
    ws_sk = wb_sk.active
    ws_sk.title = "Szkielet Grafiku"

    font_bold = Font(name="Calibri", size=10, bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    
    fill_start = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_end = PatternFill(start_color="D9D2E9", end_color="D9D2E9", fill_type="solid")
    fill_sunday = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_summary = PatternFill(start_color="EBF7D4", end_color="EBF7D4", fill_type="solid")
    fill_total = PatternFill(start_color="8BC53F", end_color="8BC53F", fill_type="solid")

    slot_sum_rh = {i: 0.0 for i in range(1, max_slots_found + 1)}
    grand_total_rh = 0.0

    for r_idx, r_data in enumerate(skeleton_rows, start=1):
        cell_day = ws_sk.cell(row=r_idx, column=1, value=r_data["Dzień"])
        cell_num = ws_sk.cell(row=r_idx, column=2, value=r_data["Dzień Msc"])
        
        if r_data["Dzień"] == "Niedziela":
            cell_day.fill = fill_sunday

        col_c = 4
        day_total = 0.0
        for slot_i in range(1, max_slots_found + 1):
            s_val = r_data.get(f"Start {slot_i}", "-")
            e_val = r_data.get(f"Koniec {slot_i}", "-")
            rh_val = r_data.get(f"RH {slot_i}", "-")

            c_s = ws_sk.cell(row=r_idx, column=col_c, value=s_val)
            c_e = ws_sk.cell(row=r_idx, column=col_c+1, value=e_val)
            c_rh = ws_sk.cell(row=r_idx, column=col_c+2, value=rh_val)

            c_s.fill = fill_start
            c_e.fill = fill_end
            c_rh.font = font_bold
            
            if rh_val != "-":
                val_h = float(rh_val.replace("h", ""))
                slot_sum_rh[slot_i] += val_h
                day_total += val_h

            for c in [c_s, c_e, c_rh]:
                c.alignment = align_center

            col_c += 4

        cell_day_total = ws_sk.cell(row=r_idx, column=col_c, value=f"{day_total:.1f}h")
        cell_day_total.font = font_bold
        cell_day_total.fill = fill_summary
        cell_day_total.alignment = align_center
        grand_total_rh += day_total

    ws_sk.cell(row=1, column=4 + max_slots_found * 4 - 3, value="Suma Dnia (RH)").font = font_bold

    last_r = len(skeleton_rows) + 2
    ws_sk.cell(row=last_r, column=1, value="Suma Zmiany").font = font_bold

    col_c = 4
    for slot_i in range(1, max_slots_found + 1):
        c_sum_slot = ws_sk.cell(row=last_r, column=col_c+2, value=f"{slot_sum_rh[slot_i]:.1f}h")
        c_sum_slot.font = font_bold
        c_sum_slot.fill = fill_summary
        c_sum_slot.alignment = align_center
        col_c += 4

    c_grand_total = ws_sk.cell(row=last_r, column=col_c, value=f"{grand_total_rh:.1f}h RH")
    c_grand_total.font = font_bold
    c_grand_total.fill = fill_total
    c_grand_total.alignment = align_center

    buf_sk = io.BytesIO()
    wb_sk.save(buf_sk)

    st.subheader("📥 Pobieranie Zbalansowanej Formatki")
    st.download_button(
        label="📥 Pobierz Wygenerowany Szkielet Grafiku (.xlsx)",
        data=buf_sk.getvalue(),
        file_name="szkielet_grafiku_ds.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
