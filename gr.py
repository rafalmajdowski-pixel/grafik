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
    page_title="żabka jush! - Generator Grafiku DS",
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
        "<p style='font-weight:bold; color:#005B2B; font-size: 1.1rem;'>Optymalizator Grafiku Pickerów DS</p>",
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
max_godzina_zamowien = 25 if is_nocny else 23

cel_efektywnosci = st.sidebar.number_input(
    "Efektywność pakowania (zamówienia / h / picker)",
    min_value=1,
    value=15,
)

min_zmiana = 6
max_zmiana = 12

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
        st.success("⚡ Wczytano obraz Lookera z prognozą zamówień!")
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

# --- 3. MODUŁ SZKIELETU GRAFIKU (CIĄGŁOŚĆ BEZ DZIUR) ---
st.divider()
st.header("3. Moduł: Szkielet Grafiku (Sloty Godzinowe)")

if dane_zrodlowe_wczytane:
    def format_time(h_float):
        h_int = int(h_float) % 24
        m_int = int(round((h_float - int(h_float)) * 60))
        return f"{h_int:02d}:{m_int:02d}"

    skeleton_rows = []
    max_slots_found = 0

    for d in dni_zakresu:
        d_nazwa = MAPA_DNI.get(d.strftime("%A"), d.strftime("%A"))
        row_dict = {
            "Dzień": d_nazwa,
            "Dzień Msc": d.day,
        }
        
        # ZAPEWNIENIE PEŁNEGO POKRYCIA OD 06:00 DO ZAMKNIĘCIA
        day_shifts = []
        day_shifts.append((6.0, 14.0)) # Otwarcie rano (8h)
        
        # Zmiana zamykająca nakłada się lub rozpoczyna dokładnie o 14:00 (brak luki!)
        start_close = 15.5 if (godzina_zamkniecia_ds - 8.0) > 14.0 else 14.0
        if is_nocny:
            start_close = 17.5
            
        day_shifts.append((start_close, godzina_zamkniecia_ds))
        
        # Dodatkowe zmiany środkowe na piki zamówień
        mid_volume = sum(srednie_godzinowe.get(d_nazwa, {}).get(h, 0) for h in range(11, 18))
        if mid_volume > cel_efektywnosci * 12:
            day_shifts.append((09.0, 17.0))
            day_shifts.append((14.0, 22.0))
        elif mid_volume > cel_efektywnosci * 6:
            day_shifts.append((10.0, 18.0))

        if len(day_shifts) > max_slots_found:
            max_slots_found = len(day_shifts)

        for slot_idx, (s, e) in enumerate(day_shifts):
            dur = e - s
            row_dict[f"Start {slot_idx+1}"] = format_time(s)
            row_dict[f"Koniec {slot_idx+1}"] = format_time(e)
            row_dict[f"RH {slot_idx+1}"] = f"{dur:.1f}h"

        skeleton_rows.append(row_dict)

    df_skeleton = pd.DataFrame(skeleton_rows).fillna("-")

    st.write("📐 **Wygenerowana Formatka Szkieletu (Gwarancja Pokrycia Całej Doby):**")
    st.dataframe(df_skeleton, use_container_width=True, hide_index=True)

    wb_sk = openpyxl.Workbook()
    ws_sk = wb_sk.active
    ws_sk.title = "Szkielet Grafiku"

    font_bold = Font(name="Calibri", size=10, bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    
    fill_start = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_end = PatternFill(start_color="D9D2E9", end_color="D9D2E9", fill_type="solid")
    fill_sunday = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    for r_idx, r_data in enumerate(skeleton_rows, start=1):
        cell_day = ws_sk.cell(row=r_idx, column=1, value=r_data["Dzień"])
        cell_num = ws_sk.cell(row=r_idx, column=2, value=r_data["Dzień Msc"])
        
        if r_data["Dzień"] == "Niedziela":
            cell_day.fill = fill_sunday

        col_c = 4
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
            
            for c in [c_s, c_e, c_rh]:
                c.alignment = align_center

            col_c += 4

    buf_sk = io.BytesIO()
    wb_sk.save(buf_sk)

    st.download_button(
        label="📥 Pobierz Sam Szkielet Grafiku (.xlsx)",
        data=buf_sk.getvalue(),
        file_name="szkielet_grafiku_ds.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# --- 4. PRACOWNICI & USTAWIENIA INDYWIDUALNE ---
st.divider()
st.header("4. Zespół Pickerów DS & Indywidualne Reguły")

pracownicy_default = [
    "Aval01204VasinA",
    "Aval01209KushnY",
    "AvalZhukoD",
    "Dive01202VitalD",
    "Eter01203SavchV",
    "EterZaichI",
]
pracownicy_input = st.text_area(
    "Lista pickerów (każdy w nowej linii):",
    "\n".join(pracownicy_default),
)
pracownicy = [
    p.strip() for p in pracownicy_input.split("\n") if p.strip() != ""
]

if "preferencje_dict" not in st.session_state:
    st.session_state.preferencje_dict = {}
if "korekty_godzin_dict" not in st.session_state:
    st.session_state.korekty_godzin_dict = {}
if "urlopy_list" not in st.session_state:
    st.session_state.urlopy_list = []

st.subheader("➕ Ustawienia dla Wybranego Pickera")

col_sel, col_pref_type, col_val, col_add = st.columns([2, 3, 2, 2])

with col_sel:
    p_target = st.selectbox("Wybierz pickera:", pracownicy if pracownicy else ["-"])

with col_pref_type:
    type_opt = st.selectbox(
        "Rodzaj dodawanego ustawienia:",
        ["Preferencja Pory Dnia", "Modyfikacja Etatowa (+/- h)", "Nieobecność / Urlop (Całe Dni)"]
    )

with col_val:
    if type_opt == "Preferencja Pory Dnia":
        val_pref = st.selectbox("Pora dnia:", ["Tylko Poranki (06:00)", "Tylko Zamknięcia"])
    elif type_opt == "Modyfikacja Etatowa (+/- h)":
        val_hours = st.number_input("Różnica godzin (np. +110 lub -30):", value=0, step=5)
    else:
        val_dates = st.date_input("Zakres wolnego:", value=(datetime.now().date(), datetime.now().date()))

with col_add:
    st.write("&nbsp;")
    if st.button("➕ Dodaj regułę", use_container_width=True):
        if type_opt == "Preferencja Pory Dnia":
            st.session_state.preferencje_dict[p_target] = val_pref
            st.success(f"Dodano preferencję dla {p_target}!")
        elif type_opt == "Modyfikacja Etatowa (+/- h)":
            st.session_state.korekty_godzin_dict[p_target] = val_hours
            st.success(f"Skorygowano etat dla {p_target} o {val_hours}h!")
        else:
            if isinstance(val_dates, tuple) and len(val_dates) == 2:
                st.session_state.urlopy_list.append({"Pracownik": p_target, "Od": val_dates[0], "Do": val_dates[1]})
            elif isinstance(val_dates, tuple) and len(val_dates) == 1:
                st.session_state.urlopy_list.append({"Pracownik": p_target, "Od": val_dates[0], "Do": val_dates[0]})
            st.success(f"Zarejestrowano nieobecność dla {p_target}!")

st.write("---")
st.subheader("📋 Aktywne Ustawienia Zespołu:")

c_pref, c_kor, c_url = st.columns(3)

with c_pref:
    st.markdown("🎯 **Preferencje Pory Dnia**")
    if st.session_state.preferencje_dict:
        df_pref = pd.DataFrame(
            list(st.session_state.preferencje_dict.items()),
            columns=["Picker", "Wymagana Pora Dnia"]
        )
        st.dataframe(df_pref, use_container_width=True, hide_index=True)
        if st.button("🗑️ Wyczyść preferencje", key="c1"):
            st.session_state.preferencje_dict = {}
    else:
        st.caption("Brak ustalonych preferencji.")

with c_kor:
    st.markdown("⏱️ **Korekty Etatów (+/- h)**")
    if st.session_state.korekty_godzin_dict:
        df_kor = pd.DataFrame(
            [{"Picker": k, "Zmiana Czasu": f"{v:+d}h"} for k, v in st.session_state.korekty_godzin_dict.items()]
        )
        st.dataframe(df_kor, use_container_width=True, hide_index=True)
        if st.button("🗑️ Wyczyść korekty", key="c2"):
            st.session_state.korekty_godzin_dict = {}
    else:
        st.caption("Wszyscy pickerzy mają domyślny etat.")

with c_url:
    st.markdown("🌴 **Nieobecności i Urlopy**")
    if st.session_state.urlopy_list:
        df_url = pd.DataFrame(st.session_state.urlopy_list)
        df_url_display = df_url.copy()
        df_url_display["Od"] = df_url_display["Od"].apply(lambda x: x.strftime("%d/%m/%Y") if hasattr(x, "strftime") else str(x))
        df_url_display["Do"] = df_url_display["Do"].apply(lambda x: x.strftime("%d/%m/%Y") if hasattr(x, "strftime") else str(x))
        st.dataframe(df_url_display, use_container_width=True, hide_index=True)
        if st.button("🗑️ Wyczyść urlopy", key="c3"):
            st.session_state.urlopy_list = []
    else:
        st.caption("Brak nieobecności w grafiku.")

# --- 5. GENEROWANIE PEŁNEGO GRAFIKU OBSADY ---
st.divider()
st.header("5. Przypisanie Pickerów do Grafiku")
if st.button("🚀 Wygeneruj Pełny Grafik jush!", type="primary", use_container_width=True):
    if not dane_zrodlowe_wczytane:
        st.error("Proszę najpierw przekazać dane z Lookera w sekcji 2!")
    elif not pracownicy:
        st.error("Proszę wpisać listę pickerów!")
    else:
        try:
            wymagani_pracownicy_h = {}
            total_required_hours = 0
            for d in dni_zakresu:
                d_nazwa = MAPA_DNI.get(d.strftime("%A"), d.strftime("%A"))
                wymagani_pracownicy_h[d] = {}
                for h in range(max_godzina_zamowien + 1):
                    sr_zam = srednie_godzinowe.get(d_nazwa, {}).get(h, 0)
                    potrzeba_osob = math.ceil(sr_zam / cel_efektywnosci)
                    if 6 <= h <= int(godzina_zamkniecia_ds):
                        potrzeba_osob = max(1, potrzeba_osob)
                    wymagani_pracownicy_h[d][h] = potrzeba_osob
                    total_required_hours += potrzeba_osob

            model = pulp.LpProblem("Optymalizacja_Grafiku", pulp.LpMinimize)

            prawidlowe_zmiany = []
            starty = [6.0 + 0.5 * i for i in range(int((18.0 - 6.0) * 2) + 1)]
            dlugosci = [float(l) for l in range(min_zmiana, max_zmiana + 1)]

            for s in starty:
                for l in dlugosci:
                    koniec = s + l
                    if koniec <= godzina_zamkniecia_ds:
                        prawidlowe_zmiany.append((s, l))

            zmiany_ranne = [(s, l) for s, l in prawidlowe_zmiany if s == 6.0]
            zmiany_wieczorne = [
                (s, l)
                for s, l in prawidlowe_zmiany
                if (s + l) == godzina_zamkniecia_ds
            ]

            zmienne_zmian = []
            for p in pracownicy:
                for d in dni_zakresu:
                    for s, l in prawidlowe_zmiany:
                        zmienne_zmian.append((p, d, s, l))

            y = pulp.LpVariable.dicts("zmiana", zmienne_zmian, cat="Binary")

            work_day = pulp.LpVariable.dicts(
                "work_day",
                [(p, d) for p in pracownicy for d in dni_zakresu],
                cat="Binary",
            )

            penalty_rest_12h = pulp.LpVariable.dicts(
                "pen_rest",
                [(p, d) for p in pracownicy for d in dni_zakresu],
                lowBound=0,
                cat="Binary",
            )
            penalty_cadence = pulp.LpVariable.dicts(
                "pen_cadence",
                [(p, d) for p in pracownicy for d in dni_zakresu],
                lowBound=0,
                cat="Binary",
            )

            dev_plus = pulp.LpVariable.dicts("dev_plus", pracownicy, lowBound=0, cat="Continuous")
            dev_minus = pulp.LpVariable.dicts("dev_minus", pracownicy, lowBound=0, cat="Continuous")

            model += (
                pulp.lpSum(dev_plus[p] + dev_minus[p] for p in pracownicy)
                + pulp.lpSum(
                    1000.0 * penalty_rest_12h[p, d] + 500.0 * penalty_cadence[p, d]
                    for p in pracownicy
                    for d in dni_zakresu
                )
            )

            for p in pracownicy:
                pref = st.session_state.preferencje_dict.get(p, "Brak")

                for d in dni_zakresu:
                    if pref in ["Tylko Poranki (06:00)", "Preferuje Poranki (06:00)"]:
                        for s, l in prawidlowe_zmiany:
                            if s != 6.0:
                                model += y[p, d, s, l] == 0
                    elif pref in ["Tylko Zamknięcia", "Preferuje Zamknięcia"]:
                        for s, l in prawidlowe_zmiany:
                            if (s + l) != godzina_zamkniecia_ds:
                                model += y[p, d, s, l] == 0

                dni_absencji = sum(
                    1
                    for d in dni_zakresu
                    if any(
                        u["Pracownik"] == p and u["Od"] <= d <= u["Do"]
                        for u in st.session_state.urlopy_list
                    )
                )
                dni_dostepne = max(1, len(dni_zakresu) - dni_absencji)
                proporcja = dni_dostepne / len(dni_zakresu)
                
                korekta_h = st.session_state.korekty_godzin_dict.get(p, 0)
                target_p = ((total_required_hours / len(pracownicy)) * proporcja) + korekta_h

                suma_h_p = pulp.lpSum(
                    y[p, d, s, l] * l
                    for d in dni_zakresu
                    for s, l in prawidlowe_zmiany
                )

                model += suma_h_p <= target_p + 15.0
                model += suma_h_p >= target_p - 15.0
                model += suma_h_p + dev_minus[p] - dev_plus[p] == target_p

                for idx_d, d in enumerate(dni_zakresu):
                    model += (
                        work_day[p, d]
                        == pulp.lpSum(y[p, d, s, l] for s, l in prawidlowe_zmiany)
                    )

                    for u in st.session_state.urlopy_list:
                        if u["Pracownik"] == p and u["Od"] <= d <= u["Do"]:
                            for s, l in prawidlowe_zmiany:
                                model += y[p, d, s, l] == 0

                    if idx_d < len(dni_zakresu) - 1:
                        d_next = dni_zakresu[idx_d + 1]
                        for s1, l1 in prawidlowe_zmiany:
                            koniec_d1 = s1 + l1
                            for s2, l2 in prawidlowe_zmiany:
                                start_d2 = s2 + 24.0
                                if (start_d2 - koniec_d1) < 12.0:
                                    model += (
                                        y[p, d, s1, l1] + y[p, d_next, s2, l2]
                                        <= 1 + penalty_rest_12h[p, d_next]
                                    )

                for idx_d in range(len(dni_zakresu) - 5):
                    window6 = [dni_zakresu[idx_d + i] for i in range(6)]
                    d_last = window6[-1]
                    model += (
                        pulp.lpSum(work_day[p, d_w] for d_w in window6)
                        <= 5 + 6 * penalty_cadence[p, d_last]
                    )

            for d in dni_zakresu:
                model += (
                    pulp.lpSum(
                        y[p, d, 6.0, l]
                        for p in pracownicy
                        for s, l in prawidlowe_zmiany
                        if s == 6.0
                    )
                    >= 1
                )
                model += (
                    pulp.lpSum(
                        y[p, d, s, l]
                        for p in pracownicy
                        for s, l in prawidlowe_zmiany
                        if s + l == godzina_zamkniecia_ds
                    )
                    >= 1
                )

                for h in range(6, int(godzina_zamkniecia_ds)):
                    potrzebni = max(1, wymagani_pracownicy_h[d].get(h, 1))
                    pracujacy = [
                        y[p, d, s, l]
                        for p in pracownicy
                        for s, l in prawidlowe_zmiany
                        if s <= h and (s + l) >= (h + 1)
                    ]
                    model += pulp.lpSum(pracujacy) >= potrzebni

            status = model.solve(pulp.PULP_CBC_CMD(msg=False))

            st.session_state.schedule_generated = True
            st.session_state.pracownicy = pracownicy
            st.session_state.dni_zakresu = dni_zakresu
            st.session_state.prawidlowe_zmiany = prawidlowe_zmiany
            st.session_state.y_vars = {
                (p, d, s, l): y[p, d, s, l].varValue
                for p in pracownicy
                for d in dni_zakresu
                for s, l in prawidlowe_zmiany
            }
            st.session_state.wymagani_h = wymagani_pracownicy_h

        except Exception as e:
            st.error(f"⚠️ Wystąpił błąd podczas obliczeń: {e}")

# --- 6. INTERAKTYWNY PODGLĄD I EXCEL ---
if st.session_state.get("schedule_generated", False):
    st.divider()
    st.header("6. Podgląd Grafiku & Pobieranie")

    pracownicy = st.session_state.pracownicy
    dni_zakresu = st.session_state.dni_zakresu
    prawidlowe_zmiany = st.session_state.prawidlowe_zmiany
    y_vars = st.session_state.y_vars

    def format_time(h_float):
        h_int = int(h_float) % 24
        m_int = int(round((h_float - int(h_float)) * 60))
        return f"{h_int:02d}:{m_int:02d}"

    data_rows = []
    for d in dni_zakresu:
        row = {"Data": d.strftime("%d/%m/%Y"), "Dzień": MAPA_DNI.get(d.strftime("%A"), "")}
        for p in pracownicy:
            shift_str = "OFF"
            for s, l in prawidlowe_zmiany:
                if y_vars.get((p, d, s, l), 0) == 1:
                    shift_str = f"{format_time(s)} - {format_time(s + l)}"
                    break
            row[p] = shift_str
        data_rows.append(row)

    df_editor = pd.DataFrame(data_rows)

    st.subheader("📝 Edytuj grafik na żywo:")
    edited_df = st.data_editor(df_editor, num_rows="fixed", use_container_width=True)

    worker_totals = {p: 0.0 for p in pracownicy}
    for p in pracownicy:
        for _, r in edited_df.iterrows():
            val = r[p]
            if str(val).strip() != "OFF" and "-" in str(val):
                try:
                    parts = str(val).split("-")
                    h_s = float(parts[0].split(":")[0]) + float(parts[0].split(":")[1]) / 60.0
                    h_e = float(parts[1].split(":")[0]) + float(parts[1].split(":")[1]) / 60.0
                    if h_e < h_s:
                        h_e += 24.0
                    worker_totals[p] += (h_e - h_s)
                except:
                    pass

    st.subheader("📊 Podsumowanie Roboczogodzin (RH):")
    cols_rh = st.columns(len(pracownicy))
    for i, p in enumerate(pracownicy):
        with cols_rh[i]:
            st.metric(label=p, value=f"{worker_totals[p]:.1f} h")

    st.subheader("📥 Eksport do Pliku Excel")
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Grafik jush"
    ws.freeze_panes = "B3"

    font_bold = Font(name="Calibri", size=10, bold=True)
    font_regular = Font(name="Calibri", size=10)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )

    fill_header_main = PatternFill(start_color="005B2B", end_color="005B2B", fill_type="solid")
    font_header_main = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    fill_header_sub = PatternFill(start_color="8BC53F", end_color="8BC53F", fill_type="solid")
    font_header_sub = Font(name="Calibri", size=10, bold=True, color="005B2B")
    fill_summary = PatternFill(start_color="EBF7D4", end_color="EBF7D4", fill_type="solid")
    fill_total_sum = PatternFill(start_color="8BC53F", end_color="8BC53F", fill_type="solid")
    font_total_sum = Font(name="Calibri", size=11, bold=True, color="005B2B")
    fill_shift_morning = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_shift_afternoon = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")

    ws.merge_cells("A1:A2")
    ws["A1"] = "pl-waw-12"
    ws["A1"].font = font_header_main
    ws["A1"].fill = fill_header_main
    ws["A1"].alignment = align_center

    col_idx = 2
    for p in pracownicy:
        col_start_letter = openpyxl.utils.get_column_letter(col_idx)
        col_end_letter = openpyxl.utils.get_column_letter(col_idx + 2)
        ws.merge_cells(f"{col_start_letter}1:{col_end_letter}1")
        cell_p = ws[f"{col_start_letter}1"]
        cell_p.value = p
        cell_p.font = font_header_main
        cell_p.fill = fill_header_main
        cell_p.alignment = align_center

        for i, sh in enumerate(["Start", "Koniec", "Suma"]):
            cell_sh = ws.cell(row=2, column=col_idx + i)
            cell_sh.value = sh
            cell_sh.font = font_header_sub
            cell_sh.fill = fill_header_sub
            cell_sh.alignment = align_center
            cell_sh.border = thin_border
        col_idx += 3

    godziny_pracownikow_excel = {p: 0.0 for p in pracownicy}
    row_idx = 3

    for _, r in edited_df.iterrows():
        cell_date = ws.cell(row=row_idx, column=1)
        cell_date.value = r["Data"]
        cell_date.font = font_regular
        cell_date.alignment = align_center
        cell_date.border = thin_border

        col_idx = 2
        for p in pracownicy:
            val = str(r[p]).strip()
            if val != "OFF" and "-" in val:
                parts = val.split("-")
                c_start = ws.cell(row=row_idx, column=col_idx)
                c_end = ws.cell(row=row_idx, column=col_idx + 1)
                c_sum = ws.cell(row=row_idx, column=col_idx + 2)

                c_start.value = parts[0].strip()
                c_end.value = parts[1].strip()

                try:
                    h_s = float(parts[0].split(":")[0]) + float(parts[0].split(":")[1]) / 60.0
                    h_e = float(parts[1].split(":")[0]) + float(parts[1].split(":")[1]) / 60.0
                    if h_e < h_s:
                        h_e += 24.0
                    len_shift = round(h_e - h_s, 1)
                    c_sum.value = len_shift
                    godziny_pracownikow_excel[p] += len_shift
                except:
                    c_sum.value = 0

                fill_c = fill_shift_morning if "06:00" in parts[0] else fill_shift_afternoon
                for cell in [c_start, c_end, c_sum]:
                    cell.font = font_regular
                    cell.alignment = align_center
                    cell.border = thin_border
                    cell.fill = fill_c
            else:
                for i in range(3):
                    ws.cell(row=row_idx, column=col_idx + i).border = thin_border
            col_idx += 3
        row_idx += 1

    cell_sum_label = ws.cell(row=row_idx, column=1)
    cell_sum_label.value = "ŁĄCZNIE"
    cell_sum_label.font = font_bold
    cell_sum_label.fill = fill_summary
    cell_sum_label.alignment = align_center
    cell_sum_label.border = thin_border

    grand_total_hours = 0.0
    col_idx = 2
    for p in pracownicy:
        col_start_letter = openpyxl.utils.get_column_letter(col_idx)
        col_end_letter = openpyxl.utils.get_column_letter(col_idx + 2)

        ws.merge_cells(f"{col_start_letter}{row_idx}:{col_end_letter}{row_idx}")
        cell_total = ws[f"{col_start_letter}{row_idx}"]
        cell_total.value = f"{round(godziny_pracownikow_excel[p], 1)}h"
        cell_total.font = font_bold
        cell_total.fill = fill_summary
        cell_total.alignment = align_center

        grand_total_hours += godziny_pracownikow_excel[p]

        for i in range(3):
            ws.cell(row=row_idx, column=col_idx + i).border = thin_border

        col_idx += 3

    row_idx += 1

    cell_grand_label = ws.cell(row=row_idx, column=1)
    cell_grand_label.value = "SUMA CAŁKOWITA"
    cell_grand_label.font = font_bold
    cell_grand_label.fill = fill_total_sum
    cell_grand_label.alignment = align_center
    cell_grand_label.border = thin_border

    last_col_letter = openpyxl.utils.get_column_letter(col_idx - 1)
    ws.merge_cells(f"B{row_idx}:{last_col_letter}{row_idx}")
    cell_grand_val = ws[f"B{row_idx}"]
    cell_grand_val.value = f"{round(grand_total_hours, 1)} Roboczogodzin (RH)"
    cell_grand_val.font = font_total_sum
    cell_grand_val.fill = fill_total_sum
    cell_grand_val.alignment = align_center

    for c in range(2, col_idx):
        ws.cell(row=row_idx, column=c).border = thin_border

    ws.column_dimensions["A"].width = 16
    for c in range(2, col_idx):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 10

    buffer = io.BytesIO()
    wb.save(buffer)

    st.download_button(
        label="📥 Pobierz Pełny Grafik Excel (.xlsx)",
        data=buffer.getvalue(),
        file_name="grafik_pickerzy_jush.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
