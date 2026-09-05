from datetime import datetime, timedelta
import io
import math
import pandas as pd
import streamlit as st

# --- KONFIGURACJA STRONY STREAMLIT ---
st.set_page_config(
    page_title="Jush! - Optymalizator Grafiku Magazynu (DS)",
    page_icon="⚡",
    layout="wide",
)

# --- STYLIZACJA W BRANDINGU JUSH! (CSS) ---
st.markdown(
    """
    <style>
    /* Kolorystyka i Styl Jush! */
    :root {
        --jush-green: #00E600;
        --jush-dark: #0A0F1D;
        --jush-purple: #7E22CE;
        --jush-light-green: #DCFCE7;
    }
    
    /* Pasek boczny */
    [data-testid="stSidebar"] {
        background-color: #0F172A !important;
        color: white !important;
    }
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    
    /* Przycisk główny (Jush Green) */
    div.stButton > button:first-child {
        background-color: #00E600 !important;
        color: #000000 !important;
        font-weight: bold !important;
        font-size: 18px !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 12px 28px !important;
        box-shadow: 0px 4px 14px rgba(0, 230, 0, 0.4) !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:first-child:hover {
        background-color: #00C800 !important;
        transform: scale(1.02) !important;
    }
    
    /* Nagłówki sekcji */
    h1, h2, h3 {
        color: #0F172A !important;
        font-family: 'Arial Black', sans-serif !important;
    }
    
    /* Plakietki i wyróżnienia */
    .jush-badge {
        background-color: #00E600;
        color: black;
        padding: 4px 12px;
        border-radius: 20px;
        font-weight: bold;
        font-size: 14px;
        display: inline-block;
        margin-bottom: 10px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Bezpieczne importy bibliotek
try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
except ImportError:
    st.error(
        "❌ Brakuje biblioteki 'openpyxl'. Upewnij się, że dodałeś 'openpyxl' do pliku requirements.txt na GitHubie!"
    )

try:
    import pulp
except ImportError:
    st.error(
        "❌ Brakuje biblioteki 'pulp'. Upewnij się, że dodałeś 'pulp' do pliku requirements.txt na GitHubie!"
    )

# --- BRANDING JUSH! - BRAND HEADER ---
col_logo, col_title = st.columns([1, 5])
with col_logo:
    st.image(
        "https://zabkagroup.com/wp-content/uploads/2022/09/Jush_logo.png",
        width=140,
    )
with col_title:
    st.markdown(
        "<h1 style='margin-bottom:0;'>jush! <span style='color:#00E600;'>DS Scheduler</span></h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='font-weight:bold; color:#64748B;'>Inteligentny system optymalizacji grafików dla Dark Store'ów Jush!</p>",
        unsafe_allow_html=True,
    )

st.divider()

# --- SIDEBAR: PARAMETRY EFEKTYWNOŚCI I OBSADY ---
st.sidebar.image(
    "https://zabkagroup.com/wp-content/uploads/2022/09/Jush_logo.png", width=100
)
st.sidebar.header("⚙️ Parametry Magazynu DS")

typ_magazynu = st.sidebar.selectbox("Typ magazynu", ["Standardowy", "Nocny"])
is_nocny = typ_magazynu == "Nocny"

godzina_otwarcia_ds = 6.0
godzina_zamkniecia_ds = 25.5 if is_nocny else 23.5  # 25.5h = 01:30 w nocy
max_godzina_zamowien = 25 if is_nocny else 23  # 25h = 01:00 w nocy

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
st.header("2. Wgraj raport zamówień z Lookera")
uploaded_file = st.file_uploader(
    "Wybierz plik CSV lub Excel z Lookera (hourly volume)", type=["csv", "xlsx"]
)

srednie_godzinowe = {}

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

        godziny_data = {
            d: {h: [] for h in range(26)} for d in MAPA_DNI.values()
        }

        for idx, row in df_raw.iterrows():
            h_val = pd.to_numeric(row[col_hour], errors="coerce")
            if pd.notna(h_val) and 0 <= int(h_val) <= 25:
                h_int = int(h_val)
                for d_nazwa, cols_list in date_cols.items():
                    for c_date in cols_list:
                        val = pd.to_numeric(
                            str(row[c_date])
                            .replace(" ", "")
                            .replace(",", "."),
                            errors="coerce",
                        )
                        if pd.notna(val):
                            godziny_data[d_nazwa][h_int].append(val)

        for d_nazwa in MAPA_DNI.values():
            srednie_godzinowe[d_nazwa] = {}
            for h in range(26):
                vals = godziny_data[d_nazwa][h]
                sr_h = sum(vals) / len(vals) if vals else 0
                srednie_godzinowe[d_nazwa][h] = sr_h

        st.success("⚡ Raport Lookera zweryfikowany pomyślnie!")

    except Exception as e:
        st.error(f"Błąd odczytu pliku z Lookera: {e}")

# --- 3. PRACOWNICI I KALENDARZ URLOPOWY ---
st.header("3. Zespół Curierów & Kurierów DS")
pracownicy_input = st.text_area(
    "Lista kurierów / pickerów (każdy w nowej linii):",
    "Aval01204VasinA\nAval01209KushnY\nAvalZhukoD\nDive01202VitalD\nEter01203SavchV\nEterZaichI",
)
pracownicy = [
    p.strip() for p in pracownicy_input.split("\n") if p.strip() != ""
]

if "urlopy_list" not in st.session_state:
    st.session_state.urlopy_list = []

st.subheader("Niedostępność kuriera:")
col_p, col_u1, col_u2, col_btn = st.columns([2, 2, 2, 1])

with col_p:
    p_select = st.selectbox("Pracownik", pracownicy if pracownicy else ["-"])
with col_u1:
    u_start = st.date_input("Start wolnego", datetime.now().date())
with col_u2:
    u_end = st.date_input("Koniec wolnego", datetime.now().date())
with col_btn:
    st.write("")
    st.write("")
    if st.button("➕ Dodaj wolne"):
        st.session_state.urlopy_list.append(
            {"Pracownik": p_select, "Od": u_start, "Do": u_end}
        )

if st.session_state.urlopy_list:
    st.write("📋 **Zarejestrowane nieobecności:**")
    st.dataframe(pd.DataFrame(st.session_state.urlopy_list))
    if st.button("🗑️ Wyczyść listę wolnych"):
        st.session_state.urlopy_list = []

# --- 4. GENEROWANIE GRAFIKU DLA DS JUSH! ---
st.header("4. Optymalizacja Grafiku")
if st.button("🚀 Wygeneruj Grafik Jush!", type="primary"):
    if not uploaded_file:
        st.error("Proszę najpierw wgrać plik z Lookera!")
    elif not pracownicy:
        st.error("Proszę wpisać listę pracowników!")
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

            dev_plus = pulp.LpVariable.dicts(
                "dev_plus", pracownicy, lowBound=0, cat="Continuous"
            )
            dev_minus = pulp.LpVariable.dicts(
                "dev_minus", pracownicy, lowBound=0, cat="Continuous"
            )
            dev_ranne = pulp.LpVariable.dicts(
                "dev_ranne", pracownicy, lowBound=0, cat="Continuous"
            )
            dev_wieczor = pulp.LpVariable.dicts(
                "dev_wieczor", pracownicy, lowBound=0, cat="Continuous"
            )

            model += pulp.lpSum(
                dev_plus[p]
                + dev_minus[p]
                + 2.0 * dev_ranne[p]
                + 2.0 * dev_wieczor[p]
                for p in pracownicy
            )

            for p in pracownicy:
                dni_absencji = 0
                for d in dni_zakresu:
                    for u in st.session_state.urlopy_list:
                        if u["Pracownik"] == p and u["Od"] <= d <= u["Do"]:
                            dni_absencji += 1
                            break

                dni_dostepne = max(1, len(dni_zakresu) - dni_absencji)
                proporcja_dostepnosci = dni_dostepne / len(dni_zakresu)
                target_p = (
                    total_required_hours / len(pracownicy)
                ) * proporcja_dostepnosci

                suma_h_p = pulp.lpSum(
                    y[p, d, s, l] * l
                    for d in dni_zakresu
                    for s, l in prawidlowe_zmiany
                )

                model += suma_h_p <= target_p + 10.0
                model += suma_h_p >= target_p - 10.0
                model += suma_h_p + dev_minus[p] - dev_plus[p] == target_p

                num_ranne = pulp.lpSum(
                    y[p, d, s, l] for d in dni_zakresu for s, l in zmiany_ranne
                )
                num_wieczorne = pulp.lpSum(
                    y[p, d, s, l]
                    for d in dni_zakresu
                    for s, l in zmiany_wieczorne
                )

                model += num_ranne - num_wieczorne <= 2.0 + dev_ranne[p]
                model += num_wieczorne - num_ranne <= 2.0 + dev_wieczor[p]

                for idx_d, d in enumerate(dni_zakresu):
                    model += (
                        work_day[p, d]
                        == pulp.lpSum(y[p, d, s, l] for s, l in prawidlowe_zmiany)
                    )

                    for u in st.session_state.urlopy_list:
                        if u["Pracownik"] == p and u["Od"] <= d <= u["Do"]:
                            for s, l in prawidlowe_zmiany:
                                model += y[p, d, s, l] == 0

                    # 12H ODPOCZYNKU
                    if idx_d < len(dni_zakresu) - 1:
                        d_next = dni_zakresu[idx_d + 1]
                        for s1, l1 in prawidlowe_zmiany:
                            koniec_dzien1 = s1 + l1
                            for s2, l2 in prawidlowe_zmiany:
                                start_dzien2 = s2 + 24.0
                                if (start_dzien2 - koniec_dzien1) < 12.0:
                                    model += (
                                        y[p, d, s1, l1] + y[p, d_next, s2, l2]
                                        <= 1
                                    )

                # REGULA: SYSTEM 4-5 DNI WORK / 1-2 OFF
                for idx_d in range(len(dni_zakresu) - 5):
                    window6 = [dni_zakresu[idx_d + i] for i in range(6)]
                    model += (
                        pulp.lpSum(work_day[p, d_w] for d_w in window6) <= 5
                    )

                for idx_d in range(len(dni_zakresu) - 2):
                    window3 = [dni_zakresu[idx_d + i] for i in range(3)]
                    has_vacation = any(
                        u["Pracownik"] == p
                        and u["Od"] <= window3[2]
                        and u["Do"] >= window3[0]
                        for u in st.session_state.urlopy_list
                    )
                    if not has_vacation:
                        model += (
                            pulp.lpSum(work_day[p, d_w] for d_w in window3)
                            >= 1
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

                for h in range(7, max_godzina_zamowien):
                    potrzebni = wymagani_pracownicy_h[d].get(h, 0)
                    if potrzebni > 0:
                        pracujacy_w_godzinie = []
                        for p in pracownicy:
                            for s, l in prawidlowe_zmiany:
                                if s <= h and (s + l) >= (h + 1):
                                    pracujacy_w_godzinie.append(y[p, d, s, l])

                        model += pulp.lpSum(pracujacy_w_godzinie) >= potrzebni

            status = model.solve(pulp.PULP_CBC_CMD(msg=False))

            if pulp.LpStatus[status] != "Optimal":
                st.error(
                    "❌ NIE MOŻNA WYGENEROWAĆ GRAFIKU JUSH! Reguły (4-5 dni pracy, 12h odpoczynku, obsługa spływu Lookera) "
                    "oraz urlopy uniemożliwiają ułożenie grafiku. Dodaj więcej pracowników na zlecenie."
                )
            else:
                # --- EXCEL STYLIZOWANY W BARWACH JUSH! ---
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Grafik Jush"
                ws.freeze_panes = "B3"

                font_bold = Font(name="Calibri", size=10, bold=True)
                font_regular = Font(name="Calibri", size=10)
                align_center = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
                thin_border = Border(
                    left=Side(style="thin", color="D3D3D3"),
                    right=Side(style="thin", color="D3D3D3"),
                    top=Side(style="thin", color="D3D3D3"),
                    bottom=Side(style="thin", color="D3D3D3"),
                )

                # PALETA JUSH!
                fill_header_main = PatternFill(
                    start_color="3B0764", end_color="3B0764", fill_type="solid"
                )  # Głęboki fiolet Jush
                font_header_main = Font(
                    name="Calibri", size=10, bold=True, color="FFFFFF"
                )

                fill_header_sub = PatternFill(
                    start_color="6B21A8", end_color="6B21A8", fill_type="solid"
                )  # Jasny fiolet
                font_header_sub = Font(
                    name="Calibri", size=10, bold=True, color="FFFFFF"
                )

                fill_summary = PatternFill(
                    start_color="E9D5FF", end_color="E9D5FF", fill_type="solid"
                )  # Jasnofioletowe podsumowanie

                fill_total_sum = PatternFill(
                    start_color="00E600", end_color="00E600", fill_type="solid"
                )  # Jush Neon Green
                font_total_sum = Font(
                    name="Calibri", size=11, bold=True, color="000000"
                )

                fill_date_weekend = PatternFill(
                    start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
                )
                fill_date_weekday = PatternFill(
                    start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"
                )

                fill_shift_morning = PatternFill(
                    start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"
                )  # Soczysta zieleń poranków
                fill_shift_afternoon = PatternFill(
                    start_color="FEF08A", end_color="FEF08A", fill_type="solid"
                )  # Ciepły żółty popołudnia

                ws.merge_cells("A1:A2")
                ws["A1"] = "pl-waw-12"
                ws["A1"].font = font_header_main
                ws["A1"].fill = fill_header_main
                ws["A1"].alignment = align_center

                def format_time(h_float):
                    h_int = int(h_float) % 24
                    m_int = int(round((h_float - int(h_float)) * 60))
                    return f"{h_int:02d}:{m_int:02d}"

                col_idx = 2
                for p in pracownicy:
                    col_start_letter = openpyxl.utils.get_column_letter(
                        col_idx
                    )
                    col_end_letter = openpyxl.utils.get_column_letter(
                        col_idx + 2
                    )

                    ws.merge_cells(f"{col_start_letter}1:{col_end_letter}1")
                    cell_p = ws[f"{col_start_letter}1"]
                    cell_p.value = p
                    cell_p.font = font_header_main
                    cell_p.fill = fill_header_main
                    cell_p.alignment = align_center

                    sub_headers = ["Start", "Koniec", "Suma"]
                    for i, sh in enumerate(sub_headers):
                        cell_sh = ws.cell(row=2, column=col_idx + i)
                        cell_sh.value = sh
                        cell_sh.font = font_header_sub
                        cell_sh.fill = fill_header_sub
                        cell_sh.alignment = align_center
                        cell_sh.border = thin_border

                    col_idx += 3

                godziny_pracownikow = {p: 0.0 for p in pracownicy}
                row_idx = 3

                for d in dni_zakresu:
                    cell_date = ws.cell(row=row_idx, column=1)
                    cell_date.value = d.strftime("%d/%m/%Y")
                    cell_date.font = font_regular
                    cell_date.alignment = align_center
                    cell_date.border = thin_border

                    if d.weekday() in [5, 6]:
                        cell_date.fill = fill_date_weekend
                    else:
                        cell_date.fill = fill_date_weekday

                    col_idx = 2
                    for p in pracownicy:
                        assigned = False
                        for s, l in prawidlowe_zmiany:
                            if y[p, d, s, l].varValue == 1:
                                c_start = ws.cell(row=row_idx, column=col_idx)
                                c_end = ws.cell(row=row_idx, column=col_idx + 1)
                                c_sum = ws.cell(row=row_idx, column=col_idx + 2)

                                c_start.value = format_time(s)
                                c_end.value = format_time(s + l)
                                c_sum.value = round(l, 1)

                                godziny_pracownikow[p] += l

                                fill_color = (
                                    fill_shift_morning
                                    if s == 6.0
                                    else fill_shift_afternoon
                                )

                                for cell in [c_start, c_end, c_sum]:
                                    cell.font = font_regular
                                    cell.alignment = align_center
                                    cell.border = thin_border
                                    cell.fill = fill_color

                                assigned = True
                                break

                        if not assigned:
                            for i in range(3):
                                cell = ws.cell(row=row_idx, column=col_idx + i)
                                cell.border = thin_border

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
                    col_start_letter = openpyxl.utils.get_column_letter(
                        col_idx
                    )
                    col_end_letter = openpyxl.utils.get_column_letter(
                        col_idx + 2
                    )

                    ws.merge_cells(
                        f"{col_start_letter}{row_idx}:{col_end_letter}{row_idx}"
                    )
                    cell_total = ws[f"{col_start_letter}{row_idx}"]
                    cell_total.value = f"{round(godziny_pracownikow[p], 1)}h"
                    cell_total.font = font_bold
                    cell_total.fill = fill_summary
                    cell_total.alignment = align_center

                    grand_total_hours += godziny_pracownikow[p]

                    for i in range(3):
                        ws.cell(row=row_idx, column=col_idx + i).border = (
                            thin_border
                        )

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
                cell_grand_val.value = (
                    f"{round(grand_total_hours, 1)} Roboczogodzin (RH)"
                )
                cell_grand_val.font = font_total_sum
                cell_grand_val.fill = fill_total_sum
                cell_grand_val.alignment = align_center

                for c in range(2, col_idx):
                    ws.cell(row=row_idx, column=c).border = thin_border

                ws.column_dimensions["A"].width = 16
                for c in range(2, col_idx):
                    col_letter = openpyxl.utils.get_column_letter(c)
                    ws.column_dimensions[col_letter].width = 10

                buffer = io.BytesIO()
                wb.save(buffer)

                st.success(
                    "⚡ Grafik Jush! wygenerowany pomyślnie! Zachowano zasady 4-5 dni pracy, 12h odpoczynku oraz wybalansowane otwarcia i zamknięcia."
                )

                st.download_button(
                    label="📥 Pobierz Grafik Jush! (.xlsx)",
                    data=buffer.getvalue(),
                    file_name="grafik_jush_ds.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except Exception as e:
            st.error(f"⚠️ Wystąpił szczegółowy błąd: {e}")
