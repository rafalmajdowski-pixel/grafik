from datetime import datetime, timedelta
import io
import math
import openpyxl
from openpyxl.styles import Alignment, Border, PatternFill, Side, Font
import pandas as pd
import pulp
import streamlit as st

st.set_page_config(page_title="Optymalizator Grafiku Magazynu", layout="wide")

st.title("📦 Optymalizator Grafiku Magazynu")

# --- SIDEBAR: PARAMETRY EFEKTYWNOŚCI I OBSADY ---
st.sidebar.header("⚙️ Parametry Magazynu")

typ_magazynu = st.sidebar.selectbox("Typ magazynu", ["Standardowy", "Nocny"])
is_nocny = typ_magazynu == "Nocny"

godzina_otwarcia_ds = 6.0
godzina_zamkniecia_ds = 25.5 if is_nocny else 23.5  # 25.5h = 01:30 w nocy
max_godzina_zamowien = 25 if is_nocny else 23       # 25h = 01:00 w nocy

cel_efektywnosci = st.sidebar.number_input(
    "Efektywność pakowania (zamówienia / h / osoba)", min_value=1, value=15
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
    value=(datetime.now().date(), datetime.now().date() + timedelta(days=6)),
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
st.header("2. Wgraj raport z Lookera")
uploaded_file = st.file_uploader(
    "Wybierz plik CSV lub Excel z Lookera", type=["csv", "xlsx"]
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

        st.success("✅ Dane z Lookera przetworzone pomyślnie.")

    except Exception as e:
        st.error(f"Błąd odczytu pliku: {e}")

# --- 3. PRACOWNICI I KALENDARZ URLOPOWY ---
st.header("3. Dostępni Pracownicy")
pracownicy_input = st.text_area(
    "Lista pracowników na zlecenie (każdy w nowej linii):",
    "Aval01204VasinA\nAval01209KushnY\nAvalZhukoD\nDive01202VitalD\nEter01203SavchV\nEterZaichI",
)
pracownicy = [
    p.strip() for p in pracownicy_input.split("\n") if p.strip() != ""
]

if "urlopy_list" not in st.session_state:
    st.session_state.urlopy_list = []

st.subheader("Niedostępność pracownika:")
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

# --- 4. GENEROWANIE GRAFIKU Z OGRANICZENIEM 12H ODPOZCZYNKU ---
st.header("4. Generowanie Grafiku")
if st.button("🚀 Wygeneruj Grafik", type="primary"):
    if not uploaded_file:
        st.error("Proszę najpierw wgrać plik z Lookera!")
    elif not pracownicy:
        st.error("Proszę wpisać listę pracowników!")
    else:
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

        zmienne_zmian = []
        for p in pracownicy:
            for d in dni_zakresu:
                for s, l in prawidlowe_zmiany:
                    zmienne_zmian.append((p, d, s, l))

        y = pulp.LpVariable.dicts("zmiana", zmienne_zmian, cat="Binary")

        srednia_godzin_na_glowe = total_required_hours / len(pracownicy)
        dev_plus = pulp.LpVariable.dicts(
            "dev_plus", pracownicy, lowBound=0, cat="Continuous"
        )
        dev_minus = pulp.LpVariable.dicts(
            "dev_minus", pracownicy, lowBound=0, cat="Continuous"
        )

        model += pulp.lpSum(dev_plus[p] + dev_minus[p] for p in pracownicy)

        for p in pracownicy:
            suma_h_p = pulp.lpSum(
                y[p, d, s, l] * l
                for d in dni_zakresu
                for s, l in prawidlowe_zmiany
            )
            model += (
                suma_h_p + dev_minus[p] - dev_plus[p] == srednia_godzin_na_glowe
            )

            for idx_d, d in enumerate(dni_zakresu):
                model += (
                    pulp.lpSum(y[p, d, s, l] for s, l in prawidlowe_zmiany) <= 1
                )

                for u in st.session_state.urlopy_list:
                    if u["Pracownik"] == p and u["Od"] <= d <= u["Do"]:
                        for s, l in prawidlowe_zmiany:
                            model += y[p, d, s, l] == 0

                # DZIENNY ODPOCZYNEK MINIMUM 12 GODZIN Z DNIA NA DZIEŃ
                if idx_d < len(dni_zakresu) - 1:
                    d_next = dni_zakresu[idx_d + 1]
                    for s1, l1 in prawidlowe_zmiany:
                        koniec_dzien1 = s1 + l1
                        for s2, l2 in prawidlowe_zmiany:
                            start_dzien2 = s2 + 24.0
                            if (start_dzien2 - koniec_dzien1) < 12.0:
                                model += y[p, d, s1, l1] + y[p, d_next, s2, l2] <= 1

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

        # WERYFIKACJA CZY ROZWIAZANIE JEST MOŻLIWE (LACK OF STAFF / 12H REST CONFLICT)
        if pulp.LpStatus[status] != "Optimal":
            st.error(
                "❌ Nie można wygenerować grafiku! Zbiór reguł (minimum 12h odpoczynku, urlopy, godziny DS) "
                "oraz zapotrzebowanie z Lookera przekraczają możliwości dostępnej liczby pracowników. "
                "Dodaj więcej pracowników na zlecenie lub zmień zakres urlopów."
            )
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Grafik"
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

            fill_header_main = PatternFill(
                start_color="8EA9DB", end_color="8EA9DB", fill_type="solid"
            )
            fill_header_sub = PatternFill(
                start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
            )
            fill_summary = PatternFill(
                start_color="B4C6E7", end_color="B4C6E7", fill_type="solid"
            )
            fill_total_sum = PatternFill(
                start_color="70AD47", end_color="70AD47", fill_type="solid"
            )
            fill_date_weekend = PatternFill(
                start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
            )
            fill_date_weekday = PatternFill(
                start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
            )

            fill_shift_morning = PatternFill(
                start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
            )
            fill_shift_afternoon = PatternFill(
                start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
            )

            ws.merge_cells("A1:A2")
            ws["A1"] = "pl-waw-12"
            ws["A1"].font = font_bold
            ws["A1"].fill = fill_header_main
            ws["A1"].alignment = align_center

            def format_time(h_float):
                h_int = int(h_float) % 24
                m_int = int(round((h_float - int(h_float)) * 60))
                return f"{h_int:02d}:{m_int:02d}"

            col_idx = 2
            for p in pracownicy:
                col_start_letter = openpyxl.utils.get_column_letter(col_idx)
                col_end_letter = openpyxl.utils.get_column_letter(col_idx + 2)

                ws.merge_cells(f"{col_start_letter}1:{col_end_letter}1")
                cell_p = ws[f"{col_start_letter}1"]
                cell_p.value = p
                cell_p.font = font_bold
                cell_p.fill = fill_header_main
                cell_p.alignment = align_center

                sub_headers = ["Start", "Koniec", "Suma"]
                for i, sh in enumerate(sub_headers):
                    cell_sh = ws.cell(row=2, column=col_idx + i)
                    cell_sh.value = sh
                    cell_sh.font = font_bold
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
                                if s <= 10.0
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
                col_start_letter = openpyxl.utils.get_column_letter(col_idx)
                col_end_letter = openpyxl.utils.get_column_letter(col_idx + 2)

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
            cell_grand_val.font = Font(
                name="Calibri", size=11, bold=True, color="FFFFFF"
            )
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
                "✅ Grafik wygenerowany pomyślnie! Zachowano wymagane 12h odpoczynku między zmianami."
            )

            st.download_button(
                label="📥 Pobierz Grafik (.xlsx)",
                data=buffer.getvalue(),
                file_name="grafik_magazyn_12h_rest.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
