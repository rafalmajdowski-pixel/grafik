from datetime import datetime, timedelta
import io
import math
import pandas as pd
import pulp
import streamlit as st

st.set_page_config(page_title="Optymalizator Grafiku Magazynu", layout="wide")

st.title("📦 Optymalizator Grafiku Magazynu")

# --- SIDEBAR: PARAMETRY EFEKTYWNOŚCI I OBSADY ---
st.sidebar.header("⚙️ Parametry Magazynu")

typ_magazynu = st.sidebar.selectbox("Typ magazynu", ["Standardowy", "Nocny"])
is_nocny = typ_magazynu == "Nocny"

# TWARDE RAMY CZASOWE
# Standard: Praca DS 06:00 - 23:30 (zamówienia 07:00 - 23:00)
# Nocny: Praca DS 06:00 - 01:30 (zamówienia 07:00 - 01:00)
godzina_otwarcia_ds = 6.0
godzina_zamkniecia_ds = 25.5 if is_nocny else 23.5  # 25.5h = 01:30 w nocy

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
    "Wskaż zakres od - do (kliknij początek i koniec w kalendarzu):",
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
st.header("2. Wgraj raport z Lookera (tabela przestawna)")
uploaded_file = st.file_uploader(
    "Wybierz plik CSV lub Excel pobrany z Lookera", type=["csv", "xlsx"]
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
            d: {h: [] for h in range(24)} for d in MAPA_DNI.values()
        }

        for idx, row in df_raw.iterrows():
            h_val = pd.to_numeric(row[col_hour], errors="coerce")
            if pd.notna(h_val) and 0 <= int(h_val) <= 23:
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
            for h in range(24):
                vals = godziny_data[d_nazwa][h]
                sr_h = sum(vals) / len(vals) if vals else 0
                srednie_godzinowe[d_nazwa][h] = sr_h

        st.success("✅ Dane z Lookera zostały pomyślnie przetworzone.")

    except Exception as e:
        st.error(f"Błąd odczytu pliku: {e}")

# --- 3. PRACOWNICI I KALENDARZ URLOPOWY ---
st.header("3. Dostępni Pracownicy (Zleceniobiorcy)")
pracownicy_input = st.text_area(
    "Lista pracowników na zlecenie (każdy w nowej linii):",
    "Jan Kowalski\nPiotr Nowak\nAnna Wiśniewska\nTomasz Zieliński\nMichał Lewandowski",
)
pracownicy = [
    p.strip() for p in pracownicy_input.split("\n") if p.strip() != ""
]

if "urlopy_list" not in st.session_state:
    st.session_state.urlopy_list = []

st.subheader("Niedostępność pracownika (urlop / brak dyspozycyjności):")
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

# --- 4. GENEROWANIE GRAFIKU Z PRECYZYJNYMI GODZINAMI DS ---
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
            for h in range(24):
                sr_zam = srednie_godzinowe.get(d_nazwa, {}).get(h, 0)
                potrzeba_osob = math.ceil(sr_zam / cel_efektywnosci)
                wymagani_pracownicy_h[d][h] = potrzeba_osob
                total_required_hours += potrzeba_osob

        model = pulp.LpProblem("Optymalizacja_Grafiku", pulp.LpMinimize)

        # GENEROWANIE ZMIAN Z KROKIEM 30 MINUT DO DOKŁADNEJ GODZINY ZAMKNIĘCIA DS
        prawidlowe_zmiany = []
        starty = [6.0 + 0.5 * i for i in range(int((22.0 - 6.0) * 2) + 1)]
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

        # RÓWNOMIERNY PODZIAŁ GODZIN ZLECENIOBIORCÓW
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

            for d in dni_zakresu:
                model += (
                    pulp.lpSum(y[p, d, s, l] for s, l in prawidlowe_zmiany) <= 1
                )

                for u in st.session_state.urlopy_list:
                    if u["Pracownik"] == p and u["Od"] <= d <= u["Do"]:
                        for s, l in prawidlowe_zmiany:
                            model += y[p, d, s, l] == 0

        # POKRYCIE SPŁYWU ZAMÓWIEŃ Z LOOKERA
        for d in dni_zakresu:
            for h in range(24):
                potrzebni = wymagani_pracownicy_h[d][h]
                if potrzebni > 0:
                    pracujacy_w_godzinie = []
                    for p in pracownicy:
                        for s, l in prawidlowe_zmiany:
                            if s <= h and (s + l) >= (h + 1):
                                pracujacy_w_godzinie.append(y[p, d, s, l])

                    model += pulp.lpSum(pracujacy_w_godzinie) >= potrzebni

        model.solve(pulp.PULP_CBC_CMD(msg=False))

        tabela = []
        godziny_pracownikow = {p: 0 for p in pracownicy}

        def format_time(h_float):
            h_int = int(h_float) % 24
            m_int = int(round((h_float - int(h_float)) * 60))
            return f"{h_int:02d}:{m_int:02d}"

        for d in dni_zakresu:
            dzien_nazwa = MAPA_DNI.get(d.strftime("%A"), d.strftime("%A"))
            row = {"Data": d.strftime("%Y-%m-%d"), "Dzień": dzien_nazwa}

            for p in pracownicy:
                assigned = False
                for s, l in prawidlowe_zmiany:
                    if y[p, d, s, l].varValue == 1:
                        s_str = format_time(s)
                        e_str = format_time(s + l)
                        l_str = f"{int(l)}h" if l.is_integer() else f"{l}h"
                        row[p] = f"{s_str} - {e_str} ({l_str})"
                        godziny_pracownikow[p] += l
                        assigned = True
                        break
                if not assigned:
                    row[p] = "OFF"

            tabela.append(row)

        row_sum = {
            "Data": "ŁĄCZNIE",
            "Dzień": "-",
            **{p: f"{round(godziny_pracownikow[p], 1)}h" for p in pracownicy},
        }

        tabela.append(row_sum)
        df_res = pd.DataFrame(tabela)

        st.success(
            "✅ Grafik wygenerowany pomyślnie! Zmiany zamykają się precyzyjnie o 23:30 (lub 01:30)."
        )
        st.dataframe(df_res, use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_res.to_excel(writer, sheet_name="Grafik Pracy", index=False)

        st.download_button(
            label="📥 Pobierz Grafik (.xlsx)",
            data=buffer.getvalue(),
            file_name="grafik_magazyn.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
