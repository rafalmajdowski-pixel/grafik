from datetime import datetime, timedelta
import io
import math
import pandas as pd
import pulp
import streamlit as st

st.set_page_config(
    page_title="Optymalizator Grafiku Magazynu", layout="wide"
)

st.title("📦 Optymalizator Grafiku Magazynu")

# --- SIDEBAR: PARAMETRY EFEKTYWNOŚCI I OBSADY ---
st.sidebar.header("⚙️ Parametry Magazynu")

typ_magazynu = st.sidebar.selectbox("Typ magazynu", ["Standardowy", "Nocny"])
is_nocny = typ_magazynu == "Nocny"

godzina_otwarcia = 18 if is_nocny else 6

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
st.header("2. Wgraj raport z Lookera (tabela przestawna godzina/dzień)")
uploaded_file = st.file_uploader(
    "Wybierz plik CSV lub Excel pobrany z Lookera", type=["csv", "xlsx"]
)

srednie_godzinowe = {}  # {Dzien_Nazwa: {godzina: srednia_zamowien}}
srednie_wolumeny_dzienne = {}

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file)

        # Identyfikacja kolumny z godzinami (Hour of Day)
        col_hour = None
        for c in df_raw.columns:
            if "hour" in str(c).lower() or "godz" in str(c).lower():
                col_hour = c
                break
        if not col_hour:
            col_hour = df_raw.columns[0]

        # Identyfikacja kolumn z datami
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

        # Analiza wiersz po wierszu dla każdej godziny
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

        # Wyliczenie średnich godzinowych oraz dziennych
        for d_nazwa in MAPA_DNI.values():
            srednie_godzinowe[d_nazwa] = {}
            suma_dzienna = 0
            for h in range(24):
                vals = godziny_data[d_nazwa][h]
                sr_h = sum(vals) / len(vals) if vals else 0
                srednie_godzinowe[d_nazwa][h] = sr_h
                suma_dzienna += sr_h
            srednie_wolumeny_dzienne[d_nazwa] = suma_dzienna

        st.success(
            "✅ Pomyślnie przeanalizowano rozkład godzinowy zamówień z Lookera!"
        )

        with st.expander("🔍 Podgląd średniego zapotrzebowania godzinowego"):
            df_godz_view = pd.DataFrame(srednie_godzinowe)
            st.dataframe(df_godz_view.style.highlight_max(axis=0))

    except Exception as e:
        st.error(f"Błąd odczytu pliku: {e}")

# --- 3. PRACOWNICI I KALENDARZ URLOPOWY ---
st.header("3. Zespół i Wolne / Urlopy")
pracownicy_input = st.text_area(
    "Lista pracowników (każdy w nowej linii):",
    "Jan Kowalski\nPiotr Nowak\nAnna Wiśniewska",
)
pracownicy = [
    p.strip() for p in pracownicy_input.split("\n") if p.strip() != ""
]

if "urlopy_list" not in st.session_state:
    st.session_state.urlopy_list = []

st.subheader("Dodaj urlop lub dzień wolny dla pracownika:")
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

# --- 4. OPTYMALIZACJA GRAFIKU DLA SZCZYTÓW GODZINOWYCH ---
st.header("4. Generowanie Grafiku")
if st.button("🚀 Wygeneruj Grafik", type="primary"):
    if not uploaded_file:
        st.error("Proszę najpierw wgrać plik z Lookera!")
    elif not pracownicy:
        st.error("Proszę wpisać listę pracowników!")
    else:
        # POBIERANIE POTRZEB GODZINOWYCH (ILE OSÓB W DANEJ GODZINIE)
        # Np. 18 zam / 15 cel = 1.2 -> wymagane min 2 osoby w tej godzinie
        wymagani_pracownicy_h = {}
        for d in dni_zakresu:
            d_nazwa = MAPA_DNI.get(d.strftime("%A"), d.strftime("%A"))
            wymagani_pracownicy_h[d] = {}
            for h in range(24):
                sr_zam = srednie_godzinowe.get(d_nazwa, {}).get(h, 0)
                potrzeba_osob = math.ceil(sr_zam / cel_efektywnosci)
                wymagani_pracownicy_h[d][h] = potrzeba_osob

        # MODEL OPTYMALIZACYJNY
        model = pulp.LpProblem("Optymalizacja_Grafiku", pulp.LpMinimize)

        # Zmienne: czy p pracuje w d, zaczynając o godzinie s i trwając l godzin
        mozliwe_starty = range(
            godzina_otwarcia, godzina_otwarcia + 8
        )  # Płynne starty
        mozliwe_dlugosci = range(min_zmiana, max_zmiana + 1)

        zmienne_zmian = []
        for p in pracownicy:
            for d in dni_zakresu:
                for s in mozliwe_starty:
                    for l in mozliwe_dlugosci:
                        zmienne_zmian.append((p, d, s % 24, l))

        y = pulp.LpVariable.dicts("zmiana", zmienne_zmian, cat="Binary")

        dev_plus = pulp.LpVariable.dicts(
            "dev_plus", pracownicy, lowBound=0, cat="Continuous"
        )
        dev_minus = pulp.LpVariable.dicts(
            "dev_minus", pracownicy, lowBound=0, cat="Continuous"
        )

        # Ograniczenie: max 1 zmiana dziennie per pracownik
        for p in pracownicy:
            for d in dni_zakresu:
                model += (
                    pulp.lpSum(
                        y[p, d, s % 24, l]
                        for s in mozliwe_starty
                        for l in mozliwe_dlugosci
                    )
                    <= 1
                )

                # Uwzględnienie urlopów
                for u in st.session_state.urlopy_list:
                    if u["Pracownik"] == p and u["Od"] <= d <= u["Do"]:
                        for s in mozliwe_starty:
                            for l in mozliwe_dlugosci:
                                model += y[p, d, s % 24, l] == 0

        # Pokrycie zapotrzebowania w KAŻDEJ GODZINIE dnia
        for d in dni_zakresu:
            for h in range(24):
                potrzebni = wymagani_pracownicy_h[d][h]
                if potrzebni > 0:
                    pracujacy_w_godzinie = []
                    for p in pracownicy:
                        for s in mozliwe_starty:
                            s_int = s % 24
                            for l in mozliwe_dlugosci:
                                # Sprawdzenie czy godzina h wpada w zakres [s_int, s_int + l]
                                if s_int <= h < s_int + l or (
                                    s_int + l > 24 and h < (s_int + l) % 24
                                ):
                                    pracujacy_w_godzinie.append(
                                        y[p, d, s_int, l]
                                    )

                    model += pulp.lpSum(pracujacy_w_godzinie) >= potrzebni

        # Równomierny podział ogólnych godzin pracy
        for p in pracownicy:
            dni_wolne_count = sum(
                1
                for d in dni_zakresu
                for u in st.session_state.urlopy_list
                if u["Pracownik"] == p and u["Od"] <= d <= u["Do"]
            )
            cel_godzin = max(0, (len(dni_zakresu) * 8) - (dni_wolne_count * 8))
            suma_h = pulp.lpSum(
                y[p, d, s % 24, l] * l
                for d in dni_zakresu
                for s in mozliwe_starty
                for l in mozliwe_dlugosci
            )
            model += suma_h + dev_minus[p] - dev_plus[p] == cel_godzin

        model += pulp.lpSum(dev_plus[p] + dev_minus[p] for p in pracownicy)
        model.solve(pulp.PULP_CBC_CMD(msg=False))

        # Tabela wyników
        tabela = []
        godziny_pracownikow = {p: 0 for p in pracownicy}

        for d in dni_zakresu:
            dzien_nazwa = MAPA_DNI.get(d.strftime("%A"), d.strftime("%A"))
            row = {"Data": d.strftime("%Y-%m-%d"), "Dzień": dzien_nazwa}
            obsada_dnia_rh = 0

            for p in pracownicy:
                assigned = False
                for s in mozliwe_starty:
                    s_int = s % 24
                    for l in mozliwe_dlugosci:
                        if y[p, d, s_int, l].varValue == 1:
                            end_h = (s_int + l) % 24
                            row[p] = f"{s_int:02d}:00 - {end_h:02d}:00 ({l}h)"
                            godziny_pracownikow[p] += l
                            obsada_dnia_rh += l
                            assigned = True
                            break
                    if assigned:
                        break
                if not assigned:
                    row[p] = "OFF"

            wol = srednie_wolumeny_dzienne.get(dzien_nazwa, 0)
            wymagane_rh = wol / cel_efektywnosci if cel_efektywnosci > 0 else 0

            row["Suma RH"] = round(obsada_dnia_rh, 1)
            row["Wymagane RH"] = round(wymagane_rh, 1)
            row["Śr. Zamówień (Looker)"] = round(wol, 1)
            row["Plan. Efektywność (zam/h)"] = (
                round(wol / obsada_dnia_rh, 1) if obsada_dnia_rh > 0 else 0
            )

            tabela.append(row)

        row_sum = {
            "Data": "ŁĄCZNIE",
            "Dzień": "-",
            **{p: f"{int(godziny_pracownikow[p])}h" for p in pracownicy},
        }
        row_sum["Suma RH"] = round(sum(r["Suma RH"] for r in tabela), 1)
        row_sum["Wymagane RH"] = round(sum(r["Wymagane RH"] for r in tabela), 1)
        row_sum["Śr. Zamówień (Looker)"] = round(
            sum(r["Śr. Zamówień (Looker)"] for r in tabela), 1
        )
        row_sum["Plan. Efektywność (zam/h)"] = (
            round(
                row_sum["Śr. Zamówień (Looker)"] / row_sum["Suma RH"], 1
            )
            if row_sum["Suma RH"] > 0
            else 0
        )

        tabela.append(row_sum)
        df_res = pd.DataFrame(tabela)

        st.success("✅ Grafik wygenerowany pomyślnie z uwzględnieniem szczytów godzinowych!")
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
