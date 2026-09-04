from datetime import datetime, timedelta
import io
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

# Godzina otwarcia magazynu
godzina_otwarcia = 18 if is_nocny else 6

cel_efektywnosci = st.sidebar.number_input(
    "Efektywność pakowania (zamówienia / h / osoba)", min_value=1, value=15
)

# Sztywne zakresy pracy zaszyte w algorytmie (6-12h)
min_zmiana = 6
max_zmiana = 12

min_obsada_otwarcie = st.sidebar.number_input(
    "Min. obsada na otwarciu (osoby)", min_value=1, value=1
)
min_obsada_zamkniecie = st.sidebar.number_input(
    "Min. obsada na zamknięciu (osoby)", min_value=1, value=1
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

# --- 2. WGRANIE DANYCH Z LOOKERA ---
st.header("2. Wgraj raport z Lookera (zamówienia z aplikacji)")
uploaded_file = st.file_uploader(
    "Wybierz plik Excel (.xlsx) lub CSV z danymi historycznymi",
    type=["xlsx", "csv"],
)

srednie_wolumeny = {}
if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_looker = pd.read_csv(uploaded_file)
        else:
            df_looker = pd.read_excel(uploaded_file)

        cols = df_looker.columns.tolist()
        col_data = st.selectbox("Wybierz kolumnę z Datą:", cols, index=0)

        num_cols = df_looker.select_dtypes(include=["number"]).columns.tolist()
        default_idx = (
            cols.index(num_cols[0]) if num_cols and num_cols[0] in cols else 0
        )

        col_wolumen = st.selectbox(
            "Wybierz kolumnę z Wolumenem (liczba zamówień):",
            cols,
            index=default_idx,
        )

        df_looker["_dt"] = pd.to_datetime(df_looker[col_data], errors="coerce")
        df_looker["Dzien_Nazwa"] = (
            df_looker["_dt"]
            .dt.day_name()
            .map(MAPA_DNI)
            .fillna(df_looker[col_data].astype(str))
        )
        df_looker["_wolumen_num"] = pd.to_numeric(
            df_looker[col_wolumen], errors="coerce"
        ).fillna(0)

        srednie_wolumeny = (
            df_looker.groupby("Dzien_Nazwa")["_wolumen_num"].mean().to_dict()
        )
        st.success("✅ Dane z Lookera zostały przetworzone pomyślnie.")
    except Exception as e:
        st.error(
            f"Błąd podczas odczytu pliku: {e}. Sprawdź wybór kolumn."
        )

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

# --- 4. OPTYMALIZACJA GRAFIKU ---
st.header("4. Generowanie Grafiku")
if st.button("🚀 Wygeneruj Grafik", type="primary"):
    if not uploaded_file:
        st.error("Proszę najpierw wgrać plik z Lookera!")
    elif not pracownicy:
        st.error("Proszę wpisać listę pracowników!")
    else:
        # WERYFIKACJA MOŻLIWOŚCI KADROWYCH
        brakujace_godziny = {}
        for d in dni_zakresu:
            dzien_nazwa = MAPA_DNI.get(d.strftime("%A"), d.strftime("%A"))
            wol = srednie_wolumeny.get(dzien_nazwa, 0)
            wymagane_rh = wol / cel_efektywnosci

            dostepni_dzisiaj = [
                p
                for p in pracownicy
                if not any(
                    u["Pracownik"] == p and u["Od"] <= d <= u["Do"]
                    for u in st.session_state.urlopy_list
                )
            ]

            max_mozliwe_rh = len(dostepni_dzisiaj) * max_zmiana

            if max_mozliwe_rh < wymagane_rh:
                roznica = round(wymagane_rh - max_mozliwe_rh, 1)
                brakujace_godziny[f"{d.strftime('%Y-%m-%d')} ({dzien_nazwa})"] = (
                    roznica,
                    len(dostepni_dzisiaj),
                )

        if brakujace_godziny:
            st.warning("⚠️ **ALERT BRAKU OBSADY W ZESPOLE!**")
            for dzien_key, (roznica, dostepni) in brakujace_godziny.items():
                st.error(
                    f"Dzień **{dzien_key}**: Dostępnych osób: **{dostepni}**. "
                    f"Brakuje co najmniej **{roznica} roboczogodzin (RH)** do obsługi wolumenu!"
                )

        # MODEL PU-LP
        model = pulp.LpProblem("Optymalizacja_Grafiku", pulp.LpMinimize)

        pracuje = pulp.LpVariable.dicts(
            "pracuje",
            [(p, d) for p in pracownicy for d in dni_zakresu],
            cat="Binary",
        )
        godziny = pulp.LpVariable.dicts(
            "godziny",
            [(p, d) for p in pracownicy for d in dni_zakresu],
            lowBound=0,
            upBound=max_zmiana,
        )

        dev_plus = pulp.LpVariable.dicts(
            "dev_plus", pracownicy, lowBound=0, cat="Continuous"
        )
        dev_minus = pulp.LpVariable.dicts(
            "dev_minus", pracownicy, lowBound=0, cat="Continuous"
        )

        for p in pracownicy:
            for d in dni_zakresu:
                model += godziny[p, d] >= min_zmiana * pracuje[p, d]
                model += godziny[p, d] <= max_zmiana * pracuje[p, d]

                for u in st.session_state.urlopy_list:
                    if u["Pracownik"] == p and u["Od"] <= d <= u["Do"]:
                        model += pracuje[p, d] == 0
                        model += godziny[p, d] == 0

        for d in dni_zakresu:
            dzien_nazwa = MAPA_DNI.get(d.strftime("%A"), d.strftime("%A"))
            wol = srednie_wolumeny.get(dzien_nazwa, 0)
            wymagane_rh = wol / cel_efektywnosci

            model += (
                pulp.lpSum(godziny[p, d] for p in pracownicy) >= wymagane_rh
            )

        for p in pracownicy:
            dni_wolne_count = sum(
                1
                for d in dni_zakresu
                for u in st.session_state.urlopy_list
                if u["Pracownik"] == p and u["Od"] <= d <= u["Do"]
            )
            cel_godzin = max(0, (len(dni_zakresu) * 8) - (dni_wolne_count * 8))
            suma_h = pulp.lpSum(godziny[p, d] for d in dni_zakresu)
            model += suma_h + dev_minus[p] - dev_plus[p] == cel_godzin

        model += pulp.lpSum(dev_plus[p] + dev_minus[p] for p in pracownicy)
        model.solve(pulp.PULP_CBC_CMD(msg=False))

        tabela = []
        godziny_pracownikow = {p: 0 for p in pracownicy}

        for d in dni_zakresu:
            dzien_nazwa = MAPA_DNI.get(d.strftime("%A"), d.strftime("%A"))
            row = {"Data": d.strftime("%Y-%m-%d"), "Dzień": dzien_nazwa}
            obsada_dnia_rh = 0

            stagger_offset = 0
            for p in pracownicy:
                val = godziny[p, d].varValue
                if val and val >= min_zmiana:
                    # Wyliczenie poprawnej godziny startu i końca
                    start_h = (godzina_otwarcia + stagger_offset) % 24
                    end_h_float = start_h + val
                    end_h = int(end_h_float) % 24
                    end_m = int((end_h_float - int(end_h_float)) * 60)

                    time_str = f"{start_h:02d}:00 - {end_h:02d}:{end_m:02d} ({round(val, 1)}h)"
                    row[p] = time_str

                    godziny_pracownikow[p] += val
                    obsada_dnia_rh += val

                    # Stopniowanie wejść kolejnych osób o 2h (max do godziny 12)
                    stagger_offset = (stagger_offset + 2) % 6
                else:
                    row[p] = "OFF"

            wol = srednie_wolumeny.get(dzien_nazwa, 0)
            row["Suma RH"] = round(obsada_dnia_rh, 1)
            row["Wymagane RH"] = round(wol / cel_efektywnosci, 1)
            row["Śr. Zamówień (Aplikacja)"] = round(wol, 1)
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
        row_sum["Śr. Zamówień (Aplikacja)"] = round(
            sum(r["Śr. Zamówień (Aplikacja)"] for r in tabela), 1
        )
        row_sum["Plan. Efektywność (zam/h)"] = (
            round(
                row_sum["Śr. Zamówień (Aplikacja)"] / row_sum["Suma RH"], 1
            )
            if row_sum["Suma RH"] > 0
            else 0
        )

        tabela.append(row_sum)
        df_res = pd.DataFrame(tabela)

        st.success("✅ Grafik wygenerowany pomyślnie!")
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
