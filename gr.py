import io
import pandas as pd
import pulp
import streamlit as st

st.set_page_config(page_title="Optymalizator Grafiku Magazynu", layout="wide")

st.title("📦 Optymalizator Grafiku Magazynu")

# --- SIDEBAR: KONFIGURACJA MAGAZYNU ---
st.sidebar.header("⚙️ Parametry Magazynu")
typ_magazynu = st.sidebar.selectbox(
    "Typ magazynu / Start zmiany",
    ["Standardowy (06:00)", "Popołudniowy (14:00)", "Nocny (22:00)"],
)
godzina_startu_baza = int(
    typ_magazynu.split("(")[1].split(":")[0]
)  # Pobiera np. 6, 14 lub 22

cel_efektywnosci = st.sidebar.number_input(
    "Cel efektywności (zamówienia/h na osobę)", min_value=1, value=15
)
min_zmiana = st.sidebar.slider("Minimalna zmiana (h)", 4, 8, 6)
max_zmiana = st.sidebar.slider("Maksymalna zmiana (h)", 8, 12, 8)
min_obsada = st.sidebar.number_input(
    "Minimalna liczba osób na zmianie", min_value=1, value=1
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

DNI_KOLEJNOSC = [
    "Poniedziałek",
    "Wtorek",
    "Środa",
    "Czwartek",
    "Piątek",
    "Sobota",
    "Niedziela",
]

# --- 1. WGRANIE DANYCH Z LOOKERA ---
st.header("1. Wgraj raport z Lookera")
uploaded_file = st.file_uploader(
    "Wybierz plik Excel (.xlsx) lub CSV z Lookera", type=["xlsx", "csv"]
)

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df_looker = pd.read_csv(uploaded_file)
    else:
        df_looker = pd.read_excel(uploaded_file)

    cols = df_looker.columns.tolist()
    col_data = st.selectbox(
        "Wybierz kolumnę z Datą lub Dniem Tygodnia:", cols, index=0
    )
    col_wolumen = st.selectbox(
        "Wybierz kolumnę z Wolumenem (liczba zamówień/paczek):",
        cols,
        index=1 if len(cols) > 1 else 0,
    )

    # Przetwarzanie dat
    df_looker["_dt"] = pd.to_datetime(df_looker[col_data], errors="coerce")
    if df_looker["_dt"].notna().any():
        df_looker["Dzien_Nazwa"] = (
            df_looker["_dt"].dt.day_name().map(MAPA_DNI)
        )
    else:
        df_looker["Dzien_Nazwa"] = df_looker[col_data].astype(str)

    srednie_wolumeny = (
        df_looker.groupby("Dzien_Nazwa")[col_wolumen].mean().to_dict()
    )

    # Obliczanie RH z uwzględnieniem minimalnej obsady
    zapotrzebowanie_rh = {}
    for d in DNI_KOLEJNOSC:
        wol = srednie_wolumeny.get(d, 0)
        calc_rh = wol / cel_efektywnosci
        # Jeśli jest jakikolwiek wolumen lub ustawiono min_obsada, przypisujemy min_zmiana * min_obsada
        if wol > 0 or min_obsada > 0:
            zapotrzebowanie_rh[d] = max(calc_rh, min_obsada * min_zmiana)
        else:
            zapotrzebowanie_rh[d] = 0.0

    # --- 2. PRACOWNICU I DNI NIEOBECNOŚCI (URLOPY) ---
    st.header("2. Pracownicy i Wybór Dni Urlopu / Nieobecności")
    pracownicy_input = st.text_area(
        "Lista pracowników (każdy w nowej linii):",
        "Jan Kowalski\nPiotr Nowak\nAnna Wiśniewska\nTomasz Zieliński\nMichał Lewandowski",
    )
    pracownicy = [
        p.strip() for p in pracownicy_input.split("\n") if p.strip() != ""
    ]

    st.subheader("Zaznacz konkretne dni nieobecności dla każdego pracownika:")
    urlopy_dni = {}
    if pracownicy:
        for p in pracownicy:
            urlopy_dni[p] = st.multiselect(
                f"Dni wolne / Urlop: **{p}**", DNI_KOLEJNOSC, key=f"u_{p}"
            )

    # --- 3. GENEROWANIE GRAFIKU ---
    st.header("3. Generowanie Grafiku")
    if st.button("🚀 Wygeneruj Zoptymalizowany Grafik", type="primary"):
        if not pracownicy:
            st.error("Dodaj przynajmniej jednego pracownika!")
        else:
            model = pulp.LpProblem(
                "Optymalizacja_Grafiku_Magazyn", pulp.LpMinimize
            )

            x = pulp.LpVariable.dicts(
                "pracuje",
                [(p, d) for p in pracownicy for d in DNI_KOLEJNOSC],
                cat="Binary",
            )
            godziny = pulp.LpVariable.dicts(
                "godziny",
                [(p, d) for p in pracownicy for d in DNI_KOLEJNOSC],
                lowBound=0,
                upBound=max_zmiana,
            )

            dev_plus = pulp.LpVariable.dicts(
                "dev_plus", pracownicy, lowBound=0, cat="Continuous"
            )
            dev_minus = pulp.LpVariable.dicts(
                "dev_minus", pracownicy, lowBound=0, cat="Continuous"
            )

            # Ograniczenia zmian i urlopów
            for p in pracownicy:
                for d in DNI_KOLEJNOSC:
                    if d in urlopy_dni[p]:
                        # Blokada pracy w wybrany dzień urlopu
                        model += x[p, d] == 0
                        model += godziny[p, d] == 0
                    else:
                        model += godziny[p, d] >= min_zmiana * x[p, d]
                        model += godziny[p, d] <= max_zmiana * x[p, d]

            # Wymóg pokrycia zapotrzebowania RH i min obsady
            for d in DNI_KOLEJNOSC:
                model += (
                    pulp.lpSum(godziny[p, d] for p in pracownicy)
                    >= zapotrzebowanie_rh[d]
                )
                if zapotrzebowanie_rh[d] > 0:
                    model += (
                        pulp.lpSum(x[p, d] for p in pracownicy) >= min_obsada
                    )

            # RÓWNOMIERNY PODZIAŁ GODZIN (cel skorygowany o wybrane dni wolne)
            for p in pracownicy:
                dni_wolne_count = len(urlopy_dni[p])
                cel_godzin_p = max(0, 40 - (dni_wolne_count * 8))
                suma_godzin_p = pulp.lpSum(godziny[p, d] for d in DNI_KOLEJNOSC)
                model += suma_godzin_p + dev_minus[p] - dev_plus[p] == cel_godzin_p

            model += pulp.lpSum(dev_plus[p] + dev_minus[p] for p in pracownicy)
            model.solve(pulp.PULP_CBC_CMD(msg=False))

            # --- 4. FORMATOWANIE WYNIKÓW (GODZINY STARTU I KOŃCA) ---
            tabela_grafik = []
            godziny_suma_pracownikow = {p: 0 for p in pracownicy}

            for d in DNI_KOLEJNOSC:
                row = {"Dzień": d}
                suma_dnia_rh = 0

                for p in pracownicy:
                    val = godziny[p, d].varValue
                    if val and val >= min_zmiana:
                        g_start = godzina_startu_baza
                        g_end = (g_start + int(val)) % 24
                        row[p] = f"{g_start:02d}:00 - {g_end:02d}:00 ({int(val)}h)"
                        suma_dnia_rh += val
                        godziny_suma_pracownikow[p] += val
                    else:
                        row[p] = "OFF"

                wol = srednie_wolumeny.get(d, 0)
                row["Obsada (Osoby)"] = sum(
                    1 for p in pracownicy if row[p] != "OFF"
                )
                row["Suma Godzin (RH)"] = round(suma_dnia_rh, 1)
                row["Wymagane RH"] = round(zapotrzebowanie_rh[d], 1)
                row["Śr. Wolumen"] = round(wol, 1)
                row["Efektywność (Paczki/h)"] = (
                    round(wol / suma_dnia_rh, 1) if suma_dnia_rh > 0 else 0
                )

                tabela_grafik.append(row)

            # Wiersz podsumowujący (Suma indywidualna per pracownik)
            row_total = {
                "Dzień": "ŁĄCZNIE GODZIN",
                **{p: f"{int(godziny_suma_pracownikow[p])}h" for p in pracownicy},
            }
            row_total["Obsada (Osoby)"] = "-"
            row_total["Suma Godzin (RH)"] = sum(
                r["Suma Godzin (RH)"] for r in tabela_grafik
            )
            row_total["Wymagane RH"] = sum(
                r["Wymagane RH"] for r in tabela_grafik
            )
            row_total["Śr. Wolumen"] = sum(
                r["Śr. Wolumen"] for r in tabela_grafik
            )
            tot_rh = row_total["Suma Godzin (RH)"]
            row_total["Efektywność (Paczki/h)"] = (
                round(row_total["Śr. Wolumen"] / tot_rh, 1) if tot_rh > 0 else 0
            )

            tabela_grafik.append(row_total)
            df_grafik = pd.DataFrame(tabela_grafik)

            st.success("✅ Grafik został pomyślnie wygenerowany!")
            st.dataframe(df_grafik, use_container_width=True)

            # --- 5. EKSPORT DO EXCELA (.XLSX) ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_grafik.to_excel(
                    writer, sheet_name="Grafik Pracy", index=False
                )

            st.download_button(
                label="📥 Pobierz Grafik w formacie Excel (.xlsx)",
                data=output.getvalue(),
                file_name="zoptymalizowany_grafik_magazynu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
