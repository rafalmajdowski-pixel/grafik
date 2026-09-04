import io
import pandas as pd
import pulp
import streamlit as st

st.set_page_config(page_title="Optymalizator Grafiku Magazynu", layout="wide")

st.title("📦 Optymalizator Grafiku Magazynu")

# --- SIDEBAR: KONFIGURACJA MAGAZYNU ---
st.sidebar.header("⚙️ Parametry Magazynu")
typ_magazynu = st.sidebar.selectbox("Typ magazynu", ["Standardowy", "Nocny"])
cel_efektywnosci = st.sidebar.number_input(
    "Cel efektywności (paczki/h)", min_value=1, value=20
)
min_zmiana = st.sidebar.slider("Minimalna zmiana (h)", 4, 8, 6)
max_zmiana = st.sidebar.slider("Maksymalna zmiana (h)", 8, 12, 10)

# Mapa dni roboczych po polsku i angielsku
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

    st.write("📋 **Podgląd wgranych danych z Lookera:**")
    st.dataframe(df_looker.head(), use_container_width=True)

    # Identyfikacja kolumn
    cols = df_looker.columns.tolist()
    col_data = st.selectbox(
        "Wybierz kolumnę z Datą lub Dniem Tygodnia:",
        cols,
        index=0 if len(cols) > 0 else 0,
    )
    col_wolumen = st.selectbox(
        "Wybierz kolumnę z Wolumenem (liczba paczek):",
        cols,
        index=1 if len(cols) > 1 else 0,
    )

    # Przekształcanie dat i grupowanie dni tygodnia
    df_looker["_dt"] = pd.to_datetime(df_looker[col_data], errors="coerce")

    if df_looker["_dt"].notna().any():
        df_looker["Dzien_Nazwa"] = (
            df_looker["_dt"].dt.day_name().map(MAPA_DNI)
        )
    else:
        df_looker["Dzien_Nazwa"] = df_looker[col_data].astype(str)

    # Porównywanie tych samych dni (Poniedziałek z Poniedziałkiem) i wyliczanie średniej
    srednie_wolumeny = (
        df_looker.groupby("Dzien_Nazwa")[col_wolumen].mean().to_dict()
    )

    # Zapotrzebowanie w roboczogodzinach (RH) per dzień
    zapotrzebowanie_rh = {
        d: round(srednie_wolumeny.get(d, 0) / cel_efektywnosci, 1)
        for d in DNI_KOLEJNOSC
    }

    # Pokazanie wyliczonego zapotrzebowania
    st.subheader("📈 Średnie zapotrzebowanie na Roboczogodziny (RH)")
    df_rh_show = pd.DataFrame(
        [
            {
                "Dzień": d,
                "Średni Wolumen": round(srednie_wolumeny.get(d, 0), 1),
                "Zapotrzebowanie (RH)": zapotrzebowanie_rh[d],
            }
            for d in DNI_KOLEJNOSC
        ]
    )
    st.dataframe(df_rh_show, use_container_width=True)

    # --- 2. PRACOWNICY I URLOPY ---
    st.header("2. Pracownicy i Urlopy")
    pracownicy_input = st.text_area(
        "Lista pracowników (każdy w nowej linii):",
        "Jan Kowalski\nPiotr Nowak\nAnna Wiśniewska\nTomasz Zieliński\nMichał Lewandowski",
    )
    pracownicy = [
        p.strip() for p in pracownicy_input.split("\n") if p.strip() != ""
    ]

    st.subheader("Dni urlopu w danym tygodniu (0 = brak urlopu)")
    urlopy_dict = {}
    if pracownicy:
        cols_p = st.columns(min(len(pracownicy), 5))
        for idx, p in enumerate(pracownicy):
            with cols_p[idx % len(cols_p)]:
                urlopy_dict[p] = st.number_input(
                    f"{p}", min_value=0, max_value=7, value=0, key=f"u_{p}"
                )

    # --- 3. GENEROWANIE GRAFIKU (PuLP) ---
    st.header("3. Generowanie Grafiku")
    if st.button("🚀 Wygeneruj Zoptymalizowany Grafik", type="primary"):
        if not pracownicy:
            st.error("Dodaj przynajmniej jednego pracownika!")
        else:
            model = pulp.LpProblem(
                "Optymalizacja_Grafiku_Magazyn", pulp.LpMinimize
            )

            # Zmienne: czy pracuje (binarna) oraz ile godzin (ciągła)
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

            # Zmienne odchyleń od równej liczby godzin
            dev_plus = pulp.LpVariable.dicts(
                "dev_plus", pracownicy, lowBound=0, cat="Continuous"
            )
            dev_minus = pulp.LpVariable.dicts(
                "dev_minus", pracownicy, lowBound=0, cat="Continuous"
            )

            # Ograniczenia długości zmian
            for p in pracownicy:
                for d in DNI_KOLEJNOSC:
                    model += godziny[p, d] >= min_zmiana * x[p, d]
                    model += godziny[p, d] <= max_zmiana * x[p, d]

            # Pokrycie dziennego zapotrzebowania na RH
            for d in DNI_KOLEJNOSC:
                model += (
                    pulp.lpSum(godziny[p, d] for p in pracownicy)
                    >= zapotrzebowanie_rh[d]
                )

            # RÓWNOMIERNY PODZIAŁ GODZIN (z korektą na urlopy)
            # Cel bazowy = 40h - (dni urlopu * 8h)
            for p in pracownicy:
                cel_godzin_p = max(0, 40 - (urlopy_dict[p] * 8))
                suma_godzin_p = pulp.lpSum(godziny[p, d] for d in DNI_KOLEJNOSC)
                model += suma_godzin_p + dev_minus[p] - dev_plus[p] == cel_godzin_p

            # FUNKCJA CELU: Minimalizacja różnic w obciążeniu godzinowym
            model += pulp.lpSum(dev_plus[p] + dev_minus[p] for p in pracownicy)

            model.solve(pulp.PULP_CBC_CMD(msg=False))

            # --- 4. TABELA WYNIKOWA ---
            tabela_grafik = []
            for d in DNI_KOLEJNOSC:
                row = {"Dzień tygodnia": d}
                suma_dnia = 0
                for p in pracownicy:
                    val = godziny[p, d].varValue
                    val_clean = round(val, 1) if val and val > 0.1 else 0
                    row[p] = val_clean
                    suma_dnia += val_clean
                row["Suma Dnia (RH)"] = round(suma_dnia, 1)
                row["Wymagane (RH)"] = zapotrzebowanie_rh[d]
                tabela_grafik.append(row)

            # Wiersz podsumowania łącznego
            row_total = {"Dzień tygodnia": "ŁĄCZNIE GODZIN"}
            for p in pracownicy:
                row_total[p] = round(
                    sum(
                        row[p]
                        for row in tabela_grafik
                        if isinstance(row[p], (int, float))
                    ),
                    1,
                )
            row_total["Suma Dnia (RH)"] = sum(
                r["Suma Dnia (RH)"] for r in tabela_grafik
            )
            row_total["Wymagane (RH)"] = sum(
                r["Wymagane (RH)"] for r in tabela_grafik
            )
            tabela_grafik.append(row_total)

            df_grafik = pd.DataFrame(tabela_grafik)

            st.success("✅ Grafik został pomyślnie wygenerowany!")
            st.dataframe(df_grafik, use_container_width=True)

            # --- 5. GENEROWANIE PLIKU EXCEL (.XLSX) ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_grafik.to_excel(
                    writer, sheet_name="Grafik Pracy", index=False
                )
                df_rh_show.to_excel(
                    writer, sheet_name="Podsumowanie RH", index=False
                )

            excel_data = output.getvalue()

            st.download_button(
                label="📥 Pobierz Grafik w formacie Excel (.xlsx)",
                data=excel_data,
                file_name="zoptymalizowany_grafik_magazynu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
