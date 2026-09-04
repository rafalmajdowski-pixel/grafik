import streamlit as st
import pandas as pd
import pulp
import math
from datetime import datetime, timedelta

st.set_page_config(page_title="Uniwersalny Optymalizator Grafiku Magazynu", layout="wide")

st.title("📦 Uniwersalny Optymalizator Grafiku Pracy Magazynu")

# --- PANEL BOCZNY: TYP MAGAZYNU I PARAMETRY ---
st.sidebar.header("🏭 1. Konfiguracja Magazynu")

typ_magazynu = st.sidebar.selectbox(
    "Typ magazynu",
    ["Standardowy (6:00 - 23:30)", "Nocny (6:00 - 01:30)"]
)

# Konfiguracja ram czasowych
if typ_magazynu == "Standardowy (6:00 - 23:30)":
    godziny_etykiety = [f"{g}:00" for g in range(6, 23)] + ["23:00-23:30"]
    sztywne_slots = ["6:00", "23:00-23:30"]
    godziny_spadku = [f"{g}:00" for g in range(7, 23)]
else: # Nocny
    godziny_etykiety = [f"{g}:00" for g in range(6, 24)] + ["0:00", "1:00-1:30"]
    sztywne_slots = ["6:00", "1:00-1:30"]
    godziny_spadku = [f"{g}:00" for g in range(7, 24)] + ["0:00"]

NUM_SLOTS = len(godziny_etykiety)

st.sidebar.header("⚙️ 2. Parametry Pracy i Obsady")
efektywnosc_target = st.sidebar.number_input("Cel efektywności (zamówienia / h / osoba)", min_value=1, max_value=200, value=20)
data_start = st.sidebar.date_input("Data początkowa grafiku", datetime.now())
okres_dni = st.sidebar.selectbox("Okres grafiku", [7, 14, 28, 30, 31], index=0)

min_shift_len = st.sidebar.slider("Min. długość zmiany (h)", 6.0, 12.0, 6.0, step=0.5)
max_shift_len = st.sidebar.slider("Max. długość zmiany (h)", 6.0, 12.0, 12.0, step=0.5)

# --- KONFIGURACJA PRACOWNIKÓW ---
st.sidebar.header("👥 3. Lista Pracowników")
imiona_input = st.sidebar.text_area(
    "Wpisz nazwiska pracowników (po jednym w wierszu)",
    "Kowalski Jan\nNowak Piotr\nWiśniewski Adam\nWójcik Michał\nKrawczyk Bartłomiej\nZalewski Tomasz\nAdamczyk Łukasz\nZielieński Paweł"
)
pracownicy = [p.strip() for p in imiona_input.split("\n") if p.strip()]

dni_daty = [data_start + timedelta(days=i) for i in range(okres_dni)]
dni_etykiety = [d.strftime("%d/%m/%Y") for d in dni_daty]

# --- SYSTEM WYKLUCZEŃ / URLOPÓW ---
st.sidebar.header("🚫 4. Wykluczenia i Urlopy")
wykluczenia = []
with st.sidebar.expander("Dodaj nieobecności (wolne dni)"):
    selected_emp = st.selectbox("Pracownik", pracownicy)
    selected_date = st.selectbox("Dzień wolny", dni_etykiety)
    if st.button("Dodaj nieobecność"):
        st.session_state.setdefault("absences", []).append((selected_emp, selected_date))

if "absences" in st.session_state and st.session_state["absences"]:
    st.sidebar.write("Zaplanowane nieobecności:")
    for emp, dt in st.session_state["absences"]:
        st.sidebar.caption(f"❌ {emp} -> {dt}")
    if st.sidebar.button("Czyść nieobecności"):
        st.session_state["absences"] = []

# --- MODUŁ IMPORTU DANYCH Z LOOKERA ---
st.subheader("📥 1. Wgraj plik z Lookera (Excel / CSV)")
uploaded_file = st.file_uploader("Przeciągnij i upuść plik z Lookera (min. 2 pełne tygodnie)", type=["xlsx", "xls", "csv"])

domyslne_dane = {d: [0] + [50 for _ in range(NUM_SLOTS-2)] + [0] for d in dni_etykiety}
df_orders = pd.DataFrame(domyslne_dane, index=godziny_etykiety)

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            df_raw = pd.read_excel(uploaded_file)
        else:
            df_raw = pd.read_csv(uploaded_file)
            
        df_clean = df_raw.iloc[1:].copy()
        hour_col = df_clean.columns[0]
        date_cols = df_clean.columns[1:]
        
        df_long = df_clean.melt(id_vars=[hour_col], value_vars=date_cols, var_name="Date", value_name="Orders")
        df_long[hour_col] = pd.to_numeric(df_long[hour_col], errors='coerce')
        df_long['Orders'] = pd.to_numeric(df_long['Orders'], errors='coerce').fillna(0)
        df_long['Date'] = pd.to_datetime(df_long['Date'], errors='coerce')
        df_long['DayOfWeek'] = df_long['Date'].dt.dayofweek
        
        # Wyliczenie średniej historycznej dla danego dnia tygodnia i godziny
        avg_by_dow_hour = df_long.groupby(['DayOfWeek', hour_col])['Orders'].mean().to_dict()
        
        for d_obj, d_str in zip(dni_daty, dni_etykiety):
            dow = d_obj.weekday()
            for g_idx, g_label in enumerate(godziny_etykiety):
                if g_label in sztywne_slots:
                    df_orders.loc[g_label, d_str] = 0
                else:
                    h_val = int(g_label.split(":")[0])
                    srednia = avg_by_dow_hour.get((dow, h_val), 0)
                    df_orders.loc[g_label, d_str] = math.ceil(srednia)

        st.success("✅ Pomyślnie załadowano Lookera i przeliczono średnie historyczne dla poszczególnych dni tygodnia!")

    except Exception as e:
        st.error(f"⚠️ Błąd podczas odczytu struktury pliku: {e}")

st.subheader("📊 Prognoza zamówień na podstawie Lookera")
edited_orders = st.data_editor(df_orders, use_container_width=True)

# Szybkie wyliczenie sumy zamówień na każdy dzień
suma_zamowien_dzien = edited_orders.sum(axis=0).to_dict()

# --- OBLICZANIE ZAPOTRZEBOWANIA ---
wymagania = {}
for d in dni_etykiety:
    req_day = []
    for g_label in godziny_etykiety:
        if g_label in sztywne_slots:
            req_day.append(1) # Sztywno min 1 osoba
        else:
            paczki = edited_orders.loc[g_label, d]
            req_day.append(max(1, math.ceil(paczki / efektywnosc_target)))
    wymagania[d] = req_day

# --- GENEROWANIE GRAFIKU (OPTYMALIZACJA) ---
if st.button("🚀 Wygeneruj Optymalny Grafik Pracy", type="primary"):
    
    możliwe_zmiany = []
    for start in range(NUM_SLOTS):
        for end in range(start + 1, NUM_SLOTS + 1):
            dur = (end - start) if end < NUM_SLOTS else (end - start - 0.5)
            if min_shift_len <= dur <= max_shift_len:
                możliwe_zmiany.append((start, end, dur))

    prob = pulp.LpProblem("Magazyn_Grafik_Master", pulp.LpMinimize)
    x = pulp.LpVariable.dicts("Shift", ((d, p, s) for d in dni_etykiety for p in pracownicy for s in range(len(możliwe_zmiany))), cat=pulp.LpBinary)
    
    # Funkcja celu: Minimalizacja sumy roboczogodzin
    prob += pulp.lpSum(x[d, p, s] * możliwe_zmiany[s][2] for d in dni_etykiety for p in pracownicy for s in range(len(możliwe_zmiany)))

    absences_list = st.session_state.get("absences", [])

    for d_idx, d in enumerate(dni_etykiety):
        for p in pracownicy:
            # Uwzględnienie wykluczeń / urlopów
            if (p, d) in absences_list:
                prob += pulp.lpSum(x[d, p, s] for s in range(len(możliwe_zmiany))) == 0
            else:
                prob += pulp.lpSum(x[d, p, s] for s in range(len(możliwe_zmiany))) <= 1
        
        # Ograniczenie pokrycia zapotrzebowania
        for t in range(NUM_SLOTS):
            working = [x[d, p, s] for p in pracownicy for s, (st_slot, end_slot, _) in enumerate(możliwe_zmiany) if st_slot <= t < end_slot]
            prob += pulp.lpSum(working) >= wymagania[d][t]

        # Ograniczenie minimum 12h odpoczynku między zmianami
        if d_idx < len(dni_etykiety) - 1:
            next_d = dni_etykiety[d_idx + 1]
            for p in pracownicy:
                for s1_idx, (_, end1, _) in enumerate(możliwe_zmiany):
                    end_hour = 6.0 + (end1 if end1 < NUM_SLOTS else (NUM_SLOTS - 0.5))
                    for s2_idx, (start2, _, _) in enumerate(możliwe_zmiany):
                        start_hour_next_day = 24.0 + 6.0 + start2
                        if (start_hour_next_day - end_hour) < 12.0:
                            prob += x[d, p, s1_idx] + x[next_d, p, s2_idx] <= 1

    solver = pulp.PULP_CBC_CMD(msg=False)
    prob.solve(solver)

    if pulp.LpStatus[prob.status] == "Optimal":
        
        # --- TWORZENIE MACIERZY GRAFIKU ---
        columns_workers = pd.MultiIndex.from_product([pracownicy, ["Start", "Koniec", "Suma"]])
        column_totals = pd.MultiIndex.from_tuples([("Suma Dnia (h)", "Total"), ("Efektywność Dnia", "paczki/h")])
        columns_all = columns_workers.append(column_totals)

        df_result = pd.DataFrame(index=dni_etykiety, columns=columns_all)

        total_orders_all_days = sum(suma_zamowien_dzien.values())
        total_hours_all_days = 0.0
        emp_monthly_hours = {p: 0.0 for p in pracownicy}

        for d in dni_etykiety:
            day_total_hours = 0.0
            for p in pracownicy:
                worked = False
                for s_idx, (st_slot, end_slot, dur) in enumerate(możliwe_zmiany):
                    if pulp.value(x[d, p, s_idx]) == 1:
                        s_time = f"{6 + st_slot:02d}:00"
                        if end_slot == NUM_SLOTS:
                            e_time = "23:30" if typ_magazynu.startswith("Standardowy") else "01:30"
                        else:
                            e_time = f"{(6 + end_slot) % 24:02d}:00"
                        
                        df_result.loc[d, (p, "Start")] = s_time
                        df_result.loc[d, (p, "Koniec")] = e_time
                        df_result.loc[d, (p, "Suma")] = f"{dur:.1f}"
                        day_total_hours += dur
                        emp_monthly_hours[p] += dur
                        worked = True
                        break
                if not worked:
                    df_result.loc[d, (p, "Start")] = ""
                    df_result.loc[d, (p, "Koniec")] = ""
                    df_result.loc[d, (p, "Suma")] = ""
            
            total_hours_all_days += day_total_hours
            df_result.loc[d, ("Suma Dnia (h)", "Total")] = f"{day_total_hours:.1f}"
            
            # Wyliczenie efektywności dziennej
            eff_day = (suma_zamowien_dzien[d] / day_total_hours) if day_total_hours > 0 else 0.0
            df_result.loc[d, ("Efektywność Dnia", "paczki/h")] = f"{eff_day:.1f}"

        # DODANIE WIERSZA PODSUMOWANIA (SUMA GODZIN DLA KAŻDEGO PRACOWNIKA)
        summary_row_name = "ŁĄCZNIE GODZIN"
        df_result.loc[summary_row_name] = ""
        for p in pracownicy:
            df_result.loc[summary_row_name, (p, "Suma")] = f"{emp_monthly_hours[p]:.1f}"
        
        df_result.loc[summary_row_name, ("Suma Dnia (h)", "Total")] = f"{total_hours_all_days:.1f}"
        
        # Wyliczenie łącznej efektywności
        total_eff = (total_orders_all_days / total_hours_all_days) if total_hours_all_days > 0 else 0.0
        df_result.loc[summary_row_name, ("Efektywność Dnia", "paczki/h")] = f"{total_eff:.1f}"

        # --- WYŚWIETLANIE WYNIKÓW ---
        st.markdown("---")
        st.subheader("📈 Efektywność Całego Grafiku")
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Łączna liczba zamówień", f"{total_orders_all_days} szt.")
        col_m2.metric("Suma przepracowanych godzin", f"{total_hours_all_days:.1f} h")
        col_m3.metric("Średnia Efektywność Grafiku", f"{total_eff:.2f} paczek / h / osoba")

        st.subheader("🗓️ Gotowy Grafik Pracy dla Kierownika")
        st.dataframe(df_result, use_container_width=True)

        csv_data = df_result.to_csv().encode('utf-8')
        st.download_button("📥 Pobierz grafik do Excela (CSV)", data=csv_data, file_name="grafik_magazyn_master.csv", mime="text/csv")

    else:
        st.error("❌ Brak możliwości ułożenia grafiku. Zwiększ liczbę pracowników lub usuń część wykluczeń/urlopów w panelu bocznym.")