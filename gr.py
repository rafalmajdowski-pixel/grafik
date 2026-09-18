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
    
    /* Panel Boczny */
    [data-testid="stSidebar"] {
        background-color: #005B2B !important;
    }
    [data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: #8BC53F !important;
    }
    
    /* Stylizacja elementów w panelu */
    [data-testid="stSidebar"] input, [data-testid="stSidebar"] select, [data-testid="stSidebar"] div[data-baseweb="select"] {
        color: #000000 !important;
        background-color: #FFFFFF !important;
        border-radius: 8px !important;
    }
    
    /* Przycisk Pobierania */
    div.stButton > button {
        background-color: #8BC53F !important;
        color: #005B2B !important;
        font-weight: 800 !important;
        font-size: 16px !important;
        border-radius: 10px !important;
        border: none !important;
        padding: 12px 28px !important;
        width: 100%;
    }
    div.stButton > button:hover {
        background-color: #A3DF52 !important;
        color: #004420 !important;
    }
    
    h1, h2, h3 {
        color: #005B2B !important;
        font-family: 'Arial Black', sans-serif !important;
    }
    </style>
""",
    unsafe_allow_html=True,
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

# --- PANEL BOCZNY (SIDEBAR) ---
with st.sidebar:
    st.image("https://zabkagroup.com/wp-content/uploads/2022/09/Jush_logo.png", width=120)
    st.title("⚡ Konfiguracja DS")
    
    st.markdown("---")
    st.subheader("1. Typ Magazynu")
    typ_magazynu = st.selectbox("Wybierz tryb pracy:", ["Standardowy (06:00 - 23:30)", "Nocny (06:00 - 01:30)"])
    is_nocny = "Nocny" in typ_magazynu
    godzina_zamkniecia_ds = 25.5 if is_nocny else 23.5

    st.markdown("---")
    st.subheader("2. Wydajność Pickera")
    cel_efektywnosci = st.slider(
        "Zamówienia / h / picker:",
        min_value=1,
        max_value=30,
        value=10,
        step=1,
    )

    st.markdown("---")
    st.subheader("3. Zrzut / Plik z Lookera")
    uploaded_file = st.file_uploader(
        "Przeciągnij lub wybierz zrzut Lookera (PNG, JPG, CSV):",
        type=["png", "jpg", "jpeg", "csv", "xlsx"]
    )

    st.markdown("---")
    st.subheader("4. Okres Grafiku")
    okres_grafiku = st.date_input(
        "Zakres dat od - do:",
        value=(datetime.now().date(), datetime.now().date() + timedelta(days=29)),
    )

if isinstance(okres_grafiku, tuple) and len(okres_grafiku) == 2:
    start_date, end_date = okres_grafiku
    dni_zakresu = [start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)]
else:
    dni_zakresu = [okres_grafiku[0]]

# --- GŁÓWNA CZĘŚĆ EKRANU ---
st.markdown(
    "<h1 style='margin-bottom:0; font-size: 2.2rem;'>żabka <span style='color:#8BC53F;'>jush!</span></h1>",
    unsafe_allow_html=True,
)
st.markdown("<p style='font-weight:bold; color:#005B2B; font-size: 1.1rem;'>Generator Szkieletu Grafiku (Shift Skeleton)</p>", unsafe_allow_html=True)
st.divider()

mock_looker_matrix = {
    "Wednesday": {7: 7, 8: 8, 9: 10, 10: 10, 11: 12, 12: 9, 13: 15, 14: 14, 15: 14, 16: 13, 17: 16, 18: 21, 19: 19, 20: 26, 21: 18, 22: 8},
    "Thursday": {7: 6, 8: 8, 9: 9, 10: 12, 11: 8, 12: 11, 13: 14, 14: 13, 15: 14, 16: 15, 17: 14, 18: 17, 19: 25, 20: 23, 21: 18, 22: 10},
    "Friday": {7: 8, 8: 8, 9: 10, 10: 12, 11: 12, 12: 11, 13: 11, 14: 14, 15: 15, 16: 16, 17: 16, 18: 22, 19: 23, 20: 21, 21: 21, 22: 10},
    "Saturday": {7: 7, 8: 12, 9: 16, 10: 24, 11: 23, 12: 20, 13: 26, 14: 20, 15: 23, 16: 24, 17: 27, 18: 33, 19: 31, 20: 31, 21: 25, 22: 9},
    "Sunday": {7: 9, 8: 15, 9: 24, 10: 24, 11: 35, 12: 37, 13: 31, 14: 33, 15: 33, 16: 33, 17: 34, 18: 39, 19: 33, 20: 39, 21: 25, 22: 12},
    "Monday": {7: 6, 8: 12, 9: 14, 10: 18, 11: 18, 12: 17, 13: 17, 14: 18, 15: 21, 16: 20, 17: 23, 18: 34, 19: 29, 20: 30, 21: 24, 22: 10},
    "Tuesday": {7: 4, 8: 10, 9: 11, 10: 16, 11: 18, 12: 15, 13: 17, 14: 17, 15: 20, 16: 21, 17: 26, 18: 32, 19: 39, 20: 30, 21: 26, 22: 9},
    "Środa": {7: 7, 8: 8, 9: 10, 10: 10, 11: 12, 12: 9, 13: 15, 14: 14, 15: 14, 16: 13, 17: 16, 18: 21, 19: 19, 20: 26, 21: 18, 22: 8},
    "Czwartek": {7: 6, 8: 8, 9: 9, 10: 12, 11: 8, 12: 11, 13: 14, 14: 13, 15: 14, 16: 15, 17: 14, 18: 17, 19: 25, 20: 23, 21: 18, 22: 10},
    "Piątek": {7: 8, 8: 8, 9: 10, 10: 12, 11: 12, 12: 11, 13: 11, 14: 14, 15: 15, 16: 16, 17: 16, 18: 22, 19: 23, 20: 21, 21: 21, 22: 10},
    "Sobota": {7: 7, 8: 12, 9: 16, 10: 24, 11: 23, 12: 20, 13: 26, 14: 20, 15: 23, 16: 24, 17: 27, 18: 33, 19: 31, 20: 31, 21: 25, 22: 9},
    "Niedziela": {7: 9, 8: 15, 9: 24, 10: 24, 11: 35, 12: 37, 13: 31, 14: 33, 15: 33, 16: 33, 17: 34, 18: 39, 19: 33, 20: 39, 21: 25, 22: 12},
    "Poniedziałek": {7: 6, 8: 12, 9: 14, 10: 18, 11: 18, 12: 17, 13: 17, 14: 18, 15: 21, 16: 20, 17: 23, 18: 34, 19: 29, 20: 30, 21: 24, 22: 10},
    "Wtorek": {7: 4, 8: 10, 9: 11, 10: 16, 11: 18, 12: 15, 13: 17, 14: 17, 15: 20, 16: 21, 17: 26, 18: 32, 19: 39, 20: 30, 21: 26, 22: 9},
}

srednie_godzinowe = {d: {h: 0.0 for h in range(26)} for d in list(MAPA_DNI.values()) + list(MAPA_DNI.keys())}

for d_name, h_dict in mock_looker_matrix.items():
    for h_val, val in h_dict.items():
        srednie_godzinowe[d_name][h_val] = float(val)

def format_time(h_float):
    h_int = int(h_float) % 24
    m_int = int(round((h_float - int(h_float)) * 60))
    return f"{h_int:02d}:{m_int:02d}"

skeleton_rows = []
max_slots_found = 0

dozwolone_zmiany = []
for s in [6.0 + 0.5 * i for i in range(int((18.0 - 6.0) * 2) + 1)]:
    for l in [float(x)/2.0 for x in range(12, 25)]:
        if s + l <= godzina_zamkniecia_ds:
            dozwolone_zmiany.append((s, l, s + l))

for d in dni_zakresu:
    d_nazwa_en = d.strftime("%A")
    d_nazwa_pl = MAPA_DNI.get(d_nazwa_en, d_nazwa_en)
    
    row_dict = {
        "Dzień": d_nazwa_pl,
        "Dzień Msc": d.day,
    }
    
    # Przeliczanie zapotrzebowania z uwzględnieniem faktu, że h oznacza przedział [h, h+1)
    req_pickers = {}
    for h in range(6, int(math.ceil(godzina_zamkniecia_ds))):
        orders_h = srednie_godzinowe.get(d_nazwa_en, {}).get(h, 0)
        if orders_h == 0:
            orders_h = srednie_godzinowe.get(d_nazwa_pl, {}).get(h, 0)
        
        needed = math.ceil(orders_h / cel_efektywnosci)
        needed = max(1, needed) # Minimum 1 picker na DS w dobie
        req_pickers[h] = needed

    prob = pulp.LpProblem("Szkielet_DS", pulp.LpMinimize)
    x = pulp.LpVariable.dicts("slot", range(len(dozwolone_zmiany)), lowBound=0, cat="Integer")
    
    kara_symetrii = []
    for i, (s, l, e) in enumerate(dozwolone_zmiany):
        odchylenie = abs(l - 7.5) * 0.1
        kara_symetrii.append(x[i] * (l + odchylenie))

    prob += pulp.lpSum(kara_symetrii)
    
    # Warunki pokrycia popytu w każdym półgodzinnym kroku doby
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
        row_dict[f"Osoba {slot_idx+1} - Zmiana"] = f"{format_time(s)} - {format_time(e)}"
        row_dict[f"Osoba {slot_idx+1} - Suma"] = f"{dur:.1f}h"

    row_dict["Suma Dnia (RH)"] = f"{sum_day_rh:.1f}h"
    skeleton_rows.append(row_dict)

df_skeleton = pd.DataFrame(skeleton_rows).fillna("-")

st.subheader("📐 Wygenerowany Szkielet Grafiku (Osoba 1, Osoba 2...)")
st.dataframe(df_skeleton, use_container_width=True, hide_index=True)

# --- TWORZENIE FORMOWANEGO EXCELA (.XLSX) ---
wb_sk = openpyxl.Workbook()
ws_sk = wb_sk.active
ws_sk.title = "Szkielet Grafiku"

font_bold = Font(name="Calibri", size=10, bold=True)
font_regular = Font(name="Calibri", size=10)
align_center = Alignment(horizontal="center", vertical="center")

fill_header_main = PatternFill(start_color="005B2B", end_color="005B2B", fill_type="solid")
font_header_main = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

fill_header_sub = PatternFill(start_color="8BC53F", end_color="8BC53F", fill_type="solid")
font_header_sub = Font(name="Calibri", size=10, bold=True, color="005B2B")

fill_shift = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
fill_sunday = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
fill_summary = PatternFill(start_color="EBF7D4", end_color="EBF7D4", fill_type="solid")
fill_total = PatternFill(start_color="8BC53F", end_color="8BC53F", fill_type="solid")

thin_border = Border(
    left=Side(style="thin", color="D3D3D3"),
    right=Side(style="thin", color="D3D3D3"),
    top=Side(style="thin", color="D3D3D3"),
    bottom=Side(style="thin", color="D3D3D3"),
)

ws_sk.merge_cells("A1:B2")
ws_sk["A1"] = "pl-waw-12"
ws_sk["A1"].font = font_header_main
ws_sk["A1"].fill = fill_header_main
ws_sk["A1"].alignment = align_center

col_idx = 3
for i in range(1, max_slots_found + 1):
    c_start_let = openpyxl.utils.get_column_letter(col_idx)
    c_end_let = openpyxl.utils.get_column_letter(col_idx + 1)
    
    ws_sk.merge_cells(f"{c_start_let}1:{c_end_let}1")
    cell_p = ws_sk[f"{c_start_let}1"]
    cell_p.value = f"Osoba {i}"
    cell_p.font = font_header_main
    cell_p.fill = fill_header_main
    cell_p.alignment = align_center

    for j, sh in enumerate(["Godziny Zmiany", "Suma"]):
        cell_sh = ws_sk.cell(row=2, column=col_idx + j)
        cell_sh.value = sh
        cell_sh.font = font_header_sub
        cell_sh.fill = fill_header_sub
        cell_sh.alignment = align_center
        cell_sh.border = thin_border

    col_idx += 2

ws_sk.merge_cells(f"{openpyxl.utils.get_column_letter(col_idx)}1:{openpyxl.utils.get_column_letter(col_idx)}2")
c_sum_head = ws_sk.cell(row=1, column=col_idx)
c_sum_head.value = "Suma Dnia (RH)"
c_sum_head.font = font_header_main
c_sum_head.fill = fill_header_main
c_sum_head.alignment = align_center

slot_sum_rh = {i: 0.0 for i in range(1, max_slots_found + 1)}
grand_total_rh = 0.0
row_idx = 3

for r_data in skeleton_rows:
    cell_day = ws_sk.cell(row=row_idx, column=1, value=r_data["Dzień"])
    cell_num = ws_sk.cell(row=row_idx, column=2, value=r_data["Dzień Msc"])
    
    for c in [cell_day, cell_num]:
        c.font = font_regular
        c.alignment = align_center
        c.border = thin_border

    if r_data["Dzień"] == "Niedziela":
        cell_day.fill = fill_sunday
        cell_num.fill = fill_sunday

    col_c = 3
    day_total = 0.0
    for slot_i in range(1, max_slots_found + 1):
        z_val = r_data.get(f"Osoba {slot_i} - Zmiana", "-")
        rh_val = r_data.get(f"Osoba {slot_i} - Suma", "-")

        c_z = ws_sk.cell(row=row_idx, column=col_c, value=z_val)
        c_rh = ws_sk.cell(row=row_idx, column=col_c+1, value=rh_val)

        if rh_val != "-":
            c_z.fill = fill_shift
            val_h = float(rh_val.replace("h", ""))
            slot_sum_rh[slot_i] += val_h
            day_total += val_h

        for c in [c_z, c_rh]:
            c.font = font_bold if c == c_rh else font_regular
            c.alignment = align_center
            c.border = thin_border

        col_c += 2

    cell_day_total = ws_sk.cell(row=row_idx, column=col_c, value=f"{day_total:.1f}h")
    cell_day_total.font = font_bold
    cell_day_total.fill = fill_summary
    cell_day_total.alignment = align_center
    cell_day_total.border = thin_border
    
    grand_total_rh += day_total
    row_idx += 1

cell_sum_label = ws_sk.cell(row=row_idx, column=1)
cell_sum_label.value = "ŁĄCZNIE"
cell_sum_label.font = font_bold
cell_sum_label.fill = fill_summary
cell_sum_label.alignment = align_center
cell_sum_label.border = thin_border

col_c = 3
for slot_i in range(1, max_slots_found + 1):
    c_let1 = openpyxl.utils.get_column_letter(col_c)
    c_let2 = openpyxl.utils.get_column_letter(col_c + 1)
    ws_sk.merge_cells(f"{c_let1}{row_idx}:{c_let2}{row_idx}")
    
    c_tot_p = ws_sk[f"{c_let1}{row_idx}"]
    c_tot_p.value = f"{slot_sum_rh[slot_i]:.1f}h"
    c_tot_p.font = font_bold
    c_tot_p.fill = fill_summary
    c_tot_p.alignment = align_center
    
    for k in range(2):
        ws_sk.cell(row=row_idx, column=col_c + k).border = thin_border
        
    col_c += 2

ws_sk.cell(row=row_idx, column=col_c).border = thin_border
row_idx += 1

cell_grand_label = ws_sk.cell(row=row_idx, column=1)
cell_grand_label.value = "SUMA CAŁKOWITA"
cell_grand_label.font = font_bold
cell_grand_label.fill = fill_total
cell_grand_label.alignment = align_center
cell_grand_label.border = thin_border

last_col_let = openpyxl.utils.get_column_letter(col_c)
ws_sk.merge_cells(f"B{row_idx}:{last_col_let}{row_idx}")
c_grand_val = ws_sk[f"B{row_idx}"]
c_grand_val.value = f"{grand_total_rh:.1f} Roboczogodzin (RH)"
c_grand_val.font = Font(name="Calibri", size=11, bold=True, color="005B2B")
c_grand_val.fill = fill_total
c_grand_val.alignment = align_center

for k in range(2, col_c + 1):
    ws_sk.cell(row=row_idx, column=k).border = thin_border

buf_sk = io.BytesIO()
wb_sk.save(buf_sk)

st.markdown("---")
st.download_button(
    label="📥 Pobierz Gotowy Grafiku (.xlsx)",
    data=buf_sk.getvalue(),
    file_name="szkielet_grafiku_ds.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
