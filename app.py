"""
Kniha jízd - generátor
Streamlit app pro automatizaci tvorby měsíční knihy jízd
"""
import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# KONFIGURACE - destinace a vzdálenosti (uprav podle potřeby)
# ============================================================
DEFAULT_DESTINACE = {
    "Praha 7 Jankovcova": 0,
    "SKLC3 (Šakoňská cesta)": 340,
    "Bratislava - hotel": 330,
    "Brno": 205,
    "Jihlava": 130,
    "Olomouc": 280,
    "Ostrava": 370,
    "Hradec Králové": 115,
    "Pardubice": 105,
    "České Budějovice": 150,
    "Plzeň": 95,
    "Karlovy Vary": 130,
    "Liberec": 105,
    "Ústí nad Labem": 90,
    "Mladá Boleslav": 70,
    "Kladno": 30,
    "Beroun": 30,
    "Mělník": 40,
    "Říčany": 25,
    "Poděbrady": 50,
    "Týnec nad Labem": 90,
    "Prostějov": 270,
    "Chomutov": 95,
    "Trnava": 380,
    "Nitra": 410,
    "Praha 4 - Pankrác": 8,
    "Praha 4 Olbrachtova (pumpa)": 8,
    "Praha HP II (pumpa)": 12,
    "Chotýčany D3 (pumpa)": 100,
    "Zelenáč D2 (pumpa)": 250,
}

# Vzdálenosti mezi vybranými páry (pokud nejede z/do Jankovcové)
KRIZOVE_VZDALENOSTI = {
    ("SKLC3 (Šakoňská cesta)", "Bratislava - hotel"): 12,
    ("Bratislava - hotel", "SKLC3 (Šakoňská cesta)"): 12,
    ("SKLC3 (Šakoňská cesta)", "Trnava"): 50,
    ("Trnava", "SKLC3 (Šakoňská cesta)"): 50,
    ("SKLC3 (Šakoňská cesta)", "Brno"): 130,
    ("Brno", "SKLC3 (Šakoňská cesta)"): 130,
    ("Hradec Králové", "Pardubice"): 25,
    ("Pardubice", "Hradec Králové"): 25,
    ("Hradec Králové", "Olomouc"): 145,
    ("Olomouc", "Ostrava"): 100,
    ("Brno", "Olomouc"): 80,
    ("Brno", "Jihlava"): 90,
    ("Plzeň", "Karlovy Vary"): 85,
    ("České Budějovice", "Plzeň"): 130,
    ("Chotýčany D3 (pumpa)", "České Budějovice"): 50,
    ("Zelenáč D2 (pumpa)", "SKLC3 (Šakoňská cesta)"): 80,
    ("Zelenáč D2 (pumpa)", "Bratislava - hotel"): 90,
}

UCEL_OPTIONS = [
    "Pracovní jednání",
    "Pracovní schůzka",
    "Pracovní cesta",
    "Schůzka u klienta",
    "Návrat",
    "Návrat + tankování",
    "Cesta + tankování",
    "Ubytování",
    "Závěrečné jednání",
]

# ============================================================
# FUNKCE
# ============================================================
def vzdalenost(odkud, kam):
    """Spočítá vzdálenost mezi dvěma body"""
    if (odkud, kam) in KRIZOVE_VZDALENOSTI:
        return KRIZOVE_VZDALENOSTI[(odkud, kam)]
    # Z/do Jankovcové
    if odkud == "Praha 7 Jankovcova":
        return st.session_state.destinace.get(kam, 0)
    if kam == "Praha 7 Jankovcova":
        return st.session_state.destinace.get(odkud, 0)
    # Pokud nic, vrátíme rozdíl od Jankovcové (hrubý odhad)
    od = st.session_state.destinace.get(odkud, 0)
    do = st.session_state.destinace.get(kam, 0)
    return abs(od - do)


def init_state():
    """Inicializace session state"""
    if "destinace" not in st.session_state:
        st.session_state.destinace = DEFAULT_DESTINACE.copy()
    if "tankovani" not in st.session_state:
        st.session_state.tankovani = []
    if "jizdy" not in st.session_state:
        st.session_state.jizdy = []
    if "stav_pocatek" not in st.session_state:
        st.session_state.stav_pocatek = 136030


def parsuj_tankovani_excel(uploaded_file):
    """Parsuje Excel s tankováními - flexibilně hledá sloupce"""
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]

    # Hledáme sloupce s datem, stavem KM, množstvím a místem
    sloupce = {col.lower(): col for col in df.columns}

    def najdi(klic_castecne):
        for low, orig in sloupce.items():
            if klic_castecne in low:
                return orig
        return None

    col_datum = najdi("datum")
    col_km = najdi("stav km") or najdi("stav")
    col_litry = najdi("množství") or najdi("mnozstvi") or najdi("litry")
    col_misto = najdi("čerpací") or najdi("cerpaci") or najdi("pumpa") or najdi("místo")
    col_cena = najdi("částka") or najdi("castka") or najdi("cena")

    tankovani = []
    for _, row in df.iterrows():
        try:
            datum = row[col_datum] if col_datum else None
            km = int(row[col_km]) if col_km and pd.notna(row[col_km]) else None
            litry = float(row[col_litry]) if col_litry and pd.notna(row[col_litry]) else None
            misto = str(row[col_misto]) if col_misto else ""
            cena = float(row[col_cena]) if col_cena and pd.notna(row[col_cena]) else None
            if km and litry:
                tankovani.append({
                    "datum": datum,
                    "km": km,
                    "litry": litry,
                    "misto": misto.strip(),
                    "cena": cena,
                })
        except Exception:
            continue
    # Seřadíme podle KM
    tankovani.sort(key=lambda x: x["km"])
    return tankovani


def export_excel(jizdy, tankovani_info, mesic_rok, stav_pocatek):
    """Vytvoří Excel knihu jízd"""
    wb = Workbook()
    ws = wb.active
    ws.title = mesic_rok

    # Hlavička metadat
    ws['A1'] = "Typ vozidla:"
    ws['B1'] = "doplnit"
    ws['A2'] = "RZ (SPZ) vozidla:"
    ws['B2'] = "5SD7805"
    ws['A3'] = "Palivo:"
    ws['B3'] = "Diesel"
    ws['A4'] = "Spotřeba paliva dle VTP"
    ws['B4'] = 3.7
    ws['A5'] = "Počáteční stav tachometru:"
    ws['B5'] = stav_pocatek

    final_km = jizdy[-1]["km_konec"] if jizdy else stav_pocatek
    ws['A6'] = "Konečný stav tachometru:"
    ws['B6'] = final_km
    ws['A7'] = "Soukromé km:"
    ws['B7'] = 0

    for i in range(1, 8):
        ws[f'A{i}'].font = Font(name='Arial', bold=True, size=10)

    # Hlavička tabulky
    headers = ['Datum', 'Výchozí místo', 'Cílové místo', 'Účel cesty',
               'Čas odjezdu', 'Čas příjezdu', 'Druh jízdy',
               'Stav km na začátku', 'Stav km na konci',
               'Počet kilometrů', 'Tankování litry', 'Cena phm', 'Řidič']
    yellow_fill = PatternFill('solid', start_color='FFF2CC')
    header_fill = PatternFill('solid', start_color='D9D9D9')
    thin = Side(border_style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    arial_bold = Font(name='Arial', bold=True, size=10)
    arial = Font(name='Arial', size=10)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    HEADER_ROW = 9
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=HEADER_ROW, column=col, value=h)
        c.font = arial_bold
        c.fill = header_fill
        c.alignment = center
        c.border = border

    # Data
    row = HEADER_ROW + 1
    for j in jizdy:
        values = [
            j["datum"],
            j["odkud"],
            j["kam"],
            j["ucel"],
            j.get("cas_odj", ""),
            j.get("cas_prij", ""),
            j["druh"],
            j["km_zacatek"],
            j["km_konec"],
            j["km_pocet"],
            j.get("litry"),
            j.get("cena"),
            j.get("ridic", "MaWy"),
        ]
        for col, v in enumerate(values, 1):
            c = ws.cell(row=row, column=col, value=v)
            c.font = arial
            c.border = border
            if col == 10:
                c.fill = yellow_fill
            if col >= 5:
                c.alignment = Alignment(horizontal='center', vertical='center')
        row += 1

    # Šířky sloupců
    widths = [11, 24, 28, 26, 11, 11, 11, 16, 16, 13, 13, 11, 9]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[HEADER_ROW].height = 32

    # Uložení do bufferu
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ============================================================
# UI
# ============================================================
st.set_page_config(page_title="Kniha jízd", page_icon="📕", layout="wide")
init_state()

st.title("📕 Kniha jízd – generátor")
st.caption("Automatizace měsíčního zápisu knihy jízd")

# --- KROK 1: Nastavení měsíce
with st.expander("⚙️ Krok 1 – nastavení měsíce", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        mesic = st.selectbox("Měsíc",
            ["leden", "únor", "březen", "duben", "květen", "červen",
             "červenec", "srpen", "září", "říjen", "listopad", "prosinec"],
            index=3)
    with col2:
        rok = st.number_input("Rok", min_value=2024, max_value=2030, value=2026)
    with col3:
        st.session_state.stav_pocatek = st.number_input(
            "Počáteční stav tachometru (km)",
            min_value=0,
            value=st.session_state.stav_pocatek,
            step=1,
        )

# --- KROK 2: Nahrání tankování
with st.expander("⛽ Krok 2 – výpis tankování", expanded=True):
    st.markdown(
        "Stáhni si výpis z Teams tabulky jako Excel a nahraj sem. "
        "Appka automaticky najde stavy KM, litry, ceny a místa."
    )
    uploaded = st.file_uploader("Nahrát Excel s tankováními", type=["xlsx", "xls"])

    if uploaded:
        try:
            st.session_state.tankovani = parsuj_tankovani_excel(uploaded)
            st.success(f"✓ Načteno {len(st.session_state.tankovani)} tankování")
        except Exception as e:
            st.error(f"Chyba při načítání: {e}")

    # Manuální přidání tankování
    if st.button("➕ Přidat tankování ručně"):
        st.session_state.tankovani.append({
            "datum": date.today(),
            "km": 0,
            "litry": 0.0,
            "misto": "",
            "cena": 0.0,
        })

    if st.session_state.tankovani:
        st.markdown("**Načtená tankování (lze upravit):**")
        for i, t in enumerate(st.session_state.tankovani):
            cols = st.columns([2, 3, 2, 2, 2, 1])
            with cols[0]:
                t["datum"] = st.date_input(
                    "Datum", value=t["datum"] if isinstance(t["datum"], (date, datetime)) else date.today(),
                    key=f"t_datum_{i}", label_visibility="collapsed"
                )
            with cols[1]:
                t["misto"] = st.text_input("Místo", value=t["misto"], key=f"t_misto_{i}", label_visibility="collapsed")
            with cols[2]:
                t["km"] = st.number_input("KM", value=int(t["km"]), key=f"t_km_{i}", label_visibility="collapsed")
            with cols[3]:
                t["litry"] = st.number_input("Litry", value=float(t["litry"]), step=0.01, key=f"t_litry_{i}", label_visibility="collapsed")
            with cols[4]:
                t["cena"] = st.number_input("Cena", value=float(t["cena"] or 0), step=0.01, key=f"t_cena_{i}", label_visibility="collapsed")
            with cols[5]:
                if st.button("🗑", key=f"t_del_{i}"):
                    st.session_state.tankovani.pop(i)
                    st.rerun()

# --- KROK 3: Jízdy
with st.expander("🚗 Krok 3 – jízdy", expanded=True):
    if not st.session_state.tankovani:
        st.info("Nejdřív nahraj nebo přidej tankování.")
    else:
        # Výpočet úseků mezi tankováními
        prev_km = st.session_state.stav_pocatek
        st.markdown("**Úseky mezi tankováními (musí přesně sedět):**")

        for i, t in enumerate(st.session_state.tankovani):
            cil_km = t["km"] - prev_km
            # Spočítat km zapsaných v jízdách v tomto úseku
            zapsano = sum(j["km_pocet"] for j in st.session_state.jizdy
                          if prev_km <= j["km_zacatek"] < t["km"])
            chybi = cil_km - zapsano
            if chybi == 0:
                st.success(f"✓ Úsek do {t['datum']} {t['misto']} (cíl {cil_km:,} km): zapsáno {zapsano:,} km – sedí")
            elif chybi > 0:
                st.error(f"✗ Úsek do {t['datum']} {t['misto']} (cíl {cil_km:,} km): zapsáno {zapsano:,} km – chybí {chybi:,} km")
            else:
                st.warning(f"⚠️ Úsek do {t['datum']} {t['misto']} (cíl {cil_km:,} km): zapsáno {zapsano:,} km – přebývá {-chybi:,} km")
            prev_km = t["km"]

        st.markdown("---")

        # Přidání jízdy
        st.markdown("**Přidat novou jízdu:**")
        c1, c2, c3, c4 = st.columns([2, 3, 3, 2])
        with c1:
            nove_datum = st.date_input("Datum", value=date(rok, 4, 1), key="nove_datum")
        with c2:
            odkud = st.selectbox("Odkud", list(st.session_state.destinace.keys()), key="nove_odkud")
        with c3:
            kam = st.selectbox("Kam", list(st.session_state.destinace.keys()), key="nove_kam", index=1)
        with c4:
            spocteno = vzdalenost(odkud, kam)
            km_pocet = st.number_input("KM", value=spocteno, min_value=0, key="nove_km")

        c5, c6, c7, c8 = st.columns([2, 2, 2, 2])
        with c5:
            cas_odj = st.text_input("Čas odjezdu", value="8:00", key="nove_odj")
        with c6:
            cas_prij = st.text_input("Čas příjezdu", value="9:30", key="nove_prij")
        with c7:
            ucel = st.selectbox("Účel", UCEL_OPTIONS, key="nove_ucel")
        with c8:
            druh = st.selectbox("Druh", ["služební", ""], key="nove_druh")

        if st.button("➕ Přidat jízdu", type="primary"):
            # Najdi navazující stav km (poslední km_konec, nebo počátek)
            if st.session_state.jizdy:
                last_km = st.session_state.jizdy[-1]["km_konec"]
            else:
                last_km = st.session_state.stav_pocatek

            nova = {
                "datum": nove_datum.strftime("%-d.%-m.%Y") if hasattr(nove_datum, "strftime") else str(nove_datum),
                "odkud": odkud,
                "kam": kam,
                "ucel": ucel,
                "cas_odj": cas_odj,
                "cas_prij": cas_prij,
                "druh": druh,
                "km_zacatek": last_km,
                "km_konec": last_km + km_pocet,
                "km_pocet": km_pocet,
            }

            # Pokud končí na pumpě, přiřadit tankování
            for t in st.session_state.tankovani:
                if t["km"] == nova["km_konec"]:
                    nova["litry"] = t["litry"]
                    nova["cena"] = t["cena"]
                    break

            st.session_state.jizdy.append(nova)
            st.rerun()

        # Tabulka jízd
        if st.session_state.jizdy:
            st.markdown("**Zapsané jízdy:**")
            df_view = pd.DataFrame(st.session_state.jizdy)
            st.dataframe(df_view, use_container_width=True, hide_index=True)

            if st.button("🗑️ Smazat poslední jízdu"):
                st.session_state.jizdy.pop()
                st.rerun()

            if st.button("🗑️ Smazat všechny jízdy"):
                st.session_state.jizdy = []
                st.rerun()

# --- Šablony rychlého přidání
with st.expander("📋 Šablony – rychlé přidání cesty", expanded=False):
    st.markdown("Tlačítka pro typické scénáře. Datum si pak doupravíš.")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**SKLC3 + přespání**")
        if st.button("Přidat cestu na SKLC3 (2 dny)"):
            datum1 = date(rok, 4, 1)
            datum2 = date(rok, 4, 2)
            last_km = st.session_state.jizdy[-1]["km_konec"] if st.session_state.jizdy else st.session_state.stav_pocatek
            st.session_state.jizdy.extend([
                {"datum": datum1.strftime("%-d.%-m.%Y"), "odkud": "Praha 7 Jankovcova",
                 "kam": "SKLC3 (Šakoňská cesta)", "ucel": "Pracovní cesta",
                 "cas_odj": "7:00", "cas_prij": "11:00", "druh": "služební",
                 "km_zacatek": last_km, "km_konec": last_km + 340, "km_pocet": 340},
                {"datum": datum1.strftime("%-d.%-m.%Y"), "odkud": "SKLC3 (Šakoňská cesta)",
                 "kam": "Bratislava - hotel", "ucel": "Ubytování",
                 "cas_odj": "17:00", "cas_prij": "17:30", "druh": "",
                 "km_zacatek": last_km + 340, "km_konec": last_km + 352, "km_pocet": 12},
                {"datum": datum2.strftime("%-d.%-m.%Y"), "odkud": "Bratislava - hotel",
                 "kam": "Praha 7 Jankovcova", "ucel": "Návrat",
                 "cas_odj": "13:00", "cas_prij": "17:30", "druh": "",
                 "km_zacatek": last_km + 352, "km_konec": last_km + 682, "km_pocet": 330},
            ])
            st.rerun()

    with col2:
        st.markdown("**Cesta do Brna**")
        if st.button("Přidat 1denní cestu Brno"):
            datum1 = date(rok, 4, 1)
            last_km = st.session_state.jizdy[-1]["km_konec"] if st.session_state.jizdy else st.session_state.stav_pocatek
            st.session_state.jizdy.extend([
                {"datum": datum1.strftime("%-d.%-m.%Y"), "odkud": "Praha 7 Jankovcova",
                 "kam": "Brno", "ucel": "Pracovní jednání",
                 "cas_odj": "7:00", "cas_prij": "9:30", "druh": "služební",
                 "km_zacatek": last_km, "km_konec": last_km + 205, "km_pocet": 205},
                {"datum": datum1.strftime("%-d.%-m.%Y"), "odkud": "Brno",
                 "kam": "Praha 7 Jankovcova", "ucel": "Návrat",
                 "cas_odj": "16:00", "cas_prij": "18:30", "druh": "",
                 "km_zacatek": last_km + 205, "km_konec": last_km + 410, "km_pocet": 205},
            ])
            st.rerun()

    with col3:
        st.markdown("**Cesta do CB**")
        if st.button("Přidat 1denní cestu ČB"):
            datum1 = date(rok, 4, 1)
            last_km = st.session_state.jizdy[-1]["km_konec"] if st.session_state.jizdy else st.session_state.stav_pocatek
            st.session_state.jizdy.extend([
                {"datum": datum1.strftime("%-d.%-m.%Y"), "odkud": "Praha 7 Jankovcova",
                 "kam": "České Budějovice", "ucel": "Pracovní jednání",
                 "cas_odj": "8:00", "cas_prij": "10:30", "druh": "služební",
                 "km_zacatek": last_km, "km_konec": last_km + 150, "km_pocet": 150},
                {"datum": datum1.strftime("%-d.%-m.%Y"), "odkud": "České Budějovice",
                 "kam": "Praha 7 Jankovcova", "ucel": "Návrat",
                 "cas_odj": "15:00", "cas_prij": "17:30", "druh": "",
                 "km_zacatek": last_km + 150, "km_konec": last_km + 300, "km_pocet": 150},
            ])
            st.rerun()

# --- Souhrnné statistiky
if st.session_state.jizdy:
    st.markdown("---")
    st.markdown("### 📊 Souhrn")
    col1, col2, col3, col4 = st.columns(4)
    najeto = sum(j["km_pocet"] for j in st.session_state.jizdy)
    litry = sum(t["litry"] for t in st.session_state.tankovani)
    cena = sum((t.get("cena") or 0) for t in st.session_state.tankovani)

    col1.metric("Najeto celkem", f"{najeto:,} km".replace(",", " "))
    col2.metric("Spotřeba", f"{litry:.2f} l")
    col3.metric("Cena PHM", f"{cena:,.0f} Kč".replace(",", " "))
    col4.metric("Konečný stav", f"{st.session_state.stav_pocatek + najeto:,} km".replace(",", " "))

# --- KROK 4: Export
st.markdown("---")
st.markdown("### ⬇️ Krok 4 – stáhnout knihu jízd")

if st.session_state.jizdy:
    excel_buffer = export_excel(
        st.session_state.jizdy,
        st.session_state.tankovani,
        f"{mesic} {rok}",
        st.session_state.stav_pocatek,
    )
    st.download_button(
        label="⬇️ Stáhnout Excel knihu jízd",
        data=excel_buffer,
        file_name=f"kniha_jizd_{mesic}_{rok}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
else:
    st.info("Přidej alespoň jednu jízdu pro export.")

# --- Sidebar: správa destinací
with st.sidebar:
    st.markdown("### 📍 Správa destinací")
    st.caption("Přidej / uprav vlastní destinace a vzdálenosti od Jankovcové.")

    nove_misto = st.text_input("Nová destinace")
    nove_km = st.number_input("Vzdálenost od Jankovcové (km)", min_value=0, value=0)
    if st.button("Přidat destinaci"):
        if nove_misto:
            st.session_state.destinace[nove_misto] = nove_km
            st.success(f"Přidáno: {nove_misto} ({nove_km} km)")
            st.rerun()

    st.markdown("---")
    st.markdown("**Aktuální destinace:**")
    for misto, km in sorted(st.session_state.destinace.items(), key=lambda x: x[1]):
        st.text(f"{misto}: {km} km")
