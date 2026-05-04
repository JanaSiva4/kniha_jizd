# 📕 Kniha jízd – generátor

Streamlit aplikace pro automatizaci tvorby měsíční knihy jízd.

## Co aplikace umí

- Načte výpis tankování z Excelu (export z Teams)
- Spočítá úseky mezi tankováními
- Pomáhá rozplánovat jízdy do úseků (s validací KM)
- Ukládá tvoje obvyklé destinace a vzdálenosti
- Šablony pro typické cesty (SKLC3, Brno, ČB...)
- Export do Excel knihy jízd

---

## 🚀 Nasazení – krok za krokem

### Krok 1: Vytvoř GitHub repo

1. Jdi na https://github.com a přihlas se
2. Vpravo nahoře klikni na **+** → **New repository**
3. Vyplň:
   - **Repository name**: `kniha-jizd`
   - **Visibility**: Private (ať to nikdo neuvidí)
   - Zaškrtni **Add a README file**
4. Klikni **Create repository**

### Krok 2: Nahraj soubory do repa

V GitHub repu klikni na **Add file** → **Upload files** a nahraj:

- `app.py` (hlavní aplikace)
- `requirements.txt` (seznam knihoven)
- `README.md` (tento soubor)

Klikni **Commit changes**.

### Krok 3: Deploy na Streamlit Cloud

1. Jdi na https://share.streamlit.io
2. Klikni **Sign in** a přihlas se přes GitHub (povol Streamlitu přístup k repu)
3. Klikni **New app** (vpravo nahoře)
4. Vyplň:
   - **Repository**: `tvoje-jmeno/kniha-jizd`
   - **Branch**: `main`
   - **Main file path**: `app.py`
5. Klikni **Deploy!**

Za 1-2 minuty ti app poběží na adrese `https://tvoje-jmeno-kniha-jizd-app-xxxx.streamlit.app`

### Krok 4: Použití každý měsíc

1. Otevři link na app
2. **Krok 1**: Vyber měsíc + zadej počáteční stav km z minulého měsíce
3. **Krok 2**: Stáhni výpis tankování z Teams (jako Excel) a nahraj
4. **Krok 3**: Naklikej jízdy mezi tankováními (sleduj zelené/červené úseky)
5. **Krok 4**: Stáhni hotový Excel ✓

---

## ⚙️ Úprava destinací

Vlevo v sidebaru můžeš:
- Přidávat nové destinace s jejich vzdáleností od Jankovcové
- Vidět seznam aktuálních destinací

Pro trvalou úpravu uprav `DEFAULT_DESTINACE` na začátku `app.py` v GitHubu.

## ⚙️ Úprava křížových vzdáleností

Pokud potřebuješ specifické vzdálenosti mezi destinacemi (které nejsou z/do Jankovcové), uprav `KRIZOVE_VZDALENOSTI` v `app.py`.

## 🐛 Něco nefunguje?

- **Excel se neparsuje**: appka hledá sloupce obsahující "datum", "stav km", "množství", "čerpací". Pokud má tvůj Excel jiné názvy, uprav funkci `parsuj_tankovani_excel`.
- **App spadla po deployi**: Streamlit Cloud má logy – klikni v dashboardu na app → **Manage app** → **Logs**.

## 💡 Tipy

- Šablony (Krok 3, sekce „Šablony") přidají rovnou 2-3 jízdy pro typické scénáře (SKLC3 s přespáním atd.). Pak si jen upravíš datum.
- Validace úseků v reálném čase: zelená = OK, červená = chybí km, oranžová = přebývá.
