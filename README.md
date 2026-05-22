# Dochádzková aplikácia

## Lokálne spustenie (testovanie)

```bash
pip install -r requirements.txt
python app.py
# Otvorte: http://localhost:5000
```

## Deploy na Render.com (krok po kroku)

### 1. GitHub
1. Vytvorte účet na https://github.com (zadarmo)
2. Kliknite **New repository** → názov: `dochadzka` → Create
3. Nahrajte všetky súbory (drag & drop cez web alebo git push)

### 2. Render.com
1. Vytvorte účet na https://render.com (zadarmo, Google login)
2. **New** → **Web Service**
3. Prepojte GitHub → vyberte repozitár `dochadzka`
4. Nastavenia:
   - **Name:** dochadzka
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
5. **Environment Variables** → Add:
   - `SECRET_KEY` = vymyslite dlhý náhodný reťazec (napr. `abc123xyz456...`)
6. Kliknite **Create Web Service**

### 3. Databáza (PostgreSQL) na Render.com
1. **New** → **PostgreSQL** → Create
2. Po vytvorení skopírujte **Internal Database URL**
3. Pridajte do Environment Variables: `DATABASE_URL` = (skopírovaná URL)

### 4. Hotovo!
Dostanete URL napr.: `https://dochadzka-xxxx.onrender.com`

## Prvé použitie
- Prvý zaregistrovaný používateľ = automaticky **správca**
- Správca vidí všetkých zamestnancov v sekcii **Správa**

## Funkcie
- ⏱ Časovač s klientmi
- ✂ Rozdelenie času na činnosti (bez prerušenia)
- 📊 Výkaz s filtrovaním
- ⬇ Export do Excelu
- 🛡 Správca vidí všetkých
