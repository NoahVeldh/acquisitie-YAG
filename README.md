# YAG Acquisitie Tool

Python CLI voor geautomatiseerde B2B acquisitie. Haalt leads op via Lusha, filtert op DNC, enrichet met contactgegevens, genereert gepersonaliseerde mails via OpenAI en verstuurt via Gmail.

Google Sheets is het centrale dashboard — alle statussen zijn live zichtbaar tijdens het draaien.

---

## Projectstructuur

```
Python/
├── main.py                      ← Enige bestand dat je start
├── requirements.txt
├── .gitignore
│
├── consultants/
│   ├── .env.example             ← Template voor nieuwe consultants (wél in git)
│   └── rick.env                 ← Ricks profiel met API keys etc. (NIET in git)
│
├── src/
│   ├── config.py                ← Kolomnummers + statusconstanten (single source of truth)
│   ├── sheets.py                ← Alle Google Sheets lees/schrijf operaties
│   ├── lusha.py                 ← Lusha search + enrich + industries ophalen
│   ├── ai_gen.py                ← OpenAI e-mail generatie
│   ├── storage.py               ← DNC lijst, suppressie, send log
│   └── gmail_send.py            ← Gmail SMTP verzending via App Password
│
├── credentials/
│   └── service_account.json     ← Google Sheets toegang (NIET in git)
│
├── data/
│   └── Niet Benaderen.xlsx      ← DNC lijst — kolom "Bedrijf" vereist
│
└── output/                      ← Automatisch aangemaakt
    ├── suppression.csv          ← Alle al verstuurde e-mailadressen
    └── send_log.csv             ← Volledige audit trail per verzending
```

---

## Eenmalige setup

### 1. Python packages

```powershell
cd Python
pip install -r requirements.txt
```

### 2. Google Sheets — Service Account

1. Ga naar [console.cloud.google.com](https://console.cloud.google.com)
2. Maak een project aan
3. **APIs & Services → Library → Google Sheets API → Enable**
4. **IAM & Admin → Service Accounts → + Create Service Account**
   - Naam: `yag-mailer` → Create → Continue → Done
5. Klik op het service account → **Keys → Add Key → JSON**
6. Download → hernoem naar `service_account.json` → zet in `credentials/`
7. Open het bestand in Notepad, kopieer het `client_email` adres
8. Ga naar je Google Sheet → **Delen** → plak het adres → Editor → Verzenden

### 3. Gmail — App Password

Per consultant eenmalig doen:

1. Ga naar [myaccount.google.com/security](https://myaccount.google.com/security) → zet **2-stapsverificatie** aan
2. Ga naar [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
3. App name: `yag-mailer` → **Create** → kopieer de 16 tekens
4. Zet in `consultants/<naam>.env` als `GMAIL_APP_PASSWORD`

### 4. Consultant profiel

```powershell
python main.py
# Kies [n] → Nieuw profiel aanmaken
```

Of handmatig:
```powershell
copy consultants\.env.example consultants\anna.env
# Open en vul in
```

### 5. Sheet kolom volgorde

De sheet moet exact deze 25 kolommen hebben in deze volgorde. Het script schrijft de header automatisch als de sheet leeg is:

```
A  Company          J  Mail Status      S  Gevallen
B  First Name       K  Datum Mail       T  Hoe contact
C  Last Name        L  Follow-up datum  U  ─── separator ───
D  Job Title        M  Reactie          V  Request ID
E  Email            N  Opmerking        W  Contact ID
F  Phone            O  ─── separator ── X  isShown
G  LinkedIn URL     P  Consultant       Y  AI Bericht
H  Enriched ✅      Q  Vestiging
I  AI Status        R  Type
```

---

## Gebruik

```powershell
cd Python
python main.py
```

Het script vraagt bij elke start wie je bent:

```
══════════════════════════════════════════════════
  YAG Acquisitie Tool
══════════════════════════════════════════════════

  Wie ben je?

    [1] Rick op het Veld  (Eindhoven-Tilburg)
    [n] Nieuw profiel aanmaken

  > 1

  ✅ Ingelogd als: Rick op het Veld
  🟡 DRY RUN  |  Max: 20 mails  |  Sheet: ...abc123

──────────────────────────────────────────────────
  [1] 📥  Leads ophalen via Lusha
  [2] 🔍  Leads enrichen (email/tel/LinkedIn)
  [3] 🤖  AI berichten genereren
  [4] ✉   Mails versturen
  [5] 📊  Overzicht bekijken
  [q] 🚪  Afsluiten
```

---

## Flow

De normale volgorde per batch is **1 → 2 → 3 → 4**:

### [1] Leads ophalen via Lusha

- Kies een ICP preset (`nl_midsized_csuite` of `nl_large_csuite`) of eigen filters
- Kies of wijzig de industrie — de volledige Lusha industrielijst wordt live opgehaald
- Kiest automatisch een willekeurige startpagina zodat je nooit dezelfde leads herhaalt
- Meta-velden (vestiging, type, hoe contact) zijn vooringevuld — gewoon Enter
- **Duplicaat check**: contacten die al in de sheet staan worden overgeslagen
- **DNC scan direct daarna**: leads van bedrijven op de Niet Benaderen lijst worden meteen gemarkeerd als 🚫 DNC en overgeslagen in alle volgende stappen

### [2] Leads enrichen

- Haalt email, telefoon en LinkedIn op voor alle leads met status `Enriched = No`
- Slaat 🚫 DNC en ⏭ AL GEMAILD rijen automatisch over
- Groepeert op Request ID (Lusha vereiste)
- Schrijft resultaten direct terug naar de sheet

### [3] AI berichten genereren

- Genereert een gepersonaliseerde e-mail per lead via OpenAI (`gpt-4.1-mini`)
- Structuur: vaste opening → AI connectiezinnen (2-3 regels over waarom jij dit bedrijf benadert) → vaste pitch → vaste signature
- Slaat leads zonder email, zonder verplichte meta-velden, en 🚫 DNC rijen over
- Bericht wordt teruggeschreven naar kolom Y (AI Bericht) in de sheet
- Fouten worden gelogd in kolom N (Opmerking)

### [4] Mails versturen

- Toont een preview van de eerste mail voor verzending
- Vraagt bevestiging met aantal te versturen mails
- DRY_RUN toggle beschikbaar vanuit het menu
- Veiligheidslagen bij verzending:
  - 🚫 DNC — tweede controle voor het geval de lijst is bijgewerkt na de search
  - ⏭ AL GEMAILD — suppression check op e-mailadres
  - Bedrijf al benaderd door een collega — check op send_log
- Verstuurt via Gmail SMTP, wacht `RATE_LIMIT_SEC` seconden tussen mails
- Logt elke verzending in `output/send_log.csv`

### [5] Overzicht

Toont tellingen per status (Enriched, AI Status, Mail Status) en huidige config.

---

## Meerdere consultants

Elke consultant heeft een eigen bestand in `consultants/`. Bij opstarten kies je wie je bent — het script laadt automatisch het juiste profiel.

| Variabele | Waarom per consultant anders |
|-----------|------------------------------|
| `SENDER_NAME` | Naam in de mail en signature |
| `SENDER_EMAIL` | Gmail account waarmee verstuurd wordt |
| `SENDER_PHONE` | Telefoonnummer in de signature |
| `GMAIL_APP_PASSWORD` | Eigen Gmail App Password |
| `VESTIGING_DEFAULT` | Vooringevuld bij leads ophalen |

`SPREADSHEET_ID`, `LUSHA_API_KEY` en `OPENAI_API_KEY` zijn gedeeld — staan bij iedereen hetzelfde.

---

## Veiligheidslagen

| Laag | Wanneer | Bron | Actie |
|------|---------|------|-------|
| DNC | Na search én voor verzending | `data/Niet Benaderen.xlsx` | 🚫 DNC — overgeslagen |
| Suppressie | Voor verzending | `output/suppression.csv` | ⏭ AL GEMAILD — overgeslagen |
| Bedrijf al benaderd | Voor verzending | `output/send_log.csv` | Overgeslagen |

De DNC check gebruikt fuzzy matching: BV/NV/Ltd worden genegeerd, samengestelde namen worden gesplitst, substrings van ≥8 tekens worden herkend.

---

## DRY_RUN

Zolang `DRY_RUN=true` staat in je `.env` worden geen echte mails verstuurd. Status in de sheet wordt `DRY RUN` in plaats van `✅ SENT`. Je kunt dit per sessie omzetten via het menu in stap [4].

---

## Veelgestelde vragen

**Ik krijg "Spreadsheet niet gevonden"**
Controleer `SPREADSHEET_ID` in je `.env` — het ID staat in de Sheet URL tussen `/d/` en `/edit`. Zorg dat de sheet gedeeld is met het `client_email` uit `service_account.json`.

**Gmail App Password werkt niet**
- 2-stapsverificatie moet aan staan op je Google account
- Gebruik het App Password (16 tekens), niet je gewone wachtwoord
- Spaties mogen erbij, het script verwijdert ze automatisch

**AI generatie mislukt**
- Controleer `OPENAI_API_KEY` in je `.env`
- Zorg dat er saldo op je OpenAI account staat
- Foutmelding staat ook in kolom N (Opmerking) in de sheet

**Ik zie steeds dezelfde leads**
Het script kiest automatisch een willekeurige pagina bij elke run. Als je toch overlap hebt, selecteer dan een ander ICP preset of pas de industrie aan.

**De sheet heeft de verkeerde kolommen**
Maak een nieuwe lege sheet en run het script — `ensure_header()` schrijft automatisch de correcte header.
