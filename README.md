# YAG Acquisitie Tool

Volledig Python CLI voor het ophalen, enrichen, AI-personaliseren en versturen van acquisitie e-mails.
Alle data leeft in Google Sheets — één sheet als central dashboard.

---

## Projectstructuur

```
yag-mailer/
├── main.py                  ← Start hier (CLI menu)
├── .env                     ← Jouw configuratie (niet committen!)
├── .env.example             ← Template voor .env
├── requirements.txt
│
├── src/
│   ├── config.py            ← Kolom mapping + status constanten
│   ├── sheets.py            ← Google Sheets lezen/schrijven
│   ├── lusha.py             ← Lusha API (search + enrich)
│   ├── ai_gen.py            ← OpenAI bericht generatie
│   ├── storage.py           ← DNC, suppressie, send log
│   ├── gmail_auth.py        ← Gmail OAuth authenticatie
│   └── gmail_send.py        ← Gmail API verzending
│
├── credentials/             ← NIET in git! (zie .gitignore)
│   ├── service_account.json ← Voor Sheets toegang
│   ├── credentials.json     ← Voor Gmail OAuth
│   └── token.json           ← Automatisch aangemaakt
│
├── data/
│   └── Niet Benaderen.xlsx  ← DNC lijst
│
└── output/                  ← Automatisch aangemaakt
    ├── suppression.csv      ← Al verstuurde e-mails
    └── send_log.csv         ← Volledige audit trail
```

---

## Setup (eenmalig)

### 1. Python omgeving

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. .env aanmaken

```bash
cp .env.example .env
# Vul .env in met jouw gegevens
```

### 3. Google Sheets — Service Account

1. Ga naar [Google Cloud Console](https://console.cloud.google.com)
2. Maak een project aan (of gebruik bestaand)
3. Enable **Google Sheets API**
4. **IAM & Admin → Service Accounts → Nieuw**
5. Download de JSON key → sla op als `credentials/service_account.json`
6. **Deel je Google Sheet** met het service account e-mailadres (Editor rechten)
7. Kopieer de Spreadsheet ID uit de URL → zet in `.env` als `SPREADSHEET_ID`

### 4. Gmail — OAuth Credentials

1. Zelfde Google Cloud project
2. Enable **Gmail API**
3. **APIs & Services → Credentials → OAuth 2.0 Client ID**
4. Type: Desktop App
5. Download JSON → sla op als `credentials/credentials.json`
6. Eerste keer `python main.py` → browser opent voor toestemming

### 5. Sheet kolom volgorde

Zorg dat je sheet **exact** deze kolomvolgorde heeft (of laat `ensure_header` het aanmaken):

```
A: Company          I: AI Status        Q: Vestiging
B: First Name       J: Mail Status      R: Type
C: Last Name        K: Datum Mail       S: Gevallen
D: Job Title        L: Follow-up datum  T: Hoe contact
E: Email            M: Reactie          U: --- separator ---
F: Phone            N: Opmerking        V: Request ID
G: LinkedIn URL     O: --- separator -- W: Contact ID
H: Enriched ✅      P: Consultant       X: isShown
                                        Y: AI Bericht
```

---

## Gebruik

```bash
python main.py
```

```
══════════════════════════════════════════════════
  YAG Acquisitie Tool
══════════════════════════════════════════════════

  [1] 📥  Leads ophalen via Lusha
  [2] 🔍  Leads enrichen (email/tel/LinkedIn)
  [3] 🤖  AI berichten genereren
  [4] ✉   Mails versturen
  [5] 📊  Overzicht bekijken
  [q] 🚪  Afsluiten
```

### Normale flow

1. **[1]** Leads ophalen → vul ICP filters in, leads worden in de sheet gezet
2. **[2]** Leads enrichen → emails, telefoon en LinkedIn worden toegevoegd
3. **[3]** AI genereren → gepersonaliseerde mail per lead, teruggeschreven naar sheet
4. **[4]** Mails versturen → DRY_RUN=true eerst, dan echt

### DRY_RUN

Zet `DRY_RUN=true` in `.env` om alles te testen zonder echte mails te sturen.
Status in sheet wordt dan `DRY RUN` in plaats van `✅ SENT`.

---

## Veiligheidslagen

Bij verzending worden drie checks gedaan:

| Check | Bron | Actie bij match |
|-------|------|----------------|
| DNC | `data/Niet Benaderen.xlsx` | Sheet: 🚫 DNC, skip |
| Suppressie | `output/suppression.csv` | Sheet: ⏭ AL GEMAILD, skip |
| Bedrijf al benaderd | `output/send_log.csv` | skip (collega bescherming) |

---

## Meerdere consultants

Elke consultant heeft zijn eigen `.env`:

```bash
cp .env .env.rick
cp .env .env.anna
# Pas SENDER_NAME, SENDER_EMAIL, TOKEN_JSON aan per persoon
```

Starten met specifieke config:
```bash
env $(cat .env.rick) python main.py
```a
