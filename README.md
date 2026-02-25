# YAG Acquisitie Tool

Python CLI voor geautomatiseerde B2B acquisitie. Haalt leads op via Lusha, filtert op DNC, enrichet met contactgegevens, genereert gepersonaliseerde HTML-mails via OpenAI en verstuurt via Gmail SMTP.

**Google Sheets is de centrale database** — alle leads, statussen en send-logs worden daar bijgehouden. Geen lokale CSV-bestanden.

---

## Projectstructuur

```
yag-mailer/
├── main.py                        ← Enige bestand dat je start
├── requirements.txt
├── .gitignore
│
├── consultants/
│   ├── .env.example               ← Template voor nieuwe consultants (wél in git)
│   └── rick.env                   ← Ricks profiel met API keys etc. (NIET in git)
│
├── src/
│   ├── config.py                  ← Kolomnummers + statusconstanten (single source of truth)
│   ├── sheets.py                  ← Alle Google Sheets lees/schrijf operaties
│   ├── lusha.py                   ← Lusha search + enrich + industries ophalen
│   ├── ai_gen.py                  ← OpenAI e-mail generatie
│   ├── storage.py                 ← DNC lijst (lokaal Excel-bestand)
│   └── gmail_send.py              ← Gmail SMTP verzending via App Password
│
├── credentials/
│   └── service_account.json       ← Google Sheets toegang (NIET in git)
│
├── data/
│   ├── Niet Benaderen.xlsx        ← DNC lijst — kolom "Bedrijf" vereist
│   └── YAG Voorstelslides.pdf     ← Optionele PDF bijlage (ATTACHMENT_PDF in .env)
│
└── output/
    └── lusha_page_state.json      ← Lusha paginateller per preset (automatisch)
```

---

## Google Sheets structuur

Het script schrijft de header automatisch als de sheet leeg is. De sheet heeft **26 kolommen**:

```
A  Company              J  Mail Status           S  Gevallen
B  First Name           K  Datum Mail            T  Hoe contact
C  Last Name            L  Follow-up datum       U  ─── separator ───
D  Job Title            M  Reactie ontvangen     V  Request ID
E  Email                N  Opmerking             W  Contact ID
F  Phone                O  ─── separator ───     X  isShown
G  LinkedIn URL         P  Consultant            Y  AI Bericht
H  Enriched ✅          Q  Vestiging             Z  AI Tokens
I  AI Status            R  Type
```

Naast het hoofdtabblad worden twee extra tabbladen automatisch aangemaakt:

| Tabblad | Inhoud |
|---------|--------|
| **Send Log** | Tijdstempel, consultant, bedrijf, e-mail, onderwerp, status per verzending |
| **DNC Archief** | Rijen die via stap 6 (opschonen) zijn verplaatst |

---

## Eenmalige setup

### 1. Python packages

```powershell
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
7. Open het bestand, kopieer het `client_email` adres
8. Ga naar je Google Sheet → **Delen** → plak het adres → Editor → Verzenden

### 3. Gmail — App Password

Per consultant eenmalig doen:

1. Ga naar [myaccount.google.com/security](https://myaccount.google.com/security) → zet **2-stapsverificatie** aan
2. Ga naar [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
3. App name: `yag-mailer` → **Create** → kopieer de 16 tekens
4. Zet dit in `consultants/<naam>.env` als `GMAIL_APP_PASSWORD`

### 4. Consultant profiel aanmaken

```powershell
python main.py
# Kies [n] → Nieuw profiel aanmaken → volg de vragen
```

Of handmatig:
```powershell
copy consultants\.env.example consultants\anna.env
# Open het bestand en vul in
```

---

## Gebruik

```powershell
python main.py
```

Het script vraagt bij elke start wie je bent en laadt automatisch het juiste profiel:

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
  Wat wil je doen?
──────────────────────────────────────────────────
  [1] 📥  Leads ophalen via Lusha
  [2] 🔍  Leads enrichen (email/tel/LinkedIn)
  [3] 🤖  AI berichten genereren
  [4] ✉   Mails versturen
  [5] 📊  Overzicht bekijken
  [6] 🧹  Sheet opschonen
  [q] 🚪  Afsluiten
```

---

## Flow per batch: 1 → 2 → 3 → 4

### [1] Leads ophalen via Lusha

- Kies een ICP preset of eigen filters:
  - `nl_midsized_csuite` — NL, 51–1000 medewerkers, C-suite
  - `nl_large_csuite` — NL, 1001–10.000 medewerkers, C-suite + Directors
- Kies of wijzig de industrie — de volledige Lusha industrielijst wordt live opgehaald
- **Paginateller**: het script onthoudt welke Lusha-pagina je de vorige keer had. Je ziet dit bij de start van stap 1:
  ```
  📄 Lusha pagina: 4  (Enter = overnemen, of typ ander getal):
  ```
  Druk Enter om door te gaan op pagina 4, of typ een ander getal om te springen.
- **Meta-velden** (Consultant, Vestiging, Type, Hoe contact) worden vooringevuld vanuit je `.env` — druk Enter om te bevestigen, of wijzig ze
- **Duplicaat check** — contacten die al in de sheet staan worden overgeslagen
- **DNC scan direct na het ophalen** — leads van geblokkeerde bedrijven worden meteen gemarkeerd als 🚫 DNC

### [2] Leads enrichen

- Haalt email, telefoon en LinkedIn URL op voor alle leads met `Enriched = No`
- Slaat 🚫 DNC en ⏭ AL GEMAILD rijen automatisch over
- Groepeert op Request ID (vereiste van Lusha API)
- Resultaten worden direct teruggeschreven naar de sheet

### [3] AI berichten genereren

- Genereert een gepersonaliseerde e-mail per lead via OpenAI (`gpt-4.1-mini`)
- Structuur:
  - **Vaste opening** — wie je bent en waarom je zelf acquireert
  - **AI connectiezinnen** (2–3 regels) — OpenAI zoekt via web search op wat het bedrijf doet en schrijft waarom jij specifiek bij hen uitkwam
  - **Vaste pitch** — recente YAG projecten
  - **Vaste signature** — naam, telefoon, LinkedIn, vestiging, YAG logo
- URLs die het model toch toevoegt worden automatisch verwijderd uit de connectiezinnen
- **Dry-run optie** — genereert een preview zonder echte OpenAI API calls; status wordt 🔴 DRY RUN (kan later opnieuw worden aangeboden voor echte generatie)
- Verplichte velden voor AI: Consultant, Vestiging, Type, Hoe contact — ontbreken ze, dan wordt de rij overgeslagen
- Meta-velden worden meegeschreven bij elke AI-generatie, zodat eventueel lege kolommen automatisch worden aangevuld vanuit je profiel

### [4] Mails versturen

- Toont een preview van de eerste mail vóór verzending
- Vraagt bevestiging + aantal te versturen mails
- DRY_RUN is per sessie te togglen vanuit het menu
- **Veiligheidslagen:**
  - 🚫 DNC — tweede controle (voor het geval de lijst is bijgewerkt na de search)
  - ⏭ AL GEMAILD — suppressie check op e-mailadres (uit sheet, alleen ✅ SENT telt)
  - Bedrijf al benaderd — check op bedrijfsnaam van eerder verstuurde mails
- Verstuurt als **HTML-mail** met:
  - YAG logo ingesloten (geen externe afbeeldingslink nodig)
  - Klikbare LinkedIn en website links in de signature
  - Plain-text fallback voor e-mailclients zonder HTML-ondersteuning
  - **Optionele PDF bijlage** (stel in via `ATTACHMENT_PDF` in je `.env`)
- Wacht `RATE_LIMIT_SEC` seconden tussen mails
- Elke verzending wordt gelogd in het **Send Log** tabblad van de sheet

### [5] Overzicht

Toont tellingen per status (Enriched, AI Status, Mail Status), tokenverbruik met geschatte kosten, en de huidige Lusha paginateller per preset (aanpasbaar vanuit dit menu).

### [6] Sheet opschonen

- **DNC rijen** worden verplaatst naar het 'DNC Archief' tabblad
- **Rijen zonder email** (na enrichment) worden verwijderd
- **🔴 DRY RUN rijen** blijven staan — die worden automatisch opnieuw aangeboden bij stap 4
- Verwijdering gebeurt in één batch-API-call om quota-fouten te voorkomen

---

## Meerdere consultants

Elke consultant heeft een eigen `.env` bestand in `consultants/`. Het script detecteert alle bestanden automatisch en toont ze als keuzemenu.

| Variabele | Waarom per consultant anders |
|-----------|------------------------------|
| `SENDER_NAME` | Naam in de mail en signature |
| `SENDER_EMAIL` | Gmail account waarmee verstuurd wordt |
| `SENDER_PHONE` | Telefoonnummer in de signature |
| `SENDER_LINKEDIN` | LinkedIn profiellink in de signature |
| `GMAIL_APP_PASSWORD` | Eigen Gmail App Password |
| `VESTIGING_DEFAULT` | Vooringevuld bij leads ophalen |
| `ATTACHMENT_PDF` | Pad naar PDF bijlage (optioneel) |

De volgende variabelen zijn gedeeld en staan bij iedereen hetzelfde: `SPREADSHEET_ID`, `LUSHA_API_KEY`, `OPENAI_API_KEY`.

---

## Alle `.env` variabelen

| Variabele | Standaard | Omschrijving |
|-----------|-----------|--------------|
| `SPREADSHEET_ID` | *(verplicht)* | ID uit de Google Sheets URL |
| `WORKSHEET_NAME` | `Sheet1` | Naam van het hoofdtabblad |
| `SERVICE_ACCOUNT_JSON` | `credentials/service_account.json` | Pad naar service account |
| `SENDER_NAME` | *(verplicht)* | Naam in Van-header en signature |
| `SENDER_EMAIL` | *(verplicht)* | Gmail adres |
| `SENDER_PHONE` | | Telefoon in signature |
| `SENDER_LINKEDIN` | | LinkedIn URL in signature |
| `GMAIL_APP_PASSWORD` | *(verplicht bij LIVE)* | 16-tekens Gmail App Password |
| `LUSHA_API_KEY` | *(verplicht)* | Lusha API key |
| `OPENAI_API_KEY` | *(verplicht)* | OpenAI API key |
| `STUDIE` | `Technische Bedrijfskunde` | In de opening van de mail |
| `UNIVERSITEIT` | `TU Eindhoven` | In de opening van de mail |
| `SUBJECT_TEMPLATE` | `Young Advisory Group x {company}` | Onderwerpregel |
| `USE_WEB_SEARCH` | `true` | OpenAI web search voor bedrijfsinfo |
| `DRY_RUN` | `true` | Geen echte mails versturen |
| `MAX_EMAILS` | `20` | Max mails per sessie |
| `RATE_LIMIT_SEC` | `2` | Seconden wachten tussen mails |
| `DNC_PATH` | `data/Niet Benaderen.xlsx` | Pad naar DNC lijst |
| `ATTACHMENT_PDF` | *(leeg = geen bijlage)* | Pad naar PDF bijlage |
| `VESTIGING_DEFAULT` | `Eindhoven-Tilburg` | Vooringevulde vestiging |
| `TYPE_DEFAULT` | `Cold` | Vooringevuld contacttype |
| `GEVALLEN_DEFAULT` | *(leeg)* | Vooringevulde sector/gevallen |
| `HOE_CONTACT_DEFAULT` | `Lusha` | Hoe contact verkregen |
| `INDUSTRY_IDS_DEFAULT` | *(leeg = alle)* | Kommagescheiden Lusha industrie-IDs |

---

## Veiligheidslagen

| Laag | Moment | Bron | Actie |
|------|--------|------|-------|
| DNC bedrijf | Direct na search | `data/Niet Benaderen.xlsx` | Gemarkeerd als 🚫 DNC |
| DNC bedrijf | Voor verzending | `data/Niet Benaderen.xlsx` | Gemarkeerd als 🚫 DNC |
| Al gemaild (email) | Voor verzending | Sheet — alleen ✅ SENT | Overgeslagen |
| Bedrijf al benaderd | Voor verzending | Sheet — alleen ✅ SENT | Overgeslagen |

De DNC check gebruikt **fuzzy matching**:
- Rechtsvorm-suffixen worden genegeerd (B.V., N.V., Ltd, GmbH, enz.)
- Samengestelde namen worden gesplitst op `;`, `|`, ` - `, `,`
- Substrings van ≥8 tekens worden herkend in beide richtingen

---

## Status-overzicht

### AI Status (kolom I)

| Status | Betekenis |
|--------|-----------|
| `PENDING` | Wacht op AI generatie |
| `RUNNING` | Bezig (crasht het script, dan blijft dit staan) |
| `✅ DONE` | Bericht gegenereerd, klaar voor verzending |
| `🔴 DRY RUN` | Preview gegenereerd, opnieuw aanbieden voor echte generatie |
| `❌ ERROR` | Fout opgetreden — zie kolom N (Opmerking) |

### Mail Status (kolom J)

| Status | Betekenis |
|--------|-----------|
| `PENDING` | Nog niet verstuurd |
| `✅ SENT` | Echte mail verstuurd |
| `🔴 DRY RUN` | Test-verzending — wordt opnieuw aangeboden bij stap 4 |
| `🚫 DNC` | Geblokkeerd door DNC lijst |
| `⏭ AL GEMAILD` | E-mailadres al eerder benaderd |
| `❌ ERROR` | SMTP fout — zie kolom N (Opmerking) |

---

## Veelgestelde vragen

**"Spreadsheet niet gevonden"**
Controleer `SPREADSHEET_ID` in je `.env` — het ID staat in de Sheet URL tussen `/d/` en `/edit`. Zorg dat de sheet gedeeld is met het `client_email` uit `service_account.json`.

**Gmail App Password werkt niet**
- 2-stapsverificatie moet aan staan op je Google account
- Gebruik het App Password (16 tekens), niet je gewone wachtwoord
- Spaties in het wachtwoord zijn toegestaan, het script verwijdert ze automatisch

**AI generatie mislukt**
- Controleer `OPENAI_API_KEY` in je `.env`
- Zorg dat er saldo op je OpenAI account staat
- De foutmelding staat ook in kolom N (Opmerking) in de sheet

**Leads die al in de sheet staan worden toch opgehaald**
Het duplicaat-check werkt op Contact ID (Lusha uniek ID). Als je handmatig rijen hebt verwijderd, worden die contacten bij een nieuwe search opnieuw toegevoegd.

**Sheet opschonen gooit een quota-fout**
Dit zou niet meer moeten voorkomen — cleanup verwijdert alle rijen in één batch-API-call.

**PDF bijlage wordt niet meegestuurd**
Controleer of `ATTACHMENT_PDF` in je `.env` naar een bestaand bestand wijst. Het script geeft een waarschuwing in de terminal als het bestand niet gevonden wordt, maar stuurt de mail wel gewoon zonder bijlage.

**AI bericht klopt maar meta-kolommen (Consultant, Vestiging etc.) zijn leeg**
Dit kon voorkomen in een oudere versie. Nu worden meta-velden altijd meegeschreven bij AI-generatie. Draai stap [3] opnieuw op de betreffende rijen — ze worden automatisch aangevuld vanuit je profiel.
