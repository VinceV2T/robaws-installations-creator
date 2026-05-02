# Robaws installations creator

Automatisch installatie-records aanmaken in Robaws op basis van nieuwe
bestelbonnen (purchase-supply-orders), zodat verkochte artikelen meteen
zichtbaar zijn doorheen het volledige proces — zonder manueel overtikken.

## De flow in jullie proces

```
Offerte ──▶ Order ──▶ Bestelbon (per leverancier)
                            │
                            ▼
                   [creator.py - dit script]
                            │
                            ▼
                Installaties tab in Robaws
                            │
                            ▼
                Onderhoudsopdrachten (later)
```

Door de installatie te koppelen aan zowel het project (via `assignedProjectId`)
als aan de bestelbon-lijn (via `materialId` op die lijn), is het artikel
zichtbaar in de offerte, het order, de bestelbon én het installatie-tabblad.
Vanuit dat installatie-tabblad kunnen later onderhoudsopdrachten aangemaakt
worden voor het toestel.

## Hoe het werkt

Dagelijks (4× verspreid via cron, zie [`.github/workflows/create-installations.yml`](.github/workflows/create-installations.yml)):

1. De masterlijst van installatie-art.nummers wordt ingelezen vanuit
   `data/*.xlsx` (kolom `Art.nummer`).
2. Robaws wordt gequeried voor bestelbonnen die in de laatste `DAYS_BACK`
   dagen zijn aangemaakt of gewijzigd.
3. Voor elke lijn op een bestelbon waar het artikel matched met de masterlijst
   **én waar de materieel-kolom (materialId) nog leeg is**:
   - de gekoppelde sales order wordt opgehaald (voor adres, klant, eindklant,
     project)
   - er wordt 1 installatie aangemaakt per stuk (op basis van `line.quantity`)
     via `POST /api/v2/installations`. Serienummer blijft bewust leeg — wordt
     later ingevuld bij levering.
   - de bestelbon-lijn wordt gepatcht met `materialId` = id van de eerste
     aangemaakte installatie (`PATCH /api/v2/purchase-supply-orders/{id}/line-items/{id}`)
4. Een JSON-rapport wordt bewaard als `create_report.json` (artifact in de
   Action). Bij echte problemen wordt een mail verstuurd.

### Wat zijn "echte problemen" voor de mail?

Enkel deze gevallen genereren een mail:

- bestelbon-lijn zonder gekoppelde sales order
- gekoppelde order zonder leveradres
- gekoppelde order zonder project
- POST naar `/api/v2/installations` faalt
- PATCH van `materialId` op de bestelbon-lijn faalt

Lege serienummers, indienststelling-data en werknemers worden **niet**
gerapporteerd — die zijn normaal in deze fase van het proces.

## Idempotency

Bij elke run worden bestelbon-lijnen waar `materialId` al gezet is overgeslagen.
Een lijn die al verwerkt is door een vorige run wordt nooit opnieuw aangeraakt,
zelfs als de cron 4× per dag draait.

## Repo structuur

```
.
├── .github/workflows/create-installations.yml   # cron + manuele trigger
├── data/                                         # artikellijsten (.xlsx)
│   └── Importlijst CP_MultiAir_2026_*.xlsx
├── creator.py                                    # hoofdscript
├── requirements.txt
├── .env.example
└── README.md
```

Voeg gerust nieuwe Excel-lijsten toe in `data/` — alle .xlsx in die map worden
samen ingelezen tot één masterset van art.nummers.

## Setup

### 1. Repo aanmaken op GitHub

```bash
# Lokaal:
cd robaws-installations-creator
git init
git add .
git commit -m "initial: installations creator"
git branch -M main
git remote add origin git@github.com:VinceV2T/robaws-installations-creator.git
git push -u origin main
```

Maak het repo aan onder `VinceV2T` als **Private**.

### 2. Secrets configureren

Settings → Secrets and variables → Actions → **New repository secret**:

| Secret             | Inhoud                                       |
|--------------------|----------------------------------------------|
| `ROBAWS_API_KEY`    | je Robaws API-key                            |
| `ROBAWS_API_SECRET` | je Robaws API-secret                         |
| `SMTP_HOST`         | bv. `smtp.office365.com` (optioneel)         |
| `SMTP_PORT`         | `587` (optioneel)                            |
| `SMTP_USER`         | mail-account voor verzenden (optioneel)      |
| `SMTP_PASSWORD`     | wachtwoord/app-password (optioneel)          |
| `MAIL_FROM`         | bv. `automation@v2technics.be` (optioneel)   |
| `MAIL_TO`           | `vincent@v2technics.be`                      |

Als je SMTP-secrets leeg laat, schrijft het script enkel het JSON-rapport en
verstuurt het geen mail.

### 3. Eerste keer dry-run draaien

GitHub → Actions → **Nightly installations creator** → **Run workflow**.

In het formulier:
- `dry_run: true`
- `days_back: 14`
- `test_bestelbon: ` (leeg, of een specifieke logicId zoals `B260412`)

Run, wacht tot klaar, klik op de run en download de `create-report` artifact.
Open `create_report.json` en bekijk wat het script **zou** doen.

### 4. Live zetten

Eens je tevreden bent met de dry-run, draai opnieuw met `dry_run: false`.
Daarna draait het script automatisch op het cron-schema.

## Lokaal testen

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# vul ROBAWS_API_KEY / SECRET in
export $(grep -v '^#' .env | xargs)
python creator.py
```

## Configuratie via env

| Variabele          | Default                  | Beschrijving                                        |
|--------------------|--------------------------|-----------------------------------------------------|
| `ROBAWS_BASE_URL`  | `https://app.robaws.com` | Robaws API host                                     |
| `ROBAWS_API_KEY`   | (verplicht)              | API-key                                             |
| `ROBAWS_API_SECRET`| (verplicht)              | API-secret                                          |
| `DRY_RUN`          | `true`                   | `false` voor effectief aanmaken                     |
| `DAYS_BACK`        | `14`                     | Hoe ver terug bestelbonnen scannen                  |
| `TEST_BESTELBON`   | `(leeg)`                 | Optioneel: enkel deze logicId verwerken             |
| `SMTP_*`, `MAIL_*` | `(leeg)`                 | Mail-config; leeg = geen mail                       |

## Verwante repos

- [`robaws-invoice-linker`](https://github.com/VinceV2T/robaws-invoice-linker)
  — koppelt aankoopfacturen aan sales orders. Zelfde authenticatiepatroon.
