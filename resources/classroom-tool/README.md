---
created: 2026-05-06
updated: 2026-05-06
created_by: claude-opus-4-7
updated_by: claude-opus-4-7
agent_version: 03.26
---

# classroom-tool

Personligt verktyg som hämtar kursdata från Google Classroom och producerar **anonymiserade** markdown-sammanställningar för en kurs (uppgifter, inlämningsstatus, antal saknas).

Riktiga elev-namn syns aldrig i normal output. Verktyget tilldelar varje elev en stabil pseudonym per kurs (`Elev 1`, `Elev 2`, ...). När du behöver veta vem `Elev 7` är skriver du ut en HTML-nyckel fysiskt och raderar filen.

## Arkitektur

Tunn Python-orkestrering ovanpå [`@googleworkspace/cli`](https://github.com/googleworkspace/cli) (`gws`). All auth och alla API-anrop går via `gws` — vi formaterar bara JSON till markdown.

```
classroom.py     CLI (argparse)
├── gws_client   Subprocess-wrapper kring `gws classroom ...` + `gws drive files export`
├── anonymize    Per-kurs alias-mapping (userId -> "Elev N")
├── summary      Markdown-generator: kurs-metadata (anonymiserad)
├── key          HTML-nyckel-generator (alias -> namn)
├── drive        Hämtar inlämnings-text från Drive (Google Docs → text/plain)
├── cache        Disk-cache för Drive-export (24h TTL, gitignored)
└── submissions  Markdown-generator: inlämnings-text (read/dump)
```

Inga Python-beroenden utöver standardbiblioteket.

## Setup

### 1. Installera `gws` och `gcloud`

```bash
npm install -g @googleworkspace/cli
curl https://sdk.cloud.google.com | bash && exec -l $SHELL
```

### 2. Konfigurera auth

```bash
# Logga in med ditt Google-konto i gcloud
gcloud auth login

# Sätt aktivt projekt (eller skapa nytt med gcloud projects create)
gcloud config set project <project-id>

# Aktivera Classroom API i projektet
gcloud services enable classroom.googleapis.com

# Skapa OAuth-client (alternativt manuellt i Cloud Console)
gws auth setup

# Lägg env-flagga om du saknar gnome-keyring/secret-tool
echo 'export GOOGLE_WORKSPACE_CLI_KEYRING_BACKEND=file' >> ~/.bashrc

# Logga in - kryssa i alla scope-toggles på consent screen
gws auth login --scopes \
  https://www.googleapis.com/auth/classroom.courses.readonly,\
https://www.googleapis.com/auth/classroom.rosters.readonly,\
https://www.googleapis.com/auth/classroom.profile.emails,\
https://www.googleapis.com/auth/classroom.coursework.students.readonly,\
https://www.googleapis.com/auth/classroom.student-submissions.students.readonly,\
https://www.googleapis.com/auth/classroom.announcements.readonly,\
https://www.googleapis.com/auth/classroom.topics.readonly,\
https://www.googleapis.com/auth/classroom.courseworkmaterials.readonly,\
https://www.googleapis.com/auth/drive.readonly,\
https://www.googleapis.com/auth/forms.body.readonly,\
https://www.googleapis.com/auth/forms.responses.readonly
```

**Forms-scopes** (sista tre): `classroom.profile.emails` behövs för att mappa Form-respondenter (email) till elev-alias. `forms.body.readonly` + `forms.responses.readonly` läser frågorna och svaren.

**SSH-tips:** OAuth-callback går till `localhost:NNNNN` på den maskin där `gws` körs. Om du är SSH-inloggad behöver du port forwarda: `~C` följt av `-L NNNNN:localhost:NNNNN` i den interaktiva SSH-prompten, eller starta en separat SSH-session med `-L`-flaggan.

## Användning

### Metadata-pipeline

```bash
# Lista mina aktiva kurser
./run.sh list

# Sammanställning (anonymiserad) för en kurs
./run.sh summary 23437726298

# Spara till vault
./run.sh summary 23437726298 -o "/home/anders/Second brain/Brain/00-Inbox/Mek24b-2026-05-06.md"

# Generera HTML-nyckel (alias -> namn) för fysisk utskrift
./run.sh key 23437726298
# -> Skriver keys/<id>-<datum>.html. Öppna, skriv ut, radera.
```

### Inlämnings-text (Drive)

`coursework_id` hittar du i URL:en till uppgiften i Classroom (`/c/<courseId>/a/<courseworkId>`) eller i `summary`-länkarna.

```bash
# Läs en specifik elevs inlämning (alias från summary)
./run.sh read 23437726298 654321 "Elev 7"
./run.sh read 23437726298 654321 7              # samma sak

# Läs alla inlämningar för en uppgift (anonymiserat) → klassanalys
./run.sh dump 23437726298 654321 -o /tmp/uppgift.md

# Hoppa över cache vid ny iteration
./run.sh dump 23437726298 654321 --no-cache

# Cache-städning
./run.sh cache --purge   # radera utgånget (default-beteende)
./run.sh cache --clear   # radera allt
```

Output är markdown med banner högst upp om GDPR-läckage i fritext, en sektion per `Elev N` med inlämnad text och bilage-titlar. Google Docs exporteras som klartext; uppladdade .docx, .pdf och .odt konverteras lokalt. Övrigt (länkar, YouTube, bilder, legacy-format) listas men öppnas i browser.

### Google Forms-uppgifter

`read` och `dump` detekterar automatiskt om uppgiften har en Form bifogad och hämtar svaren via Forms API istället för Drive. Output följer samma struktur — alias per elev, sektion per fråga, GDPR-banner.

För quiz-läge visas poäng per fråga (`✓`/`✗`) och totalt. Frågor i ordning, sortering på `Elev N`. Externa svar (respondenter som inte är medlemmar i kursens roster) listas separat som `Externt svar N` så datan inte tappas.

**Förutsättning:** "Samla in e-postadresser" måste vara på i Forms-inställningarna — annars går det inte att mappa svar till elev-alias.

**Tips för analys:** pipa direkt till en LLM-prompt:

```bash
./run.sh dump <courseId> <workId> | claude -p "Sammanfatta gemensamma missuppfattningar i klassens svar."
```

## Anonymisering

- **`aliases.json`** lagrar `course_id -> {userId: "Elev N"}`. Inga riktiga namn. Gitignored.
- Nya elever tilldelas nästa lediga nummer i kursen vid första körningen.
- En elev som finns i två kurser får olika alias i varje (per-kurs-stabilt, inte cross-course).
- `summary` visar bara metadata. **Aldrig** elevtext eller riktiga namn.
- `key` hämtar roster färskt från API:et och skriver en HTML-fil i `keys/`. Filen är gitignored. Skriv ut fysiskt och radera filen efteråt.

### GDPR — läs detta innan `read` / `dump` används

`read` och `dump` hämtar **fri elev-text**. Pseudonymiseringen byter ut namn-ramverket men kan inte filtrera bort identifierande uppgifter inuti texten ("jag heter X", "min mamma X", "i min klass på Y"). Output får därför en banner högst upp som varnar.

Konsekvenser:

- Behandla genererade markdown-filer som personuppgifter. Ligg inte kvar med dem i vaultet längre än nödvändigt.
- Spara hellre under `/tmp/` än i `Brain/00-Inbox/` när du jobbar tillfälligt.
- Cache (se nedan) ligger på disk i 24 h — `./run.sh cache --clear` om datorn ska delas eller läggas undan.

### Cache

Drive-text cachas i `cache/` (gitignored) i 24 h så att iterativa körningar inte slår API:et i onödan. Cache-nyckel = Drive file id. Filer markeras `_(cache)_` i output. För färska data: `--no-cache` per körning eller `cache --clear`.

## Filstruktur

```
classroom-tool/
├── classroom.py        # CLI
├── gws_client.py       # gws subprocess-wrapper (Classroom + Drive + Forms)
├── anonymize.py        # alias-mapping
├── summary.py          # metadata-generator (kurssammanställning)
├── key.py              # HTML-nyckel-generator
├── drive.py            # Drive-export-orkestrering
├── converters.py       # PDF/DOCX/ODT → text (pure-Python)
├── cache.py            # disk-cache för Drive-text
├── submissions.py      # markdown-generator för inlämnings-text (Drive)
├── forms.py            # markdown-generator för Forms-svar
├── run.sh              # python3 wrapper
├── aliases.json        # (gitignored) userId -> alias
├── keys/               # (gitignored) genererade HTML-nycklar
├── cache/              # (gitignored) cachad Drive-text (24h TTL)
└── README.md
```

## Begränsningar

- Auth-token bor i `~/.config/gws/credentials.enc` (krypterad). Refresh tokens upphör efter 7 dagar i Testing-läge → relogga ~1 ggr/vecka.
- `aliases.json` innehåller Google-userId:n (personidentifierare). Den är gitignored men ligger på disk.
- **Google Docs** exporteras till text via Drive API. **Uppladdade .docx, .pdf, .odt** konverteras lokalt med pure-Python (pypdf, python-docx, odfpy). Bilder, legacy .doc/.rtf, länkar och YouTube listas men exporteras inte — lös ad-hoc i browsern.
- **Forms-uppgifter** kräver att "Samla in e-postadresser" är på i Forms-inställningarna. Filuppladdnings-frågor i Forms räknas men innehållet hämtas inte.
- `submissionHistory` (revisioner, kommentarer från eleven) tas inte med — bara den senaste inlämnade texten.
