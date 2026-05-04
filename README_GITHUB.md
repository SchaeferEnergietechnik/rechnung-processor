# GitHub Actions Workflow

Automatischer Build der .exe bei jedem Push.

## Setup

### 1. Repository erstellen

Auf GitHub: https://github.com/new
- Name: `rechnung-processor` (oder wie du willst)
- Public oder Private (beides geht)

### 2. Token erstellen

GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)

**Scopes:** `repo` (für Private Repos) oder `public_repo` (für Public Repos)

### 3. Dateien pushen

```bash
# Lokalen Ordner initialisieren
cd rechnungsbearbeitung
git init
git add .
git commit -m "Initial commit"

# Mit GitHub verbinden (TOKEN ersetzen)
git remote add origin https://TOKEN@github.com/DEIN_USERNAME/rechnung-processor.git
git branch -M main
git push -u origin main
```

## Workflow

### Automatischer Build
Bei jedem Push zu `main`:
1. Python-Code wird getestet
2. .exe wird gebaut
3. Als "Artifact" hochgeladen (gültig 90 Tage)

### .exe Downloaden
GitHub → Actions → Letzter Workflow-Run → Artifacts → `RechnungProcessor-Windows`

### Manueller Build
GitHub → Actions → "Build Windows Executable" → "Run workflow"

### Release mit .exe
```bash
# Tag erstellen und pushen
git tag v1.0.0
git push origin v1.0.0
```
→ Erstellt automatisch GitHub-Release mit .exe angehängt

## Ordner-Struktur im Repo

```
rechnung-processor/
├── .github/
│   └── workflows/
│       └── build.yml          # Workflow-Definition
├── rechnung_processor.py      # Haupt-Code
├── requirements.txt           # Abhängigkeiten
└── README.md                  # Dokumentation
```

## Update-Prozess

```bash
# Änderungen machen
# ... editiere rechnung_processor.py ...

git add .
git commit -m "Fix: Bessere Datumserkennung"
git push origin main
```

→ GitHub Actions baut neue .exe automatisch (ca. 2-3 Minuten)

## Fehlerbehebung

**Build failed?**
GitHub → Actions → Roter Workflow → Logs anschauen

**Kein Artifact?**
Workflow muss erfolgreich durchlaufen (grüner Haken)

**Token abgelaufen?**
Neues Token erstellen und Remote-URL aktualisieren:
```bash
git remote set-url origin https://NEUER_TOKEN@github.com/.../...
```
