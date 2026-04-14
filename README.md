# 🏰 Schulhelfer – Primarstufe Rittergasse Basel

**Helferanmeldung für Schulanlässe** der Primarstufe Rittergasse Basel (Kindergarten & Primarschule).

Ein barrierefreies, mobil-optimiertes Tool zur Rekrutierung von Eltern-Helfern für Schulveranstaltungen.

---

## ✨ Features

- 📱 **Mobile-First Design** – Perfekt für Smartphones
- ♿ **WCAG 2.1 AA** – Vollständig barrierefrei
- 🌙 **Dark Mode** – Automatische Anpassung
- 🔒 **Keine Datenbank** – Daten in Google Sheets
- ⚡ **Schnell** – Lädt in unter 2 Sekunden
- 🔄 **Automatische Wiederholung** – Retry-Logik bei Netzwerkfehlern
- 💾 **Formular-Speicherung** – Daten werden lokal gespeichert
- 🛡️ **Sicherheit** – Rate Limiting, Input-Sanitization, Transaktionssicherheit
- 📧 **E-Mail-Benachrichtigungen** – Optionale Benachrichtigungen bei Anmeldungen
- 📊 **Audit-Log** – Vollständige Protokollierung aller Aktionen
- 📥 **Datenexport** – Export-Funktion für Anmeldungen

---

## 📁 Projektstruktur

```
schulhelfer/
├── index.html              ← Hauptseite
├── css/styles.css          ← Styling
├── js/app.js               ← Interaktion
├── google-apps-script/
│   └── Code.gs             ← Backend
└── README.md
```

---

## 🚀 Einrichtung

### Schritt 1: Google Sheet einrichten

1. Öffnen Sie [Google Sheets](https://sheets.google.com)
2. Erstellen Sie eine neue Tabelle
3. Gehen Sie zu **Erweiterungen → Apps Script**
4. Löschen Sie den vorhandenen Code
5. Kopieren Sie den Inhalt von `google-apps-script/Code.gs`
6. Speichern Sie (Ctrl+S)
7. Führen Sie `erstesSetup` aus dem Dropdown aus
8. Erlauben Sie die Berechtigungen

### Schritt 2: Als Web-App bereitstellen

1. Klicken Sie auf **Bereitstellen → Neue Bereitstellung**
2. Wählen Sie Typ: **Web-App**
3. Einstellungen:
   - **Ausführen als:** Ich
   - **Zugriff:** Jeder
4. Klicken Sie auf **Bereitstellen**
5. **Kopieren Sie die URL** (beginnt mit `https://script.google.com/...`)

### Schritt 3: GitHub Repository erstellen

1. Erstellen Sie ein neues Repository auf GitHub
2. Laden Sie alle Dateien hoch
3. Öffnen Sie `index.html`
4. Ersetzen Sie `IHRE_GOOGLE_APPS_SCRIPT_URL_HIER` mit der kopierten URL
5. Committen Sie die Änderung

### Schritt 4: GitHub Pages aktivieren

1. Gehen Sie zu **Settings → Pages**
2. Source: **Deploy from a branch**
3. Branch: **main**, Ordner: **/ (root)**
4. Klicken Sie **Save**
5. Nach ca. 1 Minute ist Ihre Seite unter `https://BENUTZERNAME.github.io/REPONAME/` verfügbar

---

## 📊 Google Sheet Struktur

### Anlässe (automatisch erstellt)
| ID | Name | Datum | Zeit | Benötigte Helfer | Angemeldete | Beschreibung |
|----|------|-------|------|------------------|-------------|--------------|
| 1  | Sommerfest | 15.07.2025 | 14:00-18:00 | 5 | 2 | Hilfe beim Grill |

### Anmeldungen (automatisch erstellt)
| Zeitstempel | Anlass-ID | Name | E-Mail | Telefon | Anlass |
|-------------|-----------|------|--------|---------|--------|
| 30.11.2024 | 1 | Max Muster | max@mail.ch | 079... | Sommerfest |

---

## 🛠 Anlässe verwalten

Im Google Sheet:
1. Öffnen Sie das Menü **🏰 Schulhelfer**
2. Wählen Sie **Neuer Anlass**
3. Füllen Sie das Formular aus

Oder direkt in der Tabelle "Anlässe":
- Neue Zeile hinzufügen
- ID muss eindeutig sein
- Datum im Format TT.MM.JJJJ

### Neue Funktionen

- **Daten exportieren**: Menü → "Daten exportieren (CSV)" – Exportiert alle Anlässe und Anmeldungen
- **Audit-Log anzeigen**: Menü → "Audit-Log anzeigen" – Zeigt alle Systemaktionen
- **E-Mail-Benachrichtigungen**: Setzen Sie `ADMIN_EMAIL` in `Code.gs` (Zeile 13) für Benachrichtigungen
- **Helferliste als Word (.docx) herunterladen** *(nur Admin)*: Pro Anlass kann eine
  Helferliste mit Name, Telefonnummer und Unterschriftsspalte als Word-Dokument
  exportiert werden – ideal als Ausdruck für den Einsatztag.

#### Admin-Zugang für die Helferliste einrichten

Die Namen der angemeldeten Helfer sind auf der öffentlichen Seite bewusst nicht
sichtbar. Der Helferlisten-Download ist deshalb durch einen Admin-Schlüssel
geschützt:

1. In `google-apps-script/Code.gs` die Konstante `ADMIN_KEY` auf einen langen,
   zufälligen Wert setzen (z.B. `var ADMIN_KEY = 'xG7p…';`) und neu
   bereitstellen.
2. Den Schlüssel einmalig in der URL anhängen:
   `https://…/schulhelfer/?admin=xG7p…`
3. Der Schlüssel wird im Browser gespeichert, anschliessend aus der URL
   entfernt, und auf jeder Anlass-Karte erscheint die Schaltfläche
   **„Helferliste (Word)"**. Unten links zeigt eine kleine Plakette
   den Admin-Modus an.
4. Zum Abmelden auf „Abmelden" in der Plakette klicken – oder Browser-Daten
   löschen.

Nicht-Admin-Besucher:innen sehen die Schaltfläche gar nicht und können die Namen
nicht abrufen.

---

## ❓ FAQ

**Wie teile ich das Tool?**  
Senden Sie die GitHub Pages URL per E-Mail an die Eltern.

**Kann ich das Design anpassen?**  
Ja! Ändern Sie die Farben in `css/styles.css` unter `:root`.

**Wie viele Helfer können sich anmelden?**  
Unbegrenzt – das Limit setzen Sie pro Anlass in der Spalte "Benötigte Helfer".

**Werden Daten geschützt?**  
Die Daten liegen in Ihrem Google Sheet. Nur Sie haben Zugriff. Das System verwendet Rate Limiting, Input-Sanitization und Transaktionssicherheit.

**Was ist das Audit-Log?**  
Alle Anmeldungen und Systemaktionen werden protokolliert. Sie finden das Log im Tab "Audit-Log" im Google Sheet.

**Wie funktioniert die Formular-Speicherung?**  
Ihre Eingaben werden lokal im Browser gespeichert und automatisch wiederhergestellt, falls die Seite versehentlich geschlossen wird.

**Was ist Rate Limiting?**  
Das System verhindert Missbrauch durch Begrenzung der Anfragen pro Zeitfenster (10 Anfragen pro Minute).

---

## 📄 Lizenz

MIT License – Frei verwendbar für Schulen.

---

**Primarstufe Rittergasse Basel**  
Kindergarten & Primarschule
