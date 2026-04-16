# Schulhelfer – Primarstufe Rittergasse Basel

**Helferanmeldung für Schulanlässe** der Primarstufe Rittergasse Basel (Kindergarten & Primarschule).

Ein barrierefreies, mobil-optimiertes Tool zur Rekrutierung von Eltern-Helfern für Schulveranstaltungen.

---

## Features

- **Mobile-First Design** – Perfekt für Smartphones
- **WCAG 2.1 AA** – Vollständig barrierefrei
- **Dark Mode** – Automatische Anpassung ans System
- **PWA** – Installierbar auf Home Screen (Android, iOS, Desktop)
- **Keine Datenbank** – Daten in Google Sheets
- **Schnell** – Statische Dateien via GitHub Pages, Cache via Service Worker
- **Automatische Wiederholung** – Retry-Logik bei Netzwerkfehlern
- **Formular-Speicherung** – Eingaben werden lokal zwischengespeichert
- **Sicherheit** – Rate Limiting, Honeypot, Input-Sanitization, Transaktionssicherheit
- **E-Mail-Benachrichtigungen** – Optionale Benachrichtigungen bei Anmeldungen
- **Audit-Log** – Vollständige Protokollierung aller Aktionen
- **Datenexport** – Export-Funktion für Anmeldungen
- **Helferliste (Word)** – Admin-geschützter .docx-Download pro Anlass
- **Kalender-Download** – iCal (.ics) Export für jeden Anlass
- **Datenschutz** – DSG-konforme Datenschutzerklärung integriert
- **Datenbereinigung** – Funktion zum Löschen alter Anlässe und Anmeldungen

---

## Projektstruktur

```
schulhelfer/
├── index.html              ← Hauptseite (PWA-fähig)
├── css/styles.css          ← Styling (Light & Dark Mode)
├── js/app.js               ← Interaktion & Client-Logik
├── service-worker.js       ← Offline-Caching
├── manifest.webmanifest    ← PWA-Manifest
├── icons/
│   ├── icon.svg            ← App-Icon
│   └── icon-maskable.svg   ← Maskable Icon (Safe Zone)
├── google-apps-script/
│   └── Code.gs             ← Backend (Google Apps Script)
└── README.md
```

---

## Einrichtung

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

### Schritt 3: GitHub Repository konfigurieren

1. Erstellen Sie ein neues Repository auf GitHub
2. Laden Sie alle Dateien hoch
3. Öffnen Sie `index.html`
4. Ersetzen Sie die `API_URL` im CONFIG-Block mit der kopierten Apps Script URL
5. Ersetzen Sie die `SPREADSHEET_URL` mit dem Link zu Ihrem Google Sheet
6. Committen Sie die Änderungen

### Schritt 4: GitHub Pages aktivieren

1. Gehen Sie zu **Settings → Pages**
2. Source: **Deploy from a branch**
3. Branch: **main**, Ordner: **/ (root)**
4. Klicken Sie **Save**
5. Nach ca. 1 Minute ist Ihre Seite verfügbar

---

## Google Sheet Struktur

### Anlässe (automatisch erstellt)
| ID | Name | Datum | Zeit | Benötigte Helfer | Angemeldete | Beschreibung |
|----|------|-------|------|------------------|-------------|--------------|
| 1  | Sommerfest | 15.07.2025 | 14:00-18:00 | 5 | 2 | Hilfe beim Grill |

### Anmeldungen (automatisch erstellt)
| Zeitstempel | Anlass-ID | Name | E-Mail | Telefon | Anlass |
|-------------|-----------|------|--------|---------|--------|
| 30.11.2024 | 1 | Max Muster | max@mail.ch | +41 79 123 45 67 | Sommerfest |

---

## Anlässe verwalten

Im Google Sheet:
1. Öffnen Sie das Menü **🏰 Schulhelfer**
2. Wählen Sie **Neuer Anlass hinzufügen**
3. Füllen Sie das Formular aus

Oder direkt in der Tabelle "Anlässe":
- Neue Zeile hinzufügen
- ID muss eindeutig sein
- Datum im Format TT.MM.JJJJ

### Menü-Funktionen

| Menüpunkt | Beschreibung |
|-----------|-------------|
| Erstes Setup | Initialisiert die Tabelle mit Beispiel-Anlässen |
| Neuer Anlass hinzufügen | Dialog zum Erstellen eines neuen Anlasses |
| Alle Anmeldungen anzeigen | Wechselt zum Anmeldungen-Tab |
| Zähler neu berechnen | Gleicht die Anmeldezahlen mit den tatsächlichen Zeilen ab |
| Daten exportieren | Erstellt einen Export-Tab mit allen Daten |
| Alte Anlässe bereinigen | Löscht vergangene Anlässe und zugehörige Anmeldungen |
| Admin-Status prüfen | Zeigt Info zum Admin-Schlüssel |
| Audit-Log anzeigen | Wechselt zum Audit-Log-Tab |

---

## Admin-Zugang (Helferliste)

Die Namen der angemeldeten Helfer sind auf der öffentlichen Seite bewusst nicht
sichtbar. Der Helferlisten-Download ist durch einen Admin-Schlüssel geschützt.

### Einrichtung

1. Im Apps Script Editor: **Projekt-Einstellungen → Script Properties**
2. Eigenschaft hinzufügen: `ADMIN_KEY` = `<langer, zufälliger Wert>` (mind. 16 Zeichen)
3. Web-App neu bereitstellen

### Verwendung

1. Den Schlüssel einmalig in der URL anhängen:
   `https://…/schulhelfer/?admin=IHR_SCHLÜSSEL`
2. Der Schlüssel wird im Browser gespeichert und aus der URL entfernt
3. Auf jeder Anlass-Karte erscheint die Schaltfläche **„Helferliste (Word)"**
4. Unten links zeigt eine kleine Plakette den Admin-Modus an
5. Zum Abmelden auf „Abmelden" in der Plakette klicken

### Helferliste (Word)

Pro Anlass kann eine Helferliste mit Name und Telefonnummer als Word-Dokument
(.docx) exportiert werden. Leere Zeilen bis zur maximalen Helferzahl bieten
Platz für kurzfristige Eintragungen am Anlass-Tag.

---

## PWA (Progressive Web App)

Die Seite kann als App auf dem Home Screen installiert werden:

- **Android / Desktop Chrome**: Der Browser bietet automatisch die Installation an, oder über das Menü → „App installieren"
- **iOS Safari**: Teilen-Symbol → „Zum Home-Bildschirm"
- **In-App-Browser** (WhatsApp, Instagram etc.): Zuerst Link kopieren und in Safari/Chrome öffnen

Die App funktioniert nach der Installation auch bei eingeschränkter Konnektivität
dank Service Worker Caching.

---

## Datenschutz

- Daten liegen im Google Sheet der Schulleitung
- Nur Name, E-Mail und optional Telefonnummer werden erhoben
- Alte Daten können über das Menü bereinigt werden
- Integrierte Datenschutzerklärung im Footer der Website
- E-Mail-Benachrichtigungen: `ADMIN_EMAIL` in `Code.gs` setzen (optional)

---

## FAQ

**Wie teile ich das Tool?**
Senden Sie die GitHub Pages URL per E-Mail oder Chat an die Eltern.

**Kann ich das Design anpassen?**
Ja! Ändern Sie die Farben in `css/styles.css` unter `:root`.

**Wie viele Helfer können sich anmelden?**
Unbegrenzt – das Limit setzen Sie pro Anlass in der Spalte "Benötigte Helfer".

**Was passiert bei vielen gleichzeitigen Anmeldungen?**
Das Backend verwendet LockService für Transaktionssicherheit. Bei hoher Last
erhalten Nutzer eine freundliche „bitte warten"-Meldung.

**Werden Daten geschützt?**
Ja. Rate Limiting, Honeypot-Felder und Input-Validierung auf Server- und
Client-Seite. Daten nur im Google Sheet der Schulleitung.

---

## Lizenz

MIT License – Frei verwendbar für Schulen.

---

**Primarstufe Rittergasse Basel**
Kindergarten & Primarschule
