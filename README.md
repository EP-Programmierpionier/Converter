# 🚀 NWG-Bericht Converter

> **Automatisierte Erstellung von NWG-Berichten aus Excel-Daten**

## 📄 Word-Vorlage Information

**Standard-Vorlage:** `Vorlagen/NWG-Bericht_Converter_Vorlage_V1.0.docx`
- Wird automatisch beim Start geladen
- Enthält Content Controls für Datenaustausch
- Kann über "Import Word-Vorlage" Button geändert werden

## 🎯 Schnellstart

### Für Benutzer:
1. **Doppelklick auf** `NWG-Bericht-Converter.exe`
2. **Excel-Datei** per Drag & Drop in die grüne Zone ziehen
3. **Energieberater** aus der Liste auswählen (automatisch aus `Vorlagen/Energieberaterliste_T2.xlsx`)
4. **"🚀 Bericht erstellen"** klicken
5. **Speicherort** wählen - fertig! 

### Für Entwickler:
1. **Doppelklick auf** `Dev/start_dev.bat`
2. Automatische Installation aller Python-Pakete
3. Anwendung startet direkt

## 📁 Saubere Struktur

```
NWG-Bericht Converter/
├── 📱 NWG-Bericht-Converter.exe   # ← Fertige Anwendung
├── 🐍 NWG_Converter.py             # ← Python-Version  
├── 📋 README.md                    # ← Diese Datei
├── 🔧 create_shortcut.ps1          # Desktop-Shortcut (optional)
├── ⚡ start_dev.bat/.ps1           # Entwicklung starten
├── 🏗️ build.bat                    # .exe erstellen (Starter)
├── 🏗️ build_app.py                 # .exe erstellen (Python)
├── 📋 requirements.txt             # Python-Abhängigkeiten
├── 📂 Vorlagen/                    # Alle Vorlagendateien
│   ├── logo.jpg                    # App-Logo
│   ├── Converter_logo.ico          # App-Icon
│   ├── Energieberaterliste_T2.xlsx # Berater-Datenbank
│   └── NWG-Bericht_Converter_Vorlage_V1.0.docx  # Standard-Vorlage
├── 📂 Logs/                        # Runtime-Protokolle
```

Hinweis: Für den Betrieb werden die Dateien im Ordner `Vorlagen/` benötigt (mindestens Beraterliste + Word-Vorlage).

## ⚡ Features

- 🎯 **Drag & Drop** - Excel-Dateien einfach reinziehen
- 👥 **Energieberater-Liste** - Automatische Auswahl aus Datenbank
- 🔄 **Content Control Ersetzung** - Intelligente Word-Verarbeitung
- 📝 **Fehlende Tags anzeigen** - Übersicht über nicht gefüllte Platzhalter
- 💾 **Pfad-Speicherung** - Merkt sich letzte Dateipfade
- 🎨 **Moderne GUI** - Saubere, benutzerfreundliche Oberfläche

---
*Erstellt mit ❤️ für effiziente NWG-Berichterstattung*