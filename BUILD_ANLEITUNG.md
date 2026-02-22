# 🚀 NWG-Bericht Converter - Anleitung zur .exe-Erstellung

## Schnellstart

1. **Doppelklick auf `build.bat`** - Das ist alles! 🎉
2. Warten bis "BUILD ERFOLGREICH ABGESCHLOSSEN!" erscheint
3. Im `Release`-Ordner finden Sie die fertige `NWG-Bericht-Converter.exe`

## Was passiert beim Build?

Das Build-Skript:
- ✅ Installiert automatisch PyInstaller
- ✅ Erstellt eine einzelne .exe-Datei (ca. 50-80 MB)
- ✅ Bindet alle Ressourcen ein (Logo, Beraterliste, Word-Vorlage)
- ✅ Setzt das Icon für die .exe
- ✅ Erstellt ein Release-Paket mit README

## Verteilung

Die fertige `NWG-Bericht-Converter.exe` kann:
- ✅ Auf jeden Windows-Computer kopiert werden
- ✅ Ohne Python-Installation ausgeführt werden
- ✅ Ohne zusätzliche Dateien laufen (alles ist eingebettet)
- ✅ Per E-Mail, USB-Stick oder Download verteilt werden

## Dateigröße

Die .exe wird etwa 50-80 MB groß, weil sie enthält:
- Python-Interpreter
- Alle Python-Bibliotheken (tkinter, pandas, python-docx, PIL, etc.)
- Ihre Anwendung + Ressourcen
- Windows-Kompatibilitäts-Layer

## Problembehandlung

**Problem**: Build bricht ab
**Lösung**: Stellen Sie sicher, dass alle benötigten Dateien vorhanden sind:
- NWG_Converter.py
- logo.jpg
- Energieberaterliste_T2.xlsx
- NWG-Bericht_Converter_Vorlage_V1.0.docx
- Converter_logo.ico

**Problem**: .exe startet nicht
**Lösung**: 
- Windows Defender/Antivirus ausschalten während der Erstellung
- .exe als "Vertrauenswürdig" markieren

**Problem**: "Datei nicht gefunden" in der .exe
**Lösung**: Alle Ressourcen wurden korrekt eingebettet, starten Sie die .exe vom Desktop aus

## Automatisches Update

Um die Anwendung zu aktualisieren:
1. Code ändern
2. `build.bat` erneut ausführen (oder `python build_app.py`)
3. Neue .exe verteilen

## Icon anpassen

Das Icon kann geändert werden durch:
1. Neue .ico-Datei als `Converter_logo.ico` speichern
2. Build-Prozess erneut ausführen

Viel Erfolg! 🎉
