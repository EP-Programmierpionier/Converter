#!/usr/bin/env python3
"""
Build-Skript für NWG-Bericht Converter
Erstellt eine ausführbare .exe-Datei mit allen benötigten Ressourcen
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def install_dependencies():
    """Installiert alle benötigten Python-Pakete."""
    packages = [
        "python-docx>=0.8.11", 
        "tkinterdnd2>=0.3.0",
        "openpyxl>=3.0.0",
        "pyinstaller>=5.0.0"
    ]
    
    print("📦 Installiere Python-Pakete...")
    for package in packages:
        try:
            result = subprocess.run([sys.executable, "-m", "pip", "install", package], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                print(f"   ✅ {package}")
            else:
                print(f"   ❌ {package} - {result.stderr}")
                return False
        except Exception as e:
            print(f"   ❌ {package} - {e}")
            return False
    
    return True

def test_imports():
    """Testet ob alle Module importiert werden können."""
    print("\n🧪 Teste Module-Imports...")
    modules = ["docx", "tkinterdnd2", "openpyxl"]
    
    for module in modules:
        try:
            __import__(module)
            print(f"   ✅ {module}")
        except ImportError as e:
            print(f"   ❌ {module} - {e}")
            return False
    
    return True

def main():
    """Hauptfunktion zum Erstellen der ausführbaren Anwendung."""


    # Aktuelles Verzeichnis
    base_dir = Path(__file__).parent
    os.chdir(base_dir)
    
    print("🚀 NWG-Bericht Converter - Build-Prozess gestartet")
    print("=" * 60)
    
    # 1. Dependencies installieren
    if not install_dependencies():
        print("\n❌ FEHLER: Konnte nicht alle Pakete installieren")
        return False
    
    # 2. Import-Test
    if not test_imports():
        print("\n❌ FEHLER: Module-Import fehlgeschlagen")
        return False
    
    # 3. Überprüfen ob alle benötigten Dateien vorhanden sind
    required_files = [
        "NWG_Converter.py",
        "Vorlagen/logo.jpg", 
        "Vorlagen/Energieberaterliste_T2.xlsx",
        "Vorlagen/NWG-Bericht_Converter_Vorlage_V1.0.docx",
        "Vorlagen/Converter_logo.ico"
    ]
    
    print("\n📋 Überprüfe benötigte Dateien...")
    missing_files = []
    for file in required_files:
        if not (base_dir / file).exists():
            missing_files.append(file)
            print(f"   ❌ {file} - FEHLT")
        else:
            print(f"   ✅ {file}")
    
    if missing_files:
        print(f"\n❌ FEHLER: Folgende Dateien fehlen: {', '.join(missing_files)}")
        return False
    
    # 4. Aufräumen - alte Build-Ordner löschen
    print("\n🧹 Aufräumen alter Build-Dateien...")
    for folder in ["build", "dist", "__pycache__"]:
        folder_path = base_dir / folder
        if folder_path.exists():
            shutil.rmtree(folder_path)
            print(f"   🗑️  {folder} gelöscht")
    
    # .spec Dateien löschen
    for spec_file in base_dir.glob("*.spec"):
        spec_file.unlink()
        print(f"   🗑️  {spec_file.name} gelöscht")
    
    # 5. PyInstaller-Befehl zusammenstellen
    print("\n⚙️  Erstelle ausführbare Datei...")
    
    # Icon-Pfad korrekt setzen (lokal im Vorlagen-Ordner)
    icon_path = base_dir / "Vorlagen" / "Converter_logo.ico"
    
    # Icon-Pfad validieren
    if not icon_path.exists():
        print(f"     Icon nicht gefunden: {icon_path}")
        print("     .exe wird ohne Icon erstellt")
        icon_param = []
    else:
        print(f"   ✅ Icon gefunden: {icon_path}")
        # Prüfe Icon-Größe
        icon_size = icon_path.stat().st_size
        print(f"   📏 Icon-Größe: {icon_size} Bytes")
        icon_param = [f"--icon={icon_path.absolute()}"]
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                           # Eine einzige .exe-Datei
        "--windowed",                          # Kein Konsolen-Fenster
        "--name=NWG-Bericht-Converter",       # Name der .exe
        *icon_param,                           # Icon für die .exe (falls vorhanden)
        "--add-data=Vorlagen/logo.png;.",   # Logo einbetten
        "--hidden-import=pandas",              # Pandas explizit einbinden
        "--hidden-import=openpyxl",            # openpyxl für Excel
        "--hidden-import=tkinterdnd2",         # Drag & Drop
        "--hidden-import=PIL",                 # Pillow für Bilder
        "--hidden-import=docx",                # python-docx für Word
        "--clean",                             # Cache bereinigen
        "--noconfirm",                         # Keine Bestätigung
        "--exclude-module=matplotlib",         # Unnötige Module ausschließen
        "--exclude-module=pandas",
        "--exclude-module=numpy",
    ]

    
    try:
        print("   🔄 PyInstaller wird ausgeführt...")
        result = subprocess.run(cmd, cwd=base_dir, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("   ✅ PyInstaller erfolgreich ausgeführt")
        else:
            print("   ❌ PyInstaller Fehler:")
            print(result.stderr)
            print(result.stdout)
            return False
    
    except Exception as e:
        print(f"   ❌ Fehler beim Ausführen von PyInstaller: {e}")
        return False
    
    # 6. Überprüfen ob .exe erstellt wurde
    exe_path = base_dir / "dist" / "NWG-Bericht-Converter.exe"
    if exe_path.exists():
        file_size = exe_path.stat().st_size / (1024 * 1024)  # MB
        print(f"   ✅ .exe-Datei erstellt: {exe_path}")
        print(f"   📦 Dateigröße: {file_size:.1f} MB")
    else:
        print("   ❌ .exe-Datei wurde nicht erstellt")
        return False
    
    # 7. .exe ins Hauptverzeichnis kopieren
    print("\n📦 Kopiere .exe ins Hauptverzeichnis...")
    main_dir = base_dir.parent  # Hauptverzeichnis (Parent von Entwicklung/)
    exe_target = main_dir / "NWG-Bericht-Converter.exe"
    
    try:
        shutil.copy2(exe_path, exe_target)
        print(f"   ✅ .exe kopiert nach: {exe_target}")
        
        # Icon-Test
        if icon_path.exists():
            print("   🎨 Icon sollte in Windows Explorer/Desktop sichtbar sein")
            print("   💡 Tipp: Falls Icon nicht sichtbar - Windows Icon Cache leeren:")
            print("        ie4uinit.exe -show")
        else:
            print("     Kein Icon eingebettet - .exe hat Standard-Icon")
            
    except Exception as e:
        print(f"   ❌ Fehler beim Kopieren: {e}")
        return False

    # 8. Release-Ordner erstellen
    print("\n📦 Erstelle Release-Paket...")
    release_dir = base_dir / "Release"
    if release_dir.exists():
        shutil.rmtree(release_dir)
    
    release_dir.mkdir()
    
    # .exe kopieren
    shutil.copy2(exe_path, release_dir / "NWG-Bericht-Converter.exe")
    
    # README erstellen
    readme_content = """# NWG-Bericht Converter

## Installation
1. Laden Sie die Datei 'NWG-Bericht-Converter.exe' herunter
2. Führen Sie die .exe-Datei aus
3. Das ist alles! 🎉

## Verwendung
1. Energieberater aus der Liste auswählen
2. Excel-Datei (Pfadfinder) laden oder per Drag & Drop hineinziehen
3. Optional: Andere Word-Vorlage importieren
4. "Bericht erstellen" klicken
5. Speicherort für den fertigen Bericht wählen

## Systemanforderungen
- Windows 10/11
- Keine zusätzliche Software erforderlich

## Support
Bei Fragen wenden Sie sich an Elia Salemi.

## Version
Version 1.0 - Build """ + str(Path().cwd().name) + """
"""
    
    with open(release_dir / "README.txt", "w", encoding="utf-8") as f:
        f.write(readme_content)
    
    print(f"   ✅ Release-Paket erstellt: {release_dir}")
    print(f"   📁 Inhalt:")
    for item in release_dir.iterdir():
        size = item.stat().st_size / (1024 * 1024) if item.is_file() else 0
        print(f"      - {item.name} ({size:.1f} MB)" if item.is_file() else f"      - {item.name}/")
    
    # 9. Aufräumen
    print("\n🧹 Aufräumen...")
    build_dir = base_dir / "build"
    if build_dir.exists():
        shutil.rmtree(build_dir)
        print("   🗑️  build/ gelöscht")
    
    for spec_file in base_dir.glob("*.spec"):
        spec_file.unlink()
        print(f"   🗑️  {spec_file.name} gelöscht")
    
    print("\n" + "=" * 60)
    print("🎉 BUILD ERFOLGREICH ABGESCHLOSSEN!")
    print(f"📦 Ihre ausführbare Anwendung befindet sich in: {release_dir}")
    print("📧 Die .exe-Datei kann jetzt verteilt werden!")
    print("=" * 60)
    
    return True

if __name__ == "__main__":
    success = main()
    if not success:
        print("\n❌ Build fehlgeschlagen!")
        sys.exit(1)
    
    # Warten auf Eingabe
    input("\nDrücken Sie Enter zum Beenden...")
