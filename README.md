# BoM zu EPD Konverter

Tool zum Konvertieren von Bill of Materials (BoM) Excel-Dateien in EPD (Environmental Product Declaration) Format.

## Installation

1. Repository klonen:
```bash
git clone <repository-url>
cd BoM_to_EPD
```

2. Python-Abhängigkeiten installieren:
```bash
pip install -r requirements.txt
```

## Verwendung

### GUI starten

```bash
python BoM_to_EPD/bom_to_epd_gui.py
```

### Workflow

1. **Dateien & Excel-Einstellungen**
   - BoM Excel-Datei auswählen
   - Mapping-Datei auswählen (Standard: `Mapping_Materials_to_Processes.xlsx`)
   - Sheet-Name eingeben
   - Start-Zeilenindex festlegen (0 = Zeile 1, 1 = Zeile 2, ...)
   - Material-Spaltenindex festlegen (0 = A, 1 = B, 2 = C, ...)
   - Amount-Spaltenindex festlegen (0 = A, 1 = B, 2 = C, ...)
   - **Materialien laden & Vorschau anzeigen** - Prüft die gefundenen Materialien

2. **Authentifizierung**
   - Prod (lca.ditwin.cloud) - URL, User, Password eingeben
   - Dev (lca.dev.ditwin.cloud) - URL, User, Password eingeben

3. **Repositories**
   - Root Repository (Ecoinvent) - fest voreingestellt, nicht änderbar
   - Target Repository - Ziel-Repository für das erstellte EPD

4. **EPD-Einstellungen**
   - EPD-Name eingeben
   - EPD-Einheit aus Dropdown wählen (kg, Item(s), m, m², etc.)

5. **API & Method Library**
   - API URL und API Key eingeben
   - Method URL und Method Name eingeben

6. **Ausgabe & Start**
   - Output-Verzeichnis auswählen
   - Option: "Fehlende Materialien automatisch überspringen"
   - **EPD erstellen** Button klicken

## Dateien

- `bom_to_epd_gui.py` - GUI-Anwendung
- `bom_to_epd.py` - Hauptlogik und API-Kommunikation
- `Mapping_Materials_to_Processes.xlsx` - Mapping-Datei für Materialien zu Ecoinvent-Prozessen

## Abhängigkeiten

- Python 3.7+
- pandas
- requests
- openpyxl (für Excel-Dateien)
- tkinter (kommt standardmäßig mit Python)

## Hinweise

- Das `results/` Verzeichnis wird automatisch erstellt
- Excel-Dateien (außer Mapping-Datei) werden nicht ins Repository aufgenommen
- Die Materialien-Vorschau zeigt fehlende A1-UUIDs rot markiert an
