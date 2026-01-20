"""
BoM zu EPD Konverter

Dieses Modul enthält die Hauptlogik zum Konvertieren von Bill of Materials (BoM) 
Excel-Dateien in EPD (Environmental Product Declaration) Format.
"""

import pandas as pd
import json
import requests
from pathlib import Path
import uuid
from typing import Optional, Callable, List, Dict, Any, Union


def read_materials_and_map(
    file_path: Union[str, Path],
    sheet_name: str,
    mapping_file: Union[str, Path],
    start_row_index: int = 0,
    material_column_index: int = 2,
    amount_column_index: int = 4
) -> pd.DataFrame:
    """
    Liest Materialien aus einer Excel-Datei und mappt sie mit Ecoinvent-Prozessen.
    
    Args:
        file_path: Pfad zur BoM Excel-Datei
        sheet_name: Name des Excel-Sheets
        mapping_file: Pfad zur Mapping-Datei (Materialien zu Ecoinvent-Prozessen)
        start_row_index: Zeilenindex, ab dem die Materialien beginnen (0 = erste Zeile)
        material_column_index: Spaltenindex für Materialnamen (0 = A, 1 = B, ...)
        amount_column_index: Spaltenindex für Mengen (0 = A, 1 = B, ...)
    
    Returns:
        DataFrame mit Materialien, Amounts, Units und UUIDs für A1 und A3
    """
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
    df = df_raw.iloc[start_row_index:, :].dropna(how='all')

    materials = []
    total_net = None
    final_prod = None

    for _, row in df.iterrows():
        mat = str(row[material_column_index]).strip()

        if not mat or mat.lower() == 'nan':
            continue

        amount_str = str(row[amount_column_index]).replace(",", ".") if pd.notna(row[amount_column_index]) else ""
        try:
            amount_val = float(amount_str)
        except ValueError:
            amount_val = 0.0

        if amount_val <= 0:
            continue

        # Spezialfall Verpackungsgewicht
        if mat.lower().startswith("total net weight material"):
            total_net = amount_val
            continue
        elif mat.lower().startswith("final product"):
            final_prod = amount_val
            continue

        # Normale Materialien
        materials.append({"Material": mat, "Amount": amount_val})

    # Verpackung berechnen
    if total_net is not None and final_prod is not None and final_prod > total_net:
        packaging_weight = final_prod - total_net
        share = packaging_weight / 3.0
        materials.extend([
            {"Material": "Packaging pallet", "Amount": share},
            {"Material": "Packaging carton", "Amount": share},
            {"Material": "Packaging film", "Amount": share}
        ])

    df_mat = pd.DataFrame(materials)

    # Mapping einlesen
    df_map = pd.read_excel(mapping_file, engine='openpyxl')
    df_map['Material_name_norm'] = df_map['Material_name'].str.strip().str.lower()
    df_mat['Material_norm'] = df_mat['Material'].str.strip().str.lower()

    # Merge mit Mapping
    df_merged = pd.merge(df_mat, df_map,
                         left_on="Material_norm",
                         right_on="Material_name_norm",
                         how="left")

    # Final Amount (A1)
    df_merged['Final_Unit_A1'] = df_merged['Process_unit_A1'].fillna('kg')
    df_merged['Final_Amount_A1'] = df_merged.apply(
        lambda r: r['Amount'] * (r['Conversion_factor_A1'] if pd.notna(r['Conversion_factor_A1']) else 1.0),
        axis=1
    )

    # Falls A3 vorhanden
    df_merged['Final_Unit_A3'] = df_merged['Process_unit_A3']
    df_merged['Final_Amount_A3'] = df_merged.apply(
        lambda r: r['Final_Amount_A1'] * (r['Conversion_factor_A3'] if pd.notna(r['Conversion_factor_A3']) else 1.0)
        if pd.notna(r['Process_uuid_A3']) else None,
        axis=1
    )

    return df_merged[['Material', 'Amount',
                      'Final_Amount_A1', 'Final_Unit_A1', 'Process_uuid_A1',
                      'Final_Amount_A3', 'Final_Unit_A3', 'Process_uuid_A3']]


def read_excel_like_reference(df: pd.DataFrame, root_repository: str) -> tuple[List[Dict], List[Dict]]:
    """
    Baut inputs und components für A1 und optional A3 aus dem DataFrame.
    
    Args:
        df: DataFrame mit Materialien und UUIDs
        root_repository: URL des Root-Repositories (Ecoinvent)
    
    Returns:
        Tuple von (inputs, components) für die API
    """
    components = []
    inputs = []

    for _, row in df.iterrows():
        # A1
        material = str(row["Material"]).strip()
        amount_a1 = float(str(row["Final_Amount_A1"]).replace(",", "."))
        unit_a1 = str(row["Final_Unit_A1"]).strip()
        epd_uuid_a1 = str(row["Process_uuid_A1"]).strip()

        local_id_a1 = str(uuid.uuid4())

        components.append({
            "id": local_id_a1,
            "name": material + " (A1)",
            "epd": epd_uuid_a1,
            "repository": root_repository
        })

        inputs.append({
            "component": local_id_a1,
            "amount": amount_a1,
            "unit": unit_a1
        })

        # A3 nur wenn vorhanden
        if pd.notna(row["Process_uuid_A3"]):
            amount_a3 = float(str(row["Final_Amount_A3"]).replace(",", ".")) if row["Final_Amount_A3"] is not None else 0.0
            unit_a3 = str(row["Final_Unit_A3"]).strip() if pd.notna(row["Final_Unit_A3"]) else unit_a1
            epd_uuid_a3 = str(row["Process_uuid_A3"]).strip()

            local_id_a3 = str(uuid.uuid4())

            components.append({
                "id": local_id_a3,
                "name": material + " (A3 process)",
                "epd": epd_uuid_a3,
                "repository": root_repository
            })

            inputs.append({
                "component": local_id_a3,
                "amount": amount_a3,
                "unit": unit_a3
            })

    return inputs, components


def generate_payload(
    full_name: str,
    inputs: List[Dict],
    components: List[Dict],
    epd_unit: str,
    target_repository: str,
    auth_list: List[Dict],
    method_lib: Dict
) -> Dict[str, Any]:
    """
    Generiert den JSON-Payload für die EPD-API.
    
    Args:
        full_name: Name des EPDs
        inputs: Liste der Input-Komponenten
        components: Liste der Komponenten
        epd_unit: Einheit des EPDs (z.B. "kg", "Item(s)")
        target_repository: URL des Ziel-Repositories
        auth_list: Liste der Authentifizierungsdaten
        method_lib: Dictionary mit Method Library Informationen
    
    Returns:
        Dictionary mit dem vollständigen Payload für die API
    """
    root_id = str(uuid.uuid4())

    root_component = {
        "id": root_id,
        "name": str(full_name).strip(),
        "inputs": inputs,
        "repository": target_repository
    }

    return {
        "auth": auth_list,
        "methodLib": method_lib,
        "root": {
            "component": root_id,
            "amount": 1,
            "unit": epd_unit,
            "modules": ["A1-A3"]
        },
        "components": components + [root_component]
    }


def save_json(data: Dict[str, Any], output_path: Path) -> None:
    """
    Speichert das JSON in einer Datei.
    
    Args:
        data: Dictionary mit den zu speichernden Daten
        output_path: Pfad zur Ausgabedatei
    """
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"JSON gespeichert unter: {output_path}")


def send_to_api(payload: Dict[str, Any], url_api: str, api_key: str) -> requests.Response:
    """
    Sendet den Payload an die EPD-API.
    
    Args:
        payload: Dictionary mit dem Payload
        url_api: URL der API
        api_key: API-Schlüssel
    
    Returns:
        Response-Objekt der API
    """
    headers = {
        "x-api-key": api_key,
        "content-type": "application/json"
    }
    print("Sende Payload an API...")
    resp = requests.post(url_api, json=payload, headers=headers)
    print(f"API Status-Code: {resp.status_code}")
    try:
        print("API Antwort:", resp.json())
    except json.JSONDecodeError:
        print("Antwort ist keine gültige JSON:", resp.text)
    return resp


def process_epd(
    main_file_path: Union[str, Path],
    sheet_name: str,
    mapping_file_path: Union[str, Path],
    full_name: str,
    epd_unit: str,
    root_repository: str,
    target_repository: str,
    start_row_index: int,
    material_column_index: int,
    amount_column_index: int,
    auth_list: List[Dict],
    method_lib: Dict,
    url_api: str,
    api_key: str,
    output_dir: Path,
    skip_missing_materials: bool = False,
    log_callback: Optional[Callable[[str], None]] = None
) -> Optional[requests.Response]:
    """
    Hauptfunktion zum Verarbeiten der EPD-Erstellung.
    
    Args:
        main_file_path: Pfad zur BoM Excel-Datei
        sheet_name: Name des Excel-Sheets
        mapping_file_path: Pfad zur Mapping-Datei
        full_name: Name des EPDs
        epd_unit: Einheit des EPDs
        root_repository: URL des Root-Repositories
        target_repository: URL des Ziel-Repositories
        start_row_index: Zeilenindex, ab dem Materialien beginnen
        material_column_index: Spaltenindex für Materialnamen
        amount_column_index: Spaltenindex für Mengen
        auth_list: Liste der Authentifizierungsdaten
        method_lib: Dictionary mit Method Library Informationen
        url_api: URL der API
        api_key: API-Schlüssel
        output_dir: Ausgabeverzeichnis
        skip_missing_materials: Wenn True, werden fehlende Materialien automatisch übersprungen
        log_callback: Optional callback-Funktion für Log-Nachrichten (z.B. für GUI)
    
    Returns:
        Response-Objekt der API oder None bei Abbruch
    """
    if log_callback:
        log_callback(f"Lese Materialien aus: {main_file_path}")

    df_for_payload = read_materials_and_map(
        main_file_path,
        sheet_name,
        mapping_file_path,
        start_row_index,
        material_column_index,
        amount_column_index
    )

    # Fehlende Materialien prüfen
    missing_a1 = df_for_payload[df_for_payload['Process_uuid_A1'].isna()]
    if not missing_a1.empty:
        missing_list = []
        for _, row in missing_a1.iterrows():
            missing_list.append(f"  - {row['Material']} ({row['Amount']})")

        warning_msg = "WARNUNG: Keine A1-UUID gefunden für:\n" + "\n".join(missing_list)

        if log_callback:
            log_callback(warning_msg)
        else:
            print(warning_msg)

        if not skip_missing_materials:
            # Bei GUI: automatisch überspringen, wenn nicht explizit erlaubt
            if log_callback:
                log_callback("Fehlende Materialien werden übersprungen.")
            else:
                # Bei CLI: automatisch überspringen (keine interaktive Eingabe mehr)
                print("Fehlende Materialien werden übersprungen.")

        df_for_payload = df_for_payload[df_for_payload['Process_uuid_A1'].notna()]

    inputs, components = read_excel_like_reference(df_for_payload, root_repository)
    payload = generate_payload(full_name, inputs, components, epd_unit, target_repository, auth_list, method_lib)
    output_path = output_dir / f"{full_name}.json"
    save_json(payload, output_path)
    resp = send_to_api(payload, url_api, api_key)
    return resp
