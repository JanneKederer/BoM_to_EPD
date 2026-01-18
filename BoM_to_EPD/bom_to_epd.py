import pandas as pd
import json
import requests
from pathlib import Path
import uuid

# ---------------------------------------------------------------------------------------------------

auth_list = [
    {
        "url": "https://lca.ditwin.cloud",
        "user": "janne_teresa_kederer_siemens_energy_com",
        "password": "Test4EPDtree!"
    },
    {
        "url": "https://lca.dev.ditwin.cloud",
        "user": "janne_teresa_kederer_siemens_energy_com",
        "password": "Test4EPDtree!"
    }
]

method_lib = {
    "url": "https://lca.ditwin.cloud",  # oder "https://lca.dev.ditwin.cloud" auch möglich?
    "name": "en15804_pef31_indata_lcia_method"
}

# Name für EPD
full_name = "Ion Exchanger - Purolite MB400 - 1kg"
epd_unit = "kg"  # Unit z.B. "Item(s)"; "kg"

# Repositories
root_repository = "https://lca.dev.ditwin.cloud/Playground/Ecoinvent_3_10_EN15804_results2"
target_repository = "https://lca.dev.ditwin.cloud/Computed/HVDC_Repo"

# BoM
main_file_path = r"C:\Users\z004ud7a\OneDrive - Siemens Energy\Documents\LCA Aufgaben\HVDC\CWC\CO2 Ion Exchanger.xlsx"
sheet_name = "4.1 Carbon Footprint Tool"
start_row_index = 5  # Zeilenindex: wo steht "Total net weight material" (0 = Zeile 1, 1 = Zeile 2 ...)
amount_column_index = 4  # Spalte in der die Amounts stehen (0 = A, 1 = B, 2 = C, 3 = D, 4 = E ...) Standard ist Spalte E = Index 4

# Mapping Datei: Materialnamen mit ecoinvent-Prozessen
mapping_file_path = Path(__file__).parent / "Mapping_Materials_to_Processes.xlsx"

output_dir = Path(r"C:\Users\z004ud7a\OneDrive - Siemens Energy\Documents\LCA Aufgaben\Modellierung\BoM_to_EPD\results")
output_dir.mkdir(parents=True, exist_ok=True)

# API
url_api = "https://olca-epd.dev.ditwin.cloud/run-epd-tree"
api_key = "develop"
headers = {
    "x-api-key": api_key,
    "content-type": "application/json"
}

# ---------------------------------------------------------------------------------------------------


def read_materials_and_map(file_path, sheet_name, mapping_file, start_row_index=5, material_column_index=2, amount_column_index=4):
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
    df = df_raw.iloc[start_row_index:, :].dropna(how='all')

    materials = []
    total_net = None
    final_prod = None

    for _, row in df.iterrows():
        mat = str(row[material_column_index]).strip()  # Materialname aus konfigurierbarer Spalte

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

    # Fehlende Materialien werden in process_epd behandelt

    # Final Amount (A1)
    df_merged['Final_Unit_A1'] = df_merged['Process_unit_A1'].fillna('kg')
    df_merged['Final_Amount_A1'] = df_merged.apply(lambda r: r['Amount'] * (r['Conversion_factor_A1'] if pd.notna(r['Conversion_factor_A1']) else 1.0), axis=1)

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


def read_excel_like_reference(df, root_repository):
    """Baut inputs und components für A1 und optional A3"""
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


def generate_payload(full_name, inputs, components, epd_unit, target_repository, auth_list, method_lib):
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


def save_json(data, output_path):
    """Speichert das JSON in einer Datei."""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"JSON gespeichert unter: {output_path}")


def send_to_api(payload, url_api, api_key):
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


def process_epd(main_file_path, sheet_name, mapping_file_path, full_name, epd_unit,
                root_repository, target_repository, start_row_index, material_column_index, amount_column_index,
                auth_list, method_lib, url_api, api_key, output_dir, skip_missing_materials=False,
                log_callback=None):
    """
    Hauptfunktion zum Verarbeiten der EPD-Erstellung.

    Args:
        skip_missing_materials: Wenn True, werden fehlende Materialien automatisch übersprungen
        log_callback: Optional callback-Funktion für Log-Nachrichten (z.B. für GUI)
    """
    if log_callback:
        log_callback(f"Lese Materialien aus: {main_file_path}")

    df_for_payload = read_materials_and_map(main_file_path, sheet_name, mapping_file_path,
                                            start_row_index, material_column_index, amount_column_index)

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
            # Interaktive Bestätigung (nur wenn nicht über GUI)
            if log_callback:
                # Bei GUI: automatisch überspringen, wenn nicht explizit erlaubt
                log_callback("Fehlende Materialien werden übersprungen.")
            else:
                while True:
                    confirm = input("Fortfahren mit diesen fehlenden Materialien? (j/n): ").strip().lower()
                    if confirm in ['ja', 'j', 'y']:
                        break
                    elif confirm in ['nein', 'n']:
                        if log_callback:
                            log_callback("Vorgang abgebrochen.")
                        else:
                            print("Vorgang abgebrochen.")
                        return None
                    else:
                        print("Bitte 'ja' oder 'nein' eingeben.")

        df_for_payload = df_for_payload[df_for_payload['Process_uuid_A1'].notna()]

    inputs, components = read_excel_like_reference(df_for_payload, root_repository)
    payload = generate_payload(full_name, inputs, components, epd_unit, target_repository, auth_list, method_lib)
    output_path = output_dir / f"{full_name}.json"
    save_json(payload, output_path)
    resp = send_to_api(payload, url_api, api_key)
    return resp


def main():
    df_for_payload = read_materials_and_map(main_file_path, sheet_name, mapping_file_path, start_row_index, 2, amount_column_index)  # material_column_index=2 (Spalte C)
    inputs, components = read_excel_like_reference(df_for_payload, root_repository)
    payload = generate_payload(full_name, inputs, components, epd_unit, target_repository, auth_list, method_lib)
    output_path = output_dir / f"{full_name}.json"
    save_json(payload, output_path)
    # print(json.dumps(payload, indent=2, ensure_ascii=False))
    send_to_api(payload, url_api, api_key)


if __name__ == "__main__":
    main()
