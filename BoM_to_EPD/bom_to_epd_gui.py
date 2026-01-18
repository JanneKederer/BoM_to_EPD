import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import threading
from bom_to_epd import process_epd, read_materials_and_map
import pandas as pd

class BoMToEPDGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("BoM zu EPD Konverter")
        self.root.geometry("800x900")
        
        # Variablen
        self.main_file_path = tk.StringVar()
        self.mapping_file_path = tk.StringVar(value=str(Path(__file__).parent / "Mapping_Materials_to_Processes.xlsx"))
        self.sheet_name = tk.StringVar(value="4.1 Carbon Footprint Tool")
        self.full_name = tk.StringVar(value="Ion Exchanger - Purolite MB400 - 1kg")
        self.epd_unit = tk.StringVar(value="kg")
        self.epd_unit_options = ["kg", "Item(s)", "m", "m²", "m³", "t", "g", "l", "ml"]
        self.start_row_index = tk.IntVar(value=5)
        self.material_column_index = tk.IntVar(value=2)  # Spalte C = Index 2
        self.amount_column_index = tk.IntVar(value=4)
        self.root_repository = tk.StringVar(value="https://lca.dev.ditwin.cloud/Playground/Ecoinvent_3_10_EN15804_results2")
        self.target_repository = tk.StringVar(value="https://lca.dev.ditwin.cloud/Computed/HVDC_Repo")
        # API & Method Library - fest voreingestellt
        self.url_api = "https://olca-epd.dev.ditwin.cloud/run-epd-tree"
        self.api_key = "develop"
        self.output_dir = tk.StringVar(value=str(Path(__file__).parent / "results"))
        self.skip_missing = tk.BooleanVar(value=False)
        
        # Auth-Einstellungen
        self.auth_url1 = tk.StringVar(value="https://lca.ditwin.cloud")
        self.auth_user1 = tk.StringVar(value="janne_teresa_kederer_siemens_energy_com")
        self.auth_password1 = tk.StringVar(value="Test4EPDtree!")
        self.auth_url2 = tk.StringVar(value="https://lca.dev.ditwin.cloud")
        self.auth_user2 = tk.StringVar(value="janne_teresa_kederer_siemens_energy_com")
        self.auth_password2 = tk.StringVar(value="Test4EPDtree!")
        
        # Method Lib - fest voreingestellt
        self.method_url = "https://lca.ditwin.cloud"
        self.method_name = "en15804_pef31_indata_lcia_method"
        
        self.create_widgets()
        
    def create_widgets(self):
        # Hauptframe mit Scrollbar
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas für Scrollbar
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        current_row = 0
        
        # ===== 1. DATEIEN & EXCEL-EINSTELLUNGEN =====
        ttk.Label(scrollable_frame, text="1. Dateien & Excel-Einstellungen", font=("Arial", 12, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", pady=(0, 5))
        current_row += 1
        
        ttk.Label(scrollable_frame, text="BoM Excel-Datei:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.main_file_path, width=50).grid(row=current_row, column=1, padx=5, pady=2)
        ttk.Button(scrollable_frame, text="Durchsuchen", command=self.browse_main_file).grid(row=current_row, column=2, padx=5, pady=2)
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Mapping-Datei:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.mapping_file_path, width=50).grid(row=current_row, column=1, padx=5, pady=2)
        ttk.Button(scrollable_frame, text="Durchsuchen", command=self.browse_mapping_file).grid(row=current_row, column=2, padx=5, pady=2)
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Sheet-Name:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.sheet_name, width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Start-Zeilenindex:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Spinbox(scrollable_frame, from_=0, to=100, textvariable=self.start_row_index, width=20).grid(row=current_row, column=1, padx=5, pady=2, sticky="w")
        ttk.Label(scrollable_frame, text="(0 = Zeile 1, 1 = Zeile 2, ...)", font=("Arial", 8)).grid(row=current_row, column=2, sticky="w", padx=5)
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Material-Spaltenindex:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Spinbox(scrollable_frame, from_=0, to=20, textvariable=self.material_column_index, width=20).grid(row=current_row, column=1, padx=5, pady=2, sticky="w")
        ttk.Label(scrollable_frame, text="(0 = A, 1 = B, 2 = C, ...)", font=("Arial", 8)).grid(row=current_row, column=2, sticky="w", padx=5)
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Amount-Spaltenindex:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Spinbox(scrollable_frame, from_=0, to=20, textvariable=self.amount_column_index, width=20).grid(row=current_row, column=1, padx=5, pady=2, sticky="w")
        ttk.Label(scrollable_frame, text="(0 = A, 1 = B, 2 = C, ...)", font=("Arial", 8)).grid(row=current_row, column=2, sticky="w", padx=5)
        current_row += 1
        
        # Materialien-Vorschau Button
        preview_frame = ttk.Frame(scrollable_frame)
        preview_frame.grid(row=current_row, column=0, columnspan=3, pady=10)
        ttk.Button(preview_frame, text="📋 Materialien laden & Vorschau anzeigen", command=self.preview_materials, width=40).pack()
        current_row += 1
        
        # ===== 2. AUTHENTIFIZIERUNG =====
        ttk.Label(scrollable_frame, text="2. Authentifizierung", font=("Arial", 12, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", pady=(15, 5))
        current_row += 1
        
        # Prod (lca.ditwin.cloud)
        ttk.Label(scrollable_frame, text="Prod (lca.ditwin.cloud) - URL:", font=("Arial", 9, "bold")).grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.auth_url1, width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Prod - User:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.auth_user1, width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Prod - Password:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.auth_password1, show="*", width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        # Dev (lca.dev.ditwin.cloud)
        ttk.Label(scrollable_frame, text="Dev (lca.dev.ditwin.cloud) - URL:", font=("Arial", 9, "bold")).grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.auth_url2, width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Dev - User:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.auth_user2, width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Dev - Password:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.auth_password2, show="*", width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        # ===== 3. REPOSITORIES =====
        ttk.Label(scrollable_frame, text="3. Repositories", font=("Arial", 12, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", pady=(15, 5))
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Root Repository (Ecoinvent):").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        root_repo_entry = ttk.Entry(scrollable_frame, textvariable=self.root_repository, width=50, state="readonly")
        root_repo_entry.grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Target Repository:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.target_repository, width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        # ===== 4. EPD-EINSTELLUNGEN =====
        ttk.Label(scrollable_frame, text="4. EPD-Einstellungen", font=("Arial", 12, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", pady=(15, 5))
        current_row += 1
        
        ttk.Label(scrollable_frame, text="EPD-Name:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.full_name, width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        ttk.Label(scrollable_frame, text="EPD-Einheit:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        epd_unit_combo = ttk.Combobox(scrollable_frame, textvariable=self.epd_unit, values=self.epd_unit_options, width=20, state="readonly")
        epd_unit_combo.grid(row=current_row, column=1, padx=5, pady=2, sticky="w")
        current_row += 1
        
        # ===== 5. AUSGABE & START =====
        ttk.Label(scrollable_frame, text="6. Ausgabe & Start", font=("Arial", 12, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", pady=(15, 5))
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Output-Verzeichnis:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.output_dir, width=50).grid(row=current_row, column=1, padx=5, pady=2)
        ttk.Button(scrollable_frame, text="Durchsuchen", command=self.browse_output_dir).grid(row=current_row, column=2, padx=5, pady=2)
        current_row += 1
        
        ttk.Checkbutton(scrollable_frame, text="Fehlende Materialien automatisch überspringen", 
                       variable=self.skip_missing).grid(row=current_row, column=0, columnspan=3, sticky="w", padx=5, pady=5)
        current_row += 1
        
        # Buttons
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=current_row, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="EPD erstellen", command=self.process_epd, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Abbrechen", command=self.root.quit, width=20).pack(side=tk.LEFT, padx=5)
        current_row += 1
        
        # Log-Ausgabe
        ttk.Label(scrollable_frame, text="Log-Ausgabe", font=("Arial", 12, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", pady=(10, 5))
        current_row += 1
        
        self.log_text = scrolledtext.ScrolledText(scrollable_frame, height=10, width=80)
        self.log_text.grid(row=current_row, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        
        # Canvas und Scrollbar packen
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Grid-Konfiguration
        scrollable_frame.columnconfigure(1, weight=1)
        
    def browse_main_file(self):
        filename = filedialog.askopenfilename(
            title="BoM Excel-Datei auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.main_file_path.set(filename)
            self.log(f"BoM-Datei ausgewählt: {filename}")
    
    def browse_mapping_file(self):
        filename = filedialog.askopenfilename(
            title="Mapping-Datei auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.mapping_file_path.set(filename)
            self.log(f"Mapping-Datei ausgewählt: {filename}")
    
    def browse_output_dir(self):
        dirname = filedialog.askdirectory(title="Output-Verzeichnis auswählen")
        if dirname:
            self.output_dir.set(dirname)
            self.log(f"Output-Verzeichnis ausgewählt: {dirname}")
    
    def preview_materials(self):
        """Lädt Materialien aus Excel und zeigt sie in einem Vorschau-Fenster an"""
        # Validierung
        if not self.main_file_path.get():
            messagebox.showerror("Fehler", "Bitte wählen Sie eine BoM Excel-Datei aus.")
            return
        if not Path(self.main_file_path.get()).exists():
            messagebox.showerror("Fehler", "Die ausgewählte BoM-Datei existiert nicht.")
            return
        if not self.mapping_file_path.get():
            messagebox.showerror("Fehler", "Bitte wählen Sie eine Mapping-Datei aus.")
            return
        if not Path(self.mapping_file_path.get()).exists():
            messagebox.showerror("Fehler", "Die ausgewählte Mapping-Datei existiert nicht.")
            return
        
        try:
            self.log("Lade Materialien...")
            # Materialien lesen (ohne API-Calls)
            df_merged = read_materials_and_map(
                self.main_file_path.get(),
                self.sheet_name.get(),
                self.mapping_file_path.get(),
                self.start_row_index.get(),
                self.material_column_index.get(),
                self.amount_column_index.get()
            )
            
            # Vorschau-Fenster erstellen
            preview_window = tk.Toplevel(self.root)
            preview_window.title("Materialien-Vorschau")
            preview_window.geometry("1000x600")
            
            # Frame für die Tabelle
            frame = ttk.Frame(preview_window, padding="10")
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Überschrift
            info_label = ttk.Label(frame, text=f"Gefundene Materialien: {len(df_merged)}", font=("Arial", 10, "bold"))
            info_label.pack(pady=(0, 10))
            
            # Treeview für die Tabelle
            tree = ttk.Treeview(frame, columns=("Material", "Amount", "Unit_A1", "UUID_A1", "Unit_A3", "UUID_A3"), show="headings", height=20)
            
            # Spalten definieren
            tree.heading("Material", text="Material")
            tree.heading("Amount", text="Amount (A1)")
            tree.heading("Unit_A1", text="Einheit (A1)")
            tree.heading("UUID_A1", text="UUID (A1)")
            tree.heading("Unit_A3", text="Einheit (A3)")
            tree.heading("UUID_A3", text="UUID (A3)")
            
            tree.column("Material", width=250)
            tree.column("Amount", width=100)
            tree.column("Unit_A1", width=100)
            tree.column("UUID_A1", width=200)
            tree.column("Unit_A3", width=100)
            tree.column("UUID_A3", width=200)
            
            # Scrollbar
            scrollbar_tree = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar_tree.set)
            
            # Daten einfügen
            missing_count = 0
            for _, row in df_merged.iterrows():
                material = str(row['Material'])
                amount = f"{row['Final_Amount_A1']:.4f}" if pd.notna(row['Final_Amount_A1']) else "-"
                unit_a1 = str(row['Final_Unit_A1']) if pd.notna(row['Final_Unit_A1']) else "-"
                uuid_a1 = str(row['Process_uuid_A1']) if pd.notna(row['Process_uuid_A1']) else "❌ FEHLT"
                unit_a3 = str(row['Final_Unit_A3']) if pd.notna(row['Final_Unit_A3']) else "-"
                uuid_a3 = str(row['Process_uuid_A3']) if pd.notna(row['Process_uuid_A3']) else "-"
                
                if pd.isna(row['Process_uuid_A1']):
                    missing_count += 1
                    tag = "missing"
                else:
                    tag = ""
                
                tree.insert("", tk.END, values=(material, amount, unit_a1, uuid_a1, unit_a3, uuid_a3), tags=(tag,))
            
            # Fehlende Materialien rot markieren
            tree.tag_configure("missing", background="#ffcccc")
            
            # Treeview und Scrollbar packen
            tree.pack(side="left", fill=tk.BOTH, expand=True)
            scrollbar_tree.pack(side="right", fill="y")
            
            # Zusammenfassung
            summary_text = f"Gesamt: {len(df_merged)} Materialien"
            if missing_count > 0:
                summary_text += f" | ⚠️ {missing_count} ohne A1-UUID"
            summary_label = ttk.Label(frame, text=summary_text, font=("Arial", 9))
            summary_label.pack(pady=(10, 0))
            
            # Schließen-Button
            close_button = ttk.Button(frame, text="Schließen", command=preview_window.destroy)
            close_button.pack(pady=10)
            
            self.log(f"Materialien-Vorschau angezeigt: {len(df_merged)} Materialien gefunden")
            if missing_count > 0:
                self.log(f"⚠️ Warnung: {missing_count} Materialien haben keine A1-UUID")
            
        except Exception as e:
            error_msg = f"Fehler beim Laden der Materialien: {str(e)}"
            messagebox.showerror("Fehler", error_msg)
            self.log(error_msg)
            import traceback
            self.log(traceback.format_exc())
    
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def validate_inputs(self):
        if not self.main_file_path.get():
            messagebox.showerror("Fehler", "Bitte wählen Sie eine BoM Excel-Datei aus.")
            return False
        if not Path(self.main_file_path.get()).exists():
            messagebox.showerror("Fehler", "Die ausgewählte BoM-Datei existiert nicht.")
            return False
        if not self.mapping_file_path.get():
            messagebox.showerror("Fehler", "Bitte wählen Sie eine Mapping-Datei aus.")
            return False
        if not Path(self.mapping_file_path.get()).exists():
            messagebox.showerror("Fehler", "Die ausgewählte Mapping-Datei existiert nicht.")
            return False
        if not self.full_name.get():
            messagebox.showerror("Fehler", "Bitte geben Sie einen EPD-Namen ein.")
            return False
        return True
    
    def process_epd(self):
        if not self.validate_inputs():
            return
        
        # UI deaktivieren während der Verarbeitung
        self.log_text.delete(1.0, tk.END)
        self.log("Starte EPD-Erstellung...")
        
        # In separatem Thread ausführen, damit UI nicht einfriert
        thread = threading.Thread(target=self._process_epd_thread)
        thread.daemon = True
        thread.start()
    
    def _process_epd_thread(self):
        try:
            # Auth-Liste erstellen
            auth_list = [
                {
                    "url": self.auth_url1.get(),
                    "user": self.auth_user1.get(),
                    "password": self.auth_password1.get()
                },
                {
                    "url": self.auth_url2.get(),
                    "user": self.auth_user2.get(),
                    "password": self.auth_password2.get()
                }
            ]
            
            # Method Lib erstellen (fest voreingestellt)
            method_lib = {
                "url": self.method_url,
                "name": self.method_name
            }
            
            # Output-Verzeichnis erstellen
            output_dir = Path(self.output_dir.get())
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # EPD verarbeiten
            self.root.after(0, self.log, f"Verarbeite Datei: {self.main_file_path.get()}")
            self.root.after(0, self.log, f"Sheet: {self.sheet_name.get()}")
            
            resp = process_epd(
                main_file_path=self.main_file_path.get(),
                sheet_name=self.sheet_name.get(),
                mapping_file_path=self.mapping_file_path.get(),
                full_name=self.full_name.get(),
                epd_unit=self.epd_unit.get(),
                root_repository=self.root_repository.get(),
                target_repository=self.target_repository.get(),
                start_row_index=self.start_row_index.get(),
                material_column_index=self.material_column_index.get(),
                amount_column_index=self.amount_column_index.get(),
                auth_list=auth_list,
                method_lib=method_lib,
                url_api=self.url_api,
                api_key=self.api_key,
                output_dir=output_dir,
                skip_missing_materials=self.skip_missing.get(),
                log_callback=lambda msg: self.root.after(0, self.log, msg)
            )
            
            if resp is not None:
                self.root.after(0, self.log, "EPD erfolgreich erstellt!")
                self.root.after(0, self.log, f"Status-Code: {resp.status_code}")
                try:
                    response_json = resp.json()
                    self.root.after(0, self.log, f"API Antwort: {response_json}")
                except:
                    self.root.after(0, self.log, f"API Antwort: {resp.text}")
                self.root.after(0, messagebox.showinfo, "Erfolg", "EPD wurde erfolgreich erstellt!")
            else:
                self.root.after(0, self.log, "Vorgang abgebrochen.")
                self.root.after(0, messagebox.showwarning, "Abgebrochen", "Der Vorgang wurde abgebrochen.")
                
        except Exception as e:
            error_msg = f"Fehler: {str(e)}"
            self.root.after(0, self.log, error_msg)
            self.root.after(0, messagebox.showerror, "Fehler", error_msg)
            import traceback
            self.root.after(0, self.log, traceback.format_exc())

if __name__ == "__main__":
    root = tk.Tk()
    app = BoMToEPDGUI(root)
    root.mainloop()

