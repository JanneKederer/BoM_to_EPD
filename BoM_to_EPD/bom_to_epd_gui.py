import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
from bom_to_epd import process_epd, read_materials_and_map
import pandas as pd

class BoMToEPDGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("BoM zu EPD Konverter")
        self.root.geometry("800x700")
        
        # Variablen
        self.main_file_path = tk.StringVar()
        self.mapping_file_path = tk.StringVar(value=str(Path(__file__).parent / "Mapping_Materials_to_Processes.xlsx"))
        self.sheet_name = tk.StringVar()
        self.full_name = tk.StringVar()
        self.epd_unit = tk.StringVar(value="kg")
        self.epd_unit_options = ["kg", "Item(s)", "m", "m²", "m³", "t", "g", "l", "ml"]
        self.material_column = tk.StringVar()
        self.amount_column = tk.StringVar()
        self.root_repository = tk.StringVar(value="https://lca.dev.ditwin.cloud/Playground/Ecoinvent_3_10_EN15804_results2")
        self.target_repository = tk.StringVar(value="https://lca.dev.ditwin.cloud/Computed/HVDC_Repo")
        # API & Method Library - fest voreingestellt
        self.url_api = "https://olca-epd.dev.ditwin.cloud/run-epd-tree"
        self.api_key = "develop"
        self.output_dir = tk.StringVar(value=str(Path(__file__).parent / "results"))
        self.skip_missing = tk.BooleanVar(value=False)
        
        # Auth-Einstellungen (URLs sind fest voreingestellt)
        self.auth_url1 = "https://lca.ditwin.cloud"
        self.auth_url2 = "https://lca.dev.ditwin.cloud"
        self.auth_user1 = tk.StringVar(value="janne_teresa_kederer_siemens_energy_com")
        self.auth_password1 = tk.StringVar(value="Test4EPDtree!")
        self.auth_user2 = tk.StringVar(value="janne_teresa_kederer_siemens_energy_com")
        self.auth_password2 = tk.StringVar(value="Test4EPDtree!")
        
        # Method Lib - fest voreingestellt
        self.method_url = "https://lca.ditwin.cloud"
        self.method_name = "en15804_pef31_indata_lcia_method"
        
        self.create_widgets()
        
    def create_widgets(self):
        # Hauptframe mit Scrollbar
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
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
        
        # ===== DATEIEN & EXCEL-EINSTELLUNGEN =====
        ttk.Label(scrollable_frame, text="Dateien & Excel-Einstellungen", font=("Arial", 12, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", pady=(0, 5))
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
        
        ttk.Label(scrollable_frame, text="Material-Spalte:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.material_column, width=5).grid(row=current_row, column=1, padx=5, pady=2, sticky="w")
        ttk.Label(scrollable_frame, text="(z.B. A, B, C, ...)", font=("Arial", 8)).grid(row=current_row, column=2, sticky="w", padx=5)
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Amount-Spalte:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.amount_column, width=5).grid(row=current_row, column=1, padx=5, pady=2, sticky="w")
        ttk.Label(scrollable_frame, text="(z.B. A, B, C, ...)", font=("Arial", 8)).grid(row=current_row, column=2, sticky="w", padx=5)
        current_row += 1
        
        # Materialien-Vorschau Button
        preview_frame = ttk.Frame(scrollable_frame)
        preview_frame.grid(row=current_row, column=0, columnspan=3, pady=10)
        ttk.Button(preview_frame, text="📋 Materialien laden & Vorschau anzeigen", command=self.preview_materials, width=40).pack()
        current_row += 1
        
        # ===== EPD-EINSTELLUNGEN =====
        ttk.Label(scrollable_frame, text="EPD-Einstellungen", font=("Arial", 12, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", pady=(15, 5))
        current_row += 1
        
        ttk.Label(scrollable_frame, text="EPD-Name:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.full_name, width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        ttk.Label(scrollable_frame, text="EPD-Einheit:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        epd_unit_combo = ttk.Combobox(scrollable_frame, textvariable=self.epd_unit, values=self.epd_unit_options, width=20, state="readonly")
        epd_unit_combo.grid(row=current_row, column=1, padx=5, pady=2, sticky="w")
        current_row += 1
        
        # ===== AUTHENTIFIZIERUNG =====
        ttk.Label(scrollable_frame, text="Authentifizierung", font=("Arial", 12, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", pady=(15, 5))
        current_row += 1
        
        # Prod (lca.ditwin.cloud)
        ttk.Label(scrollable_frame, text="Prod (lca.ditwin.cloud)", font=("Arial", 10, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", padx=5, pady=(5, 2))
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Prod - User:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.auth_user1, width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Prod - Password:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        password_frame1 = ttk.Frame(scrollable_frame)
        password_frame1.grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        self.auth_password1_entry = ttk.Entry(password_frame1, textvariable=self.auth_password1, show="*", width=45)
        self.auth_password1_entry.pack(side="left")
        self.show_password1 = tk.BooleanVar(value=False)
        ttk.Checkbutton(password_frame1, text="Anzeigen", variable=self.show_password1, command=lambda: self.toggle_password(self.auth_password1_entry, self.show_password1)).pack(side="left", padx=(5, 0))
        current_row += 1
        
        # Dev (lca.dev.ditwin.cloud)
        ttk.Label(scrollable_frame, text="Dev (lca.dev.ditwin.cloud)", font=("Arial", 10, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", padx=5, pady=(10, 2))
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Dev - User:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.auth_user2, width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Dev - Password:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        password_frame2 = ttk.Frame(scrollable_frame)
        password_frame2.grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        self.auth_password2_entry = ttk.Entry(password_frame2, textvariable=self.auth_password2, show="*", width=45)
        self.auth_password2_entry.pack(side="left")
        self.show_password2 = tk.BooleanVar(value=False)
        ttk.Checkbutton(password_frame2, text="Anzeigen", variable=self.show_password2, command=lambda: self.toggle_password(self.auth_password2_entry, self.show_password2)).pack(side="left", padx=(5, 0))
        current_row += 1
        
        # ===== REPOSITORIES =====
        ttk.Label(scrollable_frame, text="Repositories", font=("Arial", 12, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", pady=(15, 5))
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Root Repository (Ecoinvent):").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        root_repo_entry = ttk.Entry(scrollable_frame, textvariable=self.root_repository, width=50, state="readonly")
        root_repo_entry.grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        ttk.Label(scrollable_frame, text="Target Repository:").grid(row=current_row, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(scrollable_frame, textvariable=self.target_repository, width=50).grid(row=current_row, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        current_row += 1
        
        # ===== AUSGABE & START =====
        ttk.Label(scrollable_frame, text="Ausgabe & Start", font=("Arial", 12, "bold")).grid(row=current_row, column=0, columnspan=3, sticky="w", pady=(15, 5))
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
    
    def browse_mapping_file(self):
        filename = filedialog.askopenfilename(
            title="Mapping-Datei auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.mapping_file_path.set(filename)
    
    def browse_output_dir(self):
        dirname = filedialog.askdirectory(title="Output-Verzeichnis auswählen")
        if dirname:
            self.output_dir.set(dirname)
    
    def column_letter_to_index(self, letter):
        """Konvertiert Spaltenbuchstaben (A, B, C, ..., Z, AA, AB, ...) in Index (0, 1, 2, ...)"""
        if not letter:
            return 0
        letter = letter.upper().strip()
        index = 0
        for char in letter:
            if not char.isalpha():
                return 0
            index = index * 26 + (ord(char) - ord('A') + 1)
        return index - 1
    
    def toggle_password(self, entry, var):
        """Schaltet die Sichtbarkeit des Passworts um"""
        if var.get():
            entry.config(show="")
        else:
            entry.config(show="*")
    
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
            # Materialien lesen (ohne API-Calls)
            df_merged = read_materials_and_map(
                self.main_file_path.get(),
                self.sheet_name.get(),
                self.mapping_file_path.get(),
                0,  # Start immer bei Zeile 0
                self.column_letter_to_index(self.material_column.get()),
                self.column_letter_to_index(self.amount_column.get())
            )
            
            # Materialien in gefundene und fehlende trennen
            df_found = df_merged[df_merged['Process_uuid_A1'].notna()].copy()
            df_missing = df_merged[df_merged['Process_uuid_A1'].isna()].copy()
            
            # Vorschau-Fenster erstellen
            preview_window = tk.Toplevel(self.root)
            preview_window.title("Materialien-Vorschau")
            preview_window.geometry("900x600")
            
            # Hauptframe
            main_frame = ttk.Frame(preview_window, padding="10")
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
            
            # ===== FEHLENDE MATERIALIEN (ZUERST) =====
            if len(df_missing) > 0:
                missing_label = ttk.Label(scrollable_frame, text=f"Fehlende Materialien ({len(df_missing)}) - Bitte prüfen!", font=("Arial", 11, "bold"), foreground="red")
                missing_label.pack(pady=(0, 5), anchor="w")
                
                # Frame für fehlende Materialien
                missing_frame = ttk.Frame(scrollable_frame)
                missing_frame.pack(fill=tk.X, pady=(0, 15))
                
                # Treeview für fehlende Materialien
                tree_missing = ttk.Treeview(missing_frame, columns=("Material", "Amount"), show="headings", height=8)
                tree_missing.heading("Material", text="Material")
                tree_missing.heading("Amount", text="Amount")
                
                tree_missing.column("Material", width=650, minwidth=400)
                tree_missing.column("Amount", width=150, minwidth=100)
                
                # Scrollbar für fehlende
                scrollbar_missing = ttk.Scrollbar(missing_frame, orient="vertical", command=tree_missing.yview)
                tree_missing.configure(yscrollcommand=scrollbar_missing.set)
                
                # Daten einfügen
                for _, row in df_missing.iterrows():
                    material = str(row['Material'])
                    # Verwende das ursprüngliche Amount
                    if 'Amount' in row and pd.notna(row['Amount']):
                        amount = f"{row['Amount']:.4f}"
                    elif 'Final_Amount_A1' in row and pd.notna(row['Final_Amount_A1']):
                        amount = f"{row['Final_Amount_A1']:.4f}"
                    else:
                        amount = "-"
                    tree_missing.insert("", tk.END, values=(material, amount))
                
                tree_missing.pack(side="left", fill=tk.BOTH, expand=True)
                scrollbar_missing.pack(side="right", fill="y")
            
            # ===== GEFUNDENE MATERIALIEN =====
            found_label = ttk.Label(scrollable_frame, text=f"Gefundene Materialien ({len(df_found)})", font=("Arial", 11, "bold"))
            found_label.pack(pady=(10, 5) if len(df_missing) > 0 else (0, 5), anchor="w")
            
            # Frame für gefundene Materialien
            found_frame = ttk.Frame(scrollable_frame)
            found_frame.pack(fill=tk.X, pady=(0, 10))
            
            # Treeview für gefundene Materialien
            tree_found = ttk.Treeview(found_frame, columns=("Material", "Amount", "Einheit"), show="headings", height=8)
            tree_found.heading("Material", text="Material")
            tree_found.heading("Amount", text="Amount")
            tree_found.heading("Einheit", text="Einheit")
            
            tree_found.column("Material", width=500, minwidth=300)
            tree_found.column("Amount", width=150, minwidth=100)
            tree_found.column("Einheit", width=150, minwidth=100)
            
            # Scrollbar für gefundene
            scrollbar_found = ttk.Scrollbar(found_frame, orient="vertical", command=tree_found.yview)
            tree_found.configure(yscrollcommand=scrollbar_found.set)
            
            # Daten einfügen
            for _, row in df_found.iterrows():
                material = str(row['Material'])
                amount = f"{row['Final_Amount_A1']:.4f}" if pd.notna(row['Final_Amount_A1']) else "-"
                unit = str(row['Final_Unit_A1']) if pd.notna(row['Final_Unit_A1']) else "-"
                tree_found.insert("", tk.END, values=(material, amount, unit))
            
            tree_found.pack(side="left", fill=tk.BOTH, expand=True)
            scrollbar_found.pack(side="right", fill="y")
            
            # Canvas und Scrollbar packen
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Schließen-Button
            button_frame = ttk.Frame(preview_window)
            button_frame.pack(pady=(5, 10))
            close_button = ttk.Button(button_frame, text="Schließen", command=preview_window.destroy)
            close_button.pack()
            
        except Exception as e:
            error_msg = f"Fehler beim Laden der Materialien: {str(e)}"
            messagebox.showerror("Fehler", error_msg)
    
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
        
        # In separatem Thread ausführen, damit UI nicht einfriert
        thread = threading.Thread(target=self._process_epd_thread)
        thread.daemon = True
        thread.start()
    
    def _process_epd_thread(self):
        try:
            # Auth-Liste erstellen
            auth_list = [
                {
                    "url": self.auth_url1,
                    "user": self.auth_user1.get(),
                    "password": self.auth_password1.get()
                },
                {
                    "url": self.auth_url2,
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
            resp = process_epd(
                main_file_path=self.main_file_path.get(),
                sheet_name=self.sheet_name.get(),
                mapping_file_path=self.mapping_file_path.get(),
                full_name=self.full_name.get(),
                epd_unit=self.epd_unit.get(),
                root_repository=self.root_repository.get(),
                target_repository=self.target_repository.get(),
                start_row_index=0,  # Start immer bei Zeile 0
                material_column_index=self.column_letter_to_index(self.material_column.get()),
                amount_column_index=self.column_letter_to_index(self.amount_column.get()),
                auth_list=auth_list,
                method_lib=method_lib,
                url_api=self.url_api,
                api_key=self.api_key,
                output_dir=output_dir,
                skip_missing_materials=self.skip_missing.get(),
                log_callback=None
            )
            
            if resp is not None:
                self.root.after(0, messagebox.showinfo, "Erfolg", f"EPD wurde erfolgreich erstellt!\n\nStatus-Code: {resp.status_code}\n\nJSON-Datei gespeichert in:\n{output_dir}")
            else:
                self.root.after(0, messagebox.showwarning, "Abgebrochen", "Der Vorgang wurde abgebrochen.")
                
        except Exception as e:
            import traceback
            error_msg = f"Fehler: {str(e)}\n\n{traceback.format_exc()}"
            self.root.after(0, messagebox.showerror, "Fehler", error_msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = BoMToEPDGUI(root)
    root.mainloop()

