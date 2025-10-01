import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import subprocess
import os
import ctypes

# --- KMS Konfiguráció ---
KMS_SERVER = "kms.edu.hu"
SLMGR_PATH = r"C:\Windows\System32\slmgr.vbs"
OSPP_PATH_X64 = r"C:\Program Files\Microsoft Office\Office16\ospp.vbs"
OSPP_PATH_X86 = r"C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs"

# GVLK kulcsok
KEYS = {
    "Windows 10 / 11": {
        "Win 11 Pro / 10 Pro": "W269N-WFGWX-YVC9B-4J6C9-T83GX",
        "Win 11 Ent / 10 Ent": "NPPR9-FWDCX-D2C8J-H872K-2YT43",
        "Win 11 Pro N / 10 Pro N": "MH37W-N47XK-V7XM9-C7227-GCQG9",
        "Win 11 Ent N / 10 Ent N": "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4",
    },
    "Windows Server": {
        "Server 2022 Datacenter": "WX4NM-KYWYW-QJJR4-XV3QB-6VM33",
        "Server 2022 Standard": "VDYBN-27WPP-V4HQT-9VMD4-VMK7H",
        "Server 2019 Datacenter": "WMDGN-G9PQG-XVVXX-R3X43-63DFG",
        "Server 2019 Standard": "N69G4-B89J2-4G8F4-WWYCC-J464C",
        "Server 2016 Datacenter": "CB7KF-BWN84-R7R2Y-793K2-8XDDG",
        "Server 2016 Standard": "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY",
    },
    "Office 2016 - 2021": {
        "Office 2021 LTSC": "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH",
        "Office 2019": "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP",
        "Office 2016": "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99",
    }
}


class KMSActivatorApp(tk.Tk):
    def __init__(self):
        super().__init__()

        # --- Stílus Beállítások (Default téma a maximális kompatibilitásért) ---
        self.style = ttk.Style(self)
        self.style.theme_use('default') 

        # Fő Háttérszín beállítása (a tk.Tk ablakhoz)
        self.configure(bg="#f0f0f0") 

        # KMS Aktivátor Fő Cím
        self.style.configure('MainTitle.TLabel', font=("Arial", 20, "bold"), foreground="#003399")
        
        # Kredit
        self.style.configure('Credit.TLabel', font=("TkDefaultFont", 8), foreground="gray")

        # Fő Panel stílus (szekciók kerete)
        self.style.configure('Section.TLabelframe', borderwidth=1, relief="solid")
        self.style.configure('Section.TLabelframe.Label', font=("TkDefaultFont", 12, "bold")) 

        # --- Gomb stílusok (ttk.Button-okhoz) ---
        
        # 1. Státusz Ellenőrzés (Sárga/Narancs)
        self.style.configure('Status.TButton', 
                             background='#ffc04d', 
                             foreground='black', 
                             font=('TkDefaultFont', 10, 'bold'), 
                             relief='raised', 
                             borderwidth=1) 
        self.style.map('Status.TButton', 
                       background=[('active', '#ffcc00'), ('!disabled', '#ffc04d')],
                       foreground=[('active', 'black'), ('!disabled', 'black')])

        # 2. Kulcs Eltávolítása (Sötét Szürke)
        self.style.configure('Remove.TButton', 
                             background='#666666', 
                             foreground='white', 
                             font=('TkDefaultFont', 10, 'bold'), 
                             relief='raised', 
                             borderwidth=1) 
        self.style.map('Remove.TButton', 
                       background=[('active', '#888888'), ('!disabled', '#666666')],
                       foreground=[('active', 'white'), ('!disabled', 'white')])

        # 3. Aktiválás (Fényes Zöld)
        self.style.configure('Activate.TButton', 
                             background='#4CAF50', 
                             foreground='white', 
                             font=('TkDefaultFont', 10, 'bold'), 
                             relief='raised', 
                             borderwidth=1) 
        self.style.map('Activate.TButton', 
                       background=[('active', '#66BB6A'), ('!disabled', '#4CAF50')],
                       foreground=[('active', 'white'), ('!disabled', 'white')])


        self.title("KMS Aktivátor Magyar Közigazgatási EDU")
        self.geometry("900x750")
        self.minsize(850, 700)
        
        self._admin_check_result = self._is_admin()
        self._current_process = None

        # Fő keret a grid elrendezéshez
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill="both", expand=True)

        # Címkék/keretek elhelyezése a main_frame-ben (grid)
        self.main_frame.grid_columnconfigure(0, weight=1)

        self._create_ui()

        if not self._admin_check_result:
            self.after(100, self._show_admin_warning)
            
        self.after(50, self._center_window)

    def _center_window(self):
        """Ablak középre helyezése a képernyőn."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def _create_ui(self):
        """Felhasználói felület felépítése a grid elrendezésben."""
        
        row_counter = 0

        # --- 1. Felső Cím és Kreditek (Fejléc) ---
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=row_counter, column=0, pady=(20, 10), padx=20, sticky='ew')
        row_counter += 1

        ttk.Label(header_frame, text="KMS Aktivátor Magyar Közigazgatási EDU", style='MainTitle.TLabel').pack(pady=(0, 5))
        ttk.Label(header_frame, 
                  text="Készítette: Csipai Gergő 2025 (Eredeti: Ocskó Zsolt & Lovász Ádám)", 
                  style='Credit.TLabel').pack(pady=(0, 10))

        # Fő KMS Felület cím
        ttk.Label(self.main_frame, text="--- KMS Aktivátor Közigazgatási EDU felület ---", font=("Consolas", 10)).grid(row=row_counter, column=0, pady=(5, 10))
        row_counter += 1

        # --- 2. Aktiválási Szekciók Kerete ---
        sections_frame = ttk.Frame(self.main_frame, padding="10 10 10 10")
        sections_frame.grid(row=row_counter, column=0, pady=10, padx=20, sticky="ew")
        row_counter += 1
        
        # Konfiguráljuk a sections_frame oszlopait az elrendezéshez
        sections_frame.grid_columnconfigure(0, weight=1)
        sections_frame.grid_columnconfigure(1, weight=1)
        sections_frame.grid_columnconfigure(2, weight=1)
        
        # A szekciók elhelyezése a sections_frame-en belül (grid)
        
        # Windows 10/11 Szekció
        win_frame = self._create_section_frame(sections_frame, "Windows 10/11")
        win_frame.grid(row=0, column=0, padx=10, sticky="nsew")
        self._create_buttons_in_section(win_frame, KEYS["Windows 10 / 11"], self._windows_actions)

        # Windows Server Szekció
        server_frame = self._create_section_frame(sections_frame, "Windows Server")
        server_frame.grid(row=0, column=1, padx=10, sticky="nsew")
        self._create_buttons_in_section(server_frame, KEYS["Windows Server"], self._windows_actions)

        # Office Szekció
        office_frame = self._create_section_frame(sections_frame, "Office (2016-2021)")
        office_frame.grid(row=0, column=2, padx=10, sticky="nsew")
        self._create_buttons_in_section(office_frame, KEYS["Office 2016 - 2021"], self._office_actions)
        
        # --- 3. Kimeneti terület ---
        output_frame = ttk.Frame(self.main_frame, padding="10 10 10 10")
        output_frame.grid(row=row_counter, column=0, pady=(10, 5), padx=20, sticky="nsew")
        row_counter += 1
        self.main_frame.grid_rowconfigure(row_counter - 1, weight=1) # A kimeneti terület veszi fel a maradék helyet

        ttk.Label(output_frame, 
                  text="Visszajelzési Folyamat / Kimenet:", 
                  font=("TkDefaultFont", 10, "bold")).pack(pady=(5, 2), anchor="w", padx=10)

        # Text widget görgetősávval
        text_scroll = ttk.Scrollbar(output_frame)
        text_scroll.pack(side="right", fill="y")
        
        self.output = tk.Text(output_frame, height=10, font=("Consolas", 10),
                              wrap="word", yscrollcommand=text_scroll.set)
        self.output.pack(padx=10, pady=(0, 10), fill="both", expand=True)
        text_scroll.config(command=self.output.yview)

        # --- 4. Lábléc gomb (tk.Button a stabilitásért és a grid-es elhelyezésért) ---
        
        # Kilépés gomb tk.Button-ként a ttk.Button helyett (fix méret és megjelenés)
        self.exit_button = tk.Button(self.main_frame, text="Kilépés (10)", command=self.quit,
                                     bg='#D32F2F', 
                                     fg='white', 
                                     font=('TkDefaultFont', 12, 'bold'), 
                                     relief='raised',
                                     borderwidth=3, 
                                     activebackground='#E53935',
                                     activeforeground='white',
                                     padx=25, 
                                     pady=10)
        
        # Közvetlenül a fő keret aljára pakoljuk, centerben
        self.exit_button.grid(row=row_counter, column=0, pady=(10, 20))
        row_counter += 1

        # Kezdeti üzenetek a kimenetben
        self.output.insert("end", "--- KMS Aktivátor alkalmazás elindult ---\n")
        self.output.insert("end", f"--- Admin jogosultság: {'✓ MEGVAN' if self._is_admin() else '✗ NINCS'} ---\n\n")
        self.output.see("end")

    def _create_section_frame(self, parent, title):
        """Létrehoz egy keretet a címmel."""
        # A frame-en lévő borderwidth=1 és relief="solid" biztosítja a látható keretet.
        frame = ttk.Frame(parent, padding="10 10 10 10", borderwidth=1, relief="solid")
        
        # Cím
        ttk.Label(frame, text=title, font=("TkDefaultFont", 12, "bold"), foreground="black").pack(pady=(0, 10))
        
        return frame

    def _create_buttons_in_section(self, parent, keys_dict, action_handler):
        """Létrehozza a gombokat egy szekción belül."""

        # Felső vezérlőgombok kerete
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill='x', pady=(0, 15))

        # 1. Státusz Ellenőrzés gomb (Sárga)
        ttk.Button(control_frame, text="Státusz ellenőrzés", style='Status.TButton',
                   command=lambda: action_handler("check", None, None)).pack(fill='x', pady=(0, 5), padx=5)

        # 2. Kulcs Eltávolítása gomb (Sötét szürke)
        ttk.Button(control_frame, text="Kulcs eltávolítása", style='Remove.TButton',
                   command=lambda: action_handler("remove", None, None)).pack(fill='x', pady=5, padx=5)
        
        # Aktiválás Gombok (Zöld)
        for name, key in keys_dict.items():
            btn_text = f"Aktiválás: {name}"
            ttk.Button(parent, text=btn_text, style='Activate.TButton',
                       command=lambda n=name, k=key: action_handler("activate", k, n)).pack(fill='x', pady=5, padx=5)

    def _is_admin(self):
        """Rendszergazdai jogosultság ellenőrzése."""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False

    def _show_admin_warning(self):
        """Rendszergazdai figyelmeztetés megjelenítése."""
        messagebox.showwarning("Rendszergazdai Jogosultság Hiánya",
                               "A KMS aktiválási és kulcsműveletek csak rendszergazdaként futtatva működnek.\nKérjük, indítsa újra a programot 'Futtatás rendszergazdaként' opcióval.",
                               parent=self)

    def _run_command(self, command, success_msg="A parancs sikeresen lefutott."):
        """Parancs futtatása és kimenet megjelenítése."""
        if not self._is_admin():
            self.output.insert("end", "\n❌ MŰVELET MEGSZAKÍTVA: Nincs rendszergazdai jogosultság.\n")
            self.output.see("end")
            return False

        self.output.insert("end", f"\n--- Futtatás: {command} ---\n")
        self.output.see("end")
        self.update_idletasks()

        try:
            # Parancs futtatása
            self._current_process = subprocess.Popen(
                command, 
                shell=True, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE, 
                text=True, 
                encoding='utf-8', 
                errors='ignore'
            )

            stdout, stderr = self._current_process.communicate()

            # Kimenet feldolgozása
            if self._current_process.returncode == 0:
                self.output.insert("end", f"✅ {success_msg}\n")
                if stdout.strip():
                    self.output.insert("end", f"{stdout}\n")
            else:
                self.output.insert("end", f"❌ HIBA (Kód: {self._current_process.returncode})\n")
                if stderr.strip():
                    self.output.insert("end", f"Hibaüzenet:\n{stderr}\n")
                if stdout.strip():
                    self.output.insert("end", f"Kimenet:\n{stdout}\n")

        except subprocess.CalledProcessError as e:
            self.output.insert("end", f"❌ PARANCS HIBA (Kód: {e.returncode})\n")
            self.output.insert("end", f"Kimenet:\n{e.stdout}\nHiba:\n{e.stderr}\n")
        except Exception as e:
            self.output.insert("end", f"❌ VÁRATLAN HIBA: {e}\n")
        finally:
            self._current_process = None

        self.output.see("end")
        self.update_idletasks()
        return self._current_process is not None and self._current_process.returncode == 0

    def _windows_actions(self, action, key, name):
        """Windows és Server KMS műveletek kezelése."""
        if action == "check":
            self.output.insert("end", "\n--- Windows/Server státusz ellenőrzés ---\n")
            self._run_command(f'cscript //Nologo "{SLMGR_PATH}" /dlv', 
                              "Aktiválási státusz lekérdezve.")

        elif action == "remove":
            self.output.insert("end", "\n--- Windows/Server kulcs eltávolítás ---\n")
            self._run_command(f'cscript //Nologo "{SLMGR_PATH}" /upk', 
                              "Telepített kulcs eltávolítva.")
            self._run_command(f'cscript //Nologo "{SLMGR_PATH}" /cpky', 
                              "Kulcs törölve a beállításjegyzékből.")

        elif action == "activate":
            self.output.insert("end", f"\n--- Aktiválás indítása: {name} ---\n")
            steps = [
                (f'cscript //Nologo "{SLMGR_PATH}" /ipk {key}', "1/3. Kulcs telepítve"),
                (f'cscript //Nologo "{SLMGR_PATH}" /skms {KMS_SERVER}', "2/3. KMS szerver beállítva"),
                (f'cscript //Nologo "{SLMGR_PATH}" /ato', "3/3. Aktiválás megkísérelve")
            ]
            
            for cmd, msg in steps:
                if not self._run_command(cmd, msg):
                    self.output.insert("end", f"❌ A(z) {name} aktiválása megszakadt.\n")
                    break
            else:
                self.output.insert("end", f"✅ {name} aktiválási folyamat befejezve.\n")

    def _get_office_path(self):
        """Office ospp.vbs fájl elérési útjának meghatározása."""
        paths_to_check = [
            OSPP_PATH_X64,
            OSPP_PATH_X86,
            r"C:\Program Files\Microsoft Office\Office15\ospp.vbs",
            r"C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs"
        ]
        
        for path in paths_to_check:
            if os.path.exists(path):
                office_bit = "64 bites" if "x86" not in path else "32 bites"
                office_ver = "Office16" if "Office16" in path else "Office15"
                self.output.insert("end", f"[INFO] {office_bit} {office_ver} észlelve: {path}\n")
                return path
        
        self.output.insert("end", "❌ HIBA: Nem található az Office ospp.vbs fájlja.\n")
        self.output.insert("end", "⚠️ Ellenőrizze, hogy telepítve van-e az Office 2013 vagy újabb verzió.\n")
        return None

    def _office_actions(self, action, key, name):
        """Office KMS műveletek kezelése."""
        ospp_path = self._get_office_path()
        if not ospp_path:
            return

        if action == "check":
            self.output.insert("end", "\n--- Office státusz ellenőrzés ---\n")
            self._run_command(f'cscript //Nologo "{ospp_path}" /dstatus', 
                              "Aktiválási státusz lekérdezve.")
            self.output.insert("end", "\n💡 Jegyezze meg a kulcs utolsó öt karakterét az eltávolításhoz!\n")

        elif action == "remove":
            self.output.insert("end", "\n--- Office kulcs eltávolítás ---\n")
            self._prompt_for_key_removal(ospp_path)

        elif action == "activate":
            self.output.insert("end", f"\n--- Aktiválás indítása: {name} ---\n")
            steps = [
                (f'cscript //Nologo "{ospp_path}" /inpkey:{key}', "1/4. Kulcs telepítve"),
                (f'cscript //Nologo "{ospp_path}" /sethst:{KMS_SERVER}', "2/4. KMS szerver beállítva"),
                (f'cscript //Nologo "{ospp_path}" /setprt:1688', "3/4. Port beállítva"),
                (f'cscript //Nologo "{ospp_path}" /act', "4/4. Aktiválás megkísérelve")
            ]
            
            for cmd, msg in steps:
                if not self._run_command(cmd, msg):
                    self.output.insert("end", f"❌ A(z) {name} aktiválása megszakadt.\n")
                    break
            else:
                self.output.insert("end", f"✅ {name} aktiválási folyamat befejezve.\n")

    def _prompt_for_key_removal(self, ospp_path):
        """Office kulcs eltávolításához ablak."""
        top = tk.Toplevel(self)
        top.title("Office Kulcs Eltávolítása")
        top.geometry("450x180")
        top.resizable(False, False)
        top.transient(self)
        top.grab_set()
        top.lift()

        # Középre helyezés
        top.update_idletasks()
        x = (top.winfo_screenwidth() - 450) // 2
        y = (top.winfo_screenheight() - 180) // 2
        top.geometry(f"450x180+{x}+{y}")

        ttk.Label(top, 
                  text="Írja be a kulcs utolsó 5 karakterét:", 
                  font=("TkDefaultFont", 10, "bold")).pack(pady=10)

        ttk.Label(top,
                  text="(A státusz ellenőrzésnél látható, PID-ként is ismert)",
                  font=("TkDefaultFont", 8)).pack(pady=(0, 10))

        input_frame = ttk.Frame(top)
        input_frame.pack(pady=10)

        key_entry = ttk.Entry(input_frame, width=15, justify="center", font=("Consolas", 10, "bold"))
        key_entry.pack(side="left", padx=(0, 10))
        key_entry.focus()

        def remove_key():
            kulcs = key_entry.get().strip().upper()
            if len(kulcs) == 5 and all(c.isalnum() for c in kulcs):
                self.output.insert("end", f"\n🔧 Office kulcs eltávolítása ({kulcs})...\n")
                self._run_command(f'cscript //Nologo "{ospp_path}" /unpkey:{kulcs}',
                                            "Kulcs eltávolítási parancs lefutott.")
                top.destroy()
            else:
                messagebox.showerror("Hiba", 
                                     "Pontosan 5 alfanumerikus karaktert írjon be!\n\nPélda: 1A2B3", 
                                     parent=top)

        ttk.Button(input_frame, text="Eltávolítás", 
                   command=remove_key).pack(side="left")

        # Enter billentyű támogatása
        top.bind('<Return>', lambda e: remove_key())
        
        # Esc billentyű támogatása (opcionális)
        top.bind('<Escape>', lambda e: top.destroy())


if __name__ == "__main__":
    try:
        app = KMSActivatorApp()
        app.mainloop()
    except Exception as e:
        print(f"Alkalmazás indítási hiba: {e}")
