import customtkinter as ctk
from tkinter import messagebox
import subprocess
import os
import sys
import ctypes

# --- CustomTkinter Beállítások ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

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


class KMSActivatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("KMS Aktivátor Magyar Közigazgatási EDU")
        self.geometry("900x750")
        self.minsize(850, 700)

        # Középre helyezés
        self._center_window()

        # Rendszergazdai jog ellenőrzése
        self._admin_check_result = self._is_admin()

        # Változó a folyamatban lévő műveletek követésére
        self._current_process = None

        # UI felépítése
        self._create_ui()

        # Admin figyelmeztetés (ha szükséges)
        if not self._admin_check_result:
            self.after(100, self._show_admin_warning)

    def _center_window(self):
        """Ablak középre helyezése a képernyőn."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def _create_ui(self):
        """Felhasználói felület felépítése."""
        # --- Felső vezérlő panel ---
        self.control_panel = ctk.CTkFrame(self, height=80, corner_radius=10)
        self.control_panel.pack(pady=(10, 5), padx=20, fill="x")

        # Bal oldali cím rész
        title_frame = ctk.CTkFrame(self.control_panel, fg_color="transparent")
        title_frame.pack(side="left", padx=15, pady=10, fill="y")

        self.title_label = ctk.CTkLabel(title_frame,
                                        text="KMS Aktivátor Magyar Közigazgatási EDU",
                                        font=ctk.CTkFont(size=18, weight="bold"))
        self.title_label.pack(anchor="w")

        self.credit_label = ctk.CTkLabel(title_frame,
                                         text="Készítette: Csipai Gergő 2025 (Eredeti: Ocskó Zsolt & Lovász Ádám)",
                                         font=ctk.CTkFont(size=10),
                                         text_color="gray70")
        self.credit_label.pack(anchor="w")

        # Jobb oldali beállítások
        settings_frame = ctk.CTkFrame(self.control_panel, fg_color="transparent")
        settings_frame.pack(side="right", padx=15, pady=10)

        self.appearance_mode_label = ctk.CTkLabel(settings_frame, text="Megjelenés:")
        self.appearance_mode_label.pack(side="left", padx=(0, 5))

        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(settings_frame,
                                                             values=["Rendszer", "Világos", "Sötét"],
                                                             command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.set("Rendszer")
        self.appearance_mode_optionemenu.pack(side="left")

        # Admin státusz jelzés
        admin_status = "✓ Rendszergazda" if self._admin_check_result else "✗ Nincs admin jog"
        admin_color = "#28a745" if self._admin_check_result else "#dc3545"
        self.admin_label = ctk.CTkLabel(settings_frame, text=admin_status,
                                        text_color=admin_color, font=ctk.CTkFont(weight="bold"))
        self.admin_label.pack(side="left", padx=(15, 0))

        # --- Tabview ---
        self.tab_view = ctk.CTkTabview(self, width=800, height=400, corner_radius=10)
        self.tab_view.pack(pady=10, padx=20, fill="both", expand=True)

        # Tabok létrehozása
        for tab_name, keys_dict in KEYS.items():
            tab_frame = self.tab_view.add(tab_name)
            if tab_name == "Office 2016 - 2021":
                self._create_section(tab_frame, keys_dict, self._office_actions)
            else:
                self._create_section(tab_frame, keys_dict, self._windows_actions)

        # --- Kimeneti terület ---
        output_frame = ctk.CTkFrame(self, corner_radius=10)
        output_frame.pack(pady=(5, 10), padx=20, fill="both", expand=True)

        ctk.CTkLabel(output_frame,
                     text="Visszajelzési Folyamat / Kimenet:",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(10, 5), anchor="w", padx=10)

        # Gombok a kimenet felett
        button_frame = ctk.CTkFrame(output_frame, fg_color="transparent", height=40)
        button_frame.pack(fill="x", padx=10, pady=(0, 5))

        ctk.CTkButton(button_frame, text="Kimenet Törlése",
                      command=self._clear_output,
                      fg_color="#6c757d", hover_color="#5a6268",
                      width=120).pack(side="left")

        ctk.CTkButton(button_frame, text="Minden Művelet Megszakítása",
                      command=self._cancel_operations,
                      fg_color="#dc3545", hover_color="#c82333",
                      width=180).pack(side="right")

        self.output = ctk.CTkTextbox(output_frame, height=200, font=("Consolas", 10),
                                     activate_scrollbars=True, wrap="word")
        self.output.pack(padx=10, pady=(0, 10), fill="both", expand=True)
        self.output.insert("end", "--- KMS Aktivátor alkalmazás elindult ---\n")
        self.output.insert("end",
                           f"--- Admin jogosultság: {'✓ MEGVAN' if self._admin_check_result else '✗ NINCS'} ---\n\n")

        # --- Lábléc gombok ---
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.pack(pady=10)

        ctk.CTkButton(footer_frame, text="Kilépés", command=self.quit,
                      fg_color="#c43434", hover_color="#8b2525",
                      font=ctk.CTkFont(size=14, weight="bold"),
                      width=150, height=40).pack(side="left", padx=20)

    def _create_section(self, parent, keys_dict, action_handler):
        """Létrehoz egy szekciót a TabView-on belül."""
        # Fő keret
        main_frame = ctk.CTkFrame(parent, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Bal oldali vezérlőpanel
        control_frame = ctk.CTkFrame(main_frame, width=200, corner_radius=10)
        control_frame.pack(side="left", fill="y", padx=(0, 10))
        control_frame.pack_propagate(False)

        ctk.CTkLabel(control_frame, text="Általános Műveletek",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=15)

        # Gombok
        buttons = [
            ("Státusz Ellenőrzés (4)", "#ffc107", "#e0a800", "black",
             lambda: action_handler("check", None, None)),
            ("Kulcs Eltávolítása (5)", "#6c757d", "#5a6268", "white",
             lambda: action_handler("remove", None, None))
        ]

        for text, fg_color, hover_color, text_color, command in buttons:
            btn = ctk.CTkButton(control_frame, text=text, command=command,
                                fg_color=fg_color, hover_color=hover_color,
                                text_color=text_color, width=180, height=35,
                                font=ctk.CTkFont(weight="bold"))
            btn.pack(pady=8)

        # Jobb oldali aktiválási panel
        activate_frame = ctk.CTkScrollableFrame(main_frame,
                                                label_text="Aktiválási Verziók (6)",
                                                label_font=ctk.CTkFont(size=16, weight="bold"),
                                                corner_radius=10)
        activate_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))

        # Aktiválás gombok
        for name, key in keys_dict.items():
            btn = ctk.CTkButton(activate_frame, text=f"Aktiválás: {name}",
                                command=lambda n=name, k=key: action_handler("activate", k, n),
                                fg_color="#28a745", hover_color="#1e7e34",
                                font=ctk.CTkFont(size=12, weight="bold"),
                                height=40, width=400)
            btn.pack(pady=6, padx=10)

    def _is_admin(self):
        """Rendszergazdai jogosultság ellenőrzése."""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False

    def _show_admin_warning(self):
        """Rendszergazdai figyelmeztetés megjelenítése."""
        warning_window = ctk.CTkToplevel(self)
        warning_window.title("❗ Rendszergazdai Jogosultság Hiánya")
        warning_window.geometry("500x220")
        warning_window.resizable(False, False)
        warning_window.transient(self)
        warning_window.grab_set()
        warning_window.lift()

        # Középre helyezés
        warning_window.update_idletasks()
        x = (warning_window.winfo_screenwidth() - 500) // 2
        y = (warning_window.winfo_screenheight() - 220) // 2
        warning_window.geometry(f"500x220+{x}+{y}")

        ctk.CTkLabel(warning_window,
                     text="⚠️ RENDSZERGAZDAI JOGOSULTSÁG SZÜKSÉGES ⚠️",
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color="red").pack(pady=20, padx=10)

        ctk.CTkLabel(warning_window,
                     text="A KMS aktiválási és kulcsműveletek csak rendszergazdaként futtatva működnek.\nKérjük, indítsa újra a programot 'Futtatás rendszergazdaként' opcióval.",
                     wraplength=450, justify="center").pack(pady=10, padx=10)

        ctk.CTkButton(warning_window, text="Értem",
                      command=warning_window.destroy,
                      fg_color="blue", width=100).pack(pady=10)

    def change_appearance_mode_event(self, new_appearance_mode: str):
        """Megjelenési mód váltása."""
        mode_map = {"Rendszer": "System", "Világos": "Light", "Sötét": "Dark"}
        ctk.set_appearance_mode(mode_map.get(new_appearance_mode, "System"))

    def _clear_output(self):
        """Kimeneti szövegdoboz tartalmának törlése."""
        self.output.delete("1.0", "end")

    def _cancel_operations(self):
        """Futó műveletek megszakítása."""
        if self._current_process and self._current_process.poll() is None:
            self._current_process.terminate()
            self.output.insert("end", "\n--- Művelet megszakítva ---\n")
            self.output.see("end")
        else:
            self.output.insert("end", "\n--- Nincs futó művelet ---\n")
            self.output.see("end")

    def _run_command(self, command, success_msg="A parancs sikeresen lefutott."):
        """Parancs futtatása és kimenet megjelenítése."""
        if not self._admin_check_result:
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
        return self._current_process.returncode == 0 if self._current_process else False

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
        top = ctk.CTkToplevel(self)
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

        ctk.CTkLabel(top,
                     text="Írja be a kulcs utolsó 5 karakterét:",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(pady=10)

        ctk.CTkLabel(top,
                     text="(A státusz ellenőrzésnél látható, PID-ként is ismert)",
                     font=ctk.CTkFont(size=10),
                     text_color="gray70").pack(pady=(0, 10))

        input_frame = ctk.CTkFrame(top, fg_color="transparent")
        input_frame.pack(pady=10)

        key_entry = ctk.CTkEntry(input_frame, width=120, placeholder_text="pl: 1A2B3",
                                 justify="center", font=ctk.CTkFont(size=12, weight="bold"))
        key_entry.pack(side="left", padx=(0, 10))
        key_entry.focus()

        def remove_key():
            kulcs = key_entry.get().strip().upper()
            if len(kulcs) == 5 and all(c.isalnum() for c in kulcs):
                self.output.insert("end", f"\n🔧 Office kulcs eltávolítása ({kulcs})...\n")
                success = self._run_command(f'cscript //Nologo "{ospp_path}" /unpkey:{kulcs}',
                                            "Kulcs eltávolítási parancs lefutott.")
                top.destroy()
                if success:
                    self.output.insert("end", f"✅ A(z) {kulcs} kulcs eltávolítva.\n")
                else:
                    self.output.insert("end", f"❌ Nem sikerült eltávolítani a(z) {kulcs} kulcsot.\n")
            else:
                messagebox.showerror("Hiba",
                                     "Pontosan 5 alfanumerikus karaktert írjon be!\n\nPélda: 1A2B3",
                                     parent=top)

        ctk.CTkButton(input_frame, text="Eltávolítás",
                      command=remove_key,
                      fg_color="#dc3545", hover_color="#c82333",
                      width=100).pack(side="left")

        # Enter billentyű támogatása
        top.bind('<Return>', lambda e: remove_key())


if __name__ == "__main__":
    try:
        app = KMSActivatorApp()
        app.mainloop()
    except Exception as e:
        print(f"Alkalmazás indítási hiba: {e}")
        input("Nyomjon Entert a kilépéshez...")
