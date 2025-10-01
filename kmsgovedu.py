import customtkinter as ctk
from tkinter import messagebox
import subprocess
import os
import sys

# --- CustomTkinter Beállítások ---
# Kezdő mód: Rendszer
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
        self.geometry("850x850")

        # Rendszergazdai jog ellenőrzése és figyelmeztetés
        self._admin_check_result = self._is_admin()
        self._show_admin_warning()

        # --- Felső vezérlő panel (Módváltás és cím) ---
        self.control_panel = ctk.CTkFrame(self, height=50)
        self.control_panel.pack(pady=(10, 0), padx=20, fill="x")

        # Cím (Balra igazítva)
        self.title_label = ctk.CTkLabel(self.control_panel, text="KMS Aktivátor Magyar Közigazgatási EDU",
                                        font=ctk.CTkFont(size=18, weight="bold"))
        self.title_label.pack(side="left", padx=15, pady=(10, 0))

        # Készítők adatai (Kisebb betűtípussal a cím alatt)
        self.credit_label = ctk.CTkLabel(self.control_panel,
                                         text="Készítette: Csipai Gergő & Lovász Ádám 2025",
                                         font=ctk.CTkFont(size=10))
        self.credit_label.pack(side="left", padx=15, pady=(0, 10))

        # Megjelenési mód váltó (Jobbra igazítva)
        self.appearance_mode_label = ctk.CTkLabel(self.control_panel, text="Megjelenés:")
        self.appearance_mode_label.pack(side="right", padx=(10, 0), pady=10)
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.control_panel,
                                                             values=["Rendszer", "Világos", "Sötét"],
                                                             command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.set("Rendszer")
        self.appearance_mode_optionemenu.pack(side="right", padx=10, pady=10)

        # --- Tabview (Fül nézet) a funkciók szétválasztásához ---
        self.tab_view = ctk.CTkTabview(self, width=800, height=450)
        self.tab_view.pack(pady=10, padx=20, fill="x")

        # Tabok létrehozása és feltöltése
        for tab_name, keys_dict in KEYS.items():
            tab_frame = self.tab_view.add(tab_name)
            if tab_name == "Office 2016 - 2021":
                self._create_section(tab_frame, keys_dict, self._office_actions)
            else:
                self._create_section(tab_frame, keys_dict, self._windows_actions)

        # --- Kimeneti terület (Console) ---
        ctk.CTkLabel(self, text="Visszajelzési Folyamat / Kimenet:", font=ctk.CTkFont(size=14, weight="bold")).pack(
            pady=(10, 5))
        self.output = ctk.CTkTextbox(self, height=250, font=("Consolas", 10), activate_scrollbars=True)
        self.output.pack(padx=20, fill="x", expand=False)
        self.output.insert("end", "--- A műveletek kimenete itt fog megjelenni ---\n")

        # Kilépés gomb
        ctk.CTkButton(self, text="Kilépés", command=self.quit,
                      fg_color="#c43434", hover_color="#8b2525",
                      font=ctk.CTkFont(size=14, weight="bold"), width=150).pack(pady=15)

    # --- Segédfunkciók és Műveletek ---

    def _is_admin(self):
        """Rendszergazdai jogosultság ellenőrzése."""
        try:
            return os.getuid() == 0
        except AttributeError:
            try:
                import ctypes
                return ctypes.windll.shell32.IsUserAnAdmin() != 0
            except:
                return False

    def _show_admin_warning(self):
        """Külön CustomTkinter ablak a rendszergazdai figyelmeztetésnek."""
        if not self._admin_check_result:
            warning_window = ctk.CTkToplevel(self)
            warning_window.title("❗ Rendszergazdai Jogosultság Hiánya")
            warning_window.geometry("500x200")
            warning_window.transient(self)
            warning_window.grab_set()

            ctk.CTkLabel(warning_window, text="⚠️ RENDSZERGAZDAI JOGOSULTSÁG SZÜKSÉGES ⚠️",
                         font=ctk.CTkFont(size=16, weight="bold"), text_color="red").pack(pady=20, padx=10)
            ctk.CTkLabel(warning_window,
                         text="A KMS aktiválási és kulcsműveletek csak rendszergazdaként futtatva működnek.\nKérjük, indítsa újra a programot 'Futtatás rendszergazdaként' opcióval.",
                         wraplength=450).pack(pady=10, padx=10)
            ctk.CTkButton(warning_window, text="Értem", command=warning_window.destroy, fg_color="blue").pack(pady=10)
            self.wait_window(warning_window)

    def change_appearance_mode_event(self, new_appearance_mode: str):
        """Világos/Sötét mód váltása."""
        mode_map = {"Rendszer": "System", "Világos": "Light", "Sötét": "Dark"}
        ctk.set_appearance_mode(mode_map.get(new_appearance_mode, "System"))

    def _create_section(self, parent, keys_dict, action_handler):
        """Létrehoz egy szekciót a TabView-on belül."""

        control_frame = ctk.CTkFrame(parent)
        control_frame.pack(side="left", padx=20, pady=20, fill="y")

        ctk.CTkLabel(control_frame, text="Általános Műveletek", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10,
                                                                                                               padx=10)

        # Státusz/Eltávolítás gombok
        ctk.CTkButton(control_frame, text="Státusz Ellenőrzés (4)", command=lambda: action_handler("check", None, None),
                      fg_color="#ffc107", hover_color="#e0a800", text_color="black", width=200).pack(pady=10, padx=10)

        ctk.CTkButton(control_frame, text="Kulcs Eltávolítása (5)",
                      command=lambda: action_handler("remove", None, None),
                      fg_color="#6c757d", hover_color="#5a6268", text_color="white", width=200).pack(pady=10, padx=10)

        activate_frame = ctk.CTkScrollableFrame(parent, label_text="Aktiválási Verziók (6)", width=450,
                                                label_font=ctk.CTkFont(size=16, weight="bold"))
        activate_frame.pack(side="right", padx=20, pady=20, fill="both", expand=True)

        # Aktiválás gombok
        for name, key in keys_dict.items():
            ctk.CTkButton(activate_frame, text=f"Aktiválás: {name}",
                          command=lambda n=name, k=key: action_handler("activate", k, n),
                          fg_color="#28a745", hover_color="#1e7e34", font=ctk.CTkFont(size=12, weight="bold"),
                          width=400).pack(pady=5, padx=10)

    def _run_command(self, command, success_msg="A parancs sikeresen lefutott."):
        """Futtatja a parancsot és megjeleníti a kimenetet a CTkTextboxban."""

        if not self._admin_check_result:
            self.output.insert("end", "\n!!! A MŰVELET MEGSZAKÍTVA: Nincs rendszergazdai jogosultság. !!!\n")
            self.output.see("end")
            return

        self.output.insert("end", f"\n---> Futtatás: {' '.join(command)}\n")
        self.output.see("end")
        self.update_idletasks()

        try:
            result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True, encoding='utf-8',
                                    errors='ignore')

            # Visszajelzés a felhasználói felületre
            self.output.insert("end", f"[SIKER] {success_msg}\n")
            self.output.insert("end", result.stdout + "\n")

            if result.stderr:
                self.output.insert("end", f"[HIBA] a parancs futtatásakor (Stderr):\n{result.stderr}\n")

        except subprocess.CalledProcessError as e:
            self.output.insert("end",
                               f"!!! KRITIKUS HIBA (Kód: {e.returncode}): Ellenőrizd a beállításokat, vagy próbáld újra. !!!\n")
            self.output.insert("end", f"Kimenet:\n{e.stdout}\nHiba:\n{e.stderr}\n")
        except Exception as e:
            self.output.insert("end", f"[VÁRATLAN HIBA]: {e}\n")

        self.output.see("end")
        self.update_idletasks()

    # --- Windows / Server KMS Funkciók (Slmgr) ---
    def _windows_actions(self, action, key, name):
        """Kezeli a Windows és Server KMS parancsokat."""
        if action == "check":
            self.output.insert("end", "\n--- Windows/Server státusz ellenőrzés ---\n")
            self._run_command(f'cscript "{SLMGR_PATH}" /dlv', success_msg="Aktiválási státusz lekérdezve.")

        elif action == "remove":
            self.output.insert("end", "\n--- Windows/Server kulcs eltávolítás ---\n")
            self._run_command(f'cscript "{SLMGR_PATH}" /upk', success_msg="Telepített kulcs eltávolítása...")
            self._run_command(f'cscript "{SLMGR_PATH}" /cpky', success_msg="Kulcs törlése a beállításjegyzékből.")
            self.output.insert("end", "A kulcs eltávolítása megtörtént (ha volt telepített kulcs).\n")

        elif action == "activate":
            self.output.insert("end", f"\n--- Aktiválás indítása: {name} ---\n")
            self._run_command(f'cscript "{SLMGR_PATH}" /ipk {key}', success_msg="1/3. Kulcs telepítve.")
            self._run_command(f'cscript "{SLMGR_PATH}" /skms {KMS_SERVER}', success_msg="2/3. KMS szerver beállítva.")
            self._run_command(f'cscript "{SLMGR_PATH}" /ato',
                              success_msg="3/3. Aktiválás megkísérelve. Ellenőrizd a fenti kimenetet a sikerességhez.")
            self.output.insert("end", f"--- Aktiválási Folyamat Befejezve: {name} ---\n")

    # --- Office KMS Funkciók (Ospp) ---
    def _get_office_path(self):
        """Visszaadja a megfelelő ospp.vbs elérési útját."""
        if os.path.exists(OSPP_PATH_X64):
            self.output.insert("end", "[INFO] Office Verzió: 64 bites Office (Office16 mappa).\n")
            return OSPP_PATH_X64
        elif os.path.exists(OSPP_PATH_X86):
            self.output.insert("end", "[INFO] Office Verzió: 32 bites Office (Office16 mappa).\n")
            return OSPP_PATH_X86
        else:
            self.output.insert("end",
                               "!!! HIBA: Nem található az Office ospp.vbs fájlja. Ellenőrizd az Office telepítést (2016+). !!!\n")
            return None

    def _office_actions(self, action, key, name):
        """Kezeli az Office KMS parancsokat."""
        ospp_path = self._get_office_path()
        if not ospp_path:
            return

        if action == "check":
            self.output.insert("end", "\n--- Office státusz ellenőrzés ---\n")
            self._run_command(f'cscript "{ospp_path}" /dstatus', success_msg="Aktiválási státusz lekérdezve.")
            self.output.insert("end", "\n[INFORMÁCIÓ] Jegyezze meg a kulcs utolsó öt karakterét az eltávolításhoz!\n")

        elif action == "remove":
            self.output.insert("end", "\n--- Office kulcs eltávolítás ---\n")
            self._prompt_for_key_removal(ospp_path)

        elif action == "activate":
            self.output.insert("end", f"\n--- Aktiválás indítása: {name} ---\n")
            self._run_command(f'cscript "{ospp_path}" /inpkey:{key}', success_msg="1/4. Kulcs telepítve.")
            self._run_command(f'cscript "{ospp_path}" /sethst:{KMS_SERVER}', success_msg="2/4. KMS szerver beállítva.")
            self._run_command(f'cscript "{ospp_path}" /setprt:1688', success_msg="3/4. Port beállítva.")
            self._run_command(f'cscript "{ospp_path}" /act',
                              success_msg="4/4. Aktiválás megkísérelve. Ellenőrizd a fenti kimenetet a sikerességhez.")
            self.output.insert("end", f"--- Aktiválási Folyamat Befejezve: {name} ---\n")

    # --- Office Kulcs Eltávolító Ablak ---
    def _prompt_for_key_removal(self, ospp_path):
        """Külön CustomTkinter ablak a kulcs utolsó 5 karakterének bekérésére."""
        top = ctk.CTkToplevel(self)
        top.title("Office Kulcs Eltávolítása")
        top.geometry("400x150")
        top.transient(self)
        top.grab_set()

        ctk.CTkLabel(top, text="Írja be a kulcs utolsó 5 karakterét: (PID-ként is ismert)",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(pady=10)

        key_entry = ctk.CTkEntry(top, width=100)
        key_entry.pack(pady=5)

        def remove_key():
            kulcs = key_entry.get().strip().upper()
            if len(kulcs) == 5 and kulcs.isalnum():
                self.output.insert("end", f"\nOffice kulcs eltávolítása ({kulcs})...\n")
                self._run_command(f'cscript "{ospp_path}" /unpkey:{kulcs}',
                                  success_msg="Kulcs eltávolítási parancs lefutott.")
                top.destroy()
            else:
                messagebox.showerror("Hiba", "Pontosan 5 alfanumerikus karaktert írjon be!", parent=top)

        ctk.CTkButton(top, text="Eltávolítás", command=remove_key, fg_color="red").pack(pady=10)
        self.wait_window(top)


if __name__ == "__main__":
    # Győződjünk meg róla, hogy a CustomTkinter telepítve van, ha Windows alatt fut
    if sys.platform.startswith('win'):
        try:
            import customtkinter as ctk
        except ImportError:
            print("A CustomTkinter (ctk) modul nincs telepítve.")
            print("Kérjük, telepítse a 'pip install customtkinter' paranccsal.")
            sys.exit(1)

    app = KMSActivatorApp()
    app.mainloop()
