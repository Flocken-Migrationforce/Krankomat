import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import datetime
import win32com.client as win32
import os
import csv

# ----------------- Dateipfade / User-Config Dateien -----------------
TEMPLATE_BODY_PATH = "template.txt"
TEMPLATE_SUBJECT_PATH = "template-subject.txt"
EMPFAENGER_PATH = "Empfaenger.txt"
USER_VORNAME_PATH = "Mein_Vorname.txt"
USER_NACHNAME_PATH = "Mein_Nachname.txt"
USER_EMAIL_PATH = "Meine_E-Mail-Adresse.txt"
USER_MATRIKEL_PATH = "Meine_Matrikelnummer.txt"
USER_STUDIENGANG_PATH = "Mein_Studiengang.txt"
USER_STUNDENPLAN_PATH = "Stundenplan.txt"


def load_file_text(path, default=None):
    if not os.path.exists(path):
        return default
    with open(path, "r", encoding="utf-8") as f:
        return f.read().strip()
def load_email_from_berufspraxis_txt():
    text = load_file_text("Berufspraxis.txt", default="")
    for line in text.splitlines():
        if line.strip() and ";" in line:
            parts = line.split(";")
            if len(parts) >= 3:
                email_raw = parts[2].strip()
                return parse_email_cell(email_raw)
    return ""

def read_empfaenger(path):
    rows = []
    if not os.path.exists(path):
        return rows
    with open(path, newline='', encoding="utf-8") as csvfile:
        reader = csv.DictReader(csvfile, delimiter=';')
        for r in reader:
            rows.append(r)
    return rows

def parse_email_cell(cell):
    if not cell:
        return ""
    c = cell.strip()
    if "](" in c and c.endswith(")"):
        try:
            left, right = c.split("](", 1)
            inner = right.rstrip(")")
            if inner.startswith("mailto:"):
                return inner.split("mailto:",1)[1]
            return inner
        except Exception:
            pass
    if c.startswith("mailto:"):
        return c.split("mailto:",1)[1]
    return c

def render_template(template, context):
    try:
        return template.format(**context)
    except Exception as e:
        t = template
        for k,v in context.items():
            t = t.replace("{" + k + "}", str(v))
        return t

def generate_anreden(anrede_list):
    """Erzeugt die Begrüßungszeilen genau aus der Anrede-Spalte, fügt Komma an falls nötig."""
    lines = []
    for a in anrede_list:
        line = a.strip()
        # if not line.endswith(","):
        #     line += ","
        lines.append(line)
    return ",\n".join(lines)


class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tipwindow or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 10
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, background="#ffffe0", relief="solid", borderwidth=1)
        label.pack()

    def hide_tip(self, event=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()


class KrankmeldungApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Krankomat – Bereitet eine E-Mail für dich vor, damit du dich schnell krank- oder gesundmelden kannst beim ZAF, HAW-Prüfungsamt und Dozenten.")
        self.geometry("1150x800") # initiale Fenstergröße beim Öffnen

        # User Config laden
        self.user_vorname = load_file_text(USER_VORNAME_PATH, default="")
        self.user_nachname = load_file_text(USER_NACHNAME_PATH, default="")
        self.user_matrikel = load_file_text(USER_MATRIKEL_PATH, default="")
        self.user_email = load_file_text(USER_EMAIL_PATH, default="")
        self.user_studiengang = load_file_text(USER_STUDIENGANG_PATH, default="")

        self.template_body = load_file_text(TEMPLATE_BODY_PATH, default="")
        self.template_subject = load_file_text(TEMPLATE_SUBJECT_PATH, default="Krankmeldung EGOV 2025 {Datum} [{Vornamen} {Nachname}, {Matrikelnummer}]")

        # ZPD als Standard-Adressat:
        self.var_zpd = tk.BooleanVar(value=True)
        self.var_pruef = tk.BooleanVar(value=False)

        emp_rows = read_empfaenger(EMPFAENGER_PATH)

        # Platzhalter-Variable "empfaenger" ersetzen mit "Anrede" aus Empfaenger.txt
        self.greeting_items = []
        for r in emp_rows:
            anrede = (r.get("Anrede") or "").strip()
            modul = (r.get("Modul") or "").strip()
            email_raw = (r.get("Email-Adresse") or r.get("Email") or "").strip()
            email = parse_email_cell(email_raw)
            self.greeting_items.append({
                "anrede": anrede,
                "modul": modul,
                "email": email
            })
        self.gkv_var = tk.BooleanVar()
        self.matrikel_var = tk.StringVar(value=self.user_matrikel)

        self._build_top_panel()
        self._build_left_panel()
        self._build_center_panel()
        self._build_right_panel()
        self._build_output_panel()
        self._build_buttons_panel()
        self._build_email_info_panel()

        self.grid_columnconfigure(0)  # Panel 1 breiter
        # self.grid_columnconfigure(0, minsize=450)  # Panel 1 breiter
        self.grid_columnconfigure(1, weight=2, minsize=600)
        self.grid_columnconfigure(2)
        # self.grid_columnconfigure(2, minsize=350)


        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)  # Für unteres Panel rechts (Vorschau)



        # Setze Standardwerte und Events:
        # Linkes Panel: "Vorlesungszeit." ankreuzen
        if "Vorlesungszeit." in self.left_vars:
            self.left_vars["Vorlesungszeit."].set(True)

        # Center Panel: Erste ZPD Checkbox ankreuzen (z.B. über greeting_items)
        zpd_anrede = self.greeting_items[0]["anrede"] if self.greeting_items else None
        if zpd_anrede and zpd_anrede in self.prof_vars:
            self.prof_vars[zpd_anrede].set(True)

        # Datum Krankheitsbeginn auf heute setzen
        heute = datetime.datetime.now().strftime("%d.%m.%Y")
        self.entry_datum.delete(0, tk.END)
        self.entry_datum.insert(0, heute)



        self._update_preview()

    def export_output(self):
        text = self.output_text.get("1.0", "end-1c")
        file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                 filetypes=[("Textdateien", "*.txt"), ("Alle Dateien", "*.*")])
        if file_path:
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(text)
            messagebox.showinfo("Erfolg", "Datei wurde gespeichert:\n" + file_path)

    def _build_top_panel(self):
        self.top = ttk.Frame(self, padding=6)
        self.top.grid(row=0, column=0, columnspan=3, sticky="ew", padx=6, pady=6)
        # self.top.columnconfigure(7, weight=1, minsize=10)  #2510221952FF OUT.
        self.top.grid_columnconfigure(7, minsize=40)  # Breite für Datumsspalte
        self.top.grid_columnconfigure(8, minsize=100)  # Spacer-Spalte, ganz eng
        self.top.grid_columnconfigure(9, minsize=120)  # Breite für Radio Buttons

        # Erste Zeile: Name, Nachname, Datum etc.
        ttk.Label(self.top, text="Vorname:").grid(row=0, column=0, sticky="e")
        self.entry_vorname = ttk.Entry(self.top, width=20)
        self.entry_vorname.grid(row=0, column=1, padx=10)
        self.entry_vorname.insert(0, self.user_vorname)
        self.entry_vorname.bind("<KeyRelease>", lambda e: self._update_preview())

        ttk.Label(self.top, text="Nachname:").grid(row=0, column=2, sticky="e")
        self.entry_nachname = ttk.Entry(self.top, width=40)
        self.entry_nachname.grid(row=0, column=3, padx=10)
        self.entry_nachname.insert(0, self.user_nachname)
        self.entry_nachname.bind("<KeyRelease>", lambda e: self._update_preview())

        ttk.Label(self.top, text="Erster Krankheitstag:").grid(row=0, column=6, sticky="e")
        self.entry_datum = ttk.Entry(self.top, width=14)
        self.entry_datum.grid(row=0, column=7, padx=10, sticky="w")
        self.entry_datum.bind("<KeyRelease>", lambda e: self._update_preview())

        ttk.Label(self.top, text="Letzter Krankheitstag:").grid(row=1, column=6, sticky="e")
        self.entry_datum_2 = ttk.Entry(self.top, width=14)
        self.entry_datum_2.grid(row=1, column=7, padx=10, sticky="w")
        self.entry_datum_2.bind("<KeyRelease>", lambda e: self._update_preview())

        def on_datum_2_change(event=None):
            text = self.entry_datum_2.get().strip()
            if not text:
                self.meldung_var.set("Krankmeldung")
                self._update_preview()

        self.entry_datum_2.bind("<KeyRelease>", on_datum_2_change)

        # Radio-button
        self.meldung_var = tk.StringVar(value="Krankmeldung")

        rb1 = ttk.Radiobutton(self.top, text="Krankmeldung", variable=self.meldung_var, value="Krankmeldung",
                              command=self._update_preview)
        rb1.grid(row=0, column=8, sticky="w", padx=10)

        rb2 = ttk.Radiobutton(self.top, text="Gesundmeldung", variable=self.meldung_var, value="Gesundmeldung",
                              command=self._update_preview)
        rb2.grid(row=1, column=8, sticky="w", padx=(10, 0))

        # 2510221928FF OUT
        # ttk.Radiobutton(self, text="Krankmeldung", variable=self.krankmeldung, value="Krankmeldung",
        #                 command=self._update_preview).pack(anchor="w", pady=2)
        # ttk.Radiobutton(right, text="Gesundmeldung", variable=self.krankmeldung, value="Gesundmeldung",
        #                 command=self._update_preview).pack(anchor="w", pady=2)
        # ttk.Separator(right, orient="horizontal").pack(fill="x", pady=6)


        def set_heute_gesund():
            self.entry_datum_2.delete(0, tk.END)
            self.entry_datum_2.insert(0, datetime.datetime.now().strftime("%d.%m.%Y"))
            self.meldung_var.set("Gesundmeldung")
            self._update_preview()

        btn_heute_gesund = ttk.Button(self.top, text="heute wieder gesund", command=set_heute_gesund)
        btn_heute_gesund.grid(row=1, column=9, padx=10)

        # Zweite Zeile: Matrikelnummer und Absender-E-Mail
        ttk.Label(self.top, text="Matrikelnummer:").grid(row=1, column=0, sticky="e")
        ttk.Entry(self.top, textvariable=self.matrikel_var, width=15).grid(row=1, column=1, sticky="w", padx=10)

        ttk.Label(self.top, text="Absender-E-Mail:").grid(row=1, column=2, sticky="e")
        self.entry_sender_email = ttk.Entry(self.top, width=40)
        self.entry_sender_email.grid(row=1, column=3, columnspan=3, sticky="w", padx=10)
        self.entry_sender_email.insert(0, self.user_email)
        self.entry_sender_email.bind("<KeyRelease>", lambda e: self._update_preview())

    def _build_left_panel(self):
        left = ttk.LabelFrame(self, text="Heute verpasse ich krankheitsbedingt ...", padding=6)
        left.grid(row=1, column=0, sticky="nsw", padx=6, pady=6)
        opts = [
            "Vorlesungszeit.",
            "Berufspraxis.",
            "eine Prüfungsleistung / Klausur / Präsentation.",
            "die restlichen Stunden des Tages.\nHeute Morgen war ich aber schon da.\n(\"Krank im Dienst\")"
        ]
        self.left_vars = {}

        def on_berufspraxis_toggle():
            pass
        #     if self.left_vars["Berufspraxis."].get():
        #         if "Ausbildungsleitung" in self.prof_vars:
        #             self.prof_vars["Ausbildungsleitung"].set(True)
        #         email = self.load_email_from_berufspraxis_txt()
        #         if "Berufspraxis." in self.prof_vars:
        #             self.prof_vars["Berufspraxis."].set(True)
        #     else:
        #         if "Ausbildungsleitung" in self.prof_vars:
        #             self.prof_vars["Ausbildungsleitung"].set(False)
        #         if "Berufspraxis." in self.prof_vars:
        #             self.prof_vars["Berufspraxis."].set(False)
        #     self._update_preview()

        for opt in opts:
            v = tk.BooleanVar()
            if opt == "Berufspraxis.":
                cb = ttk.Checkbutton(left, text=opt, variable=v, command=on_berufspraxis_toggle)
            else:
                cb = ttk.Checkbutton(left, text=opt, variable=v, command=self._update_preview)
            cb.pack(anchor="w", pady=2)
            self.left_vars[opt] = v

    def _build_center_panel(self):
        center = ttk.LabelFrame(self, text="Krankmelden bei", padding=6)
        center.grid(row=1, column=1, sticky="nsew", padx=6, pady=6)
        center.columnconfigure(0, weight=1)
        center.rowconfigure(0, weight=1)

        # Canvas + Scrollbar erstellen
        canvas = tk.Canvas(center)
        scrollbar = ttk.Scrollbar(center, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Checkboxen aufbauen:
        self.prof_vars = {}
        row_idx = 0

        for item in self.greeting_items:
            modul = item["modul"]
            email = item["email"]
            anrede = item["anrede"]
            label = f"{modul} ({email})" if email else modul

            var = tk.BooleanVar()
            # Verbindung herstellen (korrekt gebundene Lambda-Variable!)
            var.trace_add('write', lambda *args, a=anrede: self._update_preview())
            cb = ttk.Checkbutton(scrollable_frame, text=label, variable=var)
            cb.grid(row=row_idx, column=0, sticky="w", pady=2)
            self.prof_vars[anrede] = var
            row_idx += 1

        # Attribute sichern – dann existiert scrollable_frame global
        self.center_canvas = canvas
        self.center_scrollbar = scrollbar
        self.center_scrollable_frame = scrollable_frame


        #
        #     # Label: Modul und E-Mail
        #     label = f"{modul} ({email})"
        #     var = tk.BooleanVar()
        #     self.prof_vars[anrede] = var
        #     var.trace_add('write', lambda *args, a=anrede: self._update_preview())
        #     ttk.Checkbutton(scrollable_frame, text=label, variable=var).grid(row=row_idx, column=0, sticky="w", pady=2)
        #     row_idx += 1
        #     # var.trace_add('write', lambda *args: self._update_preview())
        #     # cb = ttk.Checkbutton(center, text=label, variable=var)
        #     # cb.grid(row=row_idx, column=0, sticky="w", pady=2)
        #     # self.prof_vars[anrede] = var
        #     row_idx += 1
        #
        #     erste_anrede = ""
        #     if self.greeting_items:
        #         erste_anrede = self.greeting_items[0].get("anrede", "")
        #         if erste_anrede != "":
        #             self.template_body = self.template_body.replace("{{anrede}}", erste_anrede)
        #
        # #
        # # greeting_items[0] ist ZPD mit allen Informationen
        # if self.greeting_items:
        #     zpd = self.greeting_items[0]
        #     # Text inkl. E-Mail-Adresse wie bei anderen Empfängern
        #     text = f"{zpd['anrede']} ({zpd['email']})" if zpd.get('email') else zpd['anrede']
        #     self.var_zpd = tk.BooleanVar(value=True)
        #     ttk.Checkbutton(center, text=text, variable=self.var_zpd, command=self._update_preview).grid(row=0,
        #                                                                                                  column=0,
        #                                                                                                  sticky="w",
        #                                                                                                  pady=2)
        #
        # # Restliche Empfänger dynamisch aufbauen ...
        # for i, empfaenger in enumerate(self.greeting_items[1:], start=1):
        #     var_name = f"var_empfaenger_{i}"
        #     setattr(self, var_name, tk.BooleanVar(value=False))
        #     var = getattr(self, var_name)
        #     text = f"{empfaenger['anrede']} ({empfaenger['email']})" if empfaenger.get('email') else empfaenger[
        #         'anrede']
        #     ttk.Checkbutton(center, text=text, variable=var, command=self._update_preview).grid(row=i, column=0,
        #                                                                                         sticky="w", pady=2)


        #
        # zpd_lbl = ""
        # zpd_email = ""
        # if self.greeting_items:
        #     zpd_lbl = self.greeting_items[0]["anrede"] or ""
        #     zpd_email = self.greeting_items[0].get("email", "")
        # ttk.Checkbutton(center, text=zpd_lbl, variable=self.var_zpd, command=self._update_preview).grid(row=0, column=0, sticky="w", pady=2)
        # zpd_text = f"{zpd_lbl} ({zpd_email})" if zpd_email else zpd_lbl

        # # BooleanVar mit True initialisieren, sodass es angehakt ist
        # self.var_zpd = tk.BooleanVar(value=True)

        # ttk.Checkbutton(center, text=zpd_text, variable=self.var_zpd, command=self._update_preview).grid(row=0,
        # ttk.Checkbutton(center, text=zpd_text, variable=self.var_zpd, command=self._update_preview).grid(row=0,


        # --- Automatische Auswahl anhand Stundenplan.txt ---
        self._auto_select_by_stundenplan()

    def _auto_select_by_stundenplan(self):
        """Wählt automatisch Empfänger anhand von Stundenplan.txt und aktuellem Wochentag."""

        if not os.path.exists(USER_STUNDENPLAN_PATH):
            return

        # Krankheitsbeginn-Datum aus Eingabefeld lesen
        datum_text = self.entry_datum.get().strip()

        try:
            # Datum aus Format TT.MM.JJJJ lesen
            krank_datum = datetime.datetime.strptime(datum_text, "%d.%m.%Y").date()
        except ValueError:
            # Falls Format falsch oder leer ist, abbrechen
            return

        # Wochentag vom eingegebenen Datum (nicht 'heute')
        wochentag_map = {
            0: "Montag",
            1: "Dienstag",
            2: "Mittwoch",
            3: "Donnerstag",
            4: "Freitag",
            5: "Samstag",
            6: "Sonntag",
        }
        heute = wochentag_map[krank_datum.weekday()]

        # Stundenplanzeilen einlesen
        passende_module = []
        with open(USER_STUNDENPLAN_PATH, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                parts = line.split(";")
                if len(parts) >= 2:
                    tag, modul = parts[0].strip(), parts[1].strip()
                    if tag.lower() == heute.lower():
                        passende_module.append(modul)

        # Checkboxen automatisch aktivieren, wenn Modulnamen passen
        for item in self.greeting_items:
            modulname = item.get("modul", "").strip()
            anrede = item.get("anrede", "").strip()
            for pmod in passende_module:
                if pmod.lower() in modulname.lower():
                    if anrede in self.prof_vars:
                        self.prof_vars[anrede].set(True)
                        break

    def _build_right_panel(self):
        right = ttk.LabelFrame(self, text="Details / Optionen", padding=6)
        right.grid(row=1, column=2, sticky="nsew", padx=6, pady=6)
        # right.grid(row=1, column=7, sticky="nw", padx=6, pady=6) # 2510221926FF OUT

        def on_eau_toggle():
            if self.eau_var.get():
                self.attest_var.set(True)
            self._update_preview()  # Optionale Aktualisierung der Vorschau


        self.unfall_var = tk.BooleanVar()
        ttk.Checkbutton(right, text="Unfall", variable=self.unfall_var, command=self._update_preview).pack(anchor="w",
                                                                                                           pady=2)
        self.gkv_var = tk.BooleanVar()
        ttk.Checkbutton(right, text="GKV", variable=self.gkv_var, command=self._update_preview).pack(anchor="w", pady=2)
        self.attest_var = tk.BooleanVar()
        ttk.Checkbutton(right, text="Attest", variable=self.attest_var, command=self._update_preview).pack(anchor="w",
                                                                                                           pady=2)
        self.eau_var = tk.BooleanVar()
        eau_cb = ttk.Checkbutton(right, text="eAU", variable=self.eau_var, command=on_eau_toggle)
        eau_cb.pack(anchor="w", pady=2)


        # Multiline-Textfeld für Bemerkungen
        ttk.Label(right, text="Bemerkung / voraussichtliche Dauer:").pack(anchor="w", pady=(8, 0))
        self.bemerkung_entry = tk.Text(right, width=30, height=5, wrap=tk.WORD)
        self.bemerkung_entry.pack(anchor="w", pady=2)
        self.bemerkung_entry.bind("<KeyRelease>", lambda e: self._update_preview())

        # Variablen für Checkboxen mit Beispieltexten
        self.bemerkung_1_var = tk.BooleanVar()
        self.bemerkung_2_var = tk.BooleanVar()
        self.bemerkung_3_var = tk.BooleanVar()

        def toggle_bemerkung(var, text):
            if var.get():
                aktuell = self.bemerkung_entry.get("1.0", "end-1c").strip()
                if text not in aktuell:
                    neu = aktuell + ("\n" if aktuell else "") + text
                    self.bemerkung_entry.delete("1.0", tk.END)
                    self.bemerkung_entry.insert("1.0", neu)
                    self._update_preview()
            else:
                aktuell = self.bemerkung_entry.get("1.0", "end-1c")
                lines = aktuell.split("\n")
                if text in lines:
                    lines.remove(text)
                    neu = "\n".join(lines)
                    self.bemerkung_entry.delete("1.0", tk.END)
                    self.bemerkung_entry.insert("1.0", neu)
                    self._update_preview()

        ttk.Checkbutton(right, text="nur 1 Tag krank", variable=self.bemerkung_1_var,
                        command=lambda: toggle_bemerkung(self.bemerkung_1_var, "nur 1 Tag krank")).pack(anchor="w",
                                                                                                        pady=2)
        ttk.Checkbutton(right, text="vsl. ungefähr 3 Tage krank", variable=self.bemerkung_2_var,
                        command=lambda: toggle_bemerkung(self.bemerkung_2_var, "vsl. ungefähr 3 Tage krank")).pack(
            anchor="w", pady=2)
        ttk.Checkbutton(right, text="vsl. länger als 3 Tage, Arzt wird aufgesucht", variable=self.bemerkung_3_var,
                        command=lambda: toggle_bemerkung(self.bemerkung_3_var,
                                                         "vsl. länger als 3 Tage, Arzt wird aufgesucht")).pack(
            anchor="w", pady=2)

    def _build_output_panel(self):
        out = ttk.LabelFrame(self, text="Output / Vorschau", padding=6)
        out.grid(row=2, column=0, columnspan=3, sticky="nsew", padx=6, pady=6)
        out.columnconfigure(0, weight=1)
        out.rowconfigure(0, weight=1)

        self.preview = scrolledtext.ScrolledText(out, wrap="word")
        self.preview.grid(row=0, column=0, sticky="nsew")

    def _build_buttons_panel(self):
        btns = ttk.Frame(self, padding=6)
        btns.grid(row=3, column=0, columnspan=3, sticky="ew")
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)

        ttk.Button(btns, text="E-Mails vorbereiten und manuell abschicken", command=lambda: self._prepare_emails(send_now=False)).grid(row=0, column=0, sticky="ew", padx=6)
        ttk.Button(btns, text="E-Mails sofort abschicken", command=lambda: self._prepare_emails(send_now=True)).grid(row=0, column=1, sticky="ew", padx=6)

    def _build_email_info_panel(self):
        frame = ttk.LabelFrame(self, text="E-Mail-Adressen & Betreff zum Kopieren", padding=6)
        frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=6, pady=6)
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="In Kopie/CC: ").grid(row=0, column=0, sticky="w")
        self.text_emails = tk.Text(frame, height=3, wrap="word")
        self.text_emails.grid(row=1, column=0, sticky="ew")
        self.text_emails.config(state="disabled")

        ttk.Label(frame, text="Betreff:").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.text_subject = tk.Text(frame, height=1, wrap="word")
        self.text_subject.grid(row=3, column=0, sticky="ew")
        self.text_subject.config(state="normal") # 2510221932FF Editierbar gemacht für Benutzer.


    def _gather_context(self):
        vorname = self.entry_vorname.get().strip()
        nachname = self.entry_nachname.get().strip()

        name_field = f"{vorname} {nachname}".strip()
        name_str = f"Name: {name_field}" if vorname else f"Name:{name_field}"

        anreden_auswahl = []
        for anrede, var in self.prof_vars.items():
            if var.get():
                anreden_auswahl.append(anrede.strip())

        ctx = {
            "vorname": vorname or "Vorname",
            "nachname": nachname or "Nachname",
            "Datum": self.entry_datum.get().strip() or datetime.datetime.now().strftime("%d.%m.%Y"),
            "art": self.meldung_var.get(),
            "bemerkung": self.bemerkung_entry.get("1.0", "end-1c").strip() or "",
            "Vornamen": vorname or "Vorname",
            "Nachname": nachname or "Nachname",
            "Matrikelnummer": self.matrikel_var.get().strip() or "",
            "krankmeldung": "x" if self.meldung_var.get() == "Krankmeldung" else "",
            "gesundmeldung": "x" if self.meldung_var.get() == "Gesundmeldung" else "",
            "DatumKrank": self.entry_datum.get().strip(),
            "DatumGesund": self.entry_datum_2.get().strip() or "",
            "attest": "ja" if self.attest_var.get() else "nein",
            "eAU": "ja" if (self.eau_var.get() and self.gkv_var.get() and self.attest_var.get()) else ("ja" if self.eau_var.get() else "nein"),
            "prüfungstag": "ja" if self.left_vars["eine Prüfungsleistung / Klausur / Präsentation."].get() else "nein",
            "unfall": "ja" if self.unfall_var.get() else "nein",
            "namefield": name_str,
            "studiengang": self.user_studiengang,
            "anrede": ", ".join(anreden_auswahl),
            "bemerkung_1": "x" if self.bemerkung_1_var.get() else "",
            "bemerkung_2": "x" if self.bemerkung_2_var.get() else "",
            "bemerkung_3": "x" if self.bemerkung_3_var.get() else "",
            "Datum2": self.entry_datum_2.get().strip()
        }
        self._name_field = name_str
        return ctx

    def _update_preview(self):
        scroll_pos = self.preview.yview()
        self.preview.delete("1.0", tk.END)
        self.preview.insert(tk.END, "Hier steht der Text für die Vorschau")
        ctx = self._gather_context()

        anreden_auswahl = [anrede.strip() for anrede, var in self.prof_vars.items() if var.get()]

        anrede_text = ",\n".join(anreden_auswahl)
        ctx["anrede"] = anrede_text

        body = render_template(self.template_body, ctx)

        if anreden_auswahl:
            # Ersten Eintrag großschreiben (nur erstes Wort oder komplett)
            s = anreden_auswahl[0]
            anreden_auswahl[0] = s[0].upper() + s[1:] if s else s

        anrede_text = ",\n".join(anreden_auswahl)
        ctx["anrede"] = anrede_text

        body = render_template(self.template_body, ctx)

        # Alle ausgewählten Anreden zusammensetzen:
        # anrede_text = "\n".join(anreden_auswahl)
        # ctx["anrede"] = anrede_text

        to_list = []
        cc_list = []
        """
        # 2510262246OUT.
        selected = [a for a, v in self.prof_vars.items() if v.get()]
        to_list, cc_list = [], []

        if selected:
            # Erster Empfänger ist Hauptempfänger; alle weiteren in CC.
            first = selected[0]
            primary = next((g for g in self.greeting_items if g["anrede"] == first), None)
            if primary and primary.get("email"):
                cc_list.append(primary["email"]) #2510262244OLD: war mal to_list.append. Machte keinen Sinn für die Business Requirements.

            # Alle weiteren Empfänger als CC
            # for a in selected[1:]:
            #     g = next((g for g in self.greeting_items if g["anrede"] == a), None)
            #     if g and g.get("email"):
            #         cc_list.append(g["email"])
        """
        # ZPD (z.B. erstes Element mit entsprechendem Modul oder Anrede) immer als To-Empfänger
        zpd = next((g for g in self.greeting_items if "ZPD" in g.get("modul", "") or "ZPD" in g.get("anrede", "")), None)
        if zpd and zpd.get("email"):
            to_list = [zpd["email"]]

        # Alle Anwender-Auswahlen im mittleren Panel als CC, auch ZPD nicht doppeln
        for anrede in anreden_auswahl:
            if zpd and anrede == zpd.get("anrede"):
                continue  # ZPD nicht nochmal in CC
            g = next((g for g in self.greeting_items if g["anrede"] == anrede), None)
            if g and g.get("email"):
                cc_list.append(g["email"])

        # Generate greeting text
        greeting_text = generate_anreden(anreden_auswahl)
        ctx["anrede"] = greeting_text
        body = render_template(self.template_body, ctx)



        # Empfänger-Liste am Ende entfernen, da Begrüßung alle enthält
        # Optional kannst du hier andere Zusätze einfügen

        opts = []
        if self.unfall_var.get(): opts.append("Unfall")
        if self.attest_var.get(): opts.append("Attest")
        if self.eau_var.get(): opts.append("eAU")

        # automatische Setzung von eAU wenn GKV UND Unfall gesetzt wurden
        if self.gkv_var.get() and self.attest_var.get():
            self.eau_var.set(True)
        # else:
        #     self.eau_var.set(False)

        self.preview.delete("1.0", tk.END)
        self.preview.insert(tk.END, body)
        self.preview.yview_moveto(scroll_pos[0])

        # Betreff füllen
        subject_template = load_file_text(TEMPLATE_SUBJECT_PATH, default="Krankmeldung EGOV 2025 {Datum} [{Vornamen} {Nachname}, {Matrikelnummer}]")
        subject_filled = render_template(subject_template, {
            "Datum": ctx.get("Datum"),
            "Vornamen": ctx.get("Vornamen"),
            "Nachname": ctx.get("Nachname"),
            "Matrikelnummer": ctx.get("Matrikelnummer","")
        })

        # E-Mail-Adressen-Feld und Betreff-Feld aktualisieren
        combined_emails = "; ".join(to_list + cc_list)

        self.text_emails.config(state="normal")
        self.text_emails.delete("1.0", tk.END)
        self.text_emails.insert(tk.END, combined_emails)
        self.text_emails.config(state="disabled")

        self.text_subject.config(state="normal")
        self.text_subject.delete("1.0", tk.END)
        self.text_subject.insert(tk.END, subject_filled)
        self.text_subject.config(state="normal")



    def _load_template_from_file(self):
        path = filedialog.askopenfilename(title="Template (body) auswählen", filetypes=[("Text files","*.txt"),("All files","*.*")])
        if path:
            self.template_body = load_file_text(path, default=self.template_body)
            self._update_preview()

    def _save_text_as_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                 filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                                                 title="Text als Datei speichern")
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    text_content = self.preview.get("1.0", "end-1c")
                    f.write(text_content)
                messagebox.showinfo("Erfolg", f"Text erfolgreich gespeichert:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler beim Speichern der Datei:\n{e}")


    def _prepare_emails(self, send_now=False):
        initial_anrede = "Sehr geehrte Damen und Herren vom ZPD"
        if self.template_body is not None:
            self.template_body = self.template_body.replace("{anrede}", initial_anrede)

        self.update_idletasks()  # alle GUI-Ereignisse abarbeiten
        self._update_preview()  # Kontext inkl. Empfängerliste neu generieren


        ctx = self._gather_context()

        empfaenger_greetings = []
        for anrede, var in self.prof_vars.items():
            if var.get():
                empfaenger_greetings.append(anrede.strip())

        greeting_text = generate_anreden(empfaenger_greetings)
        ctx["empfаenger"] = greeting_text

        body = render_template(self.template_body, ctx)



        to_list = []
        cc_list = []

        # Alle markierten Empfänger aus self.prof_vars ermitteln
        selected = [a for a, v in self.prof_vars.items() if v.get()]

        if selected:
            # Erster markierter Empfänger ist Hauptempfänger (TO)
            first = selected[0]
            first_obj = next((g for g in self.greeting_items if g["anrede"] == first), None)
            if first_obj and first_obj.get("email"):
                to_list.append(first_obj["email"])

            # Alle weiteren markierten Empfänger kommen in CC
            for a in selected[1:]:
                g = next((g for g in self.greeting_items if g["anrede"] == a), None)
                if g and g.get("email"):
                    cc_list.append(g["email"])



        subject = load_file_text(TEMPLATE_SUBJECT_PATH, default="Krankmeldung EGOV 2025 {Datum} [{Vornamen} {Nachname}, {Matrikelnummer}]")
        subject_filled = render_template(subject, {
            "Datum": ctx.get("Datum"),
            "Vornamen": ctx.get("Vornamen"),
            "Nachname": ctx.get("Nachname"),
            "Matrikelnummer": ctx.get("Matrikelnummer","")
        })

        # Textfelder aktualisieren (nur zum Kopieren, nicht editierbar)
        combined_emails = "; ".join(to_list + cc_list)
        self.text_emails.config(state="normal")
        self.text_emails.delete("1.0", tk.END)
        self.text_emails.insert(tk.END, combined_emails)
        self.text_emails.config(state="disabled")

        self.text_subject.config(state="normal")
        self.text_subject.delete("1.0", tk.END)
        self.text_subject.insert(tk.END, subject_filled)
        self.text_subject.config(state="normal")


        sender_email = self.entry_sender_email.get().strip() or None

        if not to_list:
            if not messagebox.askyesno("Kein Haupt-Empfänger", "Es wurde kein Haupt-Empfänger ausgewählt. Fortfahren?"):
                return

        try:
            create_outlook_mail(to_list, cc_list, subject_filled, body, sender_email=sender_email, send_now=send_now)
            messagebox.showinfo("Fertig", "E-Mail als Entwurf erstellt." if not send_now else "E-Mail wurde gesendet.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Erstellen/Senden der E-Mail:\n{e}")


def create_outlook_mail(to_addresses, cc_addresses, subject, body, sender_email=None, send_now=False):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(to_addresses) if isinstance(to_addresses, (list,tuple)) else to_addresses or ""
    mail.CC = ";".join(cc_addresses) if cc_addresses else ""
    mail.Subject = subject
    mail.Body = body
    if sender_email:
        accounts = outlook.Session.Accounts
        for account in accounts:
            if account.SmtpAddress.lower() == sender_email.lower():
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                break
    if send_now:
        mail.Send()
    else:
        mail.Save()
        mail.Display()

if __name__ == "__main__":
    app = KrankmeldungApp()
    app.mainloop()
