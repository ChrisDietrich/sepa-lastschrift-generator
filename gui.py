#!/usr/bin/env python3
"""Grafische Oberfläche für sepa_lastschrift.py.

Einstiegspunkt für die mit PyInstaller gebaute Windows-.exe (siehe Hinweise am
Ende dieser Datei). Die eigentliche SEPA-Logik liegt unverändert in
sepa_lastschrift.py - diese Datei ist nur die Bedienoberfläche drumherum:
Excel-Datei auswählen, Fälligkeitstag eingeben, Knopf drücken.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import sepa_lastschrift as sepa


def app_dir() -> Path:
    """Ordner neben der .exe (bzw. neben diesem Skript im Entwicklungsbetrieb).

    Wichtig für PyInstaller --onefile: __file__ zeigt dort zur Laufzeit auf einen
    temporären Entpackungsordner, der bei jedem Start neu angelegt und danach
    wieder gelöscht wird. Vereinsdaten (creditor.json) müssen aber über den
    nächsten Programmstart hinweg erhalten bleiben, deshalb liegen sie neben der
    .exe selbst (sys.executable), nicht neben dem entpackten Skript.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


CREDITOR_PATH = app_dir() / "creditor.json"


def open_folder(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)])
    else:
        subprocess.run(["xdg-open", str(path)])


class CreditorDialog(tk.Toplevel):
    """Formular zum Anlegen/Bearbeiten der Vereins-/Gläubigerdaten (creditor.json)."""

    def __init__(self, parent: tk.Tk, path: Path):
        super().__init__(parent)
        self.path = path
        self.title("Vereinsdaten")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        existing: dict = {}
        if path.exists():
            try:
                existing = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                existing = {}

        pad = {"padx": 10, "pady": 4}
        frm = ttk.Frame(self, padding=12)
        frm.grid()

        ttk.Label(frm, text="Vereinsname:").grid(row=0, column=0, sticky="w", **pad)
        self.name_var = tk.StringVar(value=existing.get("name", ""))
        ttk.Entry(frm, textvariable=self.name_var, width=40).grid(row=0, column=1, **pad)

        ttk.Label(frm, text="IBAN:").grid(row=1, column=0, sticky="w", **pad)
        self.iban_var = tk.StringVar(value=existing.get("iban", ""))
        ttk.Entry(frm, textvariable=self.iban_var, width=40).grid(row=1, column=1, **pad)

        ttk.Label(frm, text="Gläubiger-ID:").grid(row=2, column=0, sticky="w", **pad)
        self.creditor_id_var = tk.StringVar(value=existing.get("creditor_id", ""))
        ttk.Entry(frm, textvariable=self.creditor_id_var, width=40).grid(row=2, column=1, **pad)
        ttk.Label(frm, text="Format DE##ZZZ#########, siehe Kontoauszug/Sparkasse",
                  foreground="#555").grid(row=3, column=1, sticky="w", padx=10)

        ttk.Label(frm, text="BIC (optional):").grid(row=4, column=0, sticky="w", **pad)
        self.bic_var = tk.StringVar(value=existing.get("bic", ""))
        ttk.Entry(frm, textvariable=self.bic_var, width=40).grid(row=4, column=1, **pad)

        ttk.Label(frm, text="Währung:").grid(row=5, column=0, sticky="w", **pad)
        self.currency_var = tk.StringVar(value=existing.get("currency", "EUR"))
        ttk.Entry(frm, textvariable=self.currency_var, width=10).grid(row=5, column=1, sticky="w", **pad)

        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=2, pady=(10, 0))
        ttk.Button(btns, text="Speichern", command=self.save).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="Abbrechen", command=self.destroy).grid(row=0, column=1, padx=6)

    def save(self) -> None:
        creditor = sepa.Creditor(
            name=self.name_var.get(),
            iban=self.iban_var.get(),
            creditor_id=self.creditor_id_var.get(),
            bic=self.bic_var.get(),
            currency=self.currency_var.get() or "EUR",
        )
        try:
            creditor = creditor.validate()
        except sepa.ValidationError as exc:
            messagebox.showerror("Fehler in den Vereinsdaten", str(exc), parent=self)
            return

        self.path.write_text(
            json.dumps(
                {
                    "name": creditor.name,
                    "iban": creditor.iban,
                    "creditor_id": creditor.creditor_id,
                    "bic": creditor.bic,
                    "currency": creditor.currency,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        messagebox.showinfo("Gespeichert", "Vereinsdaten wurden gespeichert.", parent=self)
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SEPA-Lastschrift Generator")
        self.resizable(False, False)

        self.input_path: Path | None = None

        pad = {"padx": 10, "pady": 6}
        frm = ttk.Frame(self, padding=12)
        frm.grid(sticky="nsew")

        ttk.Label(frm, text="Vereinsdaten", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        self.creditor_info_var = tk.StringVar()
        ttk.Label(frm, textvariable=self.creditor_info_var, foreground="#333", justify="left").grid(
            row=1, column=0, columnspan=2, sticky="w", **pad
        )

        ttk.Separator(frm).grid(row=2, column=0, columnspan=2, sticky="ew", pady=6)

        ttk.Label(frm, text="1. Mitgliederliste auswählen", font=("", 10, "bold")).grid(
            row=3, column=0, columnspan=2, sticky="w"
        )
        self.file_label = ttk.Label(frm, text="(keine Datei ausgewählt)", foreground="#555")
        self.file_label.grid(row=4, column=0, sticky="w", **pad)
        ttk.Button(frm, text="Datei wählen...", command=self.choose_file).grid(row=4, column=1, **pad)

        ttk.Label(frm, text="2. Fälligkeitstag (Einzugsdatum)", font=("", 10, "bold")).grid(
            row=5, column=0, columnspan=2, sticky="w"
        )
        self.date_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.date_var, width=15).grid(row=6, column=0, sticky="w", **pad)
        ttk.Button(frm, text="+7 Tage", command=self.fill_default_date).grid(row=6, column=1, sticky="w", **pad)
        ttk.Label(frm, text="Format TT.MM.JJJJ, z. B. 01.10.2026 - +7 Tage füllt einen Vorschlag ein, "
                            "bitte anhand der Vorlagefrist eurer Bank prüfen", foreground="#555",
                  wraplength=380, justify="left").grid(row=7, column=0, columnspan=2, sticky="w", **pad)

        self.advanced_visible = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frm, text="Erweiterte Einstellungen anzeigen", variable=self.advanced_visible,
            command=self.toggle_advanced,
        ).grid(row=8, column=0, columnspan=2, sticky="w", **pad)

        self.adv_frame = ttk.Frame(frm)
        self.adv_frame.grid(row=9, column=0, columnspan=2, sticky="w")

        ttk.Label(self.adv_frame, text="Lastschriftsequenz:").grid(row=0, column=0, sticky="w", **pad)
        self.sequence_var = tk.StringVar(value="RCUR")
        ttk.Combobox(
            self.adv_frame, textvariable=self.sequence_var, state="readonly",
            values=["FRST", "RCUR", "OOFF", "FNAL"], width=8,
        ).grid(row=0, column=1, sticky="w", **pad)

        ttk.Label(self.adv_frame, text="Lastschriftart:").grid(row=1, column=0, sticky="w", **pad)
        self.instrument_var = tk.StringVar(value="CORE")
        ttk.Combobox(
            self.adv_frame, textvariable=self.instrument_var, state="readonly",
            values=["CORE", "B2B"], width=8,
        ).grid(row=1, column=1, sticky="w", **pad)

        self.batch_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            self.adv_frame, text="Sammelbuchung (statt Einzelbuchung)", variable=self.batch_var,
        ).grid(row=2, column=0, columnspan=2, sticky="w", **pad)

        self.adv_frame.grid_remove()

        ttk.Separator(frm).grid(row=10, column=0, columnspan=2, sticky="ew", pady=10)

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=11, column=0, columnspan=2, sticky="w")
        ttk.Button(btn_frame, text="Vereinsdaten...", command=self.edit_creditor).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(btn_frame, text="SEPA-Datei erstellen", command=self.generate).grid(row=0, column=1)

        self.status_var = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self.status_var, foreground="#555", wraplength=380).grid(
            row=12, column=0, columnspan=2, sticky="w", pady=(10, 0)
        )

        self.refresh_creditor_display()
        if not CREDITOR_PATH.exists():
            self.after(200, self.first_run_hint)

    def first_run_hint(self) -> None:
        messagebox.showinfo(
            "Willkommen",
            "Es wurden noch keine Vereinsdaten hinterlegt.\n\n"
            "Bitte zuerst auf 'Vereinsdaten...' klicken und Name, IBAN und "
            "Gläubiger-ID des Vereins eintragen.",
        )

    def refresh_creditor_display(self) -> None:
        if not CREDITOR_PATH.exists():
            self.creditor_info_var.set("(noch keine Vereinsdaten hinterlegt - bitte 'Vereinsdaten...' klicken)")
            return
        try:
            data = json.loads(CREDITOR_PATH.read_text(encoding="utf-8"))
            self.creditor_info_var.set(
                f"Verein: {data.get('name', '')}\n"
                f"IBAN: {data.get('iban', '')}\n"
                f"Gläubiger-ID: {data.get('creditor_id', '')}"
            )
        except (OSError, json.JSONDecodeError):
            self.creditor_info_var.set("(Vereinsdaten konnten nicht gelesen werden - bitte 'Vereinsdaten...' klicken)")

    def fill_default_date(self) -> None:
        self.date_var.set((dt.date.today() + dt.timedelta(days=7)).strftime("%d.%m.%Y"))

    def toggle_advanced(self) -> None:
        if self.advanced_visible.get():
            self.adv_frame.grid()
        else:
            self.adv_frame.grid_remove()

    def choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Mitgliederliste auswählen",
            filetypes=[("Excel/CSV-Dateien", "*.xlsx *.xlsm *.csv"), ("Alle Dateien", "*.*")],
        )
        if path:
            self.input_path = Path(path)
            self.file_label.config(text=self.input_path.name, foreground="#000")

    def edit_creditor(self) -> None:
        self.wait_window(CreditorDialog(self, CREDITOR_PATH))
        self.refresh_creditor_display()

    def generate(self) -> None:
        if self.input_path is None:
            messagebox.showwarning("Hinweis", "Bitte zuerst eine Mitgliederliste auswählen.")
            return
        if not CREDITOR_PATH.exists():
            messagebox.showwarning("Hinweis", "Bitte zuerst die Vereinsdaten eintragen ('Vereinsdaten...').")
            return

        try:
            collection_date = sepa.check_date(self.date_var.get(), "Der Fälligkeitstag")
        except sepa.ValidationError as exc:
            messagebox.showerror("Ungültiges Datum", str(exc))
            return

        try:
            result = sepa.generate_sepa_file(
                input_path=self.input_path,
                creditor_path=CREDITOR_PATH,
                collection_date=collection_date,
                sequence_type=self.sequence_var.get(),
                instrument_code=self.instrument_var.get(),
                batch_booking=self.batch_var.get(),
                out_dir=self.input_path.parent,
            )
        except FileNotFoundError:
            messagebox.showerror("Fehler", f"Vereinsdaten-Datei nicht gefunden: {CREDITOR_PATH}")
            return
        except sepa.SheetNotFoundError as exc:
            messagebox.showerror("Fehler", str(exc))
            return
        except sepa.RowValidationErrors as exc:
            messagebox.showerror(
                "Fehler in der Mitgliederliste",
                "Es sind Fehler in den Daten aufgetreten - es wurde keine SEPA-Datei erstellt:\n\n"
                + "\n".join(f"- {e}" for e in exc.errors),
            )
            return
        except sepa.NoRowsError as exc:
            messagebox.showerror("Fehler", str(exc))
            return
        except (KeyError, sepa.ValidationError) as exc:
            messagebox.showerror("Fehler in den Vereinsdaten", str(exc))
            return

        self.status_var.set(f"Erstellt: {result.output_path.name}")
        open_now = messagebox.askyesno(
            "SEPA-Datei erstellt",
            "Die SEPA-XML-Datei wurde erfolgreich erstellt.\n\n"
            f"Anzahl Datensätze: {result.row_count}\n"
            f"Summe: {result.total_amount:.2f} {result.currency}\n"
            f"Datei: {result.output_path}\n\n"
            "Ordner jetzt öffnen?",
        )
        if open_now:
            open_folder(result.output_path.parent)


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()

# ---------------------------------------------------------------------------
# Hinweise zum Bauen einer Windows-.exe mit PyInstaller (auf einem Windows-PC):
#
#   pip install pyinstaller openpyxl
#   pyinstaller --onefile --windowed --name SEPA-Lastschrift-Generator gui.py
#
# Ergebnis liegt danach in dist\SEPA-Lastschrift-Generator.exe. Diese Datei kann
# einzeln weitergegeben werden; creditor.json wird beim ersten Speichern der
# Vereinsdaten automatisch daneben angelegt.
# ---------------------------------------------------------------------------
