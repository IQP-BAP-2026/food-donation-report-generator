from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


APP_NAME = "Donor Traceability Report"
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "MASTER-SHEET.xlsx"
GENERATOR_PATH = BASE_DIR / "food_generator.py"
PHOTOS_DIR = BASE_DIR / "assets" / "photos"
OUTPUT_DIR = BASE_DIR / "output"

PHOTO_SLOTS = [
    ("Closing photo 1", "closing_photo1.jpg"),
    ("Closing photo 2", "closing_photo2.jpg"),
    ("Closing photo 3", "closing_photo3.jpg"),
    ("Closing photo 4", "closing_photo4.jpg"),
    ("Custom photo 1", "custom_photo1.jpg"),
    ("Custom photo 2", "custom_photo2.jpg"),
    ("Custom photo 3", "custom_photo3.jpg"),
]


class ReportGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("900x760")
        self.minsize(820, 680)

        self.photo_vars: dict[str, tk.StringVar] = {}
        self.donors: list[str] = []
        self.status_var = tk.StringVar(value="Ready")
        self.workbook_var = tk.StringVar()
        self.selected_donor = tk.StringVar()

        self._configure_style()
        self._build_ui()
        self._refresh_workbook_status()
        self._load_donors()

    # ---------------------------------------------------------
    # UI
    # ---------------------------------------------------------

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass

        style.configure("Title.TLabel", font=("Segoe UI", 22, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground="#5f6b72")
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 11, "bold"))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))
        style.configure("Status.TLabel", font=("Segoe UI", 9), foreground="#55636a")

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=22)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root)
        header.pack(fill="x", pady=(0, 16))

        ttk.Label(header, text="Donor Traceability Report", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Manage your workbook and photos, select a donor, and run the existing report generator.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        workbook_frame = ttk.LabelFrame(root, text="1. Excel workbook", padding=14, style="Section.TLabelframe")
        workbook_frame.pack(fill="x", pady=(0, 12))

        workbook_row = ttk.Frame(workbook_frame)
        workbook_row.pack(fill="x")

        self.workbook_label = ttk.Label(workbook_row, textvariable=self.workbook_var)
        self.workbook_label.pack(side="left", fill="x", expand=True)

        ttk.Button(
            workbook_row,
            text="Upload / Replace Excel",
            command=self._replace_excel,
            style="Primary.TButton",
        ).pack(side="right")

        ttk.Label(
            workbook_frame,
            text="The selected workbook is copied into the project as MASTER-SHEET.xlsx."
        ).pack(anchor="w", pady=(8, 0))

        donor_frame = ttk.LabelFrame(root, text="2. Select company / donor", padding=14, style="Section.TLabelframe")
        donor_frame.pack(fill="x", pady=(0, 12))

        donor_row = ttk.Frame(donor_frame)
        donor_row.pack(fill="x")

        self.donor_combo = ttk.Combobox(
            donor_row,
            textvariable=self.selected_donor,
            state="normal",
            height=18,
        )
        self.donor_combo.pack(side="left", fill="x", expand=True)
        self.donor_combo.bind("<Return>", lambda _event: self._generate())

        ttk.Button(donor_row, text="Refresh", command=self._load_donors).pack(side="left", padx=(8, 0))

        ttk.Label(
            donor_frame,
            text="Donor names are read directly from the DONOR / DONANTE / EMPRESA / COMPANY column in the workbook."
        ).pack(anchor="w", pady=(8, 0))

        photos_frame = ttk.LabelFrame(root, text="3. Photos", padding=14, style="Section.TLabelframe")
        photos_frame.pack(fill="both", expand=True, pady=(0, 12))

        ttk.Label(
            photos_frame,
            text="Choose replacement images for the report slots. Leaving a slot unchanged keeps the existing file.",
            wraplength=780,
        ).pack(anchor="w", pady=(0, 10))

        canvas = tk.Canvas(photos_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(photos_frame, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>",
            lambda _event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for label, filename in PHOTO_SLOTS:
            var = tk.StringVar()
            self.photo_vars[filename] = var
            existing = PHOTOS_DIR / filename
            if existing.exists():
                var.set(str(existing))
            else:
                var.set("(not set)")

            row = ttk.Frame(scroll_frame)
            row.pack(fill="x", pady=4)

            ttk.Label(row, text=f"{label}:", width=18).pack(side="left")
            ttk.Label(row, textvariable=var, width=54).pack(side="left", fill="x", expand=True)
            ttk.Button(
                row,
                text="Choose",
                command=lambda f=filename, l=label: self._choose_photo(f, l),
            ).pack(side="right", padx=(8, 0))
            ttk.Button(
                row,
                text="Open",
                command=lambda f=filename: self._open_photo(f),
            ).pack(side="right")

        actions = ttk.Frame(root)
        actions.pack(fill="x")

        ttk.Button(actions, text="Open Output Folder", command=self._open_output).pack(side="left")
        ttk.Button(actions, text="Open Project Folder", command=self._open_project).pack(side="left", padx=(8, 0))

        self.generate_button = ttk.Button(
            actions,
            text="Generate Report",
            command=self._generate,
            style="Primary.TButton",
        )
        self.generate_button.pack(side="right")

        ttk.Label(root, textvariable=self.status_var, style="Status.TLabel").pack(anchor="w", pady=(10, 0))

    # ---------------------------------------------------------
    # Workbook / donor handling
    # ---------------------------------------------------------

    @staticmethod
    def _slug(text: object) -> str:
        value = "" if text is None else str(text).strip().lower()
        replacements = str.maketrans("áéíóúüñ", "aeiouun")
        value = value.translate(replacements)
        return "".join(ch for ch in value if ch.isalnum())

    def _detect_donor_column(self, columns: list[object]) -> str | None:
        preferred = {"donor", "donante", "empresa", "company"}
        ranked: list[tuple[int, str]] = []

        for col in columns:
            raw = str(col)
            slug = self._slug(raw)
            score = 0
            if slug in preferred:
                score = 100
            elif any(token in slug for token in preferred):
                score = 50
            if score:
                ranked.append((score, raw))

        if not ranked:
            return None

        ranked.sort(key=lambda item: (-item[0], item[1]))
        return ranked[0][1]

    def _load_donors(self) -> None:
        if not EXCEL_PATH.exists():
            self.donors = []
            self.donor_combo["values"] = []
            self.status_var.set("MASTER-SHEET.xlsx not found.")
            return

        try:
            self.status_var.set("Reading workbook...")
            self.update_idletasks()
            df = pd.read_excel(EXCEL_PATH)
            column = self._detect_donor_column(list(df.columns))

            if column is None:
                raise ValueError("Could not identify a donor column.")

            donors = (
                df[column]
                .dropna()
                .astype(str)
                .str.strip()
                .loc[lambda s: s != ""]
                .drop_duplicates()
                .tolist()
            )
            donors.sort(key=str.casefold)

            self.donors = donors
            self.donor_combo["values"] = donors

            current = self.selected_donor.get().strip()
            if current not in donors:
                self.selected_donor.set(donors[0] if donors else "")

            self.status_var.set(f"Loaded {len(donors)} donors from {EXCEL_PATH.name}.")
        except Exception as exc:
            self.donors = []
            self.donor_combo["values"] = []
            self.status_var.set("Could not read workbook.")
            messagebox.showerror(APP_NAME, f"Could not read the workbook.\n\n{exc}")

    def _refresh_workbook_status(self) -> None:
        if EXCEL_PATH.exists():
            self.workbook_var.set(f"Current workbook: {EXCEL_PATH.name}")
        else:
            self.workbook_var.set("No MASTER-SHEET.xlsx is currently installed.")

    def _replace_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[
                ("Excel workbook", "*.xlsx *.xlsm *.xls"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        source = Path(path)

        if source.resolve() == EXCEL_PATH.resolve():
            self._load_donors()
            return

        try:
            shutil.copy2(source, EXCEL_PATH)
            self._refresh_workbook_status()
            self._load_donors()
            messagebox.showinfo(APP_NAME, "Excel workbook replaced successfully.")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Could not replace the workbook.\n\n{exc}")

    # ---------------------------------------------------------
    # Photos
    # ---------------------------------------------------------

    def _choose_photo(self, filename: str, label: str) -> None:
        path = filedialog.askopenfilename(
            title=f"Choose {label}",
            filetypes=[
                ("Images", "*.jpg *.jpeg *.png *.webp"),
                ("JPEG", "*.jpg *.jpeg"),
                ("PNG", "*.png"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        PHOTOS_DIR.mkdir(parents=True, exist_ok=True)
        target = PHOTOS_DIR / filename

        try:
            shutil.copy2(path, target)
            self.photo_vars[filename].set(str(target))
            self.status_var.set(f"Updated {filename}.")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Could not copy the photo.\n\n{exc}")

    def _open_photo(self, filename: str) -> None:
        path = PHOTOS_DIR / filename
        if not path.exists():
            messagebox.showinfo(APP_NAME, f"{filename} does not exist yet.")
            return
        self._open_path(path)

    # ---------------------------------------------------------
    # Generate
    # ---------------------------------------------------------

    def _generator_command(self, donor: str) -> tuple[list[str], dict[str, str]]:
        env = os.environ.copy()

        # The existing generator prints Unicode characters (for example ✅/❌).
        # Force UTF-8 in the child process so Windows cp1252 cannot crash it.
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"

        if sys.platform.startswith("win"):
            # Prefer the Windows Python launcher for an exe build so
            # the GUI executable does not recursively launch itself.
            py_launcher = shutil.which("py")
            if py_launcher:
                return [py_launcher, "-3", "-X", "utf8", str(GENERATOR_PATH)], env

            python_exe = shutil.which("python") or shutil.which("python3")
            if python_exe:
                return [python_exe, "-X", "utf8", str(GENERATOR_PATH)], env

        # When the GUI is run as a .py file, use its current interpreter.
        return [sys.executable, "-X", "utf8", str(GENERATOR_PATH)], env

    def _generate(self) -> None:
        donor = self.selected_donor.get().strip()

        if not donor:
            messagebox.showwarning(APP_NAME, "Select a company / donor first.")
            return

        if not EXCEL_PATH.exists():
            messagebox.showerror(APP_NAME, "MASTER-SHEET.xlsx was not found.")
            return

        if not GENERATOR_PATH.exists():
            messagebox.showerror(APP_NAME, f"food_generator.py was not found.\n\n{GENERATOR_PATH}")
            return

        self.generate_button.configure(state="disabled")
        self.status_var.set(f"Generating report for {donor}...")
        self.update_idletasks()

        try:
            command, env = self._generator_command(donor)
            # Existing generator uses stdin for donor selection.
            result = subprocess.run(
                command,
                input=donor + "\n",
                text=True,
                cwd=str(BASE_DIR),
                env=env,
                capture_output=True,
                encoding="utf-8",
                errors="replace",
            )

            output = (result.stdout or "") + ("\n" + result.stderr if result.stderr else "")

            if result.returncode != 0:
                self.status_var.set("Generation failed.")
                messagebox.showerror(
                    APP_NAME,
                    "The existing food_generator.py reported an error.\n\n" + output[-6000:],
                )
                return

            self.status_var.set("Report generated successfully.")
            messagebox.showinfo(
                APP_NAME,
                f"Report generated for:\n{donor}\n\nOutput folder:\n{OUTPUT_DIR}",
            )
        except Exception as exc:
            self.status_var.set("Generation failed.")
            messagebox.showerror(APP_NAME, f"Could not launch the generator.\n\n{exc}")
        finally:
            self.generate_button.configure(state="normal")

    # ---------------------------------------------------------
    # Folders
    # ---------------------------------------------------------

    def _open_output(self) -> None:
        OUTPUT_DIR.mkdir(exist_ok=True)
        self._open_path(OUTPUT_DIR)

    def _open_project(self) -> None:
        self._open_path(BASE_DIR)

    @staticmethod
    def _open_path(path: Path) -> None:
        path = Path(path)
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Could not open:\n{path}\n\n{exc}")


def main() -> None:
    app = ReportGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
