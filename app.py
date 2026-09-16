from __future__ import annotations

import importlib
import os
import shutil
import sys
import threading
import subprocess
import time
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = "Banco de Alimentos Panamá — Reportes"
APP_VERSION = "1.0.0"

# ------------------------------------------------------------
# Runtime folders
# ------------------------------------------------------------

PACKAGED = hasattr(sys, "_MEIPASS")
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))

APP_DATA_DIR = Path(os.environ.get(
    "LOCALAPPDATA",
    str(Path.home() / "AppData" / "Local")
)) / "BAP Donor Traceability"

DOCUMENTS_DIR = Path.home() / "Documents"
OUTPUT_DIR = DOCUMENTS_DIR / "BAP Donor Traceability Reports"
if not DOCUMENTS_DIR.exists():
    OUTPUT_DIR = APP_DATA_DIR / "output"

DATA_FILES = [
    "food-traceability-template.html",
    "food-traceability.css",
    "MASTER-SHEET.xlsx",
]

PHOTO_FILES = [
    "closing_photo1.jpg",
    "closing_photo2.jpg",
    "closing_photo3.jpg",
    "closing_photo4.jpg",
    "custom_photo1.jpg",
    "custom_photo2.jpg",
    "custom_photo3.jpg",
]


def copy_initial_resources() -> None:
    """Install bundled report resources into a writable app-data folder."""
    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    (APP_DATA_DIR / "assets").mkdir(parents=True, exist_ok=True)

    # The workbook is only copied when no user workbook exists.
    for filename in DATA_FILES:
        source = RESOURCE_DIR / filename
        target = APP_DATA_DIR / filename
        if source.exists() and not target.exists():
            shutil.copy2(source, target)

    source_assets = RESOURCE_DIR / "assets"
    target_assets = APP_DATA_DIR / "assets"
    if source_assets.exists():
        for source in source_assets.rglob("*"):
            if not source.is_file():
                continue
            relative = source.relative_to(source_assets)
            target = target_assets / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            if not target.exists():
                shutil.copy2(source, target)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def configure_environment() -> None:
    """Tell the report generator where writable data/output and native DLLs live."""
    # A packaged worker receives these paths from the GUI process. Preserve them.
    os.environ.setdefault("BAP_REPORT_DATA_DIR", str(APP_DATA_DIR))
    os.environ.setdefault("BAP_REPORT_OUTPUT_DIR", str(OUTPUT_DIR))

    # In the final EXE, the required GTK/Pango DLLs are bundled under
    # _MEIPASS\weasy_dlls. Point WeasyPrint directly at that folder so the
    # recipient does NOT need MSYS2 installed. When running from source on the
    # developer machine, fall back to the normal MSYS2 UCRT64 location.
    bundled_dlls = RESOURCE_DIR / "weasy_dlls"
    if bundled_dlls.exists():
        os.environ["WEASYPRINT_DLL_DIRECTORIES"] = str(bundled_dlls)
    else:
        msys_dlls = Path(r"C:\msys64\ucrt64\bin")
        if msys_dlls.exists():
            os.environ.setdefault("WEASYPRINT_DLL_DIRECTORIES", str(msys_dlls))


def load_generator():
    """Import the existing V9 generator after paths are configured."""
    configure_environment()
    return importlib.import_module("food_generator")


class ReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("760x570")
        self.minsize(700, 520)
        self.configure(bg="#f5f7f6")

        copy_initial_resources()
        configure_environment()

        self.generator = None
        self.donors: list[str] = []
        self.selected_donor = tk.StringVar()
        self.status = tk.StringVar(value="Listo")
        self.workbook_status = tk.StringVar()
        self.last_pdf: Path | None = None

        self._style()
        self._build()
        self._load_donors()

    # ---------------------------------------------------------
    # UI
    # ---------------------------------------------------------

    def _style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("App.TFrame", background="#f5f7f6")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#f5f7f6", font=("Segoe UI", 23, "bold"))
        style.configure("Subtitle.TLabel", background="#f5f7f6", foreground="#64716b", font=("Segoe UI", 10))
        style.configure("CardTitle.TLabel", background="#ffffff", font=("Segoe UI", 11, "bold"))
        style.configure("CardText.TLabel", background="#ffffff", foreground="#66716d", font=("Segoe UI", 9))
        style.configure("Status.TLabel", background="#f5f7f6", foreground="#5d6964", font=("Segoe UI", 9))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(18, 9))
        style.configure("Big.TButton", font=("Segoe UI", 11, "bold"), padding=(22, 12))

    def _build(self) -> None:
        root = ttk.Frame(self, style="App.TFrame", padding=24)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="Reportes de Trazabilidad", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            root,
            text="Selecciona un donante y genera su reporte en PDF.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(3, 18))

        # Workbook card
        card1 = ttk.Frame(root, style="Card.TFrame", padding=16)
        card1.pack(fill="x", pady=(0, 12))
        ttk.Label(card1, text="Base de datos", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(card1, textvariable=self.workbook_status, style="CardText.TLabel").pack(anchor="w", pady=(3, 10))
        row = ttk.Frame(card1, style="Card.TFrame")
        row.pack(fill="x")
        ttk.Button(row, text="Actualizar Excel", command=self._replace_excel).pack(side="left")
        ttk.Button(row, text="Actualizar lista", command=self._load_donors).pack(side="left", padx=(8, 0))
        ttk.Button(row, text="Abrir carpeta de reportes", command=self._open_output).pack(side="right")

        # Donor card
        card2 = ttk.Frame(root, style="Card.TFrame", padding=16)
        card2.pack(fill="x", pady=(0, 12))
        ttk.Label(card2, text="Donante / empresa", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card2,
            text="Escribe para buscar o selecciona un nombre de la lista.",
            style="CardText.TLabel",
        ).pack(anchor="w", pady=(3, 8))

        donor_row = ttk.Frame(card2, style="Card.TFrame")
        donor_row.pack(fill="x")
        self.donor_combo = ttk.Combobox(
            donor_row,
            textvariable=self.selected_donor,
            state="normal",
            height=16,
        )
        self.donor_combo.pack(side="left", fill="x", expand=True)
        self.donor_combo.bind("<KeyRelease>", self._filter_donors)
        self.donor_combo.bind("<Return>", lambda _e: self._generate())

        # Actions card
        card3 = ttk.Frame(root, style="Card.TFrame", padding=16)
        card3.pack(fill="x", pady=(0, 12))
        ttk.Label(card3, text="Acciones", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card3,
            text="El PDF se guarda automáticamente en la carpeta de reportes y se abre al terminar.",
            style="CardText.TLabel",
        ).pack(anchor="w", pady=(3, 12))
        action_row = ttk.Frame(card3, style="Card.TFrame")
        action_row.pack(fill="x")
        ttk.Button(action_row, text="Fotos", command=self._photo_manager).pack(side="left")
        ttk.Button(action_row, text="Abrir último PDF", command=self._open_last_pdf).pack(side="left", padx=(8, 0))
        self.generate_button = ttk.Button(
            action_row,
            text="GENERAR REPORTE",
            command=self._generate,
            style="Big.TButton",
        )
        self.generate_button.pack(side="right")

        ttk.Label(root, textvariable=self.status, style="Status.TLabel").pack(anchor="w", pady=(4, 0))

    # ---------------------------------------------------------
    # Workbook / donor loading
    # ---------------------------------------------------------

    @staticmethod
    def _slug(value: object) -> str:
        import unicodedata
        text = "" if value is None else str(value).strip().lower()
        text = "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")
        return "".join(c for c in text if c.isalnum())

    def _detect_donor_column(self, columns: list[object]) -> str | None:
        preferred = {"donor", "donante", "empresa", "company"}
        exact: list[str] = []
        partial: list[str] = []
        for col in columns:
            raw = str(col)
            key = self._slug(raw)
            if key in preferred:
                exact.append(raw)
            elif any(token in key for token in preferred):
                partial.append(raw)
        return (exact or partial or [None])[0]

    def _load_donors(self) -> None:
        try:
            import pandas as pd
            workbook = APP_DATA_DIR / "MASTER-SHEET.xlsx"
            if not workbook.exists():
                raise FileNotFoundError("No se encontró MASTER-SHEET.xlsx")
            self.status.set("Leyendo base de datos...")
            self.update_idletasks()
            df = pd.read_excel(workbook)
            column = self._detect_donor_column(list(df.columns))
            if column is None:
                raise ValueError("No se encontró una columna DONOR / DONANTE / EMPRESA / COMPANY.")
            donors = (
                df[column]
                .dropna()
                .astype(str)
                .str.strip()
            )
            donors = sorted({d for d in donors if d}, key=str.casefold)
            self.donors = donors
            self.donor_combo["values"] = donors
            current = self.selected_donor.get().strip()
            if current not in donors:
                self.selected_donor.set(donors[0] if donors else "")
            self.workbook_status.set(f"MASTER-SHEET.xlsx · {len(donors)} donantes disponibles")
            self.status.set("Listo")
        except Exception as exc:
            self.workbook_status.set("No se pudo leer el Excel")
            self.status.set("Error al cargar la base de datos")
            messagebox.showerror(APP_NAME, f"No se pudo leer el Excel.\n\n{exc}")

    def _filter_donors(self, _event=None) -> None:
        query = self.selected_donor.get().strip().casefold()
        if not query:
            self.donor_combo["values"] = self.donors
            return
        filtered = [d for d in self.donors if query in d.casefold()]
        self.donor_combo["values"] = filtered

    def _replace_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")],
        )
        if not path:
            return
        try:
            shutil.copy2(path, APP_DATA_DIR / "MASTER-SHEET.xlsx")
            self._load_donors()
            messagebox.showinfo(APP_NAME, "La base de datos fue actualizada correctamente.")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"No se pudo actualizar el Excel.\n\n{exc}")

    # ---------------------------------------------------------
    # Photos
    # ---------------------------------------------------------

    def _photo_manager(self) -> None:
        win = tk.Toplevel(self)
        win.title("Fotos del reporte")
        win.geometry("690x420")
        win.transient(self)
        win.grab_set()

        frame = ttk.Frame(win, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Fotos personalizadas", font=("Segoe UI", 15, "bold")).pack(anchor="w")
        ttk.Label(
            frame,
            text="Puedes reemplazar las tres fotos personalizadas. Las fotos de cierre incluidas en el reporte se mantienen como respaldo.",
            wraplength=630,
        ).pack(anchor="w", pady=(4, 15))

        for i in range(1, 5):
            filename = f"closing_photo{i}.jpg"
            path = APP_DATA_DIR / "assets" / "photos" / filename
            row = ttk.Frame(frame)
            row.pack(fill="x", pady=6)
            ttk.Label(row, text=f"Foto {i}", width=10).pack(side="left")
            label = ttk.Label(row, text=str(path) if path.exists() else "Sin foto")
            label.pack(side="left", fill="x", expand=True)

            def choose(file=filename, label_widget=label):
                source = filedialog.askopenfilename(
                    parent=win,
                    title=f"Seleccionar foto {file[-5]}",
                    filetypes=[("Fotos JPEG", "*.jpg *.jpeg"), ("Todos", "*.*")],
                )
                if not source:
                    return
                target = APP_DATA_DIR / "assets" / "photos" / file
                target.parent.mkdir(parents=True, exist_ok=True)
                try:
                    shutil.copy2(source, target)
                    label_widget.configure(text=str(target))
                except Exception as exc:
                    messagebox.showerror(APP_NAME, f"No se pudo copiar la foto.\n\n{exc}", parent=win)

            ttk.Button(row, text="Elegir", command=choose).pack(side="right")

        ttk.Button(frame, text="Cerrar", command=win.destroy).pack(anchor="e", pady=(18, 0))

    # ---------------------------------------------------------
    # Generation
    # ---------------------------------------------------------

    def _generate(self) -> None:
        donor = self.selected_donor.get().strip()
        if not donor:
            messagebox.showwarning(APP_NAME, "Selecciona un donante primero.")
            return
        if not (APP_DATA_DIR / "MASTER-SHEET.xlsx").exists():
            messagebox.showerror(APP_NAME, "No se encontró la base de datos.")
            return

        self.generate_button.configure(state="disabled")
        self.status.set(f"Generando reporte para {donor}...")

        def worker() -> None:
            env = os.environ.copy()
            env["BAP_REPORT_DATA_DIR"] = str(APP_DATA_DIR)
            env["BAP_REPORT_OUTPUT_DIR"] = str(OUTPUT_DIR)
            env["PYTHONIOENCODING"] = "utf-8"
            env["PYTHONUTF8"] = "1"

            # Run WeasyPrint in a separate process. This keeps a native GTK/Pango
            # problem from freezing the GUI and lets us enforce a hard timeout.
            command = [sys.executable, "--generate", donor]
            try:
                self.after(0, lambda: self.status.set("Leyendo datos y preparando el PDF..."))
                result = subprocess.run(
                    command,
                    env=env,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=120,
                    cwd=str(RESOURCE_DIR),
                )
                if result.returncode != 0:
                    details = (result.stderr or result.stdout or "El generador terminó con un error.").strip()
                    raise RuntimeError(details[-4000:])

                # The worker prints the absolute PDF path on its final line.
                pdf_path = None
                for line in reversed(result.stdout.splitlines()):
                    candidate = Path(line.strip())
                    if candidate.suffix.lower() == ".pdf" and candidate.exists():
                        pdf_path = candidate
                        break
                if pdf_path is None:
                    raise RuntimeError("El PDF no fue creado.\n\n" + (result.stdout or ""))

                self.after(0, lambda p=pdf_path: self._generation_success(p))
            except subprocess.TimeoutExpired:
                self.after(0, lambda: self._generation_timeout())
            except Exception as exc:
                details = traceback.format_exc()
                self.after(0, lambda e=exc, d=details: self._generation_failed(e, d))

        threading.Thread(target=worker, daemon=True).start()

    def _generation_timeout(self) -> None:
        self.generate_button.configure(state="normal")
        self.status.set("La generación tardó demasiado")
        messagebox.showerror(
            APP_NAME,
            "La generación del PDF tardó más de 2 minutos y fue detenida.\n\n"
            "Esto evita que la aplicación quede congelada. Si vuelve a ocurrir, "
            "el problema está en la configuración de PDF/WeasyPrint del equipo.",
        )

    def _generation_success(self, pdf_path: Path) -> None:
        self.generate_button.configure(state="normal")
        self.last_pdf = Path(pdf_path)
        self.status.set("Reporte generado correctamente")
        answer = messagebox.askyesno(
            APP_NAME,
            f"Reporte generado correctamente.\n\n{pdf_path.name}\n\n¿Abrir el PDF ahora?",
        )
        if answer:
            self._open_path(pdf_path)

    def _generation_failed(self, exc: Exception, details: str) -> None:
        self.generate_button.configure(state="normal")
        self.status.set("No se pudo generar el reporte")
        # Keep the user-facing error short, but include useful detail.
        messagebox.showerror(APP_NAME, f"No se pudo generar el reporte.\n\n{exc}")
        print(details)

    # ---------------------------------------------------------
    # Folders / files
    # ---------------------------------------------------------

    def _open_output(self) -> None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        self._open_path(OUTPUT_DIR)

    def _open_last_pdf(self) -> None:
        if self.last_pdf and self.last_pdf.exists():
            self._open_path(self.last_pdf)
            return
        messagebox.showinfo(APP_NAME, "Todavía no se ha generado un reporte en esta sesión.")

    @staticmethod
    def _open_path(path: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", str(path)])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"No se pudo abrir:\n{path}\n\n{exc}")


def run_worker(donor: str) -> int:
    """Worker entry point used by the packaged EXE."""
    os.environ["BAP_REPORT_DATA_DIR"] = os.environ.get("BAP_REPORT_DATA_DIR", str(APP_DATA_DIR))
    os.environ["BAP_REPORT_OUTPUT_DIR"] = os.environ.get("BAP_REPORT_OUTPUT_DIR", str(OUTPUT_DIR))
    generator = load_generator()
    pdf_path = generator.generate_report_for_donor(donor)
    print(str(pdf_path), flush=True)
    return 0


def main() -> None:
    if len(sys.argv) >= 3 and sys.argv[1] == "--generate":
        raise SystemExit(run_worker(sys.argv[2]))

    app = ReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()
