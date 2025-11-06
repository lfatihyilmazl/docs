"""Wattaw A.\u015e. kartvizit ve m\u00fc\u015fteri g\u00f6r\u00fc\u015fmesi y\u00f6netim uygulamas\u0131."""
from __future__ import annotations

import importlib
import io
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, Optional


REQUIRED_PACKAGES = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "PIL": "Pillow",
    "cairosvg": "cairosvg",
}


def ensure_dependencies() -> None:
    """Install or upgrade dependencies dynamically when the app starts."""
    for module_name, package_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(module_name)
        except ImportError:  # pragma: no cover - network dependent
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "--upgrade", package_name]
            )


ensure_dependencies()

import pandas as pd  # type: ignore  # noqa: E402
from PIL import Image, ImageTk  # type: ignore  # noqa: E402
import cairosvg  # type: ignore  # noqa: E402
import tkinter as tk  # noqa: E402
from tkinter import filedialog, messagebox, ttk  # noqa: E402


DATA_COLUMNS = [
    "Kay\u0131t Zaman\u0131",
    "S\u0131rket",
    "K\u0131\u015fi",
    "\u00dcnvan",
    "Sekt\u00f6r",
    "E-posta",
    "Telefon",
    "Adres",
    "Kart Notu",
    "Kart Dosyas\u0131",
]

MEETING_COLUMNS = [
    "Kay\u0131t Zaman\u0131",
    "S\u0131rket",
    "K\u0131\u015fi",
    "G\u00f6r\u00fc\u015fme Tarihi",
    "G\u00f6r\u00fc\u015fme Saati",
    "Konu",
    "Notlar",
]


class BusinessCardApp:
    """Ana Tkinter uygulamas\u0131."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Wattaw A.\u015e. Kartvizit Y\u00f6neticisi")
        self.root.geometry("1080x720")
        self.root.minsize(960, 640)

        self.base_dir = Path(__file__).resolve().parent
        self.data_dir = self.base_dir / "data"
        self.cards_dir = self.data_dir / "business_cards"
        self.meetings_file = self.data_dir / "meetings.xlsx"
        self.customers_file = self.data_dir / "customers.xlsx"
        self.assets_dir = self.base_dir / "assets"

        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.cards_dir.mkdir(parents=True, exist_ok=True)

        self.cards_df = self._load_dataframe(self.customers_file, DATA_COLUMNS)
        self.meetings_df = self._load_dataframe(self.meetings_file, MEETING_COLUMNS)

        self.logo_images: Dict[str, ImageTk.PhotoImage] = {}
        self.selected_card_path: Optional[Path] = None

        self._load_logos()
        self._build_ui()

    # ------------------------------------------------------------------
    # Data helpers
    # ------------------------------------------------------------------
    def _load_dataframe(self, file_path: Path, columns: Iterable[str]) -> pd.DataFrame:
        if file_path.exists():
            try:
                df = pd.read_excel(file_path)
                missing = [col for col in columns if col not in df.columns]
                for column in missing:
                    df[column] = ""
                return df[columns]
            except Exception as exc:  # pragma: no cover - defensive
                messagebox.showerror(
                    "Dosya okunamad\u0131",
                    f"{file_path.name} dosyas\u0131 okunurken hata olu\u015ftu: {exc}",
                )
        return pd.DataFrame(columns=list(columns))

    def _save_dataframe(self, df: pd.DataFrame, file_path: Path) -> None:
        df.to_excel(file_path, index=False)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _load_logos(self) -> None:
        for name, filename, width in [
            ("wordmark", "wattaw_wordmark.svg", 380),
            ("icon", "wattaw_mark.svg", 80),
        ]:
            path = self.assets_dir / filename
            try:
                png_bytes = cairosvg.svg2png(url=str(path), output_width=width)
                image = Image.open(io.BytesIO(png_bytes))
                self.logo_images[name] = ImageTk.PhotoImage(image)
            except Exception:
                # If anything goes wrong we just skip the logo but keep the UI running.
                continue

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=20)
        container.pack(fill=tk.BOTH, expand=True)

        logo_frame = ttk.Frame(container)
        logo_frame.pack(fill=tk.X, pady=(0, 16))

        if "icon" in self.logo_images:
            ttk.Label(logo_frame, image=self.logo_images["icon"]).pack(side=tk.LEFT)
        ttk.Label(
            logo_frame,
            text="Wattaw A.\u015e. Kartvizit Tarama ve G\u00f6r\u00fc\u015fme Takibi",
            font=("Segoe UI", 18, "bold"),
        ).pack(side=tk.LEFT, padx=16)
        if "wordmark" in self.logo_images:
            ttk.Label(logo_frame, image=self.logo_images["wordmark"]).pack(side=tk.RIGHT)

        notebook = ttk.Notebook(container)
        notebook.pack(fill=tk.BOTH, expand=True)

        self.card_tab = ttk.Frame(notebook, padding=12)
        self.meeting_tab = ttk.Frame(notebook, padding=12)

        notebook.add(self.card_tab, text="Kartvizit Kayd\u0131")
        notebook.add(self.meeting_tab, text="G\u00f6r\u00fc\u015fme Notlar\u0131")

        self._build_card_tab()
        self._build_meeting_tab()

    def _build_card_tab(self) -> None:
        form_frame = ttk.Frame(self.card_tab)
        form_frame.pack(side=tk.TOP, fill=tk.X)

        fields = [
            ("\u015eirket Ad\u0131", "company"),
            ("\u0130lgili K\u0131\u015fi", "contact"),
            ("\u00dcnvan", "title"),
            ("Sekt\u00f6r", "sector"),
            ("E-posta", "email"),
            ("Telefon", "phone"),
            ("Adres", "address"),
        ]

        self.card_entries: Dict[str, tk.Widget] = {}

        for row, (label_text, key) in enumerate(fields):
            ttk.Label(form_frame, text=label_text).grid(row=row, column=0, sticky=tk.W, pady=4)
            if key == "sector":
                combo = ttk.Combobox(
                    form_frame,
                    values=[
                        "Demir-\u00c7elik",
                        "Enerji",
                        "Makine",
                        "M\u00fchendislik",
                        "Lojistik",
                        "Di\u011fer",
                    ],
                    state="readonly",
                )
                combo.grid(row=row, column=1, sticky=tk.EW, pady=4, padx=8)
                combo.set("Demir-\u00c7elik")
                self.card_entries[key] = combo
            else:
                entry = ttk.Entry(form_frame)
                entry.grid(row=row, column=1, sticky=tk.EW, pady=4, padx=8)
                self.card_entries[key] = entry

        form_frame.columnconfigure(1, weight=1)

        ttk.Label(form_frame, text="Karttan Notlar").grid(
            row=len(fields), column=0, sticky=tk.NW, pady=4
        )
        self.card_notes = tk.Text(form_frame, height=4)
        self.card_notes.grid(row=len(fields), column=1, sticky=tk.EW, pady=4, padx=8)

        file_frame = ttk.Frame(self.card_tab)
        file_frame.pack(fill=tk.X, pady=(12, 4))

        self.card_file_var = tk.StringVar(value="Kartvizit dosyas\u0131 se\u00e7ilmedi")
        ttk.Button(file_frame, text="Kartvizit Dosyas\u0131 Se\u00e7", command=self._select_card_file).pack(
            side=tk.LEFT
        )
        ttk.Label(file_frame, textvariable=self.card_file_var).pack(side=tk.LEFT, padx=8)

        action_frame = ttk.Frame(self.card_tab)
        action_frame.pack(fill=tk.X, pady=(4, 12))

        ttk.Button(action_frame, text="Kaydet", command=self._save_card).pack(side=tk.RIGHT)

        self.card_tree = ttk.Treeview(
            self.card_tab,
            columns=DATA_COLUMNS,
            show="headings",
            height=12,
        )
        for column in DATA_COLUMNS:
            self.card_tree.heading(column, text=column)
            self.card_tree.column(column, stretch=True, width=120)
        self.card_tree.pack(fill=tk.BOTH, expand=True)

        self._refresh_card_tree()

    def _build_meeting_tab(self) -> None:
        form_frame = ttk.Frame(self.meeting_tab)
        form_frame.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(form_frame, text="\u015eirket").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.meeting_company = ttk.Combobox(form_frame, values=self._company_list(), state="readonly")
        self.meeting_company.grid(row=0, column=1, sticky=tk.EW, pady=4, padx=8)

        ttk.Label(form_frame, text="K\u0131\u015fi").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.meeting_contact = ttk.Entry(form_frame)
        self.meeting_contact.grid(row=1, column=1, sticky=tk.EW, pady=4, padx=8)

        ttk.Label(form_frame, text="G\u00f6r\u00fc\u015fme Tarihi").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.meeting_date = ttk.Entry(form_frame)
        self.meeting_date.insert(0, datetime.now().strftime("%d.%m.%Y"))
        self.meeting_date.grid(row=2, column=1, sticky=tk.EW, pady=4, padx=8)

        ttk.Label(form_frame, text="G\u00f6r\u00fc\u015fme Saati").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.meeting_time = ttk.Entry(form_frame)
        self.meeting_time.insert(0, datetime.now().strftime("%H:%M"))
        self.meeting_time.grid(row=3, column=1, sticky=tk.EW, pady=4, padx=8)

        ttk.Label(form_frame, text="Konu").grid(row=4, column=0, sticky=tk.W, pady=4)
        self.meeting_subject = ttk.Entry(form_frame)
        self.meeting_subject.grid(row=4, column=1, sticky=tk.EW, pady=4, padx=8)

        ttk.Label(form_frame, text="Notlar").grid(row=5, column=0, sticky=tk.NW, pady=4)
        self.meeting_notes = tk.Text(form_frame, height=4)
        self.meeting_notes.grid(row=5, column=1, sticky=tk.EW, pady=4, padx=8)

        form_frame.columnconfigure(1, weight=1)

        ttk.Button(self.meeting_tab, text="G\u00f6r\u00fc\u015fmeyi Kaydet", command=self._save_meeting).pack(
            anchor=tk.E, pady=(8, 12)
        )

        self.meeting_tree = ttk.Treeview(
            self.meeting_tab,
            columns=MEETING_COLUMNS,
            show="headings",
            height=12,
        )
        for column in MEETING_COLUMNS:
            self.meeting_tree.heading(column, text=column)
            self.meeting_tree.column(column, stretch=True, width=140)
        self.meeting_tree.pack(fill=tk.BOTH, expand=True)

        self._refresh_meeting_tree()

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------
    def _select_card_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Kartvizit dosyas\u0131 se\u00e7in",
            filetypes=[("Resim dosyalar\u0131", "*.png *.jpg *.jpeg *.svg *.webp *.bmp")],
        )
        if file_path:
            self.selected_card_path = Path(file_path)
            self.card_file_var.set(self.selected_card_path.name)

    def _save_card(self) -> None:
        data = {key: self._get_widget_value(widget) for key, widget in self.card_entries.items()}
        data["notes"] = self.card_notes.get("1.0", tk.END).strip()

        if not data["company"]:
            messagebox.showwarning("Eksik bilgi", "\u015eirket ad\u0131 zorunludur.")
            return
        if not data["contact"]:
            messagebox.showwarning("Eksik bilgi", "\u0130lgili ki\u015fi alan\u0131 zorunludur.")
            return

        timestamp = datetime.now().isoformat(timespec="seconds")
        stored_path = ""
        if self.selected_card_path and self.selected_card_path.exists():
            destination = self.cards_dir / f"{timestamp.replace(':', '-')}_{self.selected_card_path.name}"
            try:
                shutil.copy2(self.selected_card_path, destination)
                stored_path = destination.relative_to(self.base_dir).as_posix()
            except Exception as exc:
                messagebox.showerror("Dosya kopyalanamad\u0131", f"Kart dosyas\u0131 kaydedilemedi: {exc}")
                return

        record = {
            "Kay\u0131t Zaman\u0131": timestamp,
            "S\u0131rket": data["company"],
            "K\u0131\u015fi": data["contact"],
            "\u00dcnvan": data["title"],
            "Sekt\u00f6r": data["sector"],
            "E-posta": data["email"],
            "Telefon": data["phone"],
            "Adres": data["address"],
            "Kart Notu": data["notes"],
            "Kart Dosyas\u0131": stored_path,
        }

        self.cards_df = pd.concat([self.cards_df, pd.DataFrame([record])], ignore_index=True)
        self._save_dataframe(self.cards_df, self.customers_file)
        self._refresh_card_tree()
        self._update_meeting_company_options()

        self._clear_card_form()
        messagebox.showinfo("Kaydedildi", "Kartvizit kayd\u0131 ba\u015far\u0131yla olu\u015fturuldu.")

    def _save_meeting(self) -> None:
        company = self.meeting_company.get()
        contact = self.meeting_contact.get().strip()
        meeting_date = self.meeting_date.get().strip()
        meeting_time = self.meeting_time.get().strip()
        subject = self.meeting_subject.get().strip()
        notes = self.meeting_notes.get("1.0", tk.END).strip()

        if not company:
            messagebox.showwarning("Eksik bilgi", "\u015eirket se\u00e7melisiniz.")
            return
        if not contact:
            messagebox.showwarning("Eksik bilgi", "G\u00f6r\u00fc\u015fme yap\u0131lan ki\u015fiyi yazmal\u0131s\u0131n\u0131z.")
            return
        if not meeting_date:
            messagebox.showwarning("Eksik bilgi", "G\u00f6r\u00fc\u015fme tarihini giriniz.")
            return

        timestamp = datetime.now().isoformat(timespec="seconds")
        record = {
            "Kay\u0131t Zaman\u0131": timestamp,
            "S\u0131rket": company,
            "K\u0131\u015fi": contact,
            "G\u00f6r\u00fc\u015fme Tarihi": meeting_date,
            "G\u00f6r\u00fc\u015fme Saati": meeting_time,
            "Konu": subject,
            "Notlar": notes,
        }

        self.meetings_df = pd.concat([self.meetings_df, pd.DataFrame([record])], ignore_index=True)
        self._save_dataframe(self.meetings_df, self.meetings_file)
        self._refresh_meeting_tree()

        self.meeting_subject.delete(0, tk.END)
        self.meeting_notes.delete("1.0", tk.END)

        messagebox.showinfo("Kaydedildi", "G\u00f6r\u00fc\u015fme kayd\u0131 olu\u015fturuldu.")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _get_widget_value(self, widget: tk.Widget) -> str:
        if isinstance(widget, ttk.Combobox):
            return widget.get().strip()
        if isinstance(widget, ttk.Entry):
            return widget.get().strip()
        return ""

    def _clear_card_form(self) -> None:
        for key, widget in self.card_entries.items():
            if isinstance(widget, ttk.Combobox):
                widget.set(widget["values"][0] if widget["values"] else "")
            elif isinstance(widget, ttk.Entry):
                widget.delete(0, tk.END)
        self.card_notes.delete("1.0", tk.END)
        self.card_file_var.set("Kartvizit dosyas\u0131 se\u00e7ilmedi")
        self.selected_card_path = None

    def _refresh_card_tree(self) -> None:
        self.card_tree.delete(*self.card_tree.get_children())
        for _, row in self.cards_df.iterrows():
            self.card_tree.insert("", tk.END, values=[row[col] for col in DATA_COLUMNS])

    def _refresh_meeting_tree(self) -> None:
        self.meeting_tree.delete(*self.meeting_tree.get_children())
        for _, row in self.meetings_df.iterrows():
            self.meeting_tree.insert("", tk.END, values=[row[col] for col in MEETING_COLUMNS])

    def _company_list(self) -> Iterable[str]:
        companies = sorted(set(self.cards_df["S\u0131rket"].dropna().tolist()))
        return [c for c in companies if c]

    def _update_meeting_company_options(self) -> None:
        companies = self._company_list()
        self.meeting_company["values"] = companies
        if companies and not self.meeting_company.get():
            self.meeting_company.set(companies[-1])


def main() -> None:
    root = tk.Tk()
    app = BusinessCardApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
