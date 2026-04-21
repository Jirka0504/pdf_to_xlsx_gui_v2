import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font


APP_TITLE = "PDF → XLSX GUI v2 (ASWO + OMNIA)"


HEADERS = [
    "Interní kód zboží",
    "Název",
    "Množství",
    "Cena celkem",
    "Zkrácená poznámka",
    "Kód kombinované nomenklatury",
    "Země původu",
    "Hmotnost",
]


def normalize_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = text.replace("￾", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def cz_number_to_float(value: str) -> float:
    value = value.replace(" ", "").replace("\xa0", "").replace(".", "").replace(",", ".")
    return float(value)


def en_number_to_float(value: str) -> float:
    value = value.replace(" ", "").replace("\xa0", "").replace(",", "")
    return float(value)


def float_to_cz_string(value: float) -> str:
    s = f"{value:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", "\xa0")
    return s


def extract_pdf_lines(pdf_path: str):
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw_line in text.splitlines():
                line = normalize_text(raw_line)
                if line:
                    lines.append(line)
    return lines


def detect_supplier(lines: list[str]) -> str:
    text = "\n".join(lines).lower()

    if "aswo czech s.r.o." in text or "číslo produktu" in text or "dodávané množství" in text:
        return "ASWO"

    if "omnia components" in text or "product code description quantity prezzo" in text:
        return "OMNIA"

    return "UNKNOWN"


# =========================
# ASWO PARSER
# =========================

def looks_like_split_code_only(line: str) -> bool:
    line = normalize_text(line)
    return bool(re.fullmatch(r"[A-Z0-9\-_/\.]+", line))


def parse_aswo(lines: list[str]):
    item_main_re = re.compile(
        r"""
        ^
        (?P<code>[A-Z0-9\-_/\.]+)\s+
        (?P<name>.+?)\s+
        (?P<ordered>\d+)\s+
        (?P<delivered>\d+)\s+
        (?P<unit_sale>\d{1,3}(?:\.\d{3})*,\d{2})\s+
        (?P<unit_net>\d{1,3}(?:\.\d{3})*,\d{2})\s+
        (?P<vat>\d{1,2})\s+
        (?P<total>\d{1,3}(?:\.\d{3})*,\d{2})
        $
        """,
        re.VERBOSE,
    )

    details_re = re.compile(
        r"""
        ^
        Celní\ kód\ zboží:\s*(?P<customs>\d+)\s+
        Hmotnost\ v\ gramech:\s*(?P<weight>\d+)\s+
        Původ:\s*(?P<origin>[A-Z]{2,3})
        $
        """,
        re.VERBOSE,
    )

    job_re = re.compile(r"^Číslo zakázky:\s*(?P<job>.+)$")
    logistics_re = re.compile(r"^Plus poměrné logistické náklady\s*(?P<log>\d{1,3}(?:\.\d{3})*,\d{2})\s*CZK$")

    skip_prefixes = (
        "Stránka:",
        "Objednávka:",
        "Zákaznické číslo:",
        "Datum vystavení",
        "dokladu / DUZP",
        "Číslo dokladu:",
        "DIČ:",
        "Údaje o vaší objednávce:",
        "Objednávka z:",
        "Jméno zákazníka:",
        "Číslo objednávky:",
        "Poznámka k dodávce",
        "Dodací adresa:",
        "ASWO Czech s.r.o.",
        "Na Prosecké vyhlídce",
        "info@aswo.cz",
        "Výkonný ředitel:",
        "Krajský soud:",
        "Místo plnění a soudní příslušnost:",
        "Commerzbank AG",
        "Číslo účtu:",
        "Kód banky:",
        "Číslo produktu",
        "Objednané množství",
        "Dodávané množství",
        "Celková cena",
        "Označení produktu",
        "Prodejní cena za jednotku",
        "Čistá cena za jednotku",
        "Sazba daně %",
        "Čistá hodnota celkem",
        "Náklady na dopravu",
        "Hodnota zboží",
        "Příplatek za velký díl",
    )

    items = []
    warnings = []

    current_item = None
    pending_code_fragment = None

    def push_current():
        nonlocal current_item
        if current_item:
            items.append(current_item)
            current_item = None

    for line in lines:
        if any(line.startswith(prefix) for prefix in skip_prefixes):
            continue

        if line in ("KTS-AME S.R.O.", "Karla Capka 60", "50002 Hradec Kralove", "CZECH REPUBLIC", "Jiří Krejčí"):
            continue

        if pending_code_fragment:
            test_line = f"{pending_code_fragment} {line}"
            m = item_main_re.match(test_line)
            if m:
                push_current()
                current_item = {
                    "code": m.group("code").strip(),
                    "name": m.group("name").strip(),
                    "qty": int(m.group("delivered")),
                    "total_base": cz_number_to_float(m.group("total")),
                    "logistics": 0.0,
                    "job": "",
                    "customs": "",
                    "origin": "",
                    "weight": "",
                }
                pending_code_fragment = None
                continue

        m = item_main_re.match(line)
        if m:
            push_current()
            current_item = {
                "code": m.group("code").strip(),
                "name": m.group("name").strip(),
                "qty": int(m.group("delivered")),
                "total_base": cz_number_to_float(m.group("total")),
                "logistics": 0.0,
                "job": "",
                "customs": "",
                "origin": "",
                "weight": "",
            }
            continue

        if looks_like_split_code_only(line):
            pending_code_fragment = line
            continue

        if current_item:
            m2 = details_re.match(line)
            if m2:
                current_item["customs"] = m2.group("customs").strip()
                current_item["origin"] = ""
                current_item["weight"] = int(m2.group("weight"))
                continue

            m3 = job_re.match(line)
            if m3:
                current_item["job"] = m3.group("job").strip()
                continue

            m4 = logistics_re.match(line)
            if m4:
                current_item["logistics"] = cz_number_to_float(m4.group("log"))
                continue

    push_current()

    result = []
    for item in items:
        total_with_logistics = item["total_base"] + item["logistics"]
        result.append({
            "Interní kód zboží": item["code"],
            "Název": item["name"],
            "Množství": item["qty"],
            "Cena celkem": total_with_logistics,
            "Zkrácená poznámka": item["job"],
            "Kód kombinované nomenklatury": item["customs"],
            "Země původu": item["origin"],
            "Hmotnost": item["weight"],
        })

    return result, warnings


# =========================
# OMNIA PARSER
# =========================

def parse_omnia(lines: list[str]):
    items = []
    warnings = []

    full_row_re = re.compile(
        r"""
        ^
        (?P<code>[A-Z0-9\-.]+)\s+
        (?P<name>.+?)\s+
        (?P<qty>\d+)\s+PZ\s+
        (?P<price>\d+(?:\.\d{2}))\s+€\s+
        (?P<total>\d+(?:\.\d{2}))\s+€
        $
        """,
        re.VERBOSE,
    )

    split_row_re = re.compile(
        r"""
        ^
        (?P<name>.+?)\s*-\s+
        (?P<tail_code>[A-Z0-9\-.]+)\s+
        (?P<qty>\d+)\s+PZ\s+
        (?P<price>\d+(?:\.\d{2}))\s+€\s+
        (?P<total>\d+(?:\.\d{2}))\s+€
        $
        """,
        re.VERBOSE,
    )

    skip_prefixes = (
        "Consegnare a:",
        "Spettabile:",
        "Vat code:",
        "TOTALE MERCE",
        "SCONTO %",
        "SPESE INCASSO",
        "TOTALE IMPONIBILE",
        "TOTALE IMPOSTA",
        "IMPONIBILE",
        "ALIQUOTA IVA",
        "TOTALE DOCUMENTO",
        "OMAGGI",
        "TOTALE DA PAGARE",
        "FATTURA",
        "OMNIA COMPONENTS",
        "Via Travnik",
        "Tel.",
        "C.F. E P.IVA",
        "Capitale Sociale",
        "N.Documento",
        "Data spedizione richiesta:",
        "http://www.omniacomponents.com",
        "PRODUCT CODE DESCRIPTION QUANTITY PREZZO SCONTO % IMPORTO TOTALE",
        "Spese di Trasporto UE Vendita - Shipping Fees",
    )

    pending_prefix = None

    for raw_line in lines:
        line = normalize_text(raw_line).replace("￾", "")

        if any(line.startswith(prefix) for prefix in skip_prefixes):
            continue

        if line.startswith("KTS - AME") or line.startswith("Karla Čapka") or line.startswith("500 02"):
            continue

        # 1) běžný kompletní řádek
        m = full_row_re.match(line)
        if m:
            items.append({
                "Interní kód zboží": m.group("code").strip(),
                "Název": m.group("name").strip(),
                "Množství": int(m.group("qty")),
                "Cena celkem": en_number_to_float(m.group("total")),
                "Zkrácená poznámka": "",
                "Kód kombinované nomenklatury": "",
                "Země původu": "",
                "Hmotnost": "",
            })
            pending_prefix = None
            continue

        # 2) případ slepeného prefixu + pokračování kódu, např. VEN149198350
        # chceme z toho udělat prefix VEN-
        m_prefix = re.fullmatch(r"([A-Z]{2,10})(\d{5,})", line)
        if m_prefix:
            pending_prefix = m_prefix.group(1).strip() + "-"
            continue

        # 3) případ samostatného prefixu zakončeného pomlčkou, např. VEN-
        if re.fullmatch(r"[A-Z0-9]+-$", line):
            pending_prefix = line.strip()
            continue

        # 4) pokud čekáme na pokračování, vezmeme další řádek
        if pending_prefix:
            m2 = split_row_re.match(line)
            if m2:
                full_code = pending_prefix + m2.group("tail_code").strip()
                items.append({
                    "Interní kód zboží": full_code,
                    "Název": m2.group("name").strip(),
                    "Množství": int(m2.group("qty")),
                    "Cena celkem": en_number_to_float(m2.group("total")),
                    "Zkrácená poznámka": "",
                    "Kód kombinované nomenklatury": "",
                    "Země původu": "",
                    "Hmotnost": "",
                })
                pending_prefix = None
                continue

            # fallback: kdyby další řádek nezačínal názvem s pomlčkou
            m3 = re.match(
                r"""
                ^
                (?P<tail_code>[A-Z0-9\-.]+)\s+
                (?P<name>.+?)\s+
                (?P<qty>\d+)\s+PZ\s+
                (?P<price>\d+(?:\.\d{2}))\s+€\s+
                (?P<total>\d+(?:\.\d{2}))\s+€
                $
                """,
                line,
                re.VERBOSE,
            )
            if m3:
                full_code = pending_prefix + m3.group("tail_code").strip()
                items.append({
                    "Interní kód zboží": full_code,
                    "Název": m3.group("name").strip(),
                    "Množství": int(m3.group("qty")),
                    "Cena celkem": en_number_to_float(m3.group("total")),
                    "Zkrácená poznámka": "",
                    "Kód kombinované nomenklatury": "",
                    "Země původu": "",
                    "Hmotnost": "",
                })
                pending_prefix = None
                continue

        # neparsované řádky jen ignorujeme

    return items, warnings

# =========================
# SAVE XLSX
# =========================

def save_xlsx(output_path: str, rows: list[dict]):
    wb = Workbook()
    ws = wb.active
    ws.title = "List1"

    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)

    row_no = 2
    for row in rows:
        ws.cell(row=row_no, column=1, value=row["Interní kód zboží"])
        ws.cell(row=row_no, column=2, value=row["Název"])
        ws.cell(row=row_no, column=3, value=row["Množství"])
        ws.cell(row=row_no, column=4, value=float_to_cz_string(row["Cena celkem"]))
        ws.cell(row=row_no, column=5, value=row["Zkrácená poznámka"])
        ws.cell(row=row_no, column=6, value=int(row["Kód kombinované nomenklatury"]) if str(row["Kód kombinované nomenklatury"]).isdigit() else row["Kód kombinované nomenklatury"])
        ws.cell(row=row_no, column=7, value=row["Země původu"])
        ws.cell(row=row_no, column=8, value=row["Hmotnost"])
        row_no += 1

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 24
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 12

    wb.save(output_path)


# =========================
# GUI
# =========================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x640")

        self.pdf_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.detected_supplier = tk.StringVar(value="Nezjištěno")

        self.preview_box = None
        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Vstupní PDF:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.pdf_path, width=95).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(frm, text="Vybrat…", command=self.pick_pdf).grid(row=0, column=2, padx=6)

        ttk.Label(frm, text="Výstupní XLSX:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.output_path, width=95).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(frm, text="Uložit jako…", command=self.pick_output).grid(row=1, column=2, padx=6)

        ttk.Label(frm, text="Detekovaný typ PDF:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.detected_supplier, width=30, state="readonly").grid(row=2, column=1, sticky="w", pady=4)

        ttk.Separator(frm, orient="horizontal").grid(row=3, column=0, columnspan=3, sticky="ew", pady=10)

        note = (
            "Verze 2:\n"
            "• Automatická detekce ASWO / OMNIA\n"
            "• ASWO: množství = dodávané množství\n"
            "• ASWO: cena celkem = celková cena + poměrné logistické náklady\n"
            "• OMNIA: umí i rozdělený kód na dalším řádku"
        )
        ttk.Label(frm, text=note, foreground="#555", justify="left").grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 10))

        self.preview_box = tk.Text(frm, height=22, wrap="none")
        self.preview_box.grid(row=5, column=0, columnspan=3, sticky="nsew")

        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=3, sticky="e", pady=12)
        ttk.Button(btns, text="Detekovat", command=self.detect_only).pack(side="left", padx=4)
        ttk.Button(btns, text="Náhled", command=self.preview).pack(side="left", padx=4)
        ttk.Button(btns, text="Převést", command=self.convert).pack(side="left", padx=4)
        ttk.Button(btns, text="Konec", command=self.destroy).pack(side="left", padx=4)

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(5, weight=1)

    def pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if path:
            self.pdf_path.set(path)
            if not self.output_path.get().strip():
                self.output_path.set(str(Path(path).with_suffix(".xlsx")))

    def pick_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.output_path.set(path)

    def _load_and_detect(self):
        pdf = self.pdf_path.get().strip()
        if not pdf:
            raise ValueError("Vyber vstupní PDF.")

        lines = extract_pdf_lines(pdf)
        supplier = detect_supplier(lines)
        self.detected_supplier.set(supplier)
        return lines, supplier

    def detect_only(self):
        try:
            _, supplier = self._load_and_detect()
        except Exception as e:
            messagebox.showerror("Chyba", str(e))
            return
        messagebox.showinfo("Detekce", f"Rozpoznaný typ PDF: {supplier}")

    def _parse(self):
        lines, supplier = self._load_and_detect()

        if supplier == "ASWO":
            return parse_aswo(lines), supplier
        if supplier == "OMNIA":
            return parse_omnia(lines), supplier

        raise ValueError("Nepodařilo se rozpoznat typ PDF. Tento dokument zatím není podporovaný.")

    def preview(self):
        try:
            (rows, warnings), supplier = self._parse()
        except Exception as e:
            messagebox.showerror("Chyba", str(e))
            return

        self.preview_box.delete("1.0", "end")
        self.preview_box.insert("1.0", f"Typ PDF: {supplier}\n")
        self.preview_box.insert("end", f"Nalezeno položek: {len(rows)}\n\n")

        for row in rows[:40]:
            self.preview_box.insert(
                "end",
                f"{row['Interní kód zboží']} | {row['Název']} | qty={row['Množství']} | total={float_to_cz_string(row['Cena celkem'])}\n"
            )

        if warnings:
            self.preview_box.insert("end", "\n--- VAROVÁNÍ ---\n")
            for w in warnings[:20]:
                self.preview_box.insert("end", w + "\n")

    def convert(self):
        out = self.output_path.get().strip()
        if not out:
            messagebox.showerror("Chyba", "Vyber výstupní XLSX.")
            return

        try:
            (rows, warnings), supplier = self._parse()
            save_xlsx(out, rows)
        except Exception as e:
            messagebox.showerror("Chyba", str(e))
            return

        msg = f"Hotovo.\nTyp PDF: {supplier}\nUloženo do:\n{out}\n\nPočet položek: {len(rows)}"
        if warnings:
            msg += f"\nVarování: {len(warnings)}"
        messagebox.showinfo("Hotovo", msg)


if __name__ == "__main__":
    App().mainloop()
