from __future__ import annotations

import csv
import os
import random
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook

import hashlib


ROOT = Path(__file__).resolve().parent
GENERATED_DIR = ROOT / "generated_files"
HISTORY_FILE = ROOT / "history.csv"
TEMPLATE_FILE = ROOT / "December2025.xlsx"

HISTORY_FIELDS = [
    "file", "datetime", "sheet1_total", "sheet2_d_total",
    "sheet2_e_total", "sheet2_f_total", "sheet2_g_total",
    "sheet3_total", "month_year",
]
CHARTS = [
    GENERATED_DIR / "sheet1_total_chart.png",
    GENERATED_DIR / "sheet2_totals_chart.png",
    GENERATED_DIR / "sheet3_total_chart.png",
]


def read_history() -> list[dict]:
    if not HISTORY_FILE.exists():
        return []
    with HISTORY_FILE.open("r", newline="", encoding="utf-8-sig") as handle:
        return list(csv.DictReader(handle))


def number(value, default=0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def latest_source(history: list[dict]) -> Path:
    # The history pointer is authoritative; filesystem timestamps are not.
    for record in reversed(history):
        filename = (record.get("file") or "").strip()
        if filename:
            candidate = GENERATED_DIR / Path(filename).name
            if candidate.is_file():
                return candidate
    return TEMPLATE_FILE


def scale_integer_values(values: list[float], target_total: float,
                         rng: random.Random, *, positive: bool = False,
                         preserve_zeros: bool = True) -> list[int]:
    """Randomize row values while keeping the column total close to target_total."""
    if not values:
        return []
    active = [i for i, value in enumerate(values)
              if not (preserve_zeros and value == 0)]
    result = [0] * len(values)
    if not active:
        return result

    weights = []
    for i in active:
        base = max(0.0, values[i])
        weights.append(max(0.01, base) * rng.uniform(0.90, 1.10))
    weight_sum = sum(weights)
    if weight_sum <= 0:
        weights = [1.0] * len(active)
        weight_sum = float(len(active))

    target = max(len(active) if positive else 0, int(round(target_total)))
    allocations = [int(round(target * weight / weight_sum)) for weight in weights]

    # Correct rounding drift without changing the zero/non-zero pattern.
    difference = target - sum(allocations)
    order = sorted(range(len(active)), key=lambda j: weights[j], reverse=True)
    step = 1 if difference > 0 else -1
    for j in (order * (abs(difference) // len(order) + 1))[:abs(difference)]:
        allocations[j] = max(1 if positive else 0, allocations[j] + step)

    for i, amount in zip(active, allocations):
        result[i] = max(1 if positive else 0, amount)
    return result


def adjust_sheet1(ws, target_total: float, rng: random.Random) -> float:
    # D/E are quantities, G is rate, F and H are formulas.
    rows = list(range(3, 36))
    original_d = [number(ws.cell(row, 4).value) for row in rows]
    original_e = [number(ws.cell(row, 5).value) for row in rows]
    rates = [number(ws.cell(row, 7).value) for row in rows]

    factors = [rng.uniform(0.92, 1.08) for _ in rows]
    raw_d = [max(0, round(v * f)) for v, f in zip(original_d, factors)]
    raw_e = [max(0, round(v * f)) for v, f in zip(original_e, factors)]

    def total(multiplier: float) -> float:
        return sum((round(d * multiplier) + round(e * multiplier)) * rate
                   for d, e, rate in zip(raw_d, raw_e, rates))

    if total(1.0) <= 0:
        raise ValueError("Sheet1 has no usable quantity/rate data.")
    multiplier = target_total / total(1.0)
    for _ in range(4):
        current = total(multiplier)
        if current <= 0:
            break
        multiplier *= target_total / current

    for row, d, e in zip(rows, raw_d, raw_e):
        ws.cell(row, 4).value = max(0, round(d * multiplier))
        ws.cell(row, 5).value = max(0, round(e * multiplier))
        ws.cell(row, 6).value = f"=D{row}+E{row}"
        ws.cell(row, 8).value = f"=F{row}*G{row}"

    calculated = sum(
        (number(ws.cell(row, 4).value) + number(ws.cell(row, 5).value))
        * number(ws.cell(row, 7).value)
        for row in rows
    )
    ws["H36"] = "=SUM(H3:H35)"
    return calculated


def update_sheet2(ws, targets: list[float], rng: random.Random) -> list[int]:
    rows = list(range(2, 197))
    totals = []
    for column, target in zip(range(4, 8), targets):
        values = [number(ws.cell(row, column).value) for row in rows]
        updated = scale_integer_values(
            values, target, rng, positive=False, preserve_zeros=True
        )
        for row, value in zip(rows, updated):
            ws.cell(row, column).value = value
        totals.append(sum(updated))

    for column in range(4, 8):
        letter = chr(64 + column)
        ws.cell(197, column).value = f"=SUM({letter}2:{letter}196)"
    return totals


def update_sheet3(ws, target: float, rng: random.Random) -> int:
    rows = list(range(2, 58))
    values = [number(ws.cell(row, 3).value) for row in rows]
    updated = scale_integer_values(
        values, target, rng, positive=True, preserve_zeros=False
    )
    for row, value in zip(rows, updated):
        ws.cell(row, 3).value = value
    ws["C58"] = "=SUM(C2:C57)"
    return sum(updated)


def write_page_headers(workbook, date_text: str) -> None:
    headers = {
        "Sheet1": f"SUMEET ELECTRONICS\nSTOCK STATEMENT AS ON {date_text}",
        "Sheet2": f"DEALER'S STATEMENT AS ON {date_text}",
        "Sheet3": f"SUNDRY CREDITOR'S LIST AS ON {date_text}",
    }
    for sheet_name, header in headers.items():
        ws = workbook[sheet_name]
        ws.oddHeader.center.text = header
        ws.oddHeader.center.size = 12
        ws.oddHeader.center.bold = True


def set_formula_calculation(workbook) -> None:
    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except AttributeError:
        pass


def write_history(history: list[dict], record: dict) -> None:
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    with HISTORY_FILE.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=HISTORY_FIELDS)
        writer.writeheader()
        for old in history:
            writer.writerow({field: old.get(field, "") for field in HISTORY_FIELDS})
        writer.writerow(record)


def make_charts(history: list[dict]) -> None:
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    records = history
    labels = [r.get("month_year", "") for r in records]

    def save_chart(filename: Path, title: str, series: list[tuple[str, list[float]]]) -> None:
        fig, ax = plt.subplots(figsize=(11, 5.5))
        for label, values in series:
            ax.plot(labels, values, marker="o", linewidth=2, label=label)
        ax.set_title(title)
        ax.set_xlabel("Month / Year")
        ax.set_ylabel("Amount (₹)")
        ax.grid(True, alpha=0.3)
        if len(series) > 1:
            ax.legend()
        ax.tick_params(axis="x", rotation=45)
        fig.tight_layout()
        fig.savefig(filename, dpi=160, format="png")
        plt.close(fig)

    save_chart(
        CHARTS[0], "Sheet1 Stock Statement Total",
        [("Sheet1 Total", [number(r.get("sheet1_total")) for r in records])],
    )
    save_chart(
        CHARTS[1], "Sheet2 Dealer Ageing Totals",
        [
            ("30 days", [number(r.get("sheet2_d_total")) for r in records]),
            ("30–60 days", [number(r.get("sheet2_e_total")) for r in records]),
            ("60–90 days", [number(r.get("sheet2_f_total")) for r in records]),
            ("90+ days", [number(r.get("sheet2_g_total")) for r in records]),
        ],
    )
    save_chart(
        CHARTS[2], "Sheet3 Sundry Creditors Total",
        [("Sheet3 Total", [number(r.get("sheet3_total")) for r in records])],
    )


def fingerprint(value):
    return hashlib.sha256(value.encode()).hexdigest()[:12]
    
def send_email(output_file: Path, totals: dict) -> None:
    sender = (os.getenv("EMAIL_ADDRESS") or "").strip()
    password = (os.getenv("EMAIL_PASSWORD") or "").strip()
    receiver = (os.getenv("RECEIVER_EMAIL") or "").strip()

    # Temporary diagnostics: verify email address secrets are loaded (never log password).
    print(f"EMAIL_ADDRESS loaded: {bool(sender)} | value: {sender or '<empty>'}")
    print(f"RECEIVER_EMAIL loaded: {bool(receiver)} | value: {receiver or '<empty>'}")
    print(f"EMAIL_PASSWORD loaded: {bool(password)}")
    if not all((sender, password, receiver)):
        print("Email skipped: configure EMAIL_ADDRESS, EMAIL_PASSWORD and RECEIVER_EMAIL secrets.")
        return

    message = EmailMessage()
    message["Subject"] = f"Generated Statement - {output_file.stem}"
    message["From"] = sender
    message["To"] = receiver
    message.set_content(
        f"""Statement generated successfully: {output_file.name}

Sheet1 total: ₹{totals['sheet1_total']:,.0f}
Sheet2 totals:
  30 days: ₹{totals['sheet2_d_total']:,.0f}
  30-60 days: ₹{totals['sheet2_e_total']:,.0f}
  60-90 days: ₹{totals['sheet2_f_total']:,.0f}
  90+ days: ₹{totals['sheet2_g_total']:,.0f}
Sheet3 total: ₹{totals['sheet3_total']:,.0f}

The Excel statement and monthly charts are attached.
"""
    )

    attachments = [output_file, *CHARTS]
    for path in attachments:
        mime = "application/octet-stream"
        if path.suffix.lower() == ".xlsx":
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif path.suffix.lower() == ".png":
            mime = "image/png"
        main_type, sub_type = mime.split("/", 1)
        message.add_attachment(
            path.read_bytes(), maintype=main_type, subtype=sub_type,
            filename=path.name
        )

    

    print("Sender fingerprint:", fingerprint(sender))
    print("Password fingerprint:", fingerprint(password))
    
    with smtplib.SMTP("smtp.gmail.com", 587, timeout=60) as smtp:
        smtp.set_debuglevel(1)
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(sender, password)
        smtp.send_message(message)
    print(f"Email sent to {receiver}.")


def main() -> None:
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    history = read_history()
    if not history:
        raise RuntimeError("history.csv is empty or missing; refusing to generate without baseline totals.")

    now = datetime.now()
    # Scheduled runs are month-end only. Manual workflow_dispatch runs are allowed anytime.
    event_name = os.getenv("GITHUB_EVENT_NAME", "")
    if event_name != "workflow_dispatch" and now.day != __import__("calendar").monthrange(now.year, now.month)[1]:
        print("Not the last day of the month; generation skipped.")
        return

    latest = history[-1]
    baseline = {
        "sheet1_total": number(latest.get("sheet1_total")),
        "sheet2_d_total": number(latest.get("sheet2_d_total")),
        "sheet2_e_total": number(latest.get("sheet2_e_total")),
        "sheet2_f_total": number(latest.get("sheet2_f_total")),
        "sheet2_g_total": number(latest.get("sheet2_g_total")),
        "sheet3_total": number(latest.get("sheet3_total")),
    }
    if any(value <= 0 for key, value in baseline.items() if key != "sheet2_g_total"):
        raise RuntimeError("One or more required baseline totals are missing or zero.")

    month_year = os.getenv("OUTPUT_MONTH_YEAR", now.strftime("%B%Y")).strip()
    date_text = now.strftime("%d-%m-%Y")
    timestamp = now.strftime("%d%H%M%S")
    output_file = GENERATED_DIR / f"{month_year}_{timestamp}.xlsx"
    source = latest_source(history)
    if not source.exists():
        raise FileNotFoundError(f"Neither latest generated workbook nor template exists: {source}")

    rng = random.Random(f"{month_year}-{timestamp}")
    targets = {
        key: value * rng.uniform(0.95, 1.05)
        for key, value in baseline.items()
    }

    workbook = load_workbook(source)
    required_sheets = {"Sheet1", "Sheet2", "Sheet3"}
    if not required_sheets.issubset(set(workbook.sheetnames)):
        raise RuntimeError(f"Workbook must contain sheets: {sorted(required_sheets)}")

    sheet1_actual = adjust_sheet1(workbook["Sheet1"], targets["sheet1_total"], rng)
    sheet2_actual = update_sheet2(
        workbook["Sheet2"],
        [targets["sheet2_d_total"], targets["sheet2_e_total"],
         targets["sheet2_f_total"], targets["sheet2_g_total"]],
        rng,
    )
    sheet3_actual = update_sheet3(workbook["Sheet3"], targets["sheet3_total"], rng)

    # Verify every generated total against the latest history baseline.
    checks = [
        ("Sheet1", sheet1_actual, baseline["sheet1_total"]),
        ("Sheet2 D", sheet2_actual[0], baseline["sheet2_d_total"]),
        ("Sheet2 E", sheet2_actual[1], baseline["sheet2_e_total"]),
        ("Sheet2 F", sheet2_actual[2], baseline["sheet2_f_total"]),
        ("Sheet2 G", sheet2_actual[3], baseline["sheet2_g_total"]),
        ("Sheet3", sheet3_actual, baseline["sheet3_total"]),
    ]
    failures = [
        f"{name}: {actual:,.0f} outside ±10% of {base:,.0f}"
        for name, actual, base in checks
        if base > 0 and not (base * 0.90 <= actual <= base * 1.10)
    ]
    if failures:
        raise RuntimeError("Generated totals failed validation: " + "; ".join(failures))

    write_page_headers(workbook, date_text)
    set_formula_calculation(workbook)
    workbook.save(output_file)

    totals = {
        "sheet1_total": round(sheet1_actual),
        "sheet2_d_total": sheet2_actual[0],
        "sheet2_e_total": sheet2_actual[1],
        "sheet2_f_total": sheet2_actual[2],
        "sheet2_g_total": sheet2_actual[3],
        "sheet3_total": sheet3_actual,
    }
    record = {
        "file": output_file.name,
        "datetime": now.strftime("%Y-%m-%d %H:%M:%S"),
        **totals,
        "month_year": month_year,
    }
    updated_history = history + [record]
    write_history(history, record)
    make_charts(updated_history)
    send_email(output_file, totals)

    print(f"Generated: {output_file.relative_to(ROOT)}")
    print("Validated totals:")
    for name, actual, base in checks:
        print(f"  {name}: {actual:,.0f} (baseline {base:,.0f})")


if __name__ == "__main__":
    main()
