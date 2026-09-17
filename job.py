from __future__ import annotations

import os
import smtplib
import sys
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

from collector import download_today_export

LONDON = ZoneInfo("Europe/London")
ACCOUNTS = {
    "RUB01", "RUB02", "RUB03", "SIM08", "SIM09",
    "SLA02", "GRP01", "SUR02", "DOR01",
}


def _normalise(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def filter_accounts(source: Path, target_day) -> tuple[Path, int, str]:
    raw = pd.read_excel(source, sheet_name=0, header=None, dtype=object)
    best = None

    # TPN exports sometimes include title rows above the table. Find the header
    # by testing each early row and selecting the column containing our account codes.
    for header_row in range(min(30, len(raw))):
        headers = [str(value).strip() if not pd.isna(value) else "" for value in raw.iloc[header_row]]
        data = raw.iloc[header_row + 1:].copy()
        data.columns = headers
        for index, column in enumerate(data.columns):
            values = data.iloc[:, index].map(_normalise)
            matches = int(values.isin(ACCOUNTS).sum())
            if best is None or matches > best[0]:
                best = (matches, header_row, index, str(column))

    if best is None or best[0] == 0:
        raise RuntimeError(
            "Could not identify the account column: none of the requested account codes "
            "were found in the downloaded TPN workbook"
        )

    _, header_row, account_index, account_heading = best
    headers = [str(value).strip() if not pd.isna(value) else "" for value in raw.iloc[header_row]]
    frame = raw.iloc[header_row + 1:].copy()
    frame.columns = headers
    account_values = frame.iloc[:, account_index].map(_normalise)
    filtered = frame.loc[account_values.isin(ACCOUNTS)].copy()
    filtered = filtered.dropna(axis=1, how="all").dropna(axis=0, how="all")

    output = source.parent / f"TPN_Account_Export_{target_day.isoformat()}.xlsx"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        filtered.to_excel(writer, sheet_name="Account Export", index=False)
        sheet = writer.book["Account Export"]
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column_cells in sheet.columns:
            width = min(
                45,
                max(10, max(len(str(cell.value or "")) for cell in column_cells) + 2),
            )
            sheet.column_dimensions[column_cells[0].column_letter].width = width

    return output, len(filtered), account_heading


def send_email(attachment: Path, row_count: int, target_day) -> None:
    username = os.getenv("MAIL_USERNAME", "sam@mcmurrayshaulage.com").strip()
    password = os.getenv("MAIL_PASSWORD", "")
    sender = os.getenv("MAIL_FROM", username).strip()
    recipient = os.getenv(
        "MAIL_TO", "customerservice@mcmurrayshaulage.com"
    ).strip()

    if not password:
        raise RuntimeError("MAIL_PASSWORD is not configured in Render")

    message = EmailMessage()
    display_date = target_day.strftime("%d/%m/%Y")
    message["From"] = sender
    message["To"] = recipient
    message["Subject"] = f"TPN Daily Account Export - {display_date}"
    message.set_content(
        f"Please find attached the TPN Browse account export for {display_date}.\n\n"
        f"Accounts: {', '.join(sorted(ACCOUNTS))}\n"
        f"Matching consignments: {row_count}\n"
    )
    message.add_attachment(
        attachment.read_bytes(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment.name,
    )

    host = os.getenv("MAIL_HOST", "smtp.office365.com")
    port = int(os.getenv("MAIL_PORT", "587"))
    use_starttls = os.getenv("MAIL_STARTTLS", "true").lower() == "true"

    with smtplib.SMTP(host, port, timeout=60) as smtp:
        smtp.ehlo()
        if use_starttls:
            smtp.starttls()
            smtp.ehlo()
        smtp.login(username, password)
        smtp.send_message(message)
    print(f"[mail] Sent {attachment.name} to {recipient}", flush=True)


def main() -> int:
    now = datetime.now(LONDON)
    force = os.getenv("FORCE_RUN", "false").lower() == "true"

    # Render cron is scheduled at both UTC possibilities for 17:00 London.
    # Only the invocation corresponding to local 17:00 performs the export.
    if not force and (now.weekday() >= 5 or now.hour != 17):
        print(
            f"[schedule] Skipped at {now.isoformat()}; waiting for a weekday 17:00 Europe/London run",
            flush=True,
        )
        return 0

    target_day = now.date()
    print(f"[job] Starting export for {target_day.isoformat()}", flush=True)
    source = download_today_export(target_day)
    attachment, row_count, account_heading = filter_accounts(source, target_day)
    print(
        f"[filter] {row_count} rows matched using column {account_heading!r}",
        flush=True,
    )
    send_email(attachment, row_count, target_day)
    print("[job] Completed", flush=True)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"[job] ERROR: {type(exc).__name__}: {exc}", file=sys.stderr, flush=True)
        raise
