# generator.py
import io, zipfile, re, unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import openpyxl
import numpy as np

RUNTIME_TZ = ZoneInfo("Australia/Melbourne")
DATE_FMT   = "%d/%m/%Y"                       # dd/mm/yyyy for Date & Flag cols

def _safe_filename(name: str) -> str:
    name = " ".join(name.split())
    name = unicodedata.normalize("NFKD", name)
    return re.sub(r'[\/\\:\*\?"<>\|]', "_", name)[:180]

def _hyperlink_email(cell):
    if cell.hyperlink and cell.hyperlink.target.startswith("mailto:"):
        target = cell.hyperlink.target[7:]
        return target.split("?")[0]
    return ""

def generate(zip_input: bytes) -> tuple[bytes, pd.DataFrame]:
    """Return (zip_bytes, updated_df) from an uploaded Excel (.xlsx) file."""

    df = pd.read_excel(io.BytesIO(zip_input), header=2)
    df.columns = [str(c).strip().lower() for c in df.columns]
    wb  = openpyxl.load_workbook(io.BytesIO(zip_input), data_only=False)
    ws  = wb.active

    name_col  = "name"
    email_col = "email address"
    flag_col  = "initial response emailed"
    date_col  = "date" if "date" in df.columns else None
    if flag_col not in df.columns:
        df[flag_col] = np.nan

    # normalise blanks
    for col in (name_col, email_col, flag_col):
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
        df[col] = df[col].replace(r'^\s*$', np.nan, regex=True)

    tz_now   = datetime.now(RUNTIME_TZ)
    date_str = tz_now.strftime("%Y%m%d")
    out_zip  = io.BytesIO()
    used     = set()

    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, row in df.iterrows():
            full_name = str(row[name_col]).strip()

            # raw email or hyperlink
            email = row[email_col]
            if pd.isna(email):
                xl_row = 4 + idx                  # adjust for title/blank/header rows
                xl_col = df.columns.get_loc(email_col) + 1
                email = _hyperlink_email(ws.cell(xl_row, xl_col))
            email = str(email).strip()

            if email.lower() in ("", "none", "nan"):
                df.at[idx, flag_col] = email if email else "No email provided"
                continue

            first = full_name.split()[0]
            body  = (
                f"Dear {first},\r\n\r\n"
                "Thank you for your application and your interest in the advertised Senior Adviser role.\r\n\r\n"
                "We are currently processing applications and will be in contact in due course.\r\n\r\n"
                "Kind regards,\r\n\r\n"
                "Fiona\r\n"
            )
            headers = [
                "MIME-Version: 1.0",
                'Content-Type: text/plain; charset="utf-8"',
                "Content-Transfer-Encoding: 8bit",
                "X-Unsent: 1",
                f"To: {email}",
                f"Recipient: {email}",
                "Subject: Thank you for your application",
            ]
            eml_bytes = ("\r\n".join(headers) + "\r\n\r\n" + body).encode("utf-8")

            base = f"{date_str} {_safe_filename(full_name)} email response.eml"
            fname = base
            seq   = 2
            while fname.lower() in used:
                fname = base.replace(".eml", f"_{seq}.eml")
                seq  += 1
            used.add(fname.lower())
            zf.writestr(fname, eml_bytes)
            df.at[idx, flag_col] = tz_now.strftime(DATE_FMT)

    # format Date + Flag cols
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.strftime(DATE_FMT)

    out_zip.seek(0)
    return out_zip.getvalue(), df
