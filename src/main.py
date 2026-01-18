# src/main.py
from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict

from src.preflight import (
    check_python,
    check_required_modules,
    check_windows_com_deps,
    print_runtime_banner,
    resolve_capabilities,
)

# Preflight: deps needed to read config
check_python()
check_required_modules([("yaml", "PyYAML")])

import yaml  # safe after preflight


# -----------------------------
# Helpers
# -----------------------------
def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def load_config(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise SystemExit(
            "\n".join(
                [
                    "Configuration file not found.",
                    f"Expected location: {path}",
                    "",
                    "Ensure config/settings.yml exists and run from the project root.",
                ]
            )
        )
    with path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return data or {}


def cfg_get(cfg: Dict[str, Any], *keys: str, default: Any = None) -> Any:
    cur: Any = cfg
    for key in keys:
        if not isinstance(cur, dict) or key not in cur:
            return default
        cur = cur[key]
    return cur


def cfg_bool(cfg: Dict[str, Any], *keys: str, default: bool = False) -> bool:
    return bool(cfg_get(cfg, *keys, default=default))


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def safe_filename(s: str, max_len: int = 80) -> str:
    s = re.sub(r"[\\/:\*\?\"<>\|\n\r\t]+", "_", str(s))
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace(" ", "_")
    return s[:max_len] if len(s) > max_len else s


@dataclass(frozen=True)
class PeriodInfo:
    year: int
    month: int
    period_id: str
    period_display: str
    payslip_date_str: str


def resolve_period(cfg: Dict[str, Any]) -> PeriodInfo:
    mode = str(cfg_get(cfg, "run", "period_mode", default="manual")).strip().lower()

    ref_str = str(cfg_get(cfg, "run", "reference_date", default="")).strip()
    if ref_str:
        try:
            ref_dt = datetime.strptime(ref_str, "%Y-%m-%d").date()
        except ValueError:
            raise SystemExit(
                "Invalid run.reference_date. Expected YYYY-MM-DD (e.g. 2026-01-25)."
            )
    else:
        ref_dt = date.today()

    if mode == "manual":
        year = int(cfg_get(cfg, "run", "manual_period", "year", default=ref_dt.year))
        month = int(cfg_get(cfg, "run", "manual_period", "month", default=ref_dt.month))
    elif mode == "auto_current_month":
        year, month = ref_dt.year, ref_dt.month
    elif mode == "auto_previous_month":
        if ref_dt.month == 1:
            year, month = ref_dt.year - 1, 12
        else:
            year, month = ref_dt.year, ref_dt.month - 1
    else:
        raise SystemExit(
            "\n".join(
                [
                    "Invalid run.period_mode.",
                    "Allowed values: auto_previous_month, auto_current_month, manual",
                ]
            )
        )

    id_format = str(cfg_get(cfg, "period", "id_format", default="{year}-{month:02d}"))
    display_format = str(
        cfg_get(cfg, "period", "display_format", default="{month_name} {year}")
    )
    payslip_date_format = str(
        cfg_get(cfg, "period", "payslip_date_format", default="%d %B %Y")
    )

    month_name = date(year, month, 1).strftime("%B")
    period_id = id_format.format(year=year, month=month, month_name=month_name)
    period_display = display_format.format(year=year, month=month, month_name=month_name)

    payslip_dt = date(year, month, 1)
    payslip_date_str = payslip_dt.strftime(payslip_date_format)

    return PeriodInfo(
        year=year,
        month=month,
        period_id=period_id,
        period_display=period_display,
        payslip_date_str=payslip_date_str,
    )


def configure_logging(cfg: Dict[str, Any], logs_dir: Path, log_filename: str) -> None:
    level_str = str(cfg_get(cfg, "logging", "level", default="INFO")).upper()
    level = getattr(logging, level_str, logging.INFO)

    log_to_console = bool(cfg_get(cfg, "logging", "log_to_console", default=True))
    log_to_file = bool(cfg_get(cfg, "logging", "log_to_file", default=True))

    handlers = []

    if log_to_console:
        handlers.append(logging.StreamHandler())

    if log_to_file:
        ensure_dir(logs_dir)
        log_file = logs_dir / log_filename
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))

    # Ensure re-runs in the same interpreter don't duplicate handlers
    root_logger = logging.getLogger()
    for h in list(root_logger.handlers):
        root_logger.removeHandler(h)

    logging.basicConfig(
        level=level,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        handlers=handlers if handlers else None,
    )


# -----------------------------
# Main
# -----------------------------
def main() -> None:
    repo = project_root()
    os.chdir(repo)

    config_path = repo / "config" / "settings.yml"
    config = load_config(config_path)

    intended_email = cfg_bool(config, "email", "enabled", default=False)
    pdf_engine = str(cfg_get(config, "pdf", "engine", default="excel")).strip().lower()
    intended_pdf = pdf_engine == "excel"

    pdf_enabled, email_enabled = resolve_capabilities(
        pdf_enabled=intended_pdf,
        email_enabled=intended_email,
        allow_dev_fallback=True,
    )

    check_windows_com_deps(pdf_enabled=pdf_enabled, email_enabled=email_enabled)

    check_required_modules(
        [
            ("pandas", "pandas"),
            ("openpyxl", "openpyxl"),
        ]
    )

    period = resolve_period(config)

    # Daily log folders + timestamped filenames
    run_dt = datetime.now()
    run_date_str = run_dt.strftime("%Y-%m-%d")  # daily folder
    run_ts_str = run_dt.strftime("%Y-%m-%d_%H-%M-%S")  # filename timestamp

    base_dir = Path(cfg_get(config, "output", "base_dir", default="output"))
    output_root = (repo / base_dir).resolve()

    xlsx_dir_name = str(cfg_get(config, "output", "structure", "xlsx_dir", default="xlsx"))
    pdf_dir_name = str(cfg_get(config, "output", "structure", "pdf_dir", default="pdf"))
    logs_dir_name = str(cfg_get(config, "output", "structure", "logs_dir", default="logs"))
    summary_dir_name = str(cfg_get(config, "output", "structure", "summary_dir", default="summary"))

    run_dir = output_root / period.period_id
    xlsx_dir = run_dir / xlsx_dir_name
    pdf_dir = run_dir / pdf_dir_name
    logs_dir = run_dir / logs_dir_name / run_date_str  # daily granularity
    summary_dir = run_dir / summary_dir_name

    for d in (xlsx_dir, pdf_dir, logs_dir, summary_dir):
        ensure_dir(d)

    log_filename = f"run_log_{run_ts_str}.log"
    configure_logging(config, logs_dir, log_filename)

    print_runtime_banner(pdf_enabled=pdf_enabled, email_enabled=email_enabled)
    print(f"Log file: {logs_dir / log_filename}")
    print("")

    workbook_path = str(cfg_get(config, "workbook", "path", default=""))
    if not workbook_path:
        raise SystemExit("workbook.path is missing in settings.yml")

    template_path = (repo / workbook_path).resolve()
    if not template_path.exists():
        raise SystemExit(f"Workbook not found at: {template_path}")

    data_source_sheet = str(cfg_get(config, "workbook", "data_source_sheet", default="Data Source"))
    employee_ref_named_range = str(cfg_get(config, "workbook", "employee_ref_named_range", default="EmployeeRef"))

    reference_col = str(cfg_get(config, "data_source", "reference_column", default="Reference Number"))
    name_col = str(cfg_get(config, "data_source", "employee_name_column", default="Employee Name"))
    email_col = str(cfg_get(config, "data_source", "email_column", default="Email"))
    required_cols = cfg_get(
        config,
        "data_source",
        "required_non_null_columns",
        default=[reference_col, name_col, email_col],
    )
    required_cols = list(required_cols) if isinstance(required_cols, (list, tuple)) else [reference_col, name_col, email_col]

    filename_pattern = str(cfg_get(config, "output", "filename_pattern", default="{ref}_{name}_{date}"))

    logger = logging.getLogger(__name__)
    logger.info("Template workbook: %s", template_path)
    logger.info("Run directory: %s", run_dir)
    logger.info("Period ID: %s", period.period_id)
    logger.info("Period display: %s", period.period_display)

    from src.data_io.load_data import load_employees
    from src.data_io.template_writer import write_employee_to_template

    employees = load_employees(
        workbook_path=str(template_path),
        sheet_name=data_source_sheet,
        reference_col=reference_col,
        name_col=name_col,
        email_col=email_col,
        required_non_null_columns=required_cols,
    )

    if not employees:
        raise SystemExit("No employees found after filtering required columns.")

    # -----------------------------
    # Optional capabilities
    # -----------------------------
    excel_exporter = None
    outlook_sender = None

    # PDF export options (mapped to ExcelPdfExporter.export signature)
    pdf_sheet_name = cfg_get(config, "pdf", "sheet_name", default=None)
    pdf_quality = str(cfg_get(config, "pdf", "quality", default="standard")).strip().lower()
    pdf_open_after = cfg_bool(config, "pdf", "open_after_publish", default=False)
    pdf_ignore_print_areas = cfg_bool(config, "pdf", "ignore_print_areas", default=False)

    # Optional email config
    sender_mailbox = cfg_get(config, "email", "sender_mailbox", default=None)

    if pdf_enabled:
        from src.pdf.excel_pdf_exporter import ExcelPdfExporter
        excel_exporter = ExcelPdfExporter()

    if email_enabled:
        from src.email.outlook_sender import OutlookEmailSender
        send_mode = str(cfg_get(config, "email", "send_mode", default="send"))
        outlook_sender = OutlookEmailSender(send_mode=send_mode)

    subject_tmpl = str(
        cfg_get(
            config,
            "email",
            "message",
            "subject",
            default="Payslip – {period_display}",
        )
    )
    body_tmpl = str(
        cfg_get(
            config,
            "email",
            "message",
            "body",
            default="Dear {name},\n\nPlease find attached your payslip for {period_display}.\n",
        )
    )

    processed = 0

    # -----------------------------
    # Main processing loop
    # -----------------------------
    for emp in employees:
        emp_ref = (emp.get("ref") or "").strip()
        emp_name = (emp.get("name") or "").strip()
        emp_email = (emp.get("email") or "").strip()

        if not emp_ref:
            logger.warning("Skipping employee with missing ref.")
            continue

        date_token = period.period_id
        raw_filename = filename_pattern.format(ref=emp_ref, name=emp_name, date=date_token)
        file_stem = safe_filename(raw_filename)

        xlsx_path = xlsx_dir / f"{file_stem}.xlsx"
        pdf_path = pdf_dir / f"{file_stem}.pdf"

        write_employee_to_template(
            template_path=str(template_path),
            output_path=str(xlsx_path),
            employee_ref=emp_ref,
            employee_ref_named_range=employee_ref_named_range,
        )

        # PDF export (Windows + Excel)
        if pdf_enabled and excel_exporter is not None:
            with excel_exporter as exporter:
                exporter.export(
                    workbook_path=str(xlsx_path),
                    output_pdf_path=str(pdf_path),
                    sheet_name=pdf_sheet_name,
                    quality=pdf_quality,
                    open_after_publish=pdf_open_after,
                    ignore_print_areas=pdf_ignore_print_areas,
                )

        # Email sending (Outlook)
        if email_enabled and outlook_sender is not None:
            if not emp_email:
                logger.warning("No email for ref=%s; skipping email.", emp_ref)
            else:
                subject = subject_tmpl.format(
                    name=emp_name,
                    ref=emp_ref,
                    period_display=period.period_display,
                    period_id=period.period_id,
                    payslip_date=period.payslip_date_str,
                )
                body = body_tmpl.format(
                    name=emp_name,
                    ref=emp_ref,
                    period_display=period.period_display,
                    period_id=period.period_id,
                    payslip_date=period.payslip_date_str,
                )

                attachment = str(pdf_path) if pdf_enabled else str(xlsx_path)

                with outlook_sender as sender:
                    sender.send_email(
                        to_address=emp_email,
                        subject=subject,
                        body=body,
                        attachment_path=attachment,
                        sender_mailbox=sender_mailbox,
                    )

        processed += 1

    print("")
    print(f"Completed. Generated payslips for {processed} employee(s).")
    print(f"Output folder: {run_dir}")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as exc:
        raise SystemExit(
            "\n".join(
                [
                    "An unexpected error occurred.",
                    f"{type(exc).__name__}: {exc}",
                    "",
                    "Review the configuration, input workbook, and logs for details.",
                ]
            )
        )
