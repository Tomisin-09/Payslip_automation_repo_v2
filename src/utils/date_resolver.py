from datetime import date, datetime
import calendar


def resolve_pay_period(run_cfg: dict, period_cfg: dict) -> dict:
    """
    Resolve the pay period based on configuration.

    Returns a dictionary with:
        - year
        - month
        - month_name
        - period_id
        - period_display
        - start_date
        - end_date
    """

    # ----------------------------------
    # Determine reference date
    # ----------------------------------
    reference_date_str = run_cfg.get("reference_date")

    if reference_date_str:
        try:
            ref_date = datetime.strptime(reference_date_str, "%Y-%m-%d").date()
        except ValueError:
            raise ValueError(
                "reference_date must be in YYYY-MM-DD format"
            )
    else:
        ref_date = date.today()

    # ----------------------------------
    # Determine target year/month
    # ----------------------------------
    mode = run_cfg.get("period_mode", "auto_previous_month")

    if mode == "manual":
        year = run_cfg["manual_period"]["year"]
        month = run_cfg["manual_period"]["month"]

    elif mode == "auto_current_month":
        year = ref_date.year
        month = ref_date.month

    elif mode == "auto_previous_month":
        year = ref_date.year
        month = ref_date.month - 1

        if month == 0:
            month = 12
            year -= 1

    else:
        raise ValueError(f"Unknown period_mode: {mode}")

    # ----------------------------------
    # Month metadata
    # ----------------------------------
    month_name = calendar.month_name[month]

    # ----------------------------------
    # Period start / end dates
    # ----------------------------------
    start_date = date(year, month, 1)
    last_day = calendar.monthrange(year, month)[1]
    end_date = date(year, month, last_day)

    # ----------------------------------
    # Formatting
    # ----------------------------------
    period_id = period_cfg["id_format"].format(
        year=year,
        month=month,
    )

    period_display = period_cfg["display_format"].format(
        year=year,
        month=month,
        month_name=month_name,
    )

    return {
        "year": year,
        "month": month,
        "month_name": month_name,
        "period_id": period_id,
        "period_display": period_display,
        "start_date": start_date,
        "end_date": end_date,
    }
