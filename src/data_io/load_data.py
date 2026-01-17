from typing import List, Dict
import pandas as pd
import logging


def load_employees(
    workbook_path: str,
    sheet_name: str,
    reference_col: str,
    name_col: str,
    email_col: str,
    required_non_null_columns: List[str],
) -> List[Dict]:
    """
    Load employee data from the Excel Data Source sheet.

    Returns a list of dicts:
    [
        {
            "ref": "...",
            "name": "...",
            "email": "..."
        }
    ]
    """

    logger = logging.getLogger(__name__)

    logger.info("Loading employees from sheet '%s'", sheet_name)

    # --------------------------------------------------
    # Read Excel
    # --------------------------------------------------
    df = pd.read_excel(
        workbook_path,
        sheet_name=sheet_name,
        dtype=str  # treat everything as string to avoid surprises
    )

    logger.info("Loaded %d rows from Data Source", len(df))

    # --------------------------------------------------
    # Validate required columns exist
    # --------------------------------------------------
    missing_cols = [
        col for col in required_non_null_columns if col not in df.columns
    ]

    if missing_cols:
        raise ValueError(
            f"Missing required columns in Data Source: {missing_cols}"
        )

    # --------------------------------------------------
    # Drop rows with nulls in required columns
    # --------------------------------------------------
    df_clean = df.dropna(subset=required_non_null_columns)

    logger.info(
        "After filtering nulls, %d employees remain",
        len(df_clean)
    )

    # --------------------------------------------------
    # Build employee records
    # --------------------------------------------------
    employees: List[Dict] = []

    for _, row in df_clean.iterrows():
        employees.append(
            {
                "ref": row[reference_col].strip(),
                "name": row[name_col].strip(),
                "email": row[email_col].strip(),
            }
        )

    logger.info("Prepared %d employee records", len(employees))

    return employees
