from pathlib import Path
import pandas as pd


REQUIRED_COLUMNS = [
    "Change ID",
    "Product Name",
    "Product Category",
    "Market",
    "Change Type",
    "Change Description",
    "Ingredient Involved",
    "Label Impacted",
    "Manufacturing Impacted",
    "Supplier Impacted",
    "Safety Data Available",
    "Existing Regulatory Approval Impacted",
    "Urgency",
    "Submitted By",
    "Submission Date",
    "Notes",
]


def read_change_intake(excel_path: str | Path) -> pd.DataFrame:
    """
    Read the Change_Intake sheet from the PharmaChangeIQ Excel template.

    Parameters
    ----------
    excel_path:
        Path to the Excel workbook.

    Returns
    -------
    pandas.DataFrame
        Cleaned dataframe containing only populated change-intake rows.
    """
    excel_path = Path(excel_path)

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    df = pd.read_excel(
        excel_path,
        sheet_name="Change_Intake",
        header=3,
        engine="openpyxl",
    )

    df = df.dropna(how="all")

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns in Change_Intake: {missing_columns}")

    df = df[REQUIRED_COLUMNS].copy()

    df = df[df["Change ID"].notna()].copy()

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()

    return df
