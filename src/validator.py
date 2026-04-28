from dataclasses import dataclass
from pathlib import Path
import pandas as pd


@dataclass
class ValidationResult:
    is_valid: bool
    errors: list[str]
    warnings: list[str]


REQUIRED_FIELDS = [
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
]


REFERENCE_FILES = {
    "Product Category": "data/reference/product_categories.csv",
    "Market": "data/reference/markets.csv",
    "Change Type": "data/reference/change_types.csv",
}


REFERENCE_VALUE_COLUMNS = {
    "Product Category": "product_category",
    "Market": "market",
    "Change Type": "change_type",
}


YES_NO_PARTIAL_FIELDS = [
    "Ingredient Involved",
    "Label Impacted",
    "Manufacturing Impacted",
    "Supplier Impacted",
    "Safety Data Available",
    "Existing Regulatory Approval Impacted",
]


URGENCY_VALUES = {"Low", "Medium", "High", "Critical"}
YES_NO_PARTIAL_VALUES = {"Yes", "No", "Partial", "Unknown", "Probably"}


def _load_allowed_values(project_root: Path, field_name: str) -> set[str]:
    reference_path = project_root / REFERENCE_FILES[field_name]
    value_column = REFERENCE_VALUE_COLUMNS[field_name]

    if not reference_path.exists():
        raise FileNotFoundError(f"Reference file not found: {reference_path}")

    reference_df = pd.read_csv(reference_path)

    if value_column not in reference_df.columns:
        raise ValueError(f"Column '{value_column}' not found in {reference_path}")

    return set(reference_df[value_column].dropna().astype(str).str.strip())


def validate_change_intake(df: pd.DataFrame, project_root: str | Path = ".") -> ValidationResult:
    """
    Validate Change_Intake data before risk scoring.

    This function checks:
    - required fields;
    - allowed categorical values;
    - duplicate Change IDs;
    - basic QA/regulatory consistency warnings.
    """
    project_root = Path(project_root)

    errors: list[str] = []
    warnings: list[str] = []

    if df.empty:
        errors.append("No populated change cases found in Change_Intake.")
        return ValidationResult(False, errors, warnings)

    missing_columns = [col for col in REQUIRED_FIELDS if col not in df.columns]
    if missing_columns:
        errors.append(f"Missing required columns: {missing_columns}")
        return ValidationResult(False, errors, warnings)

    for index, row in df.iterrows():
        row_number = index + 5
        change_id = row.get("Change ID", f"row {row_number}")

        for field in REQUIRED_FIELDS:
            value = row.get(field)
            if pd.isna(value) or str(value).strip() == "":
                errors.append(f"{change_id}: missing required field '{field}'.")

    duplicated_ids = df["Change ID"][df["Change ID"].duplicated()].tolist()
    if duplicated_ids:
        errors.append(f"Duplicate Change ID values found: {duplicated_ids}")

    for field in ["Product Category", "Market", "Change Type"]:
        allowed_values = _load_allowed_values(project_root, field)
        invalid_values = sorted(
            set(df[field].dropna().astype(str).str.strip()) - allowed_values
        )
        if invalid_values:
            errors.append(f"Invalid values for '{field}': {invalid_values}")

    for field in YES_NO_PARTIAL_FIELDS:
        invalid_values = sorted(
            set(df[field].dropna().astype(str).str.strip()) - YES_NO_PARTIAL_VALUES
        )
        if invalid_values:
            errors.append(f"Invalid values for '{field}': {invalid_values}")

    invalid_urgency = sorted(
        set(df["Urgency"].dropna().astype(str).str.strip()) - URGENCY_VALUES
    )
    if invalid_urgency:
        errors.append(f"Invalid values for 'Urgency': {invalid_urgency}")

    for _, row in df.iterrows():
        change_id = row["Change ID"]

        if row["Ingredient Involved"] == "Yes" and row["Safety Data Available"] in {"No", "Unknown"}:
            warnings.append(
                f"{change_id}: ingredient change or ingredient involvement with limited safety data. Safety/Toxicology review should be considered."
            )

        if row["Label Impacted"] == "Yes" and row["Existing Regulatory Approval Impacted"] in {"Partial", "Unknown"}:
            warnings.append(
                f"{change_id}: label is impacted but regulatory approval impact is uncertain. RA review is required before implementation."
            )

        if row["Supplier Impacted"] == "Yes" and row["Product Category"] in {"Pharma/OTC", "Medical Device"}:
            warnings.append(
                f"{change_id}: supplier change for a regulated product category. QA supplier qualification and regulatory assessment may be required."
            )

        if row["Urgency"] == "Critical":
            warnings.append(
                f"{change_id}: critical urgency selected. Escalation pathway should be documented."
            )

    is_valid = len(errors) == 0
    return ValidationResult(is_valid, errors, warnings)
