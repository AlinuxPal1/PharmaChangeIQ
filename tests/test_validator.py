import pandas as pd

from src.validator import validate_change_intake


def make_valid_case():
    return pd.DataFrame(
        [
            {
                "Change ID": "CHG-TEST-001",
                "Product Name": "Test Product",
                "Product Category": "Food Supplement",
                "Market": "EU",
                "Change Type": "Claim Update",
                "Change Description": "Test claim update.",
                "Ingredient Involved": "No",
                "Label Impacted": "Yes",
                "Manufacturing Impacted": "No",
                "Supplier Impacted": "No",
                "Safety Data Available": "Partial",
                "Existing Regulatory Approval Impacted": "Probably",
                "Urgency": "Medium",
            }
        ]
    )


def test_valid_change_intake_passes():
    df = make_valid_case()
    result = validate_change_intake(df, ".")

    assert result.is_valid is True
    assert result.errors == []


def test_missing_required_field_fails():
    df = make_valid_case()
    df.loc[0, "Product Name"] = ""

    result = validate_change_intake(df, ".")

    assert result.is_valid is False
    assert any("missing required field 'Product Name'" in error for error in result.errors)


def test_invalid_product_category_fails():
    df = make_valid_case()
    df.loc[0, "Product Category"] = "Invalid Category"

    result = validate_change_intake(df, ".")

    assert result.is_valid is False
    assert any("Invalid values for 'Product Category'" in error for error in result.errors)


def test_duplicate_change_id_fails():
    df = pd.concat([make_valid_case(), make_valid_case()], ignore_index=True)

    result = validate_change_intake(df, ".")

    assert result.is_valid is False
    assert any("Duplicate Change ID" in error for error in result.errors)


def test_supplier_warning_for_regulated_product():
    df = make_valid_case()
    df.loc[0, "Product Category"] = "Pharma/OTC"
    df.loc[0, "Change Type"] = "Packaging Supplier Change"
    df.loc[0, "Supplier Impacted"] = "Yes"

    result = validate_change_intake(df, ".")

    assert result.is_valid is True
    assert any("supplier change for a regulated product category" in warning for warning in result.warnings)
