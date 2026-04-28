import pandas as pd

from src.risk_engine import assess_changes


def test_claim_update_becomes_major_change():
    df = pd.DataFrame(
        [
            {
                "Change ID": "CHG-CLAIM",
                "Product Name": "Immune Support Supplement",
                "Product Category": "Food Supplement",
                "Market": "EU",
                "Change Type": "Claim Update",
                "Change Description": "Add immune-support claim.",
                "Ingredient Involved": "No",
                "Label Impacted": "Yes",
                "Manufacturing Impacted": "No",
                "Supplier Impacted": "No",
                "Safety Data Available": "Partial",
                "Existing Regulatory Approval Impacted": "Probably",
                "Urgency": "Medium",
                "Submitted By": "Test",
                "Submission Date": "2026-04-28",
                "Notes": "",
            }
        ]
    )

    assessments = assess_changes(df, ".")
    assessment = assessments[0]

    assert assessment.classification == "Major Change"
    assert assessment.escalation_required == "Yes"
    assert assessment.area_scores["Regulatory Impact"] == 4
    assert assessment.area_scores["Labelling Impact"] == 4
    assert assessment.primary_owner == "Regulatory Affairs"


def test_otc_packaging_supplier_change_becomes_critical():
    df = pd.DataFrame(
        [
            {
                "Change ID": "CHG-PACK",
                "Product Name": "OTC Oral Solution",
                "Product Category": "Pharma/OTC",
                "Market": "EU",
                "Change Type": "Packaging Supplier Change",
                "Change Description": "Change primary packaging supplier.",
                "Ingredient Involved": "No",
                "Label Impacted": "No",
                "Manufacturing Impacted": "No",
                "Supplier Impacted": "Yes",
                "Safety Data Available": "Yes",
                "Existing Regulatory Approval Impacted": "Probably",
                "Urgency": "High",
                "Submitted By": "Test",
                "Submission Date": "2026-04-28",
                "Notes": "",
            }
        ]
    )

    assessments = assess_changes(df, ".")
    assessment = assessments[0]

    assert assessment.classification == "Critical Change"
    assert assessment.escalation_required == "Yes"
    assert assessment.area_scores["Regulatory Impact"] >= 4
    assert assessment.area_scores["QA Impact"] >= 4


def test_cosmetic_ingredient_change_is_major_not_critical():
    df = pd.DataFrame(
        [
            {
                "Change ID": "CHG-COS",
                "Product Name": "Facial Serum",
                "Product Category": "Cosmetic",
                "Market": "EU",
                "Change Type": "Ingredient Change",
                "Change Description": "Replace fragrance component.",
                "Ingredient Involved": "Yes",
                "Label Impacted": "Yes",
                "Manufacturing Impacted": "No",
                "Supplier Impacted": "No",
                "Safety Data Available": "Partial",
                "Existing Regulatory Approval Impacted": "Unknown",
                "Urgency": "Medium",
                "Submitted By": "Test",
                "Submission Date": "2026-04-28",
                "Notes": "",
            }
        ]
    )

    assessments = assess_changes(df, ".")
    assessment = assessments[0]

    assert assessment.classification == "Major Change"
    assert assessment.escalation_required == "Yes"
    assert assessment.classification != "Critical Change"
