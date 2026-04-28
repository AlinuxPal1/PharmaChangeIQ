import pandas as pd

from src.risk_engine import assess_changes
from src.document_mapper import map_required_documents


def make_claim_update_case():
    return pd.DataFrame(
        [
            {
                "Change ID": "CHG-DOC",
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


def test_document_mapper_returns_core_documents():
    df = make_claim_update_case()
    assessment = assess_changes(df, ".")[0]

    documents = map_required_documents(df.iloc[0], assessment)
    document_names = [doc.document_name for doc in documents]

    assert "Change Control Form" in document_names
    assert "Regulatory Impact Assessment" in document_names
    assert "Label/Artwork Review" in document_names
    assert "Final Approval Memo" in document_names


def test_claim_update_requires_label_review():
    df = make_claim_update_case()
    assessment = assess_changes(df, ".")[0]

    documents = map_required_documents(df.iloc[0], assessment)

    label_review = next(doc for doc in documents if doc.document_name == "Label/Artwork Review")

    assert label_review.required == "Yes"
    assert label_review.function_owner == "Regulatory Affairs / QA"


def test_supplier_change_requires_supplier_qualification():
    df = make_claim_update_case()
    df.loc[0, "Product Category"] = "Pharma/OTC"
    df.loc[0, "Change Type"] = "Packaging Supplier Change"
    df.loc[0, "Supplier Impacted"] = "Yes"
    df.loc[0, "Label Impacted"] = "No"
    df.loc[0, "Safety Data Available"] = "Yes"

    assessment = assess_changes(df, ".")[0]
    documents = map_required_documents(df.iloc[0], assessment)

    supplier_doc = next(doc for doc in documents if doc.document_name == "Supplier Qualification")

    assert supplier_doc.required == "Yes"
    assert supplier_doc.function_owner == "QA / Supply Chain"
