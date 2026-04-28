from dataclasses import dataclass
import pandas as pd

from src.risk_engine import RiskAssessment


@dataclass
class RequiredDocument:
    change_id: str
    document_name: str
    required: str
    function_owner: str
    rationale: str


def map_required_documents(
    row: pd.Series,
    assessment: RiskAssessment,
) -> list[RequiredDocument]:
    """
    Map a change scenario to the expected documentation package.

    This function keeps document logic separate from Excel writing.
    That is important because QA/RA documentation logic should be auditable
    and testable independently from the output format.
    """
    change_id = assessment.change_id

    product_category = row["Product Category"]
    change_type = row["Change Type"]

    supplier_impacted = row["Supplier Impacted"]
    manufacturing_impacted = row["Manufacturing Impacted"]
    ingredient_involved = row["Ingredient Involved"]
    label_impacted = row["Label Impacted"]
    safety_data_available = row["Safety Data Available"]
    regulatory_approval_impacted = row["Existing Regulatory Approval Impacted"]

    documents: list[RequiredDocument] = []

    documents.append(
        RequiredDocument(
            change_id=change_id,
            document_name="Change Control Form",
            required="Yes",
            function_owner="QA",
            rationale="A controlled change record is required for all assessed changes.",
        )
    )

    documents.append(
        RequiredDocument(
            change_id=change_id,
            document_name="Regulatory Impact Assessment",
            required="Yes" if assessment.area_scores["Regulatory Impact"] >= 2 else "No",
            function_owner="Regulatory Affairs",
            rationale="Required when regulatory impact is moderate or higher, or when approval/notification impact is confirmed, likely, or unknown.",
        )
    )

    documents.append(
        RequiredDocument(
            change_id=change_id,
            document_name="Label/Artwork Review",
            required="Yes" if label_impacted == "Yes" or assessment.area_scores["Labelling Impact"] >= 2 else "No",
            function_owner="Regulatory Affairs / QA",
            rationale="Required when labels, claims, artwork, warnings, instructions, or market-facing communication are impacted.",
        )
    )

    safety_required = "No"
    if assessment.area_scores["Safety Impact"] >= 2:
        safety_required = "Yes"
    elif ingredient_involved == "Yes" or safety_data_available in {"Partial", "Unknown"}:
        safety_required = "Possibly"

    documents.append(
        RequiredDocument(
            change_id=change_id,
            document_name="Safety Assessment",
            required=safety_required,
            function_owner="Safety / Toxicology",
            rationale="Required or considered when ingredient involvement, exposure changes, missing data, or safety uncertainty are present.",
        )
    )

    documents.append(
        RequiredDocument(
            change_id=change_id,
            document_name="Supplier Qualification",
            required="Yes" if supplier_impacted == "Yes" else "No",
            function_owner="QA / Supply Chain",
            rationale="Required when raw material, packaging, or service supplier impact is identified.",
        )
    )

    documents.append(
        RequiredDocument(
            change_id=change_id,
            document_name="GMP/Quality System Impact Assessment",
            required="Yes" if product_category == "Pharma/OTC" or manufacturing_impacted == "Yes" else "Possibly",
            function_owner="QA",
            rationale="Required for pharma/OTC or manufacturing-related changes; considered for other categories when quality system impact is plausible.",
        )
    )

    stability_required = "No"
    if change_type == "Stability Related Change":
        stability_required = "Yes"
    elif product_category == "Pharma/OTC" and change_type in {
        "Packaging Supplier Change",
        "Manufacturing Site Change",
        "Ingredient Change",
        "Process Change",
    }:
        stability_required = "Possibly"

    documents.append(
        RequiredDocument(
            change_id=change_id,
            document_name="Stability Assessment",
            required=stability_required,
            function_owner="R&D / QA",
            rationale="Required or considered when the change may affect shelf life, storage conditions, packaging compatibility, or degradation profile.",
        )
    )

    variation_required = "No"
    if regulatory_approval_impacted in {"Yes", "Probably"} and product_category in {
        "Pharma/OTC",
        "Medical Device",
    }:
        variation_required = "Yes"
    elif regulatory_approval_impacted in {"Unknown", "Partial"}:
        variation_required = "Possibly"

    documents.append(
        RequiredDocument(
            change_id=change_id,
            document_name="Regulatory Variation / Notification Check",
            required=variation_required,
            function_owner="Regulatory Affairs",
            rationale="Required or considered when the change may affect authorization, notification, registration, or compliance status.",
        )
    )

    documents.append(
        RequiredDocument(
            change_id=change_id,
            document_name="Final Approval Memo",
            required="Yes",
            function_owner="QA / Regulatory Affairs",
            rationale="Required to document the final decision, owner, escalation status, and implementation conditions.",
        )
    )

    return documents
