from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
import yaml


RISK_AREAS = [
    "Regulatory Impact",
    "QA Impact",
    "Safety Impact",
    "Labelling Impact",
    "Supply Chain Impact",
    "Business Continuity Impact",
]


@dataclass
class RiskAssessment:
    change_id: str
    area_scores: dict[str, int]
    rationales: dict[str, str]
    weighted_score: float
    classification: str
    escalation_required: str
    primary_owner: str
    secondary_owner: str


def _cap_score(score: int) -> int:
    return max(0, min(score, 4))


def load_scoring_config(project_root: str | Path = ".") -> dict[str, Any]:
    config_path = Path(project_root) / "config" / "scoring_weights.yaml"

    if not config_path.exists():
        raise FileNotFoundError(f"Scoring configuration not found: {config_path}")

    with open(config_path, "r", encoding="utf-8") as file:
        return yaml.safe_load(file)


def classify_weighted_score(score: float, thresholds: dict[str, dict[str, float]]) -> str:
    for label, limits in thresholds.items():
        if limits["min"] <= score <= limits["max"]:
            return label

    if score > 4.0:
        return "Critical Change"

    return "Unclassified"


def assess_change(row: pd.Series, config: dict[str, Any]) -> RiskAssessment:
    change_id = str(row["Change ID"])

    scores = {area: 0 for area in RISK_AREAS}
    rationales = {area: [] for area in RISK_AREAS}

    product_category = row["Product Category"]
    change_type = row["Change Type"]
    market = row["Market"]

    ingredient_involved = row["Ingredient Involved"]
    label_impacted = row["Label Impacted"]
    manufacturing_impacted = row["Manufacturing Impacted"]
    supplier_impacted = row["Supplier Impacted"]
    safety_data_available = row["Safety Data Available"]
    regulatory_approval_impacted = row["Existing Regulatory Approval Impacted"]
    urgency = row["Urgency"]

    # Product category baseline
    if product_category == "Pharma/OTC":
        scores["Regulatory Impact"] += 2
        scores["QA Impact"] += 2
        rationales["Regulatory Impact"].append("Pharma/OTC products have higher baseline regulatory sensitivity.")
        rationales["QA Impact"].append("Pharma/OTC products require stronger QA oversight.")

    elif product_category == "Medical Device":
        scores["Regulatory Impact"] += 2
        scores["QA Impact"] += 2
        rationales["Regulatory Impact"].append("Medical device or borderline products require structured regulatory assessment.")
        rationales["QA Impact"].append("Medical device changes may affect controlled quality documentation.")

    elif product_category == "Food Supplement":
        scores["Regulatory Impact"] += 1
        rationales["Regulatory Impact"].append("Food supplement claims and composition may be market-sensitive.")

    elif product_category == "Cosmetic":
        scores["Safety Impact"] += 1
        scores["Labelling Impact"] += 1
        rationales["Safety Impact"].append("Cosmetic changes may require safety review depending on ingredient and exposure.")
        rationales["Labelling Impact"].append("Cosmetic formulation changes may affect ingredient declaration or warnings.")

    # Change type rules
    if change_type == "Claim Update":
        scores["Regulatory Impact"] += 3
        scores["Labelling Impact"] += 4
        rationales["Regulatory Impact"].append("Claim updates may require substantiation and regulatory review before implementation.")
        rationales["Labelling Impact"].append("Claim updates directly affect market-facing product communication.")

    elif change_type == "Ingredient Change":
        scores["Safety Impact"] += 3
        scores["Regulatory Impact"] += 2
        scores["QA Impact"] += 1
        rationales["Safety Impact"].append("Ingredient changes may alter exposure, tolerability, or toxicological profile.")
        rationales["Regulatory Impact"].append("Ingredient changes may affect compliance status or notification requirements.")
        rationales["QA Impact"].append("Ingredient changes require controlled review before release.")

    elif change_type == "Label Artwork Change":
        scores["Labelling Impact"] += 3
        scores["Regulatory Impact"] += 2
        rationales["Labelling Impact"].append("Artwork changes may affect mandatory information, warnings, claims, or instructions.")
        rationales["Regulatory Impact"].append("Label and artwork changes require regulatory verification.")

    elif change_type == "Packaging Supplier Change":
        scores["QA Impact"] += 3
        scores["Supply Chain Impact"] += 3
        rationales["QA Impact"].append("Packaging supplier changes require supplier qualification and controlled documentation.")
        rationales["Supply Chain Impact"].append("Supplier changes may affect availability, specifications, and continuity.")

    elif change_type == "Manufacturing Site Change":
        scores["QA Impact"] += 4
        scores["Regulatory Impact"] += 3
        scores["Business Continuity Impact"] += 2
        rationales["QA Impact"].append("Manufacturing site changes require strong QA/GMP impact assessment.")
        rationales["Regulatory Impact"].append("Manufacturing site changes may require regulatory notification or approval.")
        rationales["Business Continuity Impact"].append("Site transfers can affect supply continuity and launch timelines.")

    elif change_type == "Process Change":
        scores["QA Impact"] += 3
        scores["Business Continuity Impact"] += 1
        rationales["QA Impact"].append("Process changes can affect validation, specifications, and batch consistency.")
        rationales["Business Continuity Impact"].append("Process changes may affect implementation timelines.")

    elif change_type == "Raw Material Supplier Change":
        scores["QA Impact"] += 3
        scores["Supply Chain Impact"] += 3
        scores["Safety Impact"] += 1
        rationales["QA Impact"].append("Raw material supplier changes require qualification and documentation review.")
        rationales["Supply Chain Impact"].append("Raw material sourcing changes can affect supply continuity.")
        rationales["Safety Impact"].append("Raw material changes may affect impurity or safety profile.")

    elif change_type == "Specification Change":
        scores["QA Impact"] += 4
        scores["Regulatory Impact"] += 2
        rationales["QA Impact"].append("Specification changes directly affect quality acceptance criteria.")
        rationales["Regulatory Impact"].append("Specification changes may affect registered or notified product information.")

    elif change_type == "Stability Related Change":
        scores["QA Impact"] += 3
        scores["Safety Impact"] += 2
        scores["Business Continuity Impact"] += 1
        rationales["QA Impact"].append("Stability-related changes require documented quality assessment.")
        rationales["Safety Impact"].append("Stability changes may affect product performance or degradation profile.")
        rationales["Business Continuity Impact"].append("Stability work may affect timelines and release planning.")

    elif change_type == "Market Expansion":
        scores["Regulatory Impact"] += 4
        scores["Labelling Impact"] += 2
        scores["Business Continuity Impact"] += 1
        rationales["Regulatory Impact"].append("Market expansion requires market-specific regulatory assessment.")
        rationales["Labelling Impact"].append("New markets may require local label or artwork adaptation.")
        rationales["Business Continuity Impact"].append("Market expansion may require launch-readiness coordination.")

    # Binary/controlled field rules
    if ingredient_involved == "Yes":
        scores["Safety Impact"] += 2
        scores["Regulatory Impact"] += 1
        rationales["Safety Impact"].append("Ingredient involvement increases the need for safety evaluation.")
        rationales["Regulatory Impact"].append("Ingredient involvement may affect product compliance.")

    if label_impacted == "Yes":
        scores["Labelling Impact"] += 2
        scores["Regulatory Impact"] += 1
        rationales["Labelling Impact"].append("Label impact requires controlled artwork or labelling review.")
        rationales["Regulatory Impact"].append("Label impact requires regulatory verification.")

    if manufacturing_impacted == "Yes":
        scores["QA Impact"] += 2
        scores["Business Continuity Impact"] += 1
        rationales["QA Impact"].append("Manufacturing impact requires QA assessment of process and documentation.")
        rationales["Business Continuity Impact"].append("Manufacturing impact may affect implementation timelines.")

    if supplier_impacted == "Yes":
        scores["QA Impact"] += 2
        scores["Supply Chain Impact"] += 2
        rationales["QA Impact"].append("Supplier impact requires supplier qualification or documentation review.")
        rationales["Supply Chain Impact"].append("Supplier impact may affect continuity and availability.")

    if safety_data_available in {"No", "Unknown"}:
        scores["Safety Impact"] += 3
        rationales["Safety Impact"].append("Missing or unknown safety data significantly increases safety review needs.")
    elif safety_data_available == "Partial":
        scores["Safety Impact"] += 1
        rationales["Safety Impact"].append("Partial safety data requires functional confirmation before implementation.")

    if regulatory_approval_impacted == "Yes":
        scores["Regulatory Impact"] += 3
        rationales["Regulatory Impact"].append("Existing approval or notification may be directly impacted.")
    elif regulatory_approval_impacted == "Probably":
        scores["Regulatory Impact"] += 2
        rationales["Regulatory Impact"].append("Regulatory approval impact is likely and requires RA confirmation.")
    elif regulatory_approval_impacted == "Unknown":
        scores["Regulatory Impact"] += 2
        rationales["Regulatory Impact"].append("Unknown regulatory approval impact requires RA assessment.")

    if urgency == "High":
        scores["Business Continuity Impact"] += 1
        rationales["Business Continuity Impact"].append("High urgency may compress review timelines.")
    elif urgency == "Critical":
        scores["Business Continuity Impact"] += 2
        scores["QA Impact"] += 1
        rationales["Business Continuity Impact"].append("Critical urgency requires escalation and documented prioritisation.")
        rationales["QA Impact"].append("Critical urgency increases QA oversight needs.")

    if market in {"US", "Global"}:
        scores["Regulatory Impact"] += 1
        rationales["Regulatory Impact"].append("US or global market scope increases regulatory complexity.")

    # Cap scores to 0-4
    capped_scores = {area: _cap_score(score) for area, score in scores.items()}

    weights = config["risk_areas"]
    weighted_score = sum(capped_scores[area] * weights[area] for area in RISK_AREAS)
    weighted_score = round(weighted_score, 2)

    classification = classify_weighted_score(
        weighted_score,
        config["classification_thresholds"],
    )

    # Escalation override rules
    # A weighted average can underestimate regulatory-critical scenarios.
    # In QA/RA workflows, a single high-impact regulatory or labelling driver
    # may justify escalation even if other areas are low.
    if (
        capped_scores["Regulatory Impact"] >= 4
        and capped_scores["Labelling Impact"] >= 4
    ):
        classification = "Major Change"

    if (
        capped_scores["QA Impact"] >= 4
        and capped_scores["Regulatory Impact"] >= 3
    ):
        classification = "Critical Change"

    if (
        capped_scores["Safety Impact"] >= 4
        and capped_scores["Regulatory Impact"] >= 3
        and product_category in {"Pharma/OTC", "Medical Device"}
    ):
        classification = "Critical Change"

    if (
        capped_scores["Regulatory Impact"] >= 4
        and regulatory_approval_impacted in {"Yes", "Probably"}
        and product_category in {"Pharma/OTC", "Medical Device"}
    ):
        classification = "Critical Change"

    escalation_required = "Yes" if classification in {"Major Change", "Critical Change"} else "No"

    # Ownership logic
    sorted_areas = sorted(capped_scores.items(), key=lambda item: item[1], reverse=True)
    primary_area = sorted_areas[0][0]
    secondary_area = sorted_areas[1][0]

    owner_map = {
        "Regulatory Impact": "Regulatory Affairs",
        "QA Impact": "Quality Assurance",
        "Safety Impact": "Safety / Toxicology",
        "Labelling Impact": "Regulatory Affairs / QA",
        "Supply Chain Impact": "Supply Chain / QA",
        "Business Continuity Impact": "Project Management / Operations",
    }

    primary_owner = owner_map[primary_area]
    secondary_owner = owner_map[secondary_area]

    rationale_text = {
        area: " ".join(items) if items else "No major driver identified for this area."
        for area, items in rationales.items()
    }

    return RiskAssessment(
        change_id=change_id,
        area_scores=capped_scores,
        rationales=rationale_text,
        weighted_score=weighted_score,
        classification=classification,
        escalation_required=escalation_required,
        primary_owner=primary_owner,
        secondary_owner=secondary_owner,
    )


def assess_changes(df: pd.DataFrame, project_root: str | Path = ".") -> list[RiskAssessment]:
    config = load_scoring_config(project_root)
    return [assess_change(row, config) for _, row in df.iterrows()]
