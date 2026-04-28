from pathlib import Path
from datetime import datetime

import pandas as pd

from src.risk_engine import RiskAssessment


def _html_escape(value) -> str:
    if value is None:
        return ""
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def generate_portfolio_html_report(
    intake_df: pd.DataFrame,
    assessments: list[RiskAssessment],
    output_path: str | Path,
) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    assessment_by_id = {a.change_id: a for a in assessments}

    total_changes = len(assessments)
    major_changes = sum(1 for a in assessments if a.classification == "Major Change")
    critical_changes = sum(1 for a in assessments if a.classification == "Critical Change")
    escalation_required = sum(1 for a in assessments if a.escalation_required == "Yes")
    average_score = round(sum(a.weighted_score for a in assessments) / total_changes, 2)

    rows_html = ""

    for _, row in intake_df.iterrows():
        change_id = row["Change ID"]
        assessment = assessment_by_id[change_id]

        key_driver_area = max(
            assessment.area_scores,
            key=assessment.area_scores.get,
        )

        key_driver = (
            f"{key_driver_area}: score {assessment.area_scores[key_driver_area]}. "
            f"{assessment.rationales[key_driver_area]}"
        )

        classification_class = assessment.classification.lower().replace(" ", "-")

        rows_html += f"""
        <tr>
            <td>{_html_escape(change_id)}</td>
            <td>{_html_escape(row["Product Name"])}</td>
            <td>{_html_escape(row["Product Category"])}</td>
            <td>{_html_escape(row["Change Type"])}</td>
            <td>{assessment.weighted_score}</td>
            <td><span class="badge {classification_class}">{_html_escape(assessment.classification)}</span></td>
            <td>{_html_escape(assessment.escalation_required)}</td>
            <td>{_html_escape(assessment.primary_owner)}</td>
            <td>{_html_escape(key_driver)}</td>
        </tr>
        """

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>PharmaChangeIQ Portfolio Report</title>
    <style>
        body {{
            font-family: Arial, Helvetica, sans-serif;
            margin: 36px;
            color: #111827;
            background: #ffffff;
        }}
        h1 {{
            margin-bottom: 4px;
            color: #111827;
        }}
        .subtitle {{
            color: #4b5563;
            margin-top: 0;
            margin-bottom: 28px;
        }}
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 12px;
            margin-bottom: 28px;
        }}
        .kpi-card {{
            border: 1px solid #e5e7eb;
            border-radius: 10px;
            padding: 14px;
            background: #f9fafb;
        }}
        .kpi-label {{
            font-size: 12px;
            color: #6b7280;
            margin-bottom: 6px;
        }}
        .kpi-value {{
            font-size: 22px;
            font-weight: bold;
            color: #111827;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 16px;
            font-size: 13px;
        }}
        th {{
            background: #e8eef7;
            text-align: left;
            padding: 10px;
            border-bottom: 1px solid #b8c2cc;
        }}
        td {{
            vertical-align: top;
            padding: 10px;
            border-bottom: 1px solid #e5e7eb;
        }}
        .badge {{
            display: inline-block;
            padding: 4px 8px;
            border-radius: 8px;
            font-weight: bold;
            font-size: 12px;
        }}
        .minor-change {{
            background: #bbf7d0;
        }}
        .moderate-change {{
            background: #fde68a;
        }}
        .major-change {{
            background: #fdba74;
        }}
        .critical-change {{
            background: #fca5a5;
        }}
        .section-note {{
            margin-top: 26px;
            color: #4b5563;
            line-height: 1.5;
        }}
        .footer {{
            margin-top: 32px;
            font-size: 12px;
            color: #6b7280;
        }}
    </style>
</head>
<body>
    <h1>PharmaChangeIQ Portfolio Report</h1>
    <p class="subtitle">
        Excel/Python regulatory and QA change impact assessment summary.
        Generated on {generated_at}.
    </p>

    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-label">Total changes</div>
            <div class="kpi-value">{total_changes}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Major changes</div>
            <div class="kpi-value">{major_changes}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Critical changes</div>
            <div class="kpi-value">{critical_changes}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Escalations</div>
            <div class="kpi-value">{escalation_required}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Average score</div>
            <div class="kpi-value">{average_score}</div>
        </div>
    </div>

    <h2>Assessed Change Cases</h2>
    <table>
        <thead>
            <tr>
                <th>Change ID</th>
                <th>Product</th>
                <th>Category</th>
                <th>Change Type</th>
                <th>Score</th>
                <th>Classification</th>
                <th>Escalation</th>
                <th>Primary Owner</th>
                <th>Key Driver</th>
            </tr>
        </thead>
        <tbody>
            {rows_html}
        </tbody>
    </table>

    <p class="section-note">
        This report is generated from a rule-based decision-support workflow. The system is not intended to replace formal regulatory, QA, safety, or legal review. Its purpose is to demonstrate structured intake validation, risk scoring, documentation mapping, and management-ready reporting for pharma and consumer health change-control scenarios.
    </p>

    <div class="footer">
        PharmaChangeIQ — portfolio project.
    </div>
</body>
</html>
"""

    output_path.write_text(html, encoding="utf-8")
    return output_path
