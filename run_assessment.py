from pathlib import Path

from src.excel_reader import read_change_intake
from src.validator import validate_change_intake
from src.risk_engine import assess_changes
from src.excel_writer import write_assessment_results
from src.report_generator import generate_portfolio_html_report


PROJECT_ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "PharmaChangeIQ_Template.xlsx"
OUTPUT_PATH = PROJECT_ROOT / "data" / "output" / "PharmaChangeIQ_Assessed.xlsx"
REPORT_PATH = PROJECT_ROOT / "reports" / "PharmaChangeIQ_Portfolio_Report.html"


def main() -> None:
    print("PharmaChangeIQ - Intake validation and risk assessment")
    print(f"Reading template: {TEMPLATE_PATH}")

    df = read_change_intake(TEMPLATE_PATH)

    print(f"\nLoaded change cases: {len(df)}")
    print(df[["Change ID", "Product Category", "Market", "Change Type", "Urgency"]])

    result = validate_change_intake(df, PROJECT_ROOT)

    print("\nValidation result")
    print(f"Valid: {result.is_valid}")

    if result.errors:
        print("\nErrors:")
        for error in result.errors:
            print(f"- {error}")

    if result.warnings:
        print("\nWarnings:")
        for warning in result.warnings:
            print(f"- {warning}")

    if not result.is_valid:
        print("\nAssessment stopped because validation failed.")
        return

    print("\nInput validation completed successfully.")

    assessments = assess_changes(df, PROJECT_ROOT)

    print("\nRisk assessment results")
    for assessment in assessments:
        print("-" * 80)
        print(f"Change ID: {assessment.change_id}")
        print(f"Weighted score: {assessment.weighted_score}")
        print(f"Classification: {assessment.classification}")
        print(f"Escalation required: {assessment.escalation_required}")
        print(f"Primary owner: {assessment.primary_owner}")
        print(f"Secondary owner: {assessment.secondary_owner}")

    output_file = write_assessment_results(
        template_path=TEMPLATE_PATH,
        output_path=OUTPUT_PATH,
        intake_df=df,
        assessments=assessments,
    )

    report_file = generate_portfolio_html_report(
        intake_df=df,
        assessments=assessments,
        output_path=REPORT_PATH,
    )

    print("\nExcel output generated successfully.")
    print(f"Output file: {output_file}")

    print("\nHTML report generated successfully.")
    print(f"Report file: {report_file}")


if __name__ == "__main__":
    main()
