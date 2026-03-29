from config import COMPANIES
from fetch_prices import get_price_data
from load_financials import load_financials
from compute_metric import (
    calculate_price_metrics,
    metrics_to_dataframe,
    generate_company_narrative,
    generate_combined_narrative
)
from compute_financials import (
    calculate_financial_metrics,
    get_latest_year_snapshot,
    generate_financial_narrative
)
from export_excel import export_company_report, export_combined_summary

def run():
    print("Starting Automated Financial Reporting System - Step 4")

    all_metrics = []

    financials_raw_df = load_financials()
    financials_metrics_df = calculate_financial_metrics(financials_raw_df)
    latest_financial_snapshot_df = get_latest_year_snapshot(financials_metrics_df)
    financial_narrative_df = generate_financial_narrative(latest_financial_snapshot_df)

    for company in COMPANIES:
        try:
            print(f"Fetching data for {company['name']} ({company['ticker']})...")

            price_df = get_price_data(company["ticker"], period="5y")

            price_metrics = calculate_price_metrics(
                df=price_df,
                company_name=company["name"],
                ticker=company["ticker"],
                sector=company["sector"]
            )

            narrative_lines = generate_company_narrative(price_metrics)

            company_financials_df = financials_metrics_df[
                financials_metrics_df["Company Name"] == company["name"]
            ].copy()

            output_file = export_company_report(
                company_name=company["name"],
                price_df=price_df,
                metrics=price_metrics,
                narrative_lines=narrative_lines,
                company_financials_df=company_financials_df
            )

            all_metrics.append(price_metrics)

            print(f"Saved company report: {output_file}")

        except Exception as e:
            print(f"Error for {company['name']}: {e}")

    if all_metrics:
        summary_df = metrics_to_dataframe(all_metrics)
        combined_narrative = generate_combined_narrative(summary_df)

        summary_file = export_combined_summary(
            summary_df=summary_df,
            narrative_lines=combined_narrative,
            latest_financial_snapshot_df=latest_financial_snapshot_df,
            financial_narrative_df=financial_narrative_df
        )

        print(f"Saved combined summary: {summary_file}")

    print("Step 4 completed successfully.")


if __name__ == "__main__":
    run()