import pandas as pd


def calculate_financial_metrics(fin_df: pd.DataFrame) -> pd.DataFrame:
    """
    Calculate financial ratios and growth metrics company-wise.
    """
    df = fin_df.copy()

    df = df.sort_values(["Company Name", "Year"]).reset_index(drop=True)

    df["EBITDA Margin (%)"] = (df["EBITDA"] / df["Revenue"]) * 100
    df["PAT Margin (%)"] = (df["PAT"] / df["Revenue"]) * 100
    df["Debt to Equity"] = df["Total Debt"] / df["Equity"]
    df["Free Cash Flow"] = df["Operating Cash Flow"] - df["Capital Expenditure"]

    df["Revenue Growth (%)"] = (
        df.groupby("Company Name")["Revenue"].pct_change() * 100
    )
    df["PAT Growth (%)"] = (
        df.groupby("Company Name")["PAT"].pct_change() * 100
    )
    df["EPS Growth (%)"] = (
        df.groupby("Company Name")["EPS"].pct_change() * 100
    )

    numeric_cols = [
        "EBITDA Margin (%)",
        "PAT Margin (%)",
        "Debt to Equity",
        "Free Cash Flow",
        "Revenue Growth (%)",
        "PAT Growth (%)",
        "EPS Growth (%)",
    ]

    for col in numeric_cols:
        df[col] = df[col].round(2)

    return df


def get_latest_year_snapshot(financial_metrics_df: pd.DataFrame) -> pd.DataFrame:
    """
    Take the latest year available for each company.
    """
    latest_df = (
        financial_metrics_df
        .sort_values(["Company Name", "Year"])
        .groupby("Company Name", as_index=False)
        .tail(1)
        .reset_index(drop=True)
    )

    return latest_df


def _format_amount(value: float) -> str:
    if pd.isna(value):
        return "not available"
    return f"INR {value:,.2f} crore"


def _format_percentage(value: float) -> str:
    if pd.isna(value):
        return "not available"
    return f"{value:.2f}%"


def _format_ratio(value: float) -> str:
    if pd.isna(value):
        return "not available"
    return f"{value:.2f}"


def generate_financial_narrative(latest_df: pd.DataFrame) -> pd.DataFrame:
    """
    Generate narrative commentary for each company based on latest-year fundamentals.
    """
    narratives = []

    for _, row in latest_df.iterrows():
        lines = []
        financial_currency = row.get("Financial Currency")
        ebitda_source = row.get("EBITDA Source")
        currency_note = ""

        if pd.notna(financial_currency) and financial_currency != "INR":
            currency_note = (
                f" Figures were converted from {financial_currency}-reported statements into INR."
            )

        lines.append(
            f"{row['Company Name']} reported revenue of {_format_amount(row['Revenue'])} in {row['Year']}.{currency_note}"
        )

        ebitda_margin = row["EBITDA Margin (%)"]
        pat_margin = row["PAT Margin (%)"]

        if pd.notna(ebitda_margin) and pd.notna(pat_margin):
            ebitda_label = "EBITDA margin"
            if pd.notna(ebitda_source) and ebitda_source not in {"Reported EBITDA", "Unavailable"}:
                ebitda_label = "Derived EBITDA margin"
            lines.append(
                f"{ebitda_label} stood at {_format_percentage(ebitda_margin)} while PAT margin was {_format_percentage(pat_margin)}."
            )
        elif pd.notna(pat_margin):
            lines.append(
                f"PAT margin stood at {_format_percentage(pat_margin)}. EBITDA margin was not available in the source statements."
            )
        else:
            lines.append("Margin data was not fully available in the source statements.")

        lines.append(
            f"The company reported debt-to-equity of {_format_ratio(row['Debt to Equity'])} and free cash flow of {_format_amount(row['Free Cash Flow'])}."
        )

        if pd.notna(row["Revenue Growth (%)"]):
            if row["Revenue Growth (%)"] > 0:
                lines.append(
                    f"Revenue grew by {_format_percentage(row['Revenue Growth (%)'])} year-on-year."
                )
            else:
                lines.append(
                    f"Revenue declined by {_format_percentage(abs(row['Revenue Growth (%)']))} year-on-year."
                )

        if pd.notna(row["PAT Growth (%)"]):
            if row["PAT Growth (%)"] > 0:
                lines.append(
                    f"Profit after tax increased by {_format_percentage(row['PAT Growth (%)'])} year-on-year."
                )
            else:
                lines.append(
                    f"Profit after tax declined by {_format_percentage(abs(row['PAT Growth (%)']))} year-on-year."
                )

        narratives.append({
            "Company Name": row["Company Name"],
            "Ticker": row["Ticker"],
            "Year": row["Year"],
            "Narrative": " ".join(lines),
        })

    return pd.DataFrame(narratives)
