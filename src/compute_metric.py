import pandas as pd


def calculate_price_metrics(df: pd.DataFrame, company_name: str, ticker: str, sector: str) -> dict:
    if df.empty:
        raise ValueError(f"Empty dataframe received for {company_name}")

    latest_close = float(df["Close"].iloc[-1])
    first_close = float(df["Close"].iloc[0])

    five_year_return_pct = ((latest_close / first_close) - 1) * 100
    high_5y = float(df["High"].max())
    low_5y = float(df["Low"].min())
    avg_volume = float(df["Volume"].mean())
    traded_value = float((df["Close"] * df["Volume"]).sum())
    price_range_pct = (
        ((latest_close - low_5y) / (high_5y - low_5y) * 100)
        if high_5y != low_5y else 0
    )

    metrics = {
        "Company Name": company_name,
        "Ticker": ticker,
        "Sector": sector,
        "Latest Close": round(latest_close, 2),
        "5Y Return (%)": round(five_year_return_pct, 2),
        "5Y High": round(high_5y, 2),
        "5Y Low": round(low_5y, 2),
        "Average Volume": round(avg_volume, 0),
        "Approx Total Traded Value": round(traded_value, 2),
        "Position in 5Y Range (%)": round(price_range_pct, 2),
        "Data Points": len(df),
    }

    return metrics


def metrics_to_dataframe(metrics_list: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(metrics_list)


def generate_company_narrative(metrics: dict) -> list[str]:
    lines = []

    company = metrics["Company Name"]
    latest_close = metrics["Latest Close"]
    five_year_return = metrics["5Y Return (%)"]
    high_5y = metrics["5Y High"]
    low_5y = metrics["5Y Low"]
    range_position = metrics["Position in 5Y Range (%)"]
    avg_volume = metrics["Average Volume"]

    lines.append(
        f"{company} closed at INR {latest_close} based on the latest available market data."
    )

    if five_year_return > 15:
        lines.append(
            f"The stock delivered a strong 5-year return of {five_year_return}%, indicating solid market performance over the selected period."
        )
    elif five_year_return > 0:
        lines.append(
            f"The stock delivered a positive 5-year return of {five_year_return}%, reflecting moderate appreciation over the selected period."
        )
    else:
        lines.append(
            f"The stock delivered a negative 5-year return of {five_year_return}%, indicating price pressure over the selected period."
        )

    lines.append(
        f"The 5-year trading range was between INR {low_5y} and INR {high_5y}."
    )

    if range_position >= 75:
        lines.append(
            f"The stock is currently trading near the upper end of its 5-year range, at {range_position}% of the selected price band."
        )
    elif range_position >= 40:
        lines.append(
            f"The stock is trading around the middle portion of its 5-year range, at {range_position}% of the selected price band."
        )
    else:
        lines.append(
            f"The stock is trading closer to the lower end of its 5-year range, at {range_position}% of the selected price band."
        )

    lines.append(
        f"Average daily traded volume during the selected period stood at approximately {avg_volume:,.0f} shares."
    )

    return lines


def generate_combined_narrative(summary_df: pd.DataFrame) -> list[str]:
    if summary_df.empty:
        return ["No data available for summary commentary."]

    ranked_df = summary_df.sort_values("5Y Return (%)", ascending=False).reset_index(drop=True)

    top_company = ranked_df.iloc[0]
    bottom_company = ranked_df.iloc[-1]

    lines = [
        f"{top_company['Company Name']} generated the highest 5-year return among the selected companies at {top_company['5Y Return (%)']}%.",
        f"{bottom_company['Company Name']} generated the lowest 5-year return among the selected companies at {bottom_company['5Y Return (%)']}%.",
        "The comparative dashboard highlights differences in stock performance, trading range position, and liquidity across the selected companies.",
    ]

    return lines
