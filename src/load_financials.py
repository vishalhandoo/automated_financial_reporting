from __future__ import annotations

from pathlib import Path
import warnings

import pandas as pd
import yfinance as yf

from config import COMPANIES


DEFAULT_FILE_PATH = "data/company_financials.xlsx"
DEFAULT_SHEET_NAME = "Financials"
YFINANCE_CACHE_DIR = Path("data/.yf-cache")
INR_CRORE_DIVISOR = 10_000_000
FX_LOOKUP_WINDOW_DAYS = 10

REQUIRED_COLUMNS = [
    "Company Name",
    "Ticker",
    "Sector",
    "Year",
    "Revenue",
    "EBITDA",
    "PAT",
    "EPS",
    "Total Debt",
    "Equity",
    "Operating Cash Flow",
    "Capital Expenditure",
]

LINE_ITEM_ALIASES = {
    "Revenue": ["Total Revenue", "Operating Revenue", "Revenue"],
    "EBITDA": ["EBITDA", "Normalized EBITDA"],
    "EBIT": ["EBIT"],
    "Operating Income": [
        "Operating Income",
        "Operating Income As Reported",
        "Total Operating Income As Reported",
    ],
    "Pretax Income": ["Pretax Income", "Pre Tax Income"],
    "Depreciation": [
        "Reconciled Depreciation",
        "Depreciation And Amortization",
        "Depreciation",
        "Amortization",
    ],
    "PAT": [
        "Net Income",
        "Net Income Common Stockholders",
        "Net Income From Continuing Operation Net Minority Interest",
    ],
    "EPS": ["Diluted EPS", "Basic EPS"],
    "Total Debt": ["Total Debt", "Total Borrowings", "Net Debt"],
    "Equity": [
        "Common Stock Equity",
        "Stockholders Equity",
        "Total Equity Gross Minority Interest",
        "Total Equity",
    ],
    "Operating Cash Flow": [
        "Operating Cash Flow",
        "Cash Flow From Continuing Operating Activities",
        "Cash Flow From Continuing Operating Activities Before Interest Paid Supplemental Data",
    ],
    "Capital Expenditure": ["Capital Expenditure", "Capital Expenditure Reported"],
}


def _configure_yfinance_cache() -> None:
    YFINANCE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    if hasattr(yf, "set_tz_cache_location"):
        yf.set_tz_cache_location(str(YFINANCE_CACHE_DIR.resolve()))


def _flatten_download_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [col[0] if isinstance(col, tuple) else col for col in df.columns]
    return df


def _validate_financial_columns(df: pd.DataFrame) -> pd.DataFrame:
    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing columns in financial dataset: {missing_cols}")

    ordered_cols = REQUIRED_COLUMNS + [col for col in df.columns if col not in REQUIRED_COLUMNS]
    return df[ordered_cols].copy()


def _read_financials_from_excel(
    file_path: str = DEFAULT_FILE_PATH,
    sheet_name: str = DEFAULT_SHEET_NAME,
) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return _validate_financial_columns(df)


def _get_statement_value(
    statement_df: pd.DataFrame,
    labels: list[str],
    period_end: pd.Timestamp,
) -> float | None:
    if statement_df is None or statement_df.empty or period_end not in statement_df.columns:
        return None

    for label in labels:
        if label in statement_df.index:
            value = statement_df.at[label, period_end]
            if pd.notna(value):
                return float(value)

    return None


def _get_fx_rate_to_inr(
    financial_currency: str,
    period_end: pd.Timestamp,
    fx_cache: dict[tuple[str, str], float],
) -> float:
    currency = (financial_currency or "INR").upper()
    if currency == "INR":
        return 1.0

    cache_key = (currency, period_end.strftime("%Y-%m-%d"))
    if cache_key in fx_cache:
        return fx_cache[cache_key]

    pair_ticker = f"{currency}INR=X"
    start_date = (period_end - pd.Timedelta(days=FX_LOOKUP_WINDOW_DAYS)).strftime("%Y-%m-%d")
    end_date = (period_end + pd.Timedelta(days=FX_LOOKUP_WINDOW_DAYS + 1)).strftime("%Y-%m-%d")

    fx_df = yf.download(
        tickers=pair_ticker,
        start=start_date,
        end=end_date,
        interval="1d",
        auto_adjust=False,
        progress=False,
        threads=False,
    )

    fx_df = _flatten_download_columns(fx_df.reset_index())

    if fx_df.empty or "Close" not in fx_df.columns:
        raise ValueError(f"Unable to fetch INR conversion rate for {currency}")

    fx_df = fx_df[fx_df["Close"].notna()].copy()
    if fx_df.empty:
        raise ValueError(f"No INR conversion rate available for {currency} near {period_end.date()}")

    fx_df["Date"] = pd.to_datetime(fx_df["Date"])
    fx_df["Distance"] = (fx_df["Date"] - pd.Timestamp(period_end)).abs()
    fx_rate = float(fx_df.sort_values("Distance").iloc[0]["Close"])

    fx_cache[cache_key] = fx_rate
    return fx_rate


def _convert_amount_to_inr_crore(value: float | None, fx_rate_to_inr: float) -> float | None:
    if value is None or pd.isna(value):
        return None
    return round((float(value) * fx_rate_to_inr) / INR_CRORE_DIVISOR, 2)


def _convert_eps_to_inr(value: float | None, fx_rate_to_inr: float) -> float | None:
    if value is None or pd.isna(value):
        return None
    return round(float(value) * fx_rate_to_inr, 2)


def _get_ebitda_value(
    income_stmt: pd.DataFrame,
    period_end: pd.Timestamp,
) -> tuple[float | None, str]:
    reported_ebitda = _get_statement_value(
        income_stmt,
        LINE_ITEM_ALIASES["EBITDA"],
        period_end,
    )
    if reported_ebitda is not None:
        return reported_ebitda, "Reported EBITDA"

    depreciation = _get_statement_value(
        income_stmt,
        LINE_ITEM_ALIASES["Depreciation"],
        period_end,
    )

    if depreciation is not None:
        ebit = _get_statement_value(
            income_stmt,
            LINE_ITEM_ALIASES["EBIT"],
            period_end,
        )
        if ebit is not None:
            return ebit + depreciation, "Derived from EBIT + Depreciation"

        operating_income = _get_statement_value(
            income_stmt,
            LINE_ITEM_ALIASES["Operating Income"],
            period_end,
        )
        if operating_income is not None:
            return operating_income + depreciation, "Derived from Operating Income + Depreciation"

        pretax_income = _get_statement_value(
            income_stmt,
            LINE_ITEM_ALIASES["Pretax Income"],
            period_end,
        )
        if pretax_income is not None:
            return pretax_income + depreciation, "Derived from Pretax Income + Depreciation"

    return None, "Unavailable"


def _extract_company_financials(
    company: dict,
    fx_cache: dict[tuple[str, str], float],
) -> list[dict]:
    ticker_obj = yf.Ticker(company["ticker"])

    info = ticker_obj.info or {}
    financial_currency = (info.get("financialCurrency") or info.get("currency") or "INR").upper()

    income_stmt = ticker_obj.income_stmt
    balance_sheet = ticker_obj.balance_sheet
    cashflow = ticker_obj.cashflow

    if income_stmt.empty or balance_sheet.empty or cashflow.empty:
        raise ValueError("Yahoo Finance did not return a complete annual statement set")

    statement_periods = sorted(
        set(income_stmt.columns)
        .intersection(balance_sheet.columns)
        .intersection(cashflow.columns)
    )[-5:]

    if not statement_periods:
        raise ValueError("No overlapping annual statement periods were found")

    records = []

    for period_end in statement_periods:
        period_end = pd.Timestamp(period_end)
        fx_rate_to_inr = _get_fx_rate_to_inr(financial_currency, period_end, fx_cache)
        ebitda_raw, ebitda_source = _get_ebitda_value(income_stmt, period_end)

        capex_raw = _get_statement_value(
            cashflow,
            LINE_ITEM_ALIASES["Capital Expenditure"],
            period_end,
        )
        if capex_raw is not None:
            capex_raw = abs(capex_raw)

        record = {
            "Company Name": company["name"],
            "Ticker": company["ticker"],
            "Sector": company["sector"],
            "Year": f"FY{period_end.year}",
            "Revenue": _convert_amount_to_inr_crore(
                _get_statement_value(income_stmt, LINE_ITEM_ALIASES["Revenue"], period_end),
                fx_rate_to_inr,
            ),
            "EBITDA": _convert_amount_to_inr_crore(ebitda_raw, fx_rate_to_inr),
            "PAT": _convert_amount_to_inr_crore(
                _get_statement_value(income_stmt, LINE_ITEM_ALIASES["PAT"], period_end),
                fx_rate_to_inr,
            ),
            "EPS": _convert_eps_to_inr(
                _get_statement_value(income_stmt, LINE_ITEM_ALIASES["EPS"], period_end),
                fx_rate_to_inr,
            ),
            "Total Debt": _convert_amount_to_inr_crore(
                _get_statement_value(balance_sheet, LINE_ITEM_ALIASES["Total Debt"], period_end),
                fx_rate_to_inr,
            ),
            "Equity": _convert_amount_to_inr_crore(
                _get_statement_value(balance_sheet, LINE_ITEM_ALIASES["Equity"], period_end),
                fx_rate_to_inr,
            ),
            "Operating Cash Flow": _convert_amount_to_inr_crore(
                _get_statement_value(cashflow, LINE_ITEM_ALIASES["Operating Cash Flow"], period_end),
                fx_rate_to_inr,
            ),
            "Capital Expenditure": _convert_amount_to_inr_crore(capex_raw, fx_rate_to_inr),
            "Financial Currency": financial_currency,
            "FX Rate to INR": round(fx_rate_to_inr, 4),
            "EBITDA Source": ebitda_source,
            "Statement Date": period_end.date().isoformat(),
            "Source": "Yahoo Finance",
        }

        if record["Revenue"] is None and record["PAT"] is None:
            continue

        records.append(record)

    if not records:
        raise ValueError("No usable financial rows were extracted from the annual statements")

    return records


def fetch_financials_from_yfinance(companies: list[dict] | None = None) -> pd.DataFrame:
    _configure_yfinance_cache()

    company_list = companies or COMPANIES
    fx_cache: dict[tuple[str, str], float] = {}
    records: list[dict] = []
    errors: list[str] = []

    for company in company_list:
        try:
            records.extend(_extract_company_financials(company, fx_cache))
        except Exception as exc:
            errors.append(f"{company['name']} ({company['ticker']}): {exc}")

    if not records:
        raise ValueError(
            "Could not fetch any company financials from Yahoo Finance. "
            + " | ".join(errors)
        )

    if errors:
        warnings.warn(
            "Financial statements were unavailable for some companies: " + " | ".join(errors),
            stacklevel=2,
        )

    df = pd.DataFrame(records)
    return _validate_financial_columns(df)


def load_financials(
    file_path: str = DEFAULT_FILE_PATH,
    sheet_name: str = DEFAULT_SHEET_NAME,
    source: str = "auto",
    companies: list[dict] | None = None,
) -> pd.DataFrame:
    """
    Load company financial data. By default the loader first tries Yahoo Finance,
    then falls back to the local Excel file if live data is unavailable.
    """
    normalized_source = source.lower()

    if normalized_source not in {"auto", "yfinance", "excel"}:
        raise ValueError("source must be one of: 'auto', 'yfinance', or 'excel'")

    if normalized_source in {"auto", "yfinance"}:
        try:
            return fetch_financials_from_yfinance(companies=companies)
        except Exception as exc:
            if normalized_source == "yfinance":
                raise

            warnings.warn(
                f"Automatic financial fetch failed ({exc}). Falling back to '{file_path}'.",
                stacklevel=2,
            )

    return _read_financials_from_excel(file_path=file_path, sheet_name=sheet_name)
