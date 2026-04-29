# Stock Factor Analysis: Tesla (TSLA) vs Disney (DIS)

Empirical asset pricing analysis comparing TSLA and DIS over the period January 2011 – December 2025 using daily return data.

## Methods

- **CAPM** — single-factor market model
- **Fama-French 3-Factor Model (FF3)** — market, size (SMB), and value (HML) factors
- **Time series diagnostics** — stationarity (ADF), autocorrelation (ACF/PACF), rolling volatility, volatility clustering

## Data Sources

- Price data: Yahoo Finance via `yfinance` (`auto_adjust=True`)
- Factor data: `fda1_stock_factor_data.csv` (Fama-French factors, sourced locally)

## Files

| File | Description |
|------|-------------|
| `factor_model_analysis.ipynb` | Main analysis notebook |
| `fda1_stock_factor_data.csv` | Raw Fama-French factor data (percent units) |
| `merged_analysis_data.csv` | Merged daily returns + FF factors (decimal units, ready to use) |
| `pyproject.toml` | Python dependencies (managed with `uv`) |

## Setup

```bash
uv sync
jupyter lab factor_model_analysis.ipynb
```
