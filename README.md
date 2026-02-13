# Sales & Revenue Analytics

A Python-based analytics engine that processes sales data from Excel files to generate comprehensive KPIs and financial reports.

## Features

### Data Processing

- **Multi-sheet Excel processing**: Loads and merges data from 3 sheets (BD, Ciudad-Region, Presupuesto)
- **Automatic data cleaning**: Normalizes vendor names, handles whitespace, and standardizes formats
- **Smart key matching**: Matches sellers across sheets even with different name formats (FIRSTNAME LASTNAME vs LASTNAME FIRSTNAME)

### KPIs Generated

| KPI                           | Description                                         | Breakdown                            |
| ----------------------------- | --------------------------------------------------- | ------------------------------------ |
| **KPI 1 - Monthly Revenue**   | Total revenue per month                             | Total, by Region, by City, by Seller |
| **KPI 2 - Quarterly Growth**  | Quarterly revenue with % growth vs previous quarter | Total, by Region, by Seller          |
| **KPI 3 - Annual Projection** | Projected annual revenue based on available months  | Overall total                        |
| **KPI 4 - Budget Compliance** | Actual vs budget performance (%)                    | By Seller                            |

### Advanced Analysis (Exercise 3)

| Analysis                    | Description                                                       |
| --------------------------- | ----------------------------------------------------------------- |
| **Portfolio Concentration** | HHI index per seller to assess customer dependency risk           |
| **Top 5 Customer Revenue**  | Revenue from top 5 customers and % of total                       |
| **Risk Flags**              | Marks sellers with HIGH concentration (>70% from top 5 customers) |
| **Customer Reactivation**   | High-value customers inactive for 30+ days                        |

## Installation

```bash
# Clone the repository
git clone <repository-url>
cd sales-revenue-analytics

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Linux/Mac
# or: venv\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
# Run with default input file (customers_database.xlsx)
python app.py

# Run in quiet mode (no logs)
python app.py --quiet
```

### Custom Input/Output

```bash
# Specify custom input file
python app.py --input path/to/your_data.xlsx

# Specify custom output file
python app.py --output path/to/report.xlsx

# Both
python app.py --input data.xlsx --output report.xlsx
```

### Run Tests

```bash
python app.py --selftest
```

## Input File Requirements

The input Excel file must contain 3 sheets:

### Sheet: BD (Main Database)

| Column            | Description              |
| ----------------- | ------------------------ |
| Fecha Operación   | Operation date           |
| Vendedor          | Seller name              |
| Ingreso Operación | Revenue amount           |
| No. Cliente       | Customer ID              |
| Guia              | Transaction/Guide number |

### Sheet: Ciudad-Region

| Column | Description                           |
| ------ | ------------------------------------- |
| NOMBRE | Seller name (with leading ID numbers) |
| CIUDAD | City                                  |
| REGION | Region                                |

### Sheet: Presupuesto (Budget)

| Column      | Description                             |
| ----------- | --------------------------------------- |
| Vendedor    | Seller name (LASTNAME FIRSTNAME format) |
| Presupuesto | Annual budget amount                    |

## Output File Structure

The generated Excel file contains the following sheets:

```
SalesAnalysis_Report.xlsx
├── Merged_Base           # Clean, merged data (for pivot tables)
├── KPI1_Monthly_Total    # Monthly revenue (total)
├── KPI1_Monthly_Region   # Monthly revenue by region
├── KPI1_Monthly_City     # Monthly revenue by city
├── KPI1_Monthly_Seller   # Monthly revenue by seller
├── KPI2_Quarterly_Total  # Quarterly growth (total)
├── KPI2_Quarterly_Region # Quarterly growth by region
├── KPI2_Quarterly_Seller # Quarterly growth by seller
├── KPI3_Projection_Total # Annual projection
├── KPI4_Compliance_Seller# Budget compliance by seller
├── E3_Seller_Portfolio   # Portfolio concentration analysis
└── E3_Reactivate_Customers # Customers to reactivate
```

## Example Output

```
$ python app.py
[2026-02-12 10:30:15] ============================================================
[2026-02-12 10:30:15] SALES & REVENUE ANALYTICS - Starting Analysis
[2026-02-12 10:30:15] ============================================================
[2026-02-12 10:30:15] Step 1/5: Validating input file: customers_database.xlsx
[2026-02-12 10:30:15]          Input file found
[2026-02-12 10:30:15] Step 2/5: Loading Excel sheets (BD, Ciudad-Region, Presupuesto)...
[2026-02-12 10:30:16]          BD sheet: 5000 rows
[2026-02-12 10:30:16]          Ciudad-Region sheet: 50 rows
[2026-02-12 10:30:16]          Presupuesto sheet: 50 rows
[2026-02-12 10:30:16] Step 3/5: Cleaning data and merging tables...
[2026-02-12 10:30:16]          Merged base created: 5000 rows
[2026-02-12 10:30:16]          Unique sellers: 50
[2026-02-12 10:30:16]          Unique customers: 1200
[2026-02-12 10:30:16]          Date range: 2025-01-01 to 2025-06-30
[2026-02-12 10:30:16] Step 4/5: Calculating KPIs...
[2026-02-12 10:30:17]          Generated 12 Excel sheets
[2026-02-12 10:30:17] Step 5/5: Exporting to Excel: SalesAnalysis_Report.xlsx
[2026-02-12 10:30:18]          Data exported successfully
[2026-02-12 10:30:18] Analysis complete!
[2026-02-12 10:30:18] ============================================================
[2026-02-12 10:30:18] OUTPUT FILE: /path/to/SalesAnalysis_Report.xlsx
[2026-02-12 10:30:18] ============================================================
```

## Project Structure

```
sales-revenue-analytics/
├── app.py                      # Main entry point
├── src/
│   ├── utils.py                # Validation & formatting helpers
│   ├── data_processing.py      # Data loading & merging
│   ├── kpi_calculations.py     # KPI computation functions
│   └── excel_export.py         # Excel export functions
├── test/
│   └── local_test.py           # Unit tests
├── customers_database.xlsx     # Input data (not tracked in git)
├── SalesAnalysis_Report.xlsx   # Output report (not tracked in git)
├── requirements.txt            # Python dependencies
├── README.md                   # This file
└── .gitignore                  # Git ignore rules
```

## Dependencies

- Python 3.10+
- pandas
- numpy
- openpyxl

## License

MIT
