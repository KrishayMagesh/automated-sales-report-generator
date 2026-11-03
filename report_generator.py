import sys
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # save charts to files instead of opening a window
import matplotlib.pyplot as plt

# Columns this script expects (these match the Kaggle Superstore dataset).
REQUIRED = ["Order Date", "Category", "Sub-Category", "Product Name",
            "Sales", "Quantity", "Discount", "Profit"]


def load_data(path):
    """Load the CSV. The Superstore file needs latin1 encoding."""
    try:
        return pd.read_csv(path, encoding="utf-8")
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="latin1")


def clean_data(df):
    """Fix common problems: messy headers, duplicates, blanks, wrong types."""
    df = df.copy()
    rows_before = len(df)

    # 1. Standardize column names (strip spaces)
    df.columns = [c.strip() for c in df.columns]

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"CSV is missing expected columns: {missing}\n"
                         f"Columns found: {list(df.columns)}")

    # 2. Remove duplicate rows
    df = df.drop_duplicates()

    # 3. Convert types (bad values become NaN / NaT instead of crashing)
    df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
    for col in ["Sales", "Quantity", "Discount", "Profit"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # 4. Drop rows missing the essentials
    df = df.dropna(subset=["Order Date", "Sales", "Profit"])

    # 5. Clean up text columns (extra spaces, inconsistent capitalization)
    for col in ["Category", "Sub-Category", "Product Name"]:
        df[col] = df[col].astype(str).str.strip()
    df["Category"] = df["Category"].str.title()

    print(f"Cleaning: {rows_before} rows -> {len(df)} rows "
          f"({rows_before - len(df)} removed)")
    return df


def build_summary(df):
    """Return a dict of summary tables."""
    # Overall numbers
    overview = pd.DataFrame({
        "Metric": ["Total Sales", "Total Profit", "Profit Margin (%)",
                   "Total Orders (rows)", "Average Sale"],
        "Value": [
            round(df["Sales"].sum(), 2),
            round(df["Profit"].sum(), 2),
            round(df["Profit"].sum() / df["Sales"].sum() * 100, 2),
            len(df),
            round(df["Sales"].mean(), 2),
        ],
    })

    # Sales and profit by category
    by_category = (df.groupby("Category")[["Sales", "Profit"]]
                     .sum().round(2).sort_values("Sales", ascending=False))

    # Top 5 products by sales
    top_products = (df.groupby("Product Name")[["Sales", "Profit"]]
                      .sum().round(2).nlargest(5, "Sales"))

    # Monthly sales plus month-over-month % change
    monthly = (df.set_index("Order Date")["Sales"]
                 .resample("MS").sum().round(2).to_frame("Sales"))
    monthly["MoM Change (%)"] = (monthly["Sales"].pct_change() * 100).round(1)
    monthly.index = monthly.index.strftime("%Y-%m")
    monthly.index.name = "Month"

    return {
        "Overview": overview.set_index("Metric"),
        "By Category": by_category,
        "Top 5 Products": top_products,
        "Monthly Trend": monthly,
    }


def write_report(tables, out_path):
    """Write each summary table to its own sheet in an Excel file."""
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet_name, table in tables.items():
            table.to_excel(writer, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            for col in ws.columns:  # widen columns so text isn't cut off
                width = max(len(str(c.value)) for c in col if c.value is not None)
                ws.column_dimensions[col[0].column_letter].width = width + 3
    print(f"Report saved to {out_path}")


def make_chart(monthly, chart_path):
    """Save a line chart of monthly sales as a PNG image."""
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(monthly.index, monthly["Sales"], marker="o", color="#2a6fdb")

    ax.set_title("Monthly Sales", fontsize=14, fontweight="bold")
    ax.set_xlabel("Month")
    ax.set_ylabel("Sales ($)")
    ax.yaxis.set_major_formatter(lambda value, pos: f"${value:,.0f}")
    ax.grid(True, alpha=0.3)

    # Show only every 3rd month label so they don't overlap
    ax.set_xticks(range(0, len(monthly), 3))
    ax.set_xticklabels(monthly.index[::3], rotation=45, ha="right")

    fig.tight_layout()
    fig.savefig(chart_path, dpi=150)
    plt.close(fig)
    print(f"Chart saved to {chart_path}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python report_generator.py <input.csv> [output.xlsx]")
        sys.exit(1)

    in_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else "sales_report.xlsx"

    df = clean_data(load_data(in_path))
    tables = build_summary(df)
    write_report(tables, out_path)

    # Save the chart next to the report, e.g. sales_report_monthly_sales.png
    chart_path = out_path.rsplit(".", 1)[0] + "_monthly_sales.png"
    make_chart(tables["Monthly Trend"], chart_path)


if __name__ == "__main__":
    main()
