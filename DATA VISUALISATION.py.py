from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, PatternFill


BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "Warangal_Weather_Cleaned.csv"
REPORT_FILE = BASE_DIR / "Warangal_Weather_Analysis.xlsx"
CHART_DIR = BASE_DIR / "visualizations"
CHART_DIR.mkdir(exist_ok=True)

NUMERIC_COLUMNS = [
    "maximum_temperature",
    "minimum_temperature",
    "average_temperature",
    "average_humidity",
    "rainfall",
    "maximum_wind_speed",
]
MONTH_ORDER = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def add_dataframe(writer, dataframe, sheet_name, index=False):
    dataframe.to_excel(writer, sheet_name=sheet_name, index=index)
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for column in worksheet.columns:
        width = min(max(len(str(cell.value or "")) for cell in column) + 2, 30)
        worksheet.column_dimensions[column[0].column_letter].width = width


def save_charts(df, monthly_summary, weather_summary, correlation):
    plt.style.use("seaborn-v0_8-whitegrid")
    chart_files = []

    # Monthly temperature averages and range.
    fig, ax = plt.subplots(figsize=(10, 5))
    x = monthly_summary["month"]
    ax.plot(x, monthly_summary["maximum_temperature_mean"], marker="o", label="Maximum")
    ax.plot(x, monthly_summary["average_temperature_mean"], marker="o", label="Average")
    ax.plot(x, monthly_summary["minimum_temperature_mean"], marker="o", label="Minimum")
    ax.set(title="Average Temperatures by Month", xlabel="Month", ylabel="Temperature (°C)")
    ax.tick_params(axis="x", rotation=35)
    ax.legend()
    fig.tight_layout()
    path = CHART_DIR / "monthly_temperatures.png"
    fig.savefig(path, dpi=160)
    plt.show()
    plt.close(fig)
    chart_files.append(path)

    # Rainfall total by month.
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.bar(monthly_summary["month"], monthly_summary["rainfall_sum"], color="#3B82A0")
    ax.set(title="Total Rainfall by Month", xlabel="Month", ylabel="Rainfall (mm)")
    ax.tick_params(axis="x", rotation=35)
    fig.tight_layout()
    path = CHART_DIR / "monthly_rainfall.png"
    fig.savefig(path, dpi=160)
    plt.show()
    plt.close(fig)
    chart_files.append(path)

    # Weather code counts.
    fig, ax = plt.subplots(figsize=(10, 5))
    labels = weather_summary["weather code"].fillna("Missing").astype(str)
    ax.bar(labels, weather_summary["days"], color="#65A765")
    ax.set(title="Days by Weather Code", xlabel="Weather code", ylabel="Number of days")
    fig.tight_layout()
    path = CHART_DIR / "weather_code_frequency.png"
    fig.savefig(path, dpi=160)
    plt.show()
    plt.close(fig)
    chart_files.append(path)

    # Correlation matrix.
    fig, ax = plt.subplots(figsize=(8, 6))
    matrix = correlation.set_index("metric")
    image = ax.imshow(matrix.values, cmap="coolwarm", vmin=-1, vmax=1)
    ax.set_xticks(range(len(matrix.columns)), matrix.columns, rotation=45, ha="right")
    ax.set_yticks(range(len(matrix.index)), matrix.index)
    ax.set_title("Weather Metric Correlations")
    for row in range(len(matrix.index)):
        for col in range(len(matrix.columns)):
            ax.text(col, row, f"{matrix.iloc[row, col]:.2f}", ha="center", va="center", fontsize=8)
    fig.colorbar(image, ax=ax, label="Correlation")
    fig.tight_layout()
    path = CHART_DIR / "correlation_heatmap.png"
    fig.savefig(path, dpi=160)
    plt.show()
    plt.close(fig)
    chart_files.append(path)

    # Temperature and humidity relationship, with rainfall shown by color.
    fig, ax = plt.subplots(figsize=(8, 5))
    scatter = ax.scatter(
        df["average_temperature"], df["average_humidity"],
        c=df["rainfall"], cmap="Blues", alpha=0.75, edgecolors="none",
    )
    ax.set(title="Average Temperature vs Humidity", xlabel="Average temperature (°C)", ylabel="Average humidity (%)")
    fig.colorbar(scatter, ax=ax, label="Rainfall (mm)")
    fig.tight_layout()
    path = CHART_DIR / "temperature_humidity.png"
    fig.savefig(path, dpi=160)
    plt.show()
    plt.close(fig)
    chart_files.append(path)
    return chart_files


def main():
    df = pd.read_csv(DATA_FILE)
    required = {"date", "weather code", *NUMERIC_COLUMNS}
    missing_columns = required.difference(df.columns)
    if missing_columns:
        raise ValueError(f"CSV is missing required columns: {', '.join(sorted(missing_columns))}")

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)
    df["year"] = df["date"].dt.year
    df["month"] = pd.Categorical(df["date"].dt.month_name(), categories=MONTH_ORDER, ordered=True)

    numeric_summary = df[NUMERIC_COLUMNS].describe().T.reset_index().rename(columns={"index": "metric"})
    missing_values = df.isna().sum().rename("missing_count").reset_index().rename(columns={"index": "column"})
    missing_values["missing_percent"] = (missing_values["missing_count"] / len(df) * 100).round(2)

    monthly_summary = (
        df.groupby("month", observed=False)[NUMERIC_COLUMNS]
        .agg(["mean", "min", "max", "sum"])
        .round(2)
        .reset_index()
    )
    monthly_summary.columns = ["_".join(filter(None, col)).strip("_") if isinstance(col, tuple) else col for col in monthly_summary.columns]
    monthly_summary = monthly_summary.dropna(subset=["month"])

    yearly_summary = (
        df.groupby("year")[NUMERIC_COLUMNS]
        .agg(["mean", "min", "max", "sum"])
        .round(2)
        .reset_index()
    )
    yearly_summary.columns = ["_".join(filter(None, col)).strip("_") if isinstance(col, tuple) else col for col in yearly_summary.columns]

    correlation = df[NUMERIC_COLUMNS].corr().round(3).rename_axis("metric").reset_index()
    weather_summary = (
        df.groupby("weather code", dropna=False)
        .agg(days=("date", "count"))
        .reset_index()
        .sort_values("days", ascending=False)
    )
    weather_summary["percent"] = (weather_summary["days"] / len(df) * 100).round(2)

    extreme_rows = []
    for column in NUMERIC_COLUMNS:
        series = df[column].dropna()
        if series.empty:
            continue
        high, low = series.max(), series.min()
        extreme_rows.append({
            "metric": column,
            "highest_value": high,
            "highest_date": df.loc[df[column] == high, "date"].min(),
            "lowest_value": low,
            "lowest_date": df.loc[df[column] == low, "date"].min(),
        })
    extremes = pd.DataFrame(extreme_rows)

    overview = pd.DataFrame([
        ["Dataset", DATA_FILE.name],
        ["Rows", len(df)],
        ["Columns", len(df.columns)],
        ["Date range", f"{df['date'].min():%Y-%m-%d} to {df['date'].max():%Y-%m-%d}"],
        ["Weather days recorded", df["date"].nunique()],
        ["Total rainfall (mm)", round(df["rainfall"].sum(), 2)],
        ["Average temperature (°C)", round(df["average_temperature"].mean(), 2)],
        ["Average humidity (%)", round(df["average_humidity"].mean(), 2)],
    ], columns=["analysis_item", "value"])

    chart_files = save_charts(df, monthly_summary, weather_summary, correlation)

    with pd.ExcelWriter(REPORT_FILE, engine="openpyxl") as writer:
        add_dataframe(writer, overview, "Overview")
        add_dataframe(writer, missing_values, "Missing Values")
        add_dataframe(writer, numeric_summary, "Numeric Summary")
        add_dataframe(writer, monthly_summary, "Monthly Analysis")
        add_dataframe(writer, yearly_summary, "Yearly Analysis")
        add_dataframe(writer, correlation, "Correlation")
        add_dataframe(writer, weather_summary, "Weather Codes")
        add_dataframe(writer, extremes, "Weather Extremes")
        add_dataframe(writer, df, "Cleaned Data")

        chart_sheet = writer.book.create_sheet("Visualizations")
        chart_sheet["A1"] = "Weather Analysis Visualizations"
        chart_sheet["A1"].font = Font(bold=True, size=16, color="1F4E78")
        positions = ["A3", "J3", "A30", "J30", "A57"]
        for image_path, position in zip(chart_files, positions):
            image = ExcelImage(str(image_path))
            image.width = 640
            image.height = 360
            chart_sheet.add_image(image, position)

    print(f"Analysis completed successfully: {REPORT_FILE}")
    print(f"Charts saved in: {CHART_DIR}")


if __name__ == "__main__":
    main()

