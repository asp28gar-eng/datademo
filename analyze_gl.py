import pandas as pd
import os
import sys
import math
import matplotlib
matplotlib.use('Agg') # Set backend before importing pyplot
import matplotlib.pyplot as plt

def analyze_gl(file_path):
    print(f"Analyzing {file_path}...")

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

    # Clean Debit and Credit columns
    # Convert to numeric, coercing errors to NaN
    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce').fillna(0)
    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce').fillna(0)

    # Ensure Amount is numeric
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)

    # --- Statistics ---

    # Row count
    row_count = len(df)

    # Date Ranges
    # Ensure datetime format
    df['EffectiveDate'] = pd.to_datetime(df['EffectiveDate'], errors='coerce')
    df['EntryDate'] = pd.to_datetime(df['EntryDate'], errors='coerce')

    eff_date_min = df['EffectiveDate'].min()
    eff_date_max = df['EffectiveDate'].max()
    entry_date_min = df['EntryDate'].min()
    entry_date_max = df['EntryDate'].max()

    # Financial Totals
    total_debit = df['Debit'].sum()
    total_credit = df['Credit'].sum()

    # Amount Stats
    amount_sum = df['Amount'].sum()
    amount_mean = df['Amount'].mean()
    amount_min = df['Amount'].min()
    amount_max = df['Amount'].max()

    # Group By Stats
    account_type_counts = df['AccountType'].value_counts()
    source_counts = df['Source'].value_counts()

    # --- Benford's Law Analysis ---

    # Filter for non-zero amounts and take absolute value
    benford_data = df[df['Amount'] != 0]['Amount'].abs().astype(str)

    # Extract first digit
    first_digits = benford_data.str[0].astype(int)

    # Filter to ensure we only have digits 1-9 (excludes leading 0s from decimals like 0.5)
    first_digits = first_digits[first_digits.isin(range(1, 10))]

    # Calculate counts
    digit_counts = first_digits.value_counts().sort_index()
    total_count = len(first_digits)

    # Ensure all digits 1-9 are present
    for d in range(1, 10):
        if d not in digit_counts:
            digit_counts[d] = 0

    digit_counts = digit_counts.sort_index()

    # Calculate frequencies
    if total_count > 0:
        observed_freq = digit_counts / total_count
    else:
        observed_freq = pd.Series([0]*9, index=range(1, 10))

    # Expected Benford frequencies
    expected_freq = [math.log10(1 + 1/d) for d in range(1, 10)]

    # --- Output Report ---

    output_dir = 'output'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Generate Plot
    plt.figure(figsize=(10, 6))
    digits = range(1, 10)

    # Plot expected
    plt.plot(digits, expected_freq, marker='o', linestyle='-', color='r', label='Benford Expected')

    # Plot observed
    plt.bar(digits, observed_freq, alpha=0.6, color='b', label='Observed')

    plt.xlabel('First Digit')
    plt.ylabel('Frequency')
    plt.title("Benford's Law Analysis: First Digit Distribution")
    plt.xticks(digits)
    plt.legend()
    plt.grid(True, linestyle='--', alpha=0.7)

    plot_path = os.path.join(output_dir, 'benford_plot.png')
    plt.savefig(plot_path)
    plt.close()


    report_path = os.path.join(output_dir, 'report.md')

    with open(report_path, 'w') as f:
        f.write("# GL Data Analysis Report\n\n")
        f.write(f"**File Analyzed:** {file_path}\n\n")

        f.write("## Basic Statistics\n")
        f.write(f"- **Total Rows:** {row_count:,}\n")
        f.write(f"- **Total Debits:** {total_debit:,.2f}\n")
        f.write(f"- **Total Credits:** {total_credit:,.2f}\n")
        f.write(f"- **Net Amount (Sum):** {amount_sum:,.2f}\n")
        f.write(f"- **Average Amount:** {amount_mean:,.2f}\n")
        f.write(f"- **Min Amount:** {amount_min:,.2f}\n")
        f.write(f"- **Max Amount:** {amount_max:,.2f}\n\n")

        f.write("## Date Ranges\n")
        f.write("### Effective Date\n")
        f.write(f"- **Min:** {eff_date_min}\n")
        f.write(f"- **Max:** {eff_date_max}\n")
        f.write("### Entry Date\n")
        f.write(f"- **Min:** {entry_date_min}\n")
        f.write(f"- **Max:** {entry_date_max}\n\n")

        f.write("## Account Type Counts\n")
        f.write("| Account Type | Count |\n")
        f.write("|---|---:|\n")
        for type_name, count in account_type_counts.items():
            f.write(f"| {type_name} | {count:,} |\n")
        f.write("\n")

        f.write("## Source Counts\n")
        f.write("| Source | Count |\n")
        f.write("|---|---:|\n")
        for source_name, count in source_counts.items():
            f.write(f"| {source_name} | {count:,} |\n")
        f.write("\n")

        f.write("## Benford's Law Analysis\n\n")
        f.write(f"Analysis based on {total_count:,} non-zero transactions.\n\n")
        f.write("![Benford Plot](benford_plot.png)\n\n")

        f.write("| Digit | Count | Observed Freq | Expected Freq |\n")
        f.write("|---:|---:|---:|---:|\n")
        for d in range(1, 10):
            obs = observed_freq[d]
            exp = expected_freq[d-1]
            count = digit_counts[d]
            f.write(f"| {d} | {count:,} | {obs:.4f} | {exp:.4f} |\n")

    print(f"Analysis complete. Report saved to {report_path}")

if __name__ == "__main__":
    analyze_gl('je_samples (1).xlsx')
