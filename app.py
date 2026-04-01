import io
import os
import re
import streamlit as st
import pandas as pd
import pikepdf
import pdfplumber
import plotly.express as px

st.set_page_config(
    page_title="Munward Consulting | MPESA Analyzer",
    page_icon="📊📈",
    layout="wide"
)

st.markdown("""
    <style>
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 1rem;
    }
    .metric-card {
        background: #f7f9fc;
        border: 1px solid #e6ebf2;
        padding: 14px 16px;
        border-radius: 14px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }
    .section-title {
        font-size: 1.05rem;
        font-weight: 700;
        margin-bottom: 0.4rem;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊📈 Munward Consulting MPESA Analyzer")
st.markdown("**Powered by Munward Consulting Partners**")
st.caption("Upload a password-protected M-PESA statement, analyze transactions, visualize trends, and test repayment ability.")


# =========================================================
# PDF FUNCTIONS
# =========================================================
def unlock_pdf(uploaded_file, password):
    uploaded_file.seek(0)
    pdf = pikepdf.open(uploaded_file, password=password)
    output_stream = io.BytesIO()
    pdf.save(output_stream)
    output_stream.seek(0)
    return output_stream


def extract_text_from_pdf(pdf_stream):
    all_text = []
    with pdfplumber.open(pdf_stream) as pdf:
        total_pages = len(pdf.pages)
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text:
                all_text.append(f"\n--- PAGE {page_number} OF {total_pages} ---\n{text}")
    return "\n".join(all_text)


# =========================================================
# TRANSACTION CLASSIFICATION
# =========================================================
def classify_transaction(description, counterparty, amount):
    desc = str(description).lower()

    if "agent deposit" in desc:
        return "Agent Deposit"
    if "funds received from" in desc or ("received" in desc and amount > 0):
        return "Funds Received"
    if "customer transfer to" in desc:
        return "Send Money"
    if "merchant payment to" in desc:
        return "Merchant Payment"
    if "pay bill" in desc or "paybill" in desc:
        return "Pay Bill"
    if "buy goods" in desc:
        return "Buy Goods"
    if "withdraw" in desc or "withdrawal" in desc:
        return "Withdrawal"
    if "deposit" in desc:
        return "Deposit"
    if "airtime" in desc:
        return "Airtime"
    if "fuliza" in desc:
        return "Fuliza"
    if "charge" in desc or "fee" in desc or "cost" in desc:
        return "Charges"
    if "reversal" in desc:
        return "Reversal"
    if "loan" in desc:
        return "Loan"
    if amount > 0:
        return "Money In"
    return "Money Out"


def extract_reference_target(description):
    text = str(description)
    patterns = [
        r"Merchant Payment to\s+(\d+)",
        r"Customer Transfer to\s+(\d+)",
        r"Pay Bill to\s+(\d+)",
        r"Buy Goods from\s+(\d+)"
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(1)
    return ""


# =========================================================
# PARSER
# =========================================================
def parse_transactions(text):
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    transactions = []

    transaction_pattern = re.compile(
        r"^([A-Z0-9]{8,})\s+"
        r"(\d{4}-\d{2}-\d{2})\s+"
        r"(\d{2}:\d{2}:\d{2})\s+"
        r"(.+?)\s+"
        r"(-?\d[\d,]*\.\d{2})\s+"
        r"(-?\d[\d,]*\.\d{2})$"
    )

    excluded_lines = {
        "SUMMARY",
        "DETAILED STATEMENT",
        "TRANSACTION TYPE PAID IN PAID OUT",
        "Receipt No. Completion Time Details Transaction Status Paid In Withdrawn Balance",
        "M-PESA STATEMENT"
    }

    i = 0
    while i < len(lines):
        line = lines[i]

        if line in excluded_lines:
            i += 1
            continue

        match = transaction_pattern.match(line)

        if match:
            receipt = match.group(1)
            date = match.group(2)
            time = match.group(3)
            description = match.group(4).strip()
            amount = float(match.group(5).replace(",", ""))
            balance = float(match.group(6).replace(",", ""))

            counterparty = ""
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                if (
                    not transaction_pattern.match(next_line)
                    and next_line not in excluded_lines
                    and not next_line.startswith("Page ")
                    and "--- PAGE " not in next_line
                    and "STATEMENT PERIOD" not in next_line.upper()
                    and "REQUEST DATE" not in next_line.upper()
                    and "CUSTOMER NAME" not in next_line.upper()
                    and "MOBILE NUMBER" not in next_line.upper()
                    and "EMAIL ADDRESS" not in next_line.upper()
                    and "TRANSACTION TYPE" not in next_line.upper()
                    and "COMPLETION TIME" not in next_line.upper()
                    and "DETAILS TRANSACTION STATUS" not in next_line.upper()
                ):
                    counterparty = next_line
                    i += 1

            money_in = amount if amount > 0 else 0.0
            money_out = abs(amount) if amount < 0 else 0.0
            transaction_type = classify_transaction(description, counterparty, amount)
            target_code = extract_reference_target(description)

            transactions.append({
                "Receipt": receipt,
                "Date": date,
                "Time": time,
                "Description": description,
                "Counterparty": counterparty,
                "Target Code": target_code,
                "Transaction Type": transaction_type,
                "Amount": amount,
                "Money In": money_in,
                "Money Out": money_out,
                "Balance": balance
            })

        i += 1

    return pd.DataFrame(transactions)


# =========================================================
# ENRICHMENT
# =========================================================
def enrich_transactions(df):
    if df.empty:
        return df

    df = df.copy()
    df["Datetime"] = pd.to_datetime(df["Date"] + " " + df["Time"], errors="coerce")
    df["Year"] = df["Datetime"].dt.year
    df["Month"] = df["Datetime"].dt.strftime("%Y-%m")
    df["Month Name"] = df["Datetime"].dt.strftime("%b %Y")
    df["Quarter"] = (
        df["Datetime"].dt.year.astype(str)
        + "-Q"
        + df["Datetime"].dt.quarter.astype(str)
    )
    iso = df["Datetime"].dt.isocalendar()
    df["ISO Year"] = iso.year.astype(int)
    df["ISO Week"] = iso.week.astype(int)
    df["Week Label"] = df["ISO Year"].astype(str) + "-W" + df["ISO Week"].astype(str).str.zfill(2)
    df["Week Start"] = (df["Datetime"] - pd.to_timedelta(df["Datetime"].dt.weekday, unit="D")).dt.normalize()
    df["Day"] = df["Datetime"].dt.date
    df["Weekday"] = df["Datetime"].dt.day_name()
    df["Hour"] = df["Datetime"].dt.hour
    df["Direction"] = df.apply(
        lambda row: "Inflow" if row["Money In"] > 0 else ("Outflow" if row["Money Out"] > 0 else "Neutral"),
        axis=1
    )
    df["Signed Amount"] = df["Money In"] - df["Money Out"]
    return df


def add_flags(df):
    if df.empty:
        return df

    df = df.copy()
    positive_outflows = df.loc[df["Money Out"] > 0, "Money Out"]
    large_threshold = positive_outflows.quantile(0.95) if not positive_outflows.empty else 0.0

    df["Large Outflow Flag"] = df["Money Out"] >= large_threshold if large_threshold > 0 else False
    df["Round Amount Flag"] = (
        ((df["Money Out"] > 0) & (df["Money Out"] % 100 == 0))
        | ((df["Money In"] > 0) & (df["Money In"] % 100 == 0))
    )
    df["Late Night Flag"] = df["Hour"].apply(
        lambda x: bool(pd.notnull(x) and (x >= 22 or x <= 4))
    )
    flag_cols = ["Large Outflow Flag", "Round Amount Flag", "Late Night Flag"]
    df["Any Flag"] = df[flag_cols].any(axis=1)
    return df


# =========================================================
# ANALYSIS HELPERS
# =========================================================
def detect_recurring_payments(df):
    if df.empty:
        return pd.DataFrame(columns=["Counterparty", "Transaction Type", "Occurrences", "Total Money Out", "Average Amount"])

    outflow_df = df[(df["Money Out"] > 0) & (df["Counterparty"].astype(str).str.strip() != "")]
    if outflow_df.empty:
        return pd.DataFrame(columns=["Counterparty", "Transaction Type", "Occurrences", "Total Money Out", "Average Amount"])

    recurring = (
        outflow_df.groupby(["Counterparty", "Transaction Type"], as_index=False)
        .agg(
            Occurrences=("Receipt", "count"),
            Total_Money_Out=("Money Out", "sum"),
            Average_Amount=("Money Out", "mean")
        )
    )

    recurring = recurring[recurring["Occurrences"] >= 2].copy()
    recurring.rename(
        columns={
            "Total_Money_Out": "Total Money Out",
            "Average_Amount": "Average Amount"
        },
        inplace=True
    )
    return recurring.sort_values(["Occurrences", "Total Money Out"], ascending=False)


def build_summary(df):
    if df.empty:
        return pd.DataFrame([{
            "Total Transactions": 0,
            "Total Money In": 0.0,
            "Total Money Out": 0.0,
            "Net Movement": 0.0,
            "Opening Balance": 0.0,
            "Closing Balance": 0.0,
            "Average Inflow": 0.0,
            "Average Outflow": 0.0,
            "Largest Inflow": 0.0,
            "Largest Outflow": 0.0
        }])

    inflow_df = df[df["Money In"] > 0]
    outflow_df = df[df["Money Out"] > 0]
    sorted_df = df.sort_values("Datetime")

    return pd.DataFrame([{
        "Total Transactions": len(df),
        "Total Money In": df["Money In"].sum(),
        "Total Money Out": df["Money Out"].sum(),
        "Net Movement": df["Money In"].sum() - df["Money Out"].sum(),
        "Opening Balance": sorted_df["Balance"].iloc[0],
        "Closing Balance": sorted_df["Balance"].iloc[-1],
        "Average Inflow": inflow_df["Money In"].mean() if not inflow_df.empty else 0.0,
        "Average Outflow": outflow_df["Money Out"].mean() if not outflow_df.empty else 0.0,
        "Largest Inflow": inflow_df["Money In"].max() if not inflow_df.empty else 0.0,
        "Largest Outflow": outflow_df["Money Out"].max() if not outflow_df.empty else 0.0
    }])


def build_period_trend(df, analyze_by="Month"):
    if df.empty:
        return pd.DataFrame(columns=["SortKey", "Period", "Money In", "Money Out"])

    if analyze_by == "Year":
        result = (
            df.groupby("Year", as_index=False)[["Money In", "Money Out"]]
            .sum()
            .sort_values("Year")
        )
        result["SortKey"] = result["Year"].astype(str)
        result["Period"] = result["Year"].astype(str)

    elif analyze_by == "Quarter":
        result = (
            df.groupby("Quarter", as_index=False)[["Money In", "Money Out"]]
            .sum()
            .sort_values("Quarter")
        )
        result["SortKey"] = result["Quarter"]
        result["Period"] = result["Quarter"]

    elif analyze_by == "Week":
        result = (
            df.groupby(["Week Start", "Week Label"], as_index=False)[["Money In", "Money Out"]]
            .sum()
            .sort_values("Week Start")
        )
        result["SortKey"] = result["Week Start"]
        result["Period"] = result["Week Label"]

    else:
        result = (
            df.groupby(["Month", "Month Name"], as_index=False)[["Money In", "Money Out"]]
            .sum()
            .sort_values("Month")
        )
        result["SortKey"] = result["Month"]
        result["Period"] = result["Month Name"]

    return result[["SortKey", "Period", "Money In", "Money Out"]]


def build_transaction_type_summary(df):
    if df.empty:
        return pd.DataFrame(columns=["Transaction Type", "Transactions", "Money In", "Money Out"])

    result = (
        df.groupby("Transaction Type", as_index=False)
        .agg(
            Transactions=("Receipt", "count"),
            Money_In=("Money In", "sum"),
            Money_Out=("Money Out", "sum")
        )
        .sort_values(["Money_Out", "Money_In", "Transactions"], ascending=False)
    )
    result.rename(columns={"Money_In": "Money In", "Money_Out": "Money Out"}, inplace=True)
    return result


def build_top_counterparties(df, top_n=10):
    if df.empty:
        return pd.DataFrame(columns=["Counterparty", "Transactions", "Money Out"])

    outflow_df = df[(df["Money Out"] > 0) & (df["Counterparty"].astype(str).str.strip() != "")]
    if outflow_df.empty:
        return pd.DataFrame(columns=["Counterparty", "Transactions", "Money Out"])

    return (
        outflow_df.groupby("Counterparty", as_index=False)
        .agg(
            Transactions=("Receipt", "count"),
            Money_Out=("Money Out", "sum")
        )
        .sort_values(["Money_Out", "Transactions"], ascending=False)
        .head(top_n)
        .rename(columns={"Money_Out": "Money Out"})
    )


def build_weekday_summary(df):
    if df.empty:
        return pd.DataFrame(columns=["Weekday", "Money In", "Money Out"])

    weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    result = df.groupby("Weekday", as_index=False)[["Money In", "Money Out"]].sum()
    result["Weekday"] = pd.Categorical(result["Weekday"], categories=weekday_order, ordered=True)
    return result.sort_values("Weekday")


def build_balance_trend(df):
    if df.empty:
        return pd.DataFrame(columns=["Datetime", "Balance"])
    return df.sort_values("Datetime")[["Datetime", "Balance"]].copy()


def build_cashflow_by_frequency(df, frequency="Monthly"):
    if df.empty:
        return pd.DataFrame(columns=["SortKey", "Period", "Money In", "Money Out", "Net Cashflow"])

    if frequency == "Weekly":
        result = (
            df.groupby(["Week Start", "Week Label"], as_index=False)[["Money In", "Money Out"]]
            .sum()
            .sort_values("Week Start")
        )
        result["SortKey"] = result["Week Start"]
        result["Period"] = result["Week Label"]
    else:
        result = (
            df.groupby(["Month", "Month Name"], as_index=False)[["Money In", "Money Out"]]
            .sum()
            .sort_values("Month")
        )
        result["SortKey"] = result["Month"]
        result["Period"] = result["Month Name"]

    result["Net Cashflow"] = result["Money In"] - result["Money Out"]
    return result[["SortKey", "Period", "Money In", "Money Out", "Net Cashflow"]]


def compute_repayment_ability(df, frequency="Monthly", installment=0.0,
                              reducible_types=None, reduction_rate=0.0):
    if reducible_types is None:
        reducible_types = []

    cashflow = build_cashflow_by_frequency(df, frequency=frequency)

    if cashflow.empty:
        return {
            "period_cashflow": cashflow,
            "avg_income": 0.0,
            "avg_expense": 0.0,
            "avg_surplus": 0.0,
            "median_surplus": 0.0,
            "worst_surplus": 0.0,
            "safe_installment_now": 0.0,
            "installment": float(installment),
            "coverage_ratio": 0.0,
            "income_ratio": 0.0,
            "expense_reduction_needed": max(float(installment), 0.0),
            "avg_reducible_expense": 0.0,
            "possible_reduction": 0.0,
            "safe_installment_after_reduction": 0.0,
            "can_pay_now": False,
            "can_pay_after_reduction": False
        }

    avg_income = float(cashflow["Money In"].mean())
    avg_expense = float(cashflow["Money Out"].mean())
    avg_surplus = float(cashflow["Net Cashflow"].mean())
    median_surplus = float(cashflow["Net Cashflow"].median())
    worst_surplus = float(cashflow["Net Cashflow"].min())

    safe_installment_now = max(min(avg_surplus, median_surplus), 0.0)

    coverage_ratio = (avg_surplus / installment) if installment > 0 else 0.0
    income_ratio = (installment / avg_income) if avg_income > 0 else 0.0
    expense_reduction_needed = max(installment - safe_installment_now, 0.0)

    reducible_df = df[
        (df["Money Out"] > 0) &
        (df["Transaction Type"].isin(reducible_types))
    ].copy()

    if reducible_df.empty:
        avg_reducible_expense = 0.0
    else:
        reducible_cashflow = build_cashflow_by_frequency(reducible_df, frequency=frequency)
        avg_reducible_expense = float(reducible_cashflow["Money Out"].mean()) if not reducible_cashflow.empty else 0.0

    possible_reduction = avg_reducible_expense * (reduction_rate / 100.0)
    safe_installment_after_reduction = safe_installment_now + possible_reduction

    can_pay_now = installment <= safe_installment_now
    can_pay_after_reduction = installment <= safe_installment_after_reduction

    return {
        "period_cashflow": cashflow,
        "avg_income": avg_income,
        "avg_expense": avg_expense,
        "avg_surplus": avg_surplus,
        "median_surplus": median_surplus,
        "worst_surplus": worst_surplus,
        "safe_installment_now": safe_installment_now,
        "installment": float(installment),
        "coverage_ratio": coverage_ratio,
        "income_ratio": income_ratio,
        "expense_reduction_needed": expense_reduction_needed,
        "avg_reducible_expense": avg_reducible_expense,
        "possible_reduction": possible_reduction,
        "safe_installment_after_reduction": safe_installment_after_reduction,
        "can_pay_now": can_pay_now,
        "can_pay_after_reduction": can_pay_after_reduction
    }


def build_expense_reduction_candidates(df, reducible_types=None, top_n=10):
    if reducible_types is None:
        reducible_types = []

    if df.empty:
        return pd.DataFrame(columns=["Counterparty", "Transaction Type", "Transactions", "Money Out"])

    reduction_df = df[
        (df["Money Out"] > 0) &
        (df["Transaction Type"].isin(reducible_types))
    ].copy()

    if reduction_df.empty:
        return pd.DataFrame(columns=["Counterparty", "Transaction Type", "Transactions", "Money Out"])

    reduction_df["Counterparty Clean"] = reduction_df["Counterparty"].replace("", "Unspecified")

    result = (
        reduction_df.groupby(["Counterparty Clean", "Transaction Type"], as_index=False)
        .agg(
            Transactions=("Receipt", "count"),
            Money_Out=("Money Out", "sum")
        )
        .sort_values(["Money_Out", "Transactions"], ascending=False)
        .head(top_n)
        .rename(columns={
            "Counterparty Clean": "Counterparty",
            "Money_Out": "Money Out"
        })
    )
    return result


# =========================================================
# NEW HELPERS FOR INDEPENDENT HP / LOAN CALCULATIONS
# =========================================================
def build_reducible_expense_summary(df, frequency="Monthly", reducible_types=None):
    if reducible_types is None:
        reducible_types = []

    if df.empty or not reducible_types:
        return pd.DataFrame(columns=[
            "Transaction Type",
            "Transactions",
            "Total Money Out",
            "Average Period Expense",
            "Share of Reducible Expense (%)"
        ])

    reduction_df = df[
        (df["Money Out"] > 0) &
        (df["Transaction Type"].isin(reducible_types))
    ].copy()

    if reduction_df.empty:
        return pd.DataFrame(columns=[
            "Transaction Type",
            "Transactions",
            "Total Money Out",
            "Average Period Expense",
            "Share of Reducible Expense (%)"
        ])

    periods_df = build_cashflow_by_frequency(df, frequency=frequency)
    period_count = max(len(periods_df), 1)

    summary = (
        reduction_df.groupby("Transaction Type", as_index=False)
        .agg(
            Transactions=("Receipt", "count"),
            Total_Money_Out=("Money Out", "sum")
        )
    )

    summary["Average Period Expense"] = summary["Total_Money_Out"] / period_count
    total_avg_reducible = summary["Average Period Expense"].sum()

    if total_avg_reducible > 0:
        summary["Share of Reducible Expense (%)"] = (
            summary["Average Period Expense"] / total_avg_reducible * 100
        )
    else:
        summary["Share of Reducible Expense (%)"] = 0.0

    summary.rename(columns={"Total_Money_Out": "Total Money Out"}, inplace=True)

    return summary.sort_values("Average Period Expense", ascending=False)


def suggest_expense_reduction_plan(df, frequency="Monthly", target_installment=0.0, reducible_types=None):
    if reducible_types is None:
        reducible_types = []

    base_ability = compute_repayment_ability(
        df,
        frequency=frequency,
        installment=target_installment,
        reducible_types=[],
        reduction_rate=0.0
    )

    category_summary = build_reducible_expense_summary(
        df,
        frequency=frequency,
        reducible_types=reducible_types
    )

    additional_needed = max(target_installment - base_ability["safe_installment_now"], 0.0)

    if category_summary.empty:
        category_summary["Required Reduction % If Used Alone"] = []
        category_summary["Suggested Reduction % Across Selected Categories"] = []
        category_summary["Estimated Savings at Suggested %"] = []
        return {
            "base_ability": base_ability,
            "category_summary": category_summary,
            "additional_needed": additional_needed,
            "total_reducible_avg": 0.0,
            "recommended_reduction_pct": 0.0,
            "recommended_savings": 0.0,
            "max_possible_savings": 0.0,
            "can_fund_with_selected_categories": additional_needed <= 0,
            "capacity_after_recommended_reduction": base_ability["safe_installment_now"],
            "capacity_after_max_reduction": base_ability["safe_installment_now"]
        }

    total_reducible_avg = float(category_summary["Average Period Expense"].sum())
    max_possible_savings = total_reducible_avg

    if additional_needed <= 0:
        recommended_reduction_pct = 0.0
        recommended_savings = 0.0
        can_fund = True
    else:
        recommended_reduction_pct = (
            (additional_needed / total_reducible_avg) * 100
            if total_reducible_avg > 0 else 0.0
        )
        can_fund = recommended_reduction_pct <= 100
        recommended_savings = min(additional_needed, max_possible_savings)

    def single_category_required(avg_expense):
        if additional_needed <= 0:
            return 0.0
        if avg_expense <= 0:
            return None
        pct = (additional_needed / avg_expense) * 100
        return round(pct, 2)

    category_summary = category_summary.copy()
    category_summary["Required Reduction % If Used Alone"] = category_summary["Average Period Expense"].apply(single_category_required)

    applied_pct = min(recommended_reduction_pct, 100.0) if total_reducible_avg > 0 else 0.0
    category_summary["Suggested Reduction % Across Selected Categories"] = applied_pct
    category_summary["Estimated Savings at Suggested %"] = (
        category_summary["Average Period Expense"] * (applied_pct / 100.0)
    )

    capacity_after_recommended = base_ability["safe_installment_now"] + min(additional_needed, max_possible_savings)
    capacity_after_max = base_ability["safe_installment_now"] + max_possible_savings

    return {
        "base_ability": base_ability,
        "category_summary": category_summary.sort_values("Average Period Expense", ascending=False),
        "additional_needed": additional_needed,
        "total_reducible_avg": total_reducible_avg,
        "recommended_reduction_pct": recommended_reduction_pct,
        "recommended_savings": recommended_savings,
        "max_possible_savings": max_possible_savings,
        "can_fund_with_selected_categories": can_fund,
        "capacity_after_recommended_reduction": capacity_after_recommended,
        "capacity_after_max_reduction": capacity_after_max
    }


def to_excel_bytes(sheets_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, sheet_df in sheets_dict.items():
            safe_name = sheet_name[:31]
            sheet_df.to_excel(writer, sheet_name=safe_name, index=False)
    output.seek(0)
    return output


def render_metric_card(label, value):
    st.markdown(
        f"""
        <div class="metric-card">
            <div style="font-size:0.86rem;color:#5b6470;">{label}</div>
            <div style="font-size:1.35rem;font-weight:700;margin-top:4px;">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# INPUTS
# =========================================================
left, right = st.columns([1.4, 1])

with left:
    uploaded_file = st.file_uploader("Upload M-PESA Statement PDF", type=["pdf"])

with right:
    password = st.text_input("Enter PDF Password", type="password")

process_clicked = st.button("Process Statement", type="primary", use_container_width=True)

if process_clicked:
    if uploaded_file is None:
        st.error("Please upload your M-PESA PDF statement.")
    elif not password:
        st.error("Please enter the PDF password.")
    else:
        try:
            with st.spinner("Unlocking PDF..."):
                unlocked_pdf = unlock_pdf(uploaded_file, password)

            with st.spinner("Extracting text from PDF..."):
                extracted_text = extract_text_from_pdf(unlocked_pdf)

            os.makedirs("output", exist_ok=True)
            with open("output/extracted_text.txt", "w", encoding="utf-8") as f:
                f.write(extracted_text)

            with st.spinner("Parsing and enriching transactions..."):
                df = parse_transactions(extracted_text)
                df = enrich_transactions(df)
                df = add_flags(df)

            st.session_state["mpesa_df"] = df
            st.success("Statement processed successfully.")

        except pikepdf.PasswordError:
            st.error("Incorrect PDF password. Please check and try again.")
        except Exception as e:
            st.error(f"An error occurred: {e}")


# =========================================================
# DASHBOARD
# =========================================================
if "mpesa_df" in st.session_state:
    df = st.session_state["mpesa_df"].copy()

    st.sidebar.header("Analysis Filters")

    transaction_types = sorted(df["Transaction Type"].dropna().unique().tolist()) if "Transaction Type" in df.columns else []
    directions = sorted(df["Direction"].dropna().unique().tolist()) if "Direction" in df.columns else []

    analyze_by = st.sidebar.selectbox(
        "Analyze By",
        options=["Week", "Month", "Quarter", "Year"],
        index=1
    )

    if analyze_by == "Year":
        available_periods = sorted(df["Year"].dropna().astype(str).unique().tolist())
        selected_periods = st.sidebar.multiselect("Select Year(s)", options=available_periods, default=available_periods)

    elif analyze_by == "Quarter":
        available_periods = sorted(df["Quarter"].dropna().unique().tolist())
        selected_periods = st.sidebar.multiselect("Select Quarter(s)", options=available_periods, default=available_periods)

    elif analyze_by == "Week":
        week_map = (
            df[["Week Start", "Week Label"]]
            .drop_duplicates()
            .sort_values("Week Start")
        )
        available_periods = week_map["Week Label"].tolist()
        selected_periods = st.sidebar.multiselect("Select Week(s)", options=available_periods, default=available_periods)

    else:
        month_map = (
            df[["Month", "Month Name"]]
            .drop_duplicates()
            .sort_values("Month")
        )
        available_periods = month_map["Month Name"].tolist()
        selected_periods = st.sidebar.multiselect("Select Month(s)", options=available_periods, default=available_periods)

    selected_types = st.sidebar.multiselect("Transaction Type", options=transaction_types, default=transaction_types)
    selected_directions = st.sidebar.multiselect("Direction", options=directions, default=directions)
    search_text = st.sidebar.text_input("Search Counterparty / Description / Target Code")

    filtered_df = df.copy()

    if analyze_by == "Year" and selected_periods:
        filtered_df = filtered_df[filtered_df["Year"].astype(str).isin(selected_periods)]
    elif analyze_by == "Quarter" and selected_periods:
        filtered_df = filtered_df[filtered_df["Quarter"].isin(selected_periods)]
    elif analyze_by == "Week" and selected_periods:
        filtered_df = filtered_df[filtered_df["Week Label"].isin(selected_periods)]
    elif analyze_by == "Month" and selected_periods:
        filtered_df = filtered_df[filtered_df["Month Name"].isin(selected_periods)]

    if selected_types:
        filtered_df = filtered_df[filtered_df["Transaction Type"].isin(selected_types)]
    if selected_directions:
        filtered_df = filtered_df[filtered_df["Direction"].isin(selected_directions)]
    if search_text.strip():
        term = search_text.strip().lower()
        filtered_df = filtered_df[
            filtered_df["Counterparty"].astype(str).str.lower().str.contains(term, na=False)
            | filtered_df["Description"].astype(str).str.lower().str.contains(term, na=False)
            | filtered_df["Target Code"].astype(str).str.lower().str.contains(term, na=False)
        ]

    summary_df = build_summary(filtered_df)
    period_trend_df = build_period_trend(filtered_df, analyze_by=analyze_by)
    transaction_type_df = build_transaction_type_summary(filtered_df)
    top_counterparties_df = build_top_counterparties(filtered_df, top_n=10)
    weekday_df = build_weekday_summary(filtered_df)
    balance_trend_df = build_balance_trend(filtered_df)
    recurring_df = detect_recurring_payments(filtered_df)

    tabs = st.tabs([
        "Overview",
        "Transactions",
        "Recurring",
        "Hire Purchase Simulator",
        "Loan Optimizer",
        "Export"
    ])

    with tabs[0]:
        total_in = filtered_df["Money In"].sum() if not filtered_df.empty else 0.0
        total_out = filtered_df["Money Out"].sum() if not filtered_df.empty else 0.0
        net_movement = total_in - total_out

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            render_metric_card("Transactions", f"{len(filtered_df)}")
        with m2:
            render_metric_card("Money In", f"KES {total_in:,.2f}")
        with m3:
            render_metric_card("Money Out", f"KES {total_out:,.2f}")
        with m4:
            render_metric_card("Net Movement", f"KES {net_movement:,.2f}")

        c1, c2 = st.columns(2)

        with c1:
            if not period_trend_df.empty:
                period_long = period_trend_df.melt(
                    id_vars="Period",
                    value_vars=["Money In", "Money Out"],
                    var_name="Flow",
                    value_name="Amount"
                )
                fig_period = px.bar(
                    period_long,
                    x="Period",
                    y="Amount",
                    color="Flow",
                    barmode="group",
                    title=f"{analyze_by} Money In vs Money Out"
                )
                st.plotly_chart(fig_period, use_container_width=True)

        with c2:
            if not transaction_type_df.empty:
                type_long = transaction_type_df.melt(
                    id_vars="Transaction Type",
                    value_vars=["Money In", "Money Out"],
                    var_name="Flow",
                    value_name="Amount"
                )
                fig_type = px.bar(
                    type_long,
                    x="Transaction Type",
                    y="Amount",
                    color="Flow",
                    barmode="group",
                    title="Transaction Type Analysis"
                )
                st.plotly_chart(fig_type, use_container_width=True)

        c3, c4 = st.columns(2)

        with c3:
            if not balance_trend_df.empty:
                fig_balance = px.line(
                    balance_trend_df,
                    x="Datetime",
                    y="Balance",
                    title="Balance Trend"
                )
                st.plotly_chart(fig_balance, use_container_width=True)

        with c4:
            if not weekday_df.empty:
                weekday_long = weekday_df.melt(
                    id_vars="Weekday",
                    value_vars=["Money In", "Money Out"],
                    var_name="Flow",
                    value_name="Amount"
                )
                fig_weekday = px.bar(
                    weekday_long,
                    x="Weekday",
                    y="Amount",
                    color="Flow",
                    barmode="group",
                    title="Weekday Spending and Receipts"
                )
                st.plotly_chart(fig_weekday, use_container_width=True)

        d1, d2 = st.columns(2)
        with d1:
            st.markdown('<div class="section-title">Top Counterparties by Money Out</div>', unsafe_allow_html=True)
            st.dataframe(top_counterparties_df, use_container_width=True, hide_index=True)

        with d2:
            st.markdown('<div class="section-title">Transaction Type Summary</div>', unsafe_allow_html=True)
            st.dataframe(transaction_type_df, use_container_width=True, hide_index=True)

    with tabs[1]:
        st.subheader("Filtered Transactions")

        display_columns = [
            "Datetime", "Receipt", "Transaction Type", "Counterparty", "Target Code",
            "Description", "Money In", "Money Out", "Balance", "Direction",
            "Year", "Quarter", "Month Name", "Week Label"
        ]
        available_columns = [col for col in display_columns if col in filtered_df.columns]
        st.dataframe(filtered_df[available_columns], use_container_width=True, hide_index=True)

        if not filtered_df.empty:
            st.markdown("### Transaction Type vs Amount")

            transaction_type_chart_df = (
                filtered_df.groupby("Transaction Type", as_index=False)
                .agg(
                    **{
                        "Money In": ("Money In", "sum"),
                        "Money Out": ("Money Out", "sum")
                    }
                )
            )

            transaction_type_long = transaction_type_chart_df.melt(
                id_vars="Transaction Type",
                value_vars=["Money In", "Money Out"],
                var_name="Flow",
                value_name="Amount"
            )

            fig_transaction_type = px.bar(
                transaction_type_long,
                y="Transaction Type",
                x="Amount",
                color="Flow",
                barmode="group",
                orientation="h",
                title="Transaction Type vs Amount"
            )
            st.plotly_chart(fig_transaction_type, use_container_width=True)

    with tabs[2]:
        st.subheader("Recurring Payments")

        if recurring_df.empty:
            st.info("No recurring payments detected in the current filter selection.")
        else:
            st.dataframe(recurring_df, use_container_width=True, hide_index=True)

        st.markdown("### Top Counterparties by Total Amount")

        counterparty_chart_df = (
            filtered_df[
                (filtered_df["Counterparty"].astype(str).str.strip() != "")
                & (filtered_df["Money Out"] > 0)
            ]
            .groupby("Counterparty", as_index=False)
            .agg(
                **{
                    "Money Out": ("Money Out", "sum")
                }
            )
            .sort_values("Money Out", ascending=False)
            .head(10)
        )

        if counterparty_chart_df.empty:
            st.info("No counterparty outflow data available for charting.")
        else:
            fig_counterparty = px.bar(
                counterparty_chart_df,
                y="Counterparty",
                x="Money Out",
                orientation="h",
                title="Top Counterparties by Total Amount"
            )
            st.plotly_chart(fig_counterparty, use_container_width=True)

    with tabs[3]:
        st.subheader("Hire Purchase Simulator")

        if filtered_df.empty:
            st.info("No transaction data available for the current filter selection.")
        else:
            hp1, hp2, hp3 = st.columns(3)

            with hp1:
                asset_price = st.number_input(
                    "Asset Price (KES)",
                    min_value=0.0,
                    value=100000.0,
                    step=1000.0,
                    key="hp_asset_price"
                )
                deposit = st.number_input(
                    "Initial Deposit (KES)",
                    min_value=0.0,
                    value=20000.0,
                    step=1000.0,
                    key="hp_deposit"
                )

            with hp2:
                loan_months = st.number_input(
                    "Loan Duration (Periods)",
                    min_value=1,
                    value=12,
                    step=1,
                    key="hp_months"
                )
                payment_frequency = st.selectbox(
                    "Repayment Frequency",
                    options=["Monthly", "Weekly"],
                    index=0,
                    key="hp_frequency"
                )

            with hp3:
                installment = st.number_input(
                    f"{payment_frequency} Installment (KES)",
                    min_value=0.0,
                    value=8000.0,
                    step=500.0,
                    key="hp_installment"
                )

            reducible_default = [t for t in ["Airtime", "Merchant Payment", "Buy Goods", "Charges"] if t in transaction_types]
            reducible_types = st.multiselect(
                "Expense categories you may be willing to reduce",
                options=transaction_types,
                default=reducible_default,
                key="hp_reducible_types"
            )

            reduction_plan = suggest_expense_reduction_plan(
                filtered_df,
                frequency=payment_frequency,
                target_installment=installment,
                reducible_types=reducible_types
            )

            ability = reduction_plan["base_ability"]
            loan_amount = max(asset_price - deposit, 0.0)

            st.markdown("---")
            st.subheader("Repayment Ability Calculation")

            k1, k2, k3, k4 = st.columns(4)
            k1.metric(f"Avg {payment_frequency} Income", f"KES {ability['avg_income']:,.2f}")
            k2.metric(f"Avg {payment_frequency} Expenses", f"KES {ability['avg_expense']:,.2f}")
            k3.metric(f"Avg {payment_frequency} Surplus", f"KES {ability['avg_surplus']:,.2f}")
            k4.metric(f"Safe {payment_frequency} Installment Now", f"KES {ability['safe_installment_now']:,.2f}")

            k5, k6, k7, k8 = st.columns(4)
            k5.metric("Loan Amount", f"KES {loan_amount:,.2f}")
            k6.metric(f"Chosen {payment_frequency} Installment", f"KES {installment:,.2f}")
            k7.metric("Coverage Ratio", f"{ability['coverage_ratio']:.2f}x")
            k8.metric("Installment / Income", f"{ability['income_ratio'] * 100:,.1f}%")

            st.markdown("### Interpretation")

            if ability["can_pay_now"]:
                st.success(
                    f"You can already support this {payment_frequency.lower()} installment from your historical cash flow."
                )
            else:
                st.error(
                    f"You cannot currently support this {payment_frequency.lower()} installment from your present cash flow."
                )

            st.write(
                f"Current affordable installment: **KES {ability['safe_installment_now']:,.2f}** per {payment_frequency.lower()}."
            )

            if reduction_plan["additional_needed"] > 0:
                st.write(
                    f"Extra amount required to reach your chosen installment: "
                    f"**KES {reduction_plan['additional_needed']:,.2f}** per {payment_frequency.lower()}."
                )
            else:
                st.write("No expense reduction is required for this installment.")

            st.markdown("### Suggested Expense Reduction")

            s1, s2, s3, s4 = st.columns(4)
            s1.metric("Additional Savings Needed", f"KES {reduction_plan['additional_needed']:,.2f}")
            s2.metric("Reducible Expense Base", f"KES {reduction_plan['total_reducible_avg']:,.2f}")
            s3.metric("Suggested Reduction %", f"{reduction_plan['recommended_reduction_pct']:.1f}%")
            s4.metric(
                "Capacity After Suggested Reduction",
                f"KES {reduction_plan['capacity_after_recommended_reduction']:,.2f}"
            )

            if reduction_plan["additional_needed"] <= 0:
                st.success("No reduction is needed because the current cash flow already supports the installment.")
            elif reduction_plan["can_fund_with_selected_categories"]:
                st.warning(
                    f"To support this installment, reduce the selected expense categories by about "
                    f"**{reduction_plan['recommended_reduction_pct']:.1f}%** across the board."
                )
            else:
                st.error(
                    "Even if you reduce all selected categories by 100%, the chosen installment still appears too high."
                )

            if not reduction_plan["category_summary"].empty:
                st.markdown("### Categories You Can Reduce")
                st.dataframe(
                    reduction_plan["category_summary"],
                    use_container_width=True,
                    hide_index=True
                )

            period_cashflow_df = build_cashflow_by_frequency(filtered_df, frequency=payment_frequency)
            if not period_cashflow_df.empty:
                st.markdown(f"### {payment_frequency} Cashflow vs Installment")
                chart_df = period_cashflow_df.copy()
                chart_df["Installment"] = installment

                fig_loan = px.bar(
                    chart_df,
                    x="Period",
                    y="Net Cashflow",
                    title=f"{payment_frequency} Net Cashflow"
                )
                fig_loan.add_scatter(
                    x=chart_df["Period"],
                    y=chart_df["Installment"],
                    mode="lines",
                    name="Installment"
                )
                st.plotly_chart(fig_loan, use_container_width=True)

    with tabs[4]:
        st.subheader("Loan Optimizer")

        if filtered_df.empty:
            st.info("No transaction data available for the current filter selection.")
        else:
            opt1, opt2 = st.columns(2)

            with opt1:
                optimizer_asset_price = st.number_input(
                    "Target Asset Price (KES)",
                    min_value=0.0,
                    value=150000.0,
                    step=1000.0,
                    key="opt_asset_price"
                )
                optimizer_frequency = st.selectbox(
                    "Optimization Frequency",
                    options=["Monthly", "Weekly"],
                    index=0,
                    key="opt_frequency"
                )

            with opt2:
                optimizer_periods = st.number_input(
                    "Repayment Periods",
                    min_value=1,
                    value=12,
                    step=1,
                    key="opt_periods"
                )

            reducible_default_opt = [t for t in ["Airtime", "Merchant Payment", "Buy Goods", "Charges"] if t in transaction_types]
            optimizer_reducible_types = st.multiselect(
                "Expense categories included for optimization",
                options=transaction_types,
                default=reducible_default_opt,
                key="opt_reducible_types"
            )

            required_installment = optimizer_asset_price / optimizer_periods if optimizer_periods > 0 else 0.0

            optimizer_plan = suggest_expense_reduction_plan(
                filtered_df,
                frequency=optimizer_frequency,
                target_installment=required_installment,
                reducible_types=optimizer_reducible_types
            )

            optimizer_ability = optimizer_plan["base_ability"]

            safe_capacity_now = optimizer_ability["safe_installment_now"]
            safe_capacity_after_max = optimizer_plan["capacity_after_max_reduction"]

            max_loan_now = safe_capacity_now * optimizer_periods
            max_loan_after_max = safe_capacity_after_max * optimizer_periods

            recommended_deposit_now = max(optimizer_asset_price - max_loan_now, 0.0)
            recommended_deposit_after_max = max(optimizer_asset_price - max_loan_after_max, 0.0)

            o1, o2, o3, o4 = st.columns(4)
            o1.metric(f"Required {optimizer_frequency} Installment", f"KES {required_installment:,.2f}")
            o2.metric(f"Safe {optimizer_frequency} Payment Now", f"KES {safe_capacity_now:,.2f}")
            o3.metric("Max Loan Now", f"KES {max_loan_now:,.2f}")
            o4.metric("Deposit Needed Now", f"KES {recommended_deposit_now:,.2f}")

            o5, o6, o7, o8 = st.columns(4)
            o5.metric("Extra Savings Needed", f"KES {optimizer_plan['additional_needed']:,.2f}")
            o6.metric("Reducible Expense Base", f"KES {optimizer_plan['total_reducible_avg']:,.2f}")
            o7.metric("Suggested Reduction %", f"{optimizer_plan['recommended_reduction_pct']:.1f}%")
            o8.metric("Deposit Needed After Max Reduction", f"KES {recommended_deposit_after_max:,.2f}")

            st.markdown("### Interpretation")

            if optimizer_ability["can_pay_now"]:
                st.success("The target asset is affordable now from your current cash flow.")
            elif optimizer_plan["can_fund_with_selected_categories"]:
                st.warning(
                    f"The asset is not affordable now, but it may become affordable if you reduce the selected categories "
                    f"by about **{optimizer_plan['recommended_reduction_pct']:.1f}%**."
                )
            else:
                st.error(
                    "The asset still appears too expensive even if all selected reducible categories are fully cut."
                )

            st.write(
                f"To finance **KES {optimizer_asset_price:,.2f}** over **{optimizer_periods}** periods, "
                f"you need about **KES {required_installment:,.2f}** per {optimizer_frequency.lower()}."
            )

            st.write(
                f"Your current safe payment is **KES {safe_capacity_now:,.2f}** per {optimizer_frequency.lower()}."
            )

            if optimizer_plan["additional_needed"] > 0:
                st.write(
                    f"You need an extra **KES {optimizer_plan['additional_needed']:,.2f}** per "
                    f"{optimizer_frequency.lower()} to support the target."
                )

            st.markdown("### Categories You Can Reduce")
            if optimizer_plan["category_summary"].empty:
                st.info("No reducible categories found in the current filter selection.")
            else:
                st.dataframe(
                    optimizer_plan["category_summary"],
                    use_container_width=True,
                    hide_index=True
                )

    with tabs[5]:
        st.subheader("Export Filtered Analysis")

        export_ability = compute_repayment_ability(
            filtered_df,
            frequency="Monthly",
            installment=0.0,
            reducible_types=[t for t in ["Airtime", "Merchant Payment", "Buy Goods", "Charges"] if t in transaction_types],
            reduction_rate=0.0
        )

        loan_summary_df = pd.DataFrame([{
            "Average Monthly Income": export_ability["avg_income"],
            "Average Monthly Expenses": export_ability["avg_expense"],
            "Average Monthly Surplus": export_ability["avg_surplus"],
            "Median Monthly Surplus": export_ability["median_surplus"],
            "Worst Monthly Surplus": export_ability["worst_surplus"],
            "Current Safe Monthly Installment": export_ability["safe_installment_now"]
        }])

        excel_data = to_excel_bytes({
            "Transactions": filtered_df,
            "Summary": summary_df,
            f"{analyze_by}_Trend": period_trend_df.drop(columns=["SortKey"], errors="ignore"),
            "Transaction_Types": transaction_type_df,
            "Top_Counterparties": top_counterparties_df,
            "Weekday_Trend": weekday_df,
            "Balance_Trend": balance_trend_df,
            "Recurring_Payments": recurring_df,
            "Loan_Summary": loan_summary_df
        })

        st.download_button(
            label="Download Filtered Excel Report",
            data=excel_data,
            file_name="mpesa_filtered_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )