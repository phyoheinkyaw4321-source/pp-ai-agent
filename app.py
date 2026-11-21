import os
import re
from io import BytesIO

import pandas as pd
import streamlit as st
from groq import Groq

st.set_page_config(page_title="P P - AI Employee", page_icon="🤖", layout="wide")
st.title("🤖 P P AI Employee")
st.caption("Safe Mode — Original files are never overwritten.")

DEPARTMENTS = ["General Assistant", "Excel Specialist", "Daily Report Officer", "Data Analyst", "Reporting Officer", "Office Assistant"]

AUDIT_FIELDS = {
    "LAN Engineer Name": "Engineer Name",
    "Car ID": "Car ID",
    "Task Status": "Task Status",
    "Repay Status": "Repay Status",
    "City": "City",
}
EXEMPT_REPAY_STATUSES = {"engr leave", "route cancel"}


def clean_text(value):
    """Normalize text for comparisons without changing the source workbook."""
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def find_column(columns, wanted):
    wanted_key = clean_text(wanted).casefold()
    for column in columns:
        if clean_text(column).casefold() == wanted_key:
            return column
    return None


def audit_excel(upload):
    """
    Audit five requested columns.
    Engr Leave / Route Cancel exempt blanks in Engineer Name, Car ID and Task Status.
    Repay Status and City blanks are always reported.
    No source values are modified.
    """
    book = pd.ExcelFile(upload)
    summary = []
    findings = []

    for sheet_name in book.sheet_names:
        df = pd.read_excel(book, sheet_name=sheet_name)
        task_col = find_column(df.columns, "Task No")
        found_cols = {field: find_column(df.columns, field) for field in AUDIT_FIELDS}

        # Per the user's audit rule, rows with blank Task No are excluded from
        # field-error counts. Fully empty rows are therefore excluded as well.
        if task_col is not None:
            populated_row = df[task_col].map(clean_text).ne("")
        else:
            populated_row = df.apply(
                lambda row: any(clean_text(v) != "" for v in row.tolist()), axis=1
            )

        repay_col = found_cols["Repay Status"]
        if repay_col:
            repay_values = df[repay_col].map(clean_text).str.casefold()
        else:
            repay_values = pd.Series("", index=df.index, dtype="object")

        exempt_row = repay_values.isin(EXEMPT_REPAY_STATUSES)
        counts = {label: 0 for label in AUDIT_FIELDS.values()}
        sheet_findings_start = len(findings)

        for source_field, display_field in AUDIT_FIELDS.items():
            col = found_cols[source_field]
            if col is None:
                # A missing required column is a workbook-structure error, not a row error.
                findings.append({
                    "Sheet": sheet_name,
                    "Excel Row": "Header",
                    "Task No": "",
                    "Field": display_field,
                    "Error Code": "COLUMN_NOT_FOUND",
                    "Severity": "Critical",
                    "Current Value": "",
                    "Exact Error": f"Required column '{source_field}' was not found in this sheet.",
                    "Why It Is an Error": "P P cannot audit this field because the required column is missing.",
                    "Recommended Action": f"Check the Excel header. Expected column: {source_field}",
                })
                counts[display_field] += 1
                continue

            missing_mask = df[col].map(clean_text).eq("") & populated_row

            # Exempt only these three fields for the two specified Repay Status values.
            if source_field in {"LAN Engineer Name", "Car ID", "Task Status"}:
                missing_mask &= ~exempt_row

            for idx in df.index[missing_mask]:
                excel_row = int(df.index.get_loc(idx)) + 2
                task_value = clean_text(df.at[idx, task_col]) if task_col is not None else ""
                repay_value = clean_text(df.at[idx, repay_col]) if repay_col is not None else ""
                counts[display_field] += 1

                if source_field == "LAN Engineer Name":
                    detail = "Engineer Name is blank."
                    action = "Verify the correct Engineer Name with the responsible team; do not guess."
                elif source_field == "Car ID":
                    detail = "Car ID is blank."
                    action = "Report as Missing Car ID and ask the responsible team to verify; do not fill a guessed ID."
                elif source_field == "Task Status":
                    detail = "Task Status is blank."
                    action = "Verify the actual Task Status from the source system."
                elif source_field == "Repay Status":
                    detail = "Repay Status is blank."
                    action = "Verify the correct Repay Status. This field is still checked for Engr Leave and Route Cancel rows."
                else:
                    detail = "City is blank."
                    action = "Verify the correct City against the source record; do not infer it from another field."

                findings.append({
                    "Sheet": sheet_name,
                    "Excel Row": excel_row,
                    "Task No": task_value,
                    "Field": display_field,
                    "Error Code": "MISSING_VALUE",
                    "Severity": "Critical",
                    "Current Value": "",
                    "Exact Error": detail,
                    "Why It Is an Error": (
                        f"Repay Status = '{repay_value}'. "
                        "This blank is not exempt under the configured audit rules."
                        if source_field in {"LAN Engineer Name", "Car ID", "Task Status"}
                        else detail
                    ),
                    "Recommended Action": action,
                })

        task_rows = int(populated_row.sum())
        missing_task_no = 0
        if task_col is not None:
            missing_task_no = int(
                (df.apply(lambda row: any(clean_text(v) != "" for v in row.tolist()), axis=1)
                 & df[task_col].map(clean_text).eq("")).sum()
            )

        summary_row = {
            "Sheet": sheet_name,
            "Data Rows": task_rows,
            "Engineer Name Errors": counts["Engineer Name"],
            "Car ID Errors": counts["Car ID"],
            "Task Status Errors": counts["Task Status"],
            "Repay Status Errors": counts["Repay Status"],
            "City Errors": counts["City"],
            "Total Errors": sum(counts.values()),
            "Blank Task No (Excluded)": missing_task_no,
            "Notes": (
                "Engr Leave / Route Cancel blanks exempt for Engineer Name, Car ID, Task Status only. "
                "Repay Status and City are always checked. Duplicate Task No is not treated as an error."
            ),
        }
        summary.append(summary_row)

    return summary, findings


def audit_to_xlsx(summary, findings):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(summary).to_excel(writer, sheet_name="Audit Summary", index=False)
        pd.DataFrame(findings, columns=[
            "Sheet", "Excel Row", "Task No", "Field", "Error Code", "Severity",
            "Current Value", "Exact Error", "Why It Is an Error", "Recommended Action"
        ]).to_excel(writer, sheet_name="Detailed Errors", index=False)
    output.seek(0)
    return output


api_key = st.secrets.get("GROQ_API_KEY", os.getenv("GROQ_API_KEY", ""))
with st.sidebar:
    st.header("P P Settings")
    department = st.selectbox("Department", DEPARTMENTS)
    model = st.selectbox("Model", ["openai/gpt-oss-120b", "llama-3.3-70b-versatile"], index=0)
    st.caption("Safe Mode: original files are not changed.")

tab_chat, tab_excel = st.tabs(["💬 Chat", "📊 Excel Specialist"])

with tab_chat:
    if "messages" not in st.session_state:
        st.session_state.messages = []
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
    prompt = st.chat_input("P P ကို အလုပ်တစ်ခုခိုင်းပါ...")
    if prompt:
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        answer = "API ချိတ်ဆက်မှုမရှိပါ။ GROQ_API_KEY ကို Streamlit Secrets တွင် ထည့်ပါ။"
        if api_key:
            try:
                client = Groq(api_key=api_key)
                history = [{
                    "role": "system",
                    "content": (
                        f"You are P P, a careful AI employee. Department: {department}. "
                        "Answer in Burmese when appropriate. Never claim actions you did not perform. "
                        "In Safe Mode, never overwrite original files or take external actions without permission."
                    ),
                }]
                history += [
                    {"role": m["role"], "content": m["content"]}
                    for m in st.session_state.messages[-12:]
                ]
                response = client.chat.completions.create(
                    model=model, messages=history, temperature=0.2
                )
                answer = response.choices[0].message.content
            except Exception as exc:
                answer = f"AI ချိတ်ဆက်မှု အမှား: {exc}"
        st.session_state.messages.append({"role": "assistant", "content": answer})
        with st.chat_message("assistant"):
            st.markdown(answer)

with tab_excel:
    st.subheader("Excel Audit — Safe Mode")
    st.write(
        "စစ်ဆေးမည့် Column ၅ ခု: Engineer Name, Car ID, Task Status, "
        "Repay Status, City။ မူရင်းဖိုင်ကို မပြင်ပါ။"
    )
    st.info(
        "Engr Leave / Route Cancel ဖြစ်ပါက Engineer Name, Car ID, Task Status "
        "Blank ကို Error မသတ်မှတ်ပါ။ Repay Status နှင့် City ကိုတော့ အမြဲစစ်ဆေးပါမည်။"
    )
    uploaded = st.file_uploader(
        "Excel Workbook တင်ပါ", type=["xlsx", "xlsm"], key="audit_file"
    )
    if uploaded:
        try:
            summary, findings = audit_excel(uploaded)
            summary_df = pd.DataFrame(summary)
            findings_df = pd.DataFrame(findings)

            st.subheader("Audit Summary")
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

            total_errors = int(summary_df["Total Errors"].sum()) if not summary_df.empty else 0
            if total_errors:
                st.error(f"စစ်ဆေးမှုတွင် Error {total_errors} ခု တွေ့ရှိပါသည်။")
                st.subheader("Detailed Error Findings")
                st.caption(
                    "Error တစ်ခုချင်းစီတွင် Sheet, Excel Row, Task No, Field, "
                    "အမှားအကြောင်းရင်းနှင့် ပြင်ဆင်ရန်အကြံပြုချက် ပါဝင်သည်။"
                )
                st.dataframe(findings_df, use_container_width=True, hide_index=True)
            else:
                st.success("သတ်မှတ်ထားသော Column ၅ ခုတွင် Missing Data Error မတွေ့ပါ။")

            st.download_button(
                "📥 Detailed Audit Report Excel Download",
                data=audit_to_xlsx(summary, findings),
                file_name="PP_Detailed_Excel_Audit_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"Excel ဖိုင်စစ်ဆေးရာတွင် အမှား: {exc}")
