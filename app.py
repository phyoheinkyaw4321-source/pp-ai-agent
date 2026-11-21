import os
from io import BytesIO

import pandas as pd
import streamlit as st
from groq import Groq

st.set_page_config(page_title="P P - AI Employee", page_icon="🤖", layout="wide")
st.title("🤖 P P AI Employee")
st.caption("Safe Mode — Original files are never overwritten.")

DEPARTMENTS = [
    "General Assistant",
    "Excel Specialist",
    "Daily Report Officer",
    "Data Analyst",
    "Reporting Officer",
    "Office Assistant",
]

# Only these four columns are counted as Missing Data.
MISSING_FIELDS = [
    "LAN Engineer Name",
    "Car ID",
    "Task Status",
    "City",
]
EXEMPT_REPAY_STATUSES = {"engr leave", "route cancel"}

FINDING_COLUMNS = [
    "Sheet",
    "Excel Row",
    "Task No",
    "Field",
    "Issue",
    "Repay Status (Context)",
    "Reason",
    "Recommended Action",
]


def find_column(df, wanted_name):
    """Find a column without depending on capitalization or extra spaces."""
    wanted = str(wanted_name).strip().casefold()
    return next(
        (col for col in df.columns if str(col).strip().casefold() == wanted),
        None,
    )


def is_missing(value):
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def audit_excel(upload):
    """Audit workbook rows using the user's defined Missing Data rules."""
    upload.seek(0)
    book = pd.ExcelFile(upload)
    summary = []
    findings = []

    for sheet in book.sheet_names:
        df = pd.read_excel(book, sheet_name=sheet)
        task_col = find_column(df, "Task No")
        repay_col = find_column(df, "Repay Status")

        if task_col is None:
            summary.append({
                "Sheet": sheet,
                "Rows": len(df),
                "Task rows": 0,
                "Missing Engineer": 0,
                "Missing Car ID": 0,
                "Missing Task Status": 0,
                "Missing City": 0,
                "Notes": "Task No column not found; rows were not audited",
            })
            continue

        task_valid = (
            df[task_col].notna()
            & df[task_col].astype(str).str.strip().ne("")
        )
        issue_counts = {field: 0 for field in MISSING_FIELDS}

        for field in MISSING_FIELDS:
            field_col = find_column(df, field)
            if field_col is None:
                continue

            for ix in df.index:
                if not bool(task_valid.loc[ix]):
                    continue

                value = df.at[ix, field_col]
                if not is_missing(value):
                    continue

                repay_status = ""
                if repay_col is not None:
                    repay_value = df.at[ix, repay_col]
                    if not is_missing(repay_value):
                        repay_status = str(repay_value).strip()

                # Engr Leave and Route Cancel exempt only these three fields.
                # City is always checked, regardless of Repay Status.
                if (
                    field != "City"
                    and repay_status.casefold() in EXEMPT_REPAY_STATUSES
                ):
                    continue

                issue_counts[field] += 1

                if field == "LAN Engineer Name":
                    reason = "LAN Engineer Name is blank in a row with a Task No."
                    action = "Verify the correct Engineer Name with the responsible team."
                elif field == "Car ID":
                    reason = "Car ID is blank in a row with a Task No."
                    action = "Verify the Car ID with the responsible team. Do not guess or auto-fill it."
                elif field == "Task Status":
                    reason = "Task Status is blank in a row with a Task No."
                    action = "Check the original task record and confirm the correct Task Status."
                else:
                    reason = "City is blank. City is required even for Engr Leave and Route Cancel rows."
                    action = "Verify the correct City with the responsible team."

                findings.append({
                    "Sheet": sheet,
                    "Excel Row": int(df.index.get_loc(ix)) + 2,
                    "Task No": df.at[ix, task_col],
                    "Field": field,
                    "Issue": "Missing Data",
                    "Repay Status (Context)": repay_status,
                    "Reason": reason,
                    "Recommended Action": action,
                })

        summary.append({
            "Sheet": sheet,
            "Rows": len(df),
            "Task rows": int(task_valid.sum()),
            "Missing Engineer": issue_counts["LAN Engineer Name"],
            "Missing Car ID": issue_counts["Car ID"],
            "Missing Task Status": issue_counts["Task Status"],
            "Missing City": issue_counts["City"],
            "Notes": (
                "Repay Status is context only; Engr Leave and Route Cancel "
                "exempt Engineer Name, Car ID and Task Status; City is always checked; "
                "blank Task No rows and duplicate Task No are not errors."
            ),
        })

    return summary, findings


def audit_to_xlsx(summary, findings):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(summary).to_excel(
            writer, sheet_name="Audit Summary", index=False
        )
        pd.DataFrame(findings, columns=FINDING_COLUMNS).to_excel(
            writer, sheet_name="Missing Data", index=False
        )
    output.seek(0)
    return output


api_key = st.secrets.get("GROQ_API_KEY", os.getenv("GROQ_API_KEY", ""))

with st.sidebar:
    st.header("P P Settings")
    department = st.selectbox("Department", DEPARTMENTS)
    model = st.selectbox(
        "Model",
        ["openai/gpt-oss-120b", "llama-3.3-70b-versatile"],
        index=0,
    )
    st.caption(
        "Safe Mode: review and export a corrected copy. "
        "Original files are not changed."
    )

# Keep the latest audit available to Chat during the current Streamlit session.
if "messages" not in st.session_state:
    st.session_state.messages = []
if "audit_summary" not in st.session_state:
    st.session_state.audit_summary = None
if "audit_findings" not in st.session_state:
    st.session_state.audit_findings = None


tab_chat, tab_excel = st.tabs(["💬 Chat", "📊 Excel Specialist"])

with tab_chat:
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

                audit_context = "လက်ရှိ Session တွင် Excel Audit ရလဒ် မရှိသေးပါ။"
                if st.session_state.audit_summary is not None:
                    summary_text = pd.DataFrame(
                        st.session_state.audit_summary
                    ).to_string(index=False)
                    findings_df = pd.DataFrame(
                        st.session_state.audit_findings or [],
                        columns=FINDING_COLUMNS,
                    )
                    findings_text = (
                        findings_df.head(100).to_string(index=False)
                        if not findings_df.empty
                        else "Missing Data findings မတွေ့ပါ။"
                    )
                    audit_context = (
                        "လက်ရှိ Session ၏ အမှန်တကယ် Excel Audit Summary:\n"
                        + summary_text
                        + "\n\nAudit Findings (ပထမ 100 ခုအထိ):\n"
                        + findings_text
                    )

                system_prompt = f"""You are P P, a careful AI employee.
Department: {department}
Answer in Burmese when appropriate.

EXCEL AUDIT RULES — follow exactly:
1. Only LAN Engineer Name, Car ID, Task Status and City count as Missing Data.
2. Repay Status is NOT Missing Data. It is context only to explain the row.
3. If Repay Status is Engr Leave or Route Cancel, do not count missing
   LAN Engineer Name, Car ID or Task Status as errors.
4. City must ALWAYS be checked, including Engr Leave and Route Cancel rows.
5. Rows with blank Task No are excluded from these Missing Data checks.
6. Duplicate Task No is intentionally not treated as an error.
7. Never guess or automatically fill a missing Car ID.
8. Explain the actual rule used, the finding, and the recommended action.
9. When actual audit results are present, use them. Mention Sheet, Excel Row,
   Task No, missing field and Repay Status context when available.
10. Clearly distinguish observed facts from possible explanations. Do not invent
    findings or claim a workbook was audited if no audit results are available.
11. Do not overwrite original files or take external actions without permission.

CURRENT AUDIT DATA:
{audit_context}
"""
                history = [{"role": "system", "content": system_prompt}]
                history += [
                    {"role": m["role"], "content": m["content"]}
                    for m in st.session_state.messages[-12:]
                ]

                response = client.chat.completions.create(
                    model=model,
                    messages=history,
                    temperature=0.2,
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
        "Workbook ထဲက Sheet များကို စစ်ဆေးပြီး Engineer Name, Car ID, "
        "Task Status နှင့် City Missing Data များကို စာရင်းထုတ်ပေးမည်။ "
        "Repay Status ကို အကြောင်းရင်းသိရန်သာ အသုံးပြုမည်။ မူရင်းဖိုင်ကို မပြင်ပါ။"
    )

    uploaded = st.file_uploader(
        "Excel Workbook တင်ပါ",
        type=["xlsx", "xlsm"],
        key="audit_file",
    )

    if uploaded is not None:
        try:
            summary, findings = audit_excel(uploaded)
            st.session_state.audit_summary = summary
            st.session_state.audit_findings = findings

            st.subheader("Audit Summary")
            st.dataframe(pd.DataFrame(summary), use_container_width=True)

            if findings:
                st.subheader("Missing Data Findings")
                st.dataframe(
                    pd.DataFrame(findings, columns=FINDING_COLUMNS),
                    use_container_width=True,
                )
            else:
                st.success("စစ်ဆေးထားသော fields များတွင် Missing Data မတွေ့ပါ။")

            st.download_button(
                "📥 Audit Report Excel Download",
                data=audit_to_xlsx(summary, findings),
                file_name="PP_Excel_Audit_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.info(
                "Excel Audit ရလဒ်ကို Chat မှာလည်း မေးမြန်းနိုင်ပါပြီ။ "
                "တူညီတဲ့ Session အတွင်း Audit လုပ်ပြီးနောက် Chat ကို အသုံးပြုပါ။"
            )
        except Exception as exc:
            st.error(f"Excel ဖိုင်စစ်ဆေးရာတွင် အမှား: {exc}")
