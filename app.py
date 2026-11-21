import os
from io import BytesIO
import pandas as pd
import streamlit as st
from groq import Groq

st.set_page_config(page_title="P P - AI Employee", page_icon="🤖", layout="wide")
st.title("🤖 P P AI Employee")
st.caption("Safe Mode — Original files are never overwritten.")

DEPARTMENTS = ["General Assistant", "Excel Specialist", "Daily Report Officer", "Data Analyst", "Reporting Officer", "Office Assistant"]
RESOLVED = {"Resolved", "Resolved (No KPI)", "Resolved (AP Up)", "Resolved (ONU Up)"}

def audit_excel(upload):
    book = pd.ExcelFile(upload)
    summary, findings = [], []
    for sheet in book.sheet_names:
        df = pd.read_excel(book, sheet_name=sheet)
        task_col = next((c for c in df.columns if str(c).strip().lower() == "task no"), None)
        if task_col is None:
            summary.append({"Sheet": sheet, "Rows": len(df), "Task rows": 0, "Notes": "Task No column not found"})
            continue
        task_valid = df[task_col].notna() & df[task_col].astype(str).str.strip().ne("")
        repay = df.get("Repay Status", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
        leave = repay.eq("Engr Leave")
        fields = ["LAN Engineer Name", "Car ID", "Task Status"]
        issue_counts = {}
        for field in fields:
            if field in df.columns:
                missing = task_valid & df[field].isna() & ~leave
                issue_counts[field] = int(missing.sum())
                for ix in df.index[missing]:
                    findings.append({"Sheet": sheet, "Excel Row": int(ix)+2, "Task No": df.at[ix,task_col], "Field": field, "Issue": "Missing value"})
        # Duplicate Task No is intentionally not treated as an error under the user's audit rules.
        summary.append({"Sheet": sheet, "Rows": len(df), "Task rows": int(task_valid.sum()),
                        "Missing Engineer": issue_counts.get("LAN Engineer Name",0),
                        "Missing Car ID": issue_counts.get("Car ID",0),
                        "Missing Task Status": issue_counts.get("Task Status",0),
                        "Notes": "Duplicate Task No not audited; Engr Leave rows excluded"})
    return summary, findings

def audit_to_xlsx(summary, findings):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame(summary).to_excel(writer, sheet_name="Audit Summary", index=False)
        pd.DataFrame(findings, columns=["Sheet","Excel Row","Task No","Field","Issue"]).to_excel(writer, sheet_name="Missing Data", index=False)
    out.seek(0)
    return out

api_key = st.secrets.get("GROQ_API_KEY", os.getenv("GROQ_API_KEY", ""))
with st.sidebar:
    st.header("P P Settings")
    department = st.selectbox("Department", DEPARTMENTS)
    model = st.selectbox("Model", ["openai/gpt-oss-120b", "llama-3.3-70b-versatile"], index=0)
    st.caption("Safe Mode: review and export a corrected copy. Original files are not changed.")

tab_chat, tab_excel = st.tabs(["💬 Chat", "📊 Excel Specialist"])

with tab_chat:
    if "messages" not in st.session_state:
        st.session_state.messages = []
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
    prompt = st.chat_input("P P ကို အလုပ်တစ်ခုခိုင်းပါ...")
    if prompt:
        st.session_state.messages.append({"role":"user","content":prompt})
        with st.chat_message("user"): st.markdown(prompt)
        answer = "API ချိတ်ဆက်မှုမရှိပါ။ GROQ_API_KEY ကို Streamlit Secrets တွင် ထည့်ပါ။"
        if api_key:
            try:
                client = Groq(api_key=api_key)
                history = [{"role":"system","content":f"You are P P, a careful AI employee. Department: {department}. Answer in Burmese when appropriate. Never claim actions you did not perform. In Safe Mode, do not overwrite originals or take external actions; provide a separate copy and ask permission for consequential changes."}]
                history += [{"role":m["role"],"content":m["content"]} for m in st.session_state.messages[-12:]]
                resp = client.chat.completions.create(model=model, messages=history, temperature=0.2)
                answer = resp.choices[0].message.content
            except Exception as e:
                answer = f"AI ချိတ်ဆက်မှု အမှား: {e}"
        st.session_state.messages.append({"role":"assistant","content":answer})
        with st.chat_message("assistant"): st.markdown(answer)

with tab_excel:
    st.subheader("Excel Audit — Safe Mode")
    st.write("Workbook ထဲက Sheet များကို စစ်ဆေးပြီး Missing Engineer, Car ID, Task Status ကို စာရင်းထုတ်ပေးမည်။ မူရင်းဖိုင်ကို မပြင်ပါ။")
    uploaded = st.file_uploader("Excel Workbook တင်ပါ", type=["xlsx","xlsm"], key="audit_file")
    if uploaded:
        try:
            summary, findings = audit_excel(uploaded)
            st.subheader("Audit Summary")
            st.dataframe(pd.DataFrame(summary), use_container_width=True)
            if findings:
                st.subheader("Missing Data Findings")
                st.dataframe(pd.DataFrame(findings), use_container_width=True)
            else:
                st.success("စစ်ဆေးထားသော fields များတွင် Missing data မတွေ့ပါ။")
            st.download_button("📥 Audit Report Excel Download", data=audit_to_xlsx(summary,findings),
                file_name="PP_Excel_Audit_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Excel ဖိုင်စစ်ဆေးရာတွင် အမှား: {e}")
