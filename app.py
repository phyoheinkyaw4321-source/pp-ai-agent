import os
from io import BytesIO

import pandas as pd
import streamlit as st

try:
    from groq import Groq
except Exception:
    Groq = None

st.set_page_config(page_title="P P - AI Agent", page_icon="🤖", layout="wide")

st.markdown('''
<style>
    .stApp { background: #f4f6f8; }
    [data-testid="stHeader"] { background: rgba(0,0,0,0); }
    .main-title { color: #17212b; font-size: 2rem; font-weight: 750; }
    .sub-title { color: #64748b; margin-top: -8px; }
    div[data-testid="stMetric"] {
        background: white; border: 1px solid #e5e7eb;
        padding: 14px; border-radius: 12px;
    }
</style>
''', unsafe_allow_html=True)

st.markdown('<div class="main-title">🤖 P P AI Agent</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Your personal work assistant · Version 1.0</div>', unsafe_allow_html=True)


def get_api_key():
    try:
        return st.secrets.get("GROQ_API_KEY", "")
    except Exception:
        return os.getenv("GROQ_API_KEY", "")


def workbook_summary(file_bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    summaries = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, dtype=object)
        summaries.append({
            "Sheet": sheet,
            "Rows": len(df),
            "Columns": len(df.columns),
            "Blank cells": int(df.isna().sum().sum()),
            "Duplicate rows": int(df.duplicated().sum()) if len(df.columns) else 0,
            "Column names": ", ".join(map(str, df.columns)),
        })
    return pd.DataFrame(summaries)


def basic_audit(file_bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    reports = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, dtype=object)
        if df.empty and len(df.columns) == 0:
            reports.append({
                "Sheet": sheet, "Rows": 0, "Blank cells": 0,
                "Duplicate rows": 0, "Completely blank rows": 0,
            })
            continue
        reports.append({
            "Sheet": sheet,
            "Rows": len(df),
            "Blank cells": int(df.isna().sum().sum()),
            "Duplicate rows": int(df.duplicated().sum()) if len(df.columns) else 0,
            "Completely blank rows": int(df.isna().all(axis=1).sum()) if len(df.columns) else 0,
        })
    return pd.DataFrame(reports)


def ask_groq(messages, api_key, model):
    if Groq is None:
        raise RuntimeError("Groq package မတင်ရသေးပါ။ requirements.txt ကို install လုပ်ပါ။")
    client = Groq(api_key=api_key)
    response = client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=0.2,
    )
    return response.choices[0].message.content


with st.sidebar:
    st.header("⚙️ Settings")
    department = st.selectbox(
        "Department",
        ["General", "Operations", "R1", "R1S", "R2", "R2S", "MDY", "Naypyitaw"],
    )
    # Groq's supported replacement for the retired Llama 3.3 70B model.
    model = st.text_input("Groq model", value="openai/gpt-oss-120b")
    st.caption("API key ကို Streamlit Secrets ထဲမှာ GROQ_API_KEY အဖြစ် ထည့်ပါ။")
    if st.button("New Chat", use_container_width=True):
        st.session_state.chat_messages = []
        st.rerun()

if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = []

tab_chat, tab_excel, tab_help = st.tabs(["💬 AI Assistant", "📊 Excel Audit", "ℹ️ Setup"])

with tab_chat:
    st.subheader("Chat with P P")
    for msg in st.session_state.chat_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
    prompt = st.chat_input("P P ကို အလုပ်တစ်ခု ပေးပါ...")
    if prompt:
        st.session_state.chat_messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        api_key = get_api_key()
        with st.chat_message("assistant"):
            if not api_key:
                answer = (
                    "Groq API Key မတွေ့ပါ။ Streamlit Cloud ရဲ့ App settings → Secrets "
                    "ထဲမှာ GROQ_API_KEY ကို ထည့်ပြီး App ကို ပြန်ဖွင့်ပါ။"
                )
                st.warning(answer)
            else:
                system_prompt = (
                    "You are P P, a careful AI work assistant. "
                    "Help with Excel, reporting, data analysis, meetings, and office tasks. "
                    f"Current department: {department}. "
                    "Be honest about uncertainty. Never claim you edited or verified a file "
                    "unless the code actually did so. Explain important assumptions."
                )
                messages = [{"role": "system", "content": system_prompt}]
                messages.extend(st.session_state.chat_messages[-12:])
                try:
                    answer = ask_groq(messages, api_key, model)
                    st.markdown(answer)
                except Exception as exc:
                    answer = f"AI ချိတ်ဆက်မှုမှာ အမှားရှိနေပါတယ်။ အသေးစိတ်: {exc}"
                    st.error(answer)
        st.session_state.chat_messages.append({"role": "assistant", "content": answer})

with tab_excel:
    st.subheader("Excel Workbook Audit")
    st.write(
        "Excel ဖိုင်တင်ပြီး Sheet တစ်ခုချင်းစီရဲ့ Row/Column အရေအတွက်၊ "
        "Blank cells၊ Duplicate rows တို့ကို စစ်ဆေးနိုင်ပါတယ်။ မူရင်းဖိုင်ကို မပြင်ပါ။"
    )
    uploaded = st.file_uploader(
        "Excel ဖိုင်တင်ပါ",
        type=["xlsx", "xlsm"],
        help="စမ်းသပ်ရာတွင် မူရင်းဖိုင်အစား Copy ဖိုင်ကို အသုံးပြုပါ။",
    )
    if uploaded:
        file_bytes = uploaded.getvalue()
        try:
            with st.spinner("Workbook ကို စစ်ဆေးနေပါတယ်..."):
                summary = workbook_summary(file_bytes)
                audit = basic_audit(file_bytes)
            total_rows = int(summary["Rows"].sum())
            total_sheets = len(summary)
            total_blanks = int(summary["Blank cells"].sum())
            c1, c2, c3 = st.columns(3)
            c1.metric("Sheets", total_sheets)
            c2.metric("Total data rows", f"{total_rows:,}")
            c3.metric("Blank cells", f"{total_blanks:,}")
            st.markdown("#### Sheet Summary")
            st.dataframe(summary, use_container_width=True, hide_index=True)
            st.markdown("#### Basic Audit Results")
            st.dataframe(audit, use_container_width=True, hide_index=True)
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                summary.to_excel(writer, sheet_name="Workbook Summary", index=False)
                audit.to_excel(writer, sheet_name="Basic Audit", index=False)
            st.download_button(
                "⬇️ Download Audit Report",
                data=output.getvalue(),
                file_name="PP_Excel_Audit_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.info(
                "မှတ်ချက်: ဒီ Version 1.0 က အခြေခံ Workbook Audit ဖြစ်ပါတယ်။ "
                "မင်းရဲ့ လုပ်ငန်းသတ်မှတ်ချက်များအတိုင်း Engineer, Car ID, Task Status, "
                "Task No နှင့် Daily Report KPI စစ်ဆေးမှုများကို နောက်အဆင့်တွင် ထည့်သွင်းမယ်။"
            )
        except Exception as exc:
            st.error(f"Excel ဖိုင်ဖတ်ရာတွင် အမှားရှိနေပါတယ်: {exc}")

with tab_help:
    st.subheader("P P AI Agent — Version 1.0")
    st.markdown('''
**ပါဝင်ပြီးသား**
- AI chat (Groq API ချိတ်ဆက်ပြီးမှ အသုံးပြုနိုင်သည်)
- Department selection
- Excel workbook sheet summary
- Basic blank-cell / duplicate-row audit
- Audit result ကို Excel ဖိုင်အဖြစ် download လုပ်နိုင်ခြင်း

**လက်ရှိ မပါသေးသောအရာများ**
- မင်းရဲ့ Daily Report စည်းမျဉ်းအပြည့်အစုံ
- Engineer / Car ID / Task Status အလိုက် custom audit
- Auto_Pivot နှင့် Report Excel တို့ကို အလိုအလျောက် နှိုင်းယှဉ်ခြင်း
- Background / scheduled agent tasks

ဒီအရာတွေကို မင်းရဲ့ နမူနာ Excel ဖိုင်နဲ့ စမ်းသပ်ပြီး အဆင့်လိုက် ထည့်သွင်းရမယ်။
''')
