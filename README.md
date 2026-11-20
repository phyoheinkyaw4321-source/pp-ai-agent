# P P AI Agent

P P AI Agent is a Streamlit-based personal work assistant.

## Features in Version 1.0
- Groq-powered AI chat (requires an API key)
- Department selection
- Excel workbook summary across worksheets
- Basic blank-cell and duplicate-row audit
- Downloadable Excel audit report

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Configure Groq API key
Create `.streamlit/secrets.toml`:
```toml
GROQ_API_KEY = "paste-your-own-key-here"
```

Do not upload `secrets.toml` or share your API key.

## Deploy to Streamlit Community Cloud
1. Push this repository to GitHub.
2. Create an app at https://share.streamlit.io/
3. Select this repository and `app.py`.
4. Add `GROQ_API_KEY` in the app's Secrets settings.
5. Deploy.

## Important
This is a starter version. It does not yet implement the user's full Daily Report rules or guarantee unlimited Excel size. Actual capacity depends on workbook complexity and available memory.
