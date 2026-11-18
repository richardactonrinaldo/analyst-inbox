# Analyst Inbox

Students in HIFN1009 chat with **Dana Morgan**, a pretend Manager of Decision Support at a fictional Ontario hospital network. Dana sends analysis requests about a 300-row emergency department dataset, checks students' numbers against the data, gives hints instead of answers, and asks follow-up questions.

Weekly focus options: descriptive statistics, PivotTables, charts (students can attach screenshots), and presenting results.

## Deploy on Streamlit Community Cloud (free)

1. **Get an API key.** Sign in at platform.claude.com, buy a small amount of credit (e.g. $20 USD), leave auto-reload **off** so that amount is a hard cap, and create an API key.
2. **Deploy.** Sign in at share.streamlit.io with GitHub, click **Create app**, choose this repo, branch `main`, main file `app.py`.
3. **Add secrets.** In the app's **Settings > Secrets**, paste:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   CLASS_CODE = "your-class-code"
   MODEL = "claude-sonnet-5-5"
   ```
4. **Share.** Post the app URL and the class code on Blackboard. Change `CLASS_CODE` in Secrets each term (or any time it leaks).

Usage and cost appear in the Claude Console under Usage.

## Cost controls

- The class code keeps strangers from using your credits.
- Each session is capped at 40 student messages (`MAX_STUDENT_MESSAGES` in `app.py`).
- Prepaid credits with auto-reload off can't be exceeded.
- `MODEL = "claude-haiku-5-5"` cuts cost by roughly 20x.

## Run locally

```bash
pip install -r requirements.txt
cp .streamlit/secrets.example.toml .streamlit/secrets.toml   # then edit it
streamlit run app.py
```

## Change the data

`python make_dataset.py` regenerates `data/ED_Visits_2025.csv` (fixed seed, same result every time). To use your own file, replace the CSV and update `SCHEMA` and `NUMERIC` in `app.py` to match its columns.
