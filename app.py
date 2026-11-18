"""Analyst Inbox: students chat with Dana Morgan, a pretend Decision Support manager.

Runs on Streamlit. Needs these secrets (Streamlit Cloud > App > Settings > Secrets):
    ANTHROPIC_API_KEY = "sk-ant-..."
    CLASS_CODE = "something-you-give-students"   # keeps strangers from using your credits
    MODEL = "claude-sonnet-5-5"                   # optional; "claude-haiku-5-5" is cheaper
"""
import base64
import io
import json
import re
from datetime import datetime
from pathlib import Path

import anthropic
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

DATA_PATH = Path(__file__).parent / "data" / "ED_Visits_2025.csv"
MAX_STUDENT_MESSAGES = 40      # per session; caps cost if someone leaves a script running
MAX_WORKBOOK_CHARS = 20_000    # how much of an uploaded workbook Dana reads
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

st.set_page_config(page_title="Analyst Inbox", page_icon="📨", layout="wide")


@st.cache_data
def load_data() -> pd.DataFrame:
    return pd.read_csv(DATA_PATH)


DF = load_data()
NUMERIC = ["Age", "TriageLevel", "WaitTimeMin", "LengthOfStayHrs", "CostCAD"]


# ---------------------------------------------------------------- data helpers
def apply_filters(df: pd.DataFrame, filters: list | None) -> pd.DataFrame:
    for f in filters or []:
        col, op, val = f["column"], f["op"], f["value"]
        if col not in df.columns:
            raise ValueError(f"Unknown column {col}")
        if col in NUMERIC:
            s, val = df[col], float(val)
        else:
            s, val = df[col].astype(str).str.lower(), str(val).lower()
        mask = {"=": s == val, "!=": s != val, ">": s > val, "<": s < val,
                ">=": s >= val, "<=": s <= val}[op]
        df = df[mask]
    return df


def query_data(metric: str, column: str | None = None, filters: list | None = None,
               group_by: str | None = None) -> dict:
    """Compute a statistic the same way Excel does (STDEV.S, MEDIAN)."""
    df = apply_filters(DF, filters)
    if metric != "count" and column not in NUMERIC:
        raise ValueError(f"{metric} needs a numeric column: {', '.join(NUMERIC)}")

    def calc(d):
        if metric == "count":
            return int(len(d))
        s = d[column]
        v = {"sum": s.sum(), "mean": s.mean(), "median": s.median(), "min": s.min(),
             "max": s.max(), "stdev": s.std(ddof=1)}[metric]
        return None if pd.isna(v) else round(float(v), 2)

    if not group_by:
        return {"rows_matched": int(len(df)), "result": calc(df)}
    if group_by not in DF.columns:
        raise ValueError(f"Unknown group_by column {group_by}")
    keys = MONTHS if group_by == "Month" else sorted(df[group_by].unique(), key=str)
    groups = [{group_by: k if isinstance(k, str) else int(k), "n": int((df[group_by] == k).sum()),
               "result": calc(df[df[group_by] == k])} for k in keys if (df[group_by] == k).any()]
    return {"rows_matched": int(len(df)), "groups": groups}


def to_xlsx(df: pd.DataFrame, sheet: str = "Data") -> bytes:
    """Excel file with a bold, frozen header row and sensible column widths."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name=sheet)
        ws = xw.sheets[sheet]
        ws.freeze_panes = "A2"
        for i, col in enumerate(df.columns, 1):
            ws.cell(1, i).font = openpyxl.styles.Font(bold=True)
            width = max(len(str(col)), *(len(str(v)) for v in df[col].head(200))) + 2
            ws.column_dimensions[get_column_letter(i)].width = min(width, 28)
    return buf.getvalue()


def attach_dataset(filename: str, filters: list | None = None, columns: list | None = None) -> dict:
    """Dana attaches the dataset (or an extract of it) to her message as an .xlsx file."""
    df = apply_filters(DF, filters)
    if columns:
        missing = [c for c in columns if c not in DF.columns]
        if missing:
            raise ValueError(f"Unknown columns: {missing}")
        df = df[columns]
    if df.empty:
        raise ValueError("No rows match those filters, so there's nothing to attach.")
    name = re.sub(r"[^A-Za-z0-9_\-]+", "_", filename.removesuffix(".xlsx")).strip("_") or "ED_Visits_2025"
    st.session_state.outbox.append({"name": f"{name}.xlsx", "data": to_xlsx(df), "mime": XLSX_MIME})
    return {"attached": f"{name}.xlsx", "rows": int(len(df)), "columns": list(df.columns)}


def read_workbook(data: bytes, filename: str) -> str:
    """Turn a student's workbook into text Dana can review: values, and formulas with their results."""
    try:
        wb_f = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
        wb_v = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        return f"[The analyst attached {filename}, but it couldn't be opened. Ask them to re-save it as .xlsx.]"
    out = [f"[The analyst attached the workbook {filename}. Sheets: {', '.join(wb_f.sheetnames)}.]"]
    for ws_f in wb_f.worksheets:
        ws_v = wb_v[ws_f.title]
        rows, cols = ws_f.max_row, ws_f.max_column
        out.append(f"\n--- Sheet '{ws_f.title}' ({rows} rows x {cols} columns) ---")
        if getattr(ws_f, "_pivots", None):
            out.append("(This sheet contains a PivotTable; its values are shown as cells below.)")
        if getattr(ws_f, "_charts", None):
            out.append(f"(This sheet has {len(ws_f._charts)} chart(s); you can't see charts in the file, "
                       "so ask for a screenshot if you need to review one.)")
        big = rows > 60  # looks like a raw data sheet: show its header and any formulas only
        if big:
            header = [str(c.value) for c in ws_f[1] if c.value is not None]
            out.append(f"Raw data sheet. Header row: {', '.join(header)}. Showing only formula cells:")
        for row in ws_f.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                is_formula = isinstance(c.value, str) and c.value.startswith("=")
                if big and not is_formula:
                    continue
                if is_formula:
                    result = ws_v[c.coordinate].value
                    out.append(f"{c.coordinate}: {c.value}  -> {result if result is not None else '(not calculated)'}")
                else:
                    out.append(f"{c.coordinate}: {c.value}")
        text = "\n".join(out)
        if len(text) > MAX_WORKBOOK_CHARS:
            return text[:MAX_WORKBOOK_CHARS] + "\n[Workbook truncated: too large to show in full.]"
    return "\n".join(out)


TOOLS = [
    {
        "name": "query_data",
        "description": "Compute a statistic on ED_Visits_2025, the same way Excel would (stdev = STDEV.S, "
                       "median = MEDIAN). Use it to check every number the analyst reports before commenting.",
        "input_schema": {
            "type": "object",
            "properties": {
                "metric": {"type": "string", "enum": ["count", "sum", "mean", "median", "min", "max", "stdev"]},
                "column": {"type": "string", "description": "Numeric column (not needed for count): " + ", ".join(NUMERIC)},
                "filters": {"type": "array", "items": {"type": "object", "properties": {
                    "column": {"type": "string"}, "op": {"type": "string", "enum": ["=", "!=", ">", "<", ">=", "<="]},
                    "value": {"type": ["string", "number"]}}, "required": ["column", "op", "value"]}},
                "group_by": {"type": "string", "description": "Optional: Site, Month, TriageLevel, Disposition, AgeGroup, Sex, ChiefComplaint"},
            },
            "required": ["metric"],
        },
    },
    {
        "name": "attach_dataset",
        "description": "Attach an Excel file to your message: the full ED visits dataset, or an extract filtered "
                       "to the rows (and optionally columns) a request needs. The analyst downloads it from your message.",
        "input_schema": {
            "type": "object",
            "properties": {
                "filename": {"type": "string", "description": "Realistic work file name, e.g. ED_Visits_2025 or ED_Barrie_Q1_2025"},
                "filters": {"type": "array", "items": {"type": "object", "properties": {
                    "column": {"type": "string"}, "op": {"type": "string", "enum": ["=", "!=", ">", "<", ">=", "<="]},
                    "value": {"type": ["string", "number"]}}, "required": ["column", "op", "value"]}},
                "columns": {"type": "array", "items": {"type": "string"}, "description": "Optional subset of columns, in order"},
            },
            "required": ["filename"],
        },
    },
]
TOOL_FUNCS = {"query_data": query_data, "attach_dataset": attach_dataset}

# ---------------------------------------------------------------- persona
TOPICS = {
    "Week A · Descriptive statistics": (
        "Descriptive statistics in Excel: AVERAGE, MEDIAN, MIN, MAX, COUNT, COUNTIF, COUNTIFS, AVERAGEIF, "
        "STDEV.S. Requests should need 2-4 numbers, sometimes split by one category (a site or triage level)."),
    "Week B · PivotTables": (
        "PivotTables in Excel: counts and averages broken down by one or two categories (Site, Month, "
        "TriageLevel, Disposition, AgeGroup). Ask for a small summary table and one observation from it."),
    "Week C · Charts and visualizations": (
        "Charts in Excel: column/bar to compare categories, line for trends by Month, pie only for 2-4 parts "
        "of a whole. Ask for one chart for a specific audience. Judge chart choice, title, axis labels, "
        "sorting, and whether an executive gets the point in 5 seconds. Charts can't be seen inside an "
        "uploaded workbook, so ask for a screenshot of the chart."),
    "Week D · Presenting results": (
        "Presenting results: the analyst writes a 3-5 sentence summary or 3 talking points for the VP of "
        "Patient Care. Coach on leading with the answer, plain language, one key number per point, the "
        "'so what', and a recommendation. After their summary, play the VP for one tough question, then "
        "step back out as Dana to give feedback."),
}

SCHEMA = (
    "Dataset ED_Visits_2025, 300 rows, one per emergency department visit (fictional data): "
    "VisitID, VisitDate (2025-01-01 to 2025-12-31), Month (Jan..Dec), Site (Barrie, Orillia, Midland, "
    "Collingwood), Age, AgeGroup (0-17, 18-39, 40-64, 65+), Sex (F, M), TriageLevel (CTAS 1-5; 1 = most "
    "urgent), ChiefComplaint (Abdominal pain, Chest pain, Shortness of breath, Injury/Fracture, Fever, "
    "Headache, Laceration, Mental health), WaitTimeMin (minutes from triage to physician), "
    "LengthOfStayHrs, Disposition (Discharged, Admitted, LWBS = left without being seen, Transferred), CostCAD."
)


def system_prompt(topic: str, student: str) -> str:
    return f"""You are Dana Morgan, Manager of Decision Support at Lakeshore Regional Health, a fictional hospital network in Central Ontario. You are chatting with {student}, a junior healthcare data analyst on your team. In reality they are a college student in a health informatics course, a beginner with Microsoft 365 Excel. Stay in character the whole time.

{SCHEMA}

This week's focus: {TOPICS[topic]}

How you work:
- Write like a busy, supportive manager in a work chat: short, plain, friendly. Under 150 words per reply. No headings. Short bullet lists are fine.
- Each request gives: why you need it (business context), the specific question(s), who it's for, and when (e.g. "by end of day"). Keep it to this week's focus and a beginner's level.
- Send the data with attach_dataset. Attach the full dataset (filename ED_Visits_2025) with your first request. For later requests you may attach a focused extract instead (e.g. one site, one quarter, or fewer columns) when that fits the request. Mention the attachment naturally ("I've attached...").
- Now and then leave one detail a little vague so the analyst has to ask a clarifying question. Answer clarifying questions the way a real manager would.
- Before you comment on ANY number the analyst reports, check it with the query_data tool. Never guess. If they attached their workbook, review their formulas and results the same way and refer to specific cells (e.g. "your formula in B4"). Remember an extract you sent has fewer rows than the full dataset, so check against the same filters.
- If a number is right, say so specifically. If it is off, say which figure doesn't match what you're seeing and give a hint (which function, range, filter or step to recheck). Don't reveal the correct value until they've missed it twice; then show it and explain briefly.
- After each submission ask one realistic follow-up a manager would ask: what explains it, why it matters, what we should do, or who else should see it.
- Don't do the analysis for them. Show an example formula only if they're stuck after a hint.
- If they ask you to just give them the answers, or go off topic, redirect them kindly, in character.
- When a request is finished, thank them and send the next one, a little harder.
- Never mention that you are an AI, a simulation or a tool."""


# ---------------------------------------------------------------- Claude call
@st.cache_resource
def client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])


def ask_dana() -> tuple[str, list]:
    """Run the tool loop until Dana produces a text reply. Returns (text, attached files)."""
    model = st.secrets.get("MODEL", "claude-sonnet-5-5")
    system = [{"type": "text", "text": system_prompt(st.session_state.topic, st.session_state.student),
               "cache_control": {"type": "ephemeral"}}]
    msgs = st.session_state.api_messages
    st.session_state.outbox = []
    for _ in range(8):  # safety limit on tool rounds
        resp = client().messages.create(model=model, max_tokens=1024, system=system, tools=TOOLS, messages=msgs)
        u = resp.usage
        st.session_state.tokens["in"] += u.input_tokens + (u.cache_read_input_tokens or 0) + (u.cache_creation_input_tokens or 0)
        st.session_state.tokens["out"] += u.output_tokens
        msgs.append({"role": "assistant", "content": [b.model_dump(exclude_none=True) for b in resp.content]})
        if resp.stop_reason != "tool_use":
            text = "".join(b.text for b in resp.content if b.type == "text").strip()
            return text, st.session_state.outbox
        results = []
        for b in resp.content:
            if b.type == "tool_use":
                try:
                    out, err = json.dumps(TOOL_FUNCS[b.name](**b.input)), False
                except Exception as e:  # tell Dana what went wrong so she can retry
                    out, err = f"Error: {e}", True
                results.append({"type": "tool_result", "tool_use_id": b.id, "content": out, "is_error": err})
        msgs.append({"role": "user", "content": results})
    fallback = "Sorry, got pulled into a meeting. Can you resend that?"
    msgs.append({"role": "assistant", "content": fallback})
    return fallback, st.session_state.outbox


def reset(topic=None):
    st.session_state.api_messages = []
    st.session_state.chat = []  # dicts: who, text, img, files, upload
    st.session_state.tokens = {"in": 0, "out": 0}
    st.session_state.outbox = []
    if topic:
        st.session_state.topic = topic


# ---------------------------------------------------------------- sign-in gate
if "student" not in st.session_state:
    st.title("Analyst Inbox")
    st.write("Dana Morgan, Manager of Decision Support at Lakeshore Regional Health, will send you "
             "analysis requests. Do the work in Excel and reply with your results.")
    with st.form("signin"):
        name = st.text_input("Your first name")
        code = st.text_input("Class code", type="password")
        if st.form_submit_button("Open my inbox", type="primary"):
            if code.strip() != str(st.secrets.get("CLASS_CODE", "")).strip() or not st.secrets.get("CLASS_CODE"):
                st.error("That class code isn't right. Check the code your instructor posted on Blackboard.")
            elif not name.strip():
                st.error("Enter your first name so Dana knows who she's talking to.")
            else:
                st.session_state.student = name.strip()
                reset(next(iter(TOPICS)))
                st.rerun()
    st.stop()

# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.subheader("This week's focus")
    topic = st.selectbox("Focus", list(TOPICS), index=list(TOPICS).index(st.session_state.topic),
                         label_visibility="collapsed")
    if topic != st.session_state.topic:
        reset(topic)
        st.rerun()

    st.subheader("Conversation")
    transcript = "\n\n".join(
        f"{'Dana Morgan' if m['who'] == 'Dana' else st.session_state.student}: {m['text']}"
        + "".join(f"\n[Attachment: {f['name']}]" for f in m["files"])
        + (f"\n[Attachment: {m['upload']}]" if m["upload"] else "")
        for m in st.session_state.chat)
    st.download_button("Download transcript", transcript or "(No messages yet)",
                       f"transcript_{st.session_state.student}_{datetime.now():%Y%m%d}.txt",
                       use_container_width=True)
    if st.button("Start over", use_container_width=True):
        reset()
        st.rerun()
    st.caption("Submit the transcript on Blackboard when your instructor asks for it.")

    with st.expander("Lost Dana's file?"):
        st.download_button("Download the full dataset", to_xlsx(DF), "ED_Visits_2025.xlsx", XLSX_MIME,
                           use_container_width=True)

# ---------------------------------------------------------------- chat
st.markdown("#### Dana Morgan")
st.caption("Manager, Decision Support · Lakeshore Regional Health")

for i, m in enumerate(st.session_state.chat):
    with st.chat_message("assistant" if m["who"] == "Dana" else "user", avatar="🩺" if m["who"] == "Dana" else None):
        st.markdown(m["text"])
        if m["img"]:
            st.image(m["img"])
        if m["upload"]:
            st.caption(f"📎 {m['upload']}")
        for j, f in enumerate(m["files"]):
            st.download_button(f"📎 {f['name']}", f["data"], f["name"], f["mime"], key=f"dl_{i}_{j}")


def post_dana(text: str, files: list):
    st.session_state.chat.append({"who": "Dana", "text": text, "img": None, "files": files, "upload": None})


if not st.session_state.chat:
    st.info(f"Welcome to the team, {st.session_state.student}. Dana's first request is on its way.")
    st.session_state.api_messages = [{"role": "user", "content": "Hi Dana, I'm ready for my first request."}]
    try:
        with st.spinner("Dana is typing…"):
            reply, files = ask_dana()
        if not files:  # make sure the analyst always gets the data with the first request
            files = [{"name": "ED_Visits_2025.xlsx", "data": to_xlsx(DF), "mime": XLSX_MIME}]
        post_dana(reply, files)
        st.rerun()
    except anthropic.APIError as e:
        st.session_state.api_messages = []
        st.error("Dana can't be reached right now. Wait a minute and reload the page. "
                 f"(Details for your instructor: {e.__class__.__name__})")
        st.stop()

sent = sum(1 for m in st.session_state.chat if m["who"] != "Dana")
if sent >= MAX_STUDENT_MESSAGES:
    st.warning("You've reached the message limit for this session. Download your transcript, then press Start over.")
    st.stop()

prompt = st.chat_input("Reply to Dana. Attach your Excel workbook or a chart screenshot.",
                       accept_file=True, file_type=["xlsx", "png", "jpg", "jpeg"])
if prompt:
    text = (prompt.text or "").strip()
    entry = {"who": "You", "text": text, "img": None, "files": [], "upload": None}
    content = []
    for f in prompt.files[:1]:
        data, lower = f.getvalue(), f.name.lower()
        if lower.endswith(".xlsx"):
            entry["upload"] = f.name
            content.append({"type": "text", "text": read_workbook(data, f.name)})
        else:
            entry["img"] = data
            media = "image/png" if lower.endswith(".png") else "image/jpeg"
            content.append({"type": "image", "source": {"type": "base64", "media_type": media,
                                                        "data": base64.b64encode(data).decode()}})
    if not text:
        text = entry["text"] = "(See attached.)"
    content.append({"type": "text", "text": text})
    st.session_state.chat.append(entry)
    before = len(st.session_state.api_messages)
    st.session_state.api_messages.append({"role": "user", "content": content})
    try:
        with st.spinner("Dana is reviewing…" if entry["upload"] or entry["img"] else "Dana is typing…"):
            reply, files = ask_dana()
        post_dana(reply, files)
    except anthropic.APIError as e:
        st.session_state.chat.pop()
        del st.session_state.api_messages[before:]  # drop the message and any half-finished tool rounds
        st.error("Your message didn't go through. Wait a moment and send it again. "
                 f"(Details for your instructor: {e.__class__.__name__})")
        st.stop()
    st.rerun()
