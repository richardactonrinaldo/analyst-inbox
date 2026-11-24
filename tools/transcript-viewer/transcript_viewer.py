"""Analyst Inbox transcript viewer.

Turns transcripts downloaded from the Analyst Inbox apps (Excel and SQL; .html or older .txt) into a readable web page,
with tutor notes at the top: requests, predictions, possible stumbling points, questions Dana had to
repeat, long pauses, the student's reflection and Dana's end-of-day check-in.

How to use (Windows):
  - Double-click this file, pick one or more transcript .txt files, and the pages open in your browser.
  - Or drag transcript files onto this file.
  - Or from a command prompt:  python transcript_viewer.py transcript_Ethan_1997_20261121.txt

The page is saved next to each transcript as the same name with "_readable.html". When you pick
several transcripts, an index page listing them all opens too.

Needs only Python 3.9 or later (python.org). Nothing else to install.
"""
from __future__ import annotations

import html
import json
import re
import sys
import webbrowser
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

LONG_PAUSE_MIN = 15  # flag gaps between messages longer than this

# Phrases that suggest the student was stuck or Dana had to correct something.
STUDENT_STUCK = [
    r"\bwhat(?:'s| is) wrong\b", r"\bhow (?:do|can|would|should) i\b", r"\bnot (?:really |quite |totally )?sure\b", r"\bi don'?t know\b",
    r"\bi'?m stuck\b", r"\bconfus", r"\bhelp\b", r"\bdoesn'?t work\b", r"\bcan'?t get\b", r"\berror\b",
    r"\bwhy (?:does|is|did|do)\b", r"\bstruggl", r"\btripped me\b",
]
DANA_CORRECTS = [
    r"\bnot quite\b", r"\bclose\b[!,.]", r"\bdoesn'?t match\b", r"\bdon'?t match\b", r"\bisn'?t right\b",
    r"\bisn'?t correct\b", r"\bnot correct\b", r"\bwon'?t work\b", r"\bwrong\b", r"\bmistake\b", r"\boff by\b",
    r"\bhard-coded\b", r"#DIV/0", r"#VALUE", r"#N/A", r"#REF", r"\bI don'?t see\b", r"\bonly has\b",
    r"\bdouble-?check\b", r"\bsyntax\b", r"\bshould (?:read|be)\b", r"\bmissing\b", r"\bslipped\b",
]
DANA_REPEATS = [
    r"\bstill waiting\b", r"\bhaven'?t answered\b", r"\byou haven'?t\b", r"\basking (?:it )?again\b",
    r"\bstill need\b", r"\bcoming back to\b", r"\bstill open\b", r"\bdidn'?t answer\b", r"\bbefore you go\b.*\?",
    r"\bwhen you get a sec\b", r"\bwhat about the (?:first|second)\b",
]


# ---------------------------------------------------------------- parsing
@dataclass
class Message:
    idx: int
    who: str            # "dana" or "student"
    text: str
    time: datetime | None = None
    attachments: list[str] = field(default_factory=list)
    flags: list[str] = field(default_factory=list)
    kind: str = ""      # request, prediction, checkin, reflection, finish, resume, submission


@dataclass
class Transcript:
    path: Path
    student: str = "Student"
    number: str = ""
    module: str = ""
    app: str = "Excel"
    started: datetime | None = None
    messages: list[Message] = field(default_factory=list)


def parse_time(stamp: str, base: datetime | None, prev: datetime | None) -> datetime | None:
    stamp = stamp.strip()
    for fmt in ("%H:%M", "%a %b %d %H:%M"):
        try:
            t = datetime.strptime(stamp, fmt)
        except ValueError:
            continue
        ref = prev or base or datetime(1900, 1, 1)
        if fmt == "%H:%M":
            t = ref.replace(hour=t.hour, minute=t.minute, second=0, microsecond=0)
            if prev and t < prev - timedelta(minutes=1):
                t += timedelta(days=1)
        else:
            t = t.replace(year=ref.year)
        return t
    return None


EMBED = re.compile(r'<script[^>]*id="analyst-inbox-transcript"[^>]*>(.*?)</script>', re.S)


def read_transcript(path: Path) -> str:
    """Plain transcript text from a .txt file, or from the copy embedded in a student's .html transcript."""
    raw = path.read_text(encoding="utf-8", errors="replace")
    if path.suffix.lower() in (".html", ".htm"):
        m = EMBED.search(raw)
        if not m:
            raise ValueError("this HTML file isn't an Analyst Inbox transcript")
        raw = json.loads(m.group(1))
    return raw


def parse(path: Path) -> Transcript:
    raw = read_transcript(path).replace("\r\n", "\n")
    tr = Transcript(path=path, app="SQL" if "_SQL_" in path.name or "transcript_SQL" in path.name else "Excel")
    lines = raw.split("\n")

    # Header: lines before the first blank line
    i = 0
    while i < len(lines) and lines[i].strip():
        line = lines[i]
        m = re.match(r"Student:\s*(.+?)\s*(?:\(student number ending\s*(\d+)\))?\s*$", line)
        if m:
            tr.student, tr.number = m.group(1).strip(), m.group(2) or ""
        elif line.startswith("Module:"):
            tr.module = line.split(":", 1)[1].strip()
        elif line.startswith("Session started:"):
            m2 = re.search(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2})", line)
            if m2:
                tr.started = datetime.strptime(m2.group(1), "%Y-%m-%d %H:%M")
        i += 1
    if tr.started is None:
        m3 = re.search(r"_(\d{8})", path.stem)
        if m3:
            tr.started = datetime.strptime(m3.group(1), "%Y%m%d")

    start = re.compile(r"^(?:\[(?P<ts>[^\]]+)\] )?(?P<who>Dana Morgan|" + re.escape(tr.student) + r"): ?(?P<rest>.*)$")
    current: Message | None = None
    prev_time = None
    for line in lines[i:]:
        m = start.match(line)
        if m:
            if current:
                tr.messages.append(current)
            t = parse_time(m.group("ts"), tr.started, prev_time) if m.group("ts") else None
            prev_time = t or prev_time
            current = Message(idx=len(tr.messages), who="dana" if m.group("who") == "Dana Morgan" else "student",
                              text=m.group("rest"), time=t)
        elif current is not None:
            current.text += "\n" + line
    if current:
        tr.messages.append(current)

    for msg in tr.messages:
        att = re.findall(r"^\[Attachment: (.+?)\]\s*$", msg.text, re.M)
        msg.attachments = att
        msg.text = re.sub(r"^\[Attachment: .+?\]\s*$", "", msg.text, flags=re.M).strip()
    classify(tr)
    return tr


def has(patterns: list[str], text: str) -> bool:
    return any(re.search(p, text, re.I) for p in patterns)


def classify(tr: Transcript):
    wrap = False
    for msg in tr.messages:
        t = msg.text
        if msg.who == "dana":
            if t.lstrip("*").lower().startswith("end-of-day check-in"):
                msg.kind = "checkin"
            elif ((re.search(r"\bbefore you (?:start|begin|run anything|run it|run your)\b", t, re.I)
                   or re.search(r"\b(?:next one|here'?s your first one|first one is)\b", t, re.I)) and len(t) > 200):
                msg.kind = "request"
            if msg.kind != "checkin" and not re.search(r"\bbefore you go\b.*\breflection\b", t, re.I | re.S):
                if has(DANA_CORRECTS, t):
                    msg.flags.append("Dana corrected or questioned something")
                if has(DANA_REPEATS, t):
                    msg.flags.append("Dana had to repeat a question")
        else:
            if t.startswith("My prediction:"):
                msg.kind = "prediction"
            elif t.startswith("(Continuing from my earlier session"):
                msg.kind = "resume"
            elif re.match(r"I need to stop for today", t):
                msg.kind = "finish"
                wrap = True
            elif wrap:
                msg.kind = "reflection"
                wrap = False
            elif t.startswith("Query:") or "\nQuery:" in t:
                msg.kind = "submission"
            if msg.kind not in ("finish", "reflection", "prediction") and has(STUDENT_STUCK, t):
                msg.flags.append("Student asked for help or sounded stuck")
            if "Result: error" in t:
                msg.flags.append("Query returned an error")


# ---------------------------------------------------------------- rendering helpers
def inline(text: str) -> str:
    """Escape, then apply **bold**, *italic* and `code`."""
    parts = re.split(r"(`[^`]+`)", text)
    out = []
    for p in parts:
        if p.startswith("`") and p.endswith("`") and len(p) > 1:
            out.append(f"<code>{html.escape(p[1:-1])}</code>")
        else:
            s = html.escape(p)
            s = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", s)
            s = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"<em>\1</em>", s)
            out.append(s)
    return "".join(out)


def table_html(rows: list[str]) -> str:
    cells = [[c.strip() for c in r.split(" | ")] for r in rows]
    head, body = cells[0], cells[1:]
    h = "".join(f"<th>{html.escape(c)}</th>" for c in head)
    b = "".join("<tr>" + "".join(f"<td>{html.escape(c)}</td>" for c in r) + "</tr>" for r in body)
    return f'<div class="tablewrap"><table><thead><tr>{h}</tr></thead><tbody>{b}</tbody></table></div>'


def render_body(text: str) -> str:
    """Small Markdown subset: paragraphs, bullet and numbered lists, code fences, and SQL submissions."""
    lines = text.split("\n")
    out, i = [], 0
    while i < len(lines):
        line = lines[i]
        if line.strip().startswith("```"):
            lang = line.strip()[3:].strip()
            j = i + 1
            code = []
            while j < len(lines) and not lines[j].strip().startswith("```"):
                code.append(lines[j])
                j += 1
            out.append(f'<pre class="code {html.escape(lang)}">{html.escape(chr(10).join(code))}</pre>')
            i = j + 1
            continue
        if line.strip() == "Query:":
            j = i + 1
            code = []
            while j < len(lines) and not lines[j].startswith("Result:"):
                code.append(lines[j])
                j += 1
            out.append('<div class="label">Query</div><pre class="code sql">' + html.escape("\n".join(code).strip()) + "</pre>")
            i = j
            continue
        if line.startswith("Result:"):
            cls = "result error" if "error" in line[:16] else "result"
            out.append(f'<div class="{cls}">{html.escape(line)}</div>')
            j = i + 1
            rows = []
            while j < len(lines) and " | " in lines[j]:
                rows.append(lines[j])
                j += 1
            if not rows and j < len(lines) and lines[j].strip() and "|" not in lines[j] and "error" not in line:
                rows = []  # single-column result without separators
                while j < len(lines) and lines[j].strip():
                    rows.append(lines[j])
                    j += 1
            if rows:
                out.append(table_html(rows))
            i = j
            continue
        if re.match(r"^\s*[-*•] ", line):
            items = []
            while i < len(lines) and re.match(r"^\s*[-*•] ", lines[i]):
                items.append(re.sub(r"^\s*[-*•] ", "", lines[i]))
                i += 1
            out.append("<ul>" + "".join(f"<li>{inline(x)}</li>" for x in items) + "</ul>")
            continue
        if re.match(r"^\s*\d+[.)] ", line):
            items = []
            while i < len(lines) and re.match(r"^\s*\d+[.)] ", lines[i]):
                items.append(re.sub(r"^\s*\d+[.)] ", "", lines[i]))
                i += 1
            out.append("<ol>" + "".join(f"<li>{inline(x)}</li>" for x in items) + "</ol>")
            continue
        if not line.strip():
            i += 1
            continue
        para = [line]
        i += 1
        while (i < len(lines) and lines[i].strip() and not re.match(r"^\s*([-*•]|\d+[.)]) ", lines[i])
               and not lines[i].strip().startswith("```") and lines[i].strip() != "Query:"
               and not lines[i].startswith("Result:")):
            para.append(lines[i])
            i += 1
        out.append("<p>" + "<br>".join(inline(p) for p in para) + "</p>")
    return "\n".join(out)


def first_sentence(text: str, limit: int = 140) -> str:
    t = re.sub(r"\s+", " ", re.sub(r"[*`#]", "", text)).strip()
    m = re.match(r"(.+?[.!?])(\s|$)", t)
    s = m.group(1) if m else t
    return s if len(s) <= limit else s[:limit - 1].rstrip() + "…"


GREETING = re.compile(r"^(?:hi|hello|hey|thanks|thank you|welcome|nice|great|good|perfect|yes|no problem)\b", re.I)


def request_title(text: str) -> str:
    """The sentence that introduces the request, skipping feedback and greetings that come before it."""
    t = re.sub(r"\s+", " ", re.sub(r"[*`#]", "", text)).strip()
    labelled = re.search(r"\bRequest:\s*(.+)", t)
    if labelled:
        return first_sentence(labelled.group(1))
    sentences = [x for x in re.split(r"(?<=[.!?])\s+", t) if x]
    for k, sent in enumerate(sentences):
        m = re.search(r"\b(?:next one|next request|first one|here'?s your first)\b", sent, re.I)
        if m:
            if len(sent[m.end():].strip(" .:!,")) < 25 and k + 1 < len(sentences):  # "Next one." / "...your next request."
                return first_sentence(sentences[k + 1])
            return first_sentence(sent)
    body = re.sub(r"^request:\s*", "", t, flags=re.I)
    for sent in re.split(r"(?<=[.!?])\s+", body):
        sent = re.sub(r"^request:\s*", "", sent, flags=re.I)
        if sent and not GREETING.match(sent) and len(sent) > 25:
            return first_sentence(sent)
    return first_sentence(t)


def snippet(text: str, limit: int = 170) -> str:
    t = re.sub(r"\s+", " ", re.sub(r"\*\*|`|^#+ ", "", text)).strip()
    return t if len(t) <= limit else t[:limit - 1].rstrip() + "…"


def matched_sentence(text: str, patterns: list[str], limit: int = 190) -> str:
    """The sentence where the first pattern matches, so the note shows the actual correction or question."""
    t = re.sub(r"\*\*|`", "", text)
    t = re.sub(r"(^|\n)\s*(?:[-*•]|\d+[.)])\s+", r"\1", t)          # drop list markers
    t = re.sub(r"\s+", " ", t).strip()
    hits = [m.start() for p in patterns for m in [re.search(p, t, re.I)] if m]
    if not hits:
        return snippet(text, limit)
    pos = min(hits)
    start = max(t.rfind(". ", 0, pos), t.rfind("? ", 0, pos), t.rfind("! ", 0, pos), t.rfind(": ", 0, pos))
    start = start + 2 if start >= 0 else 0
    end_m = re.search(r"[.?!](\s|$)", t[pos:])
    end = pos + end_m.end() if end_m else len(t)
    if end - start < 45:                     # "Not quite, Ethan." says little; add the next sentence
        more = re.search(r"[.?!](\s|$)", t[end:])
        end = end + more.end() if more else len(t)
    s = re.sub(r"^\d+[.)] ", "", t[start:end].strip())
    return s if len(s) <= limit else s[:limit - 1].rstrip() + "…"


def tstr(t: datetime | None) -> str:
    return t.strftime("%H:%M") if t else ""


def duration(a: datetime | None, b: datetime | None) -> str:
    if not (a and b):
        return "not recorded"
    mins = int((b - a).total_seconds() // 60)
    return f"{mins // 60} h {mins % 60:02d} min" if mins >= 60 else f"{mins} min"


# ---------------------------------------------------------------- page
CSS = """
:root { --ink:#1f2a30; --muted:#5c6b73; --line:#dde3e6; --dana:#f1f5f7; --you:#e7f1ea; --accent:#0b4f5c;
        --warn:#b3541e; --warnbg:#fdf1e8; --pred:#fff7df; --ok:#2e7d4f; }
* { box-sizing: border-box; }
body { font: 15px/1.5 "Segoe UI", Calibri, Arial, sans-serif; color: var(--ink); margin: 0; background: #fafbfc; }
.wrap { max-width: 980px; margin: 0 auto; padding: 28px 24px 60px; }
h1 { font-size: 24px; margin: 0 0 4px; color: var(--accent); }
h2 { font-size: 18px; margin: 28px 0 10px; color: var(--accent); }
h3 { font-size: 15px; margin: 18px 0 6px; }
.sub { color: var(--muted); margin-bottom: 18px; }
.stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 10px; margin: 14px 0 6px; }
.stat { background: #fff; border: 1px solid var(--line); border-radius: 8px; padding: 10px 12px; }
.stat b { display: block; font-size: 20px; color: var(--accent); }
.stat span { color: var(--muted); font-size: 13px; }
.notes { background: #fff; border: 1px solid var(--line); border-radius: 10px; padding: 6px 18px 14px; }
.notes ul { margin: 6px 0 4px; padding-left: 20px; }
.notes li { margin: 4px 0; }
.notes a { color: var(--accent); text-decoration: none; }
.notes a:hover { text-decoration: underline; }
.time { color: var(--muted); font-variant-numeric: tabular-nums; }
.tag { display: inline-block; font-size: 12px; padding: 1px 8px; border-radius: 10px; margin-right: 6px;
       background: var(--warnbg); color: var(--warn); }
.tag.req { background: #e3eef1; color: var(--accent); }
.tag.pred { background: var(--pred); color: #7a5b00; }
.toolbar { margin: 10px 0 14px; display: flex; gap: 8px; flex-wrap: wrap; }
.toolbar button { font: inherit; padding: 5px 12px; border-radius: 6px; border: 1px solid var(--line);
                  background: #fff; cursor: pointer; }
.toolbar button.on { background: var(--accent); color: #fff; border-color: var(--accent); }
.msg { display: flex; margin: 10px 0; }
.msg.student { justify-content: flex-end; }
.bubble { max-width: 80%; background: var(--dana); border: 1px solid var(--line); border-radius: 12px; padding: 10px 14px; }
.student .bubble { background: var(--you); }
.msg.request .bubble { border-left: 4px solid var(--accent); }
.msg.prediction .bubble { background: var(--pred); }
.msg.checkin .bubble { max-width: 100%; background: #fff; border: 2px solid var(--accent); }
.msg.flagged .bubble { box-shadow: 0 0 0 2px var(--warn) inset; }
.who { font-size: 13px; font-weight: 600; color: var(--muted); margin-bottom: 4px; }
.bubble p { margin: 6px 0; }
.bubble ul, .bubble ol { margin: 6px 0; padding-left: 22px; }
code { background: #e9eef1; padding: 1px 4px; border-radius: 4px; font: 13px Consolas, monospace; }
pre.code { background: #1f2a30; color: #e8eef1; padding: 10px 12px; border-radius: 8px; overflow-x: auto;
           font: 13px/1.45 Consolas, monospace; white-space: pre-wrap; }
.label { font-size: 12px; font-weight: 600; color: var(--muted); margin-top: 6px; }
.result { font-size: 13px; color: var(--ok); margin: 6px 0 4px; }
.result.error { color: var(--warn); font-weight: 600; }
.tablewrap { overflow-x: auto; }
table { border-collapse: collapse; font-size: 13px; background: #fff; margin: 4px 0 6px; }
th, td { border: 1px solid var(--line); padding: 3px 8px; text-align: left; }
th { background: #eef3f5; }
.chips { margin-top: 6px; }
.chip { display: inline-block; font-size: 12px; background: #fff; border: 1px solid var(--line); border-radius: 12px;
        padding: 1px 9px; margin: 2px 4px 0 0; color: var(--muted); }
.gap { text-align: center; color: var(--warn); font-size: 13px; margin: 14px 0; }
.flagnote { font-size: 12px; color: var(--warn); margin-top: 6px; }
.hidden { display: none !important; }
.empty { color: var(--muted); }
@media print { .toolbar { display: none; } body { background: #fff; } .msg { break-inside: avoid; } }
"""

JS = """
function show(mode, btn) {
  document.querySelectorAll('.toolbar button').forEach(b => b.classList.remove('on'));
  btn.classList.add('on');
  document.querySelectorAll('.msg, .gap').forEach(el => {
    let keep = mode === 'all'
      || (mode === 'flagged' && el.classList.contains('flagged'))
      || (mode === 'requests' && (el.classList.contains('request') || el.classList.contains('prediction')
                                  || el.classList.contains('checkin')));
    el.classList.toggle('hidden', !keep);
  });
}
"""


def pause_label(m: Message) -> str:
    """What came after a pause: the message, or the file sent if it was just an attachment."""
    text = m.text if m.text not in ("(See attached.)", "") else ""
    who = "Dana" if m.who == "dana" else "student"
    if m.attachments and not text:
        return f"{who} sent {', '.join(m.attachments)}"
    return f"{who}: " + snippet(text, 110)


def build_page(tr: Transcript) -> str:
    msgs = tr.messages
    times = [m.time for m in msgs if m.time]
    start, end = (min(times), max(times)) if times else (None, None)
    student_msgs = [m for m in msgs if m.who == "student"]
    requests = [m for m in msgs if m.kind == "request"]
    preds = [m for m in msgs if m.kind == "prediction"]
    files_in = sum(len(m.attachments) for m in student_msgs)
    queries = sum(1 for m in student_msgs if m.kind == "submission")
    stuck = [m for m in msgs if any("Student" in f or "error" in f for f in m.flags)]
    corrected = [m for m in msgs if "Dana corrected or questioned something" in m.flags]
    repeats = [m for m in msgs if "Dana had to repeat a question" in m.flags]
    reflection = next((m for m in msgs if m.kind == "reflection"), None)
    checkin = next((m for m in msgs if m.kind == "checkin"), None)

    gaps = []
    prev = None
    for m in msgs:
        if m.time and prev and prev.time and (m.time - prev.time) > timedelta(minutes=LONG_PAUSE_MIN):
            gaps.append((prev, m))
        if m.time:
            prev = m

    def link(m: Message, label: str | None = None) -> str:
        return (f'<a href="#m{m.idx}"><span class="time">{tstr(m.time) or "#" + str(m.idx + 1)}</span> '
                f'{html.escape(label or snippet(m.text))}</a>')

    def listing(items, fn, empty):
        return "<ul>" + "".join(f"<li>{fn(x)}</li>" for x in items) + "</ul>" if items else f'<p class="empty">{empty}</p>'

    # Where-to-grow and focus lines from the check-in
    grow = ""
    if checkin:
        m = re.search(r"\*\*Where to grow\*\*(.+?)(\*\*Focus for next time\*\*|$)", checkin.text, re.S)
        f = re.search(r"\*\*Focus for next time\*\*\s*(.+?)(\n\n|$)", checkin.text, re.S)
        if m:
            grow = render_body(m.group(1).strip())
        if f:
            grow += f"<p><strong>Focus for next time:</strong> {inline(f.group(1).strip())}</p>"

    notes = f"""
<h3>Requests Dana sent ({len(requests)})</h3>
{listing(requests, lambda m: link(m, request_title(m.text)), "No requests detected.")}
<h3>Student predictions ({len(preds)})</h3>
{listing(preds, lambda m: link(m, m.text.replace("My prediction:", "").strip()), "No predictions recorded.")}
<h3>Possible stumbling points: student asked for help or hit an error ({len(stuck)})</h3>
{listing(stuck, lambda m: link(m, matched_sentence(m.text, STUDENT_STUCK + [r"Result: error"])), "None detected.")}
<h3>Dana corrected or questioned something ({len(corrected)})</h3>
{listing(corrected, lambda m: link(m, matched_sentence(m.text, DANA_CORRECTS)), "None detected.")}
<h3>Questions Dana had to repeat ({len(repeats)})</h3>
{listing(repeats, lambda m: link(m, matched_sentence(m.text, DANA_REPEATS)), "None. The student answered Dana's questions the first time.")}
<h3>Long pauses (over {LONG_PAUSE_MIN} min) ({len(gaps)})</h3>
{listing(gaps, lambda g: link(g[1], f"{duration(g[0].time, g[1].time)} pause, then: " + pause_label(g[1])), "None.")}
<h3>Student's reflection</h3>
{f"<p>{inline(reflection.text)}</p>" if reflection else '<p class="empty">No reflection in this transcript.</p>'}
<h3>Dana's check-in: where to grow</h3>
{grow or '<p class="empty">No end-of-day check-in in this transcript.</p>'}
"""

    convo = []
    prev = None
    for m in msgs:
        if m.time and prev and prev.time and (m.time - prev.time) > timedelta(minutes=LONG_PAUSE_MIN):
            convo.append(f'<div class="gap">⏸ {duration(prev.time, m.time)} pause</div>')
        if m.time:
            prev = m
        classes = ["msg", m.who] + ([m.kind] if m.kind else []) + (["flagged"] if m.flags else [])
        name = "Dana Morgan" if m.who == "dana" else tr.student
        tags = ""
        if m.kind == "request":
            tags += '<span class="tag req">New request</span>'
        if m.kind == "prediction":
            tags += '<span class="tag pred">Prediction</span>'
        tags += "".join(f'<span class="tag">{html.escape(f)}</span>' for f in m.flags)
        chips = "".join(f'<span class="chip">📎 {html.escape(a)}</span>' for a in m.attachments)
        convo.append(
            f'<div class="{" ".join(classes)}" id="m{m.idx}"><div class="bubble">'
            f'<div class="who">{html.escape(name)} <span class="time">{tstr(m.time)}</span></div>'
            f'{("<div>" + tags + "</div>") if tags else ""}'
            f'{render_body(m.text) if m.text else "<p class=empty>(attachment only)</p>"}'
            f'{("<div class=chips>" + chips + "</div>") if chips else ""}'
            f"</div></div>")

    date = (start or tr.started)
    title = f"{tr.student} · {tr.app} Analyst Inbox"
    sub = " · ".join(x for x in [
        f"Student number ending {tr.number}" if tr.number else "",
        tr.module, date.strftime("%A %d %B %Y") if date else ""] if x)
    stats = [
        (duration(start, end), "session length"),
        (f"{tstr(start)}–{tstr(end)}" if start else "–", "first to last message"),
        (str(len(student_msgs)), "student messages"),
        (str(len(requests)), "requests"),
        (str(queries if tr.app == "SQL" else files_in), "queries sent" if tr.app == "SQL" else "files sent by student"),
        (str(len(stuck) + len(corrected)), "flagged moments"),
    ]
    stat_html = "".join(f'<div class="stat"><b>{html.escape(v)}</b><span>{html.escape(k)}</span></div>' for v, k in stats)

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>{html.escape(title)}</title>
<meta name="viewport" content="width=device-width, initial-scale=1"><style>{CSS}</style></head>
<body><div class="wrap">
<h1>{html.escape(title)}</h1>
<div class="sub">{html.escape(sub)}<br><span class="time">Source: {html.escape(tr.path.name)}</span></div>
<div class="stats">{stat_html}</div>
<h2>Tutor notes</h2>
<div class="notes">{notes}
<p class="empty" style="margin-top:12px">Flags are found by keyword, so treat them as pointers to look at, not a verdict.
Click any item to jump to it in the conversation.</p></div>
<h2>Conversation</h2>
<div class="toolbar">
  <button class="on" onclick="show('all', this)">All messages</button>
  <button onclick="show('flagged', this)">Flagged only</button>
  <button onclick="show('requests', this)">Requests, predictions and check-in</button>
  <button onclick="window.print()">Print</button>
</div>
{"".join(convo)}
</div><script>{JS}</script></body></html>"""


def build_index(results: list[tuple[Transcript, Path]]) -> str:
    rows = []
    for tr, out in results:
        msgs = tr.messages
        times = [m.time for m in msgs if m.time]
        start, end = (min(times), max(times)) if times else (None, None)
        flagged = sum(1 for m in msgs if m.flags)
        rows.append(
            f"<tr><td><a href='{html.escape(out.name)}'>{html.escape(tr.student)}</a></td><td>{html.escape(tr.number)}</td>"
            f"<td>{html.escape(tr.app)}</td><td>{html.escape(tr.module)}</td>"
            f"<td>{(start or tr.started).strftime('%Y-%m-%d') if (start or tr.started) else ''}</td>"
            f"<td>{duration(start, end)}</td><td>{sum(1 for m in msgs if m.kind == 'request')}</td><td>{flagged}</td></tr>")
    return f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Transcripts</title>
<style>{CSS} table {{ width: 100%; font-size: 14px; }} td, th {{ padding: 6px 10px; }}</style></head>
<body><div class="wrap"><h1>Analyst Inbox transcripts</h1>
<div class="sub">{len(results)} transcript(s). Click a name to open its page.</div>
<table><thead><tr><th>Student</th><th>No.</th><th>App</th><th>Module</th><th>Date</th><th>Length</th>
<th>Requests</th><th>Flagged messages</th></tr></thead><tbody>{"".join(rows)}</tbody></table></div></body></html>"""


# ---------------------------------------------------------------- main
def pick_files() -> list[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        print("Pass transcript files on the command line, e.g. python transcript_viewer.py transcript.txt")
        return []
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    names = filedialog.askopenfilenames(title="Choose Analyst Inbox transcript(s)",
                                        filetypes=[("Transcripts", "*.html *.htm *.txt"), ("All files", "*.*")])
    root.destroy()
    return [Path(n) for n in names]


def convert(path: Path, out_dir: Path | None = None) -> tuple[Transcript, Path]:
    tr = parse(path)
    out = (out_dir or path.parent) / f"{path.stem}_readable.html"
    out.write_text(build_page(tr), encoding="utf-8")
    return tr, out


def main(argv: list[str]) -> int:
    no_open = "--no-open" in argv
    args = [a for a in argv if a != "--no-open"]
    paths = [Path(a) for a in args] if args else pick_files()
    paths = [p for p in paths if p.exists() and p.suffix.lower() in (".txt", ".html", ".htm")
             and not p.stem.endswith("_readable") and p.name != "transcripts_index.html"]
    if not paths:
        print("No transcript files chosen.")
        return 1
    results = []
    for p in paths:
        try:
            results.append(convert(p))
            print(f"Made {results[-1][1]}")
        except Exception as e:  # keep going with the other files
            print(f"Couldn't read {p.name}: {e}")
    if not results:
        return 1
    if len(results) > 1:
        index = results[0][1].parent / "transcripts_index.html"
        index.write_text(build_index(results), encoding="utf-8")
        print(f"Made {index}")
        if not no_open:
            webbrowser.open(index.resolve().as_uri())
    elif not no_open:
        webbrowser.open(results[0][1].resolve().as_uri())
    return 0


if __name__ == "__main__":
    code = main(sys.argv[1:])
    if code and not sys.argv[1:]:          # double-clicked: keep the window open so any message can be read
        input("Press Enter to close.")
    sys.exit(code)
