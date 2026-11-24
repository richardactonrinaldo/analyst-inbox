"""Build the printable alternative assignments (Excel and SQL) as PDFs via headless Chrome."""
import asyncio
import html
from pathlib import Path

OUT = Path(__file__).parent

CSS = """
@page { size: Letter; margin: 0.7in 0.75in 0.8in; }
* { box-sizing: border-box; }
body { font: 11pt/1.45 Calibri, "Segoe UI", Arial, sans-serif; color: #1d292e; margin: 0; }
h1 { font-size: 22pt; color: #0b4f5c; margin: 0 0 2pt; line-height: 1.1; }
.course { color: #5b6a71; font-size: 11pt; margin-bottom: 12pt; }
h2 { font-size: 15pt; color: #0b4f5c; margin: 18pt 0 4pt; border-bottom: 1.5pt solid #0b4f5c; padding-bottom: 2pt; break-after: avoid; }
h3 { font-size: 12pt; margin: 12pt 0 3pt; break-after: avoid; }
p { margin: 4pt 0; }
ul, ol { margin: 4pt 0 4pt 18pt; padding: 0; }
li { margin: 2pt 0; }
.box { border: 1pt solid #c9d4d8; background: #f3f7f8; border-radius: 4pt; padding: 7pt 10pt; margin: 8pt 0; }
.memo { border-left: 3pt solid #0b4f5c; padding: 2pt 0 2pt 9pt; margin: 6pt 0; font-style: italic; }
.req { margin: 8pt 0 10pt; }
.req ol, .req ul, .memo { break-inside: avoid; }
.label { font-weight: 700; color: #0b4f5c; }
.lines { height: 0; border-bottom: 0.75pt solid #9fb0b6; margin: 13pt 0 0; }
code { font: 10pt Consolas, monospace; background: #eef3f5; padding: 0 2pt; }
table { border-collapse: collapse; margin: 4pt 0; font-size: 10pt; }
th, td { border: 0.75pt solid #c9d4d8; padding: 2pt 6pt; text-align: left; }
th { background: #eef3f5; }
.small { font-size: 9.5pt; color: #5b6a71; }
.pb { break-before: page; }
"""


def lines(n=2):
    return '<div class="lines"></div>' * n


def request(n, title, memo, tasks, predict, check, interpret):
    t = "".join(f"<li>{x}</li>" for x in tasks)
    c = "".join(f"<li>{x}</li>" for x in check)
    i = "".join(f"<li>{x}</li>" for x in interpret)
    return f"""<div class="req">
<h3>Request {n}: {title}</h3>
<div class="memo">{memo}</div>
<p><span class="label">Before you start, predict:</span> {predict}</p>{lines(2)}
<p class="label" style="margin-top:8pt">Tasks</p><ol>{t}</ol>
<p class="label">Check your own work</p><ul>{c}</ul>
<p class="label">Interpret</p><ul>{i}</ul>
</div>"""


def page(title, course, intro, modules, closing):
    body = "".join(f"<h2{' class=\"pb\"' if k else ''}>{html.escape(m)}</h2>{content}" for k, (m, content) in enumerate(modules))
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>{html.escape(title)}</title>
<style>{CSS}</style></head><body>
<h1>{html.escape(title)}</h1><div class="course">{html.escape(course)}</div>
{intro}{body}{closing}</body></html>"""


HOW = """<div class="box"><p><strong>About this assignment.</strong> This is the print alternative to the Analyst
Inbox app, for students who prefer not to use an AI tool. It covers the same skills. You play a junior data
analyst at Lakeshore Regional Health (a fictional hospital network); Dana Morgan, your manager, sends the requests
below. All data is fictional.</p>
<p><strong>How to work each request</strong></p>
<ol>
<li><strong>Predict first.</strong> Write your prediction before you touch the data, with a rough number where you can and a reason. A wrong prediction is fine; it shows where your instincts need tuning.</li>
<li><strong>Do the work</strong> {how_work}</li>
<li><strong>Check your own work</strong> with the questions given. They ask what an experienced manager would ask.</li>
<li><strong>Interpret.</strong> Answer in plain sentences a busy director could read in a few seconds.</li>
</ol>
<p><strong>What to submit</strong> {submit}</p>
<p class="small">Your instructor will tell you which modules are due when. Work through the requests in order; later ones build on earlier ones.</p></div>"""

REFLECT = """<h2 class="pb">End-of-module reflection</h2>
<p>Answer these at the end of each module you complete (a short paragraph each).</p>
<ol>
<li>What tripped you up, and how did you work through it?</li>
<li>Which prediction was furthest from the result? What does that tell you about your instincts?</li>
<li>Explain one formula or query you wrote, argument by argument, as if to a classmate who missed the class.</li>
<li>What would you check next if a director asked a follow-up question about your results?</li>
</ol>""" + lines(10)

# ---------------------------------------------------------------- Excel
EX_MODULES = [
("Module A · Descriptive statistics (wait times and length of stay)",
 request(1, "ED snapshot for the operations meeting",
  "I'm pulling together a one-page ED snapshot for Thursday's operations meeting. The site directors will see it, so I need a few basic numbers, built with formulas.",
  ["The average wait time (<code>WaitTimeMin</code>) across all visits.", "The median wait time across all visits.",
   "How many visits ended in LWBS (left without being seen)."],
  "Will the median wait be above or below the average? Why?",
  ["Is every result a formula, not a typed number?", "Did you use AVERAGE rather than AVERAGEA? Would they ever differ?"],
  ["In one or two plain sentences: what does the gap between the average and the median tell the directors about our waits?"])
 + request(2, "How long is \"extremely long\"?",
  "A director will ask me how long the longest waits are, and whether they're common or rare.",
  ["The longest and shortest wait (MAX, MIN).", "The number of visits with a wait over 180 minutes, then over 120 minutes (COUNTIF).",
   "Each of those counts as a percentage of all visits, formatted as a percentage."],
  "If you change the cut-off from 180 to 120 minutes, roughly how many visits will qualify? Give a number.",
  ["What would happen to your COUNTIF if you typed the criteria without quotes?", "Does your percentage formula refer to the total with an absolute reference?"],
  ["Are very long waits common or rare? Which cut-off would you show on the one-pager, and why?"])
 + request(3, "Are our sickest patients seen faster?",
  "The site directors want to know whether our most urgent patients wait less. Use your judgement on which triage levels count as most and least urgent, and tell me what you chose.",
  ["The average wait for the most urgent group and for the least urgent group (AVERAGEIF or AVERAGEIFS).",
   "The number of visits in each group.", "The longest wait in each group."],
  "Roughly how many minutes will each group wait on average?",
  ["AVERAGEIFS treats multiple criteria as AND. How did you capture more than one triage level?",
   "How many visits sit behind your most-urgent figure? Would you trust it?"],
  ["Write the headline for the one-pager on waits by urgency, worded so nobody can misread \"lower\" triage levels."])
 + request(4, "How consistent are we?",
  "The Thursday meeting also needs a sense of how stable our ED is.",
  ["The standard deviation of wait time across all visits (STDEV.S).",
   "The average <code>LengthOfStayHrs</code> for Admitted visits and for Discharged visits.",
   "The median length of stay for Admitted visits (hint: <code>MEDIAN(FILTER(...))</code>)."],
  "Will the standard deviation be small or large compared with the average wait? Which group stays longer, and by how many hours?",
  ["Walk through your AVERAGEIF: what does each of its three arguments do?",
   "Are the median and mean length of stay close? What does that tell you?"],
  ["In one or two sentences for the directors: how consistent are our waits?"])),
("Module B · PivotTables (who comes to the ED and what happens to them)",
 request(1, "What brings people in, and what happens next?",
  "Our VP wants to understand the mix of patients. Build a PivotTable from the dataset.",
  ["A PivotTable of visit counts by <code>ChiefComplaint</code> (rows) and <code>Disposition</code> (columns), with totals.",
   "Sort the complaints from most to least common."],
  "Which complaint will be most common? Which will have the most admissions?",
  ["Do your row and column totals add up to 300 visits?", "Did you count a field that is never blank (such as VisitID)?"],
  ["Which complaint would you flag to the VP, and why: the most common one, or the one most often admitted?"])
 + request(2, "Who gets admitted?",
  "Bed planning wants to know which age groups are most likely to be admitted.",
  ["A PivotTable of visits by <code>AgeGroup</code> and <code>Disposition</code>.",
   "The admission rate for each age group (admitted ÷ all visits in that group), as a percentage. Show it as % of row total, or with a formula next to the pivot."],
  "Which age group will have the highest admission rate? Roughly what percentage?",
  ["Why is a rate more useful here than a count?", "Which age group has the fewest visits, and would you trust its rate?"],
  ["Write two sentences for bed planning on what the admission rates show."])
 + request(3, "Our busiest months",
  "Staffing wants to know when demand peaks.",
  ["A PivotTable of visit counts by <code>Month</code>, in calendar order (not alphabetical).",
   "Add the average <code>LengthOfStayHrs</code> by month to the same pivot."],
  "Which month will be busiest, and by how much compared with the quietest?",
  ["Are the months in January-to-December order? How did you get them there?"],
  ["Is there a pattern staffing should plan for, or is it too noisy to call with one year of data?"])),
("Module C · Charts and visualizations (demand and patient mix)",
 request(1, "Demand across the year",
  "The VP gives each slide about five seconds. She needs to see how visit volume moves across the year.",
  ["A line chart of visit counts by month, in calendar order.", "A headline title that states the point (for example, \"Visits peak in ...\"), and labelled axes."],
  "What shape will the line have?",
  ["Would a column chart work as well here? Why is a line usually better for months?", "Does the y-axis start at zero, and does that matter here?"],
  ["Paste the chart into your answers document with one sentence saying what the VP should take from it."])
 + request(2, "The mix of complaints",
  "For the same deck: what are people coming in for?",
  ["A bar chart of visit counts by chief complaint, sorted longest to shortest.", "A headline title and axis labels."],
  "Which three complaints will make up the most visits?",
  ["Why is a bar chart better than a pie chart with eight slices?", "Is the sort order helping the reader?"],
  ["What one action might the VP take from this chart?"])
 + request(3, "Who is admitted, and how long they stay",
  "Two last charts for bed planning.",
  ["A column chart of the admission rate by age group (from Module B, or rebuild it).",
   "A column chart of average length of stay by disposition."],
  "Which chart will make its point more strongly?",
  ["Is each chart showing a rate or a count? Is that the right choice?", "Would a reader misread anything without your title?"],
  ["Choose the one chart you would put in front of the VP and say why."])),
("Module D · Presenting results (where ED spending goes)",
 request(1, "Where does the money go?",
  "Finance asked where ED spending goes. Work out the numbers first, with formulas or a pivot.",
  ["Total and average <code>CostCAD</code> by <code>Disposition</code>.",
   "The share of total cost that comes from Admitted visits, and the share of visits that are Admitted.",
   "Average cost by AgeGroup."],
  "What share of total cost do you expect Admitted visits to account for?",
  ["Do your shares add up to 100%?", "Which number would be easiest to misread, and how will you label it?"],
  ["Which single number tells the cost story best?"])
 + request(2, "Briefing the VP",
  "The VP of Patient Care wants three talking points she can say out loud at the board meeting.",
  ["Write three talking points (one sentence each) that lead with the answer, use one key number each, and say \"so what\".",
   "Add one recommendation."],
  "What do you think the VP will ask first?",
  ["Can each point be understood without seeing your spreadsheet?", "Did you avoid jargon such as \"disposition\" or \"LWBS\" without explaining it?"],
  ["Imagine the VP asks: \"Is this just because admitted patients are sicker?\" Write how you would answer, and what data you would check next."])),
]

EX = page("Analyst Inbox: Excel Alternative Assignment", "HIFN1009 · Health Informatics and Data Analytics",
          HOW.format(how_work="in Excel on the dataset <code>ED_Visits_2025.xlsx</code> (300 ED visits from 2025, posted on Blackboard). Every result must be a formula, not a typed number.",
                     submit="your Excel workbook (one sheet per module, clearly labelled) and an answers document (Word or PDF) with your predictions, the check-your-work answers, your interpretations, screenshots of any charts, and the reflection."),
          EX_MODULES, REFLECT)

# ---------------------------------------------------------------- SQL
SQL_MODULES = [
("Module A · SELECT, WHERE and ORDER BY",
 request(1, "Barrie's longest waits",
  "Our Barrie site lead wants to see the visits with the longest waits at her site.",
  ["Look up Barrie's site_id in the <code>sites</code> table, then list the 10 longest waits at Barrie: visit_id, visit_date, triage_level and wait_time_min, longest first."],
  "Roughly how long will the longest Barrie wait be, in minutes?",
  ["Why do you need ORDER BY ... DESC and LIMIT together?", "How did you find Barrie's site_id?"],
  ["Is there anything in common among the 10 visits (triage level, time of year)?"])
 + request(2, "Non-urgent visits that waited long",
  "Quality wants to review non-urgent visits that waited a long time.",
  ["List triage level 5 visits with a wait over 120 minutes, showing visit_id, site_id and wait_time_min.",
   "Then change your query to triage level 4 or 5."],
  "How many triage 5 visits waited over 120 minutes?",
  ["When you combined levels 4 and 5, did you use IN, or AND/OR with parentheses? Why do the parentheses matter?"],
  ["Write one sentence for Quality on what you found."])),
("Module B · Aggregates and GROUP BY",
 request(1, "Volume and waits by site",
  "I need a one-table summary by site for the directors.",
  ["For each site_id: the number of visits, the average wait (rounded to one decimal), and the longest wait.",
   "Sort by average wait, longest first. Give each calculated column a clear alias."],
  "Which site will have the most visits, and which the longest average wait?",
  ["COUNT(*) and COUNT(cost_cad) give different answers on the visits table. Why? Which is right for counting visits?",
   "Why can't you use an alias in WHERE?"],
  ["Which site would you flag, and why?"])
 + request(2, "Where the money goes",
  "Finance asked for cost by disposition.",
  ["For each disposition: number of visits, total cost and average cost.", "Keep only dispositions with more than 20 visits (HAVING)."],
  "Which disposition will have the highest average cost?",
  ["Some cost_cad values are missing. How do SUM and AVG treat them?", "What is the difference between WHERE and HAVING?"],
  ["Write two sentences for Finance."])),
("Module C · JOINs",
 request(1, "Site names instead of numbers",
  "The directors don't know site IDs. Redo your site summary with names.",
  ["Join visits to sites and show site_name, number of visits and average wait.",
   "Then list every site, including any with no visits (LEFT JOIN), with its visit count."],
  "Will every site have visits?",
  ["Which site appears only with the LEFT JOIN, and why?", "Did COUNT(*) or COUNT(v.visit_id) give the right count for that site?"],
  ["What would you tell the directors about the site with no visits?"])
 + request(2, "Waits against targets",
  "Clinical leads want to compare waits with the CTAS target for each triage level.",
  ["Join visits to triage_levels and show, per triage level: the description, the target wait, the average actual wait and the number of visits.",
   "Add a column with the difference between the average wait and the target."],
  "Which triage level will be furthest over its target?",
  ["Which table holds the target, and what column links the two tables?"],
  ["Which result would you raise with clinical leads first?"])
 + request(3, "Who are our patients?",
  "Registration wants to know how many patients came back more than once, and how many are in the system but never visited.",
  ["Join patients to visits and count visits per patient; list patients with more than one visit.",
   "Count patients with no visits at all."],
  "Roughly how many patients will have visited more than once?",
  ["Why does the second task need a LEFT JOIN and a check for NULL?"],
  ["Suggest one reason a patient might be registered with no visits."])),
("Module D · CASE and presenting results",
 request(1, "Wait bands for the board",
  "The board wants waits grouped into plain-language bands.",
  ["Use CASE to label each visit's wait: \"Under 30 min\", \"30 to 119 min\", \"2 hours or more\".",
   "Count visits in each band by site name."],
  "What share of visits will wait 2 hours or more?",
  ["Do your bands overlap or leave gaps? Check the boundaries at 30 and 120.", "Why must the CASE conditions be in the right order?"],
  ["Which site has the biggest share in the longest band?"])
 + request(2, "Briefing the VP",
  "The VP of Patient Care wants three talking points from your work in this assignment.",
  ["Write three talking points (one sentence each) that lead with the answer, use one key number each, and say \"so what\".",
   "Add one recommendation, and the query you would run next to support it."],
  "What do you think the VP will ask first?",
  ["Can each point be understood without seeing your query results?"],
  ["Imagine the VP asks: \"Is Barrie just busier, or slower?\" Write how you would answer, and which query would settle it."])),
]

SQL = page("Analyst Inbox: SQL Alternative Assignment", "HIFN1009 · Health Informatics and Data Analytics",
           HOW.format(how_work="in SQL on the SQLite database <code>lakeshore_ed.db</code> (posted on Blackboard; open it with DB Browser for SQLite or a similar tool). Tables: <code>sites</code>, <code>triage_levels</code>, <code>patients</code>, <code>visits</code>. Use SELECT, WHERE, ORDER BY, aggregates, GROUP BY, HAVING, JOINs and CASE (no subqueries).",
                      submit="an answers document (Word or PDF) with, for each request, your prediction, every query you ran with its results (paste the result grid or the first rows), the check-your-work answers, your interpretations, and the reflection."),
           SQL_MODULES, REFLECT)


async def render():
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        b = await p.chromium.launch(executable_path="/usr/bin/google-chrome")
        pg = await b.new_page()
        for name, doc in (("Analyst_Inbox_Excel_Alternative_Assignment", EX), ("Analyst_Inbox_SQL_Alternative_Assignment", SQL)):
            f = OUT / f"{name}.html"
            f.write_text(doc, encoding="utf-8")
            await pg.goto(f.resolve().as_uri())
            await pg.pdf(path=str(OUT / f"{name}.pdf"), format="Letter", print_background=True,
                         display_header_footer=True, header_template="<span></span>",
                         footer_template='<div style="font:8pt Calibri,Arial;color:#5b6a71;width:100%;text-align:center">'
                                         'HIFN1009 Analyst Inbox · alternative assignment · page <span class="pageNumber"></span> of <span class="totalPages"></span></div>',
                         margin={"top": "0.7in", "bottom": "0.8in", "left": "0.75in", "right": "0.75in"})
        await b.close()

asyncio.run(render())
print("done")
