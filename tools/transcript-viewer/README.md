# Transcript viewer

Turns transcripts downloaded from the Analyst Inbox apps (Excel and SQL; .html, or older .txt files) into a readable page, with tutor
notes at the top: requests Dana sent, the student's predictions, possible stumbling points, the sentences
where Dana corrected something, questions she had to repeat, long pauses, the student's reflection and
Dana's "where to grow".

Two versions do the same thing:

| File | Use it on | How |
| --- | --- | --- |
| `transcript_viewer.html` | iPhone, iPad, any browser | Open it (or the published copy on claude.ai), tap **Choose transcript(s)**, pick transcript files. Nothing is uploaded. |
| `transcript_viewer.py` | Windows PC | Needs Python 3.9+ from python.org. Double-click it and pick transcripts, or drag files onto it. Saves `..._readable.html` next to each transcript and opens it; several files also get an index page. |

Flags come from keyword matching, so treat them as pointers to look at, not verdicts.

Student transcripts are personal data: don't commit them to this repository.
