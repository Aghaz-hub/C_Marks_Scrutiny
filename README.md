# Exam Cover Page → Excel

A Streamlit app for faculty. Upload photos, scans or PDFs of answer-script **cover pages**
(the page with the *Marks Distribution* table). The app reads the handwriting with Claude,
lets you review and correct the result, and gives you one Excel workbook with three tabs:

| Tab | What it contains |
|---|---|
| **Script Details** | Exam, subject, code, date, session, semester, number of scripts, number flagged for review |
| **Marks Check** | Per student, question-wise: parts, sum of parts, total recorded, difference, status (live Excel formulas), plus the box total and marks-in-words check |
| **Student-wise Marks** | **One row per student, sorted by roll number**: Roll No, Name, Q.No1(a), Q.No1(b) … Q.No1 Total … Total Marks, Marks in Words, Review |

Uploads can be JPG, PNG or PDF; you can add many files at once, and one PDF can hold many cover pages.
Rows are flagged **CHECK** when a question's parts do not add up to its total, the question totals do not add up to
the sheet total, the words do not match the number, a roll number is duplicated, or the model was unsure of a digit.

## Put it on GitHub and Streamlit Community Cloud

1. Create a new GitHub repository (private is recommended) and upload these files, keeping the structure:
   ```
   app.py
   requirements.txt
   README.md
   .gitignore
   .streamlit/config.toml
   .streamlit/secrets.toml.example
   ```
   Do **not** upload a real `secrets.toml` or your API key.
2. Get an API key from the Anthropic Console (console.anthropic.com).
3. Go to share.streamlit.io, sign in with GitHub, choose **New app**, pick your repository, branch `main`,
   main file `app.py`.
4. Before deploying, open **Advanced settings → Secrets** and paste:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-your-key-here"
   APP_PASSWORD = "choose-a-strong-password"
   ```
5. Click **Deploy**. Faculty open the app link and enter the password you chose; no API key is needed on their side.
   The app refuses to open if `APP_PASSWORD` is not set, so your API credits are never exposed by mistake.

If you do not set the secret, the sidebar shows a box where each user can paste their own key.

## Run locally

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # then edit and add your key
streamlit run app.py
```

## Good to know

- **Always review the CHECK rows.** Handwriting reading is good but not perfect; the totals checks exist to catch slips.
  Final marks remain the examiner's responsibility.
- **Cost:** each page is one API call. Sonnet (default) balances accuracy and cost; Haiku is cheaper but less accurate on
  handwriting. Use the sidebar "Maximum pages" limit to cap a run.
- **Scan quality:** flat page, even light, whole page in frame. Phone scanner apps (Google Drive Scan, Adobe Scan) work well.
- **Privacy:** page images are sent to the Anthropic API. Check your university's examination and data rules before using
  real scripts, and keep the repository private. Nothing is stored by the app; results live only in the browser session.
- **Different sheet layout?** Change the number of questions in the sidebar (default 10). Parts a to j are supported.
- Model names are listed in `MODELS` at the top of `app.py`; update them there if Anthropic releases newer ones.
