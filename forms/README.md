# Survey instrument (Google Form)

[`create_google_form.gs`](create_google_form.gs) is a Google Apps Script that creates the
**Student Wellbeing Survey** Google Form, plus a linked Google Sheet that collects the responses.

## Create the form (about 2 minutes)

1. Go to <https://script.google.com> and click **New project**.
2. Paste in the contents of `create_google_form.gs` and click **Save**.
3. Choose `createStudentWellbeingForm` and click **Run**. The first run asks you to allow access to Forms and Sheets.
4. Open the **Execution log**. It shows three links: the public form link to share, the edit link, and the response sheet.

## Questions and matching data columns

| # | Question | Answer type | Column in `data/student_sample_317.csv` |
|---|---|---|---|
| 0 | Consent (18+, agrees to take part) | Yes / No (No ends the form) | — |
| 1 | Gender | Male / Female / Other or prefer not to say | `Gender` |
| 2 | Age | number, 16–40 | `Age` |
| 3 | Family background | Lower Middle / Middle / Rich class | `Family_background` |
| 4 | Relationship status | Single / In a Relationship / Complicated | `Relationship_status` |
| 5 | Living style | Bachelor / Family | `Living_style` |
| 6 | Current semester | number, 0–14 | `Current_semester` |
| 7 | CGPA | number, 0–10 (10-point scale) | `CGPA` |
| 8 | Friend circle | Yes / No | `friend_circle` |
| 9 | Health condition | Good / Normal / Abnormal / Critical | `Health_condition` |
| 10 | Monthly living expense excluding study costs | Low (< 10,000) / Mid (10,000–14,999) / High (≥ 15,000) | `Living_expense` |
| 11 | Felt depressed over the past few months | Yes / No | `Depression` |

The answer options use the same wording as the dataset. Responses exported from the sheet can be renamed to
these column names and analysed with the notebook and the Stata do-file without changes.
Two exceptions: the dataset has only Male and Female, so drop any "Other / Prefer not to say" answers or code
them separately; and the analysis sample covers only students with a CGPA above 6.00, so apply the same filter
before running the models.

The analysis in this repository uses an existing survey dataset (see the report). This form is the
survey instrument for the study's variables. It is not the source of those 317 responses.
