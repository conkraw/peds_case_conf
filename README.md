# Pediatric Case Conference Builder

A Streamlit app for building standardized, interactive pediatric case conference presentations.

The app is designed around a progressive case-reveal format:

1. Title and session goal
2. Opening case stem
3. Ask the audience what history they need
4. Reveal and interpret history
5. Ask what physical exam matters
6. Reveal and interpret exam
7. Build a problem representation
8. Differential diagnosis
9. Ask what diagnostic studies to order
10. Interpret diagnostic results
11. Management decision point
12. Disease teaching section
13. Two to three major learning points
14. Return to the patient
15. Final clinical bottom line and feedback

## Key features

- Schema-driven slide fields in `case_schema.py`
- Streamlit slide/section navigation with the same boxed editing workflow as the Journal Club Builder
- Word, line, and table-length limits
- Progressive audience-prompt sections
- PowerPoint export
- One-page Word summary export
- Mentor review Word export
- Printable planning worksheet export
- Editable JSON draft download/reload
- Optional GitHub archive storage and reload using the right-side Archive controls
- Persistent case identity display so presenter, case, and learning point remain visible while building
- Archive index with the requested columns: `presenter`, `case`, and `learning_point`
- Downloadable archive index CSV

## Files

```text
case_conference_builder/
├── app.py                         # Streamlit app
├── case_schema.py                 # Case conference slide fields, defaults, limits
├── pptx_builder.py                # PowerPoint generation
├── docx_builder.py                # Word summary and mentor review generation
├── printable_form_builder.py      # Printable worksheet generation
├── github_storage.py              # Optional GitHub archive storage
├── feedback_config.py             # Feedback QR/link configuration
├── requirements.txt               # Python dependencies
├── README.md                      # This file
├── LICENSE.txt
├── .gitignore
└── .streamlit/
    └── config.toml                # Basic Streamlit theme
```

## Run locally

From the project folder:

```bash
python -m venv .venv
source .venv/bin/activate   # Mac/Linux
# .venv\Scripts\activate    # Windows PowerShell

pip install -r requirements.txt
streamlit run app.py
```

## Upload to GitHub

1. Create a new GitHub repository.
2. Upload the contents of this folder to the repository.
3. Commit to the `main` branch.
4. Deploy from Streamlit Community Cloud with `app.py` as the main file.

## Optional GitHub case archive

The app works without GitHub storage because users can download and reload JSON drafts manually.

To enable the built-in case archive:

1. Create a private GitHub repository, for example:

```text
case-conference-drafts
```

2. Create a fine-grained GitHub personal access token limited to that repository.
3. Give it `Contents: Read and write` permission.
4. Add these secrets to Streamlit Community Cloud under **App → Settings → Secrets**:

```toml
[github]
token = "github_pat_YOUR_TOKEN_HERE"
repo = "YOUR-GITHUB-USERNAME/case-conference-drafts"
branch = "main"
base_path = "case-drafts"
```

For local testing, create `.streamlit/secrets.toml` with the same content. Do not commit `.streamlit/secrets.toml`.

## Archive index

The archive index intentionally keeps the core columns simple:

```text
presenter, case, learning_point
```

The app also tracks `saved_date`, `path`, and `archive_id` internally so saved cases can be loaded back into the builder.

## Feedback link setup

Edit `feedback_config.py` and replace:

```python
FEEDBACK_QR_URL = "https://redcap.ctsi.psu.edu/surveys/?s=REPLACE_ME"
FEEDBACK_DISPLAY_URL = "https://redcap.link/replace_me"
```

with your actual REDCap, Qualtrics, or institutional feedback links.

## Customizing the case conference structure

Most edits should happen in `case_schema.py`:

- Change section names
- Change field labels
- Change default text
- Change word/line limits
- Add or remove fields
- Modify table columns
- Update helper text

The rest of the app reads from that schema.

## Privacy reminder

The builder includes a privacy check, but it does not automatically de-identify content. Presenters should remove patient names, MRNs, exact dates, rare identifiers, and identifiable images before presenting or saving a case.
