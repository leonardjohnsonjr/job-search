# Job Seeker Prompt Generator

A command-line tool that reads a list of companies and job descriptions from an Excel file, combines them with your resume, and generates structured AI prompts ready to paste into ChatGPT, Claude, Gemini, or any other AI web interface.

No API key required — just copy and paste the prompts into your preferred AI tool.

---

## Features

- **Excel-driven workflow** — maintain a single spreadsheet of target companies and job descriptions
- **External resume files** — your resume and summary live as `.txt` files, reused across every run
- **Five tailored AI prompts per role**
  - Company culture, challenges, and recent news
  - Hidden requirements beneath the formal job description
  - Honest gap analysis between your background and the role
  - Tailored resume summary mirroring the job description's language
  - Rubric evaluation to score AI responses against quality rules
- **Clean Markdown output** — each job gets its own `.md` file with fenced code blocks for easy copying
- **Terminal printing** — view all prompts directly in your console
- **Config file support** — set defaults once in `config.json`, override anytime via CLI

---

## Project Structure

```
.
├── job_report.py        # Main script
├── config.json          # Default settings (created by --save-config)
├── jobs_data.xlsx       # Excel file: one row per target job
├── resume.txt           # Your full resume as plain text
├── resume_summary.txt   # Your 2-3 sentence professional summary
└── reports/             # Generated prompt files (auto-created)
```

---

## Requirements

- Python 3.10+
- Excel file with Company and Job_Description columns (optional Run column)

Install dependencies:

```bash
pip install openpyxl
```

---

## Quick Start

### 1. Prepare your files

Ensure you have:
- `jobs_data.xlsx` — Excel file with Company and Job_Description columns (optional Run column)
- `resume.txt` — Your full resume as plain text
- `resume_summary.txt` — Your 2-3 sentence professional summary

### 2. Generate a starter config

```bash
python job_report.py --save-config
```

This creates `config.json` with default paths. Edit it to match your file locations:

```json
{
  "excel": "jobs_data.xlsx",
  "resume": "resume.txt",
  "resume_summary": "resume_summary.txt",
  "out_dir": "reports"
}
```

### 3. Run the generator

```bash
python job_report.py
```

This will:
- Read all rows from your Excel file
- Generate 5 AI prompts for each job
- Print them to the terminal
- Save a Markdown file per company in the `reports/` folder

### 4. Use the prompts

Open any generated `.md` file, copy a prompt from the code block, and paste it into [ChatGPT](https://chat.openai.com), [Claude](https://claude.ai), or your preferred AI tool.

---

## Excel Format

Your `jobs_data.xlsx` should have these columns:

| Company | Job_Description | Run |
|---|---|---|
| Costco | Full job description text here... | Y |
| Amazon | Another job description... |  |
| Google | Yet another... | Yes |

- **Company**: The company name
- **Job_Description**: The complete job posting text
- **Run** (optional): Set to 'Y', 'Yes', 'True', or leave blank to include this row in processing. Any other value (like 'N', 'No', 'False') will skip the row.

---

## Usage Examples

```bash
# Generate prompts for all jobs in the spreadsheet
python job_report.py

# Just print to terminal without saving files
python job_report.py --print-only

# Use a different Excel file
python job_report.py --excel my_jobs.xlsx

# Override resume files for a one-off run
python job_report.py --resume ~/docs/resume_v2.txt --resume-summary ~/docs/summary.txt

# Change output directory
python job_report.py --out-dir custom_prompts
```

---

## Config File

The `config.json` supports these fields:

- `excel` — Path to your Excel file
- `resume` — Path to your resume `.txt` file
- `resume_summary` — Path to your summary `.txt` file
- `out_dir` — Directory where prompt `.md` files are saved

CLI flags always override config values.

---

## Resume Files

Your resume and summary are plain `.txt` files kept separate from the spreadsheet so you can update them once and have every future run pick up the changes automatically.

| File | Description |
|---|---|
| `resume.txt` | Your complete resume — work history, skills, education |
| `resume_summary.txt` | A 2-3 sentence professional summary used in the tailored summary prompt |

File paths are resolved in this order:

1. CLI flags (`--resume`, `--resume-summary`)
2. `config.json` (`"resume"`, `"resume_summary"` keys)
3. Error with a helpful message if neither is provided

---

## Quality Rubric

Each generated prompt file includes a built-in quality rubric with 5 rules that AI responses should meet:

1. **Specificity** — References specific details from the job description, company, or resume
2. **Actionability** — Includes concrete, actionable takeaways
3. **Completeness** — Fully addresses the question without leaving gaps
4. **Candidate-Centric Framing** — Framed from the job seeker's perspective
5. **Professional Tone** — Clear, professional language without fluff

Use the 5th prompt ("Rubric Evaluation") to have the AI self-evaluate its own responses.

---

## CLI Reference

```
python job_report.py [OPTIONS]

Options:
  --excel          PATH    Path to Excel file
  --resume         PATH    Path to resume .txt file
  --resume-summary PATH    Path to resume summary .txt file
  --out-dir        PATH    Output directory for prompt files
  --config         PATH    Path to config file  (default: ./config.json)
  --save-config            Write a starter config.json and exit
  --print-only             Print prompts to terminal only, do not save files
  --help                   Show help message
```

### Examples

```bash
# Use all settings from config.json
python job_report.py

# Print to terminal without saving files
python job_report.py --print-only

# Use a different Excel file
python job_report.py --excel my_jobs.xlsx

# Override resume files for a one-off run
python job_report.py --resume ~/docs/resume_v2.txt --resume-summary ~/docs/summary.txt

# Change output directory
python job_report.py --out-dir custom_prompts
```

---

## Generated Prompts

The script generates five prompts for each job application:

| # | Prompt | Purpose |
|---|---|---|
| 1 | **Company Culture & Intel** | Summarize the company's culture, challenges, recent news, and suggest a smart interview question |
| 2 | **Hidden Job Requirements** | Uncover unstated priorities and personality traits behind the formal job description |
| 3 | **Gap Analysis** | Identify where your background is weakest for this role, with suggestions to address gaps |
| 4 | **Tailored Resume Summary** | Rewrite your resume summary to mirror the language and priorities of the job description |
| 5 | **Rubric Evaluation** | Evaluate AI responses against the 5 quality rules (use this after getting responses from prompts 1-4) |

Each prompt is saved in a clean Markdown file with fenced code blocks for easy copying and pasting into AI interfaces.

---

## .gitignore Recommendations

```gitignore
# Generated prompt files
reports/
prompts/

# Personal resume files (optional — omit if you want them versioned)
resume.txt
resume_summary.txt

# Excel files with personal data
jobs_data.xlsx
```

# macOS / editor artifacts
.DS_Store
__pycache__/
*.pyc
```

---

## License

MIT License. See [LICENSE](LICENSE) for details.
