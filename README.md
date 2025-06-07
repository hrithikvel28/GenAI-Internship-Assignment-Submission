Project Documentation:

GenAI Internship Assignment - Data Extraction & Population Task

This script dynamically maps and populates data from `source.xlsx` into `template.xlsx`, preserving its exact format using AI-powered column mapping via **Groq API** with **Llama3** model.

---

Objective:

To automate the process of filling structured Excel templates using multiple source sheets while maintaining:

- Column order
- Formatting (currency, percentages, text alignment)
- Header rows and blank cells
- Notes/comments/formulas

---

Tech Stack:

- **Language**: Python
- **Libraries**:
  - `pandas`: For data manipulation
  - `fuzzywuzzy`: Fallback for intelligent header matching
  - `openpyxl`: For reading/writing Excel files without losing formatting
  - `groq`: To call Groq API for smart mapping
- **LLM API**: Groq (`llama3-8b-8192`) for intelligent column mapping
- **Environment Source: Google Colab

---

Assumptions:

- The user has access to both `source.xlsx` and `template.xlsx`
- Template contains one main sheet (Sheet1) with hierarchical categories and subcategories
- Each facility row in the template corresponds to a facility in the source data
- All source data includes a `Facility` column and aligns with names in the template
- User provides a valid **Groq API key** when prompted



---

Setup Instructions:

1. Install dependencies:
   ```bash
   pip install pandas openpyxl fuzzywuzzy python-Levenshtein groq





Google Colab: https://colab.research.google.com/drive/1T-ZFfXr7JlTdft4kjYpXdkMKc_ELbzwE?usp=drive_link
