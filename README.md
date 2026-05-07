# EMP Electrical Quotation Estimator

AI-powered quotation generator for Electro Metal Pressings (Pvt) Ltd.  
Upload SLD drawings (PDF) + the EMP Costing Model Excel to auto-generate a priced quotation.

## Deployment on Streamlit Community Cloud

### 1. Fork / push this repo to GitHub

### 2. Add your Anthropic API key as a Streamlit Secret

In your Streamlit Cloud dashboard → App Settings → Secrets:

```toml
ANTHROPIC_API_KEY = "sk-ant-xxxxxxxxxxxxxxxxxxxx"
```

The app reads this automatically via the `anthropic` SDK (which checks `ANTHROPIC_API_KEY` env var).

### 3. Deploy

- Repository: your GitHub repo
- Branch: main
- Main file path: `app.py`

---

## How to Use

1. **Upload SLD PDFs** — all Single Line Diagram PDF sets for the project  
2. **Upload EMP Costing Model** — `Quotation_EMP_Costing_Model_V2_82.xlsm` (or newer)  
3. **Fill project details** in the sidebar (customer, ref no., etc.)  
4. Click **Extract Components** — Claude AI reads the SLDs and identifies every panel & component  
5. **Review & edit** the extracted component lists  
6. Click **Calculate & Generate** — produces the priced quotation  
7. **Download** the output Excel  

## Output Excel Contains

- **Offer** sheet — quotation letter with line items  
- **Cost Summary** sheet — material / labour / margin breakdown per panel  
- **BOM** sheet — full bill of materials with price matching status  

## Architecture

```
SLD PDFs ──► Claude Vision AI ──► Component Extraction
                                          │
EMP Costing Model ──► Price Database ─────┤
                                          ▼
                              Cost Calculator (with margin)
                                          │
                                          ▼
                              Output Quotation Excel
```
