# E-Commerce Product Data Automation Pipeline

An end-to-end automation pipeline built to handle a large-scale **e-commerce platform migration** — scraping, enriching, converting, and AI-processing thousands of product listings with minimal manual effort.

> **Business Impact:** Automated ~90% of the data migration workload, reducing processing time by 10x while maintaining high data accuracy — saving the company significant time and operational cost.

---

## What it does

| Step | Script | Description |
|------|--------|-------------|
| 1 | `getDataAuto.py` | Scrapes product **name**, **description**, and **specifications** from the source platform |
| 2 | `imageScraper.py` | Downloads product **images** — tries the primary source first, falls back to a barcode lookup service |
| 3 | `measurementConverter.py` | Adds **metric equivalents** inline to all imperial measurements (e.g. `5 in (12.7 cm)`) |
| 4 | `aiProcessor.py` | Uses **Groq AI (LLaMA 3.3 70B)** to generate SEO meta title, meta description, tags, and brand name |

---

## Project structure

```
ecommerce-migration-pipeline/
├── config.py                 # Central config — all settings in one place
├── run_all.py                # Pipeline orchestrator — run all or individual steps
├── getDataAuto.py            # Step 1 — Product data scraper
├── imageScraper.py           # Step 2 — Image downloader
├── measurementConverter.py   # Step 3 — Measurement converter
├── aiProcessor.py            # Step 4 — AI processor (Groq)
├── requirements.txt          # Python dependencies
├── .env.example              # Environment variable template
└── README.md
```

---

## The problem this solves

Migrating a large product catalog between e-commerce platforms traditionally requires a team of people manually:
- Copying product names, descriptions, and specifications one by one
- Downloading and renaming product images
- Converting imperial measurements to metric for regional markets
- Writing SEO metadata for every product

This pipeline automates all of it. A task that would take months of manual work runs in hours with a single command.

---

## Setup

### 1. Clone the repo

```bash
git clone https://github.com/your-username/ecommerce-migration-pipeline.git
cd ecommerce-migration-pipeline
```

### 2. Create a virtual environment

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Mac / Linux
source venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

### 4. Configure your `.env` file

Copy the example and fill in your values:

```bash
cp .env.example .env
```

Then edit `.env`:

```env
EXCEL_FILE_PATH=path/to/your/products.xlsx
IMAGES_FOLDER=path/to/save/images/
GROQ_API_KEYS=gsk_your_key_here
AI_MODEL=llama-3.3-70b-versatile
```

Get a free Groq API key at [console.groq.com](https://console.groq.com).

---

## Running the pipeline

### Run everything

```bash
python run_all.py
```

### Run specific steps only

```bash
python run_all.py 1          # scrape data only
python run_all.py 3 4        # convert measurements, then run AI
python run_all.py 1 2 3 4    # same as running everything
```

---

## Configuration reference

All settings live in `config.py` and are loaded from your `.env` file.

| Setting | Default | Description |
|---------|---------|-------------|
| `EXCEL_FILE_PATH` | *(required)* | Full path to your Excel file |
| `IMAGES_FOLDER` | *(required)* | Folder where product images are saved |
| `GROQ_API_KEYS` | *(required)* | Your Groq API key |
| `AI_MODEL` | `llama-3.3-70b-versatile` | Groq model to use |
| `SHEET_NAME` | `Sheet1` | Excel sheet name |
| `START_ROW` | `1` | First data row (1 = row after header) |
| `MAX_ROWS_PER_RUN` | `None` | Limit rows per run — `None` means all rows |
| `ITEM_NUMBER_COLUMN` | `D` | Column with item numbers |
| `BARCODE_COLUMN` | `F` | Column with barcodes (image fallback) |
| `COL_NAME_EN` | `H` | Output: product name |
| `COL_DESC_EN` | `J` | Output: description |
| `COL_SPECS_EN` | `L` | Output: specifications |
| `COL_META_TITLE_EN` | `N` | Output: SEO meta title |
| `COL_META_DESC_EN` | `P` | Output: SEO meta description |
| `COL_TAGS_EN` | `R` | Output: tags |
| `COL_BRAND_EN` | `T` | Output: brand name |

---

## Key technical features

- **Resumable runs** — every step skips rows that already have data, so you can stop and continue at any point without re-processing completed work
- **Dual-source image fallback** — if the primary source doesn't have an image, automatically falls back to a barcode lookup service
- **Intelligent measurement conversion** — regex engine handles fractions, mixed numbers, dimension chains, and unit abbreviations while correctly ignoring numbers that are catalogue codes or part of sentences
- **AI key rotation** — automatically rotates through multiple API keys when rate limits are hit, keeping the pipeline running without interruption
- **Fuzzy spec extraction** — multiple CSS selector fallbacks with accordion/lazy-load handling ensure specs are captured even across different page layouts
- **Safe to re-run** — anti-duplicate guards prevent stacking conversions or overwriting already-processed cells

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `playwright` | Browser automation for scraping |
| `openpyxl` | Reading and writing Excel files |
| `requests` | Downloading images |
| `groq` | AI API for metadata generation |
| `pydantic-settings` | Config management via `.env` |