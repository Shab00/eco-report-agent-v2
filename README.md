# Ecological Survey Reporting Agent v2

A production AI agent built for Greenlight Ecology Ltd that replaces 
manual Word-template report writing. Surveyors submit structured field 
data and the agent generates a fully formatted Preliminary Ecological 
Appraisal (PEA) Word document matching Greenlight Ecology's real 
report template — including branded cover page, version control table, 
survey conditions table, 12-section species assessment table, running 
headers and footers, and appendix placeholders.

## What it generates

A complete .docx PEA report with:
- Branded cover page with Greenlight Ecology logo, sage green background
- Version control table and legal disclaimer
- Survey conditions table (date, temperature, humidity, cloud cover, wind, rain)
- Full 12-section ecological assessment table with dark green section 
  headers, alternating row shading, and formal UK ecological report language
- Species sections: Habitats, Designated Sites, Invasive Species, 
  Invertebrates, Bats, Birds, Reptiles, Amphibians, Badger, Riparian, 
  Dormouse, Hedgehog
- Running header with site name and logo on every page
- Footer with company name and roman numeral page numbers
- Appendix placeholders for habitat map, location map, proposed plan, photos

## Stack

- **Backend:** FastAPI, Pydantic v2, OpenAI function calling (GPT-4o-mini)
- **Document generation:** python-docx with strict OOXML schema ordering
- **Deployment:** Render
- **Client:** Greenlight Ecology Ltd

## API

- `POST /survey` — submit survey data, returns report_id and download_url
- `GET /report/{id}/download` — download the generated .docx file
- `GET /report/{id}` — get report metadata
- `GET /health` — health check

API docs: https://eco-report-agent-v2.onrender.com/docs

## How to run locally

```bash
git clone https://github.com/Shab00/eco-report-agent-v2
cd eco-report-agent-v2
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
echo "OPENAI_API_KEY=your-key-here" > .env
echo "LOGO_PATH=/path/to/logo.png" >> .env
uvicorn main:app --reload
```

Then open http://localhost:8000/docs to explore the API.

## What I would do next with more time

- **Frontend** — a GitHub Pages form matching the portfolio style where 
  Sam can fill in survey data and download the generated report directly
- **Photo upload** — allow surveyors to upload site photos that get 
  embedded into Appendix 4 automatically
- **Designated sites lookup** — integrate with the MAGIC database API 
  to auto-populate the designated sites section from the grid reference
- **Multi-report types** — extend the document builder to support Bat 
  Survey reports and other Greenlight Ecology report formats using the 
  same brand shell
