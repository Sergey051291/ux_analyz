# UX Analyz — AI-Powered Website UX Audit

**Streamlit application for automated UX analysis of websites using LLM (Ollama or GigaChat), Selenium screenshots, and PowerPoint report generation.**

Enter a URL, capture page screenshots, run AI-driven UX evaluation against a structured prompt, and export findings as a `.pptx` presentation.

---

## Screenshots

> Add screenshots to `docs/screenshots/` and uncomment the lines below.

<!--
![Main UI](docs/screenshots/01-main.png)
![Analysis result](docs/screenshots/02-analysis.png)
![PPTX export](docs/screenshots/03-pptx.png)
-->

---

## Features

| Feature | Description |
|---------|-------------|
| **Multi-page crawl** | Analyze up to N linked pages from a starting URL |
| **Screenshot capture** | Automated browser screenshots via Selenium + Chrome |
| **LLM analysis** | UX evaluation using Ollama (local) or GigaChat (cloud) |
| **Structured prompt** | Configurable analysis template (`ux_prompt.txt`) |
| **PPTX export** | Generate PowerPoint report with screenshots and findings |
| **Analysis modes** | Basic, Extended, Full (configurable) |

---

## Tech Stack

| Layer | Technologies |
|-------|-------------|
| **UI** | Streamlit |
| **Browser automation** | Selenium + Chrome |
| **LLM** | Ollama (llama3) / GigaChat API |
| **Report generation** | python-pptx |
| **Config** | YAML (`config.yaml`) |
| **HTTP parsing** | requests, BeautifulSoup |

---

## Architecture

```
User URL input (Streamlit)
        │
        ▼
┌───────────────────┐
│  Selenium         │  Capture screenshots of pages
│  (Chrome)         │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐     ┌──────────────────┐
│  HTML extraction  │────▶│  LLM (Ollama /   │
│  + ux_prompt.txt  │     │  GigaChat)       │
└─────────┬─────────┘     └──────────────────┘
          │
          ▼
┌───────────────────┐
│  PPTX report      │  Slides with screenshots + UX findings
└───────────────────┘
```

---

## Quick Start

### Prerequisites

- Python 3.10+
- [Ollama](https://ollama.com/) running locally (`ollama serve`)
- Model: `ollama pull llama3`
- Google Chrome / Chromium installed

### Installation

```powershell
git clone https://github.com/Sergey051291/ux_analyz.git
cd ux_analyz

python -m venv .venv
.\.venv\Scripts\activate

pip install -r requirements.txt
```

### Configuration

Edit `config.yaml`:

```yaml
llm:
  provider: "ollama"
  ollama:
    base_url: "http://localhost:11434"
    model: "llama3"

prompt_template_path: "ux_prompt.txt"
```

For GigaChat, set `provider: "gigachat"` and configure credentials in `ux_analyz.py`.

### Run

Terminal 1:
```powershell
ollama serve
```

Terminal 2:
```powershell
streamlit run ux_analyz.py
```

Open `http://localhost:8501`, enter a URL, and start analysis.

---

## Project Structure

```
ux_analyz/
├── ux_analyz.py       # Main Streamlit application
├── ux_prompt.txt      # LLM prompt template for UX analysis
├── config.yaml        # Runtime configuration
├── requirements.txt
└── docs/screenshots/
```

---

## Configuration Options

| Parameter | Default | Description |
|-----------|---------|-------------|
| `max_pages` | 5 | Maximum pages to analyze per run |
| `max_html_length` | 15000 | HTML truncation limit for LLM |
| `analysis_modes` | Basic / Extended / Full | Analysis depth levels |

---

## Author

Personal project — automated UX audit tool with local LLM integration.

**Stack:** Python · Streamlit · Selenium · Ollama · python-pptx
