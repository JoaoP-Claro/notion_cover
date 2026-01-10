# Weekly Excel Range Snapshot – Full Project Plan (No Omissions)

Awesome project. Here’s a complete, company-grade plan you can implement end-to-end: architecture, repo structure, PR process, QA/QC, and CI/CD—plus ready-to-drop files/workflows.

---

## 1) Goal & Constraints (TL;DR)
- **Input:** an `.xlsx` workbook in your repo.
- **Task:** once per week, take a **picture of a specific Excel range** (no desktop Excel required), overwrite last week’s picture, and publish it as a static web page anyone can open in a browser.
- **Approach:** Python + OpenPyXL → render range to HTML → **Playwright (headless Chromium)** screenshots the HTML table → **GitHub Pages** hosts `index.html` + `latest.png`.
- **Governance:** feature branches → PRs with reviews, lint, tests → protected `main`.
- **Automation:** GitHub Actions (CI on PR; weekly scheduled deploy to Pages).

This avoids Windows/Excel COM and runs fully on Linux runners.

---

## 2) High-level Architecture

1. **Data layer**  
   - `data/source.xlsx` (versioned in the repo).  
   - Config file `config.yml` to choose sheet/range (e.g., `Sheet1!A1:D20`) and visual options.

2. **Render layer**  
   - `openpyxl` reads the sheet range (values + some basic formatting).  
   - A small renderer converts the cells to **semantic HTML** (table + inline CSS).  
   - **Playwright** renders HTML and screenshots just the table element to PNG.

3. **Publish layer**  
   - Generated files: `site/index.html`, `site/latest.png` (always the same name).  
   - **GitHub Pages** serves `/` (index) and the image at `/latest.png`.

4. **Automation**  
   - **CI (pull_request):** lint (ruff), type check (mypy), tests (pytest).  
   - **CD (schedule + manual):** weekly cron on `main` → build → deploy to Pages.

---

## 3) Repository Structure

```
excel-snapshots/
├─ .github/
│  ├─ workflows/
│  │  ├─ ci.yml                 # PR checks (ruff, mypy, pytest)
│  │  └─ deploy.yml             # weekly snapshot + Pages deploy
│  ├─ PULL_REQUEST_TEMPLATE.md
│  ├─ ISSUE_TEMPLATE/
│  │  ├─ bug_report.md
│  │  └─ feature_request.md
│  └─ CODEOWNERS
├─ data/
│  └─ source.xlsx               # your workbook
├─ renderer/
│  ├─ __init__.py
│  ├─ excel_to_html.py          # read range → HTML string
│  └─ html_template.j2          # Jinja2 template for pretty table
├─ site/                        # built artifacts (gitignored except index skeleton)
│  └─ .gitkeep
├─ tests/
│  ├─ test_range_parse.py
│  └─ test_html_render.py
├─ config.yml                   # sheet & range configuration
├─ generate_image.py            # CLI: builds HTML and latest.png
├─ requirements.txt
├─ mypy.ini
├─ pyproject.toml               # ruff config
├─ README.md
└─ LICENSE
```

---

## 4) Configuration (`config.yml`)

```yaml
workbook: "data/source.xlsx"
range: "Sheet1!A1:D20"      # Excel A1 notation
output_dir: "site"
image_name: "latest.png"
page_title: "Weekly Excel Snapshot"
table_max_width_px: 1280
cell_padding_px: 6
font_family: "Inter, Arial, sans-serif"
border_collapse: true
```

---

## 5) Dependencies (`requirements.txt`)

```txt
openpyxl==3.1.5
jinja2==3.1.4
playwright==1.46.0
pydantic==2.8.2
pyyaml==6.0.2
```

> Playwright needs a one-time browser install in CI: `playwright install --with-deps chromium`.

---

## 6) Excel → HTML Renderer (`renderer/excel_to_html.py`)

```python
from __future__ import annotations
from dataclasses import dataclass
from typing import List, Any, Tuple
import re
from openpyxl import load_workbook
from jinja2 import Environment, FileSystemLoader
from pathlib import Path

RANGE_RE = re.compile(r"^(?P<sheet>[^!]+)!(?P<start>[A-Z]+[0-9]+):(?P<end>[A-Z]+[0-9]+)$")

@dataclass
class Cell:
    value: Any
    bold: bool
    italic: bool
    bg_hex: str | None
    font_color_hex: str | None
    number_format: str | None

def _a1_to_rowcol(a1: str) -> Tuple[int, int]:
    m = re.match(r"^([A-Z]+)([0-9]+)$", a1)
    if not m:
        raise ValueError(f"Invalid A1: {a1}")
    col_letters, row = m.group(1), int(m.group(2))
    col = 0
    for ch in col_letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return row, col

def parse_range(range_str: str) -> Tuple[str, Tuple[int,int], Tuple[int,int]]:
    m = RANGE_RE.match(range_str.replace(" ", ""))
    if not m:
        raise ValueError("Range must be like 'Sheet1!A1:D20'")
    sheet = m.group("sheet")
    start = _a1_to_rowcol(m.group("start"))
    end   = _a1_to_rowcol(m.group("end"))
    return sheet, start, end

def read_range(workbook_path: str, range_str: str) -> List[List[Cell]]:
    wb = load_wb(workbook_path)
    sheet_name, (r1,c1), (r2,c2) = parse_range(range_str)
    ws = wb[sheet_name]
    table: List[List[Cell]] = []
    for r in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        row_cells: List[Cell] = []
        for cell in r:
            font = cell.font
            fill = cell.fill
            bg_hex = None
            if getattr(fill, "fgColor", None) and fill.fgColor.type == "rgb" and fill.fgColor.rgb:
                rgb = fill.fgColor.rgb
                if rgb and len(rgb) in (6,8):
                    bg_hex = "#" + (rgb[-6:])
            font_color_hex = None
            if getattr(font, "color", None) and font.color and font.color.type == "rgb" and font.color.rgb:
                rgb = font.color.rgb
                font_color_hex = "#" + (rgb[-6:])
            val = cell.value
            row_cells.append(Cell(
                value=val,
                bold=bool(getattr(font, "bold", False)),
                italic=bool(getattr(font, "italic", False)),
                bg_hex=bg_hex,
                font_color_hex=font_color_hex,
                number_format=getattr(cell, "number_format", None)
            ))
        table.append(row_cells)
    return table

def format_value(val: Any, number_format: str | None) -> str:
    if val is None:
        return ""
    if isinstance(val, (int, float)) and number_format:
        # very light formatting – expand as needed
        if "%" in number_format:
            return f"{val:.2%}"
        if "$" in number_format or "€" in number_format or "R$" in number_format:
            return f"{val:,.2f}"
        if "0.00" in number_format:
            return f"{val:,.2f}"
    return str(val)

def load_wb(path: str):
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Workbook not found: {path}")
    return load_workbook(filename=path, data_only=True)

def render_html(
    workbook_path: str,
    range_str: str,
    template_dir: str,
    title: str,
    table_max_width_px: int = 1280,
    cell_padding_px: int = 6,
    font_family: str = "Inter, Arial, sans-serif",
    border_collapse: bool = True
) -> str:
    table = read_range(workbook_path, range_str)
    rows = [
        [
            {
              "text": format_value(c.value, c.number_format),
              "bold": c.bold, "italic": c.italic,
              "bg": c.bg_hex, "fg": c.font_color_hex
            } for c in row
        ] for row in table
    ]
    env = Environment(loader=FileSystemLoader(template_dir), autoescape=True)
    tmpl = env.get_template("html_template.j2")
    return tmpl.render(
        title=title,
        rows=rows,
        table_max_width_px=table_max_width_px,
        cell_padding_px=cell_padding_px,
        font_family=font_family,
        border_collapse=border_collapse,
    )
```

---

## 7) HTML Template (`renderer/html_template.j2`)

```html
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>{{ title }}</title>
  <style>
    :root { --pad: {{ cell_padding_px }}px; --maxw: {{ table_max_width_px }}px; }
    body { font-family: {{ font_family }}; margin: 24px; }
    .wrap { max-width: var(--maxw); }
    table { width: 100%; {% if border_collapse %}border-collapse: collapse;{% endif %} }
    td, th { padding: var(--pad); border: 1px solid rgba(0,0,0,.1); }
    .meta { color: #666; margin: 8px 0 16px; font-size: 0.9rem; }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>{{ title }}</h1>
    <div class="meta">Updated: <span id="updated"></span></div>
    <div id="capture">
      <table aria-label="Excel Range">
        {% for row in rows %}
        <tr>
          {% for cell in row %}
          <td style="{% if cell.bg %}background: {{cell.bg}};{% endif %}{% if cell.fg %}color: {{cell.fg}};{% endif %}{% if cell.bold %}font-weight:bold;{% endif %}{% if cell.italic %}font-style:italic;{% endif %}">
            {{ cell.text }}
          </td>
          {% endfor %}
        </tr>
        {% endfor %}
      </table>
    </div>
    <p>Static image version (auto-refreshed weekly):</p>
    <img src="latest.png" alt="Latest snapshot" style="max-width:100%;height:auto"/>
  </div>
  <script>
    document.getElementById('updated').textContent = new Date().toLocaleString();
  </script>
</body>
</html>
```

---

## 8) Image Generator CLI (`generate_image.py`)

```python
import yaml, sys, argparse, time
from pathlib import Path
from renderer.excel_to_html import render_html
from playwright.sync_api import sync_playwright

def write_file(path: Path, content: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")

def screenshot_table(html_path: Path, out_png: Path):
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(device_scale_factor=2)  # crisp image
        page.goto(html_path.as_uri())
        page.wait_for_selector("#capture")
        loc = page.locator("#capture")
        box = loc.bounding_box()
        # some renderers need a tiny wait for fonts/layout
        time.sleep(0.2)
        loc.screenshot(path=str(out_png))
        browser.close()

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default="config.yml")
    args = ap.parse_args()

    cfg = yaml.safe_load(Path(args.config).read_text(encoding="utf-8"))
    workbook = cfg["workbook"]
    rng = cfg["range"]
    out_dir = Path(cfg.get("output_dir", "site"))
    image_name = cfg.get("image_name", "latest.png")
    title = cfg.get("page_title", "Weekly Excel Snapshot")

    html = render_html(
        workbook_path=workbook,
        range_str=rng,
        template_dir="renderer",
        title=title,
        table_max_width_px=cfg.get("table_max_width_px", 1280),
        cell_padding_px=cfg.get("cell_padding_px", 6),
        font_family=cfg.get("font_family", "Inter, Arial, sans-serif"),
        border_collapse=cfg.get("border_collapse", True)
    )

    index_html = out_dir / "index.html"
    write_file(index_html, html)

    out_png = out_dir / image_name
    screenshot_table(index_html, out_png)

    print(f"Generated {index_html} and {out_png}")

if __name__ == "__main__":
    sys.exit(main())
```

---

## 9) Tests (examples)

**`tests/test_range_parse.py`**
```python
import pytest
from renderer.excel_to_html import parse_range

def test_parse_range_ok():
    sheet, start, end = parse_range("Sheet1!A1:D20")
    assert sheet == "Sheet1"
    assert start == (1,1)
    assert end == (20,4)

def test_parse_range_invalid():
    with pytest.raises(ValueError):
        parse_range("A1:D20")  # missing sheet
```

**`tests/test_html_render.py`**
```python
from renderer.excel_to_html import render_html
def test_template_renders(tmp_path):
    html = render_html(
        "data/source.xlsx", "Sheet1!A1:A1", "renderer", "Test"
    )
    assert "<table" in html and "Test" in html
```

---

## 10) Lint/Type configs

**`pyproject.toml` (ruff)**
```toml
[tool.ruff]
line-length = 100
select = ["E","F","I","UP","B"]
```

**`mypy.ini`**
```ini
[mypy]
python_version = 3.11
ignore_missing_imports = True
```

---

## 11) GitHub Actions – CI for PRs (`.github/workflows/ci.yml`)

```yaml
name: CI
on:
  pull_request:
    branches: [ "main" ]
  push:
    branches: [ "feature/**" ]

jobs:
  test:
    runs-on: ubuntu-latest
    permissions:
      contents: read
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: "3.11" }
      - run: python -m pip install --upgrade pip
      - run: pip install -r requirements.txt pytest mypy ruff
      - run: ruff check .
      - run: mypy .
      - run: pytest -q
```

---

## 12) GitHub Actions – Weekly Snapshot & Deploy to Pages (`.github/workflows/deploy.yml`)

```yaml
name: Weekly Snapshot & Deploy
on:
  schedule:
    - cron: "0 9 * * 1"   # every Monday 09:00 UTC
  workflow_dispatch: {}    # manual run

permissions:
  contents: read
  pages: write
  id-token: write

concurrency:
  group: "pages"
  cancel-in-progress: false

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: "3.11" }
      - run: python -m pip install --upgrade pip
      - run: pip install -r requirements.txt
      - name: Install Playwright browsers
        run: python -m playwright install --with-deps chromium
      - name: Generate image & site
        run: python generate_image.py --config config.yml
      - name: Upload Pages artifact
        uses: actions/upload-pages-artifact@v3
        with:
          path: ./site

  deploy:
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    runs-on: ubuntu-latest
    needs: build
    steps:
      - id: deployment
        uses: actions/deploy-pages@v4
```

> After this workflow is merged to `main`, enable **Settings → Pages → Build and deployment → GitHub Actions**.

---

## 13) Governance: Branching, Reviews, Protections

- **Branch model:** `main` (protected) + `feature/<ticket>` branches.
- **Branch protections (Settings → Branches):**
  - Require PR reviews (≥1 or 2, as you prefer).
  - Require status checks to pass: `CI` must be green.
  - Dismiss stale approvals on new commits.
  - Require linear history (optional).
- **CODEOWNERS** (auto-request reviewers):

**`.github/CODEOWNERS`**
```
*       @your-org/data-platform
/renderer/ @your-org/python-guild
/.github/  @your-org/devops
```

- **PR template** (`.github/PULL_REQUEST_TEMPLATE.md`)
```md
## Summary
- What & why

## Changes
- Key changes

## Validation
- [ ] Unit tests added/updated
- [ ] Snapshot generated locally (`python generate_image.py`)
- [ ] Visual check of `site/index.html` and `site/latest.png`

## Risks & Rollback
- How to rollback
```

- **Issue templates** (bugs/features) assumed in repo structure above.

---

## 14) QA/QC Process

**Automated:**
- Lint (ruff), typing (mypy), unit tests (pytest).
- Build verification (CI) ensures the renderer compiles against the workbook.

**Manual (PR author):**
- Run locally:
  ```bash
  python -m venv .venv && source .venv/bin/activate
  pip install -r requirements.txt
  python -m playwright install --with-deps chromium
  python generate_image.py
  open site/index.html
  ```
- Confirm `latest.png` matches expected range & formatting.
- Check empty/merged cells display reasonably.

**Reviewers:**
- Validate `config.yml` changes (range/sheet).
- Validate visual output (attach artifact or screenshot to PR).

**Rollbacks:**
- Revert the last commit on `main` or redeploy a previous Pages artifact by re-running the workflow on an earlier SHA (via `workflow_dispatch` → “Run workflow on branch: main @ <commit>`).

---

## 15) Security & Secrets

- No external secrets required.
- Use GitHub-provided `pages:write` / `id-token:write` permissions (already set in `deploy.yml`).
- If in the future you fetch the workbook from a private source, add a repository secret and read it in CI.

---

## 16) Operations & SLO

- **Schedule:** Weekly (adjust cron as needed; Brasília time note: cron is UTC).
- **SLO:** Page available 99.9% (GitHub Pages is highly reliable).
- **Monitoring:** Optional—add a lightweight external ping (UptimeRobot) to the Pages URL.
- **Cost:** $0 on public repos; private repos with Actions usage fall under your plan minutes.

---

## 17) Variations / Alternatives

- **AWS S3 + CloudFront** instead of GitHub Pages (use OIDC to push `site/*` to S3 in the scheduled job).  
- **Google Sheets source** if you ever want live data via API.  
- **Richer styling**: expand OpenPyXL style mapping, or use pandas `Styler` to HTML first.  
- **Multiple ranges**: loop `ranges:` in `config.yml` and produce multiple PNGs + an `index.html` gallery.

---

## 18) First-Run Checklist

1. Create repo with the structure above; add your `data/source.xlsx`.
2. Commit all files; open a PR from `feature/initial`.
3. CI should pass. Review & merge.
4. In **Settings → Pages**, select **GitHub Actions**.
5. Manually trigger **Weekly Snapshot & Deploy** (Actions → deploy → Run workflow).
6. Visit your Pages URL (shown in the workflow’s **deploy** step output). You’ll see the table and the `latest.png`, and every week it updates in place.

---

If you want, share the sheet name and cell range you plan to capture—I can prefill `config.yml` and tune the HTML template (e.g., number formats, column widths, currency style `R$`, etc.).
