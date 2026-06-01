# Powerpoint Charter (Office Open XML Chart Repair)

A Python-based utility script designed to repair and convert PowerPoint presentations containing charts with broken, missing, or missing-embedded data worksheets. 

When PowerPoint charts are unlinked from their underlying Excel data sheets (often resulting in external reference errors or empty data tables when you attempt to edit them), this script processes the underlying Open XML structure to rebuild the reference mapping. It converts literal cache values (`strLit`/`numLit`) back into cell formulas (`strRef`/`numRef`), stripping broken external references and forcing PowerPoint to reconstruct an editable local data spreadsheet upon opening.

## Features

- **Automated XML Conversion:** Parses Open XML slide chart structures (`chart*.xml`) to find standalone literal elements.
- **Formula Re-mapping:** Dynamic cell range generation that converts raw literal data arrays back into Excel-compliant cellular formula strings (e.g., `Sheet1!$A$2:$B$6`).
- **External Link Disconnection:** Safely purges the `<c:externalData>` tags causing broken-link warnings.
- **Batch Processing:** Automatically loops through a designated folder, handles decompression, structures the modifications, and rebuilds valid `.pptx` archives natively.

## How it Works

1. **Decompression:** Unzips the target `.pptx` file into a temporary directory to access its structural blueprint.
2. **XML Manipulation:** Locates chart files inside `/ppt/charts/`, targeting category (`cat`) and value (`val`) nodes.
3. **Re-linking:** Swaps static literal elements out for valid reference caching tags linked to local spreadsheet ranges.
4. **Reassembly:** Packages the updated XML files back into an Office-compliant Presentation container with proper forward-slash pathing.

## Quick Start

### Prerequisites
- Python 3.x
- Standard Python libraries (`zipfile`, `shutil`, `os`, `xml.etree.ElementTree`, `pathlib`)

### Setup and Directory Structure
1. Download or clone the Jupyter Notebook (`PPTX to PPTX error chart small.ipynb`) into your local environment.
2. Create two directories alongside the notebook file:
   - `input_pptx/`
   - `output_pptx/`
3. Place the problematic PowerPoint presentations inside the `input_pptx/` folder.

### Running the Script
Open the notebook in your favored Jupyter environment (VS Code, JupyterLab, etc.) and run all cells sequentially. The script will automatically process all `.pptx` files and output the repaired versions in your `output_pptx/` folder.

```bash
# Terminal output verification will display:
Embedding chart in chart1.xml
✅ All PPTX files processed with per-chart embedding.
