## Getting Started

1. Clone the repository.
2. Install the required dependencies.
3. Configure the application.
4. Run the application using the project instructions.

Before install dependencies check you internet connection.

# OpsPilot AI / MISFlow AI

A lightweight browser-based business reporting dashboard for analyzing location-wise revenue, expense, and EBITDA from spreadsheet data.

This app lets teams upload CSV, XLS, or XLSX files, map report columns, review uploaded records, and export a cleaned summary report without sending files to a backend.

## Features

- Upload multiple CSV/XLS/XLSX files
- Drag-and-drop file upload support
- Automatic column detection for location, revenue, and expense fields
- Manual mapping controls for custom spreadsheet layouts
- Data preview before generating the report
- Location-wise KPI aggregation
- Search and sorting for business reports
- EBITDA calculation and visualization
- CSV export of generated report
- Privacy-first browser-only processing

## Tech Stack

- HTML5
- CSS3
- Vanilla JavaScript
- SheetJS (XLSX)

## Project Structure

```text
.
├── index.html
├── app.js
├── styles.css
├── README.md
└── .gitignore
```

## How it Works

1. Upload one or more spreadsheet files.
2. The app reads the document in the browser using SheetJS.
3. You map the relevant columns for location, revenue, and expense.
4. The app aggregates data by location and computes EBITDA.
5. You can search, sort, and export the final report as CSV.

## Running the App Locally

Because this project is a static front-end app, you can run it in any local web server.

### Option 1: Open directly

Open `index.html` in a browser.

### Option 2: Run a local server

```bash
python -m http.server 8000
```

Then visit:

```text
http://localhost:8000
```

## Example Workflow

1. Download the sample Excel file from the app UI.
2. Re-upload it to test the import pipeline.
3. Review mapped columns.
4. Verify the generated table and KPI cards.
5. Export the final report as CSV.

## Data Privacy

This application processes uploaded files only in the browser session. No files are uploaded to a server, database, or external API.

## Notes

- Supported file types: `.csv`, `.xls`, `.xlsx`
- Maximum file size: 10 MB per file
- The app is designed for reporting workflows and lightweight business analysis

## License

This project does not currently include a repository license file. If you plan to distribute or reuse it publicly, add a license before production use.

## Summary

OpsPilot AI / MISFlow AI is a front-end MIS reporting and business analytics tool that transforms spreadsheet data into a summarized revenue and EBITDA dashboard with zero backend dependency.
