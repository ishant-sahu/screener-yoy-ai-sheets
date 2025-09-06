YOY to Sheets (Latest Transcript with PDFLib)

This script scrapes quarterly financial data and the latest conference call transcript of a company from Screener.in
, calculates YoY growth (Sales & Profit), summarizes the transcript using OpenAI GPT, and appends the data into a Google Sheet.

It also applies conditional cell coloring in Google Sheets based on YoY growth.

📂 Project Structure
index.js # Main script
credentials.json # Google Sheets API credentials
.env # Environment variables

⚙️ Features

🔍 Scrapes quarterly sales & net profit from Screener.in

📈 Calculates YoY growth for the last 6 quarters

📄 Extracts text from PDF transcripts using pdf-lib

🤖 Uses OpenAI GPT to summarize transcripts & extract insights

📊 Appends results to Google Sheets

🎨 Applies conditional formatting:

Green if YoY growth ≥ 10%

Red if YoY growth < 10%

White if data unavailable

🔑 Requirements

Node.js (v18+ recommended)

Google Sheets API credentials (credentials.json)

Screener.in company page URL

OpenAI API key

📦 Installation
git clone <repo_url>
cd <repo_folder>
npm install axios cheerio googleapis dotenv openai pdf-lib

🔐 Setup Environment Variables

Create a .env file:

OPENAI_API_KEY=your_openai_api_key
SHEET_ID=your_google_sheet_id
SHEET_TAB=Sheet1

OPENAI_API_KEY: Your OpenAI API key

SHEET_ID: The ID of your target Google Sheet (found in the URL)

SHEET_TAB: Name of the Google Sheet tab (default: Sheet1)

🔑 Setup Google Sheets API

Go to Google Cloud Console

Enable Google Sheets API

Create a Service Account and download credentials.json

Share your Google Sheet with the service account email

🚀 Usage
node yoy-to-sheets-latest-transcript-pdflib.js "https://www.screener.in/company/ZAGGLE/#quarters"

Example Output:

🔍 Scraping quarterly data from: https://www.screener.in/company/ZAGGLE/#quarters
📄 Fetching latest transcript...
Transcript link https://www.screener.in/transcripts/123456.pdf
🤖 Getting AI-powered company insights...
📊 Row: ZAGGLE | FinTech | Payments | 15.6 | 20.1 | ... | Summary text | Guidance text
✅ Row appended and colored to "Sheet1"

🧮 How YoY is Calculated

For each of the last 6 quarters:

YOY Sales Growth (%) = ((Current Sales - Sales Last Year) / Sales Last Year) _ 100
YOY Profit Growth (%) = ((Current Profit - Profit Last Year) / Profit Last Year) _ 100

If data for the same quarter in the previous year is missing, it stores -.

📊 Google Sheet Layout
Symbol Sector Sub-Sector Q1 Sales YoY Q1 Profit YoY ... Summary Guidance
🔍 Tech Stack

Node.js for execution

Axios for HTTP requests

Cheerio for HTML scraping

Google Sheets API for data writing

OpenAI GPT-4.1-mini for AI insights

pdf-lib for PDF text extraction

🛠️ Troubleshooting
Issue Solution
No Sales/Profit rows found Screener.in changed its table structure. Update scraper selectors.
Transcript not available Company has no transcript uploaded yet.
PERMISSION_DENIED on Sheets Share sheet with service account email from credentials.json.
Error: 403 on OpenAI Check OPENAI_API_KEY.
Script fails extracting PDF text Some transcripts are image-based PDFs. Consider OCR.
📜 License

MIT License. Use at your own risk.
