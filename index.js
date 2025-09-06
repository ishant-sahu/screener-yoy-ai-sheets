/**
 * Script: yoy-to-sheets-latest-transcript-pdflib.js
 * Usage: node yoy-to-sheets-latest-transcript-pdflib.js "https://www.screener.in/company/ZAGGLE/#quarters"
 */

import axios from 'axios';
import * as cheerio from 'cheerio';
import { google } from 'googleapis';
import path from 'path';
import dotenv from 'dotenv';
import OpenAI from 'openai';
import { PDFDocument } from 'pdf-lib';

dotenv.config();

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// Google Sheets client
async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: path.join(process.cwd(), 'credentials.json'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

// Scrape quarterly data
async function scrapeQuarterlyData(url) {
  const { data } = await axios.get(url, {
    headers: { 'User-Agent': 'Mozilla/5.0' },
  });
  const $ = cheerio.load(data);

  const symbol = url.split('/')[4].toUpperCase();
  const about = $('.company-profile').text().trim() || 'No company profile.';

  const rows = [];
  $('table.data-table tr').each((_, el) => {
    const cols = $(el)
      .find('th, td')
      .map((_, td) => $(td).text().trim())
      .get();
    if (cols.length) rows.push(cols);
  });

  const quarters = rows[0].slice(1);
  const findRow = (name) =>
    rows.find((r) => r[0].toLowerCase().includes(name.toLowerCase()));

  const salesRow = findRow('sales');
  const profitRow = findRow('net profit');
  if (!salesRow || !profitRow) throw new Error('No Sales/Profit rows found.');

  const sales = salesRow
    .slice(1)
    .map((v) => parseFloat(v.replace(/,/g, '')) || 0);
  const profit = profitRow
    .slice(1)
    .map((v) => parseFloat(v.replace(/,/g, '')) || 0);

  return {
    symbol,
    about: about.slice(0, 2000),
    quarters: quarters.map((q, i) => ({
      quarter: q,
      sales: sales[i],
      profit: profit[i],
    })),
  };
}

// Calculate YOY growth, put '-' if cannot calculate
function calculateYOY(dataRows) {
  const yoy = [];

  // Map quarter name -> index for easy lookup
  const quarterMap = {};
  dataRows.forEach((q, idx) => {
    quarterMap[q.quarter] = idx;
  });

  // Assume last 6 quarters for output
  const last6Quarters = dataRows.slice(-6);

  for (const q of last6Quarters) {
    const [month, yearStr] = q.quarter.split(' ');
    const prevYear = parseInt(yearStr) - 1;
    const prevQuarter = `${month} ${prevYear}`;

    const prevIdx = quarterMap[prevQuarter];

    if (prevIdx !== undefined) {
      const salesGrowth =
        ((q.sales - dataRows[prevIdx].sales) /
          Math.abs(dataRows[prevIdx].sales)) *
        100;
      const profitGrowth =
        ((q.profit - dataRows[prevIdx].profit) /
          Math.abs(dataRows[prevIdx].profit)) *
        100;
      yoy.push(
        parseFloat(salesGrowth.toFixed(2)),
        parseFloat(profitGrowth.toFixed(2))
      );
    } else {
      yoy.push('-', '-'); // Fill with "-" if previous year data not available
    }
  }

  // Ensure the array is always 12 elements (6 quarters × 2)
  while (yoy.length < 12) {
    yoy.unshift('-', '-'); // add at the start
  }

  return yoy;
}

// Extract text from PDF using pdf-lib
async function extractTextFromPDF(buffer) {
  const pdfDoc = await PDFDocument.load(buffer);
  const pages = pdfDoc.getPages();
  let fullText = '';
  for (const page of pages) {
    try {
      const content = await page.getTextContent?.(); // fallback if available
      if (content) {
        fullText += content.items?.map((i) => i.str).join(' ') + '\n';
      }
    } catch {
      // ignore errors
    }
  }
  return fullText || '';
}

// Get latest transcript text
async function getLatestTranscript(url) {
  const { data } = await axios.get(url, {
    headers: { 'User-Agent': 'Mozilla/5.0' },
  });
  const $ = cheerio.load(data);

  // Get the documents section
  const docSection = $('#documents ul.list-links li');

  // Find the first transcript link in the documents
  let transcriptLink = null;
  docSection.each((_, li) => {
    const link = $(li).find('a[title="Raw Transcript"]').attr('href');
    if (link) {
      transcriptLink = link;
      return false; // break
    }
  });

  if (!transcriptLink) return 'Transcript not available';

  const fullLink = transcriptLink.startsWith('http')
    ? transcriptLink
    : `https://www.screener.in${transcriptLink}`;

  console.log('Transcript link', fullLink);

  const pdfResp = await axios.get(fullLink, { responseType: 'arraybuffer' });
  const pdfBuffer = Buffer.from(pdfResp.data);
  const text = await extractTextFromPDF(pdfBuffer);

  return text.slice(0, 5000); // limit to 5000 chars
}

// Get AI insights
async function getAIInsights(symbol, about, transcriptText) {
  const prompt = `
You are a financial research assistant. Extract:
1. Sector (short)
2. Sub-sector (short)
3. Detailed summary of the conference call with focus on financial performance, margin guidance, how different business segments are doing, management guidance for the future along with numbers, and key risks.

Return JSON: { "sector": "", "subSector": "", "concallSummary": "", "guidance": "" }

Company: ${symbol}
About: ${about}
Transcript: ${transcriptText}
`;

  const response = await openai.chat.completions.create({
    model: 'gpt-4.1-mini',
    messages: [{ role: 'user', content: prompt }],
    temperature: 0,
  });

  try {
    return JSON.parse(response.choices[0].message.content);
  } catch {
    return {
      sector: 'N/A',
      subSector: 'N/A',
      concallSummary: 'N/A',
      guidance: 'N/A',
    };
  }
}

// Color formatting for YOY cells, '-' is white
function getColorRequests(startRow, startCol, yoyValues) {
  const requests = [];
  for (let i = 0; i < yoyValues.length; i++) {
    const value = yoyValues[i];
    const colIndex = startCol + i;
    let color;

    if (value === '-' || value === null) {
      color = { red: 1, green: 1, blue: 1 }; // white for missing
    } else if (value >= 10) {
      color = { red: 0, green: 0.8, blue: 0 }; // green
    } else {
      color = { red: 1, green: 0, blue: 0 }; // red
    }

    requests.push({
      repeatCell: {
        range: {
          sheetId: 0,
          startRowIndex: startRow,
          endRowIndex: startRow + 1,
          startColumnIndex: colIndex,
          endColumnIndex: colIndex + 1,
        },
        cell: { userEnteredFormat: { backgroundColor: color } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });
  }
  return requests;
}

// Write row to Google Sheet
async function writeToSheet(sheetId, sheetTab, rowData, yoyValues) {
  const sheets = await getSheetsClient();

  const appendRes = await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: `${sheetTab}!A:Z`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [rowData] },
  });

  const updatedRowIndex =
    appendRes.data.updates.updatedRange.match(/\d+$/)[0] - 1;
  const startCol = 3;
  const requests = getColorRequests(updatedRowIndex, startCol, yoyValues);

  if (requests.length) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: { requests },
    });
  }

  console.log(`✅ Row appended and colored to "${sheetTab}"`);
}

// Main
async function main() {
  const url = process.argv[2];
  if (!url) {
    console.error(
      'Usage: node yoy-to-sheets-latest-transcript-pdflib.js <screener_url>'
    );
    process.exit(1);
  }

  const SHEET_ID = process.env.SHEET_ID;
  const SHEET_TAB = process.env.SHEET_TAB;

  console.log(`🔍 Scraping quarterly data from: ${url}`);
  const scraped = await scrapeQuarterlyData(url);
  const yoyNumbers = calculateYOY(scraped.quarters);

  console.log('📄 Fetching latest transcript...');
  const transcriptText = await getLatestTranscript(url);

  console.log('🤖 Getting AI-powered company insights...');
  const insights = await getAIInsights(
    scraped.symbol,
    scraped.about,
    transcriptText
  );

  const row = [
    scraped.symbol,
    insights.sector,
    insights.subSector,
    ...yoyNumbers,
    insights.concallSummary,
    insights.guidance,
  ];

  console.log(`📊 Row: ${row.join(' | ')}`);
  await writeToSheet(SHEET_ID, SHEET_TAB, row, yoyNumbers);
}

main();
