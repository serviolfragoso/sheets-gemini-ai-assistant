# Google Sheets & Gemini AI Content Assistant (Zero-Backend Architecture)

## 📌 Project Overview
A lightweight, high-performance automation tool built using **Google Apps Script** that natively integrates the **Google Gemini API** (`gemini-3.6-flash`) directly into Google Sheets. It transforms static spreadsheets into dynamic, intelligent interfaces, enabling users to process tabular data through structured system instructions and dynamic prompts at scale—all without the need for external servers or complex backends.

## 💡 The Business Problem & Solution
* **The Problem:** Academic and corporate workflows often involve repetitive, manual content generation and structuring directly inside spreadsheets, creating massive operational friction.
* **The Solution:** Developed an end-to-end automated client-side pipeline that fetches topics from active spreadsheet ranges, routes them through a controlled LLM pipeline with specialized system prompts, and seamlessly writes back formatted pedagogical or technical insights into designated columns.

## 🛠️ Tech Stack
* **Language:** JavaScript (ES6+ / Google Apps Script)
* **APIs & Ecosystem:** Google Sheets API (`SpreadsheetApp`, `UrlFetchApp`), Google Generative AI API (`Gemini`).
* **Architecture:** Serverless / *Zero-Backend* Client-Side Integration.

## ⚙️ Key Technical Highlights
* **Resilient API Handling:** Implemented an automated retry mechanism with a 5-attempt fallback loop, intelligent delay buffers (`Utilities.sleep`), and rate-limit management to gracefully handle HTTP 429 (Quota Exceeded) errors.
* **Dynamic Range Processing:** Features a custom UI menu to process either active single rows or batch-process complex highlighted ranges dynamically.
* **Separation of Concerns:** Clean architectural design decoupling the core HTTP communication layer (`ejecutarConsultaGemini`) from specialized prompt generation functions and execution triggers (`onOpen`).

## 🚀 Setup & Installation
1. Open your target Google Sheet, navigate to **Extensions > Apps Script**.
2. Paste the contents of `code.js` into the editor.
3. Configure your Gemini API key in the Apps Script project settings:
   * Go to **Project Settings (Gear icon) > Script Properties**.
   * Add a property with the key `GEMINI_API_KEY` and your actual API token as the value.
4. Refresh your Google Sheet to load the custom **`✨ IA Planeación`** menu and start processing.