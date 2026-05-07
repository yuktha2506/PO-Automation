<div align="center">
<img width="1200" height="475" alt="PO Automation Banner" src="https://github.com/user-attachments/assets/0aa67016-6eaf-458a-adb2-6e31a0763ed6" />
</div>

# PO Automation System

An AI-powered **Purchase Order (PO) Automation** application that streamlines PO processing by extracting, validating, and generating structured outputs such as **Excel reports** and related documents. The system is designed to reduce manual effort, improve accuracy, and accelerate downstream finance and QA workflows.

---

##  Features

*  Automated PO data extraction
*  Excel generation for PO summaries and reports
*  AI-assisted processing using Gemini API
*  Validation and structuring of PO information
*  Fast local development with modern Node.js tooling

---

##  Tech Stack

* **Frontend / App Framework:** Node.js
* **AI Model:** Gemini (via Gemini API)
* **File Handling:** Excel generation & processing
* **Environment Management:** dotenv

---


##  Run Locally

### Prerequisites

* Node.js (v18 or higher recommended)
* npm
* Gemini API Key

### Steps

1. **Clone the repository**

   ```bash
   git clone <your-repo-url>
   cd <project-folder>
   ```

2. **Install dependencies**

   ```bash
   npm install
   ```

3. **Configure environment variables**

   Create a `.env.local` file in the root directory and add:

   ```env
   GEMINI_API_KEY=your_gemini_api_key_here
   ```

4. **Run the application**

   ```bash
   npm run dev
   ```

5. Open your browser and access the app on the local development URL.

---

##  Output

* Generated **Excel files** containing structured PO data

---

##  Use Cases

* Automating Purchase Order documentation
* Reducing manual Excel entry for PO data

---

##  Security Notes

* Do not commit `.env.local` or API keys to version control
* Use test or dummy data during development

---

##  Contribution

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Open a Pull Request

---

 Built to simplify PO workflows with automation and AI.
