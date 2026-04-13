# Word AI Assistant

A modern, custom Microsoft Word Add-in that integrates the **Google Gemini (2.5) AI API** directly into your Word workflow. This tool completely eliminates the need to copy-paste text between Word and external AI chat interfaces.

![Word AI Assistant Interface](assets/screenshot.png) *(Add a screenshot here later)*

## ✨ Features
* **Select & Read**: Instantly pulls highlighted text from your document into the assistant.
* **Quick Actions**: One-click buttons for common editing tasks:
  * 🪄 **Fix Grammar:** Corrects spelling, grammar, and phrasing.
  * 📝 **Summarize:** Creates a concise summary of the selected text.
  * 👔 **Professional:** Rewrites text to sound polished and business-appropriate.
  * 🌍 **Translate to English:** Accurately translates Hebrew (or any language) to natural English.
* **Contextual Academic Research (📚):** A dedicated Research tab powered by Perplexity AI for seamless academic writing:
  * **Auto-Context:** Switching to the Research tab automatically pulls in whatever text you highlighted in the editor.
  * 👍 **Support Claim (`תמוך בטענה`):** Automatically finds scholarly sources that support your claim and writes an integrated academic paragraph.
  * 👎 **Contradict Claim (`סתר טענה`):** Automatically finds contrasting views and writes a critical counter-argument paragraph.
  * 💬 **Direct Quote (`מצא ציטוט ישיר`):** Finds a compelling direct quote from published literature that fits your context, complete with citation.
  * **Document Scanner Bibliography (`סריקת מסמך וביבליוגרפיה`):** Powered by Gemini's massive context window. Press one button and the AI will read your **entire Word document**, locate every single in-text citation (e.g. "Smith, 2021") you wrote manually, track down the full source on the web, and generate a complete, alphabetically sorted APA bibliography at the end of your document!
* **Custom Prompts**: Ask the AI to do anything with the selected text exactly as you would in ChatGPT or Gemini.
* **Insert at Cursor**: Seamlessly insert the AI's response back into your Word document exactly where your cursor is.
* **Secure Key Storage**: Your personal Gemini API key is securely saved locally in the Add-in's `localStorage` (via the ⚙️ Settings menu) instead of being hardcoded and exposed in the source files.

## 🚀 Prerequisites
- Node.js (v20+)
- Microsoft Word (Desktop/Mac or Word Online)
- A [Google Gemini AI API Key](https://aistudio.google.com/app/apikey)

## 🛠️ Installation & Setup

1. **Clone or Download the repository.**
2. **Install dependencies:**
   ```bash
   npm install
   ```
3. **Start the local development server:**
   ```bash
   npm run dev
   ```
   *This starts the Vite server on `https://localhost:3000`.*

---

## 📖 How to Run & Test in Word

### Option 1: Automagical Desktop Sideloading (Windows/Mac)
Open a **new terminal acting as Administrator**, navigate to this folder, and run:
   ```bash
   npx office-addin-debugging start manifest.xml desktop
   ```
*Note: Running as Administrator is required the very first time on Windows to allow the `loopback exemption` so Word can talk to your `localhost` server.*

### Option 2: Word Online (Browser)
1. Keep the `npm run dev` server running.
2. Go to [Word Online](https://www.office.com/launch/word) and create a new blank document.
3. Click **Insert > Add-ins > Manage My Add-ins > Upload My Add-in**.
4. Select the `manifest.xml` file from this project folder.
5. The "Word AI Assistant" will appear in your Home ribbon!

## 🔐 Managing the API Key
For security reasons, do **not** hardcode your API key into the HTML or Javascript.
1. Open the Add-in inside Word.
2. Click the **Settings (⚙️)** icon in the top right.
3. Paste your Gemini API Key and click **Save**. 
4. The key is now safely stored on your device and will be remembered for future sessions.

## ☁️ Study Materials From Storage Or Project Folder
The add-in now works best with local project materials, with an optional one-way sync from Firebase Storage into the project folder.

Local project folder:
1. Put your files under `public/project-materials/`.
2. List them in `public/project-materials/index.json`.
3. The add-in will load them locally and inject selected files into AI context.

Sync from Firebase Storage into the local folder:
1. In Storage, keep an `index.json` file at the bucket root or provide a direct URL to any remote index file.
2. Run:
    ```bash
    npm run sync:storage -- gs://my-study-b-b.firebasestorage.app
    ```
3. The script downloads all indexed files into `public/project-materials/` and rewrites the local `index.json`.
4. After that, run the add-in normally with `npm run dev`.

Expected `index.json` format:
```json
[
   {
      "id": "article-1",
      "title": "Political Communication Article",
      "file": "materials/political-communication.pdf",
      "type": "pdf"
   }
]
```

Supported extraction: `pdf`, `docx`, and text-based files such as `txt`, `md`, `csv`, `json`, `html`.

## 🏗️ Built With
- **Vanilla JS & HTML5**
- **Vite** - Lightning fast dev server & bundler.
- **Office.js** - Microsoft's official API to interact with Word documents.

---
*Created for personal workflow automation and productivity.*
