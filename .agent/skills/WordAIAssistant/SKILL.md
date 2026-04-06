---
name: WordAIAssistant
description: Essential guidelines and architectural context for the Word AI Assistant project to ensure precise edits, strict framework adherence, and token conservation.
---

# Word AI Assistant - Skill Context & Runbook

This document provides strict boundaries and architectural context for the "Word AI Assistant" project. By reading this, you are explicitly forbidden from initiating exploratory, token-wasting searches regarding framework structure or project setup.

## 1. Architecture & Tech Stack
* **Nature:** Microsoft Word Add-in (Taskpane).
* **Core Technologies:** Vanilla JavaScript (ES6+), Vanilla CSS3, HTML5, Office.js.
* **Build Tool:** Vite (Development Server & Bundler).
* **Constraints:** 
  * Do **NOT** introduce React, Vue, Svelte, or any JS framework.
  * Do **NOT** introduce Tailwind, Bootstrap, or any CSS framework. Stick entirely to pure Vanilla CSS.
  * All logic revolves around `src/taskpane.js`, styling in `src/taskpane.css`, and structure in `index.html`.

## 2. File Organization
* `manifest.xml`: The Add-in configuration file (Ribbon definitions, localhost pointers). Edit this only if modifying ribbon commands or permission scopes.
* `index.html`: Contains the UI layout (HTML structure only). 
* `src/taskpane.js`: The central powerhouse. Contains UI logic, API calls to Google Gemini, and all Office.js (`Word.run`) text-manipulation logic.
* `src/taskpane.css`: Central stylesheet. Maintain a dynamic, premium, modern look as per the user's aesthetic requirements.
* `vite.config.js`: Configured to run on `https://localhost:3000`.

## 3. Strict Operating Directives (Token Conservation)
1. **Targeted Edits:** Do not rewrite the entire file or use indiscriminate `sed` commands. Only use `multi_replace_file_content` or `replace_file_content` directly on the affected lines.
2. **Office.js Rules:** Every interaction with the Word document canvas MUST be wrapped in:
   ```javascript
   await Word.run(async (context) => {
       // logic here
       await context.sync();
   });
   ```
3. **API Keys:** The Google Gemini API key is managed via the UI's Settings menu and saved to the user's local `localStorage`. Never insert hard-coded dummy keys manually into the script or suggest the user paste it directly into JS.
4. **Development Flow:** If the user asks you to implement a feature, do not waste tokens analyzing multiple files—jump straight to `index.html` to add the UI elements and `src/taskpane.js` to implement the logic/bind event listeners.

## 4. Feature Set Context
* **Quick Actions:** Fix Grammar, Summarize, Rewrite Professional, Translate.
* **Academic Research (Perplexity style):** Features designed to pull highlighted text, find support/contradictions in academia, and provide scholarly quotes.
* **Bibliography Generation:** The app scans the whole document for manual citations and generates an integrated APA bibliography at the end of the document utilizing Gemini's context window.

## 5. Execution Pre-check
* Before writing any function, scan `src/taskpane.js` briefly for existing helper functions to avoid redundancy.
* Stop and ask for clarification only if the requirement drastically shifts the Vanilla JS environment or strictly conflicts with `manifest.xml`. 
