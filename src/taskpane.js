// ─── State ───────────────────────────────────────────────────────────────
let currentSelectedText = "";
let currentAIResponse = "";
let generatedImageBase64 = "";
let chatHistory = [];  // [{role:'user'|'ai', text}]
let recognition = null;
let isRecording = false;

// ─── Study Materials ────────────────────────────────────────────────
function getStudyContext() {
    const ref = document.getElementById("refMaterials")?.value.trim();
    const cls = document.getElementById("classMaterials")?.value.trim();
    let ctx = "";
    if (ref) ctx += `\n\nחומרי עזר שסופקו על ידי הסטודנט (use to ground your response):\n${ref}`;
    if (cls) ctx += `\n\nחומר לימודי שנלמד בכיתה (העדף את התאוריות והמושגים שנלמדו לפי חומר זה):\n${cls}`;
    return ctx;
}

function updateMaterialsBadge() {
    const btn = document.getElementById("toggleMaterials");
    if (!btn) return;
    const ref = document.getElementById("refMaterials")?.value.trim();
    const cls = document.getElementById("classMaterials")?.value.trim();
    btn.classList.toggle("materials-active", !!(ref || cls));
}

function saveMaterials() {
    const ref = document.getElementById("refMaterials").value;
    const cls = document.getElementById("classMaterials").value;
    localStorage.setItem("studyRefMaterials", ref);
    localStorage.setItem("studyClassMaterials", cls);
    updateMaterialsBadge();
    document.getElementById("materialsStatus").innerText = "✅ החומרים נשמרו ויוספו להקשר כל ֽ-AI";
    setTimeout(() => {
        document.getElementById("materialsPanel").classList.add("hidden");
    }, 1500);
}

function clearMaterials() {
    document.getElementById("refMaterials").value = "";
    document.getElementById("classMaterials").value = "";
    localStorage.removeItem("studyRefMaterials");
    localStorage.removeItem("studyClassMaterials");
    updateMaterialsBadge();
    document.getElementById("materialsStatus").innerText = "החומרים נמחקו";
}

function loadMaterials() {
    const ref = localStorage.getItem("studyRefMaterials");
    const cls = localStorage.getItem("studyClassMaterials");
    if (ref) document.getElementById("refMaterials").value = ref;
    if (cls) document.getElementById("classMaterials").value = cls;
    updateMaterialsBadge();
    if (ref || cls) {
        document.getElementById("materialsStatus").innerText = "חומרים טעונים ופעילים";
    }
}

// ─── File Upload for Study Materials ─────────────────────────────────────
function setupFileUploads() {
    // Set PDF.js worker source once the lib is available
    if (window.pdfjsLib) {
        pdfjsLib.GlobalWorkerOptions.workerSrc =
            "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    }

    document.getElementById("refMaterialsFile").addEventListener("change", e => {
        handleFileUpload(e.target.files, "refMaterials", "refMaterialsFileStatus");
        e.target.value = "";
    });
    document.getElementById("classMaterialsFile").addEventListener("change", e => {
        handleFileUpload(e.target.files, "classMaterials", "classMaterialsFileStatus");
        e.target.value = "";
    });

    // Drag-and-drop on the textareas
    ["refMaterials", "classMaterials"].forEach(id => {
        const statusId = id + "FileStatus";
        const ta = document.getElementById(id);
        ta.addEventListener("dragover", e => { e.preventDefault(); ta.classList.add("drag-over"); });
        ta.addEventListener("dragleave", () => ta.classList.remove("drag-over"));
        ta.addEventListener("drop", e => {
            e.preventDefault();
            ta.classList.remove("drag-over");
            if (e.dataTransfer.files.length) {
                handleFileUpload(e.dataTransfer.files, id, statusId);
            }
        });
    });
}

async function handleFileUpload(files, targetId, statusId) {
    const statusEl = document.getElementById(statusId);
    const textarea = document.getElementById(targetId);
    const names = [];

    for (const file of Array.from(files)) {
        statusEl.innerText = `⏳ קורא ${file.name}...`;
        try {
            const text = await extractTextFromFile(file);
            const separator = textarea.value.trim() ? "\n\n" : "";
            textarea.value += `${separator}--- ${file.name} ---\n${text}`;
            names.push(file.name);
        } catch (err) {
            statusEl.innerText = `❌ ${file.name}: ${err.message}`;
            return;
        }
    }

    statusEl.innerText = `✅ ${names.join(", ")}`;
    updateMaterialsBadge();
}

async function extractTextFromFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "txt" || ext === "md") return readFileAsText(file);
    if (ext === "pdf") return extractPdfText(file);
    if (ext === "docx") return extractDocxText(file);
    throw new Error(`סוג קובץ לא נתמך: .${ext}. השתמש ב-TXT, PDF או DOCX.`);
}

function readFileAsText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = () => reject(new Error("שגיאה בקריאת הקובץ"));
        reader.readAsText(file, "UTF-8");
    });
}

async function extractPdfText(file) {
    if (!window.pdfjsLib) throw new Error("PDF.js לא נטען — רענן ונסה שוב");
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const pageTexts = [];
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        pageTexts.push(content.items.map(item => item.str).join(" "));
    }
    return pageTexts.join("\n").trim();
}

async function extractDocxText(file) {
    if (!window.mammoth) throw new Error("mammoth.js לא נטען — רענן ונסה שוב");
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer });
    return result.value.trim();
}

// ─── Prompt Templates ─────────────────────────────────────────────────────
const qaPrompts = {
    fix:        "תקן שגיאות כתיב, דקדוק וניסוח. שמור על זרימה טבעית ואנושית. החזר אך ורק את הטקסט המתוקן:\n\n",
    summarize:  "סכם את העיקר בצורה שזורה וטבעית (הימנע מרשימות בולטים אם אפשר). החזר אך ורק את הסיכום:\n\n",
    professional:"שכתב לסגנון עבודה אקדמית רצינית. החזר אך ורק את הטקסט. בסוף הטקסט הוסף תגית [הכנס ספירת מילים מדויקת].\n\n",
    translate:  "Translate to fluent natural English. Return only the translation:\n\n",
    find_errors:"מצא אי-דיוקים עובדתיים על בסיס ההקשר. החזר רשימה קצרה ופשוטה של הערות:\n\n",
    plant_errors:"שכתב את הטקסט ותהליך אותו ב-3 שלבים פנימיים: 1. יצירת טקסט אקדמי מושלם. 2. תיקוף עצמי מול דרישות המשתמש (וודא שאין פספוס מושגים). 3. הנמכת משלב סמנטית (לא תחבירית!) על כ-15% מהמשפטים. החלף מילים גבוהות במילים פשוטות ותקינות (למשל: 'המנצלת' -> 'שמשתמשת'). חוקים: לעולם אל תיגע במשפט הפותח והסוגר של כל פסקה. לעולם אל תחליף בביטויי סלנג נמוכים (כגון: 'בגדול', 'תכלס', 'כאילו', 'איך ש...'). לאחר הטקסט, הוסף '--- מה שתלתי:' והסבר קצר:\n\n",
    continue:   "המשך לכתוב את הטקסט בצורה שתתחבר באופן מושלם למה שנכתב לפני כן. שמור על אותן מילים אהובות ועל אותו מקצב. החזר רק את ההמשך:\n\n"
};

const toneMap = {
    student:       "TONE: High-potential undergraduate student. Cap SUBJECTIVITY: Use 'לדעתי' or 'להבנתי' exactly once. Rest of text is objective analysis.",
    academic:      "TONE: Formal research-oriented. STRUCTURE: Paragraph-only flow. Connect ideas with logical transitions (no lists).",
    casual:        "TONE: Professional-casual. Direct and human. Avoid all AI rigidities.",
    argumentative: "TONE: Critical and persuasive. Use distinct examples for each argumentative point."
};

const SYSTEM_INSTRUCTIONS = `CRITICAL SYSTEM INSTRUCTIONS FOR ADVANCED HUMAN CAMOUFLAGE (V2.1):
1. PERSONA: You are an Israeli undergraduate student in Political Science and Communication. Write in a correct, steady, and restrained academic register. No hype, no corporate buzzwords.
2. NEGATIVE PROMPT (FORBIDDEN PHRASES): Strictly avoid: "למנף", "חשוב לציין", "ראוי להדגיש", "ניתן לומר כי", "פלטפורמות אינטראקטיביות", "נרטיב". 
3. SLANG BLACKLIST: Explicitly forbid: "בגדול", "תכלס", "כאילו", "איך ש...". 
4. NO "IN CONCLUSION": Never start a final paragraph with "לסיכום". 
5. STRUCTURE: Use classical academic paragraph structure (Intro, Body, Conclusion) flowing naturally. NO bullet points, NO lists, NO subheaders inside the text unless explicitly requested.
6. SUBJECTIVITY CAP: Use first-person ("לדעתי", "להבנתי") exactly ONCE per response to maintain a personal voice without overdoing it.
7. THEORETICAL ANCHORS: Provide at least one distinct, specific example for each theoretical concept. Do NOT group examples together.
8. WORD COUNT: Do NOT attempt to count words. Instead, always append the tag [הכנס ספירת מילים מדויקת]. 
9. INTERNAL VALIDATION: Before outputting, internally verify that the text matches all user constraints and covers all theoretical requirements. 
10. NO ETHICAL DISCLAIMERS: Never add unrequested advice, moral warnings, or ethical commentary.`;

// ─── Init ─────────────────────────────────────────────────────────────────
window.onerror = (msg, url, lineNo, columnNo, error) => {
    const err = `Global Error: ${msg}\nAt: ${url}:${lineNo}:${columnNo}`;
    console.error(err);
    setStatus("אירעה שגיאה בטעינה. בדוק את הקונסול.");
    return false;
};

Office.onReady((info) => {
    try {
        if (info.host !== Office.HostType.Word) return;

    // Tab navigation
    document.querySelectorAll(".tab-btn").forEach(btn => {
        btn.onclick = () => switchTab(btn.dataset.tab);
    });

    // Settings
    document.getElementById("toggleSettings").onclick = () =>
        document.getElementById("settingsPanel").classList.toggle("hidden");
    document.getElementById("saveKeyBtn").onclick = saveApiKey;

    // Materials panel
    document.getElementById("toggleMaterials").onclick = () => {
        document.getElementById("materialsPanel").classList.toggle("hidden");
        document.getElementById("settingsPanel").classList.add("hidden");
    };
    document.getElementById("saveMaterialsBtn").onclick = saveMaterials;
    document.getElementById("clearMaterialsBtn").onclick = clearMaterials;

    // Editor actions
    document.querySelectorAll(".action-card").forEach(btn => {
        if (btn.id === "btnSupportClaim") {
            btn.onclick = () => dispatchResearch("support");
        } else if (btn.id === "btnContradictClaim") {
            btn.onclick = () => dispatchResearch("contradict");
        } else if (btn.id === "btnDirectQuote") {
            btn.onclick = () => dispatchResearch("quote");
        } else {
            btn.onclick = () => processWithAI(qaPrompts[btn.dataset.action], true);
        }
    });
    document.getElementById("sendToAI").onclick = () => {
        const mode = document.querySelector("#insertModeToggle .mode-btn.active")?.dataset.mode;
        if (mode === "sections") {
            writeBySection();
        } else {
            processWithAI();
        }
    };

    // Mode toggle
    document.querySelectorAll("#insertModeToggle .mode-btn").forEach(btn => {
        btn.onclick = () => {
            document.querySelectorAll("#insertModeToggle .mode-btn").forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            const isSections = btn.dataset.mode === "sections";
            document.getElementById("sectionGuidelinesArea").classList.toggle("hidden", !isSections);

            if (isSections) {
                // Copy prompt → sectionGuidelines if sections field is empty
                const promptVal = document.getElementById("prompt").value.trim();
                const secVal = document.getElementById("sectionGuidelines").value.trim();
                if (promptVal && !secVal) {
                    document.getElementById("sectionGuidelines").value = promptVal;
                    document.getElementById("prompt").value = "";
                }
                document.getElementById("prompt").placeholder = "הנחיה נוספת לכלל הסעיפים (אופציונלי)...";
            } else {
                // Copy sectionGuidelines → prompt if prompt is empty
                const secVal = document.getElementById("sectionGuidelines").value.trim();
                const promptVal = document.getElementById("prompt").value.trim();
                if (secVal && !promptVal) {
                    document.getElementById("prompt").value = secVal;
                    document.getElementById("sectionGuidelines").value = "";
                }
                document.getElementById("prompt").placeholder = "שאל או בקש כל דבר...";
            }
        };
    });
    document.getElementById("insertTracked").onclick = insertAsTrackedChange;
    document.getElementById("copyResponse").onclick = copyResponse;
    document.getElementById("voiceBtn").onclick = toggleVoice;

    // Chat
    document.getElementById("sendChat").onclick = sendChatMessage;
    document.getElementById("chatInput").onkeydown = e => {
        if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendChatMessage(); }
    };

    // Studio
    document.getElementById("generateImage").onclick = generateImage;
    document.getElementById("insertImage").onclick = insertImageIntoDoc;

    // Research & Biblio
    document.getElementById("btnScanBiblio").onclick = generateDocumentBibliography;
    document.getElementById("insertResearchTracked").onclick = insertResearchTrackedChange;

    // Lecturer
    document.getElementById("runLecturerReview").onclick = runLecturerReview;
    document.getElementById("btnGenerateCoverPage").onclick = generateCoverPage;
    document.getElementById("applyAllSuggestions").onclick = applyAllSuggestions;
    document.getElementById("copyReport").onclick = () => {
        const report = document.getElementById("lecturerReport");
        navigator.clipboard.writeText(report.innerText)
            .then(() => document.getElementById("lecturerStatus").innerText = "הועתק!")
            .catch(() => {});
    };

    loadApiKey();
    loadMaterials();
    setupFileUploads();
    setupSelectionListener();
    setupDirectInsertToggle();
    setStatus("התוסף מוכן! הקליק על הלשוניות למעלה למעבר בין העורך, הצ'אט, והמרצה 🎓");
    } catch (e) {
        console.error("Init error", e);
        setStatus("שגיאה באתחול: " + e.message);
    }
});

// ─── Tab Switch ───────────────────────────────────────────────────────────
function switchTab(tabId) {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(c => c.classList.add("hidden"));
    document.querySelector(`[data-tab="${tabId}"]`).classList.add("active");
    document.getElementById(`tab-${tabId}`).classList.remove("hidden");
    
    if (tabId === "research" && currentSelectedText) {
        document.getElementById("researchQuery").value = currentSelectedText;
    }
}

// ─── Status ───────────────────────────────────────────────────────────────
function setStatus(msg) { document.getElementById("status").innerText = msg; }

// ─── API Key ──────────────────────────────────────────────────────────────
function saveApiKey() {
    const geminiKey = document.getElementById("apiKey").value.trim();
    const perplexityKey = document.getElementById("perplexityApiKey").value.trim();
    
    if (geminiKey) {
        localStorage.setItem("geminiApiKey", geminiKey);
    } else {
        localStorage.removeItem("geminiApiKey");
    }

    if (perplexityKey) {
        localStorage.setItem("perplexityApiKey", perplexityKey);
    } else {
        localStorage.removeItem("perplexityApiKey");
    }

    document.getElementById("keyStatus").innerText = "✅ ההגדרות נשמרו!";
    setTimeout(() => document.getElementById("settingsPanel").classList.add("hidden"), 1500);
}

function loadApiKey() {
    const gk = localStorage.getItem("geminiApiKey");
    if (gk) document.getElementById("apiKey").value = gk;

    const pk = localStorage.getItem("perplexityApiKey");
    if (pk) document.getElementById("perplexityApiKey").value = pk;

    if (gk || pk) {
        document.getElementById("keyStatus").innerText = "מפתחות טעונים";
    }
    // Load direct-insert preference
    const directMode = localStorage.getItem("directInsertMode") === "true";
    document.getElementById("directInsertMode").checked = directMode;
}

function getDirectInsertMode() {
    const el = document.getElementById("directInsertMode");
    return el?.checked ?? false;
}

// Save direct-insert preference whenever it changes
function setupDirectInsertToggle() {
    document.getElementById("directInsertMode").addEventListener("change", (e) => {
        localStorage.setItem("directInsertMode", e.target.checked);
    });
}

function getApiKey() {
    return localStorage.getItem("geminiApiKey") || document.getElementById("apiKey").value.trim();
}

function getPerplexityApiKey() {
    return localStorage.getItem("perplexityApiKey") || document.getElementById("perplexityApiKey").value.trim();
}

// ─── Selection Listener ───────────────────────────────────────────────────
function setupSelectionListener() {
    Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onSelectionChanged
    );
    onSelectionChanged();
}

async function onSelectionChanged() {
    try {
        await Word.run(async (ctx) => {
            const range = ctx.document.getSelection();
            range.load("text");
            await ctx.sync();

            const display = document.getElementById("selectedTextDisplay");
            const badge = document.getElementById("selectionBadge");

            if (!range.text?.trim()) {
                currentSelectedText = "";
                display.innerText = "לא סומן טקסט (משתמש בכל המסמך)";
                display.classList.add("empty-state");
                badge.className = "badge badge-empty";
                badge.innerText = "ריק";
            } else {
                currentSelectedText = range.text;
                display.innerText = currentSelectedText;
                display.classList.remove("empty-state");
                badge.className = "badge badge-filled";
                badge.innerText = `${currentSelectedText.length} תווים`;
            }
        });
    } catch(e) { console.error("selection error", e); }
}

// ─── Voice Dictation ──────────────────────────────────────────────────────
function toggleVoice() {
    if (!('webkitSpeechRecognition' in window) && !('SpeechRecognition' in window)) {
        setStatus("הדפדפן לא תומך בהכתבה קולית.");
        return;
    }

    if (isRecording) {
        recognition?.stop();
        return;
    }

    const SRClass = window.SpeechRecognition || window.webkitSpeechRecognition;
    recognition = new SRClass();
    recognition.lang = "he-IL";
    recognition.continuous = false;
    recognition.interimResults = false;

    const btn = document.getElementById("voiceBtn");
    recognition.onstart = () => {
        isRecording = true;
        btn.classList.add("recording");
        setStatus("מקשיב...");
    };
    recognition.onend = () => {
        isRecording = false;
        btn.classList.remove("recording");
        setStatus("עיבוד הדיבור...");
    };
    recognition.onerror = (e) => {
        isRecording = false;
        btn.classList.remove("recording");
        setStatus("שגיאה בקליטת קול: " + e.error);
    };
    recognition.onresult = async (e) => {
        const rawSpeech = e.results[0][0].transcript;
        setStatus(`זוהה: "${rawSpeech.slice(0,40)}..."`);
        const voicePrompt = `המשתמש אמר את הדברים הבאים (טקסט גולמי). ארגן, ניסח ונקה אותם לפסקה קריאה ומדויקת ללא שינוי תוכן:\n\n${rawSpeech}`;
        await processWithAI(voicePrompt, true, true);
    };

    recognition.start();
}

// ─── Core AI Call ─────────────────────────────────────────────────────────
async function processWithAI(overridePrompt = null, isQuickAction = false, skipSelection = false) {
    const apiKey = getApiKey();
    if (!apiKey) {
        setStatus("אנא הזן מפתח API בהגדרות (⚙️)");
        document.getElementById("settingsPanel").classList.remove("hidden");
        return;
    }

    let promptBase = overridePrompt;
    if (!isQuickAction) {
        promptBase = document.getElementById("prompt").value.trim();
        if (!promptBase) { setStatus("אנא הזן הנחיה."); return; }
        promptBase = `${promptBase}\n\n`;
    }

    // Tone override
    const tone = document.getElementById("toneSelector").value;
    const toneInstruction = tone ? `\n${toneMap[tone]}` : "";

    setStatus("קורא מסמך ומחשב...");
    setEditorBusy(true);

    try {
        const docText = await getDocumentText();

        const fullPrompt = `${SYSTEM_INSTRUCTIONS}${toneInstruction}${getStudyContext()}\n\nDocument Context (Whole Document):\n${docText}\n\n${
            currentSelectedText && !skipSelection ? `Selected Text Focus:\n${currentSelectedText}\n\n` : ""
        }User Request:\n${promptBase}`;

        const aiText = await callGeminiText(apiKey, fullPrompt);

        currentAIResponse = aiText;

        if (getDirectInsertMode()) {
            // Auto-insert as tracked change directly, skip the response panel
            try {
                await insertTextAsTracked(aiText);
                showToast("✅ הוכנס ך-Track Change");
            } catch(insertErr) {
                console.error("Auto-insert failed", insertErr);
                // Fall back to showing the panel
                document.getElementById("editorResponse").innerText = aiText;
                document.getElementById("editorResponseContainer").classList.remove("hidden");
            }
        } else {
            document.getElementById("editorResponse").innerText = aiText;
            document.getElementById("editorResponseContainer").classList.remove("hidden");
        }
        setStatus("הושלם!");
    } catch(e) {
        console.error(e);
        setStatus("שגיאה: " + e.message);
        document.getElementById("editorResponse").innerText = "אירעה שגיאה: " + e.message;
        document.getElementById("editorResponseContainer").classList.remove("hidden");
    } finally {
        setEditorBusy(false);
    }
}

function setEditorBusy(busy) {
    const btn = document.getElementById("sendToAI");
    btn.disabled = busy;
    btn.innerText = busy ? "עובד..." : "שלח ל-AI";
    document.querySelectorAll(".action-card").forEach(b => b.disabled = busy);
    document.querySelectorAll("#insertModeToggle .mode-btn").forEach(b => b.disabled = busy);
}

async function getDocumentText() {
    let text = "";
    await Word.run(async (ctx) => {
        const body = ctx.document.body;
        body.load("text");
        await ctx.sync();
        text = body.text;
    });
    return text;
}

async function callGeminiText(apiKey, prompt) {
    const res = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`,
        {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
        }
    );
    if (!res.ok) throw new Error(`API ${res.status}`);
    const data = await res.json();
    return data.candidates[0].content.parts[0].text;
}

// ─── Insert as Track Change ───────────────────────────────────────────────
async function insertTextAsTracked(text) {
    await Word.run(async (ctx) => {
        ctx.document.load("changeTrackingMode");
        await ctx.sync();
        const orig = ctx.document.changeTrackingMode;
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await ctx.sync();
        ctx.document.getSelection().insertText(text, Word.InsertLocation.after);
        await ctx.sync();
        ctx.document.changeTrackingMode = orig;
        await ctx.sync();
    });
}

async function insertAsTrackedChange() {
    if (!currentAIResponse) return;
    try {
        await insertTextAsTracked(currentAIResponse);
        setStatus("✅ הוכנס כהצעת עריכה (Track Change)!");
    } catch(e) { setStatus("שגיאה: " + e.message); }
}

function showToast(message) {
    const toast = document.getElementById("insertToast");
    toast.innerText = message;
    toast.classList.remove("hidden");
    setTimeout(() => toast.classList.add("hidden"), 2500);
}

function copyResponse() {
    if (!currentAIResponse) return;
    navigator.clipboard.writeText(currentAIResponse)
        .then(() => setStatus("הועתק ללוח!"))
        .catch(() => setStatus("שגיאה בהעתקה."));
}

// ─── Research Dispatch (routes quick vs deep) ────────────────────────────
function dispatchResearch(type) {
    const depth = document.querySelector("input[name='researchDepth']:checked")?.value ?? "quick";
    if (depth === "deep") {
        performDeepResearch(type);
    } else {
        processContextualResearch(type);
    }
}

// ─── Deep Research: Layer 1 (Perplexity) + Layer 2 (Gemini Synthesis) ───
async function performDeepResearch(type) {
    const perplexityKey = getPerplexityApiKey();
    const geminiKey = getApiKey();

    if (!perplexityKey || !geminiKey) {
        setStatus("נדרשים שני מפתחות API (Perplexity + Gemini) למחקר עמוק.");
        document.getElementById("settingsPanel").classList.remove("hidden");
        return;
    }

    const claimText = document.getElementById("researchQuery").value.trim() || currentSelectedText;
    if (!claimText) {
        setStatus("אנא סמן טקסט או הקלד את הנושא.");
        return;
    }

    const researchBtns = ["btnSupportClaim", "btnContradictClaim", "btnDirectQuote"];
    researchBtns.forEach(id => { document.getElementById(id).disabled = true; });
    document.getElementById("researchResponseContainer").classList.add("hidden");

    const progressEl = document.getElementById("deepResearchProgress");
    const step1El = document.getElementById("dstep1");
    const step2El = document.getElementById("dstep2");
    progressEl.classList.remove("hidden");
    [step1El, step2El].forEach(el => {
        el.className = "deep-step";
        el.querySelector(".step-status").innerText = "–";
    });

    function setStepState(el, state, status) {
        el.className = "deep-step " + state;
        el.querySelector(".step-status").innerText = status;
    }

    try {
        // ─── Layer 1: Perplexity – find sources ──────────────────────────
        setStepState(step1El, "active", "⏳");
        setStatus("Layer 1: מחפש מקורות אקדמיים (Perplexity)...");

        const layer1Prompts = {
            support:    "You are an academic research assistant. Find 3-5 high-quality scholarly sources that SUPPORT the given claim. For each: author(s), year, title, journal/publisher, one key argument, and a direct brief quote (max 2 sentences). Format as a numbered list.",
            contradict: "You are an academic research assistant. Find 3-5 high-quality scholarly sources that CONTRADICT or critically challenge the given claim. For each: author(s), year, title, journal/publisher, key counter-argument, and a direct brief quote (max 2 sentences). Format as a numbered list.",
            quote:      "You are an academic research assistant. Find 2-3 highly relevant DIRECT QUOTES from published academic sources on this topic. For each: the exact quote, author, year, title, and full APA citation. Be precise."
        };

        const perplexityRes = await fetch("https://api.perplexity.ai/chat/completions", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${perplexityKey}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                model: "sonar-pro",
                messages: [
                    { role: "system", content: layer1Prompts[type] },
                    { role: "user", content: `הנושא לחקר: ${claimText}` }
                ]
            })
        });

        if (!perplexityRes.ok) {
            const errData = await perplexityRes.json();
            throw new Error(errData?.error?.message || `Perplexity API ${perplexityRes.status}`);
        }

        const perplexityData = await perplexityRes.json();
        const sourcesText = perplexityData.choices[0].message.content.trim();
        setStepState(step1El, "done", "✅");

        // ─── Layer 2: Gemini – synthesize into academic paragraph ────────
        setStepState(step2El, "active", "⏳");
        setStatus("Layer 2: מסנתז לפסקה אקדמית (Gemini)...");

        const directionMap = {
            support:    "כתוב פסקה אקדמית שתומכת בטענה ומשלבת את המקורות",
            contradict: "כתוב פסקה אקדמית ביקורתית שמתמודדת עם הטענה בעזרת המקורות",
            quote:      "שלב את הציטוטים בפסקה אקדמית קוהרנטית וזורמת"
        };

        const synthesisPrompt = `${SYSTEM_INSTRUCTIONS}

You are an expert academic writing assistant. You have two inputs:
1. A claim/text from a student (in Hebrew)
2. A set of academic sources found via research

Your task: ${directionMap[type]}.

SYNTHESIS RULES:
- Write in fluent, natural academic Hebrew
- Each source must be distinctly referenced (Author, Year) woven naturally into the text — never list them as bullets
- End with a full APA reference list under the exact heading: '--- מקורות ---'
- Do not open with 'לסיכום' or any clichéd AI opener

Student's claim:
${claimText}

Sources found (Layer 1 – Perplexity):
${sourcesText}`;

        const synthesized = await callGeminiText(geminiKey, synthesisPrompt);
        setStepState(step2El, "done", "✅");

        currentAIResponse = synthesized;
        document.getElementById("researchResponseText").innerText = synthesized;
        document.getElementById("researchResponseContainer").classList.remove("hidden");
        setStatus("✅ מחקר עמוק הושלם — Perplexity + Gemini.");
        setTimeout(() => progressEl.classList.add("hidden"), 3500);

    } catch (e) {
        console.error(e);
        setStatus("שגיאה במחקר עמוק: " + e.message);
        [step1El, step2El].forEach(el => {
            if (el.classList.contains("active")) setStepState(el, "", "❌");
        });
    } finally {
        researchBtns.forEach(id => { document.getElementById(id).disabled = false; });
    }
}

// ─── Perplexity: Contextual & Document Research ───────────────────────────
async function processContextualResearch(type) {
    const apiKey = getPerplexityApiKey();
    if (!apiKey) {
        setStatus("אנא הזן מפתח Perplexity API בהגדרות (⚙️)");
        document.getElementById("settingsPanel").classList.remove("hidden");
        return;
    }

    const claimText = document.getElementById("researchQuery").value.trim() || currentSelectedText;
    if (!claimText) {
        setStatus("אנא סמן טקסט בעורך או הקלד בשדה הנושא.");
        return;
    }

    let systemPrompt = "";
    if (type === "support") {
        systemPrompt = "You are an expert academic assistant. The user provides a claim or text. Find strong, reliable academic sources that SUPPORT this claim. Write a fluent, academic paragraph in Hebrew that integrates these sources to back up the claim. At the very end of your response, you MUST provide the full formal citations in APA format under the exact heading: '--- מקורות ---'.";
    } else if (type === "contradict") {
        systemPrompt = "You are an expert academic assistant. The user provides a claim or text. Find strong, reliable academic sources that CONTRADICT or critically challenge this claim (offer a counter-argument). Write a fluent, academic paragraph in Hebrew that brings these contradicting sources to challenge the text. At the very end of your response, you MUST provide the full formal citations in APA format under the exact heading: '--- מקורות ---'.";
    } else if (type === "quote") {
        systemPrompt = "You are an expert academic assistant. The user provides a topic or claim. Find ONE specific, highly relevant DIRECT QUOTE from a published academic source. In your response: 1. Introduce the scholar and year briefly in Hebrew. 2. Provide the exact quote in quotation marks (translate to Hebrew if it's in English, or bring the English quote with a Hebrew translation). 3. Explain briefly why it's powerful. 4. At the very end, provide the full formal citation in APA format under the exact heading: '--- מקורות ---'.";
    }
    
    document.getElementById("btnSupportClaim").disabled = true;
    document.getElementById("btnContradictClaim").disabled = true;
    document.getElementById("btnDirectQuote").disabled = true;
    
    let pendingStatus = "מחפש סימוכין...";
    if (type === "support") pendingStatus = "מחפש חיזוקים לטענה במאגרי אקדמיה...";
    if (type === "contradict") pendingStatus = "מחפש סתירות וביקורות פנטנציאליות...";
    if (type === "quote") pendingStatus = "מחפש ציטוט ישיר חזק בספרות...";
    setStatus(pendingStatus);

    document.getElementById("researchResponseContainer").classList.add("hidden");

    try {
        const res = await fetch("https://api.perplexity.ai/chat/completions", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${apiKey}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                model: "sonar-pro",
                messages: [
                    { role: "system", content: systemPrompt },
                    { role: "user", content: `הנושא לחקר: ${claimText}` }
                ]
            })
        });

        if (!res.ok) {
            const errorData = await res.json();
            throw new Error(errorData?.error?.message || `API Error ${res.status}`);
        }

        const data = await res.json();
        const responseText = data.choices[0].message.content.trim();

        currentAIResponse = responseText;
        document.getElementById("researchResponseText").innerText = responseText;
        document.getElementById("researchResponseContainer").classList.remove("hidden");
        
        setStatus("✅ התקבלה תוצאת מחקר מאת Perplexity.");

    } catch (e) {
        console.error(e);
        setStatus("שגיאה בחיפוש ממוקד: " + e.message);
        document.getElementById("researchResponseText").innerText = "שגיאה מול Perplexity: " + e.message;
        document.getElementById("researchResponseContainer").classList.remove("hidden");
    } finally {
        document.getElementById("btnSupportClaim").disabled = false;
        document.getElementById("btnContradictClaim").disabled = false;
        document.getElementById("btnDirectQuote").disabled = false;
    }
}

async function insertResearchTrackedChange() {
    if (!currentAIResponse) return;
    try {
        await insertTextAsTracked("\\n" + currentAIResponse + "\\n");
        setStatus("✅ התוצאה הוכנסה למסמך כ-Track Change!");
    } catch(e) { 
        setStatus("שגיאה בהכנסה: " + e.message); 
    }
}

async function generateDocumentBibliography() {
    const apiKey = getApiKey(); // Use Gemini since it handles very large docs easily
    if (!apiKey) {
        setStatus("אנא הזן מפתח Gemini API בהגדרות לסריקת כל המסמך (⚙️)");
        document.getElementById("settingsPanel").classList.remove("hidden");
        return;
    }

    const btn = document.getElementById("btnScanBiblio");
    btn.disabled = true;
    btn.innerText = "קורא מסמך ומעבד מקורות...";
    setStatus("סורק את כל המסמך... (פעולה זו עשויה לארוך חצי דקה)");

    try {
        const docText = await getDocumentText();
        if (!docText || docText.length < 20) {
            throw new Error("המסמך ריק או קצר מדי לסריקה.");
        }

        const prompt = "You are an expert academic librarian.\n" +
            "I will provide you with a long academic document. Your task is to carefully scan the ENTIRE document from start to finish and identify every single in-text citation, scholar name mentioned, or empirical reference (e.g., \"(Smith, 2020)\", \"Cohen argued that...\", etc.).\n" +
            "Using your internet intelligence and knowledge base, DO YOUR BEST to identify the actual published source for every single mention.\n" +
            "Finally, generate ONLY a complete, unified Bibliography list of ALL the sources cited or mentioned in the text.\n" +
            "The bibliography MUST be strictly formatted in APA 7th Edition style and ALWAYS sorted alphabetically.\n\n" +
            "DO NOT write introductions. DO NOT write conclusions. DO NOT hallucinate sources not mentioned in the text.\n" +
            "RETURN ONLY THE BIBLIOGRAPHY TEXT itself.\n\n" +
            "Here is the document text to scan:\n" +
            "-------------\n" + docText;

        setStatus("בונה ביבליוגרפיה (GEMINI AI)...");
        const aiText = await callGeminiText(apiKey, prompt);

        // Insert at the very end of the document
        await Word.run(async (ctx) => {
            const body = ctx.document.body;
            const textToInsert = "\\n\\n--- רשימת מקורות (ביבליוגרפיה אוטומטית) ---\\n\\n" + aiText + "\\n";
            const insertedRange = body.insertParagraph(textToInsert, Word.InsertLocation.end);
            insertedRange.paragraphFormat.readingOrder = Word.ReadingOrder.leftToRight;
            insertedRange.font.name = "Times New Roman";
            await ctx.sync();
            setStatus("✅ הביבליוגרפיה צורפה בהצלחה לסוף המסמך!");
            showToast("✅ ביבליוגרפיה הופקה");
        });

    } catch (e) {
        console.error(e);
        setStatus("שגיאה בסריקת ביבליוגרפיה: " + e.message);
    } finally {
        btn.disabled = false;
        btn.innerText = "📑 סרוק מסמך וצור ביבליוגרפיה (APA)";
    }
}

// ─── Chat ─────────────────────────────────────────────────────────────────
async function sendChatMessage() {
    const input = document.getElementById("chatInput");
    const msg = input.value.trim();
    if (!msg) return;
    input.value = "";

    addChatBubble(msg, "user");
    chatHistory.push({ role: "user", text: msg });

    const apiKey = getApiKey();
    if (!apiKey) { addChatBubble("אנא הגדר מפתח API בהגדרות.", "ai"); return; }

    const loadingBubble = addChatBubble("חושב...", "ai loading");

    try {
        const docText = await getDocumentText();
        const historyBlock = chatHistory.slice(-8)
            .map(m => `[${m.role === "user" ? "User" : "Assistant"}]: ${m.text}`)
            .join("\n");

        const prompt = `${SYSTEM_INSTRUCTIONS}${getStudyContext()}

You are an AI assistant helping with a Word document. You are in CHAT mode - give conversational, helpful answers. You MAY use short natural replies (but no long preambles).

Document Context:
${docText}

Conversation History:
${historyBlock}

Respond only to the last user message:`;

        const aiText = await callGeminiText(apiKey, prompt);
        chatHistory.push({ role: "ai", text: aiText });
        loadingBubble.classList.remove("loading");
        loadingBubble.innerText = aiText;

        // Add "insert to doc" button
        const insertBtn = document.createElement("button");
        insertBtn.className = "chat-insert-btn";
        insertBtn.innerText = "📄 הכנס למסמך";
        insertBtn.onclick = () => insertTextAfterSelection(aiText);
        loadingBubble.parentElement.appendChild(insertBtn);

    } catch(e) {
        loadingBubble.innerText = "שגיאה: " + e.message;
        loadingBubble.classList.remove("loading");
    }
}

function addChatBubble(text, type) {
    const msgs = document.getElementById("chatMessages");
    const div = document.createElement("div");
    div.className = `chat-bubble ${type}`;
    div.innerText = text;
    msgs.appendChild(div);
    msgs.scrollTop = msgs.scrollHeight;
    return div;
}

async function insertTextAfterSelection(text) {
    try {
        await Word.run(async (ctx) => {
            ctx.document.getSelection().insertText(text, Word.InsertLocation.after);
            await ctx.sync();
            setStatus("📄 הטקסט הוכנס למסמך.");
        });
    } catch(e) { setStatus("שגיאה: " + e.message); }
}

// ─── Studio: Image Generation ─────────────────────────────────────────────
async function generateImage() {
    const apiKey = getApiKey();
    if (!apiKey) {
        document.getElementById("studioStatus").innerText = "הגדר מפתח API קודם.";
        return;
    }

    const imagePromptText = document.getElementById("imagePrompt").value.trim();
    if (!imagePromptText) {
        document.getElementById("studioStatus").innerText = "אנא הזן תיאור לתמונה.";
        return;
    }

    const btn = document.getElementById("generateImage");
    btn.disabled = true;
    btn.innerText = "יוצר תמונה...";
    document.getElementById("studioStatus").innerText = "שולח בקשה ל-Imagen...";
    document.getElementById("studioResponseContainer").classList.add("hidden");

    try {
        // Imagen API uses a different endpoint & format
        const res = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/imagen-3.0-generate-002:predict?key=${apiKey}`,
            {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    instances: [{ prompt: imagePromptText }],
                    parameters: { sampleCount: 1 }
                })
            }
        );

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err?.error?.message || `HTTP ${res.status}`);
        }

        const data = await res.json();
        const b64 = data.predictions?.[0]?.bytesBase64Encoded;
        if (!b64) throw new Error("לא התקבלה תמונה מה-API.");

        generatedImageBase64 = b64;

        const preview = document.getElementById("imagePreview");
        preview.innerHTML = `<img src="data:image/png;base64,${b64}" alt="תמונה שנוצרה" />`;
        document.getElementById("studioResponseContainer").classList.remove("hidden");
        document.getElementById("studioStatus").innerText = "✅ התמונה נוצרה!";

    } catch(e) {
        document.getElementById("studioStatus").innerText = "שגיאה: " + e.message;
    } finally {
        btn.disabled = false;
        btn.innerText = "🎨 צור תמונה";
    }
}

async function insertImageIntoDoc() {
    if (!generatedImageBase64) return;
    try {
        await Word.run(async (ctx) => {
            const range = ctx.document.getSelection();
            range.insertInlinePictureFromBase64(generatedImageBase64, Word.InsertLocation.after);
            await ctx.sync();
            setStatus("✅ התמונה הוכנסה למסמך!");
            document.getElementById("studioStatus").innerText = "✅ התמונה הוכנסה למסמך!";
        });
    } catch(e) {
        setStatus("שגיאה בהכנסת התמונה: " + e.message);
    }
}

// ─── Lecturer Mode ────────────────────────────────────────────────────────
let lecturerSuggestions = [];

const strictnessPrompts = {
    standard: "You are acting as a standard academic reviewer. Flag missing elements, logical gaps, and factual errors compared to the guidelines. Be balanced.",
    strict:   "You are a strict academic examiner. Flag EVERY deviation from the guidelines, no matter how small.",
    mentor:   "You are a supportive mentor reviewer. Give constructive feedback, suggest improvements, and point out what's done well alongside what needs work."
};

async function generateCoverPage() {
    const apiKey = getApiKey();
    if (!apiKey) {
        document.getElementById("lecturerStatus").innerText = "הגדר מפתח API קודם.";
        return;
    }

    const guidelines = document.getElementById("lecturerGuidelines").value.trim();
    if (!guidelines) {
        document.getElementById("lecturerStatus").innerText = "אנא הזן את הנחיות המרצה קודם לכן.";
        return;
    }

    const btn = document.getElementById("btnGenerateCoverPage");
    btn.disabled = true;
    btn.innerText = "מייצר דף שער...";
    document.getElementById("lecturerStatus").innerText = "מייצר מבנה לדף שער...";

    try {
        const prompt = `You are a strict Israeli academic assistant helping a student.
Based on the following lecturer guidelines, generate ONLY the TEXT of a formal academic Cover Page (דף שער) in Hebrew. 

Guidelines:
${guidelines}

DO NOT include markdown formatting like "**". Only use plain text. Do not add introductions or conclusions.
If details like Course Name, Student Name, ID, or Date are implicitly required but missing from the guidelines, add standard placeholder strings in brackets like [שם הסטודנט], [תעודת זהות], [תאריך מבוקש].
Center the text conceptually (it will be centered programmatically). Build a standard formal layout (e.g. University Name at top, Course Name, Assignment Title, Submitted to, Submitted by, Date).`;

        const coverPageText = await callGeminiText(apiKey, prompt);

        // Insert at the VERY BEGINNING of the document as a Track Change
        await Word.run(async (ctx) => {
            ctx.document.load("changeTrackingMode");
            await ctx.sync();
            const origTracking = ctx.document.changeTrackingMode;
            ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await ctx.sync();

            const body = ctx.document.body;
            // Add a clean page break after the cover page text.
            const fullText = coverPageText + "\n\n--- מעבר עמוד בדף שער באחריותך --- \n\n";
            const range = body.insertText(fullText, Word.InsertLocation.start);
            range.font.size = 14;
            range.font.name = "Ariel"; // Typical Israeli academic default, or stick to existing
            range.paragraphFormat.alignment = "Centered";

            await ctx.sync();
            
            ctx.document.changeTrackingMode = origTracking;
            await ctx.sync();

            setStatus("✅ דף שער הוקפץ לתחילת המסמך כהצעת עריכה!");
            document.getElementById("lecturerStatus").innerText = "✅ דף השער נוסף בהצלחה!";
            showToast("✅ דף שער נוצר");
        });

    } catch(e) {
        console.error(e);
        document.getElementById("lecturerStatus").innerText = "שגיאה ביצירת דף שער: " + e.message;
    } finally {
        btn.disabled = false;
        btn.innerHTML = `<span class="action-icon">📄</span><span class="action-label">צור דף שער למסמך (Track Change)</span>`;
    }
}

async function runLecturerReview() {
    const apiKey = getApiKey();
    if (!apiKey) { document.getElementById("lecturerStatus").innerText = "הגדר מפתח API קודם."; return; }

    const guidelines = document.getElementById("lecturerGuidelines").value.trim();
    if (!guidelines) { document.getElementById("lecturerStatus").innerText = "אנא הדבק הנחיות קודם."; return; }

    const strictness = document.getElementById("lecturerStrictness").value;
    const btn = document.getElementById("runLecturerReview");
    btn.disabled = true;
    btn.innerText = "בודק...";
    document.getElementById("lecturerStatus").innerText = "קורא מסמך ומנתח מול ההנחיות...";
    document.getElementById("lecturerResultContainer").classList.add("hidden");

    try {
        const docText = await getDocumentText();
        const prompt = `${strictnessPrompts[strictness]}${getStudyContext()}

You are reviewing a student's document against assignment guidelines.

ASSIGNMENT GUIDELINES:
${guidelines}

STUDENT DOCUMENT:
${docText}

Return ONLY a valid JSON array. Each item must have:
- "original": exact verbatim phrase from the document (max 80 chars)
- "replacement": corrected version
- "reason": short Hebrew explanation why (1-2 sentences)

If no issues, return []. No markdown, no intro text.`;

        const raw = await callGeminiText(apiKey, prompt);
        
        // Robust JSON Extraction
        let jsonStr = raw.trim();
        const start = jsonStr.indexOf("[");
        const end = jsonStr.lastIndexOf("]");
        if (start === -1 || end === -1) throw new Error("לא התקבלה תגובת JSON תקינה מה-AI. נסה שוב.");
        jsonStr = jsonStr.substring(start, end + 1);

        // Sanitize: 
        // 1. Remove trailing commas before ] or }
        jsonStr = jsonStr.replace(/,\s*([\]}])/g, "$1");
        // 2. Handle potential unescaped control characters in JSON strings (newlines)
        // jsonStr = jsonStr.replace(/\n/g, "\\n").replace(/\r/g, "\\r"); // Careful here, don't break the structure

        try {
            lecturerSuggestions = JSON.parse(jsonStr);
        } catch (parseError) {
            console.warn("Retrying with newline cleanup...", parseError);
            // Fallback for common AI malformation: unescaped newlines inside quote values
            const fixedJson = jsonStr.replace(/(?<=:\s*")([\s\S]*?)(?="\s*[,}\]])/g, (match) => {
                return match.replace(/\n/g, " ").replace(/\r/g, " ");
            });
            lecturerSuggestions = JSON.parse(fixedJson);
        }
        const reportEl = document.getElementById("lecturerReport");
        reportEl.innerHTML = "";

        if (lecturerSuggestions.length === 0) {
            reportEl.innerHTML = `<div style="color:var(--success);font-weight:600;">✅ לא נמצאו בעיות! העבודה עומדת בכל הדרישות.</div>`;
        } else {
            lecturerSuggestions.forEach((s, i) => {
                const block = document.createElement("div");
                block.className = "suggestion-block";
                block.innerHTML = `
                    <div class="suggestion-title">💡 הצעה ${i + 1}</div>
                    <div class="suggestion-original">${escapeHtml(s.original)}</div>
                    <div class="suggestion-new">→ ${escapeHtml(s.replacement)}</div>
                    <div class="suggestion-reason">${escapeHtml(s.reason)}</div>`;
                reportEl.appendChild(block);
            });
        }

        document.getElementById("lecturerResultContainer").classList.remove("hidden");
        document.getElementById("lecturerStatus").innerText = `✅ נמצאו ${lecturerSuggestions.length} הצעות.`;
    } catch(e) {
        console.error(e);
        document.getElementById("lecturerStatus").innerText = "שגיאה: " + e.message;
    } finally {
        btn.disabled = false;
        btn.innerText = "🎓 בדוק את העבודה";
    }
}

async function applyAllSuggestions() {
    if (!lecturerSuggestions.length) return;
    const btn = document.getElementById("applyAllSuggestions");
    btn.disabled = true;
    btn.innerText = "מחיל שינויים...";

    try {
        await Word.run(async (ctx) => {
            ctx.document.load("changeTrackingMode");
            await ctx.sync();
            const orig = ctx.document.changeTrackingMode;
            ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await ctx.sync();

            for (const s of lecturerSuggestions) {
                if (!s.original || !s.replacement) continue;
                const results = ctx.document.body.search(s.original, { matchCase: false, matchWholeWord: false });
                results.load("items");
                await ctx.sync();
                if (results.items.length > 0) {
                    results.items[0].insertText(s.replacement, Word.InsertLocation.replace);
                    await ctx.sync();
                }
            }

            ctx.document.changeTrackingMode = orig;
            await ctx.sync();
        });
        document.getElementById("lecturerStatus").innerText = `✅ ${lecturerSuggestions.length} הצעות הוחלו כ-Track Changes!`;
    } catch(e) {
        document.getElementById("lecturerStatus").innerText = "שגיאה: " + e.message;
    } finally {
        btn.disabled = false;
        btn.innerText = "✅ החל הצעות כ-Track Changes";
    }
}

function escapeHtml(str) {
    return (str || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ─── XML Section Parser (for writeBySection) ──────────────────────────────
function parseSectionsXml(raw) {
    const sections = [];
    // Find the outer <sections>...</sections> block (tolerant of whitespace/extra text)
    const outerMatch = raw.match(/<sections[\s\S]*?>([\s\S]*?)<\/sections>/i);
    const inner = outerMatch ? outerMatch[1] : raw;

    // Match every <section>...</section> block
    const sectionRegex = /<section[\s\S]*?>([\s\S]*?)<\/section>/gi;
    let match;
    while ((match = sectionRegex.exec(inner)) !== null) {
        const block = match[1];
        const headerMatch = block.match(/<header>([\s\S]*?)<\/header>/i);
        const answerMatch = block.match(/<answer>([\s\S]*?)<\/answer>/i);
        if (headerMatch && answerMatch) {
            sections.push({
                header: headerMatch[1].trim(),
                answer: answerMatch[1].trim()
            });
        }
    }
    if (!sections.length) throw new Error("לא נמצאו תגיות <section> בתשובת ה-AI.");
    return sections;
}

// ─── Robust JSON repair + parse ───────────────────────────────────────────
function repairAndParseJSON(rawStr) {
    let str = rawStr.trim();
    // Strip markdown code fences
    str = str.replace(/^```[\w]*\n?/, "").replace(/\n?```$/, "");
    const start = str.indexOf("[");
    const end = str.lastIndexOf("]");
    if (start === -1 || end === -1) throw new Error("לא נמצא מבנה JSON בתשובת ה-AI.");
    str = str.substring(start, end + 1);
    str = str.replace(/,\s*([\]}])/g, "$1");

    // First attempt: direct parse
    try { return JSON.parse(str); } catch {}

    // Second attempt: character-by-character repair of unescaped control chars inside strings
    let fixed = "";
    let inString = false;
    let escaped = false;
    for (let i = 0; i < str.length; i++) {
        const ch = str[i];
        if (escaped) { fixed += ch; escaped = false; continue; }
        if (ch === "\\") { escaped = true; fixed += ch; continue; }
        if (ch === '"') { inString = !inString; fixed += ch; continue; }
        if (inString) {
            if (ch === "\n") { fixed += "\\n"; continue; }
            if (ch === "\r") { fixed += "\\r"; continue; }
            if (ch === "\t") { fixed += "\\t"; continue; }
        }
        fixed += ch;
    }
    return JSON.parse(fixed);
}

// ─── Write By Section ─────────────────────────────────────────────────────
async function writeBySection() {
    const apiKey = getApiKey();
    if (!apiKey) {
        setStatus("הגדר מפתח Gemini API קודם (⚙️).");
        document.getElementById("settingsPanel").classList.remove("hidden");
        return;
    }

    const guidelines = document.getElementById("sectionGuidelines").value.trim()
                    || document.getElementById("prompt").value.trim();
    if (!guidelines) {
        setStatus("אנא הדבק את הנחיות/שאלות המרצה בתיבת הסעיפים.");
        document.getElementById("sectionGuidelinesArea").classList.remove("hidden");
        document.getElementById("sectionGuidelines").focus();
        return;
    }

    const sendBtn = document.getElementById("sendToAI");
    sendBtn.disabled = true;
    sendBtn.innerText = "מזהה סעיפים...";

    const progressEl = document.getElementById("sectionWriteProgress");
    const stepsEl = document.getElementById("sectionWriteSteps");
    progressEl.classList.remove("hidden");
    stepsEl.innerHTML = "";

    function addStep(label) {
        const div = document.createElement("div");
        div.className = "swp-step active";
        div.innerHTML = `<span class="swp-step-icon">⏳</span><span>${label}</span>`;
        stepsEl.appendChild(div);
        return div;
    }
    function finishStep(el, success = true) {
        el.className = "swp-step " + (success ? "done" : "error");
        el.querySelector(".swp-step-icon").innerText = success ? "✅" : "❌";
    }

    try {
        const s1 = addStep("קורא את המסמך...");
        const docText = await getDocumentText();
        if (!docText || docText.length < 10) throw new Error("המסמך ריק מדי.");
        finishStep(s1);

        const s2 = addStep("מזהה סעיפים ומנסח תוכן (Gemini)...");
        setStatus("שולח ל-Gemini...");

        const prompt = `${SYSTEM_INSTRUCTIONS}${getStudyContext()}

You are an expert academic writing assistant for an Israeli university student.
The student has pasted assignment questions/sections below. Write a complete, high-quality academic Hebrew answer for EVERY section.

═══ OUTPUT FORMAT — FOLLOW EXACTLY ═══
Use this exact XML structure. One <section> block per question/section. No text outside the tags.

<sections>
<section>
<header>EXACT verbatim text of the question/section title — copy it word-for-word, no truncation</header>
<answer>Full academic answer in Hebrew. Separate paragraphs with a blank line. NO bullet points. NO sub-headers. 2–4 flowing paragraphs per section.</answer>
</section>
<section>
<header>...</header>
<answer>...</answer>
</section>
</sections>

═══ CONTENT RULES ═══
- Each section is answered independently. NEVER merge or blend sections.
- No global intro or conclusion that spans multiple sections.
- Apply all style rules: no forbidden phrases, no first-person overuse, no "לסיכום" openers.
- If study materials were provided above, anchor answers in those concepts and theories.

Document context (tone/topic continuity — do NOT answer from this, only from the guidelines):
${docText.substring(0, 2000)}

═══ ASSIGNMENT SECTIONS TO ANSWER ═══
${guidelines}`;

        const raw = await callGeminiText(apiKey, prompt);
        finishStep(s2);

        const s3 = addStep("מפרש תוצאות...");
        let sections;
        try {
            sections = parseSectionsXml(raw);
        } catch (parseErr) {
            throw new Error("ה-AI החזיר מבנה לא-תקין, נסה שוב: " + parseErr.message);
        }
        finishStep(s3);

        if (!sections.length) {
            setStatus("לא זוהו סעיפים. ודא שההנחיות כוללות שאלות/כותרות ברורות.");
            progressEl.classList.add("hidden");
            return;
        }

        const s4 = addStep(`כותב ${sections.length} סעיפים למסמך (שאלה + תשובה)...`);
        setStatus(`מכניס ${sections.length} סעיפים כ-Track Change...`);

        await Word.run(async (ctx) => {
            ctx.document.load("changeTrackingMode");
            await ctx.sync();
            const orig = ctx.document.changeTrackingMode;
            ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await ctx.sync();

            // Anchor: insert right after the current cursor position
            let insertAfter = ctx.document.getSelection();

            for (const sec of sections) {
                // ── Question header (bold), chained after previous insert ──
                const headerPara = insertAfter.insertParagraph(sec.header, Word.InsertLocation.after);
                headerPara.font.bold = true;
                headerPara.font.size = 13;
                headerPara.spaceAfter = 4;

                let lastPara = headerPara;

                // ── Answer paragraphs (normal) ─────────────────────────
                const paragraphs = sec.answer
                    .split(/\n+/)
                    .map(p => p.trim())
                    .filter(p => p.length > 0);

                for (const para of paragraphs) {
                    const p = lastPara.insertParagraph(para, Word.InsertLocation.after);
                    p.font.bold = false;
                    p.font.size = 12;
                    p.spaceAfter = 6;
                    lastPara = p;
                }

                // ── Blank separator; becomes anchor for next section ────
                const blank = lastPara.insertParagraph("", Word.InsertLocation.after);
                await ctx.sync();
                insertAfter = blank;
            }

            ctx.document.changeTrackingMode = orig;
            await ctx.sync();
        });

        finishStep(s4);
        setStatus(`✅ ${sections.length} סעיפים נכתבו בסדר הנכון!`);
        showToast(`✅ ${sections.length} סעיפים מולאו`);

    } catch (e) {
        console.error(e);
        setStatus("שגיאה: " + e.message);
        // Mark all active steps as error
        document.querySelectorAll(".swp-step.active").forEach(el => finishStep(el, false));
    } finally {
        sendBtn.disabled = false;
        sendBtn.innerText = "שלח ל-AI";
        setTimeout(() => progressEl.classList.add("hidden"), 6000);
    }
}
