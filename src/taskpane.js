// ─── State ───────────────────────────────────────────────────────────────
let currentSelectedText = "";
let chatHistory = [];
let chatSessions = [];
let currentChatSessionId = "";
let personalStyle = createEmptyPersonalStyle();
let isChatSendInFlight = false;

// ─── Constants ───────────────────────────────────────────────────────────
const STYLE_LS_KEY = "personalStyleData";
const PAST_WORKS_CACHE_KEY = "pastWorksLearned";
const CHAT_HISTORY_LS_KEY = "chatSessionsData";
const AGENT_RULES_KEY = "agentWritingRules";
const DEBUG_MODE_KEY = "debugModeEnabled";
const ASSIGNMENT_GUIDELINES_KEY = "assignmentGuidelinesText";
const REVIEW_DEPTH_KEY = "reviewDepthLevel";
const ASSIGNMENT_GUIDELINES_FILE_NAME_KEY = "assignmentGuidelinesFileName";
const CLOUD_SELECTED_IDS_KEY = "cloudSelectedMaterialIds";
const CLOUD_SELECTED_FOLDER_KEYS_KEY = "cloudSelectedFolderKeys";
const PROJECT_MATERIALS_INDEX_URL = "/project-materials/index.json";
const MIN_LEARNED_VOCAB_COUNT = 2;
const MIN_LEARNED_PHRASE_COUNT = 2;
const MAX_COUNT_PER_FILE = 6;
const MAX_CHAT_SESSIONS = 40;
const COMMON_ACADEMIC_MARKERS = [
    "בנוסף", "עם זאת", "לכן", "למעשה", "למשל", "בהתאם לכך", "לעומת זאת",
    "מנגד", "בפרט", "כתוצאה מכך", "כלומר", "בהמשך לכך"
];
let resetAllArmedUntil = 0;
let debugEvents = [];
const RESEARCH_INTENT = /מקורות אקדמיים|ציטוט|ציטוטים|ביבליוגרפי|ביבליוגרפיה|חפש מאמר|מקורות ל|מצא מקורות|מצא ציטוט|APA|\bdoi\b|\bscholar\b|\bcitation\b|\bbibliography\b|\breferences\b|\bjournal article\b/i;
const REVIEW_INTENT = /בדיקת מרצה|בדוק עבודה|בדוק את העבודה|תן ציון לעבודה|בדוק לפי רובריקה|הערך את העבודה/i;

// ─── מיני-סוכנים ─────────────────────────────────────────────────────────
const AGENTS_CONFIG = {
    "fix":      { label: "✨ תיקון",        placeholder: "מה לתקן? (כתב, ניסוח, בהירות…)",    route: "gemini",     systemCtx: `== AGENT: PROOFREADER ==\nYou are a Hebrew proofreader. Fix grammar, spelling, phrasing and style only. Return corrected text without explanations.\n== END AGENT ==` },
    "humanize": { label: "🧬 האנשה",        placeholder: "איזה קטע להאניש?",                   route: "gemini",     systemCtx: `== AGENT: HUMANIZER ==\nYou are an aggressive Hebrew text humanizer. Break symmetry, shorten sentences over 25 words, switch passive to active, replace worn connectors. Return only the humanized text.\n== END AGENT ==` },
    "sources":  { label: "📚 מקורות",       placeholder: "על מה לחפש מקורות?",                 route: "perplexity", systemCtx: "" },
    "lecturer": { label: "🧑‍🏫 בדיקת מרצה", placeholder: "שאל את המרצה או שלח הנחיות נוספות…", route: "gemini",     systemCtx: "" },
    "continue": { label: "✍️ המשך",         placeholder: "מאיפה להמשיך לכתוב?",               route: "gemini",     systemCtx: `== AGENT: WRITER ==\nContinue writing from where the user stopped. Match their rhythm, vocabulary and academic level exactly. Write only the continuation — no meta comments.\n== END AGENT ==` },
    "summary":  { label: "📋 סיכום",        placeholder: "מה לסכם?",                           route: "gemini",     systemCtx: `== AGENT: SUMMARIZER ==\nSummarize the provided text into one flowing Hebrew paragraph — no bullet points, no numbered lists.\n== END AGENT ==` },
    "academic": { label: "👔 אקדמי",        placeholder: "מה לשכתב לסגנון אקדמי?",            route: "gemini",     systemCtx: `== AGENT: ACADEMIC WRITER ==\nRewrite the text in formal Hebrew academic style appropriate for a university assignment. Return only the rewritten text.\n== END AGENT ==` },
};

// מצב הסוכן הפעיל הנוכחי
let activeAgentId = null;

let assignmentGuidelinesFromFile = "";
let cloudMaterials = [];
let cloudSelectedMaterialIds = new Set();
let cloudSelectedFolderKeys = new Set();
let cloudMaterialTextCache = new Map();
let cloudMaterialLoadState = new Map();

// ─── System Instructions ──────────────────────────────────────────────────
const SYSTEM_INSTRUCTIONS = `CRITICAL SYSTEM INSTRUCTIONS FOR ADVANCED HUMAN CAMOUFLAGE (V3.0):

== PERSONA ==
You are a sharp, intellectually engaged Israeli undergraduate in Political Science and Communication. Your register is correct and restrained — serious without being stiff, confident without being arrogant. You think clearly and write precisely.

== NEGATIVE PROMPT — FORBIDDEN PHRASES (HARD BLOCK) ==
NEVER use any of the following: "למנף", "חשוב לציין", "ראוי להדגיש", "ניתן לומר כי", "פלטפורמות אינטראקטיביות", "נרטיב", "מורכב", "רב-ממדי", "ניואנסים", "תקופה מרתקת", "בסופו של דבר", "בהחלט", "יש לציין", "כמו כן", "בנוסף לכך", "מעבר לכך", "על כן ניתן לסכם".
NEVER open a paragraph with: "ברור כי", "אין ספק ש", "יש לזכור".
NEVER open a FINAL paragraph with "לסיכום" or "לסיכום דברים".

== SLANG BLACKLIST ==
Never use: "בגדול", "תכלס", "כאילו", "איך ש...", "בעצם".

== RHYTHM & BURSTINESS ==
Vary sentence length deliberately. After a long sentence (15+ words), follow with a short one (3-8 words). Mix compound clauses with single-punch statements.

== STRUCTURE ==
Classical paragraph flow (claim → evidence → analysis → link). NO bullet points, NO numbered lists, NO subheaders unless explicitly requested.

== SUBJECTIVITY CAP ==
Use first-person ("לדעתי", "להבנתי") EXACTLY ONCE per full response.

== THEORETICAL ANCHORS ==
Every theoretical concept must be illustrated with its own distinct, concrete historical or empirical example.

== NO ETHICAL DISCLAIMERS ==
Never add warnings or moral commentary the user did not ask for.`;

// ─── Personal Style System ────────────────────────────────────────────────
function createEmptyPersonalStyle() {
    return {
        schemaVersion: 3,
        manualVocabulary: [],
        manualPhrases: [],
        learnedVocabulary: [],
        learnedPhrases: [],
        learnedVocabularyCounts: {},
        learnedPhraseCounts: {},
        protectedVocabulary: [],
        protectedPhrases: [],
        examples: [],
        notes: "",
        learnedNotes: [],
        academic_level: "undergraduate",
        last_updated: ""
    };
}

function uniqueStrings(items) {
    return [...new Set((items || []).map(item => String(item || "").trim()).filter(Boolean))];
}

function uniqueStringsByNormalized(items) {
    const seen = new Set();
    const result = [];
    for (const item of items || []) {
        const text = String(item || "").trim();
        if (!text) continue;
        const normalized = normalizeForCompare(text);
        if (seen.has(normalized)) continue;
        seen.add(normalized);
        result.push(text);
    }
    return result;
}

function hasNormalizedItem(items, target) {
    const normalizedTarget = normalizeForCompare(target);
    return (items || []).some(item => normalizeForCompare(item) === normalizedTarget);
}

function getAggregatedCountEntries(countMap = {}) {
    const aggregate = new Map();
    for (const [rawText, rawCount] of Object.entries(countMap || {})) {
        const text = String(rawText || "").trim();
        if (!text) continue;
        const normalized = normalizeForCompare(text);
        const count = Number(rawCount) || 0;
        if (!aggregate.has(normalized)) {
            aggregate.set(normalized, { text, normalized, count: 0 });
        }
        const entry = aggregate.get(normalized);
        entry.count += count;
        if (text.length < entry.text.length) entry.text = text;
    }
    return [...aggregate.values()];
}

function mergeCountMaps(base = {}, updates = {}) {
    const merged = { ...base };
    for (const [key, value] of Object.entries(updates || {})) {
        const numericValue = Number(value) || 0;
        merged[key] = Math.max(Number(merged[key]) || 0, numericValue);
    }
    return merged;
}

function mergeExamples(base = [], updates = []) {
    const merged = [...(base || []), ...(updates || [])];
    const seen = new Set();
    return merged.filter(example => {
        const normalized = String(example || "").trim();
        if (!normalized) return false;
        const key = normalized.slice(0, 60);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    }).slice(-5);
}

function normalizePersonalStyle(raw = {}) {
    const normalized = createEmptyPersonalStyle();
    const legacyVocabulary = uniqueStrings(raw.vocabulary ?? []);
    const legacyPhrases = uniqueStrings(raw.phrases ?? []);
    const hasExplicitManualFields = Array.isArray(raw.manualVocabulary) || Array.isArray(raw.manualPhrases);

    normalized.manualVocabulary = uniqueStringsByNormalized(raw.manualVocabulary ?? []);
    normalized.manualPhrases = uniqueStringsByNormalized(raw.manualPhrases ?? []);
    normalized.learnedVocabulary = uniqueStringsByNormalized([...(raw.learnedVocabulary ?? []), ...legacyVocabulary]);
    normalized.learnedPhrases = uniqueStringsByNormalized([...(raw.learnedPhrases ?? []), ...legacyPhrases]);
    normalized.learnedVocabularyCounts = mergeCountMaps({}, raw.learnedVocabularyCounts);
    normalized.learnedPhraseCounts = mergeCountMaps({}, raw.learnedPhraseCounts);
    normalized.protectedVocabulary = uniqueStringsByNormalized(raw.protectedVocabulary ?? []);
    normalized.protectedPhrases = uniqueStringsByNormalized(raw.protectedPhrases ?? []);

    // One-time repair: older versions may have auto-learned items persisted as manual.
    // If schemaVersion is missing and no explicit manual fields existed originally,
    // move manual entries to learned so relearn/reset can clean them.
    if (!raw.schemaVersion && !hasExplicitManualFields) {
        normalized.learnedVocabulary = uniqueStringsByNormalized([...(normalized.learnedVocabulary || []), ...(normalized.manualVocabulary || [])]);
        normalized.learnedPhrases = uniqueStringsByNormalized([...(normalized.learnedPhrases || []), ...(normalized.manualPhrases || [])]);
        normalized.manualVocabulary = [];
        normalized.manualPhrases = [];
    }

    if (normalized.learnedVocabulary.length) {
        normalized.learnedVocabulary.forEach(text => {
            const normalizedKey = normalizeForCompare(text);
            const hasMatchingCount = Object.keys(normalized.learnedVocabularyCounts).some(key => normalizeForCompare(key) === normalizedKey);
            if (hasMatchingCount) return;
            normalized.learnedVocabularyCounts[text] = MIN_LEARNED_VOCAB_COUNT;
        });
    }
    if (normalized.learnedPhrases.length) {
        normalized.learnedPhrases.forEach(text => {
            const normalizedKey = normalizeForCompare(text);
            const hasMatchingCount = Object.keys(normalized.learnedPhraseCounts).some(key => normalizeForCompare(key) === normalizedKey);
            if (hasMatchingCount) return;
            normalized.learnedPhraseCounts[text] = MIN_LEARNED_PHRASE_COUNT;
        });
    }
    normalized.examples = mergeExamples([], raw.examples ?? []);
    normalized.notes = String(raw.notes ?? "").trim();
    normalized.learnedNotes = uniqueStrings(raw.learnedNotes ?? []);
    normalized.academic_level = String(raw.academic_level ?? "undergraduate");
    normalized.last_updated = String(raw.last_updated ?? "");
    normalized.schemaVersion = 3;
    return normalized;
}

function mergePersonalStyles(baseStyle, updatesStyle) {
    const base = normalizePersonalStyle(baseStyle);
    const updates = normalizePersonalStyle(updatesStyle);
    const manualVocabulary = uniqueStringsByNormalized([...base.manualVocabulary, ...updates.manualVocabulary]);
    const manualPhrases = uniqueStringsByNormalized([...base.manualPhrases, ...updates.manualPhrases]);
    const learnedVocabulary = uniqueStringsByNormalized([...base.learnedVocabulary, ...updates.learnedVocabulary])
        .filter(text => !hasNormalizedItem(manualVocabulary, text));
    const learnedPhrases = uniqueStringsByNormalized([...base.learnedPhrases, ...updates.learnedPhrases])
        .filter(text => !hasNormalizedItem(manualPhrases, text));
    const protectedVocabulary = uniqueStringsByNormalized([...base.protectedVocabulary, ...updates.protectedVocabulary])
        .filter(text => !hasNormalizedItem(manualVocabulary, text));
    const protectedPhrases = uniqueStringsByNormalized([...base.protectedPhrases, ...updates.protectedPhrases])
        .filter(text => !hasNormalizedItem(manualPhrases, text));
    return {
        ...createEmptyPersonalStyle(),
        manualVocabulary,
        manualPhrases,
        learnedVocabulary,
        learnedPhrases,
        learnedVocabularyCounts: mergeCountMaps(base.learnedVocabularyCounts, updates.learnedVocabularyCounts),
        learnedPhraseCounts: mergeCountMaps(base.learnedPhraseCounts, updates.learnedPhraseCounts),
        protectedVocabulary,
        protectedPhrases,
        examples: mergeExamples(base.examples, updates.examples),
        notes: updates.notes || base.notes,
        learnedNotes: uniqueStrings([...base.learnedNotes, ...updates.learnedNotes]).slice(-5),
        academic_level: updates.academic_level || base.academic_level || "undergraduate",
        last_updated: updates.last_updated || base.last_updated
    };
}

function getCombinedVocabulary() {
    const manualSet = new Set((personalStyle.manualVocabulary || []).map(normalizeForCompare));
    return uniqueStrings([
        ...(personalStyle.manualVocabulary || []),
        ...(personalStyle.learnedVocabulary || []).filter(text => !manualSet.has(normalizeForCompare(text)))
    ]);
}

function getCombinedPhrases() {
    const manualSet = new Set((personalStyle.manualPhrases || []).map(normalizeForCompare));
    return uniqueStrings([
        ...(personalStyle.manualPhrases || []),
        ...(personalStyle.learnedPhrases || []).filter(text => !manualSet.has(normalizeForCompare(text)))
    ]);
}

function getVisibleLearnedVocabulary() {
    const manualSet = new Set((personalStyle.manualVocabulary || []).map(normalizeForCompare));
    return (personalStyle.learnedVocabulary || []).filter(text => !manualSet.has(normalizeForCompare(text)));
}

function getVisibleLearnedPhrases() {
    const manualSet = new Set((personalStyle.manualPhrases || []).map(normalizeForCompare));
    return (personalStyle.learnedPhrases || []).filter(text => !manualSet.has(normalizeForCompare(text)));
}

function getStyleChipItems() {
    const items = [];
    const seen = new Set();
    for (const text of personalStyle.manualVocabulary || []) {
        if (seen.has(text)) continue;
        seen.add(text);
        items.push({ text, type: "vocab", source: "manual" });
    }
    for (const text of getVisibleLearnedVocabulary()) {
        if (seen.has(text)) continue;
        seen.add(text);
        items.push({ text, type: "vocab", source: "learned" });
    }
    for (const text of personalStyle.manualPhrases || []) {
        if (seen.has(text)) continue;
        seen.add(text);
        items.push({ text, type: "phrase", source: "manual" });
    }
    for (const text of getVisibleLearnedPhrases()) {
        if (seen.has(text)) continue;
        seen.add(text);
        items.push({ text, type: "phrase", source: "learned" });
    }
    return items;
}

// ─── Chat History Persistence ────────────────────────────────────────────
function createChatSession(initialTitle = "שיחה חדשה") {
    return {
        id: `chat-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        title: initialTitle,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        messages: []
    };
}

function getWelcomeMessage() {
    const userName = localStorage.getItem("userName") || "משתמש"; // הוספת שם משתמש אם קיים
    return `שלום ${userName}! אני קורא את המסמך שלך ומשלב את הסגנון האישי שלך בכל תשובה 🎯\n\nמה אני יודע לעשות:\n✍️ לכתוב, לנסח ולשכתב — בסגנון שלך\n✨ לתקן ולשפר טקסט מסומן\n📚 למצוא מקורות וציטוטים אקדמיים (Perplexity)\n🧬 לבצע האנשה אגרסיבית לטקסט\n\n פשוט שאל, או לחץ על אחד הקיצורים למטה.`;
}

function deriveChatTitle(messages = []) {
    const firstUser = (messages || []).find(message => message.role === "user" && String(message.text || "").trim());
    if (!firstUser) return "שיחה חדשה";
    const title = String(firstUser.text || "").replace(/\s+/g, " ").trim();
    return title.length > 42 ? `${title.slice(0, 42)}...` : title;
}

function saveChatSessions() {
    try {
        localStorage.setItem(CHAT_HISTORY_LS_KEY, JSON.stringify({
            currentChatSessionId,
            sessions: chatSessions
        }));
    } catch (e) {
        console.warn("Could not save chat sessions:", e);
    }
}

function loadChatSessions() {
    try {
        const raw = localStorage.getItem(CHAT_HISTORY_LS_KEY);
        if (raw) {
            const parsed = JSON.parse(raw);
            const sessions = Array.isArray(parsed?.sessions) ? parsed.sessions : [];
            chatSessions = sessions.map(session => ({
                id: String(session?.id || "").trim() || createChatSession().id,
                title: String(session?.title || "שיחה חדשה").trim() || "שיחה חדשה",
                createdAt: String(session?.createdAt || new Date().toISOString()),
                updatedAt: String(session?.updatedAt || new Date().toISOString()),
                messages: Array.isArray(session?.messages)
                    ? session.messages
                        .map(message => ({
                            role: message?.role === "ai" ? "ai" : "user",
                            text: String(message?.text || "").trim(),
                            ts: String(message?.ts || "")
                        }))
                        .filter(message => !!message.text)
                    : []
            }));
            currentChatSessionId = String(parsed?.currentChatSessionId || "").trim();
        }
    } catch (e) {
        console.warn("Could not load chat sessions:", e);
    }

    if (!chatSessions.length) {
        const session = createChatSession();
        chatSessions = [session];
        currentChatSessionId = session.id;
    }
    if (!chatSessions.some(session => session.id === currentChatSessionId)) {
        currentChatSessionId = chatSessions[0].id;
    }

    chatSessions = chatSessions
        .sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime())
        .slice(0, MAX_CHAT_SESSIONS);

    refreshChatSessionPicker();
    loadChatSession(currentChatSessionId);
}

function refreshChatSessionPicker() {
    const select = document.getElementById("chatSessionSelect");
    if (!select) return;
    select.innerHTML = "";
    chatSessions
        .slice()
        .sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime())
        .forEach(session => {
            const option = document.createElement("option");
            option.value = session.id;
            option.innerText = session.title || "שיחה חדשה";
            select.appendChild(option);
        });
    select.value = currentChatSessionId;
}

function renderChatHistory() {
    const msgs = document.getElementById("chatMessages");
    if (!msgs) return;
    msgs.innerHTML = "";
    if (!chatHistory.length) {
        const welcome = document.createElement("div");
        welcome.className = "chat-bubble ai";
        welcome.style.backgroundColor = "#f0f8ff"; // צבע רקע ייחודי
        welcome.style.borderRadius = "10px"; // פינות מעוגלות
        welcome.style.padding = "10px"; // ריווח פנימי
        welcome.innerText = getWelcomeMessage();
        msgs.appendChild(welcome);
    } else {
        chatHistory.forEach(message => {
            const bubble = document.createElement("div");
            bubble.className = `chat-bubble ${message.role}`;
            bubble.style.marginBottom = "8px"; // ריווח בין בועות
            bubble.style.padding = "10px";
            bubble.style.borderRadius = "10px";
            bubble.style.backgroundColor = message.role === "ai" ? "#e6ffe6" : "#fff"; // צבעים שונים ל-AI ולמשתמש
            bubble.innerText = message.text;
            msgs.appendChild(bubble);
        });
    }
    msgs.scrollTop = msgs.scrollHeight;
}

function getCurrentSession() {
    return chatSessions.find(session => session.id === currentChatSessionId) || null;
}

function persistCurrentSession() {
    const session = getCurrentSession();
    if (!session) return;
    session.messages = chatHistory.slice(-120);
    session.updatedAt = new Date().toISOString();
    session.title = deriveChatTitle(session.messages);
    chatSessions = chatSessions
        .map(item => item.id === session.id ? session : item)
        .sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime())
        .slice(0, MAX_CHAT_SESSIONS);
    refreshChatSessionPicker();
    saveChatSessions();
}

function loadChatSession(sessionId) {
    const session = chatSessions.find(item => item.id === sessionId);
    if (!session) return;
    currentChatSessionId = session.id;
    chatHistory = (session.messages || []).map(message => ({
        role: message.role,
        text: message.text,
        ts: message.ts || ""
    }));
    refreshChatSessionPicker();
    renderChatHistory();
    setStatus(`שיחה נטענה: ${session.title || "שיחה חדשה"}`);
}

function createNewChat() {
    const session = createChatSession();
    chatSessions.unshift(session);
    currentChatSessionId = session.id;
    chatHistory = [];
    refreshChatSessionPicker();
    renderChatHistory();
    saveChatSessions();
    setStatus("נפתחה שיחה חדשה");
}

function addMessageToHistory(role, text) {
    const normalized = String(text || "").trim();
    if (!normalized) return;
    chatHistory.push({ role, text: normalized, ts: new Date().toISOString() });
    persistCurrentSession();
}

function getChatLearningInputText() {
    const inputText = chatHistory
        .filter(message => message.role === "user")
        .map(message => String(message.text || "").trim())
        .filter(text => text.length >= 12)
        .join("\n");

    console.log("[DEBUG] Input text for learning:", inputText); // לוג לבדיקת הטקסט הנכנס
    return inputText;
}

function getChatLearningCandidates(text) {
    const source = String(text || "");
    console.log("[DEBUG] Source text for candidates:", source); // לוג לבדיקת הטקסט המקורי

    const words = source.match(/[\u0590-\u05FF]{3,24}/g) || [];
    const stopwords = new Set(["של", "עם", "אבל", "האם", "אני", "אתה", "אתם", "היא", "הוא", "זה", "זאת", "מה", "איך", "למה", "כי", "גם", "רק", "עוד", "כדי"]);
    const counts = {};
    words.forEach(word => {
        const key = normalizeForCompare(word);
        if (!key || stopwords.has(key)) return;
        counts[key] = (counts[key] || 0) + 1;
    });

    const vocabulary = Object.entries(counts)
        .filter(([, count]) => count >= 2)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 40)
        .map(([word]) => word);

    console.log("[DEBUG] Extracted vocabulary:", vocabulary); // לוג לבדיקת המילים שנלמדו

    const phrasesByCount = {};
    source
        .split(/[\n.!?]+/)
        .map(item => item.replace(/\s+/g, " ").trim())
        .filter(item => item.split(" ").length >= 2 && item.split(" ").length <= 8)
        .forEach(item => {
            const key = normalizeForCompare(item);
            phrasesByCount[key] = (phrasesByCount[key] || 0) + 1;
        });

    const phrases = Object.entries(phrasesByCount)
        .filter(([, count]) => count >= 2)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 25)
        .map(([phrase]) => phrase);

    console.log("[DEBUG] Extracted phrases:", phrases); // לוג לבדיקת הביטויים שנלמדו

    return { vocabulary, phrases };
}

function learnFromChatHistory() {
    const sourceText = getChatLearningInputText();
    if (sourceText.length < 60) return;
    const { vocabulary, phrases } = getChatLearningCandidates(sourceText);
    if (!vocabulary.length && !phrases.length) return;
    recordLearnedCandidates(vocabulary, phrases, sourceText);
    savePersonalStyle();
    renderStylePanel();
}

function sanitizeLearnedItem(item, type) {
    const normalized = String(item || "")
        .trim()
        .replace(/["'`”“]+/g, "")
        .replace(/[.,;:!?()[\]{}]+$/g, "");
    if (!normalized) return null;
    if (type === "vocab") {
        if (normalized.includes(" ")) return null;
        if (normalized.length < 3 || normalized.length > 24) return null;
        return normalized;
    }
    const words = normalized.split(/\s+/).filter(Boolean);
    if (words.length < 2 || words.length > 8) return null;
    if (normalized.length > 60) return null;
    return normalized;
}

function escapeRegExp(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function countOccurrencesInText(text, token, type) {
    if (!text || !token) return 0;
    const source = String(text).toLowerCase();
    const needle = String(token).toLowerCase();
    if (!needle) return 0;

    if (type === "vocab") {
        const re = new RegExp(`(^|[^\\p{L}])${escapeRegExp(needle)}([^\\p{L}]|$)`, "giu");
        return (source.match(re) || []).length;
    }

    // For multi-word phrases we allow simple substring frequency.
    let count = 0;
    let pos = 0;
    while (true) {
        const idx = source.indexOf(needle, pos);
        if (idx === -1) break;
        count += 1;
        pos = idx + needle.length;
    }
    return count;
}

function normalizeForCompare(text) {
    return String(text || "")
        .trim()
        .replace(/\s+/g, " ")
        .toLowerCase();
}

function buildProtectedEntries(items = [], threshold = 0) {
    return uniqueStringsByNormalized(items).map(text => ({
        text,
        normalized: normalizeForCompare(text),
        count: threshold,
        protected: true
    }));
}

function mergeCountEntriesWithProtected(countEntries = [], protectedItems = [], threshold = 0) {
    const merged = new Map();
    (countEntries || []).forEach(entry => {
        merged.set(entry.normalized, { ...entry, protected: false });
    });
    buildProtectedEntries(protectedItems, threshold).forEach(entry => {
        const existing = merged.get(entry.normalized);
        merged.set(entry.normalized, {
            text: existing?.text || entry.text,
            normalized: entry.normalized,
            count: Math.max(existing?.count || 0, entry.count),
            protected: true
        });
    });
    return [...merged.values()];
}

function ensureProtectedItem(kind, text) {
    const normalizedText = String(text || "").trim();
    if (!normalizedText) return;
    if (kind === "phrase") {
        personalStyle.protectedPhrases = personalStyle.protectedPhrases || [];
        if (!hasNormalizedItem(personalStyle.protectedPhrases, normalizedText)) {
            personalStyle.protectedPhrases.push(normalizedText);
        }
        return;
    }
    personalStyle.protectedVocabulary = personalStyle.protectedVocabulary || [];
    if (!hasNormalizedItem(personalStyle.protectedVocabulary, normalizedText)) {
        personalStyle.protectedVocabulary.push(normalizedText);
    }
}

function removeProtectedItem(kind, text) {
    const normalizedText = normalizeForCompare(text);
    if (kind === "phrase") {
        personalStyle.protectedPhrases = (personalStyle.protectedPhrases || [])
            .filter(item => normalizeForCompare(item) !== normalizedText);
        return;
    }
    personalStyle.protectedVocabulary = (personalStyle.protectedVocabulary || [])
        .filter(item => normalizeForCompare(item) !== normalizedText);
}

function recomputeLearnedLists() {
    const manualVocabularySet = new Set((personalStyle.manualVocabulary || []).map(normalizeForCompare));
    const manualPhraseSet = new Set((personalStyle.manualPhrases || []).map(normalizeForCompare));
    personalStyle.learnedVocabulary = mergeCountEntriesWithProtected(
        getAggregatedCountEntries(personalStyle.learnedVocabularyCounts || {})
            .filter(entry => entry.count >= MIN_LEARNED_VOCAB_COUNT),
        personalStyle.protectedVocabulary || [],
        MIN_LEARNED_VOCAB_COUNT
    )
        .filter(entry => !manualVocabularySet.has(entry.normalized))
        .sort((a, b) => Number(b.protected) - Number(a.protected) || b.count - a.count || a.text.localeCompare(b.text, "he"))
        .map(entry => entry.text);

    personalStyle.learnedPhrases = mergeCountEntriesWithProtected(
        getAggregatedCountEntries(personalStyle.learnedPhraseCounts || {})
            .filter(entry => entry.count >= MIN_LEARNED_PHRASE_COUNT),
        personalStyle.protectedPhrases || [],
        MIN_LEARNED_PHRASE_COUNT
    )
        .filter(entry => !manualPhraseSet.has(entry.normalized))
        .sort((a, b) => Number(b.protected) - Number(a.protected) || b.count - a.count || a.text.localeCompare(b.text, "he"))
        .map(entry => entry.text);
}

function recordLearnedCandidates(vocabulary = [], phrases = [], sourceText = "") {
    personalStyle.learnedVocabularyCounts = personalStyle.learnedVocabularyCounts || {};
    personalStyle.learnedPhraseCounts = personalStyle.learnedPhraseCounts || {};

    const vocabCandidates = uniqueStrings([...(vocabulary || []), ...COMMON_ACADEMIC_MARKERS]);
    const phraseCandidates = uniqueStrings(phrases || []);

    vocabCandidates
        .map(item => sanitizeLearnedItem(item, "vocab"))
        .filter(Boolean)
        .forEach(item => {
            const rawHits = countOccurrencesInText(sourceText, item, "vocab");
            const hits = Math.min(MAX_COUNT_PER_FILE, Math.max(1, rawHits));
            personalStyle.learnedVocabularyCounts[item] = (personalStyle.learnedVocabularyCounts[item] || 0) + hits;
        });

    phraseCandidates
        .map(item => sanitizeLearnedItem(item, "phrase"))
        .filter(Boolean)
        .forEach(item => {
            const rawHits = countOccurrencesInText(sourceText, item, "phrase");
            const hits = Math.min(MAX_COUNT_PER_FILE, Math.max(1, rawHits));
            personalStyle.learnedPhraseCounts[item] = (personalStyle.learnedPhraseCounts[item] || 0) + hits;
        });

    recomputeLearnedLists();
}

function resetLearnedStyleData() {
    personalStyle.learnedVocabulary = [];
    personalStyle.learnedPhrases = [];
    personalStyle.learnedVocabularyCounts = {};
    personalStyle.learnedPhraseCounts = {};
    personalStyle.learnedNotes = [];
    personalStyle.examples = [];
    recomputeLearnedLists();
}

async function loadPersonalStyle() {
    try {
        const res = await fetch("./my_personal_style.json");
        if (res.ok) {
            const base = await res.json();
            personalStyle = normalizePersonalStyle(base);
        }
    } catch(e) {
        console.warn("Could not load my_personal_style.json:", e);
    }
    try {
        const stored = localStorage.getItem(STYLE_LS_KEY);
        if (stored) {
            const updates = JSON.parse(stored);
            personalStyle = mergePersonalStyles(personalStyle, updates);
        }
    } catch(e) { console.warn("Could not load style from localStorage:", e); }
    recomputeLearnedLists();
    renderStylePanel();
    updateStyleBadge();
}

function savePersonalStyle() {
    personalStyle.last_updated = new Date().toISOString();
    try { localStorage.setItem(STYLE_LS_KEY, JSON.stringify(personalStyle)); }
    catch(e) { console.warn("Could not save style:", e); }
    updateStyleBadge();
}

function getExceptionsRows() {
    const manualVocabSet = new Set((personalStyle.manualVocabulary || []).map(normalizeForCompare));
    const manualPhraseSet = new Set((personalStyle.manualPhrases || []).map(normalizeForCompare));
    const rows = [];

    for (const entry of mergeCountEntriesWithProtected(
        getAggregatedCountEntries(personalStyle.learnedVocabularyCounts || {}),
        personalStyle.protectedVocabulary || [],
        MIN_LEARNED_VOCAB_COUNT
    )) {
        const numericCount = entry.count;
        const isManual = manualVocabSet.has(entry.normalized);
        const isProtected = !!entry.protected;
        const passed = (numericCount >= MIN_LEARNED_VOCAB_COUNT || isProtected) && !isManual;
        rows.push({
            text: entry.text,
            kind: "מילה",
            key: "vocab",
            count: numericCount,
            threshold: MIN_LEARNED_VOCAB_COUNT,
            status: passed ? "approved" : isManual ? "rejected" : "pending",
            reason: passed ? (isProtected ? "אושר ידנית ונשמר גם אחרי עדכון" : "נשמר כנלמד") : isManual ? "קיים ידנית" : "מתחת לסף"
        });
    }

    for (const entry of mergeCountEntriesWithProtected(
        getAggregatedCountEntries(personalStyle.learnedPhraseCounts || {}),
        personalStyle.protectedPhrases || [],
        MIN_LEARNED_PHRASE_COUNT
    )) {
        const numericCount = entry.count;
        const isManual = manualPhraseSet.has(entry.normalized);
        const isProtected = !!entry.protected;
        const passed = (numericCount >= MIN_LEARNED_PHRASE_COUNT || isProtected) && !isManual;
        rows.push({
            text: entry.text,
            kind: "ביטוי",
            key: "phrase",
            count: numericCount,
            threshold: MIN_LEARNED_PHRASE_COUNT,
            status: passed ? "approved" : isManual ? "rejected" : "pending",
            reason: passed ? (isProtected ? "אושר ידנית ונשמר גם אחרי עדכון" : "נשמר כנלמד") : isManual ? "קיים ידנית" : "מתחת לסף"
        });
    }

    rows.sort((a, b) => b.count - a.count || a.text.localeCompare(b.text, "he"));
    return rows;
}

function buildExceptionsBoardHtml(rows, expanded = false) {
    if (!rows.length) return "אין נתונים עדיין";

    return rows.map(row => {
        const statusLabel = row.status === "approved" ? "✅ מאושר" : row.status === "pending" ? "⏳ בהמתנה" : "⛔ נדחה";
        const approveButton = expanded && row.status === "pending"
            ? `<button class="btn-outline btn-mini exception-approve-btn" data-kind="${row.key}" data-text="${encodeURIComponent(row.text)}">אשר ידנית</button>`
            : "";
        return `<div class="exceptions-row ${row.status}">
            <div class="exceptions-text">${escapeHtml(row.text)}</div>
            <div class="exceptions-meta">${row.kind} · ספירה: ${row.count} · סף: ${row.threshold}</div>
            <div class="exceptions-reason">${statusLabel} · ${row.reason}</div>
            ${approveButton ? `<div class="exceptions-actions">${approveButton}</div>` : ""}
        </div>`;
    }).join("");
}

function getPersonalStyleContext() {
    const vocabulary = getCombinedVocabulary();
    const phrases = getCombinedPhrases();
    const hasContent = (vocabulary.length || phrases.length || personalStyle.examples?.length || personalStyle.notes || personalStyle.learnedNotes?.length);
    if (!hasContent) return "";

    let ctx = "\n\n== PERSONAL WRITING STYLE — apply throughout your entire response ==";
    if (vocabulary.length) {
        ctx += `\nPreferred vocabulary (weave naturally): ${vocabulary.join(", ")}`;
    }
    if (phrases.length) {
        ctx += `\nPreferred phrases/transitions: ${phrases.join("; ")}`;
    }
    if (personalStyle.notes) {
        ctx += `\nStyle notes: ${personalStyle.notes}`;
    }
    if (personalStyle.learnedNotes?.length) {
        ctx += `\nObserved style patterns from past works: ${personalStyle.learnedNotes.slice(-3).join(" | ")}`;
    }
    if (personalStyle.examples?.length) {
        const samples = personalStyle.examples.slice(-3).join("\n\n---\n\n");
        ctx += `\n\nSTYLE EXAMPLES — adopt this exact vocabulary, rhythm, and sentence structure:\n${samples}`;
    }
    ctx += "\n== END PERSONAL STYLE ==";
    return ctx;
}

function renderStylePanel() {
    const chipsEl = document.getElementById("vocabChips");
    if (chipsEl) {
        chipsEl.innerHTML = "";
        getStyleChipItems().forEach(({ text, type, source }) => {
            const chip = document.createElement("span");
            chip.className = `vocab-chip${type === "phrase" ? " phrase-chip" : ""}${source === "learned" ? " learned-chip" : ""}`;
            const isProtected = type === "phrase"
                ? hasNormalizedItem(personalStyle.protectedPhrases || [], text)
                : hasNormalizedItem(personalStyle.protectedVocabulary || [], text);
            chip.title = source === "learned"
                ? (isProtected ? "נלמד ונשמר מפני עדכון עתידי" : "נלמד אוטומטית")
                : "הוזן ידנית";
            chip.innerHTML = `${escapeHtml(text)} <button class="chip-remove" data-text="${escapeHtml(text)}" data-type="${type}" data-source="${source}">×</button>`;
            chipsEl.appendChild(chip);
        });
    }

    const notesEl = document.getElementById("styleNotes");
    if (notesEl) {
        notesEl.value = personalStyle.notes || "";
    }

    const statusEl = document.getElementById("styleStatus");
    if (statusEl) {
        const manualCount = (personalStyle.manualVocabulary?.length || 0) + (personalStyle.manualPhrases?.length || 0);
        const learnedCount = getVisibleLearnedVocabulary().length + getVisibleLearnedPhrases().length;
        const dateStr = personalStyle.last_updated
            ? new Date(personalStyle.last_updated).toLocaleDateString("he-IL")
            : "—";
        statusEl.innerText = (manualCount + learnedCount + (personalStyle.examples?.length || 0)) > 0
            ? `ידני: ${manualCount} · נלמד: ${learnedCount} · דוגמאות: ${personalStyle.examples?.length || 0} · עודכן: ${dateStr}`
            : "אין סגנון אישי מוגדר עדיין";
    }

    renderExceptionsBoard();
}

function updateStyleBadge() {
    const btn = document.getElementById("btnStyle");
    if (!btn) return;
    const count = (personalStyle.manualVocabulary?.length || 0)
        + (personalStyle.manualPhrases?.length || 0)
        + getVisibleLearnedVocabulary().length
        + getVisibleLearnedPhrases().length;
    const hasStyle = count > 0 || (personalStyle.examples?.length > 0) || !!personalStyle.notes || !!personalStyle.learnedNotes?.length;
    btn.classList.toggle("style-active", !!hasStyle);
    btn.title = hasStyle ? `סגנון אישי פעיל (${count} מילים/ביטויים)` : "סגנון אישי";
}

function approveExceptionItem(kind, text) {
    const normalizedText = String(text || "").trim();
    if (!normalizedText) return;
    ensureProtectedItem(kind, normalizedText);

    if (kind === "phrase") {
        personalStyle.learnedPhraseCounts = personalStyle.learnedPhraseCounts || {};
        const matchingKeys = Object.keys(personalStyle.learnedPhraseCounts).filter(key => normalizeForCompare(key) === normalizeForCompare(normalizedText));
        if (!matchingKeys.length) matchingKeys.push(normalizedText);
        matchingKeys.forEach(key => {
            personalStyle.learnedPhraseCounts[key] = Math.max(personalStyle.learnedPhraseCounts[key] || 0, MIN_LEARNED_PHRASE_COUNT);
        });
    } else {
        personalStyle.learnedVocabularyCounts = personalStyle.learnedVocabularyCounts || {};
        const matchingKeys = Object.keys(personalStyle.learnedVocabularyCounts).filter(key => normalizeForCompare(key) === normalizeForCompare(normalizedText));
        if (!matchingKeys.length) matchingKeys.push(normalizedText);
        matchingKeys.forEach(key => {
            personalStyle.learnedVocabularyCounts[key] = Math.max(personalStyle.learnedVocabularyCounts[key] || 0, MIN_LEARNED_VOCAB_COUNT);
        });
    }

    recomputeLearnedLists();
    savePersonalStyle();
    renderStylePanel();
    renderExceptionsBoard();
    showToast(`✅ אושר ונשמר: ${normalizedText}`);
}

function toggleExceptionsModal(show) {
    const modal = document.getElementById("exceptionsModal");
    if (!modal) return;
    modal.classList.toggle("hidden", !show);
    if (show) renderExceptionsBoard();
}

function renderExceptionsBoard() {
    const boardEl = document.getElementById("exceptionsBoard");
    const modalBoardEl = document.getElementById("exceptionsModalBoard");
    const rows = getExceptionsRows();
    if (boardEl) boardEl.innerHTML = buildExceptionsBoardHtml(rows, false);
    if (modalBoardEl) modalBoardEl.innerHTML = buildExceptionsBoardHtml(rows, true);
}

function addVocabItem() {
    const input = document.getElementById("vocabInput");
    const text = input.value.trim();
    if (!text) return;
    if (text.includes(" ")) {
        if (!personalStyle.manualPhrases) personalStyle.manualPhrases = [];
        if (!hasNormalizedItem(personalStyle.manualPhrases, text)) personalStyle.manualPhrases.push(text);
        personalStyle.learnedPhrases = (personalStyle.learnedPhrases || []).filter(item => item !== text);
    } else {
        if (!personalStyle.manualVocabulary) personalStyle.manualVocabulary = [];
        if (!hasNormalizedItem(personalStyle.manualVocabulary, text)) personalStyle.manualVocabulary.push(text);
        personalStyle.learnedVocabulary = (personalStyle.learnedVocabulary || []).filter(item => item !== text);
    }
    input.value = "";
    recomputeLearnedLists();
    renderStylePanel();
    savePersonalStyle();
}

function removeVocabItem(text, type, source = "manual") {
    if (type === "phrase") {
        if (source === "learned") {
            personalStyle.learnedPhrases = (personalStyle.learnedPhrases || []).filter(p => p !== text);
            removeProtectedItem("phrase", text);
            if (personalStyle.learnedPhraseCounts) {
                const normalizedKey = normalizeForCompare(text);
                Object.keys(personalStyle.learnedPhraseCounts).forEach(key => {
                    if (normalizeForCompare(key) === normalizedKey) delete personalStyle.learnedPhraseCounts[key];
                });
            }
        } else {
            personalStyle.manualPhrases = (personalStyle.manualPhrases || []).filter(p => p !== text);
        }
    } else {
        if (source === "learned") {
            personalStyle.learnedVocabulary = (personalStyle.learnedVocabulary || []).filter(v => v !== text);
            removeProtectedItem("vocab", text);
            if (personalStyle.learnedVocabularyCounts) {
                const normalizedKey = normalizeForCompare(text);
                Object.keys(personalStyle.learnedVocabularyCounts).forEach(key => {
                    if (normalizeForCompare(key) === normalizedKey) delete personalStyle.learnedVocabularyCounts[key];
                });
            }
        } else {
            personalStyle.manualVocabulary = (personalStyle.manualVocabulary || []).filter(v => v !== text);
        }
    }
    recomputeLearnedLists();
    renderStylePanel();
    savePersonalStyle();
}

function exportStyleJson() {
    const blob = new Blob([JSON.stringify(personalStyle, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "my_personal_style.json"; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function importStyleJson(file) {
    try {
        const text = await readFileAsText(file);
        const imported = JSON.parse(text);
        personalStyle = normalizePersonalStyle(imported);
        savePersonalStyle(); renderStylePanel(); updateStyleBadge();
        document.getElementById("styleStatus").innerText = "✅ סגנון יובא בהצלחה!";
    } catch(e) {
        document.getElementById("styleStatus").innerText = "❌ שגיאה בייבוא: " + e.message;
    }
}

function autoUpdateStyleFromInteraction(selectedText) {
    if (!selectedText || selectedText.length < 80) return;
    const newExample = selectedText.slice(0, 1200);
    const updatedExamples = mergeExamples(personalStyle.examples, [newExample]);
    if (updatedExamples.length !== (personalStyle.examples || []).length) {
        personalStyle.examples = updatedExamples;
        savePersonalStyle(); updateStyleBadge();
    }
}

async function learnStyleFromFile(file) {
    const apiKey = getApiKey();
    if (!apiKey) {
        document.getElementById("styleLearningStatus").innerText = "נדרש מפתח Gemini. הגדר בהגדרות (⚙️).";
        return;
    }
    const statusEl = document.getElementById("styleLearningStatus");
    statusEl.innerText = "⏳ קורא קובץ...";
    try {
        const text = await extractTextFromFile(file);
        if (text.length < 100) throw new Error("הקובץ קצר מדי לניתוח.");
        statusEl.innerText = "⏳ מנתח סגנון (Gemini)...";
        const analysisPrompt = `You are a writing style analyst. Analyze the Hebrew academic text below and extract the author's PERSONAL writing characteristics.
Return ONLY valid JSON (no markdown, no extra text):
{
    "vocabulary": ["distinctive Hebrew words/terms this author uses"],
    "phrases": ["transition phrases and connecting expressions this author uses"],
  "example": "A representative 2-3 sentence excerpt from the text",
  "notes": "One sentence: academic level, formality, distinctive traits"
}

Text:
${text.slice(0, 6000)}`;
        const response = await callGeminiText(apiKey, analysisPrompt);
        const jsonMatch = response.match(/\{[\s\S]*\}/);
        if (!jsonMatch) throw new Error("AI לא החזיר JSON תקין.");
        const extracted = JSON.parse(jsonMatch[0]);
        recordLearnedCandidates(extracted.vocabulary || [], extracted.phrases || [], text);
        if (extracted.example) {
            personalStyle.examples = mergeExamples(personalStyle.examples, [extracted.example]);
        }
        if (extracted.notes) {
            personalStyle.learnedNotes = uniqueStrings([...(personalStyle.learnedNotes || []), extracted.notes]).slice(-5);
        }
        savePersonalStyle(); renderStylePanel();
        statusEl.innerText = `✅ נותחו ${extracted.vocabulary?.length || 0} מילים ו-${extracted.phrases?.length || 0} ביטויים · נוספו רק פריטים שכיחים`;
    } catch(e) {
        statusEl.innerText = "❌ שגיאה: " + e.message;
        console.error(e);
    }
}

// ─── Auto-learn from past-works/ ────────────────────────────────────────
async function autoLearnFromPastWorks(force = false) {
    const apiKey = getApiKey();
    if (!apiKey) return; // אין מפתח — מדלג בשקט
    try {
        const res = await fetch('./past-works/index.json', { cache: 'no-store' });
        if (!res.ok) return;
        const { files } = await res.json();
        if (!files?.length) {
            const statusEl = document.getElementById("styleLearningStatus");
            if (statusEl && force) statusEl.innerText = "אין קבצים בתיקיית past-works/";
            return;
        }

        const learned = JSON.parse(localStorage.getItem(PAST_WORKS_CACHE_KEY) || "{}");

        const toLearn = force ? files : files.filter(f => {
            const hash = `${f.name}:${f.size}:${f.modified}`;
            return learned[f.name] !== hash;
        });
        if (!toLearn.length) return;

        setStatus(`🎓 לומד מ-${toLearn.length} עבודות קודמות...`);
        for (const fileInfo of toLearn) {
            try {
                const fileRes = await fetch(`./past-works/${encodeURIComponent(fileInfo.name)}`);
                if (!fileRes.ok) continue;
                const blob = await fileRes.blob();
                const file = new File([blob], fileInfo.name, { type: blob.type });
                await learnStyleFromFile(file);
                const hash = `${fileInfo.name}:${fileInfo.size}:${fileInfo.modified}`;
                learned[fileInfo.name] = hash;
                localStorage.setItem(PAST_WORKS_CACHE_KEY, JSON.stringify(learned));
            } catch(e) {
                console.warn(`Could not learn from ${fileInfo.name}:`, e);
            }
        }
        setStatus("✅ סגנון אישי עודכן מעבודות קודמות!");
        setTimeout(() => setStatus("מוכן! הסגנון האישי שלך פעיל 🎯"), 3000);
    } catch(e) {
        console.warn("Past-works auto-learn skipped:", e);
    }
}

async function relearnPastWorks() {
    const btn = document.getElementById("relearnPastWorksBtn");
    const statusEl = document.getElementById("styleLearningStatus");
    if (btn) btn.disabled = true;
    if (statusEl) statusEl.innerText = "🔄 מאפס מטמון למידה ומנתח מחדש...";
    try {
        resetLearnedStyleData();
        savePersonalStyle();
        renderStylePanel();
        localStorage.removeItem(PAST_WORKS_CACHE_KEY);
        await autoLearnFromPastWorks(true);
        if (statusEl) statusEl.innerText = "✅ הלמידה האוטומטית אופסה והורצה מחדש";
    } catch (e) {
        if (statusEl) statusEl.innerText = "❌ שגיאה בלמידה מחדש: " + e.message;
    } finally {
        if (btn) btn.disabled = false;
    }
}

function resetAllStyleDataOnce() {
    const now = Date.now();
    if (now > resetAllArmedUntil) {
        resetAllArmedUntil = now + 5000;
        const styleStatus = document.getElementById("styleStatus");
        if (styleStatus) {
            styleStatus.innerText = "לחץ שוב על 'איפוס מלא' בתוך 5 שניות כדי לאשר";
        }
        return;
    }
    resetAllArmedUntil = 0;

    personalStyle = createEmptyPersonalStyle();
    try { localStorage.removeItem(STYLE_LS_KEY); } catch (e) { console.warn("Could not clear style storage", e); }
    try { localStorage.removeItem(PAST_WORKS_CACHE_KEY); } catch (e) { console.warn("Could not clear past-works cache", e); }

    savePersonalStyle();
    renderStylePanel();
    updateStyleBadge();

    const notesEl = document.getElementById("styleNotes");
    if (notesEl) notesEl.value = "";

    const learnStatus = document.getElementById("styleLearningStatus");
    if (learnStatus) learnStatus.innerText = "✅ בוצע איפוס מלא לזיכרון הסגנון";

    const styleStatus = document.getElementById("styleStatus");
    if (styleStatus) styleStatus.innerText = "אופס. אפשר להתחיל ללמוד מחדש נקי";
}

// ─── Smart Routing ────────────────────────────────────────────────────────
function detectRoutingTarget(message) {
    // הסוכן הפעיל תמיד קובע את תביב הניתוב
    if (activeAgentId && AGENTS_CONFIG[activeAgentId]) {
        return AGENTS_CONFIG[activeAgentId].route;
    }
    return RESEARCH_INTENT.test(message) ? "perplexity" : "gemini";
}

function activateAgent(agentId) {
    const agent = AGENTS_CONFIG[agentId];
    if (!agent) return;
    // toggle: אם אותו סוכן כבר פעיל, כבה אותו
    if (activeAgentId === agentId) {
        activeAgentId = null;
    } else {
        activeAgentId = agentId;
    }
    renderActiveAgent();
}

function renderActiveAgent() {
    const badge = document.getElementById("activeAgentBadge");
    const chatInput = document.getElementById("chatInput");
    document.querySelectorAll(".quick-chip").forEach(chip => {
        chip.classList.toggle("agent-active", chip.dataset.agent === activeAgentId);
    });
    if (!activeAgentId || !AGENTS_CONFIG[activeAgentId]) {
        if (badge) badge.classList.add("hidden");
        if (chatInput) chatInput.placeholder = "שאל כל דבר — עריכה, ניסוח, מחקר, מקורות...";
        return;
    }
    const agent = AGENTS_CONFIG[activeAgentId];
    if (badge) {
        badge.classList.remove("hidden");
        badge.innerHTML = `<span class="agent-badge-label">${agent.label}</span><button class="agent-badge-close" title="בטל סוכן">✕</button>`;
        badge.querySelector(".agent-badge-close").onclick = () => { activeAgentId = null; renderActiveAgent(); };
    }
    if (chatInput) chatInput.placeholder = agent.placeholder;
}

function getActiveAgentContext() {
    if (!activeAgentId) return "";
    const agent = AGENTS_CONFIG[activeAgentId];
    if (!agent || !agent.systemCtx) return "";
    return `\n\n${agent.systemCtx}`;
}

// ─── Chat ─────────────────────────────────────────────────────────────────
async function sendChatMessage() {
    const input = document.getElementById("chatInput");
    const msg = input.value.trim();
    if (!msg) return;
    if (isChatSendInFlight) {
        debugLog("send.skip already-in-flight");
        return;
    }
    isChatSendInFlight = true;
    input.value = "";

    addChatBubble(msg, "user");
    addMessageToHistory("user", msg);

    const loadingBubble = addChatBubble("חושב...", "ai loading");
    const sendBtn = document.getElementById("sendChat");
    sendBtn.disabled = true;

    try {
        const routeTarget = detectRoutingTarget(msg);
        const docText = await getDocumentText();
        let aiText = "";
        debugLog(`send.start route=${routeTarget} msgLen=${msg.length} docLen=${docText.length} selLen=${currentSelectedText.length}`);

        if (routeTarget === "perplexity") {
            setStatus("🔍 מחפש מקורות אקדמיים (Perplexity)...");
            const query = currentSelectedText
                ? `${msg}\n\nהקשר מהמסמך:\n${currentSelectedText.slice(0, 800)}`
                : msg;
            aiText = await callPerplexityChat(query, docText);
            validateCitationsInText(aiText, "status");
        } else {
            const apiKey = getApiKey();
            if (!apiKey) {
                loadingBubble.innerText = "אנא הגדר מפתח Gemini API בהגדרות (⚙️).";
                loadingBubble.classList.remove("loading");
                document.getElementById("settingsPanel").classList.remove("hidden");
                return;
            }
            setStatus("✍️ מחשב...");
            const historyBlock = chatHistory.slice(-10)
                .map(m => `[${m.role === "user" ? "User" : "Assistant"}]: ${m.text.slice(0, 600)}`)
                .join("\n");
            const styleCtx = getPersonalStyleContext();
            const studyCtx = getStudyContext();
            const reviewCtx = getLecturerReviewContext(msg);
            const agentCtx = getActiveAgentContext();
            const userRules = getAgentRulesText();
            const rulesCtx = getAgentRulesContext(userRules);

            const systemPrompt = `${SYSTEM_INSTRUCTIONS}${styleCtx}${agentCtx}

== CHAT MODE — CLARIFICATION PROTOCOL ==
Before writing academic content from scratch, assess if you need more personal info.
Ask 2-4 questions ONLY when: request is vague and academic; needs personal reflection; missing key details.
Do NOT ask when: clear editing/proofreading task; sufficient context already in history.
== END PROTOCOL ==${studyCtx}${reviewCtx}${rulesCtx}
Document Context:
${docText.slice(0, 4000)}
${currentSelectedText ? `\nSelected Text:\n${currentSelectedText}` : ""}

Conversation History:
${historyBlock}

User: ${msg}

Internal instruction: silently verify full compliance with USER WRITING RULES before sending the final answer.`;
            debugLog(`send.gemini.prompt len=${systemPrompt.length} history=${chatHistory.length} rulesLen=${userRules.length}`);
            if (userRules) debugLog(`send.rules.preview ${userRules.slice(0, 120).replace(/\s+/g, " ")}`);

            loadingBubble.innerText = "";
            aiText = await callGeminiStream(apiKey, systemPrompt, partial => {
                loadingBubble.innerText = partial;
            });
            debugLog(`send.gemini.done answerLen=${aiText.length}`);
            if (currentSelectedText) autoUpdateStyleFromInteraction(currentSelectedText);
        }

        loadingBubble.classList.remove("loading");
        loadingBubble.innerText = aiText;
        addMessageToHistory("ai", aiText);
        learnFromChatHistory();
        debugLog(`send.done answerLen=${aiText.length}`);

        if (routeTarget === "perplexity") {
            const badge = document.createElement("div");
            badge.className = "source-badge";
            badge.innerText = "📡 מקורות: Perplexity";
            if (loadingBubble.parentElement) {
                loadingBubble.parentElement.insertBefore(badge, loadingBubble);
            } else {
                debugLog("send.perplexity.badge.skip no-parent");
            }
        }

        const actionsDiv = document.createElement("div");
        actionsDiv.className = "bubble-actions";
        actionsDiv.appendChild(createBubbleAction("🧩 לפי מיקומים", async () => {
            try {
                const inserted = await insertChatResponseBySearchAndInsert(aiText);
                showToast(`✅ הוכנסו ${inserted} הצעות למקומות הנכונים`);
            } catch(e) {
                debugLog(`sections.searchInsert.error ${e?.message || String(e)}`);
                setStatus("שגיאה: " + e.message);
            }
        }));
        actionsDiv.appendChild(createBubbleAction("📌 Track Change", async () => {
            try { await insertTextAsTracked("\n" + aiText); showToast("✅ הוכנס כ-Track Change"); }
            catch(e) { setStatus("שגיאה: " + e.message); }
        }));
        actionsDiv.appendChild(createBubbleAction("📄 הכנס ישירות", async () => {
            try { await insertTextDirect(aiText); showToast("✅ הוכנס ישירות"); }
            catch(e) { setStatus("שגיאה: " + e.message); }
        }));
        actionsDiv.appendChild(createBubbleAction("📋 העתק", () => {
            navigator.clipboard.writeText(aiText).then(() => setStatus("הועתק!"));
        }));
        loadingBubble.after(actionsDiv);

        setStatus("הושלם!");
    } catch(e) {
        console.error(e);
        loadingBubble.innerText = "שגיאה: " + e.message;
        loadingBubble.classList.remove("loading");
        setStatus("שגיאה: " + e.message);
        debugLog(`send.error ${e.message}`);
    } finally {
        isChatSendInFlight = false;
        sendBtn.disabled = false;
        const msgs = document.getElementById("chatMessages");
        msgs.scrollTop = msgs.scrollHeight;
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

function createBubbleAction(label, handler) {
    const btn = document.createElement("button");
    btn.className = "bubble-action-btn";
    btn.innerText = label;
    btn.onclick = handler;
    return btn;
}

function setupChatEvents() {
    const sendBtn = document.getElementById("sendChat");
    const chatInput = document.getElementById("chatInput");
    const newChatBtn = document.getElementById("newChatBtn");
    const sessionSelect = document.getElementById("chatSessionSelect");
    if (sendBtn && !sendBtn.dataset.bound) {
        sendBtn.onclick = sendChatMessage;
        sendBtn.dataset.bound = "true";
    }
    if (chatInput && !chatInput.dataset.bound) {
        chatInput.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                sendChatMessage();
            }
        });
        chatInput.dataset.bound = "true";
    }
    if (newChatBtn && !newChatBtn.dataset.bound) {
        newChatBtn.onclick = createNewChat;
        newChatBtn.dataset.bound = "true";
    }
    if (sessionSelect && !sessionSelect.dataset.bound) {
        sessionSelect.onchange = () => {
            if (sessionSelect.value) loadChatSession(sessionSelect.value);
        };
        sessionSelect.dataset.bound = "true";
    }
}

// ─── Perplexity API ───────────────────────────────────────────────────────
async function callPerplexityChat(query, docContext) {
    const apiKey = getPerplexityApiKey();
    if (!apiKey) throw new Error("לא הוגדר מפתח Perplexity API. הגדר בהגדרות (⚙️).");
    const userRules = getAgentRulesText();
    const rulesCtx = getAgentRulesContext(userRules);
    const systemPrompt = `You are an expert academic research assistant for an Israeli university student.
Find high-quality, real scholarly sources and provide academic citations.
Always respond in Hebrew. Format all citations in APA 7th edition.
CRITICAL: Only cite REAL sources you are confident exist. Never invent DOIs or page numbers.
${docContext ? `Document context:\n${docContext.slice(0, 2000)}` : ""}
${rulesCtx ? `\nUser writing rules:\n${rulesCtx.replace(/^\n+/, "")}` : ""}`;
    debugLog(`perplexity.request qLen=${query.length} promptLen=${systemPrompt.length} rulesLen=${userRules.length}`);
    const res = await fetch("https://api.perplexity.ai/chat/completions", {
        method: "POST",
        headers: { "Authorization": `Bearer ${apiKey}`, "Content-Type": "application/json" },
        body: JSON.stringify({
            model: "sonar-pro",
            messages: [
                { role: "system", content: systemPrompt },
                { role: "user", content: query }
            ]
        })
    });
    if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        debugLog(`perplexity.error status=${res.status}`);
        throw new Error(err?.error?.message || `Perplexity API ${res.status}`);
    }
    const data = await res.json();
    const content = data.choices?.[0]?.message?.content;
    if (!content) throw new Error("Perplexity: תגובה ריקה מה-API");
    debugLog(`perplexity.success answerLen=${content.length}`);
    return content;
}

// ─── Gemini API ───────────────────────────────────────────────────────────
function previewForDebug(value, maxLen = 280) {
    const text = String(value || "").replace(/\s+/g, " ").trim();
    return text.length > maxLen ? `${text.slice(0, maxLen)}...` : text;
}

function parseJsonSafe(text) {
    try {
        return JSON.parse(text);
    } catch {
        return null;
    }
}

function buildApiErrorMessage(provider, res, rawErrorBody) {
    const parsed = parseJsonSafe(rawErrorBody);
    const statusPart = `${provider} API ${res.status}${res.statusText ? ` ${res.statusText}` : ""}`;

    const errMessage = parsed?.error?.message || parsed?.message;
    const errStatus = parsed?.error?.status || parsed?.status;
    const errCode = parsed?.error?.code || parsed?.code;

    const detailParts = [];
    if (errStatus) detailParts.push(`status=${errStatus}`);
    if (errCode != null) detailParts.push(`code=${errCode}`);
    if (errMessage) detailParts.push(`message=${errMessage}`);

    if (detailParts.length) {
        return `${statusPart} | ${detailParts.join(" | ")}`;
    }
    if (rawErrorBody) {
        return `${statusPart} | body=${previewForDebug(rawErrorBody, 220)}`;
    }
    return statusPart;
}

async function callGeminiText(apiKey, prompt, options = {}) {
    const generationConfig = {
        temperature: 0.85,
        topP: 0.95,
        topK: 40,
        ...(options.generationConfig || {})
    };
    const requestBody = {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig
    };
    const bodyText = JSON.stringify(requestBody);
    debugLog(`gemini.text.request promptLen=${prompt.length} bodyLen=${bodyText.length}`);

    const res = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`,
        {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: bodyText
        }
    );

    if (!res.ok) {
        const errBody = await res.text().catch(() => "");
        const detailedError = buildApiErrorMessage("Gemini", res, errBody);
        debugLog(`gemini.text.error ${detailedError}`);
        if (errBody) debugLog(`gemini.text.errorBody ${previewForDebug(errBody, 400)}`);
        throw new Error(detailedError);
    }

    const data = await res.json();
    const text = data.candidates?.[0]?.content?.parts?.[0]?.text;
    if (text == null) {
        const finishReason = data?.candidates?.[0]?.finishReason;
        const blockReason = data?.promptFeedback?.blockReason;
        const reasonInfo = [
            finishReason ? `finishReason=${finishReason}` : "",
            blockReason ? `blockReason=${blockReason}` : ""
        ].filter(Boolean).join(" | ");
        const msg = reasonInfo
            ? `Gemini: לא התקבל טקסט (${reasonInfo})`
            : "Gemini: לא התקבל טקסט (safety filter או quota)";
        debugLog(`gemini.text.empty ${msg}`);
        throw new Error(msg);
    }
    debugLog(`gemini.text.success answerLen=${text.length}`);
    return text;
}

async function callGeminiStream(apiKey, prompt, onChunk) {
    const requestBody = {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.85, topP: 0.95, topK: 40 }
    };
    const bodyText = JSON.stringify(requestBody);
    debugLog(`gemini.stream.request promptLen=${prompt.length} bodyLen=${bodyText.length}`);

    const res = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:streamGenerateContent?alt=sse&key=${apiKey}`,
        {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: bodyText
        }
    );

    if (!res.ok) {
        const errBody = await res.text().catch(() => "");
        const detailedError = buildApiErrorMessage("Gemini", res, errBody);
        debugLog(`gemini.stream.error ${detailedError}`);
        if (errBody) debugLog(`gemini.stream.errorBody ${previewForDebug(errBody, 400)}`);
        throw new Error(detailedError);
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let fullText = "", buffer = "";
    while (true) {
        const { done, value } = await reader.read();
        if (done) { buffer += decoder.decode(); break; }
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop();
        for (const line of lines) {
            if (!line.startsWith("data: ")) continue;
            const json = line.slice(6).trim();
            if (!json || json === "[DONE]") continue;
            try {
                const parsed = JSON.parse(json);
                const chunk = parsed.candidates?.[0]?.content?.parts?.[0]?.text ?? "";
                if (chunk) { fullText += chunk; onChunk(fullText); }
            } catch { /* ignore malformed SSE */ }
        }
    }
    if (buffer.startsWith("data: ")) {
        const json = buffer.slice(6).trim();
        if (json && json !== "[DONE]") {
            try {
                const parsed = JSON.parse(json);
                const chunk = parsed.candidates?.[0]?.content?.parts?.[0]?.text ?? "";
                if (chunk) { fullText += chunk; onChunk(fullText); }
            } catch { /* ignore */ }
        }
    }
    debugLog(`gemini.stream.success answerLen=${fullText.length}`);
    return fullText;
}

// ─── Search & Insert מנגנון לפי Gemini Routing ────────────────────────────
async function buildAIRoutingMap(aiText, docContext = "") {
    const apiKey = getApiKey();
    if (!apiKey) throw new Error("נדרש מפתח Gemini");
    
    const routingPrompt = `You are an expert document router. Your task is to analyze the following Hebrew academic text and generate a routing map.

Return ONLY valid JSON with this exact shape:
{
  "suggestions": [
    {
      "targetLocation": "exact text fragment from the document where this suggestion belongs",
      "suggestionText": "the suggestion or revision for that location",
      "confidence": 0.95
    }
  ]
}

Rules:
BE PRECISE: targetLocation must be an exact substring from the document that you are confident exists.
DO NOT INVENT document text.
If you cannot find a clear target location in the original document, skip that suggestion.
Order suggestions by confidence (highest first).
Return max 10 suggestions.

Original Document:
${docContext || currentSelectedText || aiText.slice(0, 2000)}

AI Response to Route:
${aiText.slice(0, 8000)}`;
    
    debugLog(`routing.prompt.start promptLen=${routingPrompt.length}`);
    const response = await callGeminiText(apiKey, routingPrompt, {
        generationConfig: {
            temperature: 0.2,
            topP: 0.9,
            topK: 20,
            responseMimeType: "application/json"
        }
    });

    let parsed;
    try {
        parsed = extractJsonObject(response);
    } catch (e) {
        debugLog(`routing.parse.error ${e?.message || String(e)} rawPreview=${previewForDebug(response, 260)}`);
        const rescuePrompt = `Convert the following text into STRICT JSON only.
Return ONLY valid JSON with this exact shape:
{"suggestions":[{"targetLocation":"...","suggestionText":"...","confidence":0.9}]}

Text to convert:
${response.slice(0, 6000)}`;
        const rescue = await callGeminiText(apiKey, rescuePrompt, {
            generationConfig: {
                temperature: 0,
                topP: 0.1,
                topK: 1,
                responseMimeType: "application/json"
            }
        });
        parsed = extractJsonObject(rescue);
    }

    const rawSuggestions = Array.isArray(parsed)
        ? parsed
        : Array.isArray(parsed?.suggestions)
            ? parsed.suggestions
            : [];

    const suggestions = rawSuggestions
            .map(item => ({
                targetLocation: String(item?.targetLocation || "").trim(),
                suggestionText: sanitizeMarkdownText(item?.suggestionText || "").trim(),
                confidence: Number(item?.confidence) || 0
            }))
            .filter(item => item.targetLocation && item.suggestionText && item.confidence > 0.5)
            .sort((a, b) => b.confidence - a.confidence)
            .slice(0, 10);
    
    debugLog(`routing.done suggestions=${suggestions.length}`);
    return suggestions;
}

async function insertChatResponseBySearchAndInsert(aiText) {
    debugLog(`searchInsert.start aiLen=${aiText.length}`);
    setStatus("🔍 ניתוח מיקומים במסמך...");
    
    const docText = await getDocumentText();
    debugLog(`searchInsert.docText len=${docText.length}`);
    const routing = await buildAIRoutingMap(aiText, docText);
    if (!routing.length) {
        throw new Error("לא זוהו מיקומים ברורים במסמך עבור ההצעות");
    }
    
    debugLog(`searchInsert.routing count=${routing.length} targets=${routing.map(r => r.targetLocation.slice(0, 30)).join(" | ")}`);
    
    setStatus("🔎 חיפוש ומיפוי בטקסט...");
    let insertedCount = 0;
    
    await Word.run(async (ctx) => {
        ctx.document.load("changeTrackingMode");
        await ctx.sync();
        
        const originalMode = ctx.document.changeTrackingMode;
        try {
            ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await ctx.sync();
            
            for (const route of routing) {
                try {
                    debugLog(`searchInsert.searching target='${route.targetLocation.slice(0, 50)}' conf=${route.confidence.toFixed(2)}`);
                    
                    // חיפוש במסמך
                    const results = ctx.document.body.search(route.targetLocation, { matchCase: false });
                    results.load("items");
                    await ctx.sync();
                    
                    if (!results.items || results.items.length === 0) {
                        debugLog(`searchInsert.notFound target='${route.targetLocation.slice(0, 50)}'`);
                        continue;
                    }
                    
                    // הכנסת ההצעה בעקבות המיקום הראשון שנמצא
                    const targetRange = results.items[0];
                    const insertionText = `\n\n[הצעה AI: ${route.suggestionText}]\n`;
                    
                    const insertedRange = targetRange.insertText(insertionText, Word.InsertLocation.after);
                    await ctx.sync();
                    
                    // עיצוב ההצעה: כחול וברור
                    insertedRange.paragraphs.load("items");
                    await ctx.sync();
                    insertedRange.paragraphs.items.forEach(para => {
                        try {
                            para.font.color = "0070C0";
                            para.font.italic = true;
                            para.paragraphFormat.alignment = Word.Alignment.right;
                            para.paragraphFormat.readingOrder = Word.ReadingOrder.rtl;
                        } catch (fmtErr) {}
                    });
                    await ctx.sync();
                    
                    insertedCount += 1;
                    debugLog(`searchInsert.inserted target='${route.targetLocation.slice(0, 50)}' suggLen=${route.suggestionText.length}`);
                    
                } catch (itemErr) {
                    debugLog(`searchInsert.itemError target='${route.targetLocation.slice(0, 50)}' error=${itemErr?.message || String(itemErr)}`);
                }
            }
        } finally {
            ctx.document.changeTrackingMode = originalMode;
            await ctx.sync();
        }
    });
    
    debugLog(`searchInsert.done insertedCount=${insertedCount}`);
    if (!insertedCount) {
        throw new Error("לא הצליח למצוא מיקומים תואמים במסמך");
    }
    return insertedCount;
}

// ─── Legacy section insertion (keep for backwards compatibility) ────────────
function sanitizeMarkdownText(text) {
    return String(text || "")
        .replace(/```[\s\S]*?```/g, (block) => block.replace(/```/g, ""))
        .replace(/^\s{0,3}#{1,6}\s+/gm, "")
        .replace(/\*\*([^*]+)\*\*/g, "$1")
        .replace(/__([^_]+)__/g, "$1")
        .replace(/(?<!\*)\*([^*]+)\*(?!\*)/g, "$1")
        .replace(/(?<!_)_([^_]+)_(?!_)/g, "$1")
        .replace(/`([^`]+)`/g, "$1")
        .replace(/\[(.*?)\]\((.*?)\)/g, "$1")
    .replace(/\n{3,}/g, "\n\n");
}

async function insertTextAsTracked(text) {
    const cleanText = sanitizeMarkdownText(text);
    await Word.run(async (ctx) => {
        ctx.document.load("changeTrackingMode");
        await ctx.sync();
        const orig = ctx.document.changeTrackingMode;
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await ctx.sync();
        const range = ctx.document.getSelection();
        const insertedRange = range.insertText(cleanText, Word.InsertLocation.after);
        await ctx.sync();
        ctx.document.changeTrackingMode = orig;
        await ctx.sync();
        insertedRange.paragraphs.load("items");
        await ctx.sync();
        insertedRange.paragraphs.items.forEach(para => {
            para.paragraphFormat.alignment = Word.Alignment.right;
            para.paragraphFormat.readingOrder = Word.ReadingOrder.rtl;
        });
        await ctx.sync();
    });
}

async function insertTextDirect(text) {
    const cleanText = sanitizeMarkdownText(text);
    await Word.run(async (ctx) => {
        ctx.document.load("changeTrackingMode");
        await ctx.sync();
        const orig = ctx.document.changeTrackingMode;
        if (orig !== Word.ChangeTrackingMode.off) {
            ctx.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await ctx.sync();
        }
        const range = ctx.document.getSelection();
        const insertedRange = range.insertText(cleanText, Word.InsertLocation.after);
        insertedRange.paragraphs.load("items");
        await ctx.sync();
        insertedRange.paragraphs.items.forEach(para => {
            para.paragraphFormat.alignment = Word.Alignment.right;
            para.paragraphFormat.readingOrder = Word.ReadingOrder.rtl;
        });
        await ctx.sync();
        if (orig !== Word.ChangeTrackingMode.off) {
            ctx.document.changeTrackingMode = orig;
            await ctx.sync();
        }
    });
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

function normalizeHeadingForCompare(text) {
    return normalizeForCompare(text)
        .replace(/[.:]+$/g, "")
        .replace(/^[-–—•\s]+/g, "");
}

function isLikelySectionHeading(text) {
    const trimmed = String(text || "").trim();
    if (!trimmed || trimmed.length > 120) return false;
    // מספור מפורש (1. / א. / I. וכו')
    if (/^[\dIVXא-ת]+[.)\-]\s+/.test(trimmed)) return true;
    // מילות מפתח אקדמיות
    if (/^(פרק|סעיף|חלק|מבוא|דיון|מסקנות|סיכום|שאלה|טענה|רקע|שיטה|ממצאים|ניתוח|סקירת ספרות|מסגרת תיאורטית|שיטת המחקר|הקדמה|רשימת מקורות|ביבליוגרפיה|נספח|תקציר)\b/.test(trimmed)) return true;
    // נקודתיים בסוף — סימן קלאסי לכותרת
    if (/[:：]$/.test(trimmed)) return true;
    // משפט שלם — לא כותרת
    if (/[.!?]$/.test(trimmed)) return false;
    // שורה קצרה ללא תוכן מאוד ארוך — ייתכן כותרת
    const words = trimmed.split(/\s+/).filter(Boolean);
    if (words.length >= 1 && words.length <= 10) return true;
    return false;
}

async function getDocumentSections() {
    let sections = [];
    await Word.run(async (ctx) => {
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items/text");
        await ctx.sync();
        const items = paragraphs.items;
        for (let i = 0; i < items.length; i++) {
            const title = String(items[i].text || "").trim();
            if (!title || !isLikelySectionHeading(title)) continue;
            // אוסף עד 400 תווים של תוכן שאחרי הכותרת
            let content = "";
            for (let j = i + 1; j < items.length && content.length < 400; j++) {
                const line = String(items[j].text || "").trim();
                if (!line) continue;
                if (isLikelySectionHeading(line)) break;
                content += (content ? " " : "") + line;
            }
            sections.push({ index: i, title, content: content.slice(0, 400) });
        }
    });
    return sections;
}

function extractJsonObject(text) {
    const raw = String(text || "").trim();
    if (!raw) throw new Error("לא התקבל JSON תקין מה-AI");

    const direct = parseJsonSafe(raw);
    if (direct && typeof direct === "object") return direct;

    const fencedMatch = raw.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
    if (fencedMatch?.[1]) {
        const fenced = parseJsonSafe(fencedMatch[1]);
        if (fenced && typeof fenced === "object") return fenced;
    }

    const firstBrace = raw.indexOf("{");
    const firstBracket = raw.indexOf("[");
    const rootIsArray = firstBracket !== -1 && (firstBrace === -1 || firstBracket < firstBrace);

    if (rootIsArray) {
        for (let end = raw.lastIndexOf("]"); end > firstBracket; end = raw.lastIndexOf("]", end - 1)) {
            const candidate = raw.slice(firstBracket, end + 1);
            const parsed = parseJsonSafe(candidate);
            if (parsed && typeof parsed === "object") return parsed;
        }
    }

    if (firstBrace !== -1) {
        for (let end = raw.lastIndexOf("}"); end > firstBrace; end = raw.lastIndexOf("}", end - 1)) {
            const candidate = raw.slice(firstBrace, end + 1);
            const parsed = parseJsonSafe(candidate);
            if (parsed && typeof parsed === "object") return parsed;
        }
    }

    if (!rootIsArray && firstBracket !== -1) {
        for (let end = raw.lastIndexOf("]"); end > firstBracket; end = raw.lastIndexOf("]", end - 1)) {
            const candidate = raw.slice(firstBracket, end + 1);
            const parsed = parseJsonSafe(candidate);
            if (parsed && typeof parsed === "object") return parsed;
        }
    }

    throw new Error("לא התקבל JSON תקין מה-AI");
}

function normalizeSectionSuggestedText(text) {
    return sanitizeMarkdownText(text)
        .replace(/^\s*(?:[-–—*•]|(?:\d{1,2}|[א-ת]|[IVXivx]{1,6})[.)])\s+/gm, "")
        .replace(/^\s*#{1,6}\s+/gm, "")
        .replace(/\n{3,}/g, "\n\n")
        .trim();
}

function mergeSectionPlanItems(items) {
    const byHeading = new Map();
    for (const item of items || []) {
        const heading = String(item?.targetHeading || "").trim();
        const headingKey = normalizeHeadingForCompare(heading);
        const normalizedText = normalizeSectionSuggestedText(item?.suggestedText || "");
        if (!heading || !headingKey || !normalizedText) continue;
        if (!byHeading.has(headingKey)) {
            byHeading.set(headingKey, { targetHeading: heading, chunks: [] });
        }
        byHeading.get(headingKey).chunks.push(normalizedText);
    }
    return [...byHeading.values()].map(entry => ({
        targetHeading: entry.targetHeading,
        suggestedText: entry.chunks.join("\n\n")
    }));
}

function extractParagraphs(text) {
    return String(text || "")
        .split(/\n{2,}/)
        .map(p => p.trim())
        .filter(p => p.length > 10);
}

function extractFirstSentence(text) {
    const trimmed = String(text || "").trim();
    if (!trimmed) return "";
    const match = trimmed.match(/^[^.!?]*[.!?]/);
    if (!match) return trimmed.split("\n")[0].slice(0, 150).trim();
    return match[0].slice(0, 150).trim();
}

function findBestHeadingMatch(proposedHeading, sections) {
    if (!proposedHeading || !sections?.length) return -1;
    const proposed = normalizeHeadingForCompare(proposedHeading);
    if (!proposed) return -1;

    let bestMatch = { index: -1, score: 0 };

    for (let i = 0; i < sections.length; i++) {
        const existing = normalizeHeadingForCompare(sections[i].title);
        if (!existing) continue;
        if (existing === proposed) {
            return sections[i].index;
        }
        const commonWords = existing.split(/\s+/).filter(w => proposed.includes(w) && w.length > 2);
        const score = commonWords.length;
        if (score > bestMatch.score) {
            bestMatch = { index: sections[i].index, score };
        }
    }

    return bestMatch.score >= 2 ? bestMatch.index : -1;
}

async function buildSectionInsertionPlanFromParagraphs(aiText, sections) {
    debugLog(`sections.paragraph.start aiLen=${aiText.length} sections=${sections.length}`);
    const paragraphs = extractParagraphs(aiText);
    debugLog(`sections.paragraph.extracted count=${paragraphs.length}`);

    const plan = [];
    for (let i = 0; i < paragraphs.length; i++) {
        const para = paragraphs[i];
        const firstSent = extractFirstSentence(para);
        const matchedIndex = findBestHeadingMatch(firstSent, sections);

        if (matchedIndex >= 0) {
            const section = sections.find(s => s.index === matchedIndex);
            if (section) {
                plan.push({
                    targetHeading: section.title,
                    suggestedText: para,
                    fromParagraph: true
                });
                debugLog(`sections.paragraph.match i=${i} heading='${section.title}' sentLen=${firstSent.length}`);
                continue;
            }
        }

        plan.push({
            targetHeading: firstSent || `סעיף ${i + 1}`,
            suggestedText: para,
            fromParagraph: true,
            isNew: true
        });
        debugLog(`sections.paragraph.new i=${i} heading='${firstSent}' isNew=true`);
    }

    debugLog(`sections.paragraph.done planCount=${plan.length} mapped=${plan.filter(p => !p.isNew).length} new=${plan.filter(p => p.isNew).length}`);
    return plan;
}

async function buildSectionInsertionPlan(aiText, sections) {
    const apiKey = getApiKey();
    if (!apiKey) throw new Error("נדרש מפתח Gemini כדי למפות תשובה לסעיפים במסמך");
    debugLog(`sections.map.start sections=${sections.length} aiLen=${aiText.length} selLen=${currentSelectedText.length}`);

    const sectionList = sections.map(s =>
        `- כותרת: "${s.title}"${s.content ? `\n  תוכן נוכחי: "${s.content}"` : ""}`
    ).join("\n");
    const mappingPrompt = `You map a proposed Hebrew revision into an existing Hebrew academic document.
Return ONLY valid JSON with this exact shape:
{
  "sections": [
    {
      "targetHeading": "exact heading from the provided list",
      "suggestedText": "the specific revised or added text that belongs under that heading"
    }
  ]
}

CRITICAL Rules:
- Use ONLY headings from the provided list, exactly as written.
- Do NOT invent headings.
- Do NOT include headings inside suggestedText.
- suggestedText must be plain paragraphs in Hebrew (no bullet points, no numbered list markers).
- Match each part of the assistant response to the section whose CURRENT CONTENT it most directly addresses or improves.
- If the assistant response addresses multiple sections, return multiple entries — one per section.
- suggestedText must be the portion of the response relevant to THAT section only — not the full response.
- Return up to 8 relevant sections.
- If no section match is clear, return {"sections":[]}.

Document sections (heading + current content preview):
${sectionList}

Selected text context:
${currentSelectedText ? currentSelectedText.slice(0, 600) : "none"}

Assistant response to map:
${aiText.slice(0, 8000)}`;

    debugLog(`sections.map.prompt promptLen=${mappingPrompt.length} sectionListLen=${sectionList.length}`);

    const response = await callGeminiText(apiKey, mappingPrompt);
    const parsed = extractJsonObject(response);
    const mappedRaw = Array.isArray(parsed.sections)
        ? parsed.sections
            .map(item => ({
                targetHeading: String(item?.targetHeading || "").trim(),
                suggestedText: normalizeSectionSuggestedText(item?.suggestedText || "")
            }))
            .filter(item => item.targetHeading && item.suggestedText)
        : [];
    const mapped = mergeSectionPlanItems(mappedRaw);
    debugLog(`sections.map.done mapped=${mapped.length}`);
    if (mapped.length) {
        debugLog(`sections.map.targets ${mapped.slice(0, 4).map(m => m.targetHeading).join(" | ")}`);
    }
    return mapped;
}

function findMatchingParagraphIndex(paragraphs, targetHeading, usedIndices = new Set()) {
    const target = normalizeHeadingForCompare(targetHeading);
    if (!target) return -1;

    // התאמה מדויקת
    for (let index = 0; index < paragraphs.length; index += 1) {
        if (usedIndices.has(index)) continue;
        if (!isLikelySectionHeading(paragraphs[index].text)) continue;
        if (normalizeHeadingForCompare(paragraphs[index].text) === target) return index;
    }
    // התאמה חלקית — הכותרת מכילה את המחרוזת המבוקשת או להיפך
    for (let index = 0; index < paragraphs.length; index += 1) {
        if (usedIndices.has(index)) continue;
        if (!isLikelySectionHeading(paragraphs[index].text)) continue;
        const normalized = normalizeHeadingForCompare(paragraphs[index].text);
        if (!normalized) continue;
        if (normalized.includes(target) || target.includes(normalized)) return index;
    }
    return -1;
}

function findSectionIndexByHeading(sections, targetHeading) {
    const target = normalizeHeadingForCompare(targetHeading);
    if (!target) return -1;

    for (const section of sections || []) {
        if (normalizeHeadingForCompare(section.title) === target) return section.index;
    }

    const partialCandidates = [];
    for (const section of sections || []) {
        const normalized = normalizeHeadingForCompare(section.title);
        if (!normalized) continue;
        if (normalized.includes(target) || target.includes(normalized)) {
            partialCandidates.push(section.index);
        }
    }
    return partialCandidates.length === 1 ? partialCandidates[0] : -1;
}

async function applySectionSuggestionsAsTracked(suggestions, sections) {
    let insertedCount = 0;
    const newSections = [];
    
    await Word.run(async (ctx) => {
        ctx.document.load("changeTrackingMode");
        let paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items/text");
        await ctx.sync();

        const originalMode = ctx.document.changeTrackingMode;
        try {
            ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await ctx.sync();

            debugLog(`sections.apply.start suggestionsCount=${suggestions.length}`);
            for (const suggestion of suggestions) {
                if (!suggestion || !suggestion.targetHeading) {
                    debugLog(`sections.apply.skip invalid-suggestion heading=${suggestion?.targetHeading || "EMPTY"}`);
                    continue;
                }
                
                const cleanText = normalizeSectionSuggestedText(suggestion.suggestedText);
                if (!cleanText) {
                    debugLog(`sections.apply.skip heading='${suggestion.targetHeading}' reason=empty-text`);
                    continue;
                }
                debugLog(`sections.apply.process heading='${suggestion.targetHeading}' isNew=${suggestion.isNew} cleanLen=${cleanText.length}`);

                if (suggestion.isNew) {
                    newSections.push({
                        title: suggestion.targetHeading,
                        text: cleanText
                    });
                    debugLog(`sections.apply.queue-new heading='${suggestion.targetHeading}' textLen=${cleanText.length}`);
                    continue;
                }

                // חיפוש דינמי בפסקאות הנטענות כעת (במקום שימוש בsections array עם indices מיושנות)
                const target = normalizeHeadingForCompare(suggestion.targetHeading);
                let foundIndex = -1;
                for (let i = 0; i < paragraphs.items.length; i++) {
                    if (normalizeHeadingForCompare(paragraphs.items[i].text) === target) {
                        foundIndex = i;
                        break;
                    }
                }
                
                if (foundIndex === -1) {
                    for (let i = 0; i < paragraphs.items.length; i++) {
                        const normalized = normalizeHeadingForCompare(paragraphs.items[i].text);
                        if (normalized && (normalized.includes(target) || target.includes(normalized))) {
                            foundIndex = i;
                            break;
                        }
                    }
                }

                if (foundIndex === -1) {
                    debugLog(`sections.apply.skip heading='${suggestion.targetHeading}' reason=no-match-in-current`);
                    continue;
                }

                try {
                    const targetPara = paragraphs.items[foundIndex];
                    const insertText = `\n${cleanText}`;
                    debugLog(`sections.apply.insert.start heading='${suggestion.targetHeading}' idx=${foundIndex} cleanTextLen=${cleanText.length} insertTextLen=${insertText.length}`);
                    
                    const insertedRange = targetPara.insertText(insertText, Word.InsertLocation.after);
                    await ctx.sync();
                    
                    debugLog(`sections.apply.insert.after-insert heading='${suggestion.targetHeading}' idx=${foundIndex}`);
                    insertedRange.paragraphs.load("items");
                    await ctx.sync();
                    
                    insertedRange.paragraphs.items.forEach(para => {
                        try {
                            para.paragraphFormat.alignment = Word.Alignment.right;
                            para.paragraphFormat.readingOrder = Word.ReadingOrder.rtl;
                        } catch (fmtErr) {
                            debugLog(`sections.apply.format.error ${fmtErr?.message || String(fmtErr)}`);
                        }
                    });
                    insertedCount += 1;
                    debugLog(`sections.apply.ok heading='${suggestion.targetHeading}' idx=${foundIndex} textLen=${cleanText.length}`);
                    
                    paragraphs = ctx.document.body.paragraphs;
                    paragraphs.load("items/text");
                    await ctx.sync();
                } catch (insertErr) {
                    debugLog(`sections.apply.insert.error heading='${suggestion.targetHeading}' idx=${foundIndex} error=${insertErr?.message || String(insertErr)}`);
                    throw insertErr;
                }
            }

            if (newSections.length > 0) {
                try {
                    debugLog(`sections.apply.newSections.start count=${newSections.length}`);
                    const body = ctx.document.body;
                    
                    for (const newSec of newSections) {
                        try {
                            debugLog(`sections.apply.newSec.start heading='${newSec.title}' textLen=${newSec.text.length}`);
                            
                            // Reload paragraphs collection before getLast()
                            body.paragraphs.load("items");
                            await ctx.sync();
                            const lastPara = body.paragraphs.getLast();
                            if (!lastPara) {
                                debugLog(`sections.apply.newSec.error heading='${newSec.title}' reason=no-last-para`);
                                continue;
                            }
                            
                            // הוספת שורה ריקה + כותרת
                            const headingRange = lastPara.insertText(`\n\n${newSec.title}`, Word.InsertLocation.after);
                            headingRange.paragraphs.load("items");
                            await ctx.sync();
                            headingRange.paragraphs.items.forEach(p => {
                                try {
                                    p.font.bold = true;
                                    p.paragraphFormat.alignment = Word.Alignment.right;
                                    p.paragraphFormat.readingOrder = Word.ReadingOrder.rtl;
                                } catch (e) {}
                            });
                            await ctx.sync();
                            debugLog(`sections.apply.newSec.heading added='${newSec.title}'`);
                            
                            // Reload for next insertion
                            body.paragraphs.load("items");
                            await ctx.sync();
                            const bodyLastPara = body.paragraphs.getLast();
                            if (!bodyLastPara) {
                                debugLog(`sections.apply.newSec.error heading='${newSec.title}' reason=no-last-para-after-heading`);
                                continue;
                            }
                            
                            // הוספת תוכן
                            const contentRange = bodyLastPara.insertText(`\n${newSec.text}`, Word.InsertLocation.after);
                            contentRange.paragraphs.load("items");
                            await ctx.sync();
                            contentRange.paragraphs.items.forEach(p => {
                                try {
                                    p.paragraphFormat.alignment = Word.Alignment.right;
                                    p.paragraphFormat.readingOrder = Word.ReadingOrder.rtl;
                                } catch (e) {}
                            });
                            await ctx.sync();
                            
                            insertedCount += 1;
                            debugLog(`sections.apply.newSec.ok heading='${newSec.title}' textLen=${newSec.text.length}`);
                        } catch (innerErr) {
                            debugLog(`sections.apply.newSec.error heading='${newSec.title}' error=${innerErr?.message || String(innerErr)}`);
                        }
                    }
                    debugLog(`sections.apply.newSections.done count=${newSections.length}`);
                } catch (newSecErr) {
                    debugLog(`sections.apply.newSections.error ${newSecErr?.message || String(newSecErr)}`);
                }
            }
        } finally {
            ctx.document.changeTrackingMode = originalMode;
            await ctx.sync();
        }
    });

    if (!insertedCount) {
        throw new Error("לא מצאתי סעיפים תואמים במסמך עבור ההצעה הזאת");
    }
    return insertedCount;
}

async function insertChatResponseBySections(aiText) {
    debugLog(`sections.insert.start aiLen=${String(aiText || "").length}`);
    setStatus("🧩 מאתר סעיפים במסמך...");
    const sections = await getDocumentSections();
    debugLog(`sections.insert.detected count=${sections.length}`);
    if (sections.length) {
        debugLog(`sections.insert.headings ${sections.slice(0, 6).map(s => s.title).join(" | ")}`);
    }
    if (!sections.length) {
        throw new Error("לא נמצאו סעיפים ברורים במסמך. כדאי להשתמש בכותרות קצרות או מספור");
    }

    setStatus("🧩 מחלק לפסקאות ומשייך לסעיפים...");
    let plan = await buildSectionInsertionPlanFromParagraphs(aiText, sections);
    debugLog(`sections.insert.planCount ${plan.length}`);
    
    if (!plan.length) {
        debugLog(`sections.insert.fallbackToGemini reason=no-paragraphs`);
        const apiKey = getApiKey();
        if (!apiKey) {
            throw new Error("לא הוכנסה תשובה - נדרש מפתח Gemini לניתוח מתקדם");
        }
        setStatus("🧩 ממפה את תשובת ה-AI לסעיפים...");
        plan = await buildSectionInsertionPlan(aiText, sections);
        debugLog(`sections.insert.fallbackResult planCount=${plan.length}`);
    }
    
    if (!plan.length) {
        throw new Error("לא התקבלה מפת סעיפים ישימה מהתשובה הזאת");
    }

    setStatus("🧩 מכניס הצעות למסמך כ-Track Changes...");
    const insertedCount = await applySectionSuggestionsAsTracked(plan, sections);
    debugLog(`sections.insert.done inserted=${insertedCount}`);
    setStatus(`הושלם! נוספו ${insertedCount} שינויים לפי סעיפים`);
    return insertedCount;
}

function chunkText(text, maxChars = 8000) {
    if (text.length <= maxChars) return [text];
    const chunks = [];
    let start = 0;
    while (start < text.length) {
        let end = start + maxChars;
        if (end >= text.length) { chunks.push(text.slice(start)); break; }
        const paraBreak = text.lastIndexOf("\n", end);
        const sentBreak = text.lastIndexOf(". ", end);
        const breakAt = paraBreak > start + maxChars * 0.5 ? paraBreak
                      : sentBreak > start + maxChars * 0.5 ? sentBreak + 1 : end;
        chunks.push(text.slice(start, breakAt).trim());
        start = breakAt;
    }
    return chunks.filter(c => c.length > 0);
}

// ─── File Extraction ──────────────────────────────────────────────────────
async function extractTextFromFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "pdf") return extractPdfText(file);
    if (ext === "docx") return extractDocxText(file);
    if (ext === "doc") throw new Error("קובץ Word מסוג .doc לא נתמך ישירות. שמור אותו כ-.docx ונסה שוב.");
    throw new Error(`סוג קובץ לא נתמך: .${ext}. השתמש ב-PDF או DOCX.`);
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

// ─── DOI Validation (background) ─────────────────────────────────────────
async function validateCitationsInText(text, statusElId) {
    const DOI_RE = /\b10\.\d{4,}\/\S+/g;
    const dois = [...new Set((text.match(DOI_RE) || []).map(d => d.replace(/[.,;:)\]"']+$/, "")))];
    if (!dois.length) return;
    const statusEl = document.getElementById(statusElId);
    if (statusEl) statusEl.innerText = `מאמת ${dois.length} DOI...`;
    const timeout = typeof AbortSignal.timeout === "function" ? 5000 : undefined;
    const results = await Promise.allSettled(
        dois.map(doi => {
            const signal = timeout ? AbortSignal.timeout(timeout) : undefined;
            return fetch(`https://api.crossref.org/works/${doi.split("/").map(encodeURIComponent).join("/")}`, { signal })
                .then(r => ({ doi, ok: r.ok }))
                .catch(() => ({ doi, ok: null }));
        })
    );
    const failed = results.filter(r => r.status === "fulfilled" && r.value.ok === false).map(r => r.value.doi);
    const unchecked = results.filter(r => r.status === "fulfilled" && r.value.ok === null).length;
    if (!statusEl) return;
    if (failed.length === 0 && unchecked === 0) statusEl.innerText = `✅ כל ${dois.length} ה-DOI תקינים`;
    else if (failed.length > 0)                 statusEl.innerText = `⚠️ ${failed.length} DOI לא תקינים: ${failed.join(", ")}`;
    else                                        statusEl.innerText = `ℹ️ לא ניתן לאמת ${unchecked} DOI`;
}

// ─── Utilities ────────────────────────────────────────────────────────────
function setStatus(msg) { const el = document.getElementById("status"); if (el) el.innerText = msg; }

function showToast(message) {
    const toast = document.getElementById("insertToast");
    if (!toast) return;
    toast.innerText = message;
    toast.classList.remove("hidden");
    setTimeout(() => toast.classList.add("hidden"), 2500);
}

function escapeHtml(str) {
    return (str || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ─── RoamingSettings helpers ──────────────────────────────────────────────
function rsGet(key) {
    try {
        const val = Office.context.roamingSettings.get(key);
        if (val !== undefined && val !== null) return val;
    } catch(e) { /* fallback */ }
    return localStorage.getItem(key);
}

function rsSet(key, value) {
    try {
        Office.context.roamingSettings.set(key, value);
        Office.context.roamingSettings.saveAsync(result => {
            if (result.status !== Office.AsyncResultStatus.Succeeded)
                console.warn("roamingSettings save failed:", result.error?.message);
        });
    } catch(e) { /* fallthrough */ }
    localStorage.setItem(key, value);
}

function rsRemove(key) {
    try { Office.context.roamingSettings.remove(key); Office.context.roamingSettings.saveAsync(); }
    catch(e) { /* fallthrough */ }
    localStorage.removeItem(key);
}

// ─── API Key Management ───────────────────────────────────────────────────
function saveApiKey() {
    const geminiKey = document.getElementById("apiKey").value.trim();
    const perplexityKey = document.getElementById("perplexityApiKey").value.trim();
    if (geminiKey) rsSet("geminiApiKey", geminiKey); else rsRemove("geminiApiKey");
    if (perplexityKey) rsSet("perplexityApiKey", perplexityKey); else rsRemove("perplexityApiKey");
    document.getElementById("keyStatus").innerText = "✅ נשמר!";
    setTimeout(() => document.getElementById("settingsPanel").classList.add("hidden"), 1500);
}

function loadApiKey() {
    const gk = rsGet("geminiApiKey");
    if (gk) document.getElementById("apiKey").value = gk;
    const pk = rsGet("perplexityApiKey");
    if (pk) document.getElementById("perplexityApiKey").value = pk;
    if (gk || pk) document.getElementById("keyStatus").innerText = "מפתחות טעונים";
}

function getApiKey() { return rsGet("geminiApiKey") || document.getElementById("apiKey").value.trim(); }
function getPerplexityApiKey() { return rsGet("perplexityApiKey") || document.getElementById("perplexityApiKey").value.trim(); }

// ─── Study Materials ──────────────────────────────────────────────────────
function setCloudLoading(isLoading, message = "") {
    const el = document.getElementById("cloudLoading");
    if (!el) return;
    el.innerText = message || "⏳ טוען נתוני ענן...";
    el.classList.toggle("hidden", !isLoading);
}

function setCloudAuthStatus(message, isError = false) {
    const el = document.getElementById("cloudAuthStatus");
    if (!el) return;
    el.innerText = message;
    el.style.color = isError ? "#b42318" : "";
}

function setCloudMaterialsStatus(message, isError = false) {
    const el = document.getElementById("cloudMaterialsStatus");
    if (!el) return;
    el.innerText = message;
    el.style.color = isError ? "#b42318" : "";
}

function safeCloudLog(action, details = {}) {
    debugLog(`cloud.${action} ${JSON.stringify(details).slice(0, 180)}`);
}

function normalizeMaterialFileName(material) {
    return material.fileName || material.name || material.title || material.id || "material";
}

function getMaterialExt(material) {
    const explicit = String(material.type || "").toLowerCase().trim();
    if (explicit === "application/pdf" || explicit === "pdf") return "pdf";
    if (explicit === "docx" || explicit.includes("word")) return "docx";
    if (explicit === "text" || explicit.startsWith("text/")) return "txt";

    const fileName = normalizeMaterialFileName(material);
    const parts = fileName.split(".");
    return parts.length > 1 ? parts.pop().toLowerCase() : "";
}

function toPublicDownloadUrl(rawUrl, fallbackObjectPath = "") {
    const trimmed = String(rawUrl || "").trim();
    if (!trimmed) return "";
    if (trimmed.startsWith("gs://")) {
        const noScheme = trimmed.slice(5);
        const slashAt = noScheme.indexOf("/");
        if (slashAt < 0) {
            if (!fallbackObjectPath) return "";
            const bucketOnly = noScheme.trim();
            if (!bucketOnly) return "";
            return `https://firebasestorage.googleapis.com/v0/b/${encodeURIComponent(bucketOnly)}/o/${encodeURIComponent(fallbackObjectPath)}?alt=media`;
        }
        if (slashAt === 0) return "";
        const bucket = noScheme.slice(0, slashAt);
        const objectPath = noScheme.slice(slashAt + 1) || fallbackObjectPath;
        if (!objectPath) return "";
        return `https://firebasestorage.googleapis.com/v0/b/${encodeURIComponent(bucket)}/o/${encodeURIComponent(objectPath)}?alt=media`;
    }
    return trimmed;
}

function inferFileNameFromUrl(url) {
    try {
        const u = new URL(url);
        const path = decodeURIComponent(u.pathname || "");
        const last = path.split("/").filter(Boolean).pop() || "material.pdf";
        if (last.includes(".")) return last;
        return `${last}.pdf`;
    } catch {
        return "material.pdf";
    }
}

async function importPublicCloudUrlIntoContext(rawUrl) {
    const input = String(rawUrl || "").trim();
    const isBucketOnlyGs = /^gs:\/\/[^/]+\/?$/.test(input);
    const sourceUrl = isBucketOnlyGs
        ? toPublicDownloadUrl(input, "index.json")
        : toPublicDownloadUrl(input);
    if (!sourceUrl) throw new Error("קישור לא תקין");

    if (/\.json(\?|$)/i.test(sourceUrl)) {
        const remote = await fetchMaterialsFromIndexUrl(sourceUrl, "storage");
        if (!remote.length) {
            if (isBucketOnlyGs) {
                throw new Error("לא נמצא index.json בשורש ה-bucket. הוסף קובץ אינדקס או ספק קישור ישיר לקובץ.");
            }
            throw new Error("לא נמצאו חומרים בקובץ האינדקס שסופק");
        }
        return { kind: "index", count: remote.length, materials: remote };
    }

    const fileName = inferFileNameFromUrl(sourceUrl);
    const material = {
        id: `storage-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        courseId: "storage",
        title: fileName,
        name: fileName,
        fileName,
        type: fileName.split(".").pop().toLowerCase(),
        directUrl: sourceUrl,
    };
    return { kind: "single", count: 1, materials: [material] };
}

async function syncStorageToLocal(source) {
    const response = await fetch("/api/storage-sync", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ source }),
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok || !payload?.ok) {
        throw new Error(payload?.error || "סנכרון לתיקייה המקומית נכשל");
    }
    return payload.data || { syncedFiles: [], count: 0, indexCount: 0 };
}

function materialSignature(material) {
    const name = String(material.fileName || material.title || material.id || "").trim().toLowerCase();
    const ext = String(getMaterialExt(material) || "").trim().toLowerCase();
    const sourceUrl = String(material.syncedFrom || material.directUrl || "").trim().toLowerCase();
    return `${name}::${ext}::${sourceUrl}`;
}

function areMaterialSetsDifferent(primary = [], secondary = []) {
    const a = new Set(primary.map(materialSignature));
    const b = new Set(secondary.map(materialSignature));
    if (a.size !== b.size) return true;
    for (const sig of a) {
        if (!b.has(sig)) return true;
    }
    return false;
}

function mergeMaterials(primary = [], secondary = []) {
    const merged = [];
    const seen = new Set();
    for (const material of [...primary, ...secondary]) {
        const sig = materialSignature(material);
        if (seen.has(sig)) continue;
        seen.add(sig);
        merged.push(material);
    }
    return merged;
}

function isLikelyProcessedMaterial(material) {
    const title = String(material.title || material.name || material.fileName || "").toLowerCase();
    const type = String(material.type || "").toLowerCase();
    const tags = Array.isArray(material.tags) ? material.tags.map((t) => String(t).toLowerCase()) : [];
    const source = `${title} ${type} ${tags.join(" ")}`;

    // מסנן סיכומים/נגזרות כדי להעדיף חומרי גלם מקוריים.
    const processedHints = [
        "summary", "summarized", "abstract", "flashcard", "processed", "generated",
        "סיכום", "תקציר", "תמצית", "כרטיסיות", "מעובד", "נגזר", "מחולל"
    ];
    return processedHints.some((hint) => source.includes(hint));
}

function filterRawCloudMaterials(materials = []) {
    const all = Array.isArray(materials) ? materials : [];
    return all.filter((material) => !isLikelyProcessedMaterial(material));
}

function getCloudCacheKey(material) {
    const courseId = material.courseId || "global";
    return `${courseId}:${material.id}`;
}

function normalizeFolderKey(value) {
    return String(value || "")
        .trim()
        .replace(/\\/g, "/")
        .replace(/^\/+|\/+$/g, "")
        .toLowerCase();
}

function extractFolderFromFilePath(filePath) {
    const normalized = String(filePath || "").trim().replace(/\\/g, "/");
    if (!normalized) return "";
    const parts = normalized.split("/").filter(Boolean);
    if (parts.length <= 1) return "";
    return parts.slice(0, -1).join("/");
}

function normalizeTocEntries(rawToc) {
    if (!Array.isArray(rawToc)) return [];
    const entries = [];
    for (const item of rawToc) {
        if (typeof item === "string") {
            const title = item.trim();
            if (title) entries.push({ title, id: "" });
            continue;
        }
        const title = String(item?.title || item?.label || item?.name || "").trim();
        if (!title) continue;
        entries.push({ title, id: String(item?.id || item?.anchor || "").trim() });
    }
    return entries;
}

function getMaterialCompositeId(material) {
    return `${material.courseId || ""}:${material.id}`;
}

function getFilteredCloudMaterials() {
    if (!cloudSelectedFolderKeys.size) return cloudMaterials;
    return cloudMaterials.filter((material) => cloudSelectedFolderKeys.has(material.folderKey || "root"));
}

function pruneCloudSelectedMaterialIds() {
    const available = new Set(cloudMaterials.map((material) => getMaterialCompositeId(material)));
    cloudSelectedMaterialIds = new Set([...cloudSelectedMaterialIds].filter((id) => available.has(id)));
}

function syncMaterialSelectionWithSelectedFolders() {
    if (!cloudSelectedFolderKeys.size) return;
    // שומר בחירות ידניות קיימות ומאמץ לפי תיקיות רק כשאין בחירה קודמת.
    if (cloudSelectedMaterialIds.size > 0) return;
    for (const material of cloudMaterials) {
        if (!cloudSelectedFolderKeys.has(material.folderKey || "root")) continue;
        cloudSelectedMaterialIds.add(getMaterialCompositeId(material));
    }
}

function loadCloudSelectionState() {
    try {
        const raw = rsGet(CLOUD_SELECTED_IDS_KEY);
        const parsed = raw ? JSON.parse(raw) : [];
        cloudSelectedMaterialIds = new Set(Array.isArray(parsed) ? parsed : []);
    } catch {
        cloudSelectedMaterialIds = new Set();
    }
}

function saveCloudSelectionState() {
    rsSet(CLOUD_SELECTED_IDS_KEY, JSON.stringify([...cloudSelectedMaterialIds]));
}

function loadCloudFolderSelectionState() {
    try {
        const raw = rsGet(CLOUD_SELECTED_FOLDER_KEYS_KEY);
        const parsed = raw ? JSON.parse(raw) : [];
        cloudSelectedFolderKeys = new Set(
            Array.isArray(parsed)
                ? parsed.map((key) => normalizeFolderKey(key || "root")).filter(Boolean)
                : []
        );
    } catch {
        cloudSelectedFolderKeys = new Set();
        rsSet(CLOUD_SELECTED_FOLDER_KEYS_KEY, "[]");
    }
}

function saveCloudFolderSelectionState() {
    rsSet(CLOUD_SELECTED_FOLDER_KEYS_KEY, JSON.stringify([...cloudSelectedFolderKeys]));
}

function renderCloudCourses() {
    // מצב ללא קורסים: הרשימה מגיעה מתיקיית הפרויקט.
}

function renderCloudMaterials() {
    const list = document.getElementById("cloudMaterialsList");
    const visibleMaterials = getFilteredCloudMaterials();
    if (!list) return;
    if (!cloudMaterials.length) {
        list.innerText = "לא נמצאו חומרים. אפשר להוסיף קבצים לתיקייה project-materials";
        renderCloudFolderFilters();
        renderCloudIndex();
        return;
    }

    if (!visibleMaterials.length) {
        list.innerText = "לא נמצאו חומרים בתיקיות שנבחרו";
        renderCloudFolderFilters();
        renderCloudIndex();
        return;
    }

    list.innerHTML = visibleMaterials.map((material) => {
        const id = getMaterialCompositeId(material);
        const checked = cloudSelectedMaterialIds.has(id) ? "checked" : "";
        const state = cloudMaterialLoadState.get(id) || "לא נטען";
        const cls = state.includes("שגיאה") ? "error" : state === "נטען" ? "ready" : "";
        const title = escapeHtml(material.title || material.name || material.id);
        const folderText = material.folderLabel && material.folderLabel !== "root" ? ` · ${material.folderLabel}` : "";
        const meta = escapeHtml(`${(material.type || getMaterialExt(material) || "file").toUpperCase()}${folderText}${material.updatedAt ? " · מעודכן" : ""}`);
        const toc = Array.isArray(material.tocEntries) && material.tocEntries.length
            ? `<div class="cloud-material-toc">אינדקס: ${material.tocEntries.map((entry) => escapeHtml(entry.title)).join(" • ")}</div>`
            : "";
        return `<label class="cloud-material-row">
            <input type="checkbox" class="cloud-material-check" data-material-id="${escapeHtml(id)}" ${checked} />
            <div>
                <div class="cloud-material-title">${title}</div>
                <div class="cloud-material-meta">${meta}</div>
                ${toc}
            </div>
            <div class="cloud-material-state ${cls}">${escapeHtml(state)}</div>
        </label>`;
    }).join("");
    renderCloudFolderFilters();
    renderCloudIndex();
}

function renderCloudFolderFilters() {
    const box = document.getElementById("cloudFolderFilters");
    if (!box) return;
    const byFolder = new Map();
    for (const material of cloudMaterials) {
        const key = material.folderKey || "root";
        const info = byFolder.get(key) || { label: material.folderLabel || "root", count: 0 };
        info.count += 1;
        byFolder.set(key, info);
    }

    if (!byFolder.size) {
        box.innerHTML = "";
        return;
    }

    box.innerHTML = [...byFolder.entries()]
        .sort((a, b) => a[1].label.localeCompare(b[1].label, "he"))
        .map(([key, info]) => {
            const checked = cloudSelectedFolderKeys.has(key) ? "checked" : "";
            return `<label class="cloud-folder-filter-row">
                <input type="checkbox" class="cloud-folder-filter-check" data-folder-key="${escapeHtml(key)}" ${checked} />
                <span>${escapeHtml(info.label)} (${info.count})</span>
            </label>`;
        })
        .join("");
}

function renderCloudIndex() {
    const box = document.getElementById("cloudIndexList");
    if (!box) return;
    const visibleMaterials = getFilteredCloudMaterials()
        .filter((material) => cloudSelectedMaterialIds.has(getMaterialCompositeId(material)));
    if (!visibleMaterials.length) {
        box.innerHTML = "אין עדיין תוכן להצגה";
        return;
    }

    const grouped = new Map();
    for (const material of visibleMaterials) {
        const key = material.folderKey || "root";
        if (!grouped.has(key)) grouped.set(key, { label: material.folderLabel || "root", items: [] });
        grouped.get(key).items.push(material);
    }

    box.innerHTML = [...grouped.values()]
        .sort((a, b) => a.label.localeCompare(b.label, "he"))
        .map((group) => {
            const items = group.items
                .sort((a, b) => String(a.title || "").localeCompare(String(b.title || ""), "he"))
                .map((material) => {
                    const tocItems = Array.isArray(material.tocEntries) && material.tocEntries.length
                        ? `<ul class="cloud-index-sublist">${material.tocEntries.map((entry) => `<li>${escapeHtml(entry.title)}</li>`).join("")}</ul>`
                        : "";
                    return `<li><span class="cloud-index-item-title">${escapeHtml(material.title || material.fileName || material.id)}</span>${tocItems}</li>`;
                })
                .join("");
            return `<details class="cloud-index-group" open>
                <summary>${escapeHtml(group.label)} (${group.items.length})</summary>
                <ul class="cloud-index-list-items">${items}</ul>
            </details>`;
        })
        .join("");
}

async function extractTextFromCloudBlob(blob, material) {
    const ext = getMaterialExt(material);
    const maxBytes = 12 * 1024 * 1024;
    if (blob.size > maxBytes) {
        throw new Error("הקובץ גדול מדי לחילוץ טקסט בצד לקוח");
    }
    if (ext === "pdf" || ext === "docx") {
        const file = new File([blob], normalizeMaterialFileName(material), { type: blob.type || "application/octet-stream" });
        return extractTextFromFile(file);
    }
    if (["txt", "md", "csv", "json", "html", "htm"].includes(ext) || (blob.type || "").startsWith("text/")) {
        return blob.text();
    }
    throw new Error(`סוג קובץ לא נתמך לחילוץ אוטומטי: ${ext || "unknown"}`);
}

async function cacheCloudMaterialText(material, forced = false) {
    const id = `${material.courseId || ""}:${material.id}`;
    if (!forced && cloudMaterialTextCache.has(id)) return cloudMaterialTextCache.get(id);

    cloudMaterialLoadState.set(id, "טוען...");
    renderCloudMaterials();
    safeCloudLog("material.load.start", { id });

    try {
        if (!material.directUrl) {
            throw new Error("לא נמצא קישור ישיר לקובץ");
        }
        const response = await fetch(material.directUrl);
        if (!response.ok) {
            throw new Error(`הקישור לא זמין (${response.status})`);
        }
        const blob = await response.blob();
        const text = String(await extractTextFromCloudBlob(blob, material) || "").trim();
        if (!text) throw new Error("לא זוהה טקסט בקובץ");
        cloudMaterialTextCache.set(id, text);
        cloudMaterialLoadState.set(id, "נטען");
        renderCloudMaterials();
        safeCloudLog("material.load.done", { id, len: text.length });
        return text;
    } catch (e) {
        cloudMaterialLoadState.set(id, `שגיאה: ${e.message}`);
        renderCloudMaterials();
        safeCloudLog("material.load.error", { id, msg: e.message });
        throw e;
    }
}

async function fetchProjectMaterials() {
    return fetchMaterialsFromIndexUrl(PROJECT_MATERIALS_INDEX_URL, "project");
}

async function fetchMaterialsFromIndexUrl(indexUrl, sourceKey = "project") {
    try {
        const urlWithBust = `${indexUrl}${indexUrl.includes("?") ? "&" : "?"}t=${Date.now()}`;
        const response = await fetch(urlWithBust);
        if (!response.ok) return [];
        const rows = await response.json();
        if (!Array.isArray(rows)) return [];
        const baseUrl = new URL(indexUrl, window.location.href);

        return rows
            .map((row, idx) => {
                const filePath = String(row.file || row.path || "").trim();
                const directUrlRaw = String(row.url || (filePath ? filePath : "")).trim();
                const directUrl = directUrlRaw
                    ? new URL(directUrlRaw, baseUrl).toString()
                    : "";
                if (!directUrl) return null;
                const fileName = filePath.split("/").pop() || `material-${idx + 1}`;
                const id = String(row.id || filePath || `material-${idx + 1}`);
                const rawFolder = String(row.folder || row.topic || row.course || extractFolderFromFilePath(filePath) || "root").trim();
                const folderKey = normalizeFolderKey(rawFolder || "root") || "root";
                return {
                    id: `${sourceKey}-${id}`,
                    courseId: sourceKey,
                    title: row.title || fileName,
                    fileName: row.fileName || fileName,
                    filePath,
                    type: row.type || "",
                    folderKey,
                    folderLabel: rawFolder || "root",
                    tocEntries: normalizeTocEntries(row.toc),
                    directUrl,
                    syncedFrom: row.syncedFrom || "",
                };
            })
            .filter(Boolean);
    } catch {
        return [];
    }
}

async function loadCloudCourses(force = false) {
    setCloudLoading(true, "⏳ טוען חומרים מתיקיית הפרויקט...");
    try {
        const fetchedMaterials = await fetchProjectMaterials();
        cloudMaterials = filterRawCloudMaterials(fetchedMaterials);
        pruneCloudSelectedMaterialIds();
        syncMaterialSelectionWithSelectedFolders();
        saveCloudSelectionState();
        renderCloudMaterials();
        if (!cloudMaterials.length) {
            setCloudMaterialsStatus("לא נמצאו חומרים בתיקייה project-materials (בדוק index.json)");
        } else {
            setCloudMaterialsStatus(`נטענו ${cloudMaterials.length} חומרים מתיקיית הפרויקט`);
        }
        safeCloudLog("project.materials.loaded", { count: cloudMaterials.length, force });
        await preloadSelectedCloudMaterials();
    } catch (e) {
        setCloudMaterialsStatus("שגיאה בטעינת חומרים מקומיים: " + e.message, true);
    } finally {
        setCloudLoading(false);
    }
}

async function loadCloudMaterials(courseId, force = false) {
    await loadCloudCourses(force);
}

async function onCloudMaterialSelectionChange(materialId, checked) {
    const material = cloudMaterials.find((m) => getMaterialCompositeId(m) === materialId);
    if (!material) return;

    if (checked) {
        cloudSelectedMaterialIds.add(materialId);
        saveCloudSelectionState();
        try {
            await cacheCloudMaterialText(material);
            setCloudMaterialsStatus(`החומר '${material.title || material.id}' נוסף כהקשר`);
        } catch (e) {
            cloudSelectedMaterialIds.delete(materialId);
            saveCloudSelectionState();
            setCloudMaterialsStatus(`לא ניתן לטעון '${material.title || material.id}': ${e.message}`, true);
        }
    } else {
        cloudSelectedMaterialIds.delete(materialId);
        saveCloudSelectionState();
        setCloudMaterialsStatus("עודכן סט חומרים להקשר");
    }
    renderCloudMaterials();
}

async function preloadSelectedCloudMaterials() {
    const selectedNow = cloudMaterials.filter((m) => cloudSelectedMaterialIds.has(getMaterialCompositeId(m)));
    if (!selectedNow.length) return;

    setCloudMaterialsStatus(`טוען ${selectedNow.length} חומרים מסומנים...`);
    const results = await Promise.allSettled(selectedNow.map((material) => cacheCloudMaterialText(material)));
    const failed = results.filter((r) => r.status === "rejected").length;
    if (failed) {
        setCloudMaterialsStatus(`נטענו ${selectedNow.length - failed}/${selectedNow.length} חומרים`, true);
    } else {
        setCloudMaterialsStatus(`נטענו ${selectedNow.length} חומרים מסומנים`);
    }
}

async function selectCloudMaterialsByFilePaths(filePaths = []) {
    if (!cloudMaterials.length || !filePaths.length) return;
    const wanted = new Set(filePaths.map((value) => String(value || "").trim().toLowerCase()));
    const selectedIds = cloudMaterials
        .filter((material) => wanted.has(String(material.filePath || material.fileName || "").trim().toLowerCase()))
        .map((material) => getMaterialCompositeId(material));
    if (!selectedIds.length) return;
    for (const id of selectedIds) cloudSelectedMaterialIds.add(id);
    saveCloudSelectionState();
    renderCloudMaterials();
    await preloadSelectedCloudMaterials();
}

function buildCloudStudyContext() {
    const blocks = [];
    for (const material of getFilteredCloudMaterials()) {
        const id = getMaterialCompositeId(material);
        if (!cloudSelectedMaterialIds.has(id)) continue;
        const cachedText = cloudMaterialTextCache.get(id);
        if (!cachedText) continue;
        const tocLine = Array.isArray(material.tocEntries) && material.tocEntries.length
            ? `\nאינדקס: ${material.tocEntries.map((entry) => entry.title).join(" | ")}`
            : "";
        blocks.push(`=== חומר: ${material.title || material.id} (${material.folderLabel || "root"}) ===${tocLine}\n${cachedText.slice(0, 5500)}`);
    }

    if (!blocks.length) return "";
    return `\n\nCloud Study Materials (retrieved from Firebase):\n${blocks.join("\n\n")}`;
}

function setupCloudEvents() {
    const materialsList = document.getElementById("cloudMaterialsList");
    const folderFilters = document.getElementById("cloudFolderFilters");
    const publicUrlInput = document.getElementById("cloudPublicUrl");
    const importUrlBtn = document.getElementById("cloudImportUrlBtn");
    const downloadToLocalToggle = document.getElementById("cloudDownloadToLocalToggle");

    if (downloadToLocalToggle) {
        const localSyncAvailable = !!import.meta.env.DEV && /localhost|127\.0\.0\.1/i.test(window.location.hostname || "");
        downloadToLocalToggle.disabled = !localSyncAvailable;
        if (!localSyncAvailable) {
            downloadToLocalToggle.checked = false;
            setCloudMaterialsStatus("הורדה לתיקייה המקומית זמינה רק בזמן npm run dev על localhost");
        }
    }

    if (materialsList && !materialsList.dataset.bound) {
        materialsList.addEventListener("change", async (e) => {
            const cb = e.target.closest(".cloud-material-check");
            if (!cb) return;
            await onCloudMaterialSelectionChange(cb.dataset.materialId, cb.checked);
        });
        materialsList.dataset.bound = "true";
    }

    if (folderFilters && !folderFilters.dataset.bound) {
        folderFilters.addEventListener("change", async (e) => {
            const cb = e.target.closest(".cloud-folder-filter-check");
            if (!cb) return;
            const folderKey = normalizeFolderKey(cb.dataset.folderKey || "root") || "root";
            if (cb.checked) {
                cloudSelectedFolderKeys.add(folderKey);
                for (const material of cloudMaterials) {
                    if ((material.folderKey || "root") !== folderKey) continue;
                    cloudSelectedMaterialIds.add(getMaterialCompositeId(material));
                }
                saveCloudSelectionState();
                saveCloudFolderSelectionState();
                renderCloudMaterials();
                await preloadSelectedCloudMaterials();
                setCloudMaterialsStatus("התיקייה נוספה להקשר");
                return;
            }

            cloudSelectedFolderKeys.delete(folderKey);
            for (const material of cloudMaterials) {
                if ((material.folderKey || "root") !== folderKey) continue;
                cloudSelectedMaterialIds.delete(getMaterialCompositeId(material));
            }
            saveCloudSelectionState();
            saveCloudFolderSelectionState();
            renderCloudMaterials();
            setCloudMaterialsStatus("התיקייה הוסרה מהקשר");
        });
        folderFilters.dataset.bound = "true";
    }

    if (importUrlBtn && !importUrlBtn.dataset.bound) {
        importUrlBtn.onclick = async () => {
            const rawUrl = (publicUrlInput?.value || "").trim();
            const shouldDownloadToLocal = !!downloadToLocalToggle?.checked;
            if (!rawUrl) {
                setCloudMaterialsStatus("אין קישור סטורג׳. נטען רק מתיקיית הפרויקט.");
                await loadCloudCourses(true);
                return;
            }

            setCloudLoading(true, shouldDownloadToLocal ? "⏳ מוריד לתיקייה המקומית..." : "⏳ טוען קודם מהסטורג׳...");
            try {
                if (shouldDownloadToLocal) {
                    const syncResult = await syncStorageToLocal(rawUrl);
                    await loadCloudCourses(true);
                    await selectCloudMaterialsByFilePaths(syncResult.syncedFiles || []);
                    setCloudMaterialsStatus(`הורדו ${syncResult.count || 0} חומרים לתיקייה המקומית ונטענו בהצלחה`);
                    if (publicUrlInput) publicUrlInput.value = "";
                    return;
                }

                const storageResult = await importPublicCloudUrlIntoContext(rawUrl);
                const storageMaterials = filterRawCloudMaterials(storageResult.materials || []);
                const projectMaterials = filterRawCloudMaterials(await fetchProjectMaterials());

                let finalMaterials = storageMaterials;
                if (projectMaterials.length && areMaterialSetsDifferent(storageMaterials, projectMaterials)) {
                    const includeProject = window.confirm("זוהו קבצים שונים בין הסטורג׳ לתיקיית הפרויקט. לגשת גם לתיקייה המקומית?");
                    if (includeProject) {
                        finalMaterials = mergeMaterials(storageMaterials, projectMaterials);
                    }
                }

                cloudMaterials = finalMaterials;
                pruneCloudSelectedMaterialIds();
                syncMaterialSelectionWithSelectedFolders();
                saveCloudSelectionState();
                renderCloudMaterials();
                await preloadSelectedCloudMaterials();
                setCloudMaterialsStatus(`נטענו ${cloudMaterials.length} חומרים (סטורג׳${cloudMaterials.length !== storageMaterials.length ? " + תיקייה" : " בלבד"})`);
                if (publicUrlInput) publicUrlInput.value = "";
            } catch (e) {
                setCloudMaterialsStatus(`שגיאה בטעינת קישור: ${e.message}`, true);
            } finally {
                setCloudLoading(false);
            }
        };
        importUrlBtn.dataset.bound = "true";
    }
}

function setupCloudAuthObserver() {
    setCloudAuthStatus("ללא התחברות");
    setCloudMaterialsStatus("אפשר לטעון מתיקיית project-materials או מקישור ציבורי");
    loadCloudCourses(true);
}

function getStudyContext() {
    const cloud = buildCloudStudyContext();
    let ctx = "";
    if (cloud) ctx += cloud;
    return ctx;
}

function getReviewDepthLabel(value) {
    if (value === "strict") return "מחמירה";
    if (value === "basic") return "בסיסית";
    return "רגילה";
}

function getAssignmentGuidelinesText() {
    const fromInput = (document.getElementById("assignmentGuidelines")?.value || "").trim();
    const fromFile = String(assignmentGuidelinesFromFile || "").trim();
    return fromInput || fromFile;
}

function getLecturerReviewContext(message = "") {
    const depth = document.getElementById("reviewDepth")?.value || rsGet(REVIEW_DEPTH_KEY) || "standard";
    const guidelines = getAssignmentGuidelinesText();
    const shouldApply = REVIEW_INTENT.test(String(message || ""));
    if (!shouldApply) return "";
    return `\n\n== LECTURER REVIEW MODE ==
You are grading a student's assignment.
Return your answer in Hebrew only.
Step 1 (mandatory): verify guideline coverage in the assignment text BEFORE grading.
Output this section first:
כותרת: "בדיקת כיסוי הנחיות"
1) סטטוס כיסוי: "מלא" או "חלקי"
2) הנחיות שזוהו בטקסט: רשימה קצרה
3) הנחיות חסרות/לא מזוהות: רשימה קצרה

If guidelines are missing or coverage is partial, STOP and end with:
"שלח את ההנחיות החסרות ואבצע בדיקת מרצה מלאה."
In that case DO NOT provide a grade yet.

Only if coverage is full, continue with structured feedback in this exact order:
1) ציון כללי (0-100)
2) מה עובד טוב (עד 3 נקודות)
3) בעיות עיקריות (מובחן לפי סעיפים/פסקאות כשאפשר)
4) הצעת שיפור קצרה לכל בעיה (משפט אחד לכל בעיה)
5) התראת AI: סמן רק אם יש חשד גבוה מאוד לניסוח שנוצר ב-AI.
Use one of these labels only:
- "ללא חשד משמעותי"
- "⚠️ חשד גבוה ל-AI"
If you mark "⚠️ חשד גבוה ל-AI", include a short reason tied to concrete text signals.
Never claim certainty; phrase as high-confidence suspicion.
Review strictness level: ${getReviewDepthLabel(depth)}.
${guidelines ? `Assignment guidelines/rubric to enforce:\n${guidelines.slice(0, 6000)}` : "No rubric supplied. Treat as missing guidelines and ask the user to provide them before grading."}
== END LECTURER REVIEW MODE ==`;
}

function getAgentRulesText() {
    return (document.getElementById("agentRules")?.value || rsGet(AGENT_RULES_KEY) || "").trim();
}

function getAgentRulesContext(rulesInput = "") {
    const rules = String(rulesInput || "").trim();
    if (!rules) return "";
    return `\n\n== USER WRITING RULES (HARD REQUIREMENTS) ==
These rules are mandatory.
If any stylistic instruction in this prompt conflicts with these rules, USER WRITING RULES win.
Apply them to wording, structure, tone, and output format.
${rules}
== END USER RULES ==`;
}

function isDebugEnabled() {
    const toggle = document.getElementById("debugModeToggle");
    if (toggle) return !!toggle.checked;
    return rsGet(DEBUG_MODE_KEY) === "1";
}

function renderDebugLog() {
    const logEl = document.getElementById("debugLog");
    if (!logEl) return;
    if (!isDebugEnabled()) {
        logEl.innerText = "דיבאג כבוי";
        return;
    }
    logEl.innerText = debugEvents.length ? debugEvents.join("\n") : "ממתין לאירועים...";
    logEl.scrollTop = logEl.scrollHeight;
}

function debugLog(message) {
    if (!isDebugEnabled()) return;
    const ts = new Date().toLocaleTimeString("he-IL", { hour12: false });
    const line = `[${ts}] ${message}`;
    debugEvents.push(line);
    debugEvents = debugEvents.slice(-120);
    console.log(`[WordAI][DEBUG] ${line}`);
    renderDebugLog();
}

function saveMaterials() {
    rsSet(AGENT_RULES_KEY, document.getElementById("agentRules")?.value || "");
    rsSet(ASSIGNMENT_GUIDELINES_KEY, document.getElementById("assignmentGuidelines")?.value || "");
    rsSet(REVIEW_DEPTH_KEY, document.getElementById("reviewDepth")?.value || "standard");
    rsSet(DEBUG_MODE_KEY, document.getElementById("debugModeToggle")?.checked ? "1" : "0");
    const statusEl = document.getElementById("materialsStatus");
    if (statusEl) statusEl.innerText = "✅ ההגדרות נשמרו";
    debugLog(`settings.saved rulesLen=${getAgentRulesText().length}`);
    setTimeout(() => document.getElementById("settingsPanel").classList.add("hidden"), 1500);
}

function loadMaterials() {
    const rules = rsGet(AGENT_RULES_KEY);
    const assignmentGuidelines = rsGet(ASSIGNMENT_GUIDELINES_KEY);
    const reviewDepth = rsGet(REVIEW_DEPTH_KEY);
    const fileName = rsGet(ASSIGNMENT_GUIDELINES_FILE_NAME_KEY);
    const debugEnabled = rsGet(DEBUG_MODE_KEY) === "1";
    if (rules) document.getElementById("agentRules").value = rules;
    if (assignmentGuidelines) document.getElementById("assignmentGuidelines").value = assignmentGuidelines;
    if (reviewDepth) document.getElementById("reviewDepth").value = reviewDepth;
    const guidelinesStatus = document.getElementById("assignmentGuidelinesStatus");
    if (guidelinesStatus && fileName) guidelinesStatus.innerText = `נטען לאחרונה: ${fileName}`;
    const quickGuidelinesStatus = document.getElementById("quickGuidelinesStatus");
    if (quickGuidelinesStatus && fileName) quickGuidelinesStatus.innerText = `קובץ הנחיות אחרון: ${fileName}`;
    const debugToggle = document.getElementById("debugModeToggle");
    if (debugToggle) debugToggle.checked = debugEnabled;
    renderDebugLog();
}

async function loadAssignmentGuidelinesFromFile(file) {
    const settingsStatusEl = document.getElementById("assignmentGuidelinesStatus");
    const quickStatusEl = document.getElementById("quickGuidelinesStatus");
    if (settingsStatusEl) settingsStatusEl.innerText = "⏳ קורא הנחיות...";
    if (quickStatusEl) quickStatusEl.innerText = "⏳ קורא קובץ הנחיות...";
    try {
        const text = await extractTextFromFile(file);
        assignmentGuidelinesFromFile = String(text || "").trim();
        const guidelinesInput = document.getElementById("assignmentGuidelines");
        if (guidelinesInput && assignmentGuidelinesFromFile) {
            guidelinesInput.value = assignmentGuidelinesFromFile;
        }
        if (assignmentGuidelinesFromFile) {
            rsSet(ASSIGNMENT_GUIDELINES_KEY, assignmentGuidelinesFromFile);
        }
        rsSet(ASSIGNMENT_GUIDELINES_FILE_NAME_KEY, file.name);
        if (settingsStatusEl) {
            settingsStatusEl.innerText = assignmentGuidelinesFromFile
                ? `✅ נטען ${file.name} (${assignmentGuidelinesFromFile.length} תווים)`
                : `⚠️ ${file.name} נטען אבל לא זוהה בו טקסט`;
        }
        if (quickStatusEl) {
            quickStatusEl.innerText = assignmentGuidelinesFromFile
                ? `✅ ההנחיות נטענו מ-${file.name}`
                : `⚠️ לא זוהה טקסט ב-${file.name}`;
        }
        setStatus(assignmentGuidelinesFromFile ? "✅ הנחיות הוטענו מקובץ" : "⚠️ הקובץ נטען ללא טקסט");
    } catch (e) {
        if (settingsStatusEl) settingsStatusEl.innerText = "❌ שגיאה בקריאת קובץ ההנחיות: " + e.message;
        if (quickStatusEl) quickStatusEl.innerText = "❌ שגיאה בקריאת הקובץ: " + e.message;
        assignmentGuidelinesFromFile = "";
        setStatus("❌ שגיאה בקריאת הנחיות מהקובץ");
    }
}

// ─── Selection Listener ───────────────────────────────────────────────────
function setupSelectionListener() {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    onSelectionChanged();
}

async function onSelectionChanged() {
    try {
        await Word.run(async (ctx) => {
            const range = ctx.document.getSelection();
            range.load("text");
            await ctx.sync();
            const preview = document.getElementById("selectedPreview");
            const badge = document.getElementById("selectionBadge");
            if (!range.text?.trim()) {
                currentSelectedText = "";
                if (preview) { preview.innerText = "לא נבחר טקסט — ה-AI קורא את כל המסמך"; preview.classList.add("empty-state"); }
                if (badge) badge.innerText = "";
            } else {
                currentSelectedText = range.text;
                if (preview) {
                    preview.innerText = currentSelectedText.length > 120
                        ? currentSelectedText.slice(0, 120) + "..." : currentSelectedText;
                    preview.classList.remove("empty-state");
                }
                if (badge) badge.innerText = `${currentSelectedText.length} תווים`;
            }
        });
    } catch(e) { console.error("selection error:", e); }
}

// ─── Init ─────────────────────────────────────────────────────────────────
window.onerror = (msg, url, lineNo) => {
    console.error(`Global Error: ${msg} @ ${url}:${lineNo}`);
    setStatus("שגיאה בטעינה — בדוק קונסול.");
    return false;
};

window.sendChatMessage = sendChatMessage;
setupChatEvents();

Office.onReady(async (info) => {
    try {
        if (info.host !== Office.HostType.Word) return;

        // Settings panel
        document.getElementById("btnSettings").onclick = () => {
            document.getElementById("settingsPanel").classList.toggle("hidden");
            document.getElementById("stylePanel").classList.add("hidden");
        };
        document.getElementById("saveKeyBtn").onclick = saveApiKey;
        const saveMaterialsBtn = document.getElementById("saveMaterialsBtn");
        if (saveMaterialsBtn) saveMaterialsBtn.onclick = saveMaterials;
        loadCloudSelectionState();
        loadCloudFolderSelectionState();
        setupCloudEvents();
        const debugToggle = document.getElementById("debugModeToggle");
        if (debugToggle) {
            debugToggle.onchange = () => {
                rsSet(DEBUG_MODE_KEY, debugToggle.checked ? "1" : "0");
                if (!debugToggle.checked) debugEvents = [];
                renderDebugLog();
                if (debugToggle.checked) debugLog("debug.enabled");
            };
        }

        // Style panel
        document.getElementById("btnStyle").onclick = () => {
            document.getElementById("stylePanel").classList.toggle("hidden");
            document.getElementById("settingsPanel").classList.add("hidden");
        };
        document.getElementById("addVocabBtn").onclick = addVocabItem;
        document.getElementById("vocabInput").onkeydown = e => {
            if (e.key === "Enter") { e.preventDefault(); addVocabItem(); }
        };
        document.getElementById("vocabChips").onclick = e => {
            const btn = e.target.closest(".chip-remove");
            if (btn) removeVocabItem(btn.dataset.text, btn.dataset.type, btn.dataset.source);
        };
        document.getElementById("saveStyleBtn").onclick = () => {
            personalStyle.notes = document.getElementById("styleNotes").value.trim();
            savePersonalStyle(); renderStylePanel();
            document.getElementById("styleStatus").innerText = "✅ נשמר!";
            setTimeout(() => document.getElementById("stylePanel").classList.add("hidden"), 1200);
        };
        document.getElementById("openExceptionsModalBtn").onclick = () => toggleExceptionsModal(true);
        document.getElementById("closeExceptionsModalBtn").onclick = () => toggleExceptionsModal(false);
        document.getElementById("exceptionsModal").onclick = (e) => {
            if (e.target.id === "exceptionsModal") toggleExceptionsModal(false);
        };
        document.getElementById("exceptionsModalBoard").onclick = (e) => {
            const btn = e.target.closest(".exception-approve-btn");
            if (!btn) return;
            approveExceptionItem(btn.dataset.kind, decodeURIComponent(btn.dataset.text || ""));
        };
        document.getElementById("refreshExceptionsBtn").onclick = renderExceptionsBoard;
        document.getElementById("relearnPastWorksBtn").onclick = relearnPastWorks;
        document.getElementById("resetAllStyleBtn").onclick = resetAllStyleDataOnce;
        document.getElementById("exportStyleBtn").onclick = exportStyleJson;
        document.getElementById("importStyleFile").onchange = e => {
            if (e.target.files[0]) importStyleJson(e.target.files[0]);
            e.target.value = "";
        };
        document.getElementById("styleLearningFile").onchange = async e => {
            if (e.target.files[0]) await learnStyleFromFile(e.target.files[0]);
            e.target.value = "";
        };
        document.getElementById("assignmentGuidelinesFile").onchange = async e => {
            if (e.target.files[0]) await loadAssignmentGuidelinesFromFile(e.target.files[0]);
            e.target.value = "";
        };
        document.getElementById("assignmentGuidelinesFileQuick").onchange = async e => {
            if (e.target.files[0]) await loadAssignmentGuidelinesFromFile(e.target.files[0]);
            e.target.value = "";
        };

        // Chat
        setupChatEvents();
        // Quick chips — activate mini-agent mode on click
        document.querySelectorAll(".quick-chip").forEach(chip => {
            chip.onclick = () => activateAgent(chip.dataset.agent);
        });
        renderActiveAgent(); // סנכרון מצב התחלתי

        // PDF.js worker
        if (window.pdfjsLib) {
            pdfjsLib.GlobalWorkerOptions.workerSrc =
                "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
        }

        loadApiKey();
        loadMaterials();
        setupCloudAuthObserver();
        debugLog("init.materials.loaded");
        loadChatSessions();
        debugLog("init.chat.loaded");
        await loadPersonalStyle();
        debugLog("init.personalStyle.loaded");
        setupSelectionListener();
        debugLog("init.selection.listener.ready");

        // לומד מעבודות קודמות ב-background (לא חוסם את ה-UI)
        autoLearnFromPastWorks();
        debugLog("init.autoLearn.triggered");

        setStatus("מוכן! הסגנון האישי שלך פעיל 🎯");
    } catch(e) {
        console.error("Init error:", e);
        setStatus("שגיאה באתחול: " + e.message);
    }
});
