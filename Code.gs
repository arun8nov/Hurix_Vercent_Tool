/**
 * Vercent Voice Agent for English Learning
 * Server-Side Script (Code.gs)
 */

// Global constant to hold sheet ID if not bound to a spreadsheet. 
// If bound, it automatically uses SpreadsheetApp.getActiveSpreadsheet().
var SPREADSHEET_ID = "1RognFtyIZc6V27zVYx1JPcyXeLwop83IlS3tr67A5hc"; 

/**
 * Serves the web app UI
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile("Index");
  
  // Get active user email to personalize the dashboard
  var email = "Unknown User";
  try {
    email = Session.getActiveUser().getEmail();
  } catch (err) {
    Logger.log("Failed to fetch email: " + err.toString());
  }
  
  template.userEmail = email;
  
  return template.evaluate()
    .setTitle("Vercent English Voice Agent")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper to include HTML partials (CSS/JS)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Returns the active spreadsheet object
 */
function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Initializes Sheet tabs if they do not exist
 */
function initDatabase() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 10 seconds timeout
    
    var ss = getSpreadsheet();
    if (!ss) {
      return "Error: No active spreadsheet bound. Please open a Google Sheet and click Extension > Apps Script, or set your SPREADSHEET_ID in Code.gs.";
    }
    
    // Create EnglishReports tab if missing
    var reportsSheet = ss.getSheetByName("EnglishReports");
    if (!reportsSheet) {
      reportsSheet = ss.insertSheet("EnglishReports");
      reportsSheet.appendRow([
        "Timestamp", 
        "User Email", 
        "Scenario", 
        "Total Turns", 
        "Avg Grammar Score", 
        "Avg Vocab Score", 
        "Speech Rate (WPM)", 
        "Duration (Secs)",
        "Session ID"
      ]);
      reportsSheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    }
    
    // Create VoiceLogs tab if missing
    var logsSheet = ss.getSheetByName("VoiceLogs");
    if (!logsSheet) {
      logsSheet = ss.insertSheet("VoiceLogs");
      logsSheet.appendRow([
        "Timestamp", 
        "User Email", 
        "Session ID", 
        "Turn Number", 
        "User Spoke", 
        "AI Spoke", 
        "Grammar Score",
        "Vocab Score",
        "Corrections Json"
      ]);
      logsSheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    }
    
    return "Database initialized successfully.";
  } catch (e) {
    Logger.log("DB Init Error: " + e.toString());
    return "Error initializing database: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fetches the user email for frontend JavaScript access
 */
function getUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || "anonymous@vercent.com";
  } catch (e) {
    return "anonymous@vercent.com";
  }
}

/**
 * Connects to Gemini 1.5 Flash using the user's API key
 * and returns the parsed JSON schema response.
 */
function fetchGeminiResponse(chatHistory, userPrompt, apiKey, systemPrompt) {
  if (!apiKey || apiKey.trim() === "") {
    throw new Error("Gemini API Key is missing. Please enter it in the settings panel.");
  }
  
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey;
  
  // Format conversation history for Gemini
  var contents = [];
  
  // Feed past conversation turns
  if (chatHistory && chatHistory.length > 0) {
    for (var i = 0; i < chatHistory.length; i++) {
      contents.push({
        role: chatHistory[i].role === "assistant" ? "model" : "user",
        parts: [{ text: chatHistory[i].text }]
      });
    }
  }
  
  // Append current user prompt
  contents.push({
    role: "user",
    parts: [{ text: userPrompt }]
  });
  
  // Request response in JSON with strict Schema
  var payload = {
    contents: contents,
    systemInstruction: {
      parts: [{ text: systemPrompt }]
    },
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: {
        type: "OBJECT",
        properties: {
          reply: { 
            type: "STRING", 
            description: "Encouraging English reply to user's prompt (keep it under 3 sentences, spoken style)." 
          },
          corrections: {
            type: "ARRAY",
            description: "Grammatical or vocabulary mistakes that the user made in their spoken prompt.",
            items: {
              type: "OBJECT",
              properties: {
                original: { type: "STRING", description: "The exact wrong word or phrase user spoke." },
                corrected: { type: "STRING", description: "The correct English alternative." },
                explanation: { type: "STRING", description: "A simple, quick grammar advice (under 15 words)." }
              },
              required: ["original", "corrected", "explanation"]
            }
          },
          grammarScore: { 
            type: "INTEGER", 
            description: "Score out of 100 for the user's grammatical accuracy in this turn." 
          },
          vocabularyScore: { 
            type: "INTEGER", 
            description: "Score out of 100 for the user's word selection complexity and suitability." 
          }
        },
        required: ["reply", "corrections", "grammarScore", "vocabularyScore"]
      }
    }
  };
  
  var options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    if (responseCode !== 200) {
      var errorJson = JSON.parse(responseText);
      var errMsg = (errorJson.error && errorJson.error.message) ? errorJson.error.message : "HTTP Error " + responseCode;
      throw new Error("Gemini API Error: " + errMsg);
    }
    
    var responseJson = JSON.parse(responseText);
    
    if (responseJson.candidates && responseJson.candidates[0] && responseJson.candidates[0].content && responseJson.candidates[0].content.parts[0]) {
      return responseJson.candidates[0].content.parts[0].text;
    } else {
      throw new Error("Invalid response format received from Gemini.");
    }
  } catch (e) {
    Logger.log("Gemini API Call failed: " + e.toString());
    throw new Error(e.message || e.toString());
  }
}

/**
 * Saves the session evaluation report and logs to Google Sheets securely
 */
function saveSession(sessionData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 15 seconds queue timeout
    
    var ss = getSpreadsheet();
    if (!ss) {
      throw new Error("Database Spreadsheet is not open or bound.");
    }
    
    var email = getUserEmail();
    var timestamp = new Date();
    
    // 1. Write summary to EnglishReports
    var reportsSheet = ss.getSheetByName("EnglishReports");
    if (!reportsSheet) {
      initDatabase();
      reportsSheet = ss.getSheetByName("EnglishReports");
    }
    
    reportsSheet.appendRow([
      timestamp,
      email,
      sessionData.scenario,
      sessionData.totalTurns,
      sessionData.avgGrammar,
      sessionData.avgVocab,
      sessionData.wpm,
      sessionData.duration,
      sessionData.sessionId
    ]);
    
    // 2. Write details to VoiceLogs
    var logsSheet = ss.getSheetByName("VoiceLogs");
    if (!logsSheet) {
      initDatabase();
      logsSheet = ss.getSheetByName("VoiceLogs");
    }
    
    var turns = sessionData.turns;
    if (turns && turns.length > 0) {
      for (var i = 0; i < turns.length; i++) {
        var turn = turns[i];
        logsSheet.appendRow([
          timestamp,
          email,
          sessionData.sessionId,
          i + 1,
          turn.userSpoke,
          turn.aiSpoke,
          turn.grammarScore,
          turn.vocabScore,
          JSON.stringify(turn.corrections || [])
        ]);
      }
    }
    
    return { success: true, message: "Session results saved successfully to Google Sheets." };
  } catch (e) {
    Logger.log("Failed to save session: " + e.toString());
    return { success: false, message: "Error saving to Sheet: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fetches all report logs from the Sheet to render numerical analytics
 */
function getReportsSummary() {
  try {
    var ss = getSpreadsheet();
    if (!ss) return [];
    
    var sheet = ss.getSheetByName("EnglishReports");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; // Only headers
    
    var currentUser = getUserEmail();
    var isAdmin = false; // We can designate specific admins if needed (e.g. check domain or role)
    
    var reports = [];
    // Skip headers (row 0)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // Simple authorization filter: regular users see only their own data. Admin sees everything.
      if (row[1] === currentUser || isAdmin) {
        reports.push({
          timestamp: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "",
          email: row[1],
          scenario: row[2],
          turns: row[3],
          grammar: row[4],
          vocab: row[5],
          wpm: row[6],
          duration: row[7],
          sessionId: row[8]
        });
      }
    }
    return reports.reverse(); // Newest first
  } catch (e) {
    Logger.log("Error loading reports: " + e.toString());
    return [];
  }
}
