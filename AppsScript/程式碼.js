// 💡 1. 讀取試算表中的「系統設定」分頁
function getSystemSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('系統設定');
  const settings = {};
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) {
        settings[data[i][0].toString().trim()] = data[i][1] ? data[i][1].toString().trim() : "";
      }
    }
  }
  return settings;
}

function doPost(e) {
  try {
    const settings = getSystemSettings();
    const json = JSON.parse(e.postData.contents);
    const event = json.events[0];

    if (!event) return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }));

    const replyToken = event.replyToken;
    const userId = event.source.userId;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const className = settings.CLASS_NAME || "班級";

    // 👤 讀取/更新「使用者對照」分頁（含學生姓名、暱稱、情緒標記）
    const userInfo = recordOrUpdateUser(ss, userId, settings.LINE_ACCESS_TOKEN);
    const parentName = userInfo.displayName;
    let studentName = userInfo.studentName;
    const userRow = userInfo.row;

    // 🌟 1. 處理「加入好友/歡迎訊息 (follow)」事件
    if (event.type === 'follow') {
      const welcomeMsg = `親愛的 ${parentName} 您好：\n\n歡迎加入！我是【${className}小幫手】🤖！\n\n為了提供更完善的服務，請先回覆告知您家「學生的姓名」（例如直接輸入：王小明 或 學生：王小明），系統將自動為您綁定對照喔！`;
      replyMessage(replyToken, welcomeMsg, settings.LINE_ACCESS_TOKEN);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }));
    }

    let userMessage = "";

    // 💡 處理語音訊息與文字訊息
    if (event.message.type === 'text') {
      userMessage = event.message.text.trim();
    } else if (event.message.type === 'audio') {
      userMessage = processAudioMessage(event.message.id, settings);
      if (!userMessage || userMessage.includes("失敗")) {
        const greeting = studentName ? `親愛的 ${studentName} 家長您好：` : `親愛的 ${parentName} 您好：`;
        replyMessage(replyToken, `${greeting}\n\n⚠️ 語音辨識失敗，請再試一次或改用文字輸入。`, settings.LINE_ACCESS_TOKEN);
        return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }));
      }
    } else {
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }));
    }

    // 🌟 2. 判斷是否為「登記/修改學生姓名」
    const isNameCommand = userMessage.startsWith("學生") || userMessage.startsWith("孩子") || userMessage.startsWith("名字");
    const cleanName = userMessage.replace(/^(?:學生姓名|學生|孩子|名字)[:：\s]*/, "").trim();
    const isPotentialName = /^[\u4e00-\u9fa5]{2,4}$/.test(cleanName) && !/問|幾|啥|誰|請|選|選單|試算|行程|考試|放學|網址|公告|段考|請假/.test(cleanName);

    if ((isNameCommand && cleanName.length >= 2) || (!studentName && isPotentialName)) {
      saveStudentName(ss, userRow, cleanName);
      studentName = cleanName;
      const newGreeting = `親愛的 ${cleanName} 家長您好：`;
      const confirmMsg = `${newGreeting}\n\n✅ 已成功登記學生姓名為「${cleanName}」！\n\n現在您可以隨時向我詢問：\n1️⃣ 班級規範與常見問題 (QA)\n2️⃣ 學校行事曆與班級行程\n3️⃣ 學校官網最新公告\n\n（若日後需修改學生姓名，隨時輸入「學生：新姓名」即可！）`;
      replyMessage(replyToken, confirmMsg, settings.LINE_ACCESS_TOKEN);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }));
    }

    // 動態判斷當前稱呼抬頭
    const greeting = studentName ? `親愛的 ${studentName} 家長您好：` : `親愛的 ${parentName} 您好：`;

    let sheetQA = ss.getSheetByName('QA');

    // 💡 管理員指令 (支援文字與語音)
    if (userId === settings.ADMIN_USER_ID) {
      if (userMessage === "選單" || userMessage === "指令" || userMessage === "指令表" || userMessage === "幫助") {
        replyMessage(replyToken, getAdminMenuText(className), settings.LINE_ACCESS_TOKEN);
        return ContentService.createTextOutput(JSON.stringify({ status: 'success' }));
      }

      if (userMessage.startsWith("新增QA") || userMessage.startsWith("新增關鍵字")) {
        handleAdminAddQA(userMessage, replyToken, sheetQA, settings.LINE_ACCESS_TOKEN);
        return ContentService.createTextOutput(JSON.stringify({ status: 'success' }));
      }

      if (userMessage.startsWith("新增行程") || userMessage.startsWith("新增行事曆") || userMessage.startsWith("新增日曆")) {
        const teacherCalId = settings.TEACHER_CALENDAR_ID || "primary";
        handleAdminAddCalendar(userMessage, replyToken, teacherCalId, settings.LINE_ACCESS_TOKEN);
        return ContentService.createTextOutput(JSON.stringify({ status: 'success' }));
      }
    }

    // 讀取快取資料
    const qaContext = getQAContext(sheetQA);
    const calContext = getSchoolCalendarFromSheet(ss);
    const newsContext = getSchoolNewsContext(ss);
    const fullContext = `${qaContext}\n${calContext}\n${newsContext}`;

    // 呼叫 Gemini AI
    const aiResponse = callGeminiWithFallback(userMessage, fullContext, settings.GEMINI_API_KEY, className);

    // 若 AI 偵測到情緒抱怨，僅自動在試算表中標註「需要特別關注」，不改變回覆語法
    if (aiResponse.autoFlagUser && userRow > 0) {
      autoMarkEmotionalUser(ss, userRow, "需要特別關注");
    }

    // 🌟 3. LINE 回應與 Email 觸發
    const replyContent = aiResponse.replyText || `我是【${className}小幫手】🤖。\n\n若需要導師親自協助，請留下您想請教的事項，我會為您轉達給導師喔！`;
    const formattedReply = `${greeting}\n\n${replyContent}`;
    
    replyMessage(replyToken, formattedReply, settings.LINE_ACCESS_TOKEN);

    // 只要 needEmail 為 true，發送 Email 給導師
    if (aiResponse.needEmail) {
      sendTeacherEmail(settings.TEACHER_EMAIL, parentName, studentName, userId, userMessage, aiResponse.reason, className);
    }

  } catch (err) {
    Logger.log("❌ doPost 錯誤: " + err.toString());
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'success' }));
}

// 👤 記錄並更新「使用者對照」分頁
function recordOrUpdateUser(ss, userId, lineToken) {
  if (!userId) return { displayName: "家長", userTag: "", studentName: "", row: -1 };

  let sheet = ss.getSheetByName('使用者對照');
  if (!sheet) {
    sheet = ss.insertSheet('使用者對照');
    sheet.appendRow(['LINE User ID', 'LINE 暱稱', '最後互動時間', '特別標記/情緒備註', '學生姓名']);
    sheet.getRange("A1:E1").setBackground("#e0e0e0").setFontWeight("bold");
  }

  const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
  
  let displayName = "家長";
  try {
    const url = `https://api.line.me/v2/bot/profile/${userId}`;
    const response = UrlFetchApp.fetch(url, { 'headers': { 'Authorization': 'Bearer ' + lineToken }, 'muteHttpExceptions': true });
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.displayName) displayName = data.displayName;
    }
  } catch (err) {
    Logger.log("⚠️ 取得 LINE Profile 失敗: " + err.toString());
  }

  const data = sheet.getDataRange().getValues();
  let foundRow = -1;
  let existingTag = "";
  let existingStudentName = "";

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      foundRow = i + 1;
      existingTag = data[i][3] ? data[i][3].toString().trim() : "";
      existingStudentName = data[i][4] ? data[i][4].toString().trim() : "";
      break;
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 2).setValue(displayName);
    sheet.getRange(foundRow, 3).setValue(nowStr);
  } else {
    sheet.appendRow([userId, displayName, nowStr, "", ""]);
    foundRow = sheet.getLastRow();
  }

  return { displayName: displayName, userTag: existingTag, studentName: existingStudentName, row: foundRow };
}

// 📝 儲存學生姓名至試算表
function saveStudentName(ss, row, studentName) {
  try {
    const sheet = ss.getSheetByName('使用者對照');
    if (sheet && row > 1) {
      sheet.getRange(row, 5).setValue(studentName);
      Logger.log(`✅ 已更新第 ${row} 列學生姓名為：${studentName}`);
    }
  } catch (e) {
    Logger.log("⚠️ 儲存學生姓名失敗: " + e.toString());
  }
}

// 🏷 自動在試算表中標註情緒家長（僅記錄於後台）
function autoMarkEmotionalUser(ss, row, tagValue) {
  try {
    const sheet = ss.getSheetByName('使用者對照');
    if (sheet && row > 1) {
      const currentTag = sheet.getRange(row, 4).getValue();
      if (!currentTag) {
        sheet.getRange(row, 4).setValue(tagValue);
        Logger.log(`🏷️ 已自動為列 ${row} 使用者標註：${tagValue}`);
      }
    }
  } catch (e) {}
}

// 📋 產生管理員指令表文字
function getAdminMenuText(className) {
  return `📋 【${className} 導師管理專用指令表】\n\n` +
         `導師可透過「文字輸入」或「直接傳送語音」發送以下指令：\n\n` +
         `1️⃣ 新增常見問答 (QA)\n` +
         `🔹 格式：新增QA 關鍵字 回應內容\n` +
         `🔹 範例：新增QA 請假 學生請假請至校務系統提出申請。\n` +
         `🔹 說明：關鍵字與內容間可用「空白」或「|」隔開。\n\n` +
         `2️⃣ 新增行程至個人日曆\n` +
         `🔹 指定時間：新增行程 MM/DD hh:mm 行程名稱\n` +
         `   範例：新增行程 9/29 11:30 考試\n` +
         `🔹 全天行程：新增行程 MM/DD 行程名稱\n` +
         `   範例：新增行程 10/20 校慶\n` +
         `💡 提示：行程會新增至您的個人日曆並同步給 AI 查詢！\n\n` +
         `3️⃣ 隨時查閱選單\n` +
         `🔹 輸入：選單、指令、指令表\n\n` +
         `🎙️ 語音小技巧：\n` +
         `您可以直接對機器人錄音，例如講出：「新增行程 9月29日 11點30分 考試」或「選單」，AI 聽懂後就會為您執行！`;
}

// 🎙️ 下載 LINE 語音並使用 Gemini 辨識轉文字
function processAudioMessage(messageId, settings) {
  try {
    const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
    const response = UrlFetchApp.fetch(url, {
      'headers': { 'Authorization': 'Bearer ' + settings.LINE_ACCESS_TOKEN },
      'muteHttpExceptions': true
    });

    if (response.getResponseCode() !== 200) return "（語音下載失敗）";

    const blob = response.getBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType() || "audio/m4a";

    return callGeminiAudioTranscribe(base64Data, mimeType, settings.GEMINI_API_KEY);
  } catch (err) {
    return "（語音處理失敗）";
  }
}

// 🎙️️ Gemini 專屬語音轉文字
function callGeminiAudioTranscribe(base64Data, mimeType, apiKey) {
  const dynamicModel = getLatestFlashModelName(apiKey);
  const candidateModels = Array.from(new Set([
    "gemini-flash-lite-latest",
    dynamicModel,
    "gemini-flash-latest",
    "gemini-3.6-flash",
    "gemini-2.5-flash"
  ].filter(Boolean)));

  const systemPrompt = "請將這段語音完整轉寫為繁體中文文字。不要加入開場白或解釋，只需精準輸出聽到的語音文字內容。";

  for (const modelName of candidateModels) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
    const payload = {
      "contents": [{
        "parts": [
          { "inline_data": { "mime_type": mimeType, "data": base64Data } },
          { "text": systemPrompt }
        ]
      }]
    };

    try {
      const options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        const result = JSON.parse(res.getContentText());
        return result.candidates[0].content.parts[0].text.trim();
      }
    } catch (err) {}
  }
  return "（語音辨識失敗）";
}

// 📅 導師新增行程到 Google 日曆
function handleAdminAddCalendar(userMessage, replyToken, calendarId, lineToken) {
  let cleanMsg = userMessage.replace(/^新增行程|^新增行事曆|^新增日曆/, "").trim();
  
  const regex = /^(\d{1,2})[\/\-](\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?\s+(.+)$/;
  const match = cleanMsg.match(regex);

  if (!match) {
    replyMessage(replyToken, "⚠️ 新增行程格式不正確！\n\n範例：\n新增行程 9/29 11:30 考試\n新增行程 10/15 校慶", lineToken);
    return;
  }

  const month = parseInt(match[1], 10) - 1;
  const day = parseInt(match[2], 10);
  const hour = match[3] !== undefined ? parseInt(match[3], 10) : null;
  const minute = match[4] !== undefined ? parseInt(match[4], 10) : null;
  const title = match[5].trim();

  const now = new Date();
  let year = now.getFullYear();

  let startDate;
  const isAllDay = (hour === null || minute === null);

  if (isAllDay) {
    startDate = new Date(year, month, day, 0, 0, 0);
  } else {
    startDate = new Date(year, month, day, hour, minute, 0);
  }

  if (startDate.getTime() < now.getTime()) {
    year += 1;
    if (isAllDay) {
      startDate = new Date(year, month, day, 0, 0, 0);
    } else {
      startDate = new Date(year, month, day, hour, minute, 0);
    }
  }

  try {
    const cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) {
      replyMessage(replyToken, "❌ 找不到指定日曆，請確認 TEACHER_CALENDAR_ID 設定！", lineToken);
      return;
    }

    if (isAllDay) {
      cal.createAllDayEvent(title, startDate);
    } else {
      const endDate = new Date(startDate.getTime() + 60 * 60 * 1000);
      cal.createEvent(title, startDate, endDate);
    }

    const timeStr = isAllDay 
      ? `${year}/${month+1}/${day} (全天)` 
      : `${year}/${month+1}/${day} ${hour}:${minute < 10 ? '0' + minute : minute}`;

    replyMessage(replyToken, `✅ 已成功加入教師個人日曆！\n\n📌 行程：${title}\n📅 時間：${timeStr}`, lineToken);
    
    syncSchoolCalendarToSheet();
  } catch (err) {
    replyMessage(replyToken, "❌ 新增行程失敗：" + err.toString(), lineToken);
  }
}

// 📧 寄送 Email 通知導師
function sendTeacherEmail(teacherEmail, parentName, studentName, userId, userMessage, reason, className) {
  if (!teacherEmail) return;
  const subject = `【${className} LINE 機器人家長訊息待處理通知】來自 ${studentName ? studentName + ' 家長' : parentName}`;
  const body = `親愛的老師，您好：\n\n${className} LINE 班級機器人收到了一則家長訊息，需請您撥空處理：\n\n🎓 學生姓名：${studentName || "未登記"}\n👤 家長 LINE 暱稱：${parentName}\n🆔 家長 LINE ID：${userId}\n💬 訊息內容：${userMessage}\n🔍 AI 判斷原因：${reason || '需導師親自聯繫處理'}\n\n請於方便時撥空處理，謝謝！`;

  try {
    MailApp.sendEmail(teacherEmail, subject, body);
  } catch (err1) {
    try {
      GmailApp.sendEmail(teacherEmail, subject, body);
    } catch (err2) {
      Logger.log(`❌ Email 寄送失敗: ${err2.toString()}`);
    }
  }
}

// 💡 動態抓取 Flash 模型名稱
function getLatestFlashModelName(apiKey) {
  const defaultModel = "gemini-flash-lite-latest";
  if (!apiKey) return defaultModel;

  const listUrl = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
  
  try {
    const response = UrlFetchApp.fetch(listUrl, { "muteHttpExceptions": true });
    const result = JSON.parse(response.getContentText());
    
    if (result.models && Array.isArray(result.models)) {
      const flashModels = result.models
        .filter(m => {
          const name = m.name.toLowerCase();
          const hasGen = m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent");
          const isFlash = name.includes("flash");
          const isExcluded = name.includes("tts") || name.includes("image") || name.includes("computer-use") || name.includes("omni") || name.includes("exp");
          return hasGen && isFlash && !isExcluded;
        })
        .map(m => m.name.replace("models/", ""));

      if (flashModels.length > 0) {
        flashModels.sort((a, b) => b.localeCompare(a, undefined, { numeric: true, sensitivity: 'base' }));
        return flashModels[0];
      }
    }
  } catch (err) {
    Logger.log("⚠️ 取得模型清單失敗，使用預設值: " + err.toString());
  }

  return defaultModel;
}

// 💡 整合後的模型呼叫順序
function callGeminiWithFallback(userMessage, fullContext, apiKey, className) {
  const dynamicModel = getLatestFlashModelName(apiKey);
  const candidateModels = Array.from(new Set([
    "gemini-flash-lite-latest",
    dynamicModel,
    "gemini-flash-latest",
    "gemini-3.6-flash",
    "gemini-2.5-flash"
  ].filter(Boolean)));

  for (const modelName of candidateModels) {
    const res = callGeminiSingleModel(userMessage, fullContext, modelName, apiKey, className);
    if (res.success) return res.data;
  }
  return { replyText: "AI 服務暫時無回應，已為您將訊息記錄，請導師撥空處理。", needEmail: true, reason: "AI 連線異常備援", autoFlagUser: false };
}

// 🧠 Gemini 核心智慧判斷（保持統一平穩專業語氣）
function callGeminiSingleModel(userMessage, fullContext, modelName, apiKey, className) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

  const systemPrompt = `你是一位態度平穩、專業、客觀且值得信賴的班級助理機器人【${className}小幫手】。請**嚴格根據**【參考資料】處理提問。

【參考資料】：
${fullContext}

【邏輯判斷與回應規則】：

1. **資料庫有明確解答 / 一般日常問候**：
   - replyText: 給出精準解答或親切問候。
   - needEmail: false
   - autoFlagUser: false
   - reason: "已完成自動解答"

2. **訊息屬於「純宣洩情緒/抱怨」且「沒有說明具體事件/需求」**：
   - replyText: 保持平穩專業態度說明：「了解您的關切，為了能最有效率協助您並轉達導師，請詳細說明您『遇到的具體狀況/經過』，我會立即為您通知導師發送郵件。」
   - needEmail: false
   - autoFlagUser: true （💡 標記此位家長需特別關注，但不改變回覆語風）
   - reason: "純抱怨無事由，引導提供詳細資訊"

3. **訊息中有說明「具體事件/狀況」（例：受傷、體育課、請假、發生什麼事、疑問等）**：
   - ⚠️ **語風控制**：語氣必須**平穩、專業、客觀、簡潔**，嚴禁使用非專業的誇張情緒字眼（如「好心疼」、「委屈難受」）。
   - replyText: 簡短客觀回應已收到反映，並明確告知已發送 Email 轉達導師。
     💡 **標準範例**：「已收到您反映關於孩子在 [事由/地點] 發生的狀況。我已將相關訊息透過 Email 轉發給導師，導師會在上班時間親自與您聯繫處理，謝謝您的告知。」
   - needEmail: true
   - autoFlagUser: false
   - reason: "已取得具體事由，已發送 Email 轉報導師"

【格式要求】：
- 務必輸出合法 JSON 格式：
  { "replyText": "...", "needEmail": true/false, "autoFlagUser": true/false, "reason": "..." }
- 開頭請勿加上打招呼語（如「你好」、「親愛的家長您好」），外層系統會自動統一加上稱呼。`;

  const payload = {
    "contents": [{ "parts": [{ "text": `${systemPrompt}\n\n使用者訊息：${userMessage}` }] }],
    "generationConfig": { "response_mime_type": "application/json" }
  };

  try {
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) return { success: false };

    const result = JSON.parse(response.getContentText());
    let rawText = result.candidates[0].content.parts[0].text.replace(/```json/g, "").replace(/```/g, "").trim();
    return { success: true, data: JSON.parse(rawText) };
  } catch (err) {
    return { success: false };
  }
}

// 💡 輔助函式：取得日曆物件，若使用者尚未訂閱，自動嘗試為其訂閱
function getOrSubscribeCalendar(calendarId) {
  if (!calendarId) return null;
  try {
    let cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) {
      cal = CalendarApp.subscribeToCalendar(calendarId);
    }
    return cal;
  } catch (e) {
    return null;
  }
}

// 📅 智慧排程任務：同步日曆（包含 Google 日曆 API + 公開 iCal 爬取雙備援）
function syncSchoolCalendarToSheet() {
  try {
    const settings = getSystemSettings();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('行事曆Cache');
    if (!sheet) sheet = ss.insertSheet('行事曆Cache');
    
    sheet.clearContents();
    sheet.appendRow(['更新時間', '行程標題', '開始日期', '結束日期', '來源']);

    const now = new Date();
    const endDate = new Date(now.getTime() + 180 * 24 * 60 * 60 * 1000);
    const updateTime = Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd HH:mm");

    const allEvents = [];

    // 1. 抓取學校日曆 (優先用 CalendarApp，抓不到則啟動 iCal 公開網址爬取備援)
    if (settings.SCHOOL_CALENDAR_ID) {
      let schoolEventsCount = 0;
      try {
        const schoolCal = getOrSubscribeCalendar(settings.SCHOOL_CALENDAR_ID);
        if (schoolCal) {
          const events = schoolCal.getEvents(now, endDate);
          events.forEach(e => {
            allEvents.push({ title: e.getTitle(), start: e.getStartTime(), end: e.getEndTime(), source: "學校" });
            schoolEventsCount++;
          });
        }
      } catch (e) {
        Logger.log("⚠️ CalendarApp 讀取學校日曆失敗，準備啟動 iCal 備援: " + e.toString());
      }

      // 🛡️ 備援機制：如果 CalendarApp 未抓到任何行程，自動透過 HTTP 直接爬取公開 iCal 訂閱檔
      if (schoolEventsCount === 0) {
        Logger.log("ℹ️ 啟動爬取備援機制：直接抓取公開 iCal 檔...");
        const backupEvents = fetchPublicIcalEvents(settings.SCHOOL_CALENDAR_ID);
        backupEvents.forEach(e => allEvents.push(e));
      }
    }

    // 2. 抓取教師個人日曆
    const teacherCalId = settings.TEACHER_CALENDAR_ID || "primary";
    try {
      const teacherCal = getOrSubscribeCalendar(teacherCalId);
      if (teacherCal && teacherCalId !== settings.SCHOOL_CALENDAR_ID) {
        const events = teacherCal.getEvents(now, endDate);
        events.forEach(e => {
          allEvents.push({ title: e.getTitle(), start: e.getStartTime(), end: e.getEndTime(), source: "班級" });
        });
      }
    } catch (e) {}

    // 按時間排序
    allEvents.sort((a, b) => a.start - b.start);

    // 寫入快取分頁
    allEvents.forEach(e => {
      const startStr = Utilities.formatDate(e.start, "Asia/Taipei", "yyyy/MM/dd HH:mm");
      const endStr = Utilities.formatDate(e.end, "Asia/Taipei", "yyyy/MM/dd HH:mm");
      sheet.appendRow([updateTime, e.title, startStr, endStr, e.source]);
    });

    Logger.log(`✅ 日曆同步完成，共抓取 ${allEvents.length} 筆行程！`);

  } catch (err) {
    Logger.log("❌ 同步日曆失敗: " + err.toString());
  }
}

// 🌐 備援爬取函式：直接爬取 Google 日曆公開的 basic.ics 訂閱檔
function fetchPublicIcalEvents(calendarId) {
  if (!calendarId) return [];
  const icsUrl = `https://calendar.google.com/calendar/ical/${encodeURIComponent(calendarId)}/public/basic.ics`;
  try {
    const response = UrlFetchApp.fetch(icsUrl, { "muteHttpExceptions": true });
    if (response.getResponseCode() !== 200) {
      Logger.log("⚠️ 爬取 iCal 失敗，HTTP 回應碼: " + response.getResponseCode());
      return [];
    }
    const icsText = response.getContentText("UTF-8");
    return parseIcsEvents(icsText);
  } catch (e) {
    Logger.log("⚠️ 爬取 iCal 發生例外: " + e.toString());
    return [];
  }
}

// 🧩 解析 iCal (.ics) 文字內容為行程物件
function parseIcsEvents(icsText) {
  const events = [];
  const now = new Date();
  const maxDate = new Date(now.getTime() + 180 * 24 * 60 * 60 * 1000);
  const blocks = icsText.split("BEGIN:VEVENT");

  for (let i = 1; i < blocks.length; i++) {
    const block = blocks[i].split("END:VEVENT")[0];
    const summaryMatch = block.match(/SUMMARY:(.*)/);
    const dtstartMatch = block.match(/DTSTART(?:;[^:]+)?:(\d{8}(?:T\d{6}Z?)?)/);
    const dtendMatch = block.match(/DTEND(?:;[^:]+)?:(\d{8}(?:T\d{6}Z?)?)/);

    if (summaryMatch && dtstartMatch) {
      const title = summaryMatch[1].trim();
      const startDate = parseIcsDate(dtstartMatch[1]);
      let endDate = dtendMatch ? parseIcsDate(dtendMatch[1]) : startDate;

      if (startDate && startDate >= new Date(now.getTime() - 24 * 60 * 60 * 1000) && startDate <= maxDate) {
        events.push({ title: title, start: startDate, end: endDate, source: "學校" });
      }
    }
  }
  return events;
}

// 🗓️ 解析 iCal 日期格式 (YYYYMMDD 或 YYYYMMDDTHHMMSSZ)
function parseIcsDate(dateStr) {
  if (!dateStr) return null;
  const y = parseInt(dateStr.substring(0, 4), 10);
  const m = parseInt(dateStr.substring(4, 6), 10) - 1;
  const d = parseInt(dateStr.substring(6, 8), 10);
  let h = 0, min = 0;
  if (dateStr.includes("T") && dateStr.length >= 13) {
    h = parseInt(dateStr.substring(9, 11), 10);
    min = parseInt(dateStr.substring(11, 13), 10);
  }
  return new Date(y, m, d, h, min);
}

// 排程任務：抓取學校官網公告 (近 1 個月)
function fetchSchoolAnnouncements() {
  try {
    const settings = getSystemSettings();
    if (!settings.SCHOOL_ANNOUNCEMENT_URL) return;

    const response = UrlFetchApp.fetch(settings.SCHOOL_ANNOUNCEMENT_URL, { "muteHttpExceptions": true });
    const html = response.getContentText("UTF-8");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('學校公告');
    if (!sheet) {
      sheet = ss.insertSheet('學校公告');
      sheet.appendRow(['抓取時間', '公告標題', '公告連結', '發布日期']);
    }

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();
    }

    const now = new Date();
    const nowStr = Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd HH:mm");
    const thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

    const linkRegex = /<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
    let match;
    const seenTitles = new Set();
    let addedCount = 0;

    while ((match = linkRegex.exec(html)) !== null) {
      let link = match[1].trim();
      let rawTitle = match[2].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();

      if (link.startsWith("/")) {
        const domainMatch = settings.SCHOOL_ANNOUNCEMENT_URL.match(/^(https?:\/\/[^\/]+)/);
        const domain = domainMatch ? domainMatch[1] : "";
        link = domain + link;
      }

      if (rawTitle.length <= 5 || rawTitle.includes("更多") || seenTitles.has(rawTitle)) {
        continue;
      }

      const itemDate = parseDateFromText(match[0] + " " + rawTitle);
      let dateStr = "";

      if (itemDate) {
        if (itemDate < thirtyDaysAgo) {
          continue;
        }
        dateStr = Utilities.formatDate(itemDate, "Asia/Taipei", "yyyy/MM/dd");
      }

      seenTitles.add(rawTitle);
      sheet.appendRow([nowStr, rawTitle, link, dateStr]);
      addedCount++;

      if (addedCount >= 30) break;
    }

    Logger.log(`✅ 已成功抓取 ${addedCount} 筆近一個月學校公告！`);
  } catch (err) {
    Logger.log("❌ 抓取公告失敗: " + err.toString());
  }
}

// 💡 輔助解析日期函式
function parseDateFromText(text) {
  if (!text) return null;
  let match = text.match(/(20\d{2})[\/\.\-](\d{1,2})[\/\.\-](\d{1,2})/);
  if (match) {
    return new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10));
  }
  match = text.match(/(1\d{2})[\/\.\-](\d{1,2})[\/\.\-](\d{1,2})/);
  if (match) {
    let adYear = parseInt(match[1], 10) + 1911;
    return new Date(adYear, parseInt(match[2], 10) - 1, parseInt(match[3], 10));
  }
  return null;
}

// 輔助函式
function getQAContext(sheet) {
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  let qaText = "【班級規範與常見問題】：\n";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      qaText += `【${data[i][0]}】: ${data[i][1]}\n`;
    }
  }
  return qaText;
}

function getSchoolCalendarFromSheet(ss) {
  const sheet = ss.getSheetByName('行事曆Cache');
  if (!sheet || sheet.getLastRow() <= 1) return "【未來半年行事曆與班級行程】：（目前無預定行程）\n";

  const data = sheet.getDataRange().getValues();
  let calText = "【未來半年行事曆與班級行程】：\n";
  for (let i = 1; i < data.length; i++) {
    const sourceStr = data[i][4] ? `[${data[i][4]}] ` : "";
    calText += `- ${sourceStr}${data[i][2]}: ${data[i][1]}\n`;
  }
  return calText;
}

function getSchoolNewsContext(ss) {
  const sheet = ss.getSheetByName('學校公告');
  if (!sheet || sheet.getLastRow() <= 1) return "【學校近一個月官網公告】：（目前無公告）\n";

  const data = sheet.getDataRange().getValues();
  let newsText = "【學校近一個月官網公告】：\n";
  const maxRows = Math.min(data.length, 25);
  for (let i = 1; i < maxRows; i++) {
    if (data[i][1] && data[i][2]) {
      const dateInfo = data[i][3] ? `[${data[i][3]}] ` : "";
      newsText += `- ${dateInfo}${data[i][1]} (網址: ${data[i][2]})\n`;
    }
  }
  return newsText;
}

function handleAdminAddQA(userMessage, replyToken, sheet, lineToken) {
  let cleanMsg = userMessage.replace(/^新增QA|^新增關鍵字/, "").trim();
  let keyword = "", content = "";

  if (cleanMsg.includes("|")) {
    const parts = cleanMsg.split("|");
    keyword = parts[0].trim();
    content = parts.slice(1).join("|").trim();
  } else {
    const match = cleanMsg.match(/^(\S+)\s+([\s\S]+)$/);
    if (match) {
      keyword = match[1].trim();
      content = match[2].trim();
    }
  }

  if (keyword && content) {
    sheet.appendRow([keyword, content]);
    replyMessage(replyToken, `✅ 已成功新增 QA 到試算表！\n\n📌 關鍵字：${keyword}\n📝 回應內容：${content}`, lineToken);
  } else {
    replyMessage(replyToken, `⚠️ 新增失敗，格式：新增QA 關鍵字 回應內容`, lineToken);
  }
}

function replyMessage(replyToken, text, lineToken) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    'headers': { 'Content-Type': 'application/json; charset=UTF-8', 'Authorization': 'Bearer ' + lineToken },
    'method': 'post',
    'payload': JSON.stringify({ 'replyToken': replyToken, 'messages': [{ 'type': 'text', 'text': text }] })
  });
}

// 🧪 專用檢測權限與設定的測試函式
function testPermissions() {
  Logger.log("=== 🔍 開始檢測 LINE Bot 與 GAS 權限及設定 ===");
  const settings = getSystemSettings();

  Logger.log("✅ 1. 系統設定讀取成功！");
  Logger.log("   - CLASS_NAME: " + (settings.CLASS_NAME ? settings.CLASS_NAME : "⚠️ 未設定 (預設為 '班級')"));

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetQA = ss.getSheetByName('QA');
    if (!sheetQA) {
      sheetQA = ss.insertSheet('QA');
      sheetQA.appendRow(['關鍵字', '回應內容']);
    }
    Logger.log("✅ 2. 試算表 QA 分頁讀寫權限正常！");
  } catch (e) {
    Logger.log("❌ 2. 試算表寫入失敗: " + e.toString());
  }

  try {
    if (settings.SCHOOL_CALENDAR_ID) {
      const cal = getOrSubscribeCalendar(settings.SCHOOL_CALENDAR_ID);
      if (cal) Logger.log("✅ 3A. 學校日曆讀取/訂閱成功！名稱: " + cal.getName());
      else Logger.log("ℹ️ 3A. 無法透過 CalendarApp 存取，將自動啟用 iCal 網址直接爬取備援。");
    }
  } catch (e) {
    Logger.log("❌ 3A. 學校日曆讀取失敗: " + e.toString());
  }

  try {
    const teacherCalId = settings.TEACHER_CALENDAR_ID || "primary";
    const teacherCal = getOrSubscribeCalendar(teacherCalId);
    if (teacherCal) Logger.log("✅ 3B. 教師個人日曆連結成功！名稱: " + teacherCal.getName());
    else Logger.log("❌ 3B. 找不到教師個人日曆！");
  } catch (e) {
    Logger.log("❌ 3B. 教師日曆存取失敗: " + e.toString());
  }

  try {
    if (settings.TEACHER_EMAIL) {
      sendTeacherEmail(settings.TEACHER_EMAIL, "【測試員】", "測試學生", "TestID", "權限與寄信測試", "測試用途", settings.CLASS_NAME || "班級");
      Logger.log("✅ 4. 測試信件已成功發送至 " + settings.TEACHER_EMAIL);
    }
  } catch (e) {
    Logger.log("❌ 4. Email 發送失敗: " + e.toString());
  }

  try {
    if (settings.GEMINI_API_KEY) {
      const res = callGeminiWithFallback("你好", "測試上下文", settings.GEMINI_API_KEY, settings.CLASS_NAME || "班級");
      if (res) Logger.log("✅ 5. Gemini API 測試成功！");
    }
  } catch (e) {
    Logger.log("❌ 5. Gemini API 呼叫異常: " + e.toString());
  }

  try {
    if (settings.LINE_ACCESS_TOKEN) {
      const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
        headers: { 'Authorization': 'Bearer ' + settings.LINE_ACCESS_TOKEN },
        muteHttpExceptions: true
      });
      if (res.getResponseCode() === 200) {
        const botInfo = JSON.parse(res.getContentText());
        Logger.log("✅ 6. LINE Access Token 有效！名稱: " + botInfo.displayName);
      }
    }
  } catch (e) {
    Logger.log("❌ 6. LINE API 連線異常: " + e.toString());
  }

  Logger.log("=== 🏁 檢測完成 ===");
}