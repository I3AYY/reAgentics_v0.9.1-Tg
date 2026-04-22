// =========================================================================
// System: reAgentics - Laboratory Reagent Management System
// Version: 0.9.1 (AI Core)
// Developer: P. PURICUMPEE & AI Assistant
// Description: Backend Google Apps Script (Server-side logic)
// Update: Added Reagent Type, UI Enhancements, Database Column Shift, and Telegram Topic Support
// =========================================================================

// -------------------------------------------------------------------------
// 1. DATABASE CONFIGURATION (SaaS Architecture)
// -------------------------------------------------------------------------
const DEFAULT_DB = {
  MAIN: 'ใส่_ID_ไฟล์_reAgentics_DB_ที่นี่',       // Items, Stock_Balance
  UNIT: 'ใส่_ID_ไฟล์_reAgentics_Units_ที่นี่',     // Units, ReagUnits, Analyzers, storageLocation, Company, ReagTypes
  USER: 'ใส่_ID_ไฟล์_reAgentics_User_ที่นี่',      // User
  CONFIG: 'ใส่_ID_ไฟล์_reAgentics_Config_ที่นี่',  // Sticker_Config, App_Logo, Year_Config
  LOG: 'ใส่_ID_ไฟล์_reAgentics_Log_ที่นี่',        // System_Logs
  FOLDER_PROFILE: '', // Folder ID สำหรับโปรไฟล์
  FOLDER_LOGO: ''     // Folder ID สำหรับโลโก้
};

function getDbConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    MAIN: props.getProperty('DB_MAIN') || DEFAULT_DB.MAIN,
    UNIT: props.getProperty('DB_UNIT') || DEFAULT_DB.UNIT,
    USER: props.getProperty('DB_USER') || DEFAULT_DB.USER,
    CONFIG: props.getProperty('DB_CONFIG') || DEFAULT_DB.CONFIG,
    LOG: props.getProperty('DB_LOG') || DEFAULT_DB.LOG,
    FOLDER_PROFILE: props.getProperty('FOLDER_PROFILE') || DEFAULT_DB.FOLDER_PROFILE,
    FOLDER_LOGO: props.getProperty('FOLDER_LOGO') || DEFAULT_DB.FOLDER_LOGO
  };
}

function checkDatabaseSetup() {
  const config = getDbConfig();
  if (!config.MAIN || config.MAIN.includes('ใส่_ID_ไฟล์') || !config.USER || config.USER.includes('ใส่_ID_ไฟล์')) {
    throw new Error("คุณยังไม่ได้ตั้งค่า Google Sheet ID ครับ กรุณานำ ID มาใส่ในหน้า 'ตั้งค่าฐานข้อมูล' ให้ครบถ้วน");
  }
}

function deleteOldDriveFile(oldUrl) {
  if (oldUrl && oldUrl.includes("drive.google.com")) {
    try {
      let fileId = "";
      const idMatch = oldUrl.match(/id=([^&]+)/);
      const dMatch = oldUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (idMatch && idMatch[1]) fileId = idMatch[1];
      else if (dMatch && dMatch[1]) fileId = dMatch[1];
      if (fileId) DriveApp.getFileById(fileId).setTrashed(true);
    } catch (err) { console.log("Delete old file error: " + err); }
  }
}

// -------------------------------------------------------------------------
// 2. CORE WEB APP FUNCTIONS
// -------------------------------------------------------------------------
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('reAgentics | Lab Inventory System (v0.9.1 AI Core)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function safeString(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (val === null || val === undefined) return "";
  return String(val).trim();
}

// -------------------------------------------------------------------------
// 3. SYSTEM LOGGING & TELEGRAM
// -------------------------------------------------------------------------
function logSystem(action, detail, userId) {
  try {
    checkDatabaseSetup(); 
    const config = getDbConfig();
    if (config.LOG && !config.LOG.includes('ใส่_ID_ไฟล์')) {
      const logSS = SpreadsheetApp.openById(config.LOG);
      let sheet = logSS.getSheetByName('System_Logs');
      if (!sheet) {
        sheet = logSS.insertSheet('System_Logs');
        sheet.appendRow(['Timestamp', 'UserID', 'Action', 'Detail']);
        sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#e2e8f0");
        sheet.setFrozenRows(1);
      }
      let timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      sheet.appendRow([timeStr, userId, action, detail]);
    }
  } catch(e) { console.error("Log Sys Error: " + e); }
}

function sendTelegramNotification(unitName, messageText) {
  try {
    const config = getDbConfig();
    if (!config.UNIT || config.UNIT.includes('ใส่_ID_ไฟล์')) return;
    const ss = SpreadsheetApp.openById(config.UNIT);
    const sheet = ss.getSheetByName('Units');
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    let token = '', rawChatId = '';
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(unitName).trim()) {
        token = String(data[i][5]).trim(); 
        rawChatId = String(data[i][6]).trim(); 
        break;
      }
    }

    if (token && rawChatId) {
      let finalChatId = rawChatId;
      let topicId = null;

      // เช็คว่ามีการใส่ Topic ID พ่วงมาด้วยหรือไม่ (เช่น -1001234567890:45)
      if (rawChatId.includes(":")) {
        let parts = rawChatId.split(":");
        finalChatId = parts[0];
        topicId = parseInt(parts[1], 10);
      }

      const url = `https://api.telegram.org/bot${token}/sendMessage`;
      const payload = { 
        chat_id: finalChatId, 
        text: messageText, 
        parse_mode: 'HTML' 
      };

      // ถ้ามี Topic ID ให้แนบ message_thread_id เข้าไปด้วย
      if (topicId) {
        payload.message_thread_id = topicId;
      }

      const options = { 
        method: 'post', 
        contentType: 'application/json', 
        payload: JSON.stringify(payload), 
        muteHttpExceptions: true 
      };
      
      UrlFetchApp.fetch(url, options);
    }
  } catch(e) { console.error("Telegram Error: ", e); }
}

// -------------------------------------------------------------------------
// 4. AUTHENTICATION & SECURITY
// -------------------------------------------------------------------------
function verifyLogin(userId, password) {
  try {
    checkDatabaseSetup(); 
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.USER);
    const sheet = ss.getSheetByName("User");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(userId) && String(data[i][2]) === String(password)) {
        let status = String(data[i][8] || "ปกติ").trim();
        if (status === "ระงับการใช้งาน") {
          logSystem("Login Blocked", "Suspended user tried to login", userId);
          return { success: false, message: `บัญชีผู้ใช้ ${userId} ถูกระงับ โปรดติดต่อผู้ดูแลระบบ` };
        }

        let email = data[i][3];
        let name = data[i][0];
        let otpStatus = String(data[i][9] || "ON").trim().toUpperCase();
        
        if (!email && otpStatus === "ON") {
          logSystem("Login Failed", "Account missing email", userId);
          return { success: false, message: "บัญชีนี้ยังไม่ได้ตั้งค่า Email กรุณาติดต่อ Admin" };
        }

        if (otpStatus === "OFF") {
          let userProfile = getUserProfileById(userId);
          if (userProfile) {
            logSystem("Login Success", "User authenticated (OTP Bypassed)", userId);
            let availableYears = getAvailableYears();
            return { success: true, bypassed: true, message: "เข้าสู่ระบบสำเร็จ (OTP Bypassed)", user: userProfile, years: availableYears };
          } else { return { success: false, message: "ไม่พบข้อมูลโปรไฟล์ผู้ใช้งาน" }; }
        }

        let cache = CacheService.getScriptCache();
        let existingOtp = cache.get("OTP_" + userId);
        
        if (existingOtp) {
          logSystem("Login Info", "User logged in with active OTP session", userId);
          return { success: true, message: "ระบบได้ส่ง OTP ไปก่อนหน้านี้แล้ว กรุณาใช้รหัสเดิม (อายุรหัส 5 นาที)", email: email, userId: userId, otpExists: true };
        } else {
          let otpResult = generateAndSendOTP(userId, email, name);
          if (otpResult.success) {
            logSystem("OTP Requested", "New OTP sent to email", userId);
            return { success: true, message: "กรุณาตรวจสอบ OTP ที่อีเมลของคุณ", email: email, userId: userId };
          } else {
            logSystem("OTP Error", otpResult.error, userId);
            return { success: false, message: "ไม่สามารถส่งอีเมล OTP ได้: " + otpResult.error };
          }
        }
      }
    }
    logSystem("Login Failed", "Invalid credentials", userId);
    return { success: false, message: "UserID หรือ Password ไม่ถูกต้อง!" };
  } catch (error) { return { success: false, message: "ระบบฐานข้อมูลขัดข้อง: " + error.message }; }
}

function apiRequestNewOTP(userId) {
  try {
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.USER);
    const sheet = ss.getSheetByName("User");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(userId)) {
        let email = data[i][3]; let name = data[i][0];
        if (!email) return { success: false, message: "บัญชีนี้ยังไม่ได้ตั้งค่า Email" };
        
        CacheService.getScriptCache().remove("OTP_" + userId);
        let otpResult = generateAndSendOTP(userId, email, name);
        
        if (otpResult.success) {
          logSystem("OTP Resent", "User explicitly requested a new OTP", userId);
          return { success: true, message: "ส่ง OTP รหัสใหม่เรียบร้อยแล้ว" };
        } else { return { success: false, message: "เกิดข้อผิดพลาด: " + otpResult.error }; }
      }
    }
    return { success: false, message: "ไม่พบข้อมูลผู้ใช้งาน" };
  } catch (error) { return { success: false, message: error.message }; }
}

function generateAndSendOTP(userId, email, name) {
  try {
    let otp = Math.floor(100000 + Math.random() * 900000).toString(); 
    CacheService.getScriptCache().put("OTP_" + userId, otp, 300);
    
    let logoUrl = "https://cdn-icons-png.flaticon.com/512/3003/3003251.png"; 
    try {
      const config = getDbConfig();
      if(config.CONFIG && !config.CONFIG.includes('ใส่_ID_ไฟล์')) {
        const configSS = SpreadsheetApp.openById(config.CONFIG);
        const logoSheet = configSS.getSheetByName('App_Logo');
        if (logoSheet && logoSheet.getLastRow() > 1) { logoUrl = logoSheet.getRange(2, 2).getValue() || logoUrl; }
      }
    } catch(e) {}

    const htmlTemplate = `
        <div style="font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif; color: #1e293b; max-width: 600px; margin: 0 auto; padding: 20px;">
            <div style="text-align: center; margin-bottom: 24px;">
                <img src="${logoUrl}" alt="reAgentics Logo" style="width: 56px; height: 56px; border-radius: 12px; object-fit: contain; vertical-align: middle;">
                <span style="font-size: 28px; font-weight: 700; color: #0ea5e9; vertical-align: middle; margin-left: 12px; letter-spacing: -0.5px; display: inline-block;">reAgentics</span>
            </div>
            <h2 style="font-size: 22px; font-weight: 500; text-align: center; margin-bottom: 24px; color: #334155;">กรุณายืนยันตัวตนของคุณ, <strong>${name}</strong></h2>
            <div style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 24px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
                <p style="margin-top: 0; margin-bottom: 16px; font-size: 15px; text-align: center;">นี่คือรหัส OTP สำหรับเข้าสู่ระบบบริหารจัดการน้ำยา:</p>
                <div style="text-align: center; font-size: 36px; font-family: ui-monospace, SFMono-Regular, Consolas, monospace; font-weight: 700; letter-spacing: 10px; color: #0f172a; margin: 28px 0; background-color: #f8fafc; padding: 16px; border-radius: 8px;">${otp}</div>
                <p style="font-size: 14px; margin-bottom: 16px; text-align: center; color: #475569;">รหัสนี้มีอายุการใช้งาน <strong>5 นาที</strong> และใช้ได้เพียงครั้งเดียว</p>
                <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 24px 0;">
                <p style="font-size: 13px; margin-bottom: 12px; color: #64748b;"><strong style="color: #ef4444;">ข้อควรระวัง (PDPA):</strong> โปรดอย่าแชร์รหัสนี้กับบุคคลอื่น ทีมงาน reAgentics จะไม่ขอรหัสผ่านหรือ OTP ของคุณผ่านช่องทางใดๆ โดยเด็ดขาด</p>
            </div>
        </div>
    `;

    MailApp.sendEmail({ to: email, subject: "รหัส OTP สำหรับเข้าสู่ระบบ reAgentics", htmlBody: htmlTemplate, name: "reAgentics LIS" });
    return { success: true };
  } catch (error) { return { success: false, error: error.message }; }
}

function verifyOTP(userId, inputOtp) {
  try {
    checkDatabaseSetup();
    let cache = CacheService.getScriptCache();
    let cachedOtp = cache.get("OTP_" + userId);
    
    if (!cachedOtp) {
      logSystem("Login Failed", "Expired or missing OTP", userId);
      return { success: false, message: "OTP หมดอายุหรือไม่ถูกต้อง กรุณาเข้าสู่ระบบใหม่" };
    }
    
    if (cachedOtp === inputOtp.toString()) {
      cache.remove("OTP_" + userId);
      let userProfile = getUserProfileById(userId);
      if(userProfile) {
        logSystem("Login Success", "User successfully authenticated", userId);
        let availableYears = getAvailableYears();
        return { success: true, message: "เข้าสู่ระบบสำเร็จ", user: userProfile, years: availableYears };
      } else { return { success: false, message: "ไม่พบข้อมูลโปรไฟล์ผู้ใช้งาน" }; }
    } else {
      logSystem("Login Failed", "Invalid OTP entered", userId);
      return { success: false, message: "รหัส OTP ไม่ถูกต้อง" };
    }
  } catch (e) { return { success: false, message: "Verify Error: " + e.message }; }
}

function verifyPasswordOnly(userId, password) {
  try {
    checkDatabaseSetup();
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.USER);
    const sheet = ss.getSheetByName('User');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(userId) && String(data[i][2]) === String(password)) {
        let status = String(data[i][8] || "ปกติ").trim();
        if (status === "ระงับการใช้งาน") {
          logSystem("Unlock Blocked", "Suspended user tried to unlock screen", userId);
          return { success: false, message: `บัญชีผู้ใช้ ${userId} ถูกระงับ โปรดติดต่อผู้ดูแลระบบ` };
        }
        logSystem("Unlock Screen", "Successfully unlocked screen", userId);
        return { success: true };
      }
    }
    logSystem("Unlock Failed", "Invalid password during screen unlock", userId);
    return { success: false, message: 'รหัสผ่านไม่ถูกต้อง' };
  } catch (e) { return { success: false, message: 'System Error: ' + e.message }; }
}

function getUserProfileById(userId) {
  const config = getDbConfig();
  const ss = SpreadsheetApp.openById(config.USER);
  const sheet = ss.getSheetByName("User");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(userId)) { 
      let profile = { name: data[i][0], userId: data[i][1], group: data[i][4], role: data[i][5], unitIdRaw: data[i][6], image: data[i][7] || "", allowedUnits: [] };
      try {
        if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
          const unitSS = SpreadsheetApp.openById(config.UNIT);
          const unitSheet = unitSS.getSheetByName("Units");
          if (unitSheet) {
            const unitData = unitSheet.getDataRange().getValues();
            const role = String(profile.role).toUpperCase();
            const isAdmin = role === 'ADMIN';
            for (let r = 1; r < unitData.length; r++) {
              const uGroup = String(unitData[r][0]).trim();
              const uName = String(unitData[r][1]).trim(); 
              if (isAdmin) { if (uName && !profile.allowedUnits.includes(uName)) profile.allowedUnits.push(uName); } 
              else { if (uGroup === String(profile.group).trim()) { if (uName && !profile.allowedUnits.includes(uName)) profile.allowedUnits.push(uName); } }
            }
          }
        }
      } catch (e) { if (profile.unitIdRaw) profile.allowedUnits = String(profile.unitIdRaw).split(',').map(s => s.trim()); }
      return profile;
    }
  }
  return null;
}

function getAvailableYears() {
  try {
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) return [];
    const dbSS = SpreadsheetApp.openById(config.CONFIG); 
    let sheet = dbSS.getSheetByName('Year_Config');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    let years = [];
    for (let i = 1; i < data.length; i++) {
      if(data[i][0] && (data[i][2] === 'Connected' || !data[i][2])) years.push(String(data[i][0]));
    }
    return years.length > 0 ? years : [];
  } catch (e) { return []; }
}

// -------------------------------------------------------------------------
// 4.5 USER MANAGEMENT API (ADMIN SYSTEM)
// -------------------------------------------------------------------------
function apiGetUsersList(actionByUserId) {
  try {
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.USER);
    const sheet = ss.getSheetByName("User");
    const data = sheet.getDataRange().getValues();
    
    let usersList = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) { 
        usersList.push({
          originalUserId: data[i][1], name: data[i][0], userId: data[i][1], email: data[i][3],
          group: data[i][4], role: data[i][5], unitIdRaw: data[i][6] || '', status: data[i][8] || 'ปกติ', otpStatus: data[i][9] || 'ON'
        });
      }
    }
    return { success: true, data: usersList };
  } catch (error) { return { success: false, message: error.message }; }
}

function apiSaveUserAdmin(payload, actionByUserId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.USER);
    const sheet = ss.getSheetByName("User");
    const data = sheet.getDataRange().getValues();
    
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(payload.originalUserId)) { rowIdx = i + 1; break; }
    }
    
    if (rowIdx === -1) throw new Error("ไม่พบข้อมูลผู้ใช้งานที่ต้องการแก้ไข หรืออาจถูกลบไปแล้ว");
    
    let role = String(payload.role).trim().toUpperCase();
    let unitIdRaw = String(payload.unitIdRaw).trim();
    if (role === 'ADMIN') unitIdRaw = 'ALL'; else if (role === 'USER') unitIdRaw = '';
    
    sheet.getRange(rowIdx, 1).setValue(payload.name); sheet.getRange(rowIdx, 2).setValue(payload.userId);
    sheet.getRange(rowIdx, 4).setValue(payload.email); sheet.getRange(rowIdx, 5).setValue(payload.group);
    sheet.getRange(rowIdx, 6).setValue(role); sheet.getRange(rowIdx, 7).setValue(unitIdRaw);
    sheet.getRange(rowIdx, 9).setValue(payload.status); sheet.getRange(rowIdx, 10).setValue(payload.otpStatus || 'ON'); 
    
    SpreadsheetApp.flush();
    logSystem("Admin Action", `Updated user details for UserID: ${payload.userId}`, actionByUserId);
    return { success: true, message: "บันทึกข้อมูลผู้ใช้งานเรียบร้อยแล้ว" };
  } catch (error) { return { success: false, message: error.message }; } finally { lock.releaseLock(); }
}

// -------------------------------------------------------------------------
// 5. DATABASE CONFIG MANAGER
// -------------------------------------------------------------------------
function apiGetDbConfig() {
  try {
    const config = getDbConfig();
    let years = []; let unitFolders = []; let deliveryNoteFolders = []; let telegramConfigs = [];
    try {
      if(config.CONFIG && !config.CONFIG.includes('ใส่_ID_ไฟล์')) {
        const configSS = SpreadsheetApp.openById(config.CONFIG); 
        let sheet = configSS.getSheetByName('Year_Config');
        if (sheet) {
          const data = sheet.getDataRange().getValues();
          for (let i = 1; i < data.length; i++) {
            if(data[i][0]) {
              let status = data[i][2] || 'Connected'; 
              years.push({ year: String(data[i][0]), fileId: String(data[i][1]), status: status });
            }
          }
        }
      }
    } catch(e) {}

    try {
      if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
         const unitSS = SpreadsheetApp.openById(config.UNIT);
         let unitSheet = unitSS.getSheetByName('Units');
         if(unitSheet) {
             const data = unitSheet.getDataRange().getValues();
             for(let i = 1; i < data.length; i++) {
                 if(data[i][1]) { 
                     let uName = String(data[i][1]).trim();
                     unitFolders.push({ name: uName, folderId: String(data[i][3] || '').trim() });
                     deliveryNoteFolders.push({ name: uName, folderId: String(data[i][4] || '').trim() });
                     telegramConfigs.push({ name: uName, botToken: String(data[i][5] || '').trim(), chatId: String(data[i][6] || '').trim() });
                 }
             }
         }
      }
    } catch(e) {}

    return { 
      success: true, 
      config: { 
        mainId: config.MAIN, unitId: config.UNIT, userId: config.USER, configId: config.CONFIG, logId: config.LOG, 
        folderProfile: config.FOLDER_PROFILE, folderLogo: config.FOLDER_LOGO,
        years: years, unitFolders: unitFolders, deliveryNoteFolders: deliveryNoteFolders, telegramConfigs: telegramConfigs
      } 
    };
  } catch (e) { return { success: false, message: e.message }; }
}

function apiSaveCoreDbConfig(payload, userId) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (payload.mainId !== undefined) props.setProperty('DB_MAIN', payload.mainId.trim());
    if (payload.unitId !== undefined) props.setProperty('DB_UNIT', payload.unitId.trim());
    if (payload.userId !== undefined) props.setProperty('DB_USER', payload.userId.trim());
    if (payload.configId !== undefined) props.setProperty('DB_CONFIG', payload.configId.trim());
    if (payload.logId !== undefined) props.setProperty('DB_LOG', payload.logId.trim());
    if (payload.folderProfile !== undefined) props.setProperty('FOLDER_PROFILE', payload.folderProfile.trim());
    if (payload.folderLogo !== undefined) props.setProperty('FOLDER_LOGO', payload.folderLogo.trim());

    const config = getDbConfig();
    if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
        const unitSS = SpreadsheetApp.openById(config.UNIT);
        let unitSheet = unitSS.getSheetByName('Units');
        if(unitSheet) {
            const data = unitSheet.getDataRange().getValues();
            for(let i = 1; i < data.length; i++) {
                let uName = String(data[i][1]).trim();
                if (payload.unitFolders && payload.unitFolders.length > 0) {
                    let matchImage = payload.unitFolders.find(u => u.name === uName);
                    if(matchImage) unitSheet.getRange(i + 1, 4).setValue(matchImage.folderId); 
                }
                if (payload.deliveryNoteFolders && payload.deliveryNoteFolders.length > 0) {
                    let matchPDF = payload.deliveryNoteFolders.find(u => u.name === uName);
                    if(matchPDF) unitSheet.getRange(i + 1, 5).setValue(matchPDF.folderId); 
                }
                if (payload.telegramConfigs && payload.telegramConfigs.length > 0) {
                    let matchTel = payload.telegramConfigs.find(u => u.name === uName);
                    if(matchTel) {
                        unitSheet.getRange(i + 1, 6).setValue(matchTel.botToken); 
                        unitSheet.getRange(i + 1, 7).setValue(matchTel.chatId); 
                    }
                }
            }
        }
    }

    logSystem("Update DB Config", "Admin updated core database & folder configurations", userId);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function apiCreateYearSheet(year, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); checkDatabaseSetup();
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("กรุณาระบุ Sheet ID สำหรับไฟล์ Config ก่อนสร้างปีงบประมาณ");
    const configSS = SpreadsheetApp.openById(config.CONFIG); 
    let yearSheet = configSS.getSheetByName('Year_Config');
    
    if (!yearSheet) {
      yearSheet = configSS.insertSheet('Year_Config');
      yearSheet.appendRow(['Year', 'FileID', 'Status']);
      yearSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e2e8f0");
      yearSheet.setFrozenRows(1);
    }
    
    const data = yearSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(year)) return { success: false, message: `มีการตั้งค่าไฟล์ของปี ${year}อยู่แล้วในระบบครับ` };
    }
    
    const fileName = `reAgentics_Transactions_${year}`;
    const newSS = SpreadsheetApp.create(fileName);
    const fileId = newSS.getId();
    let sheet = newSS.getSheets()[0];
    sheet.setName(String(year));
    const headers = ['transactionID', 'timestamp', 'type', 'itemID', 'lot', 'expiry_Date', 'quantity', 'actionBy_UserID', 'Transport_Temp', 'Transport_Speed', 'Delivery_Note_URL'];
    sheet.appendRow(headers);
    sheet.getRange("A1:K1").setFontWeight("bold").setBackground("#f8fafc");
    sheet.setFrozenRows(1);
    
    yearSheet.appendRow([year, fileId, 'Connected']);
    logSystem("Create Year Sheet", `Created new transaction file for year ${year} (ID: ${fileId})`, userId);
    return { success: true, fileId: fileId, year: year, status: 'Connected' };
  } catch (e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

function apiManualAddYearSheet(year, fileId, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); checkDatabaseSetup();
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("Missing Config DB ID");
    const configSS = SpreadsheetApp.openById(config.CONFIG); 
    let yearSheet = configSS.getSheetByName('Year_Config');
    
    if (!yearSheet) {
      yearSheet = configSS.insertSheet('Year_Config');
      yearSheet.appendRow(['Year', 'FileID', 'Status']);
      yearSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e2e8f0");
      yearSheet.setFrozenRows(1);
    }
    
    const data = yearSheet.getDataRange().getValues();
    let isExist = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(year)) {
        yearSheet.getRange(i + 1, 2).setValue(fileId); yearSheet.getRange(i + 1, 3).setValue('Connected'); isExist = true; break;
      }
    }
    if (!isExist) yearSheet.appendRow([year, fileId, 'Connected']);

    try { SpreadsheetApp.openById(fileId); } catch(err) { throw new Error("ไม่สามารถเข้าถึงไฟล์ Sheet ID ที่ระบุได้ กรุณาตรวจสอบสิทธิ์การเข้าถึง"); }
    
    logSystem("Manual Connect Year", `Manually connected file for year ${year} (ID: ${fileId})`, userId);
    return { success: true, fileId: fileId, year: year, status: 'Connected' };
  } catch (e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

function apiDisconnectYearSheet(year, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("Missing Config DB ID");
    const configSS = SpreadsheetApp.openById(config.CONFIG); 
    let yearSheet = configSS.getSheetByName('Year_Config');
    if (!yearSheet) return { success: false, message: 'ไม่พบตารางตั้งค่าปี' };

    const data = yearSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(year)) {
        yearSheet.getRange(i + 1, 3).setValue('Disconnected'); 
        logSystem("Disconnect Year", `Disconnected transaction file for year ${year}`, userId);
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลปีที่ต้องการระงับการเชื่อมต่อ' };
  } catch (e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

function apiDeleteYearSheet(year, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("Missing Config DB ID");
    const configSS = SpreadsheetApp.openById(config.CONFIG); 
    let yearSheet = configSS.getSheetByName('Year_Config');
    if (!yearSheet) return { success: false, message: 'ไม่พบตารางตั้งค่าปี' };

    const data = yearSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(year)) {
        yearSheet.deleteRow(i + 1); 
        logSystem("Delete Year Link", `Removed year ${year} from config database`, userId);
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลปีที่ต้องการลบ' };
  } catch (e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

// -------------------------------------------------------------------------
// 6. IMAGE, PROFILE & PDF APIs
// -------------------------------------------------------------------------
function apiGetSystemLogo() {
  try {
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) return { success: false };
    const sysSS = SpreadsheetApp.openById(config.CONFIG); 
    let sheet = sysSS.getSheetByName('App_Logo');
    if (sheet) {
      let url = sheet.getRange("B2").getValue();
      if (!url) url = sheet.getRange("B1").getValue();
      if (url) return { success: true, url: url };
    }
    return { success: false };
  } catch(e) { return { success: false }; }
}

function apiChangePassword(userId, newPassword) {
  try {
    const config = getDbConfig();
    const userSS = SpreadsheetApp.openById(config.USER); 
    const sheet = userSS.getSheetByName('User'); 
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { 
      if (String(data[i][1]) === String(userId)) { 
        sheet.getRange(i + 1, 3).setValue(newPassword); 
        logSystem("Change Password", "Updated password", userId); 
        return { status: 'success' }; 
      } 
    }
    return { status: 'error', message: 'User not found' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function apiSaveProfileImage(userId, base64Data) {
  try {
    const config = getDbConfig();
    const userSS = SpreadsheetApp.openById(config.USER); 
    const sheet = userSS.getSheetByName('User'); 
    const data = sheet.getDataRange().getDisplayValues();
    
    let rowIndex = -1; let oldFileUrl = "";
    for (let i = 1; i < data.length; i++) { 
      if (String(data[i][1]) === String(userId)) { rowIndex = i + 1; oldFileUrl = data[i][7]; break; } 
    }
    if (rowIndex === -1) return { status: 'error', message: 'User not found' };
    deleteOldDriveFile(oldFileUrl);
    
    let folder;
    if (config.FOLDER_PROFILE) { try { folder = DriveApp.getFolderById(config.FOLDER_PROFILE); } catch(e) {} }
    if (!folder) {
        const folderName = "reAgentics_Profiles";
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) folder = folders.next(); else folder = DriveApp.createFolder(folderName);
    }
    
    const contentType = base64Data.substring(5, base64Data.indexOf(';')); 
    let ext = "png"; if (contentType.includes("gif")) ext = "gif"; else if (contentType.includes("jpeg") || contentType.includes("jpg")) ext = "jpg";
    const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7));
    const blob = Utilities.newBlob(bytes, contentType, `profile_${userId}_${Date.now()}.${ext}`); 
    const file = folder.createFile(blob); 
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    
    const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s800`; 
    sheet.getRange(rowIndex, 8).setValue(fileUrl); 
    logSystem("Change Profile Pic", "Updated profile image", userId); 
    return { status: 'success', url: fileUrl };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function apiSaveSystemLogo(base64Data, userId) {
  try {
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("กรุณาตั้งค่า Config ID ในหน้าตั้งค่าฐานข้อมูลก่อน");
    const sysSS = SpreadsheetApp.openById(config.CONFIG); 
    let sheet = sysSS.getSheetByName('App_Logo');
    if (!sheet) { sheet = sysSS.insertSheet('App_Logo'); }
    
    let oldFileUrl = "";
    try { oldFileUrl = sheet.getRange("B2").getValue() || sheet.getRange("B1").getValue(); } catch(e) {}
    deleteOldDriveFile(oldFileUrl);

    let folder;
    if (config.FOLDER_LOGO) { try { folder = DriveApp.getFolderById(config.FOLDER_LOGO); } catch(e) {} }
    if (!folder) {
        const folderName = "reAgentics_Logos";
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) folder = folders.next(); else folder = DriveApp.createFolder(folderName);
    }
    
    const contentType = base64Data.substring(5, base64Data.indexOf(';')); 
    let ext = "png"; if (contentType.includes("gif")) ext = "gif"; else if (contentType.includes("jpeg") || contentType.includes("jpg")) ext = "jpg";
    const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7)); 
    const blob = Utilities.newBlob(bytes, contentType, `app_logo_${Date.now()}.${ext}`);
    const file = folder.createFile(blob); 
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    
    const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s800`;
    sheet.clear();
    sheet.getRange("A1").setValue("Name").setFontWeight("bold"); sheet.getRange("B1").setValue("Url").setFontWeight("bold");
    sheet.getRange("A2").setValue("MainLogo"); sheet.getRange("B2").setValue(fileUrl);
    SpreadsheetApp.flush();
    logSystem("Change Logo", "Updated system logo", userId); 
    return { status: 'success', url: fileUrl };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function apiUploadReagentImage(base64Data, unitName, oldUrl, userId) {
    try {
        const config = getDbConfig();
        let targetFolderId = "";
        
        if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
             const unitSS = SpreadsheetApp.openById(config.UNIT);
             let unitSheet = unitSS.getSheetByName('Units');
             if(unitSheet) {
                 const data = unitSheet.getDataRange().getValues();
                 for(let i = 1; i < data.length; i++) {
                     if(String(data[i][1]).trim() === String(unitName).trim()) {
                         targetFolderId = String(data[i][3] || '').trim(); break;
                     }
                 }
             }
        }
        deleteOldDriveFile(oldUrl);

        let folder;
        if (targetFolderId) { try { folder = DriveApp.getFolderById(targetFolderId); } catch(e) {} }
        if (!folder) {
            const folderName = "reAgentics_Items_" + unitName;
            const folders = DriveApp.getFoldersByName(folderName);
            if (folders.hasNext()) folder = folders.next(); else folder = DriveApp.createFolder(folderName);
        }

        const contentType = base64Data.substring(5, base64Data.indexOf(';')); 
        let ext = "png"; if (contentType.includes("gif")) ext = "gif"; else if (contentType.includes("jpeg") || contentType.includes("jpg")) ext = "jpg";
        const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7));
        const blob = Utilities.newBlob(bytes, contentType, `item_${Date.now()}.${ext}`); 
        const file = folder.createFile(blob); 
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
        
        const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s800`; 
        logSystem("Upload Reagent Image", `Uploaded image to folder for unit: ${unitName}`, userId);
        return { success: true, url: fileUrl };
    } catch(e) { return { success: false, message: e.toString() }; }
}

function apiUploadDeliveryNote(base64Data, unitName, userId) {
    try {
        const config = getDbConfig();
        let targetFolderId = ""; let prefix = unitName; 
        
        if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
             const unitSS = SpreadsheetApp.openById(config.UNIT);
             let unitSheet = unitSS.getSheetByName('Units');
             if(unitSheet) {
                 const data = unitSheet.getDataRange().getValues();
                 for(let i = 1; i < data.length; i++) {
                     if(String(data[i][1]).trim() === String(unitName).trim()) {
                         prefix = String(data[i][2]).trim() || unitName; 
                         targetFolderId = String(data[i][4] || '').trim(); break;
                     }
                 }
             }
        }
        
        let folder;
        if (targetFolderId) { try { folder = DriveApp.getFolderById(targetFolderId); } catch(e) {} }
        if (!folder) {
            const folderName = "reAgentics_DeliveryNotes_" + unitName;
            const folders = DriveApp.getFoldersByName(folderName);
            if (folders.hasNext()) folder = folders.next(); else folder = DriveApp.createFolder(folderName);
        }

        const today = new Date();
        const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
        const randomStr = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
        const fileName = `${prefix}_${dateStr}_${randomStr}.pdf`;

        const contentType = 'application/pdf';
        const base64Clean = base64Data.split(',')[1];
        const bytes = Utilities.base64Decode(base64Clean);
        const blob = Utilities.newBlob(bytes, contentType, fileName); 
        const file = folder.createFile(blob); 
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
        
        const fileUrl = file.getUrl(); 
        logSystem("Upload Delivery Note", `Uploaded PDF to folder for unit: ${unitName}`, userId);
        return { success: true, url: fileUrl };
    } catch(e) { return { success: false, message: e.toString() }; }
}

// -------------------------------------------------------------------------
// 7. STICKER CONFIG API
// -------------------------------------------------------------------------
function apiGetStickerConfig() {
  try {
    const dbConfig = getDbConfig();
    if(!dbConfig.CONFIG || dbConfig.CONFIG.includes('ใส่_ID_ไฟล์')) {
      return { status: 'success', config: getDefaultStickerConfig() };
    }

    const sysSS = SpreadsheetApp.openById(dbConfig.CONFIG);
    let sheet = sysSS.getSheetByName('Sticker_Config');
    let config = getDefaultStickerConfig();

    if (!sheet) {
      sheet = sysSS.insertSheet('Sticker_Config');
      sheet.appendRow(['Key', 'Value', 'Description']);
      sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f1f5f9");
      
      const descriptions = {
        width: "ความกว้างของสติ๊กเกอร์ (mm)", height: "ความสูงของสติ๊กเกอร์ (mm)",
        autoPrintCount: "จำนวนแผ่นที่จะพิมพ์อัตโนมัติเมื่อรับเข้าเสร็จ", manualPrintCount: "จำนวนแผ่นที่จะพิมพ์เมื่อกดปุ่มพิมพ์จากหน้าจอ",
        barcodeHeight: "ความสูงของเส้นบาร์โค้ด (px)", barcodeWidth: "สเกลความกว้างเส้นบาร์โค้ด", layoutJSON: "พิกัด X/Y, ขนาด, การหมุน ของแต่ละองค์ประกอบ (ห้ามแก้ไขด้วยมือ)"
      };

      for (let key in config) { sheet.appendRow([key, config[key], descriptions[key]]); }
      sheet.setColumnWidth(1, 150); sheet.setColumnWidth(2, 250); sheet.setColumnWidth(3, 300);
    } else {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (config.hasOwnProperty(data[i][0])) {
          let val = data[i][1];
          if (val === 'true' || val === true) config[data[i][0]] = true;
          else if (val === 'false' || val === false) config[data[i][0]] = false;
          else if (data[i][0] === 'layoutJSON') config[data[i][0]] = String(val);
          else config[data[i][0]] = Number(val) || val;
        }
      }
    }
    return { status: 'success', config: config };
  } catch (e) { return { status: 'error', message: 'Get Sticker Config Error: ' + e.message }; }
}

function getDefaultStickerConfig() {
  return {
    width: 50, height: 30, autoPrintCount: 2, manualPrintCount: 1, barcodeHeight: 35, barcodeWidth: 1.5,   
    layoutJSON: JSON.stringify({
      cyto:    { x: 25, y: 4, size: 11, rot: 0, visible: true, bold: true, font: 'Montserrat' }, 
      name:    { x: 25, y: 9, size: 9, rot: 0, visible: false, bold: false, font: 'Montserrat' }, 
      age:     { x: 10, y: 26, size: 10, rot: 0, visible: true, bold: true, font: 'Roboto Mono' }, 
      spec:    { x: 40, y: 26, size: 10, rot: 0, visible: true, bold: true, font: 'Roboto Mono' }, 
      unit:    { x: 25, y: 28, size: 8, rot: 0, visible: false, bold: false, font: 'Montserrat' }, 
      bar:     { x: 25, y: 13, rot: 0, visible: true, width: 1.5 },
      barText: { x: 25, y: 21, size: 11, rot: 0, visible: true, bold: true, font: 'Roboto Mono' } 
    })
  };
}

function apiSaveStickerConfig(newConfig, userId) {
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try {
    const dbConfig = getDbConfig();
    if(!dbConfig.CONFIG || dbConfig.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("กรุณาตั้งค่า Config ID ในหน้าตั้งค่าฐานข้อมูลก่อน");

    const sysSS = SpreadsheetApp.openById(dbConfig.CONFIG);
    let sheet = sysSS.getSheetByName('Sticker_Config');
    if (!sheet) return { status: 'error', message: 'Sticker_Config sheet not found' };

    const data = sheet.getDataRange().getValues();
    for (let key in newConfig) {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(newConfig[key]);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow([key, newConfig[key], "Auto-generated field"]);
    }
    logSystem("Update Config", "Admin updated Sticker Configuration Layout", userId);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: 'Save Sticker Config Error: ' + e.message }; } finally { lock.releaseLock(); }
}

// -------------------------------------------------------------------------
// 8. DATA FETCHING (DROPDOWNS & AUTO-IDs) 
// -------------------------------------------------------------------------
function apiGetFormOptions() {
  try {
    const config = getDbConfig();
    if(!config.UNIT || config.UNIT.includes('ใส่_ID_ไฟล์')) return { success: false, message: 'ไม่ได้ตั้งค่า Unit DB' };
    
    const unitSS = SpreadsheetApp.openById(config.UNIT);
    const options = { units: [], reagUnits: [], analyzers: [], storageLocations: [], companies: [], reagTypes: [] };
    
    let sheet = unitSS.getSheetByName('Units');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) {
        if(data[i][1]) {
          options.units.push({ group: String(data[i][0]).trim(), name: String(data[i][1]).trim(), prefix: String(data[i][2]).trim() });
        }
      }
    }
    
    sheet = unitSS.getSheetByName('ReagUnits');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) { if(data[i][0]) options.reagUnits.push(data[i][0]); }
    }
    
    // NEW: Fetch Reagent Types from ReagTypes tab
    let typeSheet = unitSS.getSheetByName('ReagTypes');
    if (typeSheet) {
      const data = typeSheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) { if(data[i][0]) options.reagTypes.push(data[i][0]); }
    }
    
    sheet = unitSS.getSheetByName('Analyzers');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) { if(data[i][0] && data[i][1]) options.analyzers.push({ unit: data[i][0], name: data[i][1] }); }
    }
    
    sheet = unitSS.getSheetByName('storageLocation');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) { if(data[i][0]) options.storageLocations.push(data[i][0]); }
    }

    let companySheet = unitSS.getSheetByName('Company');
    if (companySheet) {
      const data = companySheet.getRange("A2:A" + companySheet.getLastRow()).getValues();
      const uniqueCompanies = [...new Set(data.map(r => String(r[0]).trim()).filter(String))];
      options.companies = uniqueCompanies;
    }
    return { success: true, data: options };
  } catch(e) { return { success: false, message: e.message }; }
}

function apiGetNextItemID(unitName) {
  try {
    const config = getDbConfig();
    const unitSS = SpreadsheetApp.openById(config.UNIT);
    const unitSheet = unitSS.getSheetByName('Units');
    if (!unitSheet) throw new Error("ไม่พบแท็บ Units ในฐานข้อมูลหน่วยงาน");
    
    const unitData = unitSheet.getDataRange().getValues();
    let prefix = "";
    for(let i=1; i<unitData.length; i++) {
      if(String(unitData[i][1]).trim() === String(unitName).trim()) { prefix = String(unitData[i][2]).trim(); break; }
    }
    if(!prefix) return { success: false, message: "ไม่พบรหัส Prefix สำหรับหน่วยงานนี้" };
    
    const mainSS = SpreadsheetApp.openById(config.MAIN);
    const itemSheet = mainSS.getSheetByName('Items');
    let maxNum = 0;
    
    if (itemSheet) {
      const itemData = itemSheet.getDataRange().getValues();
      for(let i=1; i<itemData.length; i++) {
        const id = String(itemData[i][0]).trim();
        if(id.startsWith(prefix + "-")) {
          const numPart = parseInt(id.replace(prefix + "-", ""), 10);
          if(!isNaN(numPart) && numPart > maxNum) maxNum = numPart;
        }
      }
    }
    const nextId = prefix + "-" + String(maxNum + 1).padStart(3, '0');
    return { success: true, nextId: nextId };
  } catch(e) { return { success: false, message: e.message }; }
}

function apiGetActiveLots(itemID) {
  try {
    const config = getDbConfig();
    const mainSS = SpreadsheetApp.openById(config.MAIN);
    const stockSheet = mainSS.getSheetByName('Stock_Balance');
    if(!stockSheet) return { success: true, lots: [] };
    
    const data = stockSheet.getDataRange().getValues();
    const lots = [];
    for(let i=1; i<data.length; i++) {
      if(String(data[i][0]).trim() === String(itemID).trim() && Number(data[i][4]) >= 1) {
        lots.push({ lot: String(data[i][2]), exp: safeString(data[i][3]), qty: Number(data[i][4]), unit: String(data[i][5]) });
      }
    }
    return { success: true, lots: lots };
  } catch(e) { return { success: false, message: e.message }; }
}

// -------------------------------------------------------------------------
// 9. INVENTORY & TRANSACTION ENGINE 
// -------------------------------------------------------------------------
function getItemsData(yearSheetId, dashboardFilterMonth) {
  try {
    checkDatabaseSetup();
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.MAIN);
    
    let unitGroupMap = {};
    try {
      if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
        const unitSS = SpreadsheetApp.openById(config.UNIT);
        const unitSheet = unitSS.getSheetByName('Units');
        if (unitSheet) {
          const unitData = unitSheet.getDataRange().getValues();
          for (let i=1; i<unitData.length; i++) {
             let group = String(unitData[i][0]).trim(); let uName = String(unitData[i][1]).trim();
             if(uName) unitGroupMap[uName] = group;
          }
        }
      }
    } catch(e) {}

    const itemSheet = ss.getSheetByName("Items");
    if (!itemSheet) throw new Error("ไม่พบชีต 'Items' ในฐานข้อมูลหลัก");
    const itemData = itemSheet.getDataRange().getDisplayValues();
    
    let stockSheet = ss.getSheetByName("Stock_Balance");
    let stockData = [];
    if (stockSheet) { stockData = stockSheet.getDataRange().getValues(); }
    
    const today = new Date();
    today.setHours(0,0,0,0);
    const balanceMap = {};
    if (stockData.length > 1) {
      for (let r = 1; r < stockData.length; r++) {
        let itemId = String(stockData[r][0]).trim();
        let qty = Number(stockData[r][4]) || 0;
        let expStr = stockData[r][3];
        
        if (!balanceMap[itemId]) balanceMap[itemId] = { totalQty: 0, earliestExp: Infinity };
        if(qty > 0) {
            balanceMap[itemId].totalQty += qty;
            if (expStr && expStr !== '-') {
                let expDate = new Date(expStr).getTime();
                if (!isNaN(expDate) && expDate < balanceMap[itemId].earliestExp) {
                    balanceMap[itemId].earliestExp = expDate;
                }
            }
        }
      }
    }

    let txReportMap = {};
    if (yearSheetId && config.CONFIG && !config.CONFIG.includes('ใส่_ID_ไฟล์')) {
        try {
            const configSS = SpreadsheetApp.openById(config.CONFIG);
            let yearSheet = configSS.getSheetByName('Year_Config');
            if (yearSheet) {
                let transFileId = "";
                let yData = yearSheet.getDataRange().getValues();
                for (let i = 1; i < yData.length; i++) {
                    if (String(yData[i][0]) === String(yearSheetId) && yData[i][2] !== 'Disconnected') { transFileId = yData[i][1]; break; }
                }
                
                if (transFileId) {
                    const transSS = SpreadsheetApp.openById(transFileId);
                    let tSheet = transSS.getSheetByName(String(yearSheetId));
                    if (tSheet && tSheet.getLastRow() > 1) {
                        const tData = tSheet.getDataRange().getValues();
                        for (let r = 1; r < tData.length; r++) {
                            let timestamp = new Date(tData[r][1]);
                            let tType = String(tData[r][2]).trim().toUpperCase();
                            let tItemId = String(tData[r][3]).trim();
                            let tQty = Number(tData[r][6]) || 0;
                            
                            let includeRow = true;
                            if (dashboardFilterMonth && dashboardFilterMonth !== 'All') {
                                let filterM = Number(dashboardFilterMonth);
                                if (timestamp.getMonth() + 1 !== filterM) { includeRow = false; }
                            }
                            
                            if (includeRow) {
                                if(!txReportMap[tItemId]) txReportMap[tItemId] = { rx: 0, disp: 0 };
                                if(tType === 'RECEIVE') txReportMap[tItemId].rx += tQty;
                                else if(tType === 'DISPENSE') txReportMap[tItemId].disp += tQty;
                            }
                        }
                    }
                }
            }
        } catch(e) { console.error("Error generating report data", e); }
    }
    
    const resultData = [];
    const reportData = [];

    // UPDATE FOR SHIFTED COLUMNS: ReagType is now at index 5. Images at 10, etc.
    for (let i = 1; i < itemData.length; i++) {
      let row = itemData[i];
      let itemId = String(row[0]).trim();
      if (!itemId) continue; 
      
      let currentBalance = balanceMap[itemId] ? balanceMap[itemId].totalQty : 0;
      let earliestExp = balanceMap[itemId] ? balanceMap[itemId].earliestExp : Infinity;
      
      let expStatus = 'Active';
      if (earliestExp !== Infinity) {
          const diffDays = (earliestExp - today.getTime()) / (1000 * 60 * 60 * 24);
          if (diffDays < 0) expStatus = 'Expired';
          else if (diffDays <= 60) expStatus = 'Expiring'; 
      } else if (currentBalance === 0) { expStatus = '-'; }

      let uName = String(row[4]).trim();
      let group = unitGroupMap[uName] || 'ไม่ระบุ';
      
      let itemObj = {
        itemID: itemId, itemName: row[1], minLevel: row[2], unit: row[3], unitID: uName, group: group,
        reagType: String(row[5] || '').trim(), // Fetched new Reagent Type column
        analyzer: row[6], storageTemp: row[7], storageLocation: row[8], status: String(row[9]).trim(),
        expStatus: expStatus, image: String(row[10] || '').trim(), company: String(row[11] || '').trim(), price: Number(row[12]) || 0, balance: currentBalance 
      };
      resultData.push(itemObj);

      if (yearSheetId) {
          reportData.push({ ...itemObj, receiveSum: txReportMap[itemId] ? txReportMap[itemId].rx : 0, dispenseSum: txReportMap[itemId] ? txReportMap[itemId].disp : 0 });
      }
    }
    return { success: true, data: resultData, reportData: reportData };
  } catch (error) { return { success: false, message: error.message }; }
}

function apiGetTransactionLogs(payload) {
  try {
    const { yearSheetId, startDate, endDate } = payload;
    const config = getDbConfig();
    
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("Missing Config DB ID");
    const configSS = SpreadsheetApp.openById(config.CONFIG); 
    let yearSheet = configSS.getSheetByName('Year_Config');
    if(!yearSheet) throw new Error("ไม่พบตารางตั้งค่าปีในระบบ");

    let transFileId = "";
    let yData = yearSheet.getDataRange().getValues();
    for (let i = 1; i < yData.length; i++) {
        if (String(yData[i][0]) === String(yearSheetId) && yData[i][2] !== 'Disconnected') { transFileId = yData[i][1]; break; }
    }
    if (!transFileId) throw new Error("ไม่พบไฟล์ฐานข้อมูลประวัติการทำรายการสำหรับปี " + yearSheetId);

    const transSS = SpreadsheetApp.openById(transFileId);
    const tSheet = transSS.getSheetByName(String(yearSheetId));
    if (!tSheet) return { success: true, logs: [] };

    const mainSS = SpreadsheetApp.openById(config.MAIN);
    const iSheet = mainSS.getSheetByName("Items");
    const iData = iSheet.getDataRange().getValues();
    let itemMap = {};
    for(let r=1; r<iData.length; r++) {
        itemMap[String(iData[r][0]).trim()] = { 
            unitID: String(iData[r][4]).trim(),
            unit: String(iData[r][3]).trim(), 
            image: String(iData[r][10] || '').trim() // Grab image from shifted column K (index 10)
        };
    }

    let stockMap = {};
    const stockSheet = mainSS.getSheetByName("Stock_Balance");
    if(stockSheet) {
        const sData = stockSheet.getDataRange().getValues();
        for(let r=1; r<sData.length; r++) {
            let sId = String(sData[r][0]).trim();
            let sLot = String(sData[r][2]).trim().toUpperCase();
            stockMap[sId + '|' + sLot] = {
                exp: sData[r][3],
                unit: String(sData[r][5]).trim()
            };
        }
    }

    let unitGroupMap = {};
    if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
        const unitSS = SpreadsheetApp.openById(config.UNIT);
        const unitSheet = unitSS.getSheetByName('Units');
        if (unitSheet) {
          const unitData = unitSheet.getDataRange().getValues();
          for (let i=1; i<unitData.length; i++) { unitGroupMap[String(unitData[i][1]).trim()] = String(unitData[i][0]).trim(); }
        }
    }

    const tData = tSheet.getDataRange().getValues();
    const result = [];
    let startMs = startDate ? new Date(startDate).setHours(0,0,0,0) : 0;
    let endMs = endDate ? new Date(endDate).setHours(23,59,59,999) : Infinity;

    for (let i = tData.length - 1; i >= 1; i--) {
        let tsStr = tData[i][1];
        let timeMs = 0;
        if (typeof tsStr === 'string' && tsStr.includes('/')) {
            let parts = tsStr.split(' ');
            let dateParts = parts[0].split('/');
            if(dateParts.length === 3) {
                timeMs = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]).getTime();
            }
        } else {
             let timestamp = new Date(tData[i][1]);
             timeMs = timestamp.getTime();
        }

        if (timeMs >= startMs && timeMs <= endMs) {
            let itemId = String(tData[i][3]).trim();
            let lotId = String(tData[i][4]).trim().toUpperCase();
            let itemDetail = itemMap[itemId] || { unitID: 'Unknown', unit: '', image: '' };
            let unitID = itemDetail.unitID;
            let group = unitGroupMap[unitID] || 'ไม่ระบุ';

            let foundName = "";
            for(let k=1; k<iData.length; k++){ if(String(iData[k][0]).trim() === itemId) { foundName = String(iData[k][1]); break;} }

            let sInfo = stockMap[itemId + '|' + lotId] || {};
            let finalExp = sInfo.exp || tData[i][5]; 
            let finalUnit = sInfo.unit || itemDetail.unit || '';
            
            let displayTimestamp = "";
            if (typeof tsStr === 'string' && tsStr.includes('/')) {
                displayTimestamp = tsStr; 
            } else {
                let d = new Date(tsStr);
                displayTimestamp = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
            }

            result.push({
                transId: tData[i][0],
                timestamp: displayTimestamp,
                type: tData[i][2],
                itemID: itemId,
                itemName: foundName || itemId,
                lot: lotId,
                exp: safeString(finalExp),
                qty: tData[i][6],
                unit: finalUnit, 
                userId: tData[i][7],
                unitID: unitID,
                group: group,
                image: itemDetail.image, // Export Image to frontend for Logs
                transportTemp: tData[i][8],
                transportSpeed: tData[i][9]
            });
        }
    }
    return { success: true, logs: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function apiRegisterNewItem(payload, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.MAIN);
    const sheet = ss.getSheetByName("Items");
    
    const existingData = sheet.getRange("A:A").getValues().flat();
    if(existingData.includes(payload.itemID)) throw new Error(`รหัสน้ำยา ${payload.itemID} มีอยู่ในระบบแล้ว กรุณาลองใหม่อีกครั้ง`);
    
    // UPDATED WITH REAGTYPE at position 5
    sheet.appendRow([ payload.itemID, payload.itemName, payload.minLevel, payload.unit, payload.unitID, payload.reagType, payload.analyzer, payload.storageTemp, payload.storageLocation, 'Active', payload.image || '', payload.company || '', payload.price || 0 ]);
    SpreadsheetApp.flush(); 
    logSystem("Register Item", `Registered new item: ${payload.itemID}`, userId);
    return { success: true, message: 'ลงทะเบียนน้ำยาใหม่เรียบร้อยแล้ว' };
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

function apiUpdateItem(payload, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.MAIN);
    const sheet = ss.getSheetByName("Items");
    const data = sheet.getDataRange().getValues();
    
    let targetRow = -1;
    for(let i=1; i<data.length; i++) { if(String(data[i][0]).trim() === String(payload.itemID).trim()) { targetRow = i + 1; break; } }
    if(targetRow === -1) throw new Error("ไม่พบรายการที่ต้องการแก้ไข");
    
    // SHIFTED COLUMNS FOR UPDATE
    sheet.getRange(targetRow, 2).setValue(payload.itemName); sheet.getRange(targetRow, 3).setValue(payload.minLevel);
    sheet.getRange(targetRow, 4).setValue(payload.unit); sheet.getRange(targetRow, 5).setValue(payload.unitID);
    sheet.getRange(targetRow, 6).setValue(payload.reagType); // NEW
    sheet.getRange(targetRow, 7).setValue(payload.analyzer); sheet.getRange(targetRow, 8).setValue(payload.storageTemp);
    sheet.getRange(targetRow, 9).setValue(payload.storageLocation); sheet.getRange(targetRow, 10).setValue(payload.status);
    
    if (payload.image !== undefined) { sheet.getRange(targetRow, 11).setValue(payload.image); }
    sheet.getRange(targetRow, 12).setValue(payload.company || ''); sheet.getRange(targetRow, 13).setValue(payload.price || 0);
    
    SpreadsheetApp.flush();
    logSystem("Update Item", `Updated details for item: ${payload.itemID} | Status: ${payload.status}`, userId);
    return { success: true, message: 'อัปเดตข้อมูลน้ำยาเรียบร้อยแล้ว' };
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

// -------------------------------------------------------------------------
// OPTIMIZED BATCH TRANSACTION PROCESSING
// -------------------------------------------------------------------------
function processTransaction(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const config = getDbConfig();
    const { type, yearSheetId, userId, items, transportSpeed, deliveryNoteUrl } = payload;
    const dbSS = SpreadsheetApp.openById(config.MAIN);
    
    let stockSheet = dbSS.getSheetByName('Stock_Balance');
    if (!stockSheet) {
      stockSheet = dbSS.insertSheet('Stock_Balance');
      stockSheet.appendRow(['ItemID', 'ItemName', 'Lot', 'EXP', 'Qty', 'Unit', 'LastUpdate']);
      stockSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#f8fafc");
      stockSheet.setFrozenRows(1);
    }
    let stockData = stockSheet.getDataRange().getValues();
    
    const itemSheet = dbSS.getSheetByName('Items');
    const itemDataArr = itemSheet.getDataRange().getValues();
    let itemToUnitMap = {};
    for(let i=1; i<itemDataArr.length; i++) { 
      itemToUnitMap[String(itemDataArr[i][0]).trim()] = String(itemDataArr[i][4]).trim(); 
    }

    // -- เตรียมไฟล์ Log ล่วงหน้า (เปิดไฟล์แค่รอบเดียว) --
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("Missing Config DB ID");
    const configSS = SpreadsheetApp.openById(config.CONFIG);
    let yearSheet = configSS.getSheetByName('Year_Config');
    if(!yearSheet) throw new Error("ไม่พบตารางตั้งค่าปี");
    
    let transFileId = "";
    let yData = yearSheet.getDataRange().getValues();
    for (let i = 1; i < yData.length; i++) {
      if (String(yData[i][0]) === String(yearSheetId)) { 
        if(yData[i][2] === 'Disconnected') throw new Error("ไฟล์ปี " + yearSheetId + " ถูกระงับการเชื่อมต่อ กรุณาเชื่อมต่อก่อนทำรายการ");
        transFileId = yData[i][1]; break; 
      }
    }
    if (!transFileId) throw new Error("ไม่พบไฟล์ Transactions สำหรับปี " + yearSheetId);
    
    const transSS = SpreadsheetApp.openById(transFileId);
    let logSheet = transSS.getSheetByName(String(yearSheetId));
    if (!logSheet) {
      logSheet = transSS.insertSheet(String(yearSheetId));
      logSheet.appendRow(['transactionID', 'timestamp', 'type', 'itemID', 'lot', 'expiry_Date', 'quantity', 'actionBy_UserID', 'Transport_Temp', 'Transport_Speed', 'Delivery_Note_URL']);
      logSheet.getRange("A1:K1").setFontWeight("bold").setBackground("#f8fafc");
      logSheet.setFrozenRows(1);
    }

    const timestamp = new Date();
    let timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    let txSummaryForTelegram = {}; 
    let logRowsToAppend = [];

    // -- จัดการอัปเดต Stock และเตรียมข้อมูล Log --
    for (let i = 0; i < items.length; i++) {
      let item = items[i];
      let reqItemId = String(item.itemID).trim();
      let reqLot = String(item.lot).trim().toUpperCase(); 
      let reqQty = Number(item.qty);
      let targetUnit = itemToUnitMap[reqItemId] || 'Unknown';

      let rowToUpdate = -1; let currentQty = 0;
      
      // ค้นหา Lot ในคลัง
      for (let r = 1; r < stockData.length; r++) {
        let sheetItemId = String(stockData[r][0]).trim();
        let sheetLot = String(stockData[r][2]).trim().toUpperCase();
        if (sheetItemId === reqItemId && sheetLot === reqLot) { 
          rowToUpdate = r + 1; 
          currentQty = Number(stockData[r][4]) || 0; 
          
          if (type === 'DISPENSE') {
            item.exp = stockData[r][3];
            item.unit = stockData[r][5];
          }
          break; 
        }
      }

      if (!txSummaryForTelegram[targetUnit]) txSummaryForTelegram[targetUnit] = [];
      txSummaryForTelegram[targetUnit].push(`- ${item.itemName} (${reqLot}) x ${reqQty} ${item.unit || ''}`);
      
      let newQty = currentQty;
      if (type === 'RECEIVE') {
        newQty = currentQty + reqQty;
        if (rowToUpdate === -1) {
          stockSheet.appendRow([item.itemID, item.itemName, reqLot, item.exp, newQty, item.unit, timeStr]);
        } else { 
          stockSheet.getRange(rowToUpdate, 5).setValue(newQty); 
          stockSheet.getRange(rowToUpdate, 7).setValue(timeStr); 
        }
      } else if (type === 'DISPENSE') {
        if (rowToUpdate === -1) throw new Error(`ไม่พบข้อมูล Lot: ${reqLot} ของน้ำยารหัส ${reqItemId} ในคลัง กรุณาตรวจสอบให้แน่ใจ`);
        if (currentQty < reqQty) throw new Error(`ยอดคงเหลือของ ${reqItemId} (Lot: ${reqLot}) ไม่เพียงพอ (มีอยู่ ${currentQty} แต่ต้องการเบิก ${reqQty})`);
        newQty = currentQty - reqQty;
        stockSheet.getRange(rowToUpdate, 5).setValue(newQty);
        stockSheet.getRange(rowToUpdate, 7).setValue(timeStr);
      }

      let randCode = String(Math.floor(100 + Math.random() * 900));
      let transId = "TX-" + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyMMddHHmmss") + randCode + i;

      logRowsToAppend.push([
        transId, timeStr, type, item.itemID, reqLot, item.exp || "-", reqQty, userId, item.transportTemp || "", transportSpeed || "", deliveryNoteUrl || ""
      ]);
    }
    
    // -- เขียน Log ลงชีตพร้อมกันทีเดียวแบบรวดเร็ว --
    if (logRowsToAppend.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, logRowsToAppend.length, logRowsToAppend[0].length).setValues(logRowsToAppend);
    }
    
    SpreadsheetApp.flush(); 

    // ส่งแจ้งเตือน Telegram
    Object.keys(txSummaryForTelegram).forEach(unitName => {
        let msg = `📢 <b>แจ้งเตือนการทำรายการ (reAgentics)</b>\nประเภท: ${type === 'RECEIVE' ? '📥 รับน้ำยาเข้าคลัง' : '📤 เบิกใช้น้ำยา'}\nผู้ทำรายการ: ${userId}\nเวลา: ${timeStr}\nหน่วยงาน: ${unitName}\n\n<b>รายการน้ำยา:</b>\n${txSummaryForTelegram[unitName].join('\n')}`;
        sendTelegramNotification(unitName, msg);
    });
    
    logSystem("Transaction Success", `Processed ${type} for ${items.length} items (Batched)`, userId);
    return { success: true, message: `บันทึกรายการ ${type === 'RECEIVE' ? 'รับเข้า' : 'เบิกใช้'} จำนวน ${items.length} รายการ สำเร็จ` };
  } catch (e) {
    logSystem("Transaction Failed", e.message, payload.userId || "Unknown");
    return { success: false, message: e.message };
  } finally { lock.releaseLock(); }
}
