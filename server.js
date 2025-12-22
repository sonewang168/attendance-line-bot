/**
 * 🎓 學生簽到系統 - LINE BOT 後端
 * 功能：GPS 定位簽到、遲到判定、缺席追蹤、Google Sheets 整合
 */

const express = require('express');
const path = require('path');
const line = require('@line/bot-sdk');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const cron = require('node-cron');
require('dotenv').config();

const app = express();

// ===== LINE Bot 設定 =====
const lineConfig = {
    channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
    channelSecret: process.env.LINE_CHANNEL_SECRET
};

const lineClient = new line.Client(lineConfig);

// ===== Google Sheets 設定 =====
let doc;
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

async function initGoogleSheets() {
    const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
        scopes: SCOPES,
    });
    
    doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    console.log('📊 Google Sheets 連線成功:', doc.title);
}

// ===== 工具函數 =====

/**
 * 計算兩點間的距離（公尺）- Haversine 公式
 */
function calculateDistance(lat1, lon1, lat2, lon2) {
    const R = 6371000; // 地球半徑（公尺）
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
              Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
              Math.sin(dLon/2) * Math.sin(dLon/2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
    return R * c;
}

/**
 * 格式化日期時間
 */
function formatDateTime(date) {
    return date.toLocaleString('zh-TW', { 
        timeZone: 'Asia/Taipei',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    });
}

/**
 * 取得今天日期字串
 */
function getTodayString() {
    return new Date().toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' });
}

// ===== Google Sheets 操作 =====

/**
 * 取得或建立工作表
 */
async function getOrCreateSheet(title, headers) {
    let sheet = doc.sheetsByTitle[title];
    if (!sheet) {
        sheet = await doc.addSheet({ title, headerValues: headers });
    }
    return sheet;
}

/**
 * 取得學生資料
 */
async function getStudent(lineUserId) {
    const sheet = await getOrCreateSheet('學生名單', [
        '學號', '姓名', '班級', 'LINE_ID', 'LINE名稱', '註冊時間', '狀態'
    ]);
    const rows = await sheet.getRows();
    return rows.find(row => row.get('LINE_ID') === lineUserId);
}

/**
 * 註冊學生
 */
async function registerStudent(lineUserId, lineName, studentId, studentName, className) {
    const sheet = await getOrCreateSheet('學生名單', [
        '學號', '姓名', '班級', 'LINE_ID', 'LINE名稱', '註冊時間', '狀態'
    ]);
    
    // 檢查學號是否已被使用
    const rows = await sheet.getRows();
    const existing = rows.find(row => row.get('學號') === studentId);
    if (existing) {
        if (existing.get('LINE_ID') === lineUserId) {
            return { success: false, message: '您已經註冊過了！' };
        }
        return { success: false, message: '此學號已被其他帳號綁定！' };
    }
    
    await sheet.addRow({
        '學號': studentId,
        '姓名': studentName,
        '班級': className,
        'LINE_ID': lineUserId,
        'LINE名稱': lineName,
        '註冊時間': formatDateTime(new Date()),
        '狀態': '正常'
    });
    
    return { success: true, message: '註冊成功！' };
}

/**
 * 取得課程資料
 */
async function getCourse(courseId) {
    const sheet = await getOrCreateSheet('課程列表', [
        '課程ID', '科目', '班級', '教師', '上課時間', '教室', 
        '教室緯度', '教室經度', '簽到範圍(公尺)', '遲到標準(分鐘)', 
        '通知教師', '通知家長', '狀態', '建立時間'
    ]);
    const rows = await sheet.getRows();
    return rows.find(row => row.get('課程ID') === courseId);
}

/**
 * 取得今日課程活動
 */
async function getTodaySession(courseId) {
    const today = getTodayString();
    const sheet = await getOrCreateSheet('簽到活動', [
        '活動ID', '課程ID', '日期', '開始時間', '結束時間', 'QR碼內容', '狀態'
    ]);
    const rows = await sheet.getRows();
    return rows.find(row => 
        row.get('課程ID') === courseId && 
        row.get('日期') === today &&
        row.get('狀態') === '進行中'
    );
}

/**
 * 記錄簽到
 */
async function recordAttendance(sessionId, studentId, status, lateMinutes = 0, gpsLat = '', gpsLon = '') {
    const sheet = await getOrCreateSheet('簽到紀錄', [
        '活動ID', '學號', '簽到時間', '狀態', '遲到分鐘', 'GPS緯度', 'GPS經度', '備註'
    ]);
    
    // 檢查是否已簽到
    const rows = await sheet.getRows();
    const existing = rows.find(row => 
        row.get('活動ID') === sessionId && 
        row.get('學號') === studentId
    );
    
    if (existing) {
        return { success: false, message: '您已經簽到過了！', status: existing.get('狀態') };
    }
    
    await sheet.addRow({
        '活動ID': sessionId,
        '學號': studentId,
        '簽到時間': formatDateTime(new Date()),
        '狀態': status,
        '遲到分鐘': lateMinutes,
        'GPS緯度': gpsLat,
        'GPS經度': gpsLon,
        '備註': ''
    });
    
    // 更新統計
    await updateStatistics(studentId, status);
    
    return { success: true, message: '簽到成功！', status };
}

/**
 * 更新統計資料
 */
async function updateStatistics(studentId, status) {
    const sheet = await getOrCreateSheet('出席統計', [
        '學號', '姓名', '班級', '出席次數', '遲到次數', '缺席次數', '出席率', '最後更新'
    ]);
    
    const rows = await sheet.getRows();
    let statRow = rows.find(row => row.get('學號') === studentId);
    
    // 取得學生資料
    const studentSheet = doc.sheetsByTitle['學生名單'];
    const studentRows = await studentSheet.getRows();
    const student = studentRows.find(row => row.get('學號') === studentId);
    
    if (!statRow) {
        // 建立新統計
        statRow = await sheet.addRow({
            '學號': studentId,
            '姓名': student ? student.get('姓名') : '',
            '班級': student ? student.get('班級') : '',
            '出席次數': 0,
            '遲到次數': 0,
            '缺席次數': 0,
            '出席率': '0%',
            '最後更新': formatDateTime(new Date())
        });
    }
    
    // 更新計數
    let attended = parseInt(statRow.get('出席次數')) || 0;
    let late = parseInt(statRow.get('遲到次數')) || 0;
    let absent = parseInt(statRow.get('缺席次數')) || 0;
    
    if (status === '已報到') attended++;
    else if (status === '遲到') { attended++; late++; }
    else if (status === '缺席') absent++;
    
    const total = attended + absent;
    const rate = total > 0 ? Math.round((attended / total) * 100) : 0;
    
    statRow.set('出席次數', attended);
    statRow.set('遲到次數', late);
    statRow.set('缺席次數', absent);
    statRow.set('出席率', `${rate}%`);
    statRow.set('最後更新', formatDateTime(new Date()));
    await statRow.save();
}

/**
 * 取得班級列表
 */
async function getClasses() {
    const sheet = await getOrCreateSheet('班級列表', [
        '班級代碼', '班級名稱', '導師', '人數', '建立時間'
    ]);
    const rows = await sheet.getRows();
    return rows.map(row => ({
        code: row.get('班級代碼'),
        name: row.get('班級名稱')
    }));
}

// ===== LINE Bot 訊息處理 =====

// 用戶狀態暫存（實際應用建議用 Redis）
const userStates = new Map();

/**
 * 處理 Webhook 事件
 */
async function handleEvent(event) {
    if (event.type !== 'message' && event.type !== 'postback') {
        return null;
    }
    
    const userId = event.source.userId;
    const userProfile = await lineClient.getProfile(userId);
    const userName = userProfile.displayName;
    
    // 處理 Postback（按鈕回應）
    if (event.type === 'postback') {
        return handlePostback(event, userId, userName);
    }
    
    // 處理位置訊息（GPS 簽到）
    if (event.message.type === 'location') {
        return handleLocation(event, userId);
    }
    
    // 處理文字訊息
    if (event.message.type === 'text') {
        const text = event.message.text.trim();
        
        // 檢查是否為簽到連結
        if (text.startsWith('簽到:')) {
            return handleCheckinRequest(event, userId, text);
        }
        
        // 檢查用戶狀態（是否在註冊流程中）
        const state = userStates.get(userId);
        if (state) {
            return handleRegistrationFlow(event, userId, userName, text, state);
        }
        
        // 一般指令
        return handleCommand(event, userId, userName, text);
    }
    
    return null;
}

/**
 * 處理一般指令
 */
async function handleCommand(event, userId, userName, text) {
    const student = await getStudent(userId);
    
    switch(text) {
        case '註冊':
        case '綁定':
            if (student) {
                return replyText(event, `✅ 您已經註冊過了！\n\n📋 您的資料：\n學號：${student.get('學號')}\n姓名：${student.get('姓名')}\n班級：${student.get('班級')}`);
            }
            // 開始註冊流程
            userStates.set(userId, { step: 'studentId' });
            return replyText(event, '📝 開始註冊\n\n請輸入您的【學號】：');
        
        case '我的資料':
        case '查詢':
            if (!student) {
                return replyText(event, '❌ 您尚未註冊！\n\n請輸入「註冊」開始綁定學號。');
            }
            return replyStudentInfo(event, student);
        
        case '出席紀錄':
        case '統計':
            if (!student) {
                return replyText(event, '❌ 您尚未註冊！\n\n請輸入「註冊」開始綁定學號。');
            }
            return replyAttendanceStats(event, student.get('學號'));
        
        case '說明':
        case '幫助':
        case 'help':
            return replyHelp(event);
        
        default:
            if (!student) {
                return replyText(event, `👋 歡迎 ${userName}！\n\n您尚未註冊，請輸入「註冊」綁定學號後才能使用簽到功能。\n\n輸入「說明」查看更多指令。`);
            }
            return replyText(event, `👋 ${student.get('姓名')} 同學您好！\n\n📌 可用指令：\n• 我的資料\n• 出席紀錄\n• 說明\n\n📍 簽到請掃描教師提供的 QR Code`);
    }
}

/**
 * 處理註冊流程
 */
async function handleRegistrationFlow(event, userId, userName, text, state) {
    switch(state.step) {
        case 'studentId':
            // 驗證學號格式（可自訂）
            if (!/^\d{6,10}$/.test(text)) {
                return replyText(event, '❌ 學號格式不正確！\n\n請輸入 6-10 位數字的學號：');
            }
            userStates.set(userId, { ...state, step: 'studentName', studentId: text });
            return replyText(event, `學號：${text} ✓\n\n請輸入您的【姓名】：`);
        
        case 'studentName':
            if (text.length < 2 || text.length > 10) {
                return replyText(event, '❌ 姓名長度應為 2-10 個字！\n\n請重新輸入您的【姓名】：');
            }
            userStates.set(userId, { ...state, step: 'className', studentName: text });
            
            // 顯示班級選擇
            const classes = await getClasses();
            if (classes.length > 0) {
                return replyClassSelection(event, classes, text);
            }
            return replyText(event, `姓名：${text} ✓\n\n請輸入您的【班級】（例如：801、802）：`);
        
        case 'className':
            const result = await registerStudent(
                userId, 
                userName, 
                state.studentId, 
                state.studentName, 
                text
            );
            userStates.delete(userId);
            
            if (result.success) {
                return replyText(event, `🎉 註冊成功！\n\n📋 您的資料：\n學號：${state.studentId}\n姓名：${state.studentName}\n班級：${text}\n\n現在可以使用簽到功能了！`);
            }
            return replyText(event, `❌ ${result.message}`);
    }
}

/**
 * 處理簽到請求
 */
async function handleCheckinRequest(event, userId, text) {
    const student = await getStudent(userId);
    if (!student) {
        return replyText(event, '❌ 您尚未註冊！\n\n請先輸入「註冊」綁定學號。');
    }
    
    // 解析簽到碼
    const parts = text.replace('簽到:', '').split('|');
    if (parts.length < 2) {
        return replyText(event, '❌ 無效的簽到碼！');
    }
    
    const [courseId, sessionId] = parts;
    
    // 取得課程資訊
    const course = await getCourse(courseId);
    if (!course) {
        return replyText(event, '❌ 找不到此課程！');
    }
    
    // 取得今日活動
    const session = await getTodaySession(courseId);
    if (!session || session.get('活動ID') !== sessionId) {
        return replyText(event, '❌ 此簽到活動已結束或不存在！');
    }
    
    // 儲存待簽到資訊
    userStates.set(userId, { 
        step: 'waitingLocation',
        courseId,
        sessionId,
        courseName: course.get('科目'),
        classroomLat: parseFloat(course.get('教室緯度')),
        classroomLon: parseFloat(course.get('教室經度')),
        checkRadius: parseInt(course.get('簽到範圍(公尺)')) || 50,
        lateMinutes: parseInt(course.get('遲到標準(分鐘)')) || 10,
        startTime: session.get('開始時間')
    });
    
    // 請求位置
    return replyLocationRequest(event, course.get('科目'));
}

/**
 * 處理位置訊息
 */
async function handleLocation(event, userId) {
    const state = userStates.get(userId);
    if (!state || state.step !== 'waitingLocation') {
        return replyText(event, '❌ 請先掃描簽到 QR Code！');
    }
    
    const { latitude, longitude } = event.message;
    const student = await getStudent(userId);
    
    // 計算距離
    const distance = calculateDistance(
        latitude, longitude,
        state.classroomLat, state.classroomLon
    );
    
    // 檢查是否在範圍內
    if (distance > state.checkRadius) {
        userStates.delete(userId);
        return replyText(event, 
            `🚫 簽到失敗！\n\n您不在教室範圍內。\n📍 與教室距離：${Math.round(distance)} 公尺\n📏 允許範圍：${state.checkRadius} 公尺\n\n請到教室後再試一次。`
        );
    }
    
    // 計算是否遲到
    const now = new Date();
    const [startHour, startMin] = state.startTime.split(':').map(Number);
    const startDate = new Date();
    startDate.setHours(startHour, startMin, 0, 0);
    
    const diffMinutes = Math.floor((now - startDate) / 60000);
    let status = '已報到';
    let lateMinutes = 0;
    
    if (diffMinutes > state.lateMinutes) {
        status = '遲到';
        lateMinutes = diffMinutes;
    }
    
    // 記錄簽到
    const result = await recordAttendance(
        state.sessionId,
        student.get('學號'),
        status,
        lateMinutes,
        latitude.toString(),
        longitude.toString()
    );
    
    userStates.delete(userId);
    
    if (!result.success) {
        return replyText(event, `ℹ️ ${result.message}\n\n狀態：${result.status}`);
    }
    
    // 簽到成功訊息
    let message = '';
    if (status === '已報到') {
        message = `✅ 簽到成功！\n\n📚 課程：${state.courseName}\n⏰ 時間：${formatDateTime(now)}\n📍 狀態：準時報到\n\n繼續保持！💪`;
    } else {
        message = `⚠️ 簽到成功（遲到）\n\n📚 課程：${state.courseName}\n⏰ 時間：${formatDateTime(now)}\n📍 狀態：遲到 ${lateMinutes} 分鐘\n\n下次請準時到達！`;
    }
    
    return replyText(event, message);
}

/**
 * 處理 Postback
 */
async function handlePostback(event, userId, userName) {
    const data = event.postback.data;
    const params = new URLSearchParams(data);
    const action = params.get('action');
    
    if (action === 'selectClass') {
        const className = params.get('class');
        const state = userStates.get(userId);
        if (state && state.step === 'className') {
            const result = await registerStudent(
                userId, 
                userName, 
                state.studentId, 
                state.studentName, 
                className
            );
            userStates.delete(userId);
            
            if (result.success) {
                return replyText(event, `🎉 註冊成功！\n\n📋 您的資料：\n學號：${state.studentId}\n姓名：${state.studentName}\n班級：${className}\n\n現在可以使用簽到功能了！`);
            }
            return replyText(event, `❌ ${result.message}`);
        }
    }
    
    return null;
}

// ===== 回覆訊息函數 =====

function replyText(event, text) {
    return lineClient.replyMessage(event.replyToken, {
        type: 'text',
        text: text
    });
}

function replyLocationRequest(event, courseName) {
    return lineClient.replyMessage(event.replyToken, {
        type: 'text',
        text: `📚 準備簽到：${courseName}\n\n請點擊下方「+」按鈕，選擇「位置訊息」分享您的位置來完成簽到。`,
        quickReply: {
            items: [{
                type: 'action',
                action: {
                    type: 'location',
                    label: '📍 分享位置簽到'
                }
            }]
        }
    });
}

function replyClassSelection(event, classes, studentName) {
    const columns = classes.slice(0, 10).map(c => ({
        type: 'action',
        action: {
            type: 'postback',
            label: c.name || c.code,
            data: `action=selectClass&class=${c.code}`
        }
    }));
    
    return lineClient.replyMessage(event.replyToken, {
        type: 'text',
        text: `姓名：${studentName} ✓\n\n請選擇您的班級：`,
        quickReply: { items: columns }
    });
}

async function replyStudentInfo(event, student) {
    const statsSheet = doc.sheetsByTitle['出席統計'];
    let stats = null;
    if (statsSheet) {
        const rows = await statsSheet.getRows();
        stats = rows.find(row => row.get('學號') === student.get('學號'));
    }
    
    let message = `📋 學生資料\n\n`;
    message += `👤 姓名：${student.get('姓名')}\n`;
    message += `🔢 學號：${student.get('學號')}\n`;
    message += `🏫 班級：${student.get('班級')}\n`;
    message += `📅 註冊時間：${student.get('註冊時間')}\n`;
    
    if (stats) {
        message += `\n📊 出席統計\n`;
        message += `✅ 出席：${stats.get('出席次數')} 次\n`;
        message += `⚠️ 遲到：${stats.get('遲到次數')} 次\n`;
        message += `❌ 缺席：${stats.get('缺席次數')} 次\n`;
        message += `📈 出席率：${stats.get('出席率')}`;
    }
    
    return replyText(event, message);
}

async function replyAttendanceStats(event, studentId) {
    const sheet = doc.sheetsByTitle['簽到紀錄'];
    if (!sheet) {
        return replyText(event, '📊 尚無簽到紀錄');
    }
    
    const rows = await sheet.getRows();
    const records = rows.filter(row => row.get('學號') === studentId)
        .slice(-10)
        .reverse();
    
    if (records.length === 0) {
        return replyText(event, '📊 尚無簽到紀錄');
    }
    
    let message = '📊 最近 10 筆簽到紀錄\n\n';
    records.forEach((record, index) => {
        const status = record.get('狀態');
        const icon = status === '已報到' ? '✅' : status === '遲到' ? '⚠️' : '❌';
        message += `${icon} ${record.get('簽到時間')}\n`;
        if (status === '遲到') {
            message += `   遲到 ${record.get('遲到分鐘')} 分鐘\n`;
        }
    });
    
    return replyText(event, message);
}

function replyHelp(event) {
    const message = `📖 使用說明\n\n` +
        `【學生指令】\n` +
        `• 註冊 - 綁定學號\n` +
        `• 我的資料 - 查看個人資訊\n` +
        `• 出席紀錄 - 查看簽到記錄\n` +
        `• 說明 - 顯示此說明\n\n` +
        `【簽到方式】\n` +
        `1. 掃描教師提供的 QR Code\n` +
        `2. 分享您的位置\n` +
        `3. 系統自動完成簽到\n\n` +
        `⚠️ 注意：必須在教室範圍內才能簽到！`;
    
    return replyText(event, message);
}

// ===== 缺席檢查排程 =====

async function checkAbsences() {
    console.log('⏰ 執行缺席檢查...');
    
    try {
        const sessionSheet = doc.sheetsByTitle['簽到活動'];
        const recordSheet = doc.sheetsByTitle['簽到紀錄'];
        const studentSheet = doc.sheetsByTitle['學生名單'];
        
        if (!sessionSheet || !studentSheet) return;
        
        const sessions = await sessionSheet.getRows();
        const now = new Date();
        
        for (const session of sessions) {
            if (session.get('狀態') !== '進行中') continue;
            
            // 檢查是否已結束
            const [endHour, endMin] = session.get('結束時間').split(':').map(Number);
            const endTime = new Date();
            endTime.setHours(endHour, endMin, 0, 0);
            
            if (now > endTime) {
                // 標記缺席的學生
                const courseSheet = doc.sheetsByTitle['課程列表'];
                const courses = await courseSheet.getRows();
                const course = courses.find(c => c.get('課程ID') === session.get('課程ID'));
                
                if (course) {
                    const className = course.get('班級');
                    const students = await studentSheet.getRows();
                    const classStudents = students.filter(s => s.get('班級') === className);
                    
                    const records = recordSheet ? await recordSheet.getRows() : [];
                    
                    for (const student of classStudents) {
                        const hasRecord = records.some(r => 
                            r.get('活動ID') === session.get('活動ID') &&
                            r.get('學號') === student.get('學號')
                        );
                        
                        if (!hasRecord) {
                            // 記錄缺席
                            await recordAttendance(
                                session.get('活動ID'),
                                student.get('學號'),
                                '缺席'
                            );
                            
                            // 發送缺席通知
                            try {
                                await lineClient.pushMessage(student.get('LINE_ID'), {
                                    type: 'text',
                                    text: `❌ 缺席通知\n\n您已被標記為缺席：\n📚 課程：${course.get('科目')}\n📅 日期：${session.get('日期')}\n\n如有疑問請聯繫教師。`
                                });
                            } catch (e) {
                                console.error('發送通知失敗:', e);
                            }
                        }
                    }
                }
                
                // 更新活動狀態
                session.set('狀態', '已結束');
                await session.save();
            }
        }
        
        console.log('✅ 缺席檢查完成');
    } catch (error) {
        console.error('缺席檢查錯誤:', error);
    }
}

// 每 5 分鐘檢查一次
cron.schedule('*/5 * * * *', checkAbsences);

// ===== Express 路由 =====

// 靜態檔案（教師管理介面）
app.use(express.static(path.join(__dirname, 'public')));

// 首頁路由
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.use('/webhook', line.middleware(lineConfig));

app.post('/webhook', (req, res) => {
    Promise.all(req.body.events.map(handleEvent))
        .then((result) => res.json(result))
        .catch((err) => {
            console.error('Webhook Error:', err);
            res.status(500).end();
        });
});

// ===== API 端點 =====
app.use(express.json());

// CORS 設定
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE');
    res.header('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') return res.sendStatus(200);
    next();
});

// === 班級 API ===
app.get('/api/classes', async (req, res) => {
    try {
        const sheet = await getOrCreateSheet('班級列表', ['班級代碼', '班級名稱', '導師', '人數', '建立時間']);
        const rows = await sheet.getRows();
        res.json(rows.map(r => ({
            code: r.get('班級代碼'),
            name: r.get('班級名稱'),
            teacher: r.get('導師'),
            count: r.get('人數') || 0
        })));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/classes', async (req, res) => {
    try {
        const { code, name, teacher } = req.body;
        const sheet = await getOrCreateSheet('班級列表', ['班級代碼', '班級名稱', '導師', '人數', '建立時間']);
        await sheet.addRow({
            '班級代碼': code,
            '班級名稱': name,
            '導師': teacher || '',
            '人數': 0,
            '建立時間': new Date().toLocaleString('zh-TW')
        });
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.put('/api/classes/:code', async (req, res) => {
    try {
        const { code } = req.params;
        const { name, teacher } = req.body;
        const sheet = doc.sheetsByTitle['班級列表'];
        if (!sheet) return res.json({ success: false });
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('班級代碼') === code);
        if (row) {
            if (name) row.set('班級名稱', name);
            if (teacher !== undefined) row.set('導師', teacher);
            await row.save();
        }
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.delete('/api/classes/:code', async (req, res) => {
    try {
        const { code } = req.params;
        const sheet = doc.sheetsByTitle['班級列表'];
        if (!sheet) return res.json({ success: true });
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('班級代碼') === code);
        if (row) await row.delete();
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 取得單一班級的學生
app.get('/api/classes/:code/students', async (req, res) => {
    try {
        const { code } = req.params;
        const sheet = doc.sheetsByTitle['學生名單'];
        if (!sheet) return res.json([]);
        const rows = await sheet.getRows();
        const students = rows.filter(r => r.get('班級') === code);
        res.json(students.map(s => ({
            studentId: s.get('學號'),
            name: s.get('姓名'),
            lineName: s.get('LINE名稱'),
            registeredAt: s.get('註冊時間')
        })));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// === 課程 API ===
app.get('/api/courses', async (req, res) => {
    try {
        const sheet = await getOrCreateSheet('課程列表', [
            '課程ID', '科目', '班級', '教師', '上課時間', '教室',
            '教室緯度', '教室經度', '簽到範圍', '遲到標準', '狀態', '建立時間'
        ]);
        const rows = await sheet.getRows();
        res.json(rows.map(r => ({
            id: r.get('課程ID'),
            subject: r.get('科目'),
            classCode: r.get('班級'),
            teacher: r.get('教師'),
            time: r.get('上課時間'),
            room: r.get('教室'),
            lat: parseFloat(r.get('教室緯度')) || 0,
            lon: parseFloat(r.get('教室經度')) || 0,
            radius: parseInt(r.get('簽到範圍')) || 50,
            lateMinutes: parseInt(r.get('遲到標準')) || 10,
            status: r.get('狀態') || '啟用'
        })));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/courses', async (req, res) => {
    try {
        const { subject, classCode, teacher, time, room, lat, lon, radius, lateMinutes } = req.body;
        const sheet = await getOrCreateSheet('課程列表', [
            '課程ID', '科目', '班級', '教師', '上課時間', '教室',
            '教室緯度', '教室經度', '簽到範圍', '遲到標準', '狀態', '建立時間'
        ]);
        const courseId = 'C' + Date.now();
        await sheet.addRow({
            '課程ID': courseId,
            '科目': subject,
            '班級': classCode,
            '教師': teacher || '',
            '上課時間': time || '',
            '教室': room || '',
            '教室緯度': lat,
            '教室經度': lon,
            '簽到範圍': radius || 50,
            '遲到標準': lateMinutes || 10,
            '狀態': '啟用',
            '建立時間': new Date().toLocaleString('zh-TW')
        });
        res.json({ success: true, courseId });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.delete('/api/courses/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const sheet = doc.sheetsByTitle['課程列表'];
        if (!sheet) return res.json({ success: true });
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('課程ID') === id);
        if (row) await row.delete();
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// === 學生 API ===
app.get('/api/students', async (req, res) => {
    try {
        const sheet = await getOrCreateSheet('學生名單', [
            '學號', '姓名', '班級', 'LINE_ID', 'LINE名稱', '註冊時間', '狀態'
        ]);
        const rows = await sheet.getRows();
        res.json(rows.map(r => ({
            studentId: r.get('學號'),
            name: r.get('姓名'),
            classCode: r.get('班級'),
            lineId: r.get('LINE_ID'),
            lineName: r.get('LINE名稱'),
            registeredAt: r.get('註冊時間'),
            status: r.get('狀態')
        })));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// === 簽到活動 API ===
app.get('/api/sessions', async (req, res) => {
    try {
        const sheet = await getOrCreateSheet('簽到活動', [
            '活動ID', '課程ID', '日期', '開始時間', '結束時間', 'QR碼內容', '狀態'
        ]);
        const rows = await sheet.getRows();
        
        // 取得課程資料以顯示名稱
        const courseSheet = doc.sheetsByTitle['課程列表'];
        const courses = courseSheet ? await courseSheet.getRows() : [];
        const courseMap = {};
        courses.forEach(c => {
            courseMap[c.get('課程ID')] = { subject: c.get('科目'), classCode: c.get('班級') };
        });
        
        res.json(rows.map(r => {
            const courseId = r.get('課程ID');
            const course = courseMap[courseId] || {};
            return {
                id: r.get('活動ID'),
                courseId: courseId,
                courseName: course.subject || '未知課程',
                classCode: course.classCode || '',
                date: r.get('日期'),
                startTime: r.get('開始時間'),
                endTime: r.get('結束時間'),
                qrContent: r.get('QR碼內容'),
                status: r.get('狀態')
            };
        }));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/sessions', async (req, res) => {
    try {
        const { courseId, date, startTime, endTime } = req.body;
        const sheet = await getOrCreateSheet('簽到活動', [
            '活動ID', '課程ID', '日期', '開始時間', '結束時間', 'QR碼內容', '狀態'
        ]);
        const sessionId = `S${Date.now()}`;
        const qrContent = `簽到:${courseId}|${sessionId}`;
        await sheet.addRow({
            '活動ID': sessionId,
            '課程ID': courseId,
            '日期': date,
            '開始時間': startTime,
            '結束時間': endTime,
            'QR碼內容': qrContent,
            '狀態': '進行中'
        });
        res.json({ success: true, sessionId, qrContent });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.put('/api/sessions/:id/end', async (req, res) => {
    try {
        const { id } = req.params;
        const sheet = doc.sheetsByTitle['簽到活動'];
        if (!sheet) return res.json({ success: false });
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('活動ID') === id);
        if (row) {
            row.set('狀態', '已結束');
            await row.save();
        }
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// === 簽到紀錄 API ===
app.get('/api/records', async (req, res) => {
    try {
        const sheet = await getOrCreateSheet('簽到紀錄', [
            '活動ID', '學號', '簽到時間', '狀態', '遲到分鐘', 'GPS緯度', 'GPS經度', '備註'
        ]);
        const rows = await sheet.getRows();
        
        // 取得學生資料
        const studentSheet = doc.sheetsByTitle['學生名單'];
        const students = studentSheet ? await studentSheet.getRows() : [];
        const studentMap = {};
        students.forEach(s => {
            studentMap[s.get('學號')] = s.get('姓名');
        });
        
        // 取得活動資料
        const sessionSheet = doc.sheetsByTitle['簽到活動'];
        const sessions = sessionSheet ? await sessionSheet.getRows() : [];
        const sessionMap = {};
        sessions.forEach(s => {
            sessionMap[s.get('活動ID')] = { courseId: s.get('課程ID'), date: s.get('日期') };
        });
        
        // 取得課程資料
        const courseSheet = doc.sheetsByTitle['課程列表'];
        const courses = courseSheet ? await courseSheet.getRows() : [];
        const courseMap = {};
        courses.forEach(c => {
            courseMap[c.get('課程ID')] = c.get('科目');
        });
        
        res.json(rows.map(r => {
            const sessionId = r.get('活動ID');
            const session = sessionMap[sessionId] || {};
            const courseName = courseMap[session.courseId] || '未知';
            const studentId = r.get('學號');
            return {
                sessionId: sessionId,
                studentId: studentId,
                studentName: studentMap[studentId] || '未知',
                courseName: courseName,
                date: session.date || '',
                time: r.get('簽到時間'),
                status: r.get('狀態'),
                lateMinutes: r.get('遲到分鐘')
            };
        }));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// === 儀表板統計 API ===
app.get('/api/dashboard', async (req, res) => {
    try {
        const today = new Date().toISOString().split('T')[0];
        
        // 學生數
        const studentSheet = doc.sheetsByTitle['學生名單'];
        const students = studentSheet ? await studentSheet.getRows() : [];
        
        // 今日紀錄
        const recordSheet = doc.sheetsByTitle['簽到紀錄'];
        const records = recordSheet ? await recordSheet.getRows() : [];
        
        const sessionSheet = doc.sheetsByTitle['簽到活動'];
        const sessions = sessionSheet ? await sessionSheet.getRows() : [];
        const todaySessionIds = sessions.filter(s => s.get('日期') === today).map(s => s.get('活動ID'));
        
        const todayRecords = records.filter(r => todaySessionIds.includes(r.get('活動ID')));
        
        const attended = todayRecords.filter(r => r.get('狀態') === '已報到').length;
        const late = todayRecords.filter(r => r.get('狀態') === '遲到').length;
        const absent = todayRecords.filter(r => r.get('狀態') === '缺席').length;
        
        // 最近紀錄
        const recentRecords = records.slice(-10).reverse().map(r => ({
            studentId: r.get('學號'),
            time: r.get('簽到時間'),
            status: r.get('狀態')
        }));
        
        res.json({
            totalStudents: students.length,
            todayAttended: attended,
            todayLate: late,
            todayAbsent: absent,
            recentRecords: recentRecords
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// === 健康檢查 ===
app.get('/api/health', (req, res) => {
    res.json({ status: 'ok', time: new Date().toISOString() });
});

// === 統計 API ===
app.get('/api/stats/attendance', async (req, res) => {
    try {
        const recordSheet = doc.sheetsByTitle['簽到紀錄'];
        const studentSheet = doc.sheetsByTitle['學生名單'];
        if (!recordSheet || !studentSheet) {
            return res.json({ overall: 0, students: [] });
        }
        
        const records = await recordSheet.getRows();
        const students = await studentSheet.getRows();
        
        // 計算整體出席率
        const total = records.length;
        const attended = records.filter(r => r.get('狀態') === '已報到').length;
        const late = records.filter(r => r.get('狀態') === '遲到').length;
        const absent = records.filter(r => r.get('狀態') === '缺席').length;
        const overall = total > 0 ? Math.round((attended + late) / total * 100) : 0;
        
        // 計算每位學生的出席率
        const studentStats = [];
        for (const student of students) {
            const studentId = student.get('學號');
            const studentRecords = records.filter(r => r.get('學號') === studentId);
            const sTotal = studentRecords.length;
            const sAttended = studentRecords.filter(r => r.get('狀態') === '已報到').length;
            const sLate = studentRecords.filter(r => r.get('狀態') === '遲到').length;
            const sAbsent = studentRecords.filter(r => r.get('狀態') === '缺席').length;
            const rate = sTotal > 0 ? Math.round((sAttended + sLate) / sTotal * 100) : 100;
            
            studentStats.push({
                studentId,
                name: student.get('姓名'),
                classCode: student.get('班級'),
                total: sTotal,
                attended: sAttended,
                late: sLate,
                absent: sAbsent,
                rate
            });
        }
        
        // 排序：出席率低的在前
        studentStats.sort((a, b) => a.rate - b.rate);
        
        res.json({
            overall,
            totalRecords: total,
            attended,
            late,
            absent,
            students: studentStats,
            lowAttendance: studentStats.filter(s => s.rate < 80),
            warnings: studentStats.filter(s => s.rate < 60)
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 取得學生連續缺席狀況
app.get('/api/stats/consecutive-absent', async (req, res) => {
    try {
        const recordSheet = doc.sheetsByTitle['簽到紀錄'];
        const studentSheet = doc.sheetsByTitle['學生名單'];
        if (!recordSheet || !studentSheet) {
            return res.json({ alerts: [] });
        }
        
        const records = await recordSheet.getRows();
        const students = await studentSheet.getRows();
        const alerts = [];
        
        for (const student of students) {
            const studentId = student.get('學號');
            const studentRecords = records
                .filter(r => r.get('學號') === studentId)
                .sort((a, b) => new Date(b.get('簽到時間')) - new Date(a.get('簽到時間')));
            
            // 計算連續缺席次數
            let consecutive = 0;
            for (const r of studentRecords) {
                if (r.get('狀態') === '缺席') {
                    consecutive++;
                } else {
                    break;
                }
            }
            
            if (consecutive >= 2) {
                alerts.push({
                    studentId,
                    name: student.get('姓名'),
                    classCode: student.get('班級'),
                    lineId: student.get('LINE_ID'),
                    consecutiveAbsent: consecutive,
                    level: consecutive >= 5 ? 'critical' : consecutive >= 3 ? 'warning' : 'notice'
                });
            }
        }
        
        alerts.sort((a, b) => b.consecutiveAbsent - a.consecutiveAbsent);
        res.json({ alerts });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// === 通知 API ===
// 發送上課提醒
app.post('/api/notify/remind', async (req, res) => {
    try {
        const { courseId, message } = req.body;
        const studentSheet = doc.sheetsByTitle['學生名單'];
        const courseSheet = doc.sheetsByTitle['課程列表'];
        
        if (!studentSheet || !courseSheet) {
            return res.json({ success: false, message: '找不到資料表' });
        }
        
        const courses = await courseSheet.getRows();
        const course = courses.find(c => c.get('課程ID') === courseId);
        if (!course) {
            return res.json({ success: false, message: '找不到課程' });
        }
        
        const classCode = course.get('班級');
        const students = await studentSheet.getRows();
        const classStudents = students.filter(s => s.get('班級') === classCode && s.get('LINE_ID'));
        
        // 發送 LINE 通知
        const notifications = [];
        for (const student of classStudents) {
            const lineId = student.get('LINE_ID');
            if (lineId) {
                try {
                    await lineClient.pushMessage(lineId, {
                        type: 'text',
                        text: message || `📢 上課提醒\n\n${course.get('科目')} 即將開始！\n⏰ ${course.get('上課時間')}\n📍 ${course.get('教室')}\n\n請準時出席！`
                    });
                    notifications.push({ studentId: student.get('學號'), status: 'sent' });
                } catch (e) {
                    notifications.push({ studentId: student.get('學號'), status: 'failed', error: e.message });
                }
            }
        }
        
        res.json({ success: true, sent: notifications.filter(n => n.status === 'sent').length, total: classStudents.length, details: notifications });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 發送缺席通知
app.post('/api/notify/absent', async (req, res) => {
    try {
        const { studentId, sessionId, courseName } = req.body;
        const studentSheet = doc.sheetsByTitle['學生名單'];
        
        if (!studentSheet) {
            return res.json({ success: false, message: '找不到學生資料' });
        }
        
        const students = await studentSheet.getRows();
        const student = students.find(s => s.get('學號') === studentId);
        
        if (!student || !student.get('LINE_ID')) {
            return res.json({ success: false, message: '學生未綁定 LINE' });
        }
        
        await lineClient.pushMessage(student.get('LINE_ID'), {
            type: 'text',
            text: `⚠️ 缺席通知\n\n${student.get('姓名')} 同學，您在「${courseName}」課程中被記錄為缺席。\n\n如有疑問請聯繫老師。`
        });
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 發送連續缺席警告
app.post('/api/notify/warning', async (req, res) => {
    try {
        const { studentId, consecutiveCount } = req.body;
        const studentSheet = doc.sheetsByTitle['學生名單'];
        
        if (!studentSheet) {
            return res.json({ success: false });
        }
        
        const students = await studentSheet.getRows();
        const student = students.find(s => s.get('學號') === studentId);
        
        if (!student || !student.get('LINE_ID')) {
            return res.json({ success: false, message: '學生未綁定 LINE' });
        }
        
        const level = consecutiveCount >= 5 ? '🚨 嚴重警告' : consecutiveCount >= 3 ? '⚠️ 警告' : '📢 提醒';
        
        await lineClient.pushMessage(student.get('LINE_ID'), {
            type: 'text',
            text: `${level}\n\n${student.get('姓名')} 同學，您已連續 ${consecutiveCount} 次缺席！\n\n請盡快與老師聯繫說明情況。持續缺席可能影響您的學業成績。`
        });
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 批次發送通知
app.post('/api/notify/batch', async (req, res) => {
    try {
        const { type, targets, message } = req.body;
        const studentSheet = doc.sheetsByTitle['學生名單'];
        
        if (!studentSheet) {
            return res.json({ success: false });
        }
        
        const students = await studentSheet.getRows();
        let sent = 0, failed = 0;
        
        for (const studentId of targets) {
            const student = students.find(s => s.get('學號') === studentId);
            if (student && student.get('LINE_ID')) {
                try {
                    await lineClient.pushMessage(student.get('LINE_ID'), {
                        type: 'text',
                        text: message
                    });
                    sent++;
                } catch {
                    failed++;
                }
            } else {
                failed++;
            }
        }
        
        res.json({ success: true, sent, failed });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 取得通知設定
app.get('/api/settings/notifications', async (req, res) => {
    try {
        const sheet = await getOrCreateSheet('系統設定', ['設定項目', '設定值']);
        const rows = await sheet.getRows();
        const settings = {};
        rows.forEach(r => {
            settings[r.get('設定項目')] = r.get('設定值');
        });
        res.json({
            remindBeforeClass: settings['上課提醒'] === 'true',
            remindMinutes: parseInt(settings['提醒分鐘']) || 10,
            notifyAbsent: settings['缺席通知'] === 'true',
            notifyParent: settings['通知家長'] === 'true',
            warningThreshold: parseInt(settings['警告門檻']) || 3,
            weeklyReport: settings['週報'] === 'true'
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 儲存通知設定
app.post('/api/settings/notifications', async (req, res) => {
    try {
        const { remindBeforeClass, remindMinutes, notifyAbsent, notifyParent, warningThreshold, weeklyReport } = req.body;
        const sheet = await getOrCreateSheet('系統設定', ['設定項目', '設定值']);
        
        // 清空舊設定
        const rows = await sheet.getRows();
        for (const row of rows) {
            await row.delete();
        }
        
        // 寫入新設定
        await sheet.addRows([
            { '設定項目': '上課提醒', '設定值': remindBeforeClass ? 'true' : 'false' },
            { '設定項目': '提醒分鐘', '設定值': remindMinutes || 10 },
            { '設定項目': '缺席通知', '設定值': notifyAbsent ? 'true' : 'false' },
            { '設定項目': '通知家長', '設定值': notifyParent ? 'true' : 'false' },
            { '設定項目': '警告門檻', '設定值': warningThreshold || 3 },
            { '設定項目': '週報', '設定值': weeklyReport ? 'true' : 'false' }
        ]);
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 上課結束自動處理
app.post('/api/sessions/:id/complete', async (req, res) => {
    try {
        const { id } = req.params;
        const sessionSheet = doc.sheetsByTitle['簽到活動'];
        const recordSheet = doc.sheetsByTitle['簽到紀錄'];
        const studentSheet = doc.sheetsByTitle['學生名單'];
        const courseSheet = doc.sheetsByTitle['課程列表'];
        
        if (!sessionSheet) return res.json({ success: false });
        
        // 更新活動狀態
        const sessions = await sessionSheet.getRows();
        const session = sessions.find(s => s.get('活動ID') === id);
        if (!session) return res.json({ success: false, message: '找不到活動' });
        
        session.set('狀態', '已結束');
        await session.save();
        
        // 找出未簽到的學生，標記為缺席
        const courseId = session.get('課程ID');
        const courses = await courseSheet.getRows();
        const course = courses.find(c => c.get('課程ID') === courseId);
        if (!course) return res.json({ success: true, marked: 0 });
        
        const classCode = course.get('班級');
        const students = await studentSheet.getRows();
        const classStudents = students.filter(s => s.get('班級') === classCode);
        
        const records = await recordSheet.getRows();
        const sessionRecords = records.filter(r => r.get('活動ID') === id);
        const checkedInIds = sessionRecords.map(r => r.get('學號'));
        
        let marked = 0;
        const absentStudents = [];
        
        for (const student of classStudents) {
            const studentId = student.get('學號');
            if (!checkedInIds.includes(studentId)) {
                // 標記缺席
                await recordSheet.addRow({
                    '活動ID': id,
                    '學號': studentId,
                    '簽到時間': new Date().toLocaleString('zh-TW'),
                    '狀態': '缺席',
                    '遲到分鐘': 0,
                    'GPS緯度': '',
                    'GPS經度': '',
                    '備註': '系統自動標記'
                });
                marked++;
                absentStudents.push({
                    studentId,
                    name: student.get('姓名'),
                    lineId: student.get('LINE_ID')
                });
            }
        }
        
        res.json({ 
            success: true, 
            marked, 
            absentStudents,
            courseName: course.get('科目')
        });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 請假系統 API ===
// 取得請假列表
app.get('/api/leaves', async (req, res) => {
    try {
        const sheet = await getOrCreateSheet('請假紀錄', ['請假ID', '學號', '姓名', '班級', '日期', '節次', '請假類型', '原因', '狀態', '申請時間', '審核時間', '審核備註']);
        const rows = await sheet.getRows();
        res.json(rows.map(r => ({
            id: r.get('請假ID'),
            studentId: r.get('學號'),
            name: r.get('姓名'),
            classCode: r.get('班級'),
            date: r.get('日期'),
            periods: r.get('節次'),
            type: r.get('請假類型'),
            reason: r.get('原因'),
            status: r.get('狀態'),
            appliedAt: r.get('申請時間'),
            reviewedAt: r.get('審核時間'),
            reviewNote: r.get('審核備註')
        })));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 學生申請請假
app.post('/api/leaves', async (req, res) => {
    try {
        const { studentId, date, periods, type, reason } = req.body;
        const sheet = await getOrCreateSheet('請假紀錄', ['請假ID', '學號', '姓名', '班級', '日期', '節次', '請假類型', '原因', '狀態', '申請時間', '審核時間', '審核備註']);
        const studentSheet = doc.sheetsByTitle['學生名單'];
        
        if (!studentSheet) return res.json({ success: false, message: '找不到學生資料' });
        
        const students = await studentSheet.getRows();
        const student = students.find(s => s.get('學號') === studentId);
        if (!student) return res.json({ success: false, message: '學生不存在' });
        
        const leaveId = 'L' + Date.now();
        await sheet.addRow({
            '請假ID': leaveId,
            '學號': studentId,
            '姓名': student.get('姓名'),
            '班級': student.get('班級'),
            '日期': date,
            '節次': periods,
            '請假類型': type || '事假',
            '原因': reason || '',
            '狀態': '待審核',
            '申請時間': new Date().toLocaleString('zh-TW'),
            '審核時間': '',
            '審核備註': ''
        });
        
        res.json({ success: true, leaveId });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 審核請假
app.put('/api/leaves/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const { status, note } = req.body;
        const sheet = doc.sheetsByTitle['請假紀錄'];
        if (!sheet) return res.json({ success: false });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('請假ID') === id);
        if (!row) return res.json({ success: false, message: '找不到請假單' });
        
        row.set('狀態', status);
        row.set('審核時間', new Date().toLocaleString('zh-TW'));
        row.set('審核備註', note || '');
        await row.save();
        
        // 發送通知給學生
        const studentSheet = doc.sheetsByTitle['學生名單'];
        if (studentSheet) {
            const students = await studentSheet.getRows();
            const student = students.find(s => s.get('學號') === row.get('學號'));
            if (student && student.get('LINE_ID')) {
                const statusText = status === '已核准' ? '✅ 已核准' : '❌ 已駁回';
                try {
                    await lineClient.pushMessage(student.get('LINE_ID'), {
                        type: 'text',
                        text: `📋 請假審核結果\n\n${statusText}\n日期：${row.get('日期')}\n節次：${row.get('節次')}\n${note ? '備註：' + note : ''}`
                    });
                } catch (e) { console.log('LINE 通知失敗:', e.message); }
            }
        }
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 刪除請假
app.delete('/api/leaves/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const sheet = doc.sheetsByTitle['請假紀錄'];
        if (!sheet) return res.json({ success: true });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('請假ID') === id);
        if (row) await row.delete();
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 家長管理 API ===
// 綁定家長 LINE
app.post('/api/students/:id/parent', async (req, res) => {
    try {
        const { id } = req.params;
        const { parentLineId, parentName } = req.body;
        const sheet = doc.sheetsByTitle['學生名單'];
        if (!sheet) return res.json({ success: false });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('學號') === id);
        if (!row) return res.json({ success: false, message: '學生不存在' });
        
        row.set('家長LINE_ID', parentLineId);
        row.set('家長姓名', parentName || '');
        await row.save();
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 發送通知給家長
app.post('/api/notify/parent', async (req, res) => {
    try {
        const { studentId, message, type } = req.body;
        const studentSheet = doc.sheetsByTitle['學生名單'];
        if (!studentSheet) return res.json({ success: false });
        
        const students = await studentSheet.getRows();
        const student = students.find(s => s.get('學號') === studentId);
        
        if (!student || !student.get('家長LINE_ID')) {
            return res.json({ success: false, message: '家長未綁定 LINE' });
        }
        
        let text = message;
        if (!text) {
            if (type === 'absent') {
                text = `📢 家長您好\n\n您的孩子 ${student.get('姓名')} 今日有缺席紀錄，請關心了解。\n\n如有疑問請與學校聯繫。`;
            } else if (type === 'warning') {
                text = `⚠️ 重要通知\n\n您的孩子 ${student.get('姓名')} 近期出席狀況異常，已連續多次缺席。\n\n請儘速與學校聯繫了解情況。`;
            }
        }
        
        await lineClient.pushMessage(student.get('家長LINE_ID'), { type: 'text', text });
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 批次通知家長
app.post('/api/notify/parents-batch', async (req, res) => {
    try {
        const { studentIds, message, type } = req.body;
        const studentSheet = doc.sheetsByTitle['學生名單'];
        if (!studentSheet) return res.json({ success: false });
        
        const students = await studentSheet.getRows();
        let sent = 0, failed = 0;
        
        for (const studentId of studentIds) {
            const student = students.find(s => s.get('學號') === studentId);
            if (student && student.get('家長LINE_ID')) {
                try {
                    let text = message || `📢 家長您好\n\n您的孩子 ${student.get('姓名')} 的出席狀況需要您關注。\n\n詳情請與學校聯繫。`;
                    await lineClient.pushMessage(student.get('家長LINE_ID'), { type: 'text', text });
                    sent++;
                } catch { failed++; }
            } else { failed++; }
        }
        
        res.json({ success: true, sent, failed });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 週報 API ===
// 產生週報
app.get('/api/reports/weekly', async (req, res) => {
    try {
        const { weekStart, weekEnd } = req.query;
        const recordSheet = doc.sheetsByTitle['簽到紀錄'];
        const studentSheet = doc.sheetsByTitle['學生名單'];
        const classSheet = doc.sheetsByTitle['班級列表'];
        
        if (!recordSheet || !studentSheet) {
            return res.json({ success: false, message: '資料表不存在' });
        }
        
        const records = await recordSheet.getRows();
        const students = await studentSheet.getRows();
        const classes = classSheet ? await classSheet.getRows() : [];
        
        // 過濾本週紀錄
        const weekRecords = records.filter(r => {
            const date = r.get('簽到時間')?.split(' ')[0];
            return date >= weekStart && date <= weekEnd;
        });
        
        const total = weekRecords.length;
        const attended = weekRecords.filter(r => r.get('狀態') === '已報到').length;
        const late = weekRecords.filter(r => r.get('狀態') === '遲到').length;
        const absent = weekRecords.filter(r => r.get('狀態') === '缺席').length;
        const rate = total > 0 ? Math.round((attended + late) / total * 100) : 0;
        
        // 各班統計
        const classSummary = [];
        for (const cls of classes) {
            const code = cls.get('班級代碼');
            const classStudents = students.filter(s => s.get('班級') === code).map(s => s.get('學號'));
            const classRecords = weekRecords.filter(r => classStudents.includes(r.get('學號')));
            const cTotal = classRecords.length;
            const cAttended = classRecords.filter(r => r.get('狀態') === '已報到').length;
            const cLate = classRecords.filter(r => r.get('狀態') === '遲到').length;
            const cAbsent = classRecords.filter(r => r.get('狀態') === '缺席').length;
            
            classSummary.push({
                code, name: cls.get('班級名稱'),
                total: cTotal, attended: cAttended, late: cLate, absent: cAbsent,
                rate: cTotal > 0 ? Math.round((cAttended + cLate) / cTotal * 100) : 100
            });
        }
        
        // 問題學生
        const problemStudents = [];
        for (const student of students) {
            const studentId = student.get('學號');
            const studentRecords = weekRecords.filter(r => r.get('學號') === studentId);
            const sAbsent = studentRecords.filter(r => r.get('狀態') === '缺席').length;
            const sLate = studentRecords.filter(r => r.get('狀態') === '遲到').length;
            
            if (sAbsent >= 2 || sLate >= 3) {
                problemStudents.push({ studentId, name: student.get('姓名'), classCode: student.get('班級'), absent: sAbsent, late: sLate });
            }
        }
        
        res.json({ success: true, weekStart, weekEnd, summary: { total, attended, late, absent, rate }, classSummary, problemStudents });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 發送週報通知
app.post('/api/reports/weekly/send', async (req, res) => {
    try {
        const { report, teacherLineId } = req.body;
        
        let text = `📊 週報 (${report.weekStart} ~ ${report.weekEnd})\n\n`;
        text += `📈 整體統計\n`;
        text += `• 出席率：${report.summary.rate}%\n`;
        text += `• 出席：${report.summary.attended} 次\n`;
        text += `• 遲到：${report.summary.late} 次\n`;
        text += `• 缺席：${report.summary.absent} 次\n\n`;
        
        if (report.problemStudents?.length > 0) {
            text += `⚠️ 需關注學生\n`;
            for (const s of report.problemStudents.slice(0, 5)) {
                text += `• ${s.name} (${s.classCode}): 缺席${s.absent}次, 遲到${s.late}次\n`;
            }
        } else {
            text += `✅ 本週無異常狀況\n`;
        }
        
        if (teacherLineId) {
            await lineClient.pushMessage(teacherLineId, { type: 'text', text });
        }
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 匯出報表 API ===
app.get('/api/export/attendance', async (req, res) => {
    try {
        const { format, startDate, endDate, classCode } = req.query;
        const recordSheet = doc.sheetsByTitle['簽到紀錄'];
        const studentSheet = doc.sheetsByTitle['學生名單'];
        
        if (!recordSheet || !studentSheet) return res.json({ success: false });
        
        const records = await recordSheet.getRows();
        const students = await studentSheet.getRows();
        
        let filtered = records;
        if (startDate) filtered = filtered.filter(r => r.get('簽到時間')?.split(' ')[0] >= startDate);
        if (endDate) filtered = filtered.filter(r => r.get('簽到時間')?.split(' ')[0] <= endDate);
        if (classCode) {
            const classStudentIds = students.filter(s => s.get('班級') === classCode).map(s => s.get('學號'));
            filtered = filtered.filter(r => classStudentIds.includes(r.get('學號')));
        }
        
        const data = filtered.map(r => {
            const student = students.find(s => s.get('學號') === r.get('學號'));
            return {
                日期: r.get('簽到時間')?.split(' ')[0] || '',
                時間: r.get('簽到時間')?.split(' ')[1] || '',
                學號: r.get('學號'),
                姓名: student?.get('姓名') || '',
                班級: student?.get('班級') || '',
                狀態: r.get('狀態'),
                遲到分鐘: r.get('遲到分鐘') || 0,
                備註: r.get('備註') || ''
            };
        });
        
        res.json({ success: true, data, format });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 學生管理 API ===
// 新增學生（手動）
app.post('/api/students', async (req, res) => {
    try {
        const { studentId, name, classCode, phone, parentPhone, parentLineId } = req.body;
        const sheet = await getOrCreateSheet('學生名單', ['學號', '姓名', '班級', 'LINE_ID', '電話', '家長電話', '家長LINE_ID', '註冊時間']);
        
        // 檢查學號是否已存在
        const rows = await sheet.getRows();
        const exists = rows.find(r => r.get('學號') === studentId);
        if (exists) {
            return res.json({ success: false, message: '學號已存在' });
        }
        
        await sheet.addRow({
            '學號': studentId,
            '姓名': name,
            '班級': classCode,
            'LINE_ID': '',
            '電話': phone || '',
            '家長電話': parentPhone || '',
            '家長LINE_ID': parentLineId || '',
            '註冊時間': new Date().toLocaleString('zh-TW')
        });
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 更新學生
app.put('/api/students/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const { name, classCode, phone, parentPhone } = req.body;
        const sheet = doc.sheetsByTitle['學生名單'];
        if (!sheet) return res.json({ success: false });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('學號') === id);
        if (!row) return res.json({ success: false, message: '學生不存在' });
        
        if (name) row.set('姓名', name);
        if (classCode) row.set('班級', classCode);
        if (phone !== undefined) row.set('電話', phone);
        if (parentPhone !== undefined) row.set('家長電話', parentPhone);
        await row.save();
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 刪除學生
app.delete('/api/students/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const sheet = doc.sheetsByTitle['學生名單'];
        if (!sheet) return res.json({ success: true });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('學號') === id);
        if (row) await row.delete();
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 班級管理 API 增強 ===
// 更新班級
app.put('/api/classes/:code', async (req, res) => {
    try {
        const { code } = req.params;
        const { name, division, teacher } = req.body;
        const sheet = doc.sheetsByTitle['班級列表'];
        if (!sheet) return res.json({ success: false });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('班級代碼') === code);
        if (!row) return res.json({ success: false, message: '班級不存在' });
        
        if (name) row.set('班級名稱', name);
        if (division) row.set('部別', division);
        if (teacher) row.set('導師', teacher);
        await row.save();
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 刪除班級
app.delete('/api/classes/:code', async (req, res) => {
    try {
        const { code } = req.params;
        const sheet = doc.sheetsByTitle['班級列表'];
        if (!sheet) return res.json({ success: true });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('班級代碼') === code);
        if (row) await row.delete();
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 課程管理 API 增強 ===
// 更新課程
app.put('/api/courses/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const { subject, classCode, day, period, time, room, lat, lon } = req.body;
        const sheet = doc.sheetsByTitle['課程列表'];
        if (!sheet) return res.json({ success: false });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('課程ID') === id);
        if (!row) return res.json({ success: false, message: '課程不存在' });
        
        if (subject) row.set('科目', subject);
        if (classCode) row.set('班級', classCode);
        if (day !== undefined) row.set('星期', day);
        if (period !== undefined) row.set('節次', period);
        if (time) row.set('上課時間', time);
        if (room) row.set('教室', room);
        if (lat !== undefined) row.set('GPS緯度', lat);
        if (lon !== undefined) row.set('GPS經度', lon);
        await row.save();
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 刪除課程
app.delete('/api/courses/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const sheet = doc.sheetsByTitle['課程列表'];
        if (!sheet) return res.json({ success: true });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('課程ID') === id);
        if (row) await row.delete();
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 匯出 Excel ===
app.get('/api/export/excel', async (req, res) => {
    try {
        const { startDate, endDate, classCode, type } = req.query;
        const recordSheet = doc.sheetsByTitle['簽到紀錄'];
        const studentSheet = doc.sheetsByTitle['學生名單'];
        const courseSheet = doc.sheetsByTitle['課程列表'];
        
        if (!recordSheet || !studentSheet) {
            return res.status(400).json({ success: false, message: '資料表不存在' });
        }
        
        const records = await recordSheet.getRows();
        const students = await studentSheet.getRows();
        const courses = courseSheet ? await courseSheet.getRows() : [];
        
        let data = [];
        
        if (type === 'summary') {
            // 學生出席率摘要
            for (const student of students) {
                const studentId = student.get('學號');
                if (classCode && student.get('班級') !== classCode) continue;
                
                const studentRecords = records.filter(r => {
                    const date = r.get('簽到時間')?.split(' ')[0];
                    const matchDate = (!startDate || date >= startDate) && (!endDate || date <= endDate);
                    return r.get('學號') === studentId && matchDate;
                });
                
                const total = studentRecords.length;
                const attended = studentRecords.filter(r => r.get('狀態') === '已報到').length;
                const late = studentRecords.filter(r => r.get('狀態') === '遲到').length;
                const absent = studentRecords.filter(r => r.get('狀態') === '缺席').length;
                const rate = total > 0 ? Math.round((attended + late) / total * 100) : 100;
                
                data.push({
                    學號: studentId,
                    姓名: student.get('姓名'),
                    班級: student.get('班級'),
                    總堂數: total,
                    出席: attended,
                    遲到: late,
                    缺席: absent,
                    出席率: rate + '%'
                });
            }
        } else {
            // 詳細出缺紀錄
            for (const r of records) {
                const date = r.get('簽到時間')?.split(' ')[0];
                if (startDate && date < startDate) continue;
                if (endDate && date > endDate) continue;
                
                const student = students.find(s => s.get('學號') === r.get('學號'));
                if (classCode && student?.get('班級') !== classCode) continue;
                
                const course = courses.find(c => c.get('課程ID') === r.get('課程ID'));
                
                data.push({
                    日期: date || '',
                    時間: r.get('簽到時間')?.split(' ')[1] || '',
                    學號: r.get('學號'),
                    姓名: student?.get('姓名') || '',
                    班級: student?.get('班級') || '',
                    課程: course?.get('科目') || '',
                    狀態: r.get('狀態'),
                    遲到分鐘: r.get('遲到分鐘') || 0,
                    備註: r.get('備註') || ''
                });
            }
        }
        
        // 產生 CSV
        if (data.length === 0) {
            return res.json({ success: false, message: '無資料' });
        }
        
        const headers = Object.keys(data[0]);
        const csv = '\uFEFF' + headers.join(',') + '\n' + 
            data.map(row => headers.map(h => '"' + (row[h] || '').toString().replace(/"/g, '""') + '"').join(',')).join('\n');
        
        res.setHeader('Content-Type', 'text/csv; charset=utf-8');
        res.setHeader('Content-Disposition', 'attachment; filename=attendance_' + new Date().toISOString().split('T')[0] + '.csv');
        res.send(csv);
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 手動調整出席紀錄 ===
app.put('/api/records/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const { status, note } = req.body;
        const sheet = doc.sheetsByTitle['簽到紀錄'];
        if (!sheet) return res.json({ success: false });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.rowNumber.toString() === id || r.get('活動ID') + '_' + r.get('學號') === id);
        if (!row) return res.json({ success: false, message: '找不到紀錄' });
        
        if (status) row.set('狀態', status);
        if (note !== undefined) row.set('備註', note);
        row.set('修改時間', new Date().toLocaleString('zh-TW'));
        await row.save();
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 新增手動出席紀錄
app.post('/api/records/manual', async (req, res) => {
    try {
        const { studentId, courseId, date, status, note } = req.body;
        const sheet = await getOrCreateSheet('簽到紀錄', ['活動ID', '學號', '簽到時間', '狀態', '遲到分鐘', 'GPS緯度', 'GPS經度', '備註', '修改時間']);
        
        await sheet.addRow({
            '活動ID': 'MANUAL_' + Date.now(),
            '學號': studentId,
            '簽到時間': date + ' 00:00:00',
            '狀態': status || '已報到',
            '遲到分鐘': 0,
            'GPS緯度': '',
            'GPS經度': '',
            '備註': note || '手動新增',
            '修改時間': new Date().toLocaleString('zh-TW')
        });
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 獎勵系統 ===
// 取得全勤學生
app.get('/api/rewards/perfect-attendance', async (req, res) => {
    try {
        const { startDate, endDate } = req.query;
        const recordSheet = doc.sheetsByTitle['簽到紀錄'];
        const studentSheet = doc.sheetsByTitle['學生名單'];
        
        if (!recordSheet || !studentSheet) {
            return res.json({ students: [] });
        }
        
        const records = await recordSheet.getRows();
        const students = await studentSheet.getRows();
        const perfectStudents = [];
        
        for (const student of students) {
            const studentId = student.get('學號');
            const studentRecords = records.filter(r => {
                const date = r.get('簽到時間')?.split(' ')[0];
                const matchDate = (!startDate || date >= startDate) && (!endDate || date <= endDate);
                return r.get('學號') === studentId && matchDate;
            });
            
            const total = studentRecords.length;
            if (total === 0) continue;
            
            const absent = studentRecords.filter(r => r.get('狀態') === '缺席').length;
            const late = studentRecords.filter(r => r.get('狀態') === '遲到').length;
            
            if (absent === 0 && late === 0) {
                perfectStudents.push({
                    studentId,
                    name: student.get('姓名'),
                    classCode: student.get('班級'),
                    lineId: student.get('LINE_ID'),
                    totalClasses: total
                });
            }
        }
        
        res.json({ students: perfectStudents });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 發送獎勵通知
app.post('/api/rewards/send', async (req, res) => {
    try {
        const { studentIds, message } = req.body;
        const studentSheet = doc.sheetsByTitle['學生名單'];
        if (!studentSheet) return res.json({ success: false });
        
        const students = await studentSheet.getRows();
        let sent = 0;
        
        for (const studentId of studentIds) {
            const student = students.find(s => s.get('學號') === studentId);
            if (student && student.get('LINE_ID')) {
                try {
                    const text = message || `🏆 恭喜！\n\n${student.get('姓名')} 同學，您達成全勤！\n\n感謝您的認真出席，繼續保持！💪`;
                    await lineClient.pushMessage(student.get('LINE_ID'), { type: 'text', text });
                    sent++;
                } catch (e) { console.log('發送失敗:', e.message); }
            }
        }
        
        res.json({ success: true, sent });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 課前提醒排程 ===
app.post('/api/reminders/schedule', async (req, res) => {
    try {
        const { courseId, minutesBefore } = req.body;
        // 這裡可以整合 node-cron 或其他排程工具
        // 目前先返回成功，實際排程需要額外設定
        res.json({ success: true, message: '提醒已排程' });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 立即發送課前提醒
app.post('/api/reminders/send-now', async (req, res) => {
    try {
        const { courseId } = req.body;
        const courseSheet = doc.sheetsByTitle['課程列表'];
        const studentSheet = doc.sheetsByTitle['學生名單'];
        
        if (!courseSheet || !studentSheet) {
            return res.json({ success: false });
        }
        
        const courses = await courseSheet.getRows();
        const course = courses.find(c => c.get('課程ID') === courseId);
        if (!course) return res.json({ success: false, message: '找不到課程' });
        
        const classCode = course.get('班級');
        const students = await studentSheet.getRows();
        const classStudents = students.filter(s => s.get('班級') === classCode && s.get('LINE_ID'));
        
        let sent = 0;
        for (const student of classStudents) {
            try {
                await lineClient.pushMessage(student.get('LINE_ID'), {
                    type: 'text',
                    text: `⏰ 上課提醒\n\n${course.get('科目')} 即將開始！\n📍 ${course.get('教室') || '教室'}\n⏰ ${course.get('上課時間')}\n\n請準時出席！`
                });
                sent++;
            } catch (e) { }
        }
        
        res.json({ success: true, sent });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 多位置 GPS 管理 ===
app.get('/api/locations', async (req, res) => {
    try {
        const sheet = await getOrCreateSheet('GPS位置', ['位置ID', '名稱', '緯度', '經度', '半徑', '備註']);
        const rows = await sheet.getRows();
        res.json(rows.map(r => ({
            id: r.get('位置ID'),
            name: r.get('名稱'),
            lat: parseFloat(r.get('緯度')) || 0,
            lon: parseFloat(r.get('經度')) || 0,
            radius: parseInt(r.get('半徑')) || 50,
            note: r.get('備註')
        })));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/locations', async (req, res) => {
    try {
        const { name, lat, lon, radius, note } = req.body;
        const sheet = await getOrCreateSheet('GPS位置', ['位置ID', '名稱', '緯度', '經度', '半徑', '備註']);
        
        const locationId = 'LOC_' + Date.now();
        await sheet.addRow({
            '位置ID': locationId,
            '名稱': name,
            '緯度': lat,
            '經度': lon,
            '半徑': radius || 50,
            '備註': note || ''
        });
        
        res.json({ success: true, locationId });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

app.delete('/api/locations/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const sheet = doc.sheetsByTitle['GPS位置'];
        if (!sheet) return res.json({ success: true });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('位置ID') === id);
        if (row) await row.delete();
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// === 圖表數據 API ===
app.get('/api/charts/attendance-trend', async (req, res) => {
    try {
        const { days } = req.query;
        const numDays = parseInt(days) || 7;
        const recordSheet = doc.sheetsByTitle['簽到紀錄'];
        
        if (!recordSheet) {
            return res.json({ data: [] });
        }
        
        const records = await recordSheet.getRows();
        const today = new Date();
        const data = [];
        
        for (let i = numDays - 1; i >= 0; i--) {
            const date = new Date(today);
            date.setDate(date.getDate() - i);
            const dateStr = date.toISOString().split('T')[0];
            
            const dayRecords = records.filter(r => r.get('簽到時間')?.startsWith(dateStr));
            const total = dayRecords.length;
            const attended = dayRecords.filter(r => r.get('狀態') === '已報到').length;
            const late = dayRecords.filter(r => r.get('狀態') === '遲到').length;
            const absent = dayRecords.filter(r => r.get('狀態') === '缺席').length;
            const rate = total > 0 ? Math.round((attended + late) / total * 100) : 0;
            
            data.push({
                date: dateStr,
                label: (date.getMonth() + 1) + '/' + date.getDate(),
                total,
                attended,
                late,
                absent,
                rate
            });
        }
        
        res.json({ data });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.get('/api/charts/class-comparison', async (req, res) => {
    try {
        const recordSheet = doc.sheetsByTitle['簽到紀錄'];
        const studentSheet = doc.sheetsByTitle['學生名單'];
        const classSheet = doc.sheetsByTitle['班級列表'];
        
        if (!recordSheet || !studentSheet) {
            return res.json({ data: [] });
        }
        
        const records = await recordSheet.getRows();
        const students = await studentSheet.getRows();
        const classes = classSheet ? await classSheet.getRows() : [];
        const data = [];
        
        // 取得所有班級代碼
        const classCodes = [...new Set(students.map(s => s.get('班級')))];
        
        for (const code of classCodes) {
            const classStudents = students.filter(s => s.get('班級') === code);
            const studentIds = classStudents.map(s => s.get('學號'));
            const classRecords = records.filter(r => studentIds.includes(r.get('學號')));
            
            const total = classRecords.length;
            const attended = classRecords.filter(r => r.get('狀態') === '已報到').length;
            const late = classRecords.filter(r => r.get('狀態') === '遲到').length;
            const rate = total > 0 ? Math.round((attended + late) / total * 100) : 0;
            
            const classInfo = classes.find(c => c.get('班級代碼') === code);
            
            data.push({
                code,
                name: classInfo?.get('班級名稱') || code,
                studentCount: classStudents.length,
                rate
            });
        }
        
        data.sort((a, b) => b.rate - a.rate);
        res.json({ data });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// ===== 啟動伺服器 =====

const PORT = process.env.PORT || 3000;

initGoogleSheets()
    .then(() => {
        app.listen(PORT, () => {
            console.log(`🚀 簽到系統已啟動，埠號 ${PORT}`);
        });
    })
    .catch(err => {
        console.error('初始化失敗:', err);
        process.exit(1);
    });
