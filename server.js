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
