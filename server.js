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
    // 統一使用 YYYY-MM-DD 格式
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
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
    try {
        const sheet = await getOrCreateSheet('學生名單', [
            '學號', '姓名', '班級', 'LINE_ID', 'LINE名稱', '註冊時間', '狀態'
        ]);
        const rows = await sheet.getRows();
        return rows.find(row => row.get('LINE_ID') === lineUserId);
    } catch (error) {
        console.error('❌ getStudent 錯誤:', error);
        return null;
    }
}

/**
 * 註冊學生
 */
/**
 * 註冊學生
 * 支援「同一學號換手機/換 LINE」自動覆寫 LINE_ID
 */
async function registerStudent(lineUserId, lineName, studentId, studentName, className) {
    try {
        await doc.loadInfo();
        const sheet = await getOrCreateSheet('學生名單', [
            '學號', '姓名', '班級', 'LINE_ID', 'LINE名稱', '註冊時間', '狀態'
        ]);
        
        // 檢查學號是否已存在
        const rows = await sheet.getRows();
        const existing = rows.find(row => row.get('學號') === studentId);
        
        if (existing) {
            const oldLineId = existing.get('LINE_ID') || '';
            
            // 1️⃣ 完全同一個 LINE 帳號：視為重複註冊
            if (oldLineId === lineUserId) {
                return { success: false, message: '您已經註冊過了！' };
            }
            
            // 2️⃣ 學號已存在但 LINE_ID 不同：視為「換手機 / 換 LINE 帳號」
            existing.set('LINE_ID', lineUserId);
            existing.set('LINE名稱', lineName);
            existing.set('狀態', '正常');
            existing.set('註冊時間', formatDateTime(new Date()));
            await existing.save();
            
            console.log(`🔄 學號 ${studentId} 重新綁定 LINE_ID，舊=${oldLineId} 新=${lineUserId}`);
            
            return { 
                success: true, 
                message: '偵測到您使用新裝置，已為您更新綁定資料。' 
            };
        }
        
        // 3️⃣ 全新註冊
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
    } catch (error) {
        console.error('❌ registerStudent 錯誤:', error);
        return { success: false, message: '註冊失敗: ' + error.message };
    }
}

/**
 * 取得課程資料
 */
async function getCourse(courseId) {
    try {
        // 強制重新載入整個文檔
        const sheet = doc.sheetsByTitle['課程列表'];
        if (!sheet) {
            console.log('❌ 課程列表不存在');
            return null;
        }
        
        // 使用 limit 參數讀取
        const rows = await sheet.getRows({ limit: 500 });
        
        const course = rows.find(row => row.get('課程ID') === courseId);
        if (course) {
            const radius = course.get('簽到範圍');
            console.log(`📖 讀取課程 ${courseId}:`, {
                科目: course.get('科目'),
                簽到範圍: radius,
                簽到範圍類型: typeof radius
            });
        } else {
            console.log(`❌ 找不到課程 ${courseId}`);
        }
        return course;
    } catch (error) {
        console.error('getCourse 錯誤:', error);
        return null;
    }
}

/**
 * 取得今日課程活動
 */
async function getTodaySession(courseId) {
    try {
        const today = getTodayString();
        const sheet = await getOrCreateSheet('簽到活動', [
            '活動ID', '課程ID', '日期', '開始時間', '結束時間', 'QR碼內容', '狀態'
        ]);
        const rows = await sheet.getRows();
        
        // 找今天的活動（不限制狀態，只要不是「已結束」）
        const session = rows.find(row => {
            const rowCourseId = row.get('課程ID');
            const rowDate = row.get('日期');
            const rowStatus = row.get('狀態');
            
            // 日期可能是不同格式，都嘗試匹配
            const dateMatch = rowDate === today || 
                             rowDate === today.replace(/-/g, '/') ||
                             rowDate?.includes(today.split('-')[1] + '/' + today.split('-')[2]);
            
            return rowCourseId === courseId && 
                   dateMatch && 
                   rowStatus !== '已結束';
        });
        
        return session;
    } catch (error) {
        console.error('❌ getTodaySession 錯誤:', error);
        return null;
    }
}

/**
 * 檢查是否已簽到
 */
async function checkExistingAttendance(sessionId, studentId) {
    try {
        const sheet = doc.sheetsByTitle['簽到紀錄'];
        if (!sheet) return null;
        
        const rows = await sheet.getRows();
        return rows.find(row => 
            row.get('活動ID') === sessionId && 
            row.get('學號') === studentId
        );
    } catch (e) {
        console.error('檢查簽到錯誤:', e);
        return null;
    }
}

/**
 * 記錄簽到並發送通知
 */
async function recordAttendance(sessionId, studentId, status, lateMinutes = 0, gpsLat = '', gpsLon = '', sendNotification = true) {
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
    
    // 發送簽到狀態通知（準時、遲到、缺席都發送）
    if (sendNotification) {
        try {
            const studentSheet = doc.sheetsByTitle['學生名單'];
            if (studentSheet) {
                const students = await studentSheet.getRows();
                const student = students.find(s => s.get('學號') === studentId);
                
                if (student && student.get('LINE_ID')) {
                    // 取得課程資訊
                    const sessionSheet = doc.sheetsByTitle['簽到活動'];
                    const sessions = await sessionSheet.getRows();
                    const session = sessions.find(s => s.get('活動ID') === sessionId);
                    
                    if (session) {
                        const courseSheet = doc.sheetsByTitle['課程列表'];
                        const courses = await courseSheet.getRows();
                        const course = courses.find(c => c.get('課程ID') === session.get('課程ID'));
                        
                        if (course) {
                            let notifyText = '';
                            if (status === '已報到') {
                                notifyText = `✅ 簽到成功\n\n📚 課程：${course.get('科目')}\n📅 日期：${session.get('日期')}\n✨ 狀態：準時報到\n\n繼續保持！💪`;
                            } else if (status === '遲到') {
                                notifyText = `⚠️ 遲到通知\n\n📚 課程：${course.get('科目')}\n📅 日期：${session.get('日期')}\n⏰ 遲到：${lateMinutes} 分鐘\n\n請下次準時出席！`;
                            } else if (status === '缺席') {
                                notifyText = `❌ 缺席通知\n\n📚 課程：${course.get('科目')}\n📅 日期：${session.get('日期')}\n\n如有疑問請聯繫教師。`;
                            }
                            
                            if (notifyText) {
                                await lineClient.pushMessage(student.get('LINE_ID'), {
                                    type: 'text',
                                    text: notifyText
                                });
                                console.log(`✉️ 已發送${status}通知給 ${studentId}`);
                            }
                        }
                    }
                }
            }
        } catch (e) {
            console.error('發送簽到通知失敗:', e.message);
        }
    }
    
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
    try {
        const sheet = await getOrCreateSheet('班級列表', [
            '班級代碼', '班級名稱', '導師', '人數', '建立時間'
        ]);
        const rows = await sheet.getRows();
        return rows.map(row => ({
            code: row.get('班級代碼'),
            name: row.get('班級名稱')
        }));
    } catch (error) {
        console.error('❌ getClasses 錯誤:', error);
        return [];
    }
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
    let userName = '同學';
    try {
        const userProfile = await lineClient.getProfile(userId);
        userName = userProfile.displayName || '同學';
    } catch (e) {
        // 無法取得用戶資料，使用預設名稱
    }
    
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
        if (text.startsWith('直接簽到:')) {
            return handleDirectCheckin(event, userId, text);
        }
        
        if (text.startsWith('GPS簽到:')) {
            return handleGPSCheckin(event, userId, text);
        }
        
        // 舊版相容
        if (text.startsWith('簽到:')) {
            return handleCheckinRequest(event, userId, text);
        }
        
        // 檢查用戶狀態（是否在流程中）
        const state = userStates.get(userId);
        if (state) {
            if (state.step === 'addNewClass') {
                return handleAddNewClass(event, userId, text, state);
            }
            if (state.step === 'removeClass') {
                return handleRemoveClass(event, userId, text, state);
            }
            return handleRegistrationFlow(event, userId, userName, text, state);
        }
        
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
        
        case '解除綁定':
        case '取消綁定':
            if (!student) {
                return replyText(event, '❌ 您尚未綁定帳號！');
            }
            // 確認解除綁定
            userStates.set(userId, { step: 'confirmUnbind', studentId: student.get('學號') });
            return replyText(event, `⚠️ 確認解除綁定？\n\n學號：${student.get('學號')}\n姓名：${student.get('姓名')}\n\n輸入「確認」解除綁定，或輸入其他文字取消。`);
        
        case '確認':
            const state = userStates.get(userId);
            if (state && state.step === 'confirmUnbind') {
                // 執行解除綁定
                try {
                    const studentSheet = doc.sheetsByTitle['學生名單'];
                    const rows = await studentSheet.getRows();
                    const studentRow = rows.find(r => r.get('學號') === state.studentId);
                    if (studentRow) {
                        studentRow.set('LINE_ID', '');
                        studentRow.set('LINE名稱', '');
                        await studentRow.save();
                    }
                    userStates.delete(userId);
                    return replyText(event, '✅ 已解除綁定！\n\n感謝您這學期的使用。\n如需重新綁定，請輸入「註冊」。');
                } catch (e) {
                    return replyText(event, '❌ 解除綁定失敗，請稍後再試。');
                }
            }
            return replyText(event, '❌ 無效的操作。');
        
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
        
        case '我的班級':
        case '班級資料':
            if (!student) {
                return replyText(event, '❌ 您尚未註冊！\n\n請輸入「註冊」開始綁定學號。');
            }
            return replyClassDetails(event, student);
        
        case '加入班級':
        case '新班級':
            if (!student) {
                return replyText(event, '❌ 您尚未註冊！\n\n請先輸入「註冊」綁定學號後，再加入班級。');
            }
            userStates.set(userId, { step: 'addNewClass', studentId: student.get('學號') });
            const availableClasses = await getClasses();
            const currentClasses = (student.get('班級') || '').split('、').map(c => c.trim()).filter(c => c);
            const newClasses = availableClasses.filter(c => !currentClasses.includes(c.code));
            if (newClasses.length === 0) {
                userStates.delete(userId);
                return replyText(event, '📋 您已加入所有可用班級！\n\n目前班級：' + currentClasses.join('、'));
            }
            let classListMsg = '📝 加入新班級\n\n您目前的班級：' + (currentClasses.length > 0 ? currentClasses.join('、') : '無') + '\n\n可加入的班級：\n';
            newClasses.forEach(c => { classListMsg += '• ' + c.code + ' - ' + c.name + '\n'; });
            classListMsg += '\n請輸入要加入的【班級代碼】：';
            return replyText(event, classListMsg);
        
        case '退出班級':
            if (!student) {
                return replyText(event, '❌ 您尚未註冊！');
            }
            const myClasses = (student.get('班級') || '').split('、').map(c => c.trim()).filter(c => c);
            if (myClasses.length <= 1) {
                return replyText(event, '❌ 您只有一個班級，無法退出！\n\n如需完全解除綁定，請輸入「解除綁定」。');
            }
            userStates.set(userId, { step: 'removeClass', studentId: student.get('學號'), currentClasses: myClasses });
            return replyText(event, '📝 退出班級\n\n您目前的班級：\n' + myClasses.join('、') + '\n\n請輸入要退出的【班級代碼】：');
        
        case '全部紀錄':
        case '所有紀錄':
            if (!student) {
                return replyText(event, '❌ 您尚未註冊！');
            }
            return replyAllClassesAttendance(event, student);
        
        case '說明':
        case '幫助':
        case 'help':
            return replyHelp(event);
        
        case '我的ID':
        case 'myid':
            // 除錯用：顯示用戶的 LINE ID
            const storedLineId = student ? student.get('LINE_ID') : '未註冊';
            return replyText(event, `🔍 LINE ID 資訊\n\n📱 您目前的 ID：\n${userId}\n\n📋 試算表中的 ID：\n${storedLineId}\n\n${userId === storedLineId ? '✅ ID 一致' : '❌ ID 不一致！'}`);
        
        default:
            if (!student) {
                return replyText(event, `👋 歡迎 ${userName}！\n\n您尚未註冊，請輸入「註冊」綁定學號後才能使用簽到功能。\n\n輸入「說明」查看更多指令。`);
            }
            return replyText(event, `👋 ${student.get('姓名')} 同學您好！\n\n📌 可用指令：\n• 我的資料\n• 我的班級\n• 出席紀錄\n• 全部紀錄\n• 加入班級\n• 退出班級\n• 說明\n\n📍 簽到請掃描教師提供的 QR Code`);
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
 * 直接簽到（掃老師手機 QR Code）
 * 不需要 GPS 驗證，直接簽到成功
 */
async function handleDirectCheckin(event, userId, text) {
    const student = await getStudent(userId);
    console.log('📲 直接簽到 - userId =', userId, ' student =', student ? student.get('學號') : '未找到');
    if (!student) {
        return replyText(event, '❌ 您尚未註冊！\n\n請先輸入「註冊」綁定學號。');
    }
    
    const parts = text.replace('直接簽到:', '').split('|');
    if (parts.length < 2) {
        return replyText(event, '❌ 無效的簽到碼！');
    }
    
    const [courseId, sessionId] = parts;
    
    const course = await getCourse(courseId);
    if (!course) {
        return replyText(event, '❌ 找不到此課程！');
    }
    
    // 取得活動
    let session = await getTodaySession(courseId);
    if (!session) {
        const sessionSheet = await getOrCreateSheet('簽到活動', [
            '活動ID', '課程ID', '日期', '開始時間', '結束時間', 'QR碼內容', '狀態'
        ]);
        const rows = await sessionSheet.getRows();
        session = rows.find(r => r.get('活動ID') === sessionId && r.get('狀態') !== '已結束');
    }
    
    if (!session) {
        return replyText(event, '❌ 此簽到活動已結束或不存在！');
    }
    
    const actualSessionId = session.get('活動ID');
    
    // 檢查是否已簽到
    const existingRecord = await checkExistingAttendance(actualSessionId, student.get('學號'));
    if (existingRecord) {
        return replyText(event, `✅ 您已經簽到過了！\n\n📚 課程：${course.get('科目')}\n⏰ 簽到時間：${existingRecord.get('簽到時間')}`);
    }
    
    // 計算是否遲到
    const startTime = session.get('開始時間');
    const lateMinutes = parseInt(course.get('遲到標準')) || 10;
    const now = new Date();
    const [startHour, startMin] = (startTime || '08:00').split(':').map(Number);
    const startDate = new Date();
    startDate.setHours(startHour, startMin, 0, 0);
    
    const diffMinutes = Math.floor((now - startDate) / 60000);
    const status = diffMinutes > lateMinutes ? '遲到' : '已報到';
    
    // 記錄簽到（不記錄 GPS）
    const result = await recordAttendance(
        actualSessionId,
        student.get('學號'),
        status,
        diffMinutes > lateMinutes ? diffMinutes : 0,
        '', ''
    );
    
    if (result.success) {
        const emoji = status === '已報到' ? '✅' : '⚠️';
        let msg = `${emoji} 簽到成功！\n\n📚 課程：${course.get('科目')}\n👤 學生：${student.get('姓名')}\n📍 方式：掃描 QR Code\n✨ 狀態：${status}`;
        if (status === '遲到') {
            msg += `\n⏰ 遲到 ${diffMinutes} 分鐘`;
        }
        return replyText(event, msg);
    } else {
        return replyText(event, `❌ 簽到失敗：${result.message}`);
    }
}

/**
 * GPS 簽到（學生點連結自己簽到）
 * 需要 GPS 驗證
 */
async function handleGPSCheckin(event, userId, text) {
    const student = await getStudent(userId);
    console.log('📍 GPS 簽到 - userId =', userId, ' student =', student ? student.get('學號') : '未找到');
    if (!student) {
        return replyText(event, `❌ 找不到您的帳號！\n\n📱 收到的 ID：\n${userId}\n\n請輸入「我的ID」比對，或輸入「註冊」重新綁定。`);
    }
    
    const parts = text.replace('GPS簽到:', '').split('|');
    if (parts.length < 2) {
        return replyText(event, '❌ 無效的簽到碼！');
    }
    
    const [courseId, sessionId] = parts;
    
    const course = await getCourse(courseId);
    if (!course) {
        return replyText(event, '❌ 找不到此課程！');
    }
    
    // 取得活動
    let session = await getTodaySession(courseId);
    if (!session) {
        const sessionSheet = await getOrCreateSheet('簽到活動', [
            '活動ID', '課程ID', '日期', '開始時間', '結束時間', 'QR碼內容', '狀態'
        ]);
        const rows = await sessionSheet.getRows();
        session = rows.find(r => r.get('活動ID') === sessionId && r.get('狀態') !== '已結束');
    }
    
    if (!session) {
        return replyText(event, '❌ 此簽到活動已結束或不存在！');
    }
    
    const actualSessionId = session.get('活動ID');
    
    // 檢查是否已簽到
    const existingRecord = await checkExistingAttendance(actualSessionId, student.get('學號'));
    if (existingRecord) {
        return replyText(event, `✅ 您已經簽到過了！\n\n📚 課程：${course.get('科目')}\n⏰ 簽到時間：${existingRecord.get('簽到時間')}`);
    }
    
    // 取得簽到設定（從 Google Sheets 直接讀取）
    const classroomLat = parseFloat(course.get('教室緯度')) || 0;
    const classroomLon = parseFloat(course.get('教室經度')) || 0;
    const rawRadius = course.get('簽到範圍');
    
    // 詳細記錄讀取到的值
    console.log('🔍 簽到範圍原始值:', {
        rawRadius,
        type: typeof rawRadius,
        isEmpty: rawRadius === '',
        isNull: rawRadius === null,
        isUndefined: rawRadius === undefined
    });
    
    // 解析 radius
    let checkRadius;
    if (rawRadius === '' || rawRadius === undefined || rawRadius === null) {
        checkRadius = 100;  // 預設值
        console.log('⚠️ 使用預設值 100');
    } else {
        checkRadius = parseInt(rawRadius);
        console.log('✅ 解析後的 checkRadius:', checkRadius);
    }
    
    console.log('📍 GPS 簽到設定:', { 
        courseId, 
        科目: course.get('科目'),
        classroomLat, 
        classroomLon, 
        rawRadius, 
        checkRadius 
    });
    
    // 簽到模式判斷
    // -1: 現場簽到（只能掃 QR Code，不能用連結）
    if (checkRadius === -1) {
        return replyText(event, '📱 此課程設定為「現場簽到」\n\n請到教室掃描老師手機上的 QR Code 簽到。');
    }
    
    // 0 或無設定: 不限制（線上課程），直接簽到
    // 有設定 GPS 座標且 checkRadius > 0: 需要 GPS 驗證
    if (classroomLat !== 0 && classroomLon !== 0 && checkRadius > 0) {
        userStates.set(userId, { 
            step: 'waitingLocation',
            courseId,
            sessionId: actualSessionId,
            courseName: course.get('科目'),
            classroomLat,
            classroomLon,
            checkRadius,
            lateMinutes: parseInt(course.get('遲到標準')) || 10,
            startTime: session.get('開始時間')
        });
        
        return lineClient.replyMessage(event.replyToken, {
            type: 'template',
            altText: '📍 請傳送您的位置以完成簽到',
            template: {
                type: 'buttons',
                title: `📍 GPS 簽到 - ${course.get('科目')}`,
                text: `請傳送位置驗證\n允許範圍：${checkRadius} 公尺`,
                actions: [
                    {
                        type: 'uri',
                        label: '📍 傳送我的位置',
                        uri: 'https://line.me/R/nv/location'
                    }
                ]
            }
        });
    }
    
    // 不限制 GPS（線上課程），直接簽到
    const startTime = session.get('開始時間');
    const lateMinutes = parseInt(course.get('遲到標準')) || 10;
    const now = new Date();
    const [startHour, startMin] = (startTime || '08:00').split(':').map(Number);
    const startDate = new Date();
    startDate.setHours(startHour, startMin, 0, 0);
    
    const diffMinutes = Math.floor((now - startDate) / 60000);
    const status = diffMinutes > lateMinutes ? '遲到' : '已報到';
    
    const result = await recordAttendance(
        actualSessionId,
        student.get('學號'),
        status,
        diffMinutes > lateMinutes ? diffMinutes : 0,
        '', ''
    );
    
    if (result.success) {
        const emoji = status === '已報到' ? '✅' : '⚠️';
        let msg = `${emoji} 簽到成功！\n\n📚 課程：${course.get('科目')}\n👤 學生：${student.get('姓名')}\n✨ 狀態：${status}`;
        if (status === '遲到') {
            msg += `\n⏰ 遲到 ${diffMinutes} 分鐘`;
        }
        return replyText(event, msg);
    } else {
        return replyText(event, `❌ 簽到失敗：${result.message}`);
    }
}

/**
 * 處理簽到請求（舊版相容 - 直接簽到）
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
    
    const course = await getCourse(courseId);
    if (!course) {
        return replyText(event, '❌ 找不到此課程！');
    }
    
    let session = await getTodaySession(courseId);
    if (!session) {
        const sessionSheet = await getOrCreateSheet('簽到活動', [
            '活動ID', '課程ID', '日期', '開始時間', '結束時間', 'QR碼內容', '狀態'
        ]);
        const rows = await sessionSheet.getRows();
        session = rows.find(r => r.get('活動ID') === sessionId && r.get('狀態') !== '已結束');
    }
    
    if (!session) {
        return replyText(event, '❌ 此簽到活動已結束或不存在！');
    }
    
    const actualSessionId = session.get('活動ID');
    
    const existingRecord = await checkExistingAttendance(actualSessionId, student.get('學號'));
    if (existingRecord) {
        return replyText(event, `✅ 您已經簽到過了！\n\n📚 課程：${course.get('科目')}\n⏰ 簽到時間：${existingRecord.get('簽到時間')}`);
    }
    
    // 舊版直接簽到（不需要 GPS）
    const startTime = session.get('開始時間');
    const lateMinutes = parseInt(course.get('遲到標準')) || 10;
    const now = new Date();
    const [startHour, startMin] = (startTime || '08:00').split(':').map(Number);
    const startDate = new Date();
    startDate.setHours(startHour, startMin, 0, 0);
    
    const diffMinutes = Math.floor((now - startDate) / 60000);
    const status = diffMinutes > lateMinutes ? '遲到' : '已報到';
    
    const result = await recordAttendance(
        actualSessionId,
        student.get('學號'),
        status,
        diffMinutes > lateMinutes ? diffMinutes : 0,
        '', ''
    );
    
    if (result.success) {
        const emoji = status === '已報到' ? '✅' : '⚠️';
        let msg = `${emoji} 簽到成功！\n\n📚 課程：${course.get('科目')}\n👤 學生：${student.get('姓名')}\n✨ 狀態：${status}`;
        if (status === '遲到') {
            msg += `\n⏰ 遲到 ${diffMinutes} 分鐘`;
        }
        return replyText(event, msg);
    } else {
        return replyText(event, `❌ 簽到失敗：${result.message}`);
    }
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
    
    if (!student) {
        userStates.delete(userId);
        return replyText(event, '❌ 找不到您的學生資料！\n\n請先輸入「註冊」綁定學號。');
    }
    
    // 每次都重新讀取課程設定（確保使用最新的簽到範圍）
    const course = await getCourse(state.courseId);
    if (!course) {
        userStates.delete(userId);
        return replyText(event, '❌ 課程不存在！');
    }
    
    // 重新讀取最新的簽到設定
    const classroomLat = parseFloat(course.get('教室緯度')) || state.classroomLat;
    const classroomLon = parseFloat(course.get('教室經度')) || state.classroomLon;
    const rawRadius = course.get('簽到範圍');
    const checkRadius = rawRadius !== '' && rawRadius !== undefined && rawRadius !== null ? parseInt(rawRadius) : state.checkRadius;
    
    console.log('位置驗證 - 最新設定:', { courseId: state.courseId, checkRadius, rawRadius });
    
    // 計算距離
    const distance = calculateDistance(
        latitude, longitude,
        classroomLat, classroomLon
    );
    
    // 使用最新的設定範圍
    const allowedRadius = checkRadius;
    
    // 檢查是否在範圍內
    if (distance > allowedRadius) {
        // 不刪除狀態，允許重試
        state.retryCount = (state.retryCount || 0) + 1;
        
        // 最多重試 3 次
        if (state.retryCount >= 3) {
            userStates.delete(userId);
            return replyText(event, 
                `🚫 簽到失敗！\n\n已重試 ${state.retryCount} 次仍不在範圍內。\n📍 您的位置距離：${Math.round(distance)} 公尺\n📏 允許範圍：${allowedRadius} 公尺\n\n💡 建議：\n1. 到戶外或窗邊重新定位\n2. 聯繫老師使用現場 QR Code 簽到`
            );
        }
        
        // 允許重試
        return lineClient.replyMessage(event.replyToken, {
            type: 'template',
            altText: '📍 位置驗證失敗，請重試',
            template: {
                type: 'buttons',
                title: '📍 位置不在範圍內',
                text: `您的距離：${Math.round(distance)} 公尺\n允許範圍：${allowedRadius} 公尺\n\n請移動到教室範圍內重試`,
                actions: [
                    {
                        type: 'uri',
                        label: '🔄 重新傳送位置',
                        uri: 'https://line.me/R/nv/location'
                    }
                ]
            }
        });
    }
    
    // 計算是否遲到
    const now = new Date();
    const [startHour, startMin] = (state.startTime || '08:00').split(':').map(Number);
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
        message = `✅ 簽到成功！\n\n📚 課程：${state.courseName}\n⏰ 時間：${formatDateTime(now)}\n📍 距離教室：${Math.round(distance)} 公尺\n✨ 狀態：準時報到\n\n繼續保持！💪`;
    } else {
        message = `⚠️ 簽到成功（遲到）\n\n📚 課程：${state.courseName}\n⏰ 時間：${formatDateTime(now)}\n📍 距離教室：${Math.round(distance)} 公尺\n⏰ 遲到 ${lateMinutes} 分鐘\n\n下次請準時到達！`;
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
        `【基本指令】\n` +
        `• 註冊 - 綁定學號\n` +
        `• 我的資料 - 查看個人資訊\n` +
        `• 出席紀錄 - 最近簽到記錄\n` +
        `• 解除綁定 - 解除 LINE 綁定\n\n` +
        `【班級管理】\n` +
        `• 我的班級 - 查看班級詳細資料\n` +
        `• 加入班級 - 加入新的班級\n` +
        `• 退出班級 - 退出指定班級\n` +
        `• 全部紀錄 - 所有班級出缺席統計\n\n` +
        `【簽到方式】\n` +
        `掃描教師 QR Code → 分享位置 → 完成\n\n` +
        `💡 一個學號可加入多個班級`;
    
    return replyText(event, message);
}

// 處理加入新班級
async function handleAddNewClass(event, userId, text, state) {
    const classCode = text.trim();
    const allClasses = await getClasses();
    const targetClass = allClasses.find(c => c.code === classCode || c.code.toUpperCase() === classCode.toUpperCase());
    
    if (!targetClass) {
        userStates.delete(userId);
        return replyText(event, '❌ 找不到班級「' + text + '」！\n\n請重新輸入「加入班級」。');
    }
    
    try {
        await doc.loadInfo();
        const studentSheet = doc.sheetsByTitle['學生名單'];
        const rows = await studentSheet.getRows();
        const studentRow = rows.find(r => r.get('學號') === state.studentId);
        
        if (studentRow) {
            const currentClasses = (studentRow.get('班級') || '').split('、').map(c => c.trim()).filter(c => c);
            if (currentClasses.includes(targetClass.code)) {
                userStates.delete(userId);
                return replyText(event, '❌ 您已在「' + targetClass.code + '」班級中！');
            }
            currentClasses.push(targetClass.code);
            studentRow.set('班級', currentClasses.join('、'));
            await studentRow.save();
            userStates.delete(userId);
            return replyText(event, '✅ 成功加入班級！\n\n🏫 ' + targetClass.code + ' - ' + targetClass.name + '\n\n📋 您的所有班級：\n' + currentClasses.join('、'));
        }
        userStates.delete(userId);
        return replyText(event, '❌ 找不到您的資料。');
    } catch (e) {
        console.error('加入班級錯誤:', e);
        userStates.delete(userId);
        return replyText(event, '❌ 加入失敗: ' + e.message);
    }
}

// 處理退出班級
async function handleRemoveClass(event, userId, text, state) {
    const classCode = text.trim();
    if (!state.currentClasses.includes(classCode)) {
        userStates.delete(userId);
        return replyText(event, '❌ 您不在「' + classCode + '」班級中！');
    }
    
    try {
        await doc.loadInfo();
        const studentSheet = doc.sheetsByTitle['學生名單'];
        const rows = await studentSheet.getRows();
        const studentRow = rows.find(r => r.get('學號') === state.studentId);
        
        if (studentRow) {
            const newClasses = state.currentClasses.filter(c => c !== classCode);
            studentRow.set('班級', newClasses.join('、'));
            await studentRow.save();
            userStates.delete(userId);
            return replyText(event, '✅ 已退出班級「' + classCode + '」！\n\n📋 目前班級：\n' + newClasses.join('、'));
        }
        userStates.delete(userId);
        return replyText(event, '❌ 操作失敗。');
    } catch (e) {
        userStates.delete(userId);
        return replyText(event, '❌ 退出失敗: ' + e.message);
    }
}

// 回覆班級詳細資料
async function replyClassDetails(event, student) {
    const classesStr = student.get('班級') || '';
    const studentClasses = classesStr.split('、').map(c => c.trim()).filter(c => c);
    
    if (studentClasses.length === 0) {
        return replyText(event, '❌ 您尚未加入任何班級！\n\n請輸入「加入班級」。');
    }
    
    const classSheet = doc.sheetsByTitle['班級列表'];
    const courseSheet = doc.sheetsByTitle['課程列表'];
    
    let msg = '🏫 我的班級資料\n━━━━━━━━━━━━━━━\n';
    msg += '👤 ' + student.get('姓名') + ' (' + student.get('學號') + ')\n';
    msg += '📚 共 ' + studentClasses.length + ' 個班級\n\n';
    
    for (const classCode of studentClasses) {
        msg += '【' + classCode + '】';
        if (classSheet) {
            const classRows = await classSheet.getRows();
            const classInfo = classRows.find(r => r.get('班級代碼') === classCode);
            if (classInfo) {
                msg += ' ' + (classInfo.get('班級名稱') || '') + '\n';
                msg += '   👨‍🏫 導師：' + (classInfo.get('導師') || '未設定') + '\n';
            } else {
                msg += '\n';
            }
        } else {
            msg += '\n';
        }
        if (courseSheet) {
            const courseRows = await courseSheet.getRows();
            const classCourses = courseRows.filter(r => r.get('班級') === classCode);
            msg += '   📖 課程：' + classCourses.length + ' 門\n';
        }
        msg += '\n';
    }
    
    msg += '💡「加入班級」可加入新班級';
    return replyText(event, msg);
}

// 回覆所有班級出缺席
async function replyAllClassesAttendance(event, student) {
    const studentId = student.get('學號');
    const recordSheet = doc.sheetsByTitle['簽到紀錄'];
    
    if (!recordSheet) {
        return replyText(event, '📊 尚無簽到紀錄');
    }
    
    const allRecords = await recordSheet.getRows();
    const studentRecords = allRecords.filter(r => r.get('學號') === studentId);
    
    if (studentRecords.length === 0) {
        return replyText(event, '📊 尚無簽到紀錄');
    }
    
    let attend = 0, late = 0, absent = 0;
    studentRecords.forEach(r => {
        const status = r.get('狀態');
        if (status === '已報到') attend++;
        else if (status === '遲到') late++;
        else if (status === '缺席') absent++;
    });
    
    const total = attend + late + absent;
    const rate = total > 0 ? Math.round((attend + late) / total * 100) : 0;
    
    let msg = '📊 出缺席統計\n━━━━━━━━━━━━━━━\n';
    msg += '👤 ' + student.get('姓名') + '\n\n';
    msg += '✅ 出席：' + attend + ' 次\n';
    msg += '⚠️ 遲到：' + late + ' 次\n';
    msg += '❌ 缺席：' + absent + ' 次\n';
    msg += '📈 出席率：' + rate + '%';
    
    return replyText(event, msg);
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
            // 只處理「進行中」的活動
            if (session.get('狀態') !== '進行中') continue;
            
            // 檢查是否已結束
            const endTimeStr = session.get('結束時間');
            if (!endTimeStr) continue;
            const [endHour, endMin] = endTimeStr.split(':').map(Number);
            const endTime = new Date();
            endTime.setHours(endHour, endMin, 0, 0);
            
            if (now > endTime) {
                console.log('📝 處理結束的活動:', session.get('活動ID'));
                
                // 先更新活動狀態為「處理中」避免重複處理
                session.set('狀態', '處理中');
                await session.save();
                
                // 標記缺席的學生
                const courseSheet = doc.sheetsByTitle['課程列表'];
                const courses = await courseSheet.getRows();
                const course = courses.find(c => c.get('課程ID') === session.get('課程ID'));
                
                if (course) {
                    const className = course.get('班級');
                    const students = await studentSheet.getRows();
                    const classStudents = students.filter(s => (s.get('班級') || '').split(/[,、]/).map(c => c.trim()).includes(className));
                    
                    const records = recordSheet ? await recordSheet.getRows() : [];
                    
                    for (const student of classStudents) {
                        const hasRecord = records.some(r => 
                            r.get('活動ID') === session.get('活動ID') &&
                            r.get('學號') === student.get('學號')
                        );
                        
                        if (!hasRecord) {
                            // 記錄缺席（只會記錄一次）
                            const result = await recordAttendance(
                                session.get('活動ID'),
                                student.get('學號'),
                                '缺席'
                            );
                            
                            // 只有成功記錄才發送通知（確保只發一次）
                            if (result.success && student.get('LINE_ID')) {
                                try {
                                    await lineClient.pushMessage(student.get('LINE_ID'), {
                                        type: 'text',
                                        text: `❌ 缺席通知\n\n您已被標記為缺席：\n📚 課程：${course.get('科目')}\n📅 日期：${session.get('日期')}\n\n如有疑問請聯繫教師。`
                                    });
                                    console.log('✉️ 已發送缺席通知給', student.get('學號'));
                                } catch (e) {
                                    console.error('發送通知失敗:', e.message);
                                }
                            }
                        }
                    }
                }
                
                // 更新活動狀態為「已結束」
                session.set('狀態', '已結束');
                await session.save();
                console.log('✅ 活動已結束:', session.get('活動ID'));
            }
        }
        
        console.log('✅ 缺席檢查完成');
    } catch (error) {
        console.error('缺席檢查錯誤:', error);
    }
}

// 每 10 分鐘檢查一次（減少干擾）
cron.schedule('*/10 * * * *', checkAbsences);

// ===== 學期結束通知 =====
async function checkSemesterEnd() {
    console.log('📅 檢查學期結束...');
    
    try {
        const settingsSheet = doc.sheetsByTitle['系統設定'];
        if (!settingsSheet) return;
        
        const settings = await settingsSheet.getRows();
        let semesterEnd = '';
        for (const s of settings) {
            if (s.get('設定項目') === '結業日期') {
                semesterEnd = s.get('設定值');
                break;
            }
        }
        
        if (!semesterEnd) return;
        
        const now = new Date();
        const endDate = new Date(semesterEnd);
        const today = getTodayString();
        
        // 檢查是否是學期最後一天
        if (today !== semesterEnd) return;
        
        // 檢查是否已經發送過通知
        const reminderSheet = await getOrCreateSheet('提醒紀錄', ['課程ID', '日期', '類型', '發送時間']);
        const reminders = await reminderSheet.getRows();
        const alreadySent = reminders.some(r => 
            r.get('日期') === today && 
            r.get('類型') === '學期結束'
        );
        
        if (alreadySent) return;
        
        // 取得最後一堂課的結束時間
        const courseSheet = doc.sheetsByTitle['課程列表'];
        const sessionSheet = doc.sheetsByTitle['簽到活動'];
        
        if (!courseSheet || !sessionSheet) return;
        
        const sessions = await sessionSheet.getRows();
        const todaySessions = sessions.filter(s => s.get('日期') === today);
        
        if (todaySessions.length === 0) return;
        
        // 找最後結束的課程
        let lastEndTime = 0;
        for (const session of todaySessions) {
            const endTimeStr = session.get('結束時間');
            if (endTimeStr) {
                const [h, m] = endTimeStr.split(':').map(Number);
                const endMinutes = h * 60 + m;
                if (endMinutes > lastEndTime) {
                    lastEndTime = endMinutes;
                }
            }
        }
        
        // 檢查現在是否在最後一堂課結束後 30 分鐘
        const currentMinutes = now.getHours() * 60 + now.getMinutes();
        if (currentMinutes >= lastEndTime + 30 && currentMinutes <= lastEndTime + 40) {
            console.log('📢 發送學期結束通知...');
            
            // 發送解除綁定說明給所有學生
            const studentSheet = doc.sheetsByTitle['學生名單'];
            if (studentSheet) {
                const students = await studentSheet.getRows();
                
                for (const student of students) {
                    if (student.get('LINE_ID')) {
                        try {
                            await lineClient.pushMessage(student.get('LINE_ID'), {
                                type: 'text',
                                text: `📚 學期結束通知\n\n親愛的 ${student.get('姓名')} 同學：\n\n本學期課程已全部結束，感謝您這學期的配合！\n\n📌 解除 LINE BOT 綁定方式：\n1. 進入此聊天室\n2. 點右上角「≡」選單\n3. 選擇「封鎖」即可解除\n\n或輸入「解除綁定」由系統處理。\n\n🎉 祝您假期愉快！`
                            });
                        } catch (e) {
                            console.error('發送學期結束通知失敗:', e.message);
                        }
                    }
                }
                
                // 記錄已發送
                await reminderSheet.addRow({
                    '課程ID': 'SEMESTER_END',
                    '日期': today,
                    '類型': '學期結束',
                    '發送時間': formatDateTime(now)
                });
                
                console.log('✅ 學期結束通知已發送');
            }
        }
    } catch (error) {
        console.error('學期結束通知錯誤:', error);
    }
}

// 每 10 分鐘檢查一次學期結束
cron.schedule('*/10 * * * *', checkSemesterEnd);

// ===== 自動上課提醒排程 =====
async function autoClassReminder() {
    console.log('⏰ 檢查上課提醒...');
    
    try {
        // 取得學期設定
        const settingsSheet = doc.sheetsByTitle['系統設定'];
        let remindMinutes = 30; // 預設提前 30 分鐘提醒
        let autoRemind = true;
        
        if (settingsSheet) {
            const settings = await settingsSheet.getRows();
            for (const s of settings) {
                if (s.get('設定項目') === '上課提醒') autoRemind = s.get('設定值') === 'true';
                if (s.get('設定項目') === '提醒分鐘') remindMinutes = parseInt(s.get('設定值')) || 30;
            }
        }
        
        if (!autoRemind) {
            console.log('自動提醒已關閉');
            return;
        }
        
        // 取得今天星期幾
        const now = new Date();
        const dayOfWeek = now.getDay(); // 0=日, 1=一, ... 6=六
        const currentHour = now.getHours();
        const currentMin = now.getMinutes();
        const currentTotalMin = currentHour * 60 + currentMin;
        
        // 取得今天的課程
        const courseSheet = doc.sheetsByTitle['課程列表'];
        if (!courseSheet) return;
        
        const courses = await courseSheet.getRows();
        const todayCourses = courses.filter(c => parseInt(c.get('星期')) === dayOfWeek && c.get('狀態') === '啟用');
        
        if (todayCourses.length === 0) {
            console.log('今天沒有課程');
            return;
        }
        
        // 取得已發送的提醒記錄（避免重複發送）
        const reminderSheet = await getOrCreateSheet('提醒紀錄', ['課程ID', '日期', '類型', '發送時間']);
        const reminders = await reminderSheet.getRows();
        const today = getTodayString();
        
        for (const course of todayCourses) {
            const courseId = course.get('課程ID');
            const courseTime = course.get('上課時間') || '';
            const [startTime] = courseTime.split('-');
            
            if (!startTime) continue;
            
            const [startHour, startMin] = startTime.split(':').map(Number);
            const startTotalMin = startHour * 60 + startMin;
            const reminderTime = startTotalMin - remindMinutes;
            
            // 檢查是否到了提醒時間（允許 5 分鐘誤差）
            if (currentTotalMin >= reminderTime && currentTotalMin <= reminderTime + 5) {
                // 檢查今天是否已發送過提醒
                const alreadySent = reminders.some(r => 
                    r.get('課程ID') === courseId && 
                    r.get('日期') === today && 
                    r.get('類型') === '上課提醒'
                );
                
                if (alreadySent) {
                    console.log(`課程 ${courseId} 今日已發送提醒`);
                    continue;
                }
                
                console.log(`📢 發送上課提醒: ${course.get('科目')}`);
                
                // 自動建立簽到活動
                const sessionSheet = await getOrCreateSheet('簽到活動', [
                    '活動ID', '課程ID', '日期', '開始時間', '結束時間', 'QR碼內容', '狀態'
                ]);
                
                const sessionId = `S${Date.now()}`;
                // 老師手機 QR Code 用「直接簽到」，學生連結用「GPS簽到」
                const qrContent = `直接簽到:${courseId}|${sessionId}`;
                const gpsCheckinCode = `GPS簽到:${courseId}|${sessionId}`;
                const [, endTime] = courseTime.split('-');
                
                await sessionSheet.addRow({
                    '活動ID': sessionId,
                    '課程ID': courseId,
                    '日期': today,
                    '開始時間': startTime,
                    '結束時間': endTime || '',
                    'QR碼內容': qrContent,
                    '狀態': '進行中'
                });
                
                // 發送 LINE 通知給學生（使用 GPS 簽到連結）
                const classCode = course.get('班級');
                const studentSheet = doc.sheetsByTitle['學生名單'];
                if (studentSheet) {
                    const students = await studentSheet.getRows();
                    const classStudents = students.filter(s => (s.get('班級') || '').split(/[,、]/).map(c => c.trim()).includes(classCode) && s.get('LINE_ID'));
                    
                    const botId = process.env.LINE_BOT_ID;
                    // 學生連結使用 GPS 簽到
                    const checkinUrl = `https://line.me/R/oaMessage/${botId}/?${encodeURIComponent(gpsCheckinCode)}`;
                    
                    for (const student of classStudents) {
                        try {
                            await lineClient.pushMessage(student.get('LINE_ID'), {
                                type: 'template',
                                altText: `📢 上課提醒 - ${course.get('科目')}`,
                                template: {
                                    type: 'buttons',
                                    title: `📢 ${course.get('科目')} 即將上課`,
                                    text: `⏰ ${courseTime}\n📍 ${course.get('教室') || '教室'}\n\n${remindMinutes} 分鐘後上課`,
                                    actions: [
                                        {
                                            type: 'uri',
                                            label: '📱 點我簽到',
                                            uri: checkinUrl
                                        }
                                    ]
                                }
                            });
                        } catch (e) {
                            console.error(`發送提醒失敗 ${student.get('學號')}:`, e.message);
                        }
                    }
                    
                    console.log(`✅ 已發送 ${classStudents.length} 則提醒`);
                }
                
                // 記錄已發送
                await reminderSheet.addRow({
                    '課程ID': courseId,
                    '日期': today,
                    '類型': '上課提醒',
                    '發送時間': now.toLocaleString('zh-TW')
                });
            }
        }
        
        console.log('✅ 上課提醒檢查完成');
    } catch (error) {
        console.error('上課提醒錯誤:', error);
    }
}

// 每分鐘檢查一次（確保不會錯過提醒時間）
cron.schedule('* * * * *', autoClassReminder);


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
        const sheet = await getOrCreateSheet('班級列表', ['班級代碼', '班級名稱', '部別', '導師', '人數', '建立時間']);
        const rows = await sheet.getRows();
        
        // 取得學生名單來計算人數
        let studentCounts = {};
        try {
            const studentSheet = doc.sheetsByTitle['學生名單'];
            if (studentSheet) {
                const students = await studentSheet.getRows();
                students.forEach(s => {
                    const classCode = s.get('班級');
                    if (classCode) {
                        studentCounts[classCode] = (studentCounts[classCode] || 0) + 1;
                    }
                });
            }
        } catch (e) {
            console.log('計算學生人數失敗:', e.message);
        }
        
        res.json(rows.map(r => ({
            code: r.get('班級代碼'),
            name: r.get('班級名稱'),
            division: r.get('部別') || 'day',
            teacher: r.get('導師'),
            count: studentCounts[r.get('班級代碼')] || 0
        })));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/classes', async (req, res) => {
    try {
        const { code, name, division, teacher } = req.body;
        const sheet = await getOrCreateSheet('班級列表', ['班級代碼', '班級名稱', '部別', '導師', '人數', '建立時間']);
        await sheet.addRow({
            '班級代碼': code,
            '班級名稱': name,
            '部別': division || 'day',
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
        const { name, division, teacher } = req.body;
        const sheet = doc.sheetsByTitle['班級列表'];
        if (!sheet) return res.json({ success: false, message: '資料表不存在' });
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('班級代碼') === code);
        if (!row) return res.json({ success: false, message: '班級不存在' });
        
        if (name) row.set('班級名稱', name);
        if (division) row.set('部別', division);
        if (teacher !== undefined) row.set('導師', teacher);
        await row.save();
        
        res.json({ success: true });
    } catch (error) {
        console.error('更新班級錯誤:', error);
        res.status(500).json({ success: false, error: error.message });
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
        const students = rows.filter(r => (r.get('班級') || '').split(/[,、]/).map(c => c.trim()).includes(code));
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
            '課程ID', '科目', '班級', '教師', '星期', '節次', '上課時間', '教室',
            '教室緯度', '教室經度', '簽到範圍', '遲到標準', '狀態', '建立時間'
        ]);
        const rows = await sheet.getRows();
        res.json(rows.map(r => ({
            id: r.get('課程ID'),
            subject: r.get('科目'),
            name: r.get('科目'),
            classCode: r.get('班級'),
            teacher: r.get('教師'),
            day: parseInt(r.get('星期')) || 1,
            period: parseInt(r.get('節次')) || 1,
            time: r.get('上課時間'),
            room: r.get('教室'),
            lat: parseFloat(r.get('教室緯度')) || 0,
            lon: parseFloat(r.get('教室經度')) || 0,
            radius: r.get('簽到範圍') !== '' && r.get('簽到範圍') !== undefined ? parseInt(r.get('簽到範圍')) : 100,
            lateMinutes: parseInt(r.get('遲到標準')) || 10,
            status: r.get('狀態') || '啟用'
        })));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/courses', async (req, res) => {
    try {
        const { subject, name, classCode, teacher, day, period, time, room, lat, lon, radius, lateMinutes } = req.body;
        const sheet = await getOrCreateSheet('課程列表', [
            '課程ID', '科目', '班級', '教師', '星期', '節次', '上課時間', '教室',
            '教室緯度', '教室經度', '簽到範圍', '遲到標準', '狀態', '建立時間'
        ]);
        const courseId = 'C' + Date.now();
        await sheet.addRow({
            '課程ID': courseId,
            '科目': subject || name,
            '班級': classCode,
            '教師': teacher || '',
            '星期': day || 1,
            '節次': period || 1,
            '上課時間': time || '',
            '教室': room || '',
            '教室緯度': lat || 0,
            '教室經度': lon || 0,
            '簽到範圍': radius !== undefined ? radius : 100,
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

// 更新課程
app.put('/api/courses/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const { subject, classCode, day, period, time, room, lat, lon, radius } = req.body;
        console.log('📝 更新課程請求:', id, { radius, radiusType: typeof radius });
        
        // 強制刷新
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle['課程列表'];
        if (!sheet) return res.json({ success: false, message: '資料表不存在' });
        
        await sheet.loadHeaderRow();
        const rows = await sheet.getRows({ limit: 500 });
        const row = rows.find(r => r.get('課程ID') === id);
        if (!row) return res.json({ success: false, message: '課程不存在' });
        
        // 記錄更新前的值
        const oldRadius = row.get('簽到範圍');
        console.log('📝 更新前簽到範圍:', oldRadius);
        
        if (subject) row.set('科目', subject);
        if (classCode) row.set('班級', classCode);
        if (day !== undefined) row.set('星期', day);
        if (period !== undefined) row.set('節次', period);
        if (time) row.set('上課時間', time);
        if (room !== undefined) row.set('教室', room);
        if (lat !== undefined) row.set('教室緯度', lat);
        if (lon !== undefined) row.set('教室經度', lon);
        if (radius !== undefined) {
            // 確保存入數字
            row.set('簽到範圍', parseInt(radius));
        }
        
        await row.save();
        
        // 驗證：重新讀取確認更新成功
        await doc.loadInfo();
        const verifySheet = doc.sheetsByTitle['課程列表'];
        await verifySheet.loadHeaderRow();
        const verifyRows = await verifySheet.getRows({ limit: 500 });
        const verifyRow = verifyRows.find(r => r.get('課程ID') === id);
        const newRadius = verifyRow ? verifyRow.get('簽到範圍') : '找不到';
        
        console.log('✅ 更新後簽到範圍:', newRadius, '(預期:', radius, ')');
        
        res.json({ success: true, radius: newRadius, oldRadius, requestedRadius: radius });
    } catch (error) {
        console.error('更新課程錯誤:', error);
        res.status(500).json({ success: false, error: error.message });
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
        // 老師手機 QR Code 用「直接簽到」
        const qrContent = `直接簽到:${courseId}|${sessionId}`;
        // 學生連結用「GPS簽到」
        const gpsCheckinCode = `GPS簽到:${courseId}|${sessionId}`;
        await sheet.addRow({
            '活動ID': sessionId,
            '課程ID': courseId,
            '日期': date,
            '開始時間': startTime,
            '結束時間': endTime,
            'QR碼內容': qrContent,
            '狀態': '進行中'
        });
        res.json({ success: true, sessionId, qrContent, gpsCheckinCode });
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
// 發送上課提醒（附帶簽到連結）
// === 通知 API ===
// 發送上課提醒（附帶簽到連結，含防重複機制）
app.post('/api/notify/remind', async (req, res) => {
    try {
        const { courseId, sessionId, message } = req.body;
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
        const classStudents = students.filter(s => 
            (s.get('班級') || '')
                .split(/[,、]/)
                .map(c => c.trim())
                .includes(classCode) && 
            s.get('LINE_ID')
        );
        
        // 讀取提醒紀錄，避免同一節課重複發送手動提醒
        const reminderSheet = await getOrCreateSheet('提醒紀錄', [
            '課程ID', '日期', '類型', '發送時間', '活動ID'
        ]);
        const reminders = await reminderSheet.getRows();
        const today = getTodayString();
        
        if (sessionId) {
            const alreadySent = reminders.some(r =>
                r.get('課程ID') === courseId &&
                r.get('日期') === today &&
                (r.get('類型') === '手動上課提醒') &&
                (r.get('活動ID') || '') === String(sessionId)
            );
            
            if (alreadySent) {
                console.log(`⛔ 手動提醒略過：課程 ${courseId} 活動 ${sessionId} 今日已發送過`);
                return res.json({ 
                    success: false, 
                    message: '今日此節課已發送過簽到提醒，不再重複推播。' 
                });
            }
        }
        
        // 建立簽到連結（學生使用 GPS 簽到）
        const botId = process.env.LINE_BOT_ID;
        const checkinCode = sessionId ? `GPS簽到:${courseId}|${sessionId}` : '';
        const checkinUrl = checkinCode 
            ? `https://line.me/R/oaMessage/${botId}/?${encodeURIComponent(checkinCode)}` 
            : '';
        
        let sent = 0;
        
        // 發送 LINE 通知
        for (const student of classStudents) {
            const lineId = student.get('LINE_ID');
            if (!lineId) continue;
            
            try {
                if (checkinUrl) {
                    // 帶簽到按鈕
                    await lineClient.pushMessage(lineId, {
                        type: 'template',
                        altText: `📢 上課提醒 - ${course.get('科目')}`,
                        template: {
                            type: 'buttons',
                            title: `📢 ${course.get('科目')} 上課提醒`,
                            text: `⏰ ${course.get('上課時間')}
📍 ${course.get('教室') || '教室'}

請點擊下方按鈕簽到`,
                            actions: [
                                {
                                    type: 'uri',
                                    label: '📱 點我簽到',
                                    uri: checkinUrl
                                }
                            ]
                        }
                    });
                } else {
                    // 純文字提醒
                    await lineClient.pushMessage(lineId, {
                        type: 'text',
                        text: message || 
                            `📢 上課提醒

${course.get('科目')} 即將開始！
⏰ ${course.get('上課時間')}
📍 ${course.get('教室') || '教室'}`
                    });
                }
                sent++;
            } catch (e) {
                console.error(`發送提醒失敗 ${student.get('學號')}:`, e.message);
            }
        }
        
        // 寫入提醒紀錄
        await reminderSheet.addRow({
            '課程ID': courseId,
            '日期': today,
            '類型': '手動上課提醒',
            '發送時間': formatDateTime(new Date()),
            '活動ID': sessionId || ''
        });
        
        console.log(`✅ 手動上課提醒已發送：課程 ${courseId}，人數 ${sent}`);
        return res.json({ success: true, sent });
    } catch (error) {
        console.error('手動上課提醒錯誤:', error);
        return res.status(500).json({ success: false, message: error.message });
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
// 取得設定（通用）
app.get('/api/settings', async (req, res) => {
    try {
        const sheet = await getOrCreateSheet('系統設定', ['設定項目', '設定值']);
        const rows = await sheet.getRows();
        const settings = {};
        rows.forEach(r => {
            settings[r.get('設定項目')] = r.get('設定值');
        });
        res.json({
            remindBeforeClass: settings['上課提醒'] !== 'false',
            remindMinutes: parseInt(settings['提醒分鐘']) || 30,
            notifyAbsent: settings['缺席通知'] === 'true',
            notifyParent: settings['通知家長'] === 'true',
            warningThreshold: parseInt(settings['警告門檻']) || 3,
            weeklyReport: settings['週報'] === 'true',
            semesterStart: settings['開學日期'] || '',
            semesterEnd: settings['結業日期'] || ''
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 儲存設定（通用）
app.post('/api/settings', async (req, res) => {
    try {
        const { remindBeforeClass, remindMinutes, notifyAbsent, notifyParent, warningThreshold, weeklyReport, semesterStart, semesterEnd } = req.body;
        const sheet = await getOrCreateSheet('系統設定', ['設定項目', '設定值']);
        
        // 更新或新增設定
        const rows = await sheet.getRows();
        const settingsMap = {};
        rows.forEach(r => { settingsMap[r.get('設定項目')] = r; });
        
        const updateOrAdd = async (key, value) => {
            if (settingsMap[key]) {
                settingsMap[key].set('設定值', value);
                await settingsMap[key].save();
            } else {
                await sheet.addRow({ '設定項目': key, '設定值': value });
            }
        };
        
        if (remindBeforeClass !== undefined) await updateOrAdd('上課提醒', remindBeforeClass ? 'true' : 'false');
        if (remindMinutes !== undefined) await updateOrAdd('提醒分鐘', remindMinutes);
        if (notifyAbsent !== undefined) await updateOrAdd('缺席通知', notifyAbsent ? 'true' : 'false');
        if (notifyParent !== undefined) await updateOrAdd('通知家長', notifyParent ? 'true' : 'false');
        if (warningThreshold !== undefined) await updateOrAdd('警告門檻', warningThreshold);
        if (weeklyReport !== undefined) await updateOrAdd('週報', weeklyReport ? 'true' : 'false');
        if (semesterStart !== undefined) await updateOrAdd('開學日期', semesterStart);
        if (semesterEnd !== undefined) await updateOrAdd('結業日期', semesterEnd);
        
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

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
        const classStudents = students.filter(s => (s.get('班級') || '').split(/[,、]/).map(c => c.trim()).includes(classCode));
        
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
            const classStudents = students.filter(s => (s.get('班級') || '').split(/[,、]/).map(c => c.trim()).includes(code)).map(s => s.get('學號'));
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
            const classStudentIds = students.filter(s => (s.get('班級') || '').split(/[,、]/).map(c => c.trim()).includes(classCode)).map(s => s.get('學號'));
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
        const { studentId, name, classCode, lineId, lineName, phone, parentPhone, parentLineId } = req.body;
        const sheet = await getOrCreateSheet('學生名單', ['學號', '姓名', '班級', 'LINE_ID', 'LINE名稱', '電話', '家長電話', '家長LINE_ID', '註冊時間']);
        
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
            'LINE_ID': lineId || '',
            'LINE名稱': lineName || '',
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
        const { name, classCode, lineId, lineName, phone, parentPhone } = req.body;
        const sheet = doc.sheetsByTitle['學生名單'];
        if (!sheet) return res.json({ success: false });
        
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('學號') === id);
        if (!row) return res.json({ success: false, message: '學生不存在' });
        
        if (name) row.set('姓名', name);
        if (classCode) row.set('班級', classCode);
        if (lineId !== undefined) row.set('LINE_ID', lineId);
        if (lineName !== undefined) row.set('LINE名稱', lineName);
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
        const classStudents = students.filter(s => (s.get('班級') || '').split(/[,、]/).map(c => c.trim()).includes(classCode) && s.get('LINE_ID'));
        
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

// ===== 測試驗證 API =====

// 檢查學生 LINE 綁定狀態（除錯用）
app.get('/api/debug/students', async (req, res) => {
    try {
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle['學生名單'];
        if (!sheet) return res.json({ error: '學生名單不存在' });
        
        await sheet.loadHeaderRow();
        const rows = await sheet.getRows({ limit: 100 });
        
        const students = rows.map(r => ({
            學號: r.get('學號'),
            姓名: r.get('姓名'),
            班級: r.get('班級'),
            LINE_ID: r.get('LINE_ID') ? (r.get('LINE_ID').substring(0, 15) + '...') : '未綁定',
            LINE_ID長度: (r.get('LINE_ID') || '').length,
            已綁定: !!r.get('LINE_ID')
        }));
        
        res.json({
            欄位名稱: sheet.headerValues,
            總學生數: rows.length,
            已綁定數: students.filter(s => s.已綁定).length,
            學生列表: students
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 檢查特定 LINE_ID 是否存在（除錯用）
app.get('/api/debug/check-lineid/:lineId', async (req, res) => {
    try {
        const { lineId } = req.params;
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle['學生名單'];
        if (!sheet) return res.json({ error: '學生名單不存在' });
        
        await sheet.loadHeaderRow();
        const rows = await sheet.getRows({ limit: 1000 });
        
        // 精確比對
        const exactMatch = rows.find(r => r.get('LINE_ID') === lineId);
        
        // trim 後比對
        const trimMatch = rows.find(r => (r.get('LINE_ID') || '').trim() === lineId);
        
        // 部分比對（前 20 字元）
        const partialMatches = rows.filter(r => {
            const storedId = r.get('LINE_ID') || '';
            return storedId.includes(lineId.substring(0, 20)) || lineId.includes(storedId.substring(0, 20));
        });
        
        res.json({
            查詢的LINE_ID: lineId,
            LINE_ID長度: lineId.length,
            精確比對: exactMatch ? { 姓名: exactMatch.get('姓名'), 學號: exactMatch.get('學號') } : null,
            trim比對: trimMatch ? { 姓名: trimMatch.get('姓名'), 學號: trimMatch.get('學號') } : null,
            部分比對: partialMatches.map(r => ({
                姓名: r.get('姓名'),
                學號: r.get('學號'),
                LINE_ID: r.get('LINE_ID')
            })),
            欄位名稱: sheet.headerValues
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 檢查課程簽到範圍設定（除錯用）
app.get('/api/debug/course/:id', async (req, res) => {
    try {
        const { id } = req.params;
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle['課程列表'];
        if (!sheet) return res.json({ error: '課程列表不存在' });
        
        await sheet.loadHeaderRow();
        const rows = await sheet.getRows({ limit: 500 });
        const row = rows.find(r => r.get('課程ID') === id);
        
        if (!row) return res.json({ error: '課程不存在', courseId: id });
        
        const rawRadius = row.get('簽到範圍');
        res.json({
            courseId: id,
            科目: row.get('科目'),
            班級: row.get('班級'),
            教室緯度: row.get('教室緯度'),
            教室經度: row.get('教室經度'),
            簽到範圍_原始值: rawRadius,
            簽到範圍_類型: typeof rawRadius,
            簽到範圍_解析: parseInt(rawRadius),
            所有欄位: sheet.headerValues
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 測試學期結束通知
app.post('/api/test/semester-end', async (req, res) => {
    try {
        const { classCode } = req.body;
        if (!classCode) return res.json({ success: false, message: '請選擇班級' });
        
        const studentSheet = doc.sheetsByTitle['學生名單'];
        if (!studentSheet) return res.json({ success: false, message: '學生名單不存在' });
        
        const students = await studentSheet.getRows();
        const classStudents = students.filter(s => (s.get('班級') || '').split(/[,、]/).map(c => c.trim()).includes(classCode) && s.get('LINE_ID'));
        
        let count = 0;
        for (const student of classStudents) {
            try {
                await lineClient.pushMessage(student.get('LINE_ID'), {
                    type: 'text',
                    text: `📚 【測試】學期結束通知\n\n親愛的 ${student.get('姓名')} 同學：\n\n本學期課程已全部結束，感謝您這學期的配合！\n\n📌 解除 LINE BOT 綁定方式：\n1. 進入此聊天室\n2. 點右上角「≡」選單\n3. 選擇「封鎖」即可解除\n\n或輸入「解除綁定」由系統處理。\n\n🎉 祝您假期愉快！\n\n⚠️ 這是測試訊息`
                });
                count++;
            } catch (e) {
                console.error('發送測試通知失敗:', e.message);
            }
        }
        
        res.json({ success: true, count, message: `已發送 ${count} 則通知` });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

// 測試簽到狀態通知（準時/遲到/缺席）
app.post('/api/test/checkin-notify', async (req, res) => {
    try {
        const { classCode, status, lateMinutes } = req.body;
        if (!classCode) return res.json({ success: false, message: '請選擇班級' });
        
        const studentSheet = doc.sheetsByTitle['學生名單'];
        if (!studentSheet) return res.json({ success: false, message: '學生名單不存在' });
        
        const students = await studentSheet.getRows();
        const classStudents = students.filter(s => (s.get('班級') || '').split(/[,、]/).map(c => c.trim()).includes(classCode) && s.get('LINE_ID'));
        
        const today = getTodayString();
        let notifyText = '';
        
        if (status === '已報到') {
            notifyText = `✅ 【測試】簽到成功\n\n📚 課程：測試課程\n📅 日期：${today}\n✨ 狀態：準時報到\n\n繼續保持！💪\n\n⚠️ 這是測試訊息`;
        } else if (status === '遲到') {
            notifyText = `⚠️ 【測試】遲到通知\n\n📚 課程：測試課程\n📅 日期：${today}\n⏰ 遲到：${lateMinutes || 15} 分鐘\n\n請下次準時出席！\n\n⚠️ 這是測試訊息`;
        } else if (status === '缺席') {
            notifyText = `❌ 【測試】缺席通知\n\n📚 課程：測試課程\n📅 日期：${today}\n\n如有疑問請聯繫教師。\n\n⚠️ 這是測試訊息`;
        }
        
        let count = 0;
        for (const student of classStudents) {
            try {
                await lineClient.pushMessage(student.get('LINE_ID'), {
                    type: 'text',
                    text: notifyText
                });
                count++;
            } catch (e) {
                console.error('發送測試通知失敗:', e.message);
            }
        }
        
        res.json({ success: true, count, message: `已發送 ${count} 則通知` });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

// 測試上課提醒
app.post('/api/test/reminder', async (req, res) => {
    try {
        const { classCode } = req.body;
        if (!classCode) return res.json({ success: false, message: '請選擇班級' });
        
        const studentSheet = doc.sheetsByTitle['學生名單'];
        if (!studentSheet) return res.json({ success: false, message: '學生名單不存在' });
        
        const students = await studentSheet.getRows();
        const classStudents = students.filter(s => (s.get('班級') || '').split(/[,、]/).map(c => c.trim()).includes(classCode) && s.get('LINE_ID'));
        
        // 建立測試簽到連結
        const botId = process.env.LINE_BOT_ID;
        const testCode = `GPS簽到:TEST|TEST${Date.now()}`;
        const checkinUrl = `https://line.me/R/oaMessage/${botId}/?${encodeURIComponent(testCode)}`;
        
        let count = 0;
        for (const student of classStudents) {
            try {
                await lineClient.pushMessage(student.get('LINE_ID'), {
                    type: 'template',
                    altText: '📢 【測試】上課提醒',
                    template: {
                        type: 'buttons',
                        title: '📢 【測試】上課提醒',
                        text: `⏰ 08:00-09:00\n📍 測試教室\n\n30 分鐘後上課\n\n⚠️ 這是測試訊息`,
                        actions: [
                            {
                                type: 'uri',
                                label: '📱 點我簽到（測試）',
                                uri: checkinUrl
                            }
                        ]
                    }
                });
                count++;
            } catch (e) {
                console.error('發送測試提醒失敗:', e.message);
            }
        }
        
        res.json({ success: true, count, message: `已發送 ${count} 則提醒` });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
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
