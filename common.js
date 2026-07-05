// 共用工具：Firebase 連線、Excel 解析、行事曆工作日計算
const firebaseConfig = {
    databaseURL: "https://super-ting-default-rtdb.firebaseio.com"
};

let db = null;
try {
    firebase.initializeApp(firebaseConfig);
    db = firebase.database();
} catch(e) {
    console.log("Firebase 初始化失敗，使用本地模式", e);
}

// 右上角同步狀態指示燈
function setSyncStatus(connected) {
    let badge = document.getElementById('syncStatus');
    if (!badge) {
        badge = document.createElement('div');
        badge.id = 'syncStatus';
        badge.style.cssText = 'position:fixed;top:8px;right:8px;z-index:9999;padding:4px 10px;border-radius:12px;font-size:12px;color:white;box-shadow:0 2px 6px rgba(0,0,0,0.2);';
        document.body.appendChild(badge);
    }
    if (connected) {
        badge.textContent = '🟢 雲端同步中';
        badge.style.background = '#28a745';
    } else {
        badge.textContent = '🔴 離線（僅本機）';
        badge.style.background = '#dc3545';
    }
}
if (db) {
    db.ref('.info/connected').on('value', s => setSyncStatus(s.val() === true));
} else {
    setSyncStatus(false);
}

function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function fmtDate(d) {
    return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

function todayStr() { return fmtDate(new Date()); }

// 數字解析：容忍千分位逗號與空白
function num(v) {
    const n = Number(String(v ?? '').replace(/[,\s]/g, ''));
    return isNaN(n) ? 0 : n;
}

// 日期解析：容忍 Excel 日期物件、序號、2026/7/2、2026-07-02 等格式
function parseDateCell(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date && !isNaN(v)) return fmtDate(v);
    if (typeof v === 'number') {
        const d = new Date(Math.round((v - 25569) * 86400 * 1000));
        return d.getUTCFullYear() + '-' + String(d.getUTCMonth()+1).padStart(2,'0') + '-' + String(d.getUTCDate()).padStart(2,'0');
    }
    const s = String(v).trim().replace(/[./]/g, '-');
    const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return m[1] + '-' + m[2].padStart(2,'0') + '-' + m[3].padStart(2,'0');
    return '';
}

// Excel 欄位標題別名：不同系統匯出的標題都能對上
const COL_ALIASES = {
    mat:      ['料號','物料','品號','物料編號','料件編號','Material','material'],
    name:     ['品名','品名規格','名稱','品名/規格','Description'],
    qty:      ['數量','需求數量','在途數量','Qty','QTY','qty'],
    ltDays:   ['基本交期','基本交期(天)','交期天數','交期(天)','計劃交貨時間','計畫交貨時間','計劃交期','交貨時間','採購前置時間','前置時間','L/T','LT'],
    dueDate:  ['交期','交貨日','到貨日','預交日','交期日期','預計到貨日'],
    orderNo:  ['工單號','工單','製令','製令單號','單號'],
    parent:   ['成品料號','主件料號','母件料號','父階料號','父階'],
    child:    ['子件料號','元件料號','子階料號','組件料號','元件','子件'],
    usage:    ['用量','組成用量','單位用量','每單位用量'],
    stock:    ['庫存','庫存數量','現有庫存','未限制庫存'],
    startDate:['開始日期','開始日','開工日','預計開工'],
    owner:    ['負責人','機台','負責人/機台'],
    po:       ['採購單號','採購單','PO','PO單號'],
    note:     ['備註','說明','Remark']
};

// 從第一列資料的標題找出各欄位對應的實際標題
function buildHeaderMap(row) {
    const keys = Object.keys(row);
    const map = {};
    Object.entries(COL_ALIASES).forEach(([field, aliases]) => {
        const k = keys.find(k => aliases.includes(String(k).trim()));
        if (k !== undefined) map[field] = k;
    });
    return map;
}

function cell(row, map, field) {
    return map[field] !== undefined ? row[map[field]] : '';
}

// 讀取 Excel 檔第一個工作表 → 物件陣列（以標題列為 key）
function readExcelFile(file, cb) {
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
            const ws = wb.Sheets[wb.SheetNames[0]];
            cb(null, XLSX.utils.sheet_to_json(ws, { defval: '' }));
        } catch (err) { cb(err); }
    };
    reader.onerror = () => cb(new Error('讀檔失敗'));
    reader.readAsArrayBuffer(file);
}

// 讀取 Excel 第一個工作表 → 陣列的陣列（保留原始儲存格，含無標題資料）
function readExcelFileAoa(file, cb) {
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: false });
            const ws = wb.Sheets[wb.SheetNames[0]];
            cb(null, XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }));
        } catch (err) { cb(err); }
    };
    reader.onerror = () => cb(new Error('讀檔失敗'));
    reader.readAsArrayBuffer(file);
}

// 下載匯入範本
function downloadTemplate(filename, headers, example) {
    const ws = XLSX.utils.aoa_to_sheet([headers, example]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '匯入資料');
    XLSX.writeFile(wb, filename);
}

// ===== 行事曆（小時制）=====
// 新版行事曆 cal = { defaultWeekday: 7.8, days: { '2026-01-01': 小時數 } }
// 舊版相容：overrides = { '2026-01-01': 'off' | 'on' }
const DEFAULT_WEEKDAY_HOURS = 7.8;

// 判斷是否為新版小時制行事曆物件
function isHoursCal(cal) {
    return cal && typeof cal === 'object' && ('days' in cal || 'defaultWeekday' in cal);
}

// 某一天的可用工時（小時）
function dayHours(dateStr, cal) {
    if (isHoursCal(cal)) {
        const days = cal.days || {};
        if (days[dateStr] !== undefined && days[dateStr] !== null && days[dateStr] !== '') return Number(days[dateStr]);
        const dw = new Date(dateStr + 'T00:00:00').getDay();
        if (dw === 0 || dw === 6) return 0; // 預設週六日休
        return Number(cal.defaultWeekday != null ? cal.defaultWeekday : DEFAULT_WEEKDAY_HOURS);
    }
    // 舊版 on/off：上班日給預設工時、休息日 0
    return isWorkday(dateStr, cal) ? DEFAULT_WEEKDAY_HOURS : 0;
}

// 預設週一～週五上班、週六日休息；支援新版(小時>0視為上班)與舊版 on/off
function isWorkday(dateStr, cal) {
    if (isHoursCal(cal)) return dayHours(dateStr, cal) > 0;
    const ov = (cal || {})[dateStr];
    if (ov === 'on') return true;
    if (ov === 'off') return false;
    const day = new Date(dateStr + 'T00:00:00').getDay();
    return day >= 1 && day <= 5;
}

// 產能游標消耗：從 cursor { date, remain(當天剩餘工時) } 起，消耗 needHours 小時
// 回傳 { start, end, cursor }；start=實際開工日、end=完工日、cursor=消耗後的新游標
// 用於「一條產線依序佔用產能」，可跨日、可從當天剩餘時數接續
function consumeHours(cursor, needHours, cal) {
    let { date, remain } = cursor;
    let need = Number(needHours) || 0;
    let guard = 0;
    // 先移動到有可用工時的日期
    const ensureCapacity = () => {
        let g = 0;
        while (remain <= 0 && g++ < 3700) {
            const d = new Date(date + 'T00:00:00'); d.setDate(d.getDate() + 1); date = fmtDate(d);
            remain = dayHours(date, cal);
        }
    };
    ensureCapacity();
    if (need <= 0) return { start: date, end: date, cursor: { date, remain } };
    const start = date;
    let last = date;
    while (need > 1e-9 && guard++ < 5000) {
        if (remain <= 0) { ensureCapacity(); }
        const use = Math.min(remain, need);
        remain -= use; need -= use; last = date;
        if (need > 1e-9 && remain <= 1e-9) ensureCapacity();
    }
    return { start, end: last, cursor: { date, remain } };
}

// 從 fromStr（含當天）起，依行事曆逐日消耗可用工時，累計滿 needHours 的那一天即結束日
// 回傳 { start, end }：start=第一個有工時可用的日期，end=工時累計達標的日期
function scheduleByHours(fromStr, needHours, cal) {
    const d = new Date(fromStr + 'T00:00:00');
    let remain = Number(needHours) || 0, guard = 0, start = null, last = fmtDate(d);
    if (remain <= 0) return { start: fmtDate(d), end: fmtDate(d), spanDays: 0 };
    while (remain > 0 && guard++ < 3700) {
        const ds = fmtDate(d);
        const h = dayHours(ds, cal);
        if (h > 0) {
            if (!start) start = ds;
            remain -= h;
            last = ds;
        }
        if (remain > 0) d.setDate(d.getDate() + 1);
    }
    return { start: start || fmtDate(d), end: last };
}

// 從 fromStr 起算 n 個日曆天後的日期（單純加天數，不跳假日）
function addDays(fromStr, n) {
    const d = new Date(fromStr + 'T00:00:00');
    d.setDate(d.getDate() + Number(n || 0));
    return fmtDate(d);
}

// 從 fromStr 起算 n 個工作天後的日期（跳過休假日）
function addWorkdays(fromStr, n, overrides) {
    const d = new Date(fromStr + 'T00:00:00');
    let count = 0, guard = 0;
    while (count < n && guard++ < 3700) {
        d.setDate(d.getDate() + 1);
        if (isWorkday(fmtDate(d), overrides)) count++;
    }
    return fmtDate(d);
}
