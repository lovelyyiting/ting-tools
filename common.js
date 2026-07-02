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
    ltDays:   ['基本交期','基本交期(天)','交期天數','交期(天)','採購前置時間','前置時間','L/T','LT'],
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

// 下載匯入範本
function downloadTemplate(filename, headers, example) {
    const ws = XLSX.utils.aoa_to_sheet([headers, example]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '匯入資料');
    XLSX.writeFile(wb, filename);
}

// ===== 行事曆工作日 =====
// 預設週一～週五上班、週六日休息；overrides = { '2026-01-01': 'off' | 'on' }
function isWorkday(dateStr, overrides) {
    const ov = (overrides || {})[dateStr];
    if (ov === 'on') return true;
    if (ov === 'off') return false;
    const day = new Date(dateStr + 'T00:00:00').getDay();
    return day >= 1 && day <= 5;
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
