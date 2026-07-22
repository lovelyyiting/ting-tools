// 共用密碼門檻（簡易保護）：輸入正確密碼才能進站，記住 30 天。
// 說明：這是「輕量遮蔽」，可擋掉知道網址就想直接用的人；但資料庫本身仍為開放狀態，
//       真正要鎖死請改用「登入＋Firebase 權限規則」。
// 換密碼：把下面 PASS_HASH 換成新密碼的 SHA-256 雜湊（不要放明碼）。
(function () {
    const PASS_HASH = '0b26478de84be38e49aa45505ccb74f40477d9bcd1540e16dab78fb4eebc393e'; // 密碼 = rting（換密碼改此雜湊）
    const KEY = 'ting_gate_ok';
    const DAYS = 30;

    // 已解鎖且未過期 → 直接放行
    try {
        const v = JSON.parse(localStorage.getItem(KEY) || 'null');
        if (v && v.h === PASS_HASH && v.exp > Date.now()) return;
    } catch (e) {}

    // 先把畫面藏起來，避免內容閃一下
    const hideStyle = document.createElement('style');
    hideStyle.id = 'gate-hide';
    hideStyle.textContent = 'body{visibility:hidden !important;}';
    (document.head || document.documentElement).appendChild(hideStyle);

    async function sha256(str) {
        const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
        return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
    }

    function buildOverlay() {
        const ov = document.createElement('div');
        ov.id = 'gate-overlay';
        ov.style.cssText = 'visibility:visible;position:fixed;inset:0;z-index:2147483647;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#8c9a76,#6e7a58);font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,sans-serif;';
        ov.innerHTML =
            '<div style="background:#fff;padding:28px 24px;border-radius:14px;box-shadow:0 10px 30px rgba(0,0,0,.25);width:280px;text-align:center;">' +
            '<div style="font-size:34px;">🔒</div>' +
            '<div style="font-size:16px;font-weight:700;margin:8px 0 4px;">請輸入密碼</div>' +
            '<div style="font-size:12px;color:#888;margin-bottom:14px;">SUPER TING 工作工具</div>' +
            '<input id="gate-pw" type="password" autocomplete="off" style="width:100%;padding:10px;border:1px solid #ccc;border-radius:8px;font-size:15px;box-sizing:border-box;" placeholder="密碼">' +
            '<div id="gate-err" style="color:#dc3545;font-size:12px;height:16px;margin:6px 0;"></div>' +
            '<button id="gate-btn" style="width:100%;padding:10px;border:0;border-radius:8px;background:#6e7a58;color:#fff;font-size:15px;font-weight:700;cursor:pointer;">進入</button>' +
            '</div>';
        document.body.appendChild(ov);
        const pw = ov.querySelector('#gate-pw');
        const err = ov.querySelector('#gate-err');
        pw.focus();
        async function tryUnlock() {
            const h = await sha256(pw.value);
            if (h === PASS_HASH) {
                localStorage.setItem(KEY, JSON.stringify({ h, exp: Date.now() + DAYS * 864e5 }));
                ov.remove();
                const s = document.getElementById('gate-hide'); if (s) s.remove();
            } else {
                err.textContent = '密碼錯誤';
                pw.value = ''; pw.focus();
            }
        }
        ov.querySelector('#gate-btn').addEventListener('click', tryUnlock);
        pw.addEventListener('keydown', e => { if (e.key === 'Enter') tryUnlock(); });
    }

    if (document.body) buildOverlay();
    else document.addEventListener('DOMContentLoaded', buildOverlay);
})();
