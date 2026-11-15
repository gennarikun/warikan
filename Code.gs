/** @OnlyCurrentDoc */
// ふたりの割り勘：記録の置き場（Google スプレッドシートに付ける Apps Script）2026-11-15
// 表「payments」と「settings」を使う。列は見出しの名前で読む（並びに頼らない）
const PCOLS = ['id','date','payer','amount','ratioA','title','createdAt','updatedAt','settledAt','deletedAt'];

function sheet_(name, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(cols); }
  return sh;
}
function rows_(sh) {
  const v = sh.getDataRange().getValues(); const head = v.shift() || [];
  return { head, list: v.map((r, i) => { const o = {_row: i + 2}; head.forEach((h, j) => o[h] = r[j] instanceof Date ? Utilities.formatDate(r[j], 'Asia/Tokyo', 'yyyy-MM-dd') : r[j]); return o; }) };
}
function settings_() {
  const sh = sheet_('settings', ['key','value']); const o = {};
  rows_(sh).list.forEach(r => { if (r.key) o[r.key] = String(r.value); });
  return o;
}
// 文字は先頭に ' を付けて入れる（「=」で始まる名前や費用名を計算式にさせない・日付に変えさせない）
function txt_(v) { v = String(v === undefined || v === null ? '' : v); return v === '' ? '' : "'" + v; }
// 毎月の決まった支払い：決まった形の物だけ残す
function recurring_(v) {
  let a; try { a = JSON.parse(v); } catch (e) { return null; }
  if (!Array.isArray(a)) return null;
  return JSON.stringify(a.slice(0, 20).filter(r => r && /^[a-z0-9]{1,24}$/.test(r.id) && Number(r.amount) > 0 && (r.payer === 'A' || r.payer === 'B') && Number(r.day) >= 1 && Number(r.day) <= 28 && /^\d{4}-\d{2}$/.test(r.start))
    .map(r => ({ id: r.id, title: String(r.title || '').slice(0, 30), amount: Math.round(Number(r.amount)), payer: r.payer, ratioA: Math.min(100, Math.max(0, Number(r.ratioA) || 0)), day: Math.round(Number(r.day)), start: r.start })));
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() {
  const r = rows_(sheet_('payments', PCOLS));
  const payments = r.list.map(p => { const o = {}; PCOLS.forEach(c => o[c] = p[c] === undefined ? '' : p[c]); o.date = String(o.date); o.amount = Number(o.amount); o.ratioA = Number(o.ratioA); return o; });
  return out_({ payments, settings: settings_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000); // 2人が同時に書いても1件ずつ
  try {
    const b = JSON.parse(e.postData.contents); const now = "'" + new Date().toISOString(); // 先頭の ' で文字のまま残す（日付に変えられないように）
    const sh = sheet_('payments', PCOLS); const r = rows_(sh); const col = c => r.head.indexOf(c) + 1;
    if (b.action === 'add' || b.action === 'update') {
      const p = b.p; if (!p || !p.id || !(Number(p.amount) > 0)) return out_({ ok: false, error: '金額が正しくありません' });
      if (p.payer !== 'A' && p.payer !== 'B') return out_({ ok: false, error: '払った人が正しくありません' });
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(p.date))) return out_({ ok: false, error: '日付が正しくありません' });
      p.ratioA = Math.min(100, Math.max(0, Number(p.ratioA) || 0)); p.amount = Math.round(Number(p.amount)); p.title = String(p.title || '').slice(0, 60);
      const hit = r.list.find(x => x.id === p.id);  // 同じ id は上書き＝2回届いても1件
      if (hit) {
        if (hit.settledAt) return out_({ ok: false, error: '精算済みは直せません' });
        ['date','payer','amount','ratioA','title'].forEach(c => sh.getRange(hit._row, col(c)).setValue(c === 'date' || c === 'title' ? txt_(p[c]) : p[c]));
        sh.getRange(hit._row, col('updatedAt')).setValue(now);
      } else {
        sh.appendRow(PCOLS.map(c => c === 'date' || c === 'title' || c === 'id' ? txt_(p[c]) : c === 'createdAt' || c === 'updatedAt' ? now : (p[c] === undefined ? '' : p[c])));
      }
    } else if (b.action === 'delete') {
      const hit = r.list.find(x => x.id === b.id); if (hit && !hit.settledAt) sh.getRange(hit._row, col('deletedAt')).setValue(now);
    } else if (b.action === 'settle') {
      r.list.forEach(x => { if (!x.settledAt && !x.deletedAt) sh.getRange(x._row, col('settledAt')).setValue("'" + b.at); });
    } else if (b.action === 'unsettle') {
      r.list.forEach(x => { if (String(x.settledAt) === String(b.at)) sh.getRange(x._row, col('settledAt')).setValue(''); });
    } else if (b.action === 'saveSettings') {
      const s = sheet_('settings', ['key','value']); const sr = rows_(s);
      const ok = { nameA: 1, nameB: 1, titles: 1, imgA: 1, imgB: 1, recurring: 1 };
      Object.keys(b.s || {}).forEach(k => { if (!ok[k]) return; let v = String(b.s[k] === undefined || b.s[k] === null ? '' : b.s[k]);
        if ((k === 'imgA' || k === 'imgB') && v && !/^data:image\/(jpeg|png|webp|gif);base64,[A-Za-z0-9+/=]+$/.test(v)) return;
        if (k === 'recurring') { v = recurring_(v); if (v === null) return; }
        if (v.length > 45000) return; const hit = sr.list.find(x => x.key === k); if (hit) s.getRange(hit._row, 2).setValue(txt_(v)); else s.appendRow([k, txt_(v)]); });
    } else return out_({ ok: false, error: '知らない操作です' });
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}
