import { bitable, FieldType } from '@lark-base-open/js-sdk';
import PptxGenJS from 'pptxgenjs';

/* ===== 設定：預設值，會被同資料夾的 config.json 覆蓋 ===== */
const DEFAULTS = {
  '專案表_標題欄位': '活動名稱',
  '專案表_關聯欄位': '關聯曝光需求',
  '專案表_報告存放欄位': '曝光報告',
  '封面_顯示欄位': ['POC', '優惠期間(起)', '優惠期間(訖)'],
  '曝光表_要截圖欄位': '要截圖',
  '曝光表_沒勾要截圖但有截圖也放入': false,
  '曝光表_渠道欄位': '渠道類別',
  '曝光表_時間欄位': '曝光時間',
  '曝光表_狀態欄位': '曝光狀態',
  '曝光表_備註欄位': '曝光備註',
  '曝光表_負責人欄位': '渠道 Owner',
  '曝光表_截圖欄位': '曝光截圖提供',
  '曝光表_連結欄位': '曝光連結提供',
  '每頁最多幾張截圖': 4,
  '報告名稱': '曝光報告',
  '主色': 'FF5B00',
};
let CFG = { ...DEFAULTS };
async function loadConfig() {
  try {
    const r = await withTimeout(fetch('config.json?v=' + Date.now(), { cache: 'no-store' }), 5000, '讀取設定');
    if (r.ok) CFG = { ...DEFAULTS, ...(await r.json()) };
  } catch { /* 用預設值 */ }
}

/* ===== 小工具 ===== */
const $ = (s) => document.querySelector(s);
const URL_RE = /https?:\/\/[^\s,，、）)]+/g;
const pad = (n) => String(n).padStart(2, '0');
const today = () => { const d = new Date(); return { dash: `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`, slash: `${d.getFullYear()}/${pad(d.getMonth() + 1)}/${pad(d.getDate())}` }; };
const safeName = (s) => (s || '未命名').replace(/\//g, '-').replace(/[\\:*?"<>|\n\r]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
const clip = (s, n) => (s && s.length > n ? s.slice(0, n - 1) + '…' : s || '');
const blobToDataUrl = (b) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(b); });
const imgSize = (src) => new Promise((res) => { const i = new Image(); i.onload = () => res([i.naturalWidth || 4, i.naturalHeight || 3]); i.onerror = () => res([4, 3]); i.src = src; });
const byName = (fields, name) => name && fields.find((f) => f.name.trim() === String(name).trim());

const state = { busy: false, lastPptx: null, lastName: '' };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
function withTimeout(promise, ms, label) {
  return Promise.race([promise, sleep(ms).then(() => { throw new Error('逾時：' + label); })]);
}
function stage(text) { const el = $('#empty'); if (!$('#card').hidden) return; el.hidden = false; el.textContent = text; }

/* ===== 讀取 Lark 資料 ===== */
async function str(table, field, rid) {
  if (!field) return '';
  try { return ((await table.getCellString(field.id, rid)) || '').trim(); } catch { return ''; }
}
function linkIds(v) {
  if (!v) return [];
  const arr = Array.isArray(v) ? v : [v];
  return arr.flatMap((x) => x?.recordIds || x?.record_ids || (x?.id ? [x.id] : []));
}
function linkTableId(v, meta) {
  const arr = Array.isArray(v) ? v : [v];
  return arr.find((x) => x?.tableId || x?.table_id)?.tableId || arr.find((x) => x?.table_id)?.table_id || meta?.property?.tableId;
}

async function readLink(table, linkF, rid) {
  const tries = [
    ['方法一', () => table.getCellValue(linkF.id, rid)],
    ['方法二', async () => (await table.getField(linkF.id)).getValue(rid)],
    ['方法三', async () => (await table.getRecordById(rid))?.fields?.[linkF.id]],
  ];
  const errs = [];
  for (const [name, fn] of tries) {
    try {
      const v = await withTimeout(fn(), 5000, '讀取關聯欄位（' + name + '）');
      if (linkIds(v).length) { state.linkMethod = name; return v; }
      errs.push(name + '：沒有讀到關聯記錄');
    } catch (e) { errs.push(name + '：' + e.message); }
  }
  state.linkDebug = errs.join('；');
  return null;
}

async function getProject() {
  stage('讀取目前選取的位置…');
  const sel = await withTimeout(bitable.base.getSelection(), 8000, '讀取選取位置');
  if (!sel.tableId || !sel.recordId) return { reason: 'none' };
  stage('讀取表格欄位…');
  const table = await withTimeout(bitable.base.getTableById(sel.tableId), 8000, '開啟表格');
  const fields = await withTimeout(table.getFieldMetaList(), 8000, '讀取欄位清單');
  const linkF = byName(fields, CFG['專案表_關聯欄位']);
  if (!linkF) return { reason: 'wrongTable' };
  const titleF = byName(fields, CFG['專案表_標題欄位']) || fields[0];
  stage('讀取專案名稱…');
  const title = (await withTimeout(str(table, titleF, sel.recordId), 8000, '讀取專案名稱').catch(() => '')) || '未命名專案';
  stage('讀取關聯的曝光需求…');
  const linkVal = await readLink(table, linkF, sel.recordId);
  const ids = linkIds(linkVal);
  const expTableId = linkTableId(linkVal, linkF) || (await withTimeout(table.getField(linkF.id).then((f) => f.getMeta()), 8000, '關聯欄位設定').catch(() => null))?.property?.tableId;
  const outF = fields.find((f) => f.type === FieldType.Attachment && f.name.trim() === CFG['專案表_報告存放欄位']);
  return { sel, table, fields, title, ids, expTableId, outF };
}

async function readExposures(p, onProgress) {
  if (!p.ids.length || !p.expTableId) return { items: [], skipped: 0 };
  const t = await bitable.base.getTableById(p.expTableId);
  const fs = await t.getFieldMetaList();
  const F = {
    need: byName(fs, CFG['曝光表_要截圖欄位']),
    channel: byName(fs, CFG['曝光表_渠道欄位']),
    time: byName(fs, CFG['曝光表_時間欄位']),
    status: byName(fs, CFG['曝光表_狀態欄位']),
    note: byName(fs, CFG['曝光表_備註欄位']),
    owner: byName(fs, CFG['曝光表_負責人欄位']),
    shots: byName(fs, CFG['曝光表_截圖欄位']),
    link: byName(fs, CFG['曝光表_連結欄位']),
  };
  const shotsField = F.shots ? await t.getField(F.shots.id) : null;
  const items = []; let skipped = 0; let done = 0;
  await Promise.all(p.ids.map(async (rid) => withTimeout((async () => {
    const need = F.need ? (await t.getCellValue(F.need.id, rid).catch(() => false)) === true : true;
    const shotVal = F.shots ? await t.getCellValue(F.shots.id, rid).catch(() => null) : null;
    const shotCount = Array.isArray(shotVal) ? shotVal.length : 0;
    if (!need && !(CFG['曝光表_沒勾要截圖但有截圖也放入'] && shotCount)) { skipped++; onProgress?.(++done, p.ids.length); return; }
    const [channel, timeStr, status, note, owner, linkStr] = await Promise.all([F.channel, F.time, F.status, F.note, F.owner, F.link].map((f) => str(t, f, rid)));
    const timeRaw = F.time ? await t.getCellValue(F.time.id, rid).catch(() => null) : null;
    let urls = [];
    if (shotsField && shotCount) { try { urls = (await shotsField.getAttachmentUrls(rid)) || []; } catch { urls = []; } }
    const links = [...new Set((linkStr.match(URL_RE) || []))];
    items.push({ rid, channel: channel || '未填渠道', timeStr, timeRaw: typeof timeRaw === 'number' ? timeRaw : Infinity, status, note, owner, linkText: linkStr, links, urls, shots: [] });
    onProgress?.(++done, p.ids.length);
  })(), 20000, '讀取曝光需求').catch(() => { skipped++; })));
  items.sort((a, b) => a.channel.localeCompare(b.channel, 'zh-Hant') || a.timeRaw - b.timeRaw);
  return { items, skipped };
}

async function readCover(p) {
  const out = [];
  for (const name of CFG['封面_顯示欄位'] || []) {
    const f = byName(p.fields, name);
    const v = await str(p.table, f, p.sel.recordId);
    if (v) out.push([name, v]);
  }
  return out;
}

/* ===== 產生 PPT ===== */
const FONT = 'Arial', INK = '1F2328', MUTED = '6B7280', LINE = 'DDE1E6', LINK = '1A5FD6';

function header(pptx, s, title, sub) {
  const accent = CFG['主色'];
  s.background = { color: 'FFFFFF' };
  s.addText(title, { x: 0.6, y: 0.4, w: 10.2, h: 0.65, fontSize: 26, bold: true, color: INK, fontFace: FONT, margin: 0, fit: 'shrink' });
  s.addShape(pptx.ShapeType.roundRect, { x: 11.13, y: 0.47, w: 1.6, h: 0.42, rectRadius: 0.1, fill: { color: 'FFF1E8' }, line: { color: 'FFF1E8' } });
  s.addText(CFG['報告名稱'], { x: 11.13, y: 0.47, w: 1.6, h: 0.42, fontSize: 12, bold: true, color: accent, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
  if (sub) s.addText(sub, { x: 0.6, y: 1.05, w: 12.1, h: 0.35, fontSize: 13, color: MUTED, fontFace: FONT, margin: 0 });
}

function placeholder(pptx, s, x, y, w, h, text) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: 'F3F4F6' }, line: { color: 'E5E7EB', width: 0.75, dashType: 'dash' } });
  s.addText(text, { x, y, w, h, fontSize: 18, color: '9CA3AF', fontFace: FONT, align: 'center', valign: 'middle' });
}

function grid(pptx, s, shots, left, top, width, bottom) {
  const n = shots.length, cols = n === 1 ? 1 : n <= 4 ? 2 : 3, rows = Math.ceil(n / cols), gap = 0.25;
  const cw = (width - gap * (cols - 1)) / cols, ch = (bottom - top - gap * (rows - 1)) / rows;
  shots.forEach((sh, i) => {
    const c = i % cols, r = Math.floor(i / cols), sc = Math.min(cw / sh.w, ch / sh.h), w = sh.w * sc, h = sh.h * sc;
    const x = left + c * (cw + gap) + (cw - w) / 2, y = top + r * (ch + gap) + (ch - h) / 2;
    s.addImage({ data: sh.dataUrl, x, y, w, h });
    s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: 'FFFFFF', transparency: 100 }, line: { color: LINE, width: 0.75 } });
  });
}

async function buildPptx(title, cover, items) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.title = `${title} ${CFG['報告名稱']}`;
  const accent = CFG['主色'];
  const withShot = items.filter((i) => i.shots.length).length;
  const withLink = items.filter((i) => i.links.length).length;

  // 封面
  const c = pptx.addSlide();
  c.background = { color: 'FFFFFF' };
  c.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 1.3, w: 1.6, h: 0.42, rectRadius: 0.1, fill: { color: 'FFF1E8' }, line: { color: 'FFF1E8' } });
  c.addText(CFG['報告名稱'], { x: 0.8, y: 1.3, w: 1.6, h: 0.42, fontSize: 12, bold: true, color: accent, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
  c.addText(title, { x: 0.8, y: 1.95, w: 11.7, h: 1.5, fontSize: 36, bold: true, color: INK, fontFace: FONT, margin: 0, valign: 'top', fit: 'shrink' });
  const stats = [[String(items.length), '則曝光'], [String(withShot), '則已有截圖'], [String(withLink), '則已有連結']];
  stats.forEach(([num, label], i) => {
    const x = 0.8 + i * 2.6;
    c.addText(num, { x, y: 3.75, w: 2.4, h: 0.8, fontSize: 40, bold: true, color: i === 0 ? accent : INK, fontFace: FONT, margin: 0 });
    c.addText(label, { x, y: 4.55, w: 2.4, h: 0.35, fontSize: 13, color: MUTED, fontFace: FONT, margin: 0 });
  });
  const info = cover.map(([k, v]) => `${k}：${clip(v, 60)}`).concat([`產出日期：${today().slash}`]);
  c.addText(info.join('\n'), { x: 0.8, y: 5.4, w: 11.7, h: 1.4, fontSize: 13, color: MUTED, fontFace: FONT, margin: 0, valign: 'top', paraSpaceAfter: 4 });

  // 總覽表
  if (items.length) {
    const ROWS = 12;
    for (let i = 0; i < items.length; i += ROWS) {
      const s = pptx.addSlide();
      const pages = Math.ceil(items.length / ROWS);
      header(pptx, s, pages > 1 ? `曝光總覽（${i / ROWS + 1}/${pages}）` : '曝光總覽', `${title}`);
      const hdr = ['渠道', '曝光時間', '狀態', '截圖', '連結'].map((t) => ({ text: t, options: { bold: true, color: MUTED, fill: { color: 'F6F7F9' } } }));
      const rows = items.slice(i, i + ROWS).map((it) => [
        clip(it.channel, 30), it.timeStr || '—', it.status || '—',
        { text: it.shots.length ? `${it.shots.length} 張` : '待補', options: { color: it.shots.length ? INK : 'B45309' } },
        { text: it.links.length ? '有' : '待補', options: { color: it.links.length ? INK : 'B45309' } },
      ]);
      s.addTable([hdr, ...rows], { x: 0.6, y: 1.6, w: 12.13, colW: [4.4, 2.6, 2.2, 1.4, 1.53], fontSize: 12, fontFace: FONT, color: INK, border: { type: 'solid', pt: 0.5, color: 'E5E7EB' }, rowH: 0.38, valign: 'middle' });
    }
  }

  // 每則曝光
  const per = Math.max(1, Number(CFG['每頁最多幾張截圖']) || 4);
  for (const it of items) {
    const chunks = it.shots.length ? [] : [[]];
    for (let i = 0; i < it.shots.length; i += per) chunks.push(it.shots.slice(i, i + per));
    chunks.forEach((chunk, ci) => {
      const s = pptx.addSlide();
      const t = chunks.length > 1 ? `${it.channel}（${ci + 1}/${chunks.length}）` : it.channel;
      header(pptx, s, t, [it.timeStr, it.status].filter(Boolean).join(' · ') || '曝光時間待補');
      let y = 1.5;
      if (it.note) {
        s.addText(clip(it.note.replace(/\s+/g, ' '), 140), { x: 0.6, y, w: 12.1, h: 0.5, fontSize: 12, color: INK, fontFace: FONT, margin: 0, valign: 'top', fit: 'shrink' });
        y += 0.55;
      }
      const top = y + 0.1, bottom = 6.45;
      if (chunk.length) grid(pptx, s, chunk, 0.6, top, 12.13, bottom);
      else placeholder(pptx, s, 0.6, top, 12.13, bottom - top, '截圖待補');
      const linkRuns = it.links.length
        ? [{ text: '連結：', options: { color: MUTED } }, { text: clip(it.links[0], 85), options: { color: LINK, underline: true, hyperlink: { url: it.links[0] } } }]
        : [{ text: '連結：', options: { color: MUTED } }, { text: it.linkText ? clip(it.linkText, 85) : '待補', options: { color: it.linkText ? INK : 'B45309' } }];
      s.addText(linkRuns, { x: 0.6, y: 6.6, w: 9.4, h: 0.35, fontSize: 12, fontFace: FONT, margin: 0 });
      if (it.owner) s.addText(`渠道 Owner：${it.owner}`, { x: 10.1, y: 6.6, w: 2.63, h: 0.35, fontSize: 11, color: MUTED, fontFace: FONT, margin: 0, align: 'right' });
    });
  }
  return pptx;
}

/* ===== 介面 ===== */
function setStatus(kind, text) { const el = $('#status'); el.className = 'status ' + kind; el.textContent = text; el.hidden = !text; }
function setSteps(active, done = []) {
  document.querySelectorAll('#steps li').forEach((li) => { const k = li.dataset.k; li.className = done.includes(k) ? 'done' : k === active ? 'active' : ''; });
  $('#steps').hidden = !active && !done.length;
}
function setBusy(b) { state.busy = b; $('#go').disabled = b || !state.ready; $('#go').textContent = b ? '產生中…' : '產生報告'; }
function resetResult() { $('#result').hidden = true; setStatus('', ''); setSteps(null); }

async function refreshPreview() {
  if (state.busy) return;
  state.ready = false;
  try {
    const p = await getProject();
    if (p.reason === 'none') { show('empty', `在「專案資訊」表點選一個專案，就能產生它的${CFG['報告名稱']}。`); return setBusy(false); }
    if (p.reason === 'wrongTable') { show('empty', `目前不在專案表。請切到「專案資訊」，點選一個專案。`); return setBusy(false); }
    $('#empty').hidden = true; $('#card').hidden = false;
    $('#rec-title').textContent = p.title;
    $('#rec-meta').textContent = p.ids.length ? `關聯 ${p.ids.length} 則曝光需求，計算中…` : ('這個專案讀不到關聯的曝光需求' + (state.linkDebug ? `（${state.linkDebug}）` : ''));
    $('#out-hint').textContent = (p.outF ? `報告會存到這個專案的「${p.outF.name}」欄位，舊版本會保留。` : `專案表沒有「${CFG['專案表_報告存放欄位']}」附件欄位，報告只會下載到電腦。`) + (state.linkMethod && state.linkMethod !== '方法一' ? `（讀取方式：${state.linkMethod}）` : '');
    state.ready = p.ids.length > 0;
    setBusy(false);
    if (p.ids.length) {
      const key = p.sel.recordId; state.previewKey = key;
      const { items } = await readExposures(p);
      if (state.previewKey !== key) return;
      const shot = items.filter((i) => i.urls.length).length;
      $('#rec-meta').textContent = items.length ? `${items.length} 則要截圖的曝光，${shot} 則已有截圖` : '沒有勾選「要截圖」的曝光需求';
    }
  } catch (e) { console.error(e); setStatus('bad', '讀不到表格資料：' + (e?.message || e) + '。請截圖給 Sarah。'); $('#empty').textContent = '讀取失敗'; }
}
function show(id, text) { $('#card').hidden = true; $('#empty').hidden = false; $('#empty').textContent = text; $('#out-hint').textContent = ''; }

async function generate() {
  if (state.busy) return;
  setBusy(true); resetResult();
  try {
    setSteps('read');
    const p = await getProject();
    if (p.reason) { setSteps(null); setStatus('bad', '請先在「專案資訊」表點選一個專案。'); return; }
    const cover = await readCover(p);
    const { items } = await readExposures(p, (d, n) => { $('#step-read').textContent = `讀取曝光需求（${d}/${n}）`; });
    $('#step-read').textContent = '讀取曝光需求';
    if (!items.length) { setSteps(null); setStatus('warn', '這個專案沒有勾選「要截圖」的曝光需求，所以沒有產生報告。'); return; }

    setSteps('download', ['read']);
    const total = items.reduce((a, i) => a + i.urls.length, 0); let got = 0, failed = 0;
    for (const it of items) {
      for (const url of it.urls) {
        $('#step-download').textContent = `下載截圖（${got + failed + 1}/${total}）`;
        try {
          const r = await fetch(url); if (!r.ok) throw new Error(r.status);
          const dataUrl = await blobToDataUrl(await r.blob());
          const [w, h] = await imgSize(dataUrl);
          it.shots.push({ dataUrl, w, h }); got++;
        } catch { failed++; }
      }
    }
    $('#step-download').textContent = '下載截圖';

    setSteps('build', ['read', 'download']);
    const pptx = await buildPptx(p.title, cover, items);
    const name = `${safeName(p.title)}_${CFG['報告名稱']}_${today().dash}.pptx`;
    state.lastPptx = pptx; state.lastName = name;

    let saved = false;
    if (p.outF) {
      setSteps('save', ['read', 'download', 'build']);
      const blob = await pptx.write({ outputType: 'blob' });
      const file = new File([blob], name, { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
      try {
        const [token] = await bitable.base.batchUploadFile([file]);
        const existing = (await p.table.getCellValue(p.outF.id, p.sel.recordId).catch(() => null)) || [];
        await p.table.setCellValue(p.outF.id, p.sel.recordId, [...existing, { name, size: file.size, type: file.type, token, timeStamp: Date.now() }]);
        saved = true;
      } catch {
        try { const f = await p.table.getField(p.outF.id); await f.setValue(p.sel.recordId, file); saved = true; } catch { /* 改用下載 */ }
      }
    }
    setSteps(null, saved ? ['read', 'download', 'build', 'save'] : ['read', 'download', 'build']);
    const missingShot = items.filter((i) => !i.shots.length).length;
    const msg = [`報告完成：${items.length} 則曝光，${got} 張截圖。`];
    if (missingShot) msg.push(`${missingShot} 則還沒有截圖，已留「待補」。`);
    if (failed) msg.push(`${failed} 張截圖下載失敗，沒有放進報告。`);
    if (p.outF) msg.push(saved ? `已存到「${p.outF.name}」欄位。` : `存回「${p.outF.name}」欄位失敗，請改用下面的按鈕下載。`);
    setStatus(failed || (p.outF && !saved) ? 'warn' : 'ok', msg.join(''));
    $('#result').hidden = false; $('#file-name').textContent = name;
    if (!p.outF) await download();
  } catch (e) {
    console.error(e); setSteps(null);
    setStatus('bad', '產生報告時發生錯誤，請重新整理頁面再試一次。');
  } finally { setBusy(false); }
}

async function download() {
  if (!state.lastPptx) return;
  try { await state.lastPptx.writeFile({ fileName: state.lastName }); } catch { setStatus('bad', '下載失敗，請再按一次。'); }
}

async function applyTheme() {
  try { const t = await bitable.bridge.getTheme(); document.documentElement.dataset.theme = t === 'DARK' ? 'dark' : 'light'; } catch { /* 預設淺色 */ }
}

window.addEventListener('DOMContentLoaded', async () => {
  $('#go').onclick = generate; $('#dl').onclick = download;
  await loadConfig();
  $('#app-title').textContent = CFG['報告名稱'];
  applyTheme();
  try { bitable.bridge.onThemeChange(applyTheme); } catch { /* 忽略 */ }
  refreshPreview();
  try { bitable.base.onSelectionChange(() => { if (!state.busy) { resetResult(); refreshPreview(); } }); } catch { /* 忽略 */ }
});
