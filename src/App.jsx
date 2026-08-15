import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import {
  Package, Search, ScanLine, Plus, Minus, Download, Home, List, X, Check,
  Camera, AlertCircle, Trash2, Edit3, Save, Image as ImageIcon, Upload, User,
  Shield, Eye, EyeOff, ClipboardCheck, Lock, LogOut, Database, Cloud,
  RefreshCw, Settings as SettingsIcon, CheckCircle2, XCircle, Layers,
  FileSpreadsheet, ArrowRight, FileCheck, WifiOff, Zap, Send, Clock,
  ThumbsUp, ThumbsDown, Inbox, ArrowLeftRight, Receipt, MapPin, Users,
} from 'lucide-react';

const INVOICE_API = "/api/claude";
const DRIVE_FOLDER_DEFAULT = "1Egu5XH0UInn4ol6V2FlI06-TMUdnXO4D";
const MODELS = [
  { id: "claude-opus-4-6",    label: "Opus 4.6" },
  { id: "claude-sonnet-4-6",  label: "Sonnet 4.6 ✦" },
  { id: "claude-haiku-4-5-20251001", label: "Haiku 4.5" },
];
const INVOICE_PROMPT = `ตอบด้วย JSON เท่านั้น ห้ามพิมพ์ข้อความอื่นใดก่อนหรือหลัง JSON ห้าม backtick ห้าม markdown
สกัดข้อมูลจากใบกำกับภาษี/ใบส่งของที่อัปโหลดมา และตอบเป็น JSON ตามโครงสร้างนี้:
{
  "invoice_no": string|null,
  "invoice_date": string|null,
  "vendor_name": string|null,
  "vendor_tax_id": string|null,
  "document_type": "invoice"|"receipt"|"credit_note"|"debit_note"|"quotation"|"other"|null,
  "vendor_address": string|null,
  "vendor_branch": string|null,
  "price_type": "incl"|"excl",
  "products": [
    {
      "no": number|null,
      "description": string|null,
      "carton_size": number|null,
      "carton": number|null,
      "carton_unit": string|null,
      "ea": number|null,
      "ea_unit": string|null,
      "qty": number|null,
      "price_ea": number|null,
      "amount": number|null,
      "special_discount": number|null,
      "vat": "v"|"n"|null
    }
  ]
}

กฎสำคัญ:
- price_type="incl" ถ้าราคารวม VAT แล้ว, "excl" ถ้าแยก VAT
- vat="v" ถ้ามี VAT 7%, "n" ถ้าไม่มี VAT

⚠️ กฎเรื่อง carton_size / carton / ea — สำคัญมาก:
- carton_size = จำนวนชิ้นต่อ 1 ลัง/กล่อง/แพ็ค (pack size)
- carton = จำนวนลัง/แพ็ค (number of packs)
- ea = จำนวนชิ้นเศษ (loose pieces)
- ระบบจะคำนวณ qty อัตโนมัติ: qty = carton_size × carton + ea
- ห้ามกรอก qty โดยตรง — ระบบจะคำนวณจาก cs/ca/ea เอง

🔍 การตีความ pack pattern (สำคัญ):
- ถ้า "รายละเอียด/Description" มี pattern แบบ "x 12", "X 6", "(12 ชิ้น)", "x12 pcs", "12 ขวด/แพ็ค", "8 ชิ้น x 6" เป็นต้น
  → ตัวเลขนั้นคือ carton_size (จำนวนต่อแพ็ค)
  → คอลัมน์ Quantity = carton (จำนวนแพ็คที่ซื้อ)
  → ea = 0
  ตัวอย่าง: "น้ำทิพย์ 550 มล. x 12" Quantity=3 EACH → carton_size=12, carton=3, ea=0 → qty=36
  ตัวอย่าง: "ผ้าอนามัย 8 ชิ้น x 6" Quantity=1 EACH → carton_size=6, carton=1, ea=0 → qty=6

- ถ้าหน่วย (unit) เป็น Kilogram/กก./น้ำหนัก หรือ EACH ที่ไม่มี pack pattern
  → ใส่ Quantity ใน ea ทั้งหมด, carton_size=null, carton=null
  ตัวอย่าง: "บล็อกโคลีนอก 1 กก." Quantity=1.358 KG → ea=1.358, carton_size=null, carton=null

- ถ้าไม่แน่ใจว่าตัวเลขใน description คือ pack size หรือไม่ → ใส่ใน ea

⚠️ กฎเรื่อง price_ea (ราคาต่อชิ้น) — สำคัญมาก:
- price_ea ต้องเป็น "ราคาต่อ 1 ชิ้น" เสมอ (per piece / per smallest unit)
- ถ้าใบกำกับแสดง "ราคา/หน่วย" เป็น ราคาต่อลัง/กล่อง/หีบ → ให้หารด้วย carton_size ก่อน
  ตัวอย่าง: บิลเขียน 606.48 บาท/ลัง, carton_size=12 → price_ea = 606.48/12 = 50.54
- ถ้าไม่แน่ใจ ให้ปล่อย price_ea = null ระบบจะคำนวณจาก amount/qty เอง

⚠️ กฎเรื่อง amount (ยอดก่อนหักส่วนลด):
- amount = ราคารวม "ก่อน" หักส่วนลด = ราคา/หน่วย × จำนวน
- ถ้าใบกำกับแสดงคอลัมน์ "จำนวนเงิน" หรือ "amount" เป็นยอดหลังหักแล้ว → ให้บวก special_discount กลับ
  ตัวอย่าง: บิลเขียน จำนวนเงิน=1177.54, ส่วนลด=35.42 → amount = 1177.54 + 35.42 = 1212.96

⚠️ กฎเรื่อง carton_unit / ea_unit:
- carton_unit = หน่วยลัง เช่น "ลัง", "กล่อง", "หีบ", "แพ็ค"
- ea_unit = หน่วยชิ้น/เศษ เช่น "ชิ้น", "ขวด", "ซอง", "กก", "ก."
- ถ้าอ่านหน่วยไม่ชัด ปล่อย null ห้ามเดา

- สกัดทุกรายการสินค้า ไม่ละเว้น
- invoice_date ให้เป็น DD/MM/YYYY หรือ YYYY-MM-DD`;

const BARCODE_PROMPT = list => `คุณได้รับรูปภาพสินค้าและรายการชื่อสินค้าจากใบกำกับ
จงหาบาร์โค้ดจากรูป (EAN-13 หรือรหัสสินค้า) และจับคู่กับชื่อในรายการ
ตอบเป็น JSON array เท่านั้น:
[{"barcode":"xxx","match":"ชื่อสินค้าที่ตรงกัน","description_image":"ชื่อสินค้าจากรูป"}]
รายการสินค้า:
${list}`;

const STEPS = ["อัปโหลด", "สแกนสินค้า", "ตรวจสอบ", "สรุป"];


const storage = {
  get:    async (key)        => { try { const v = localStorage.getItem(key); return v != null ? { value: v } : null; } catch { return null; } },
  set:    async (key, value) => { try { localStorage.setItem(key, String(value)); return { key, value }; } catch { return null; } },
  delete: async (key)        => { try { localStorage.removeItem(key); return { key, deleted: true }; } catch { return null; } },
};
const safeGet = (k, def = '') => { try { return localStorage.getItem(k) ?? def; } catch { return def; } };
const safeSet = (k, v) => { try { localStorage.setItem(k, v); } catch {} };

function useDebouncedStorage(key, value, ready, delay = 800) {
  const timerRef = useRef(null);
  const pendingRef = useRef(false);
  useEffect(() => {
    if (!ready) return;
    if (timerRef.current) clearTimeout(timerRef.current);
    pendingRef.current = true;
    timerRef.current = setTimeout(async () => {
      if (!pendingRef.current) return;
      pendingRef.current = false;
      try { await storage.set(key, JSON.stringify(value)); } catch {}
    }, delay);
    return () => { if (timerRef.current) clearTimeout(timerRef.current); };
  }, [key, value, ready, delay]);
}

function useWinWidth() {
  const [w, setW] = useState(typeof window !== 'undefined' ? window.innerWidth : 800);
  useEffect(() => {
    const fn = () => setW(window.innerWidth);
    window.addEventListener('resize', fn);
    return () => window.removeEventListener('resize', fn);
  }, []);
  return w;
}

async function supabaseFindProduct(cfg, code) {
  const { url, anonKey, tableName, stockTableName } = cfg;
  if (!url || !anonKey) return null;
  const base = url.replace(/\/$/, '');
  const h = { apikey: anonKey, Authorization: `Bearer ${anonKey}` };
  const col = encodeURIComponent('"รหัสสินค้า"');
  const priceTable = tableName || 'product_price';
  const stockTable = stockTableName || 'product_stock';

  // Query both tables in parallel
  const [priceRes, stockRes] = await Promise.all([
    fetch(`${base}/rest/v1/${priceTable}?${col}=eq.${encodeURIComponent(code)}&limit=1`, { headers: h }),
    fetch(`${base}/rest/v1/${stockTable}?${col}=eq.${encodeURIComponent(code)}&limit=1`, { headers: h }),
  ]);

  let priceRow = null, stockRow = null;
  if (priceRes.ok) { const arr = await priceRes.json(); if (arr.length) priceRow = arr[0]; }
  if (stockRes.ok) { const arr = await stockRes.json(); if (arr.length) stockRow = arr[0]; }

  if (!priceRow && !stockRow) {
    if (!priceRes.ok && !stockRes.ok) throw new Error(`Supabase ${priceRes.status}/${stockRes.status}`);
    return null;
  }
  // Merge: prefer price row fields, fill missing from stock row
  return { ...(stockRow||{}), ...(priceRow||{}) };
}

function mapSupabaseRow(row, fallbackCode) {
  const get = (...keys) => {
    for (const k of keys) {
      if (row[k] != null && row[k] !== '') return row[k];
      for (let i = 1; i <= 3; i++) { const kn = k + '-' + i; if (row[kn] != null && row[kn] !== '') return row[kn]; }
    }
    return '';
  };
  const id = String(get('รหัสสินค้า', 'product_code', 'code', 'id') || fallbackCode || '');
  return {
    source: 'supabase', id,
    barcode: String(get('บาร์โค้ด', 'barcode') || id || fallbackCode || ''),
    productCode: id,
    name: String(get('ชื่อสินค้า', 'product_name', 'name') || '(ไม่มีชื่อ)'),
    category: String(get('ประเภท', 'category') || 'อื่นๆ'),
    unit: String(get('หน่วย', 'unit') || 'ชิ้น'),
    price: parseFloat(String(get('ราคา', 'price', 'ราคาขาย') || '0').replace(/[^\d.-]/g, '')) || 0,
    cost:  parseFloat(String(get('ทุนเฉลี่ย', 'ต้นทุน', 'cost', 'ราคาทุน') || '0').replace(/[^\d.-]/g, '')) || 0,
  };
}

async function sbFetch(url, key, table, rawQS) {
  const res = await fetch(`${url.replace(/\/$/, '')}/rest/v1/${table}?${rawQS}`, {
    headers: { apikey: key, Authorization: `Bearer ${key}` }
  });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  return res.json();
}

async function lookupVendorREST(sbUrl, sbKey, vendorName) {
  if (!sbUrl || !sbKey || !vendorName) return null;
  const h = { apikey: sbKey, Authorization: `Bearer ${sbKey}` };
  const col = encodeURIComponent('"ชื่อ-นามสกุล"');
  const sel = encodeURIComponent('"รหัส"');
  const get = async (op, kw) => {
    const val = op === 'ilike' ? `*${encodeURIComponent(kw)}*` : encodeURIComponent(kw);
    const r = await fetch(`${sbUrl}/rest/v1/vendor_info?select=${sel}&${col}=${op}.${val}&limit=1`, { headers: h });
    if (!r.ok) return null;
    const d = await r.json();
    return d[0]?.['รหัส'] ?? null;
  };
  const stripped = vendorName.trim().replace(/\s*(จำกัด|มหาชน|co\.?,?\s*ltd\.?|บจก\.?|หจก\.?|บมจ\.?)/gi, '').trim();
  return (await get('eq', vendorName)) ?? (await get('ilike', stripped)) ?? (await get('ilike', stripped.split(/\s+/)[0]));
}

function toYMD(raw) {
  if (!raw) return null;
  const s = String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const parts = s.split(/[\/\-\.]/);
  if (parts.length === 3) {
    let [a, b, c] = parts;
    if (String(a).length === 4) return `${a}-${String(b).padStart(2,'0')}-${String(c).padStart(2,'0')}`;
    const year = +c > 2400 ? +c - 543 : +c < 100 ? +c + 2000 : +c;
    return `${year}-${String(b).padStart(2,'0')}-${String(a).padStart(2,'0')}`;
  }
  const dt = new Date(s);
  if (!isNaN(dt)) return dt.toISOString().slice(0, 10);
  return null;
}

async function imgToBase64(file) {
  const ok = ['image/jpeg','image/png','image/gif','image/webp'];
  if (ok.includes(file.type)) {
    const img = new Image(), url = URL.createObjectURL(file);
    await new Promise((res, rej) => { img.onload = res; img.onerror = rej; img.src = url; });
    const c = document.createElement('canvas');
    c.width = img.width; c.height = img.height;
    c.getContext('2d').drawImage(img, 0, 0);
    URL.revokeObjectURL(url);
    return { base64: c.toDataURL('image/jpeg', 0.92).split(',')[1], mediaType: 'image/jpeg' };
  }
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res({ base64: r.result.split(',')[1], mediaType: file.type });
    r.onerror = rej;
    r.readAsDataURL(file);
  });
}

async function callClaude(content, extra = {}, model) {
  const r = await fetch(INVOICE_API, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model, max_tokens: 8000, temperature: 0, messages: [{ role: 'user', content }], ...extra }),
  });
  if (!r.ok) { const txt = await r.text(); throw new Error('API ' + r.status + ': ' + txt.slice(0, 150)); }
  return r.json();
}

function extractJSON(text) {
  if (!text) throw new Error('empty');
  const t = text.replace(/```json|```/g, '').trim();
  if (t.startsWith('[')) { const e = t.lastIndexOf(']'); if (e > 0) return JSON.parse(t.slice(0, e+1)); }
  const start = t.indexOf('{'); const end = t.lastIndexOf('}');
  if (start < 0 || end < 0) throw new Error('no JSON');
  return JSON.parse(t.slice(start, end + 1));
}

function recalc(p) {
  const cs = p.carton_size != null && p.carton_size !== '' ? +p.carton_size : null;
  const ca = p.carton != null && p.carton !== '' ? +p.carton : null;
  const ea = p.ea != null && p.ea !== '' ? +p.ea : 0;
  const am = p.amount != null && p.amount !== '' ? +p.amount : null;
  const sd = p.special_discount != null && p.special_discount !== '' ? +p.special_discount : 0;

  // qty: 4 branches per spec
  let qty;
  if (cs != null && ca != null)      qty = +(cs * ca + ea).toFixed(4);
  else if (ca != null)               qty = +(ca + ea).toFixed(4);
  else if (cs != null && ea > 0)     qty = +(cs * ea).toFixed(4);
  else if (ea > 0)                   qty = +ea.toFixed(4);
  else                               qty = null;

  // price_ea locked once user-entered
  const price_ea = p.price_ea != null && p.price_ea !== '' ? +p.price_ea
                 : (qty != null && qty !== 0 && am != null ? +(am / qty).toFixed(4) : null);

  const total     = (qty != null && price_ea != null) ? +(qty * price_ea - sd).toFixed(2) : null;
  const amount_sd = am != null ? +(am - sd).toFixed(2) : null;
  const diff      = (amount_sd != null && total != null) ? +(amount_sd - total).toFixed(2) : null;

  const vatCode = p.vat ?? 'v', pt = p._pt ?? 'incl';
  const excl_vat = total != null
    ? (vatCode === 'v' ? (pt === 'incl' ? +(total/1.07).toFixed(2) : total) : total)
    : null;
  const vat_amt = total != null
    ? (vatCode === 'v' ? (pt === 'incl' ? +(total - total/1.07).toFixed(2) : +(total*0.07).toFixed(2)) : 0)
    : null;

  return { ...p, qty, price_ea, total, amount_sd, diff, excl_vat, vat_amt };
}

function vatSummary(products = []) {
  let sdTot = 0, excl = 0, vatAmt = 0;
  for (const p of products) {
    const sd = p.special_discount != null ? +p.special_discount : 0;
    const net = p.total != null ? +p.total : 0, pt = p._pt ?? 'incl';
    sdTot += sd;
    excl += p.excl_vat != null ? +p.excl_vat : (p.vat === 'v' ? (pt === 'incl' ? +(net/1.07) : net) : net);
    vatAmt += p.vat_amt != null ? +p.vat_amt : 0;
  }
  const netTotal = excl + vatAmt;
  return { sdTot: +sdTot.toFixed(2), netTotal: +netTotal.toFixed(2), excl: +excl.toFixed(2), vatAmt: +vatAmt.toFixed(2) };
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob), a = document.createElement('a');
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
}

function buildStockExcelRows(data, docNo, countedAt) {
  const header = ['รหัสสินค้า', 'ชื่อสินค้า', 'หน่วย', 'จำนวนนับ', 'Location', 'ราคาขาย', 'ทุนเฉลี่ย', 'เวลาสแกน'];
  const rows = data.map(d => [
    d.barcode, d.productName, d.unit || '', d.qty,
    d.location || '', d.price || 0, d.cost || 0,
    d.scannedAt ? new Date(d.scannedAt).toLocaleString('th-TH') : '',
  ]);
  const meta = [
    [`เอกสารเลขที่: ${docNo || ''}`],
    [`วันเวลาส่ง: ${countedAt ? new Date(countedAt).toLocaleString('th-TH') : ''}`],
    [],
    header,
  ];
  return [...meta, ...rows];
}

function downloadStockExcel(data, filename, docNo, submittedAt) {
  const allRows = buildStockExcelRows(data, docNo, submittedAt);
  const ws = XLSX.utils.aoa_to_sheet(allRows);
  ws['!cols'] = [{ wch: 20 }, { wch: 30 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 20 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Stock');
  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  downloadBlob(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);
}

function downloadStockCSV(data, filename) {
  const csv = data.map(d => `${d.barcode},${d.qty},${d.price || 0},0`).join('\n');
  downloadBlob(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' }), filename);
}

function openPDFPrint(sub) {
  const OVERLAY_ID = '__pdf_overlay__', STYLE_ID = '__pdf_style__';
  document.getElementById(OVERLAY_ID)?.remove();
  document.getElementById(STYLE_ID)?.remove();
  const dateStr = new Date(sub.submittedAt).toLocaleDateString('th-TH');
  const rows = sub.data.map(d => `<tr><td>${d.barcode}</td><td>${d.productName || ''}</td><td style="text-align:center">${d.qty}</td><td style="text-align:center">${d.unit || ''}</td>${d.location ? `<td>${d.location}</td>` : '<td>-</td>'}<td style="text-align:right">${Number(d.price||0).toLocaleString()}</td></tr>`).join('');
  if (!document.getElementById('__sarabun_font__')) {
    const link = document.createElement('link');
    link.id = '__sarabun_font__'; link.rel = 'stylesheet';
    link.href = 'https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap';
    document.head.appendChild(link);
  }
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = `@media print{body>*:not(#${OVERLAY_ID}){display:none!important}#${OVERLAY_ID}{display:block!important;position:static!important;background:#fff;padding:0;margin:0;box-shadow:none}@page{size:A4 portrait;margin:12mm}}#${OVERLAY_ID}{position:fixed;inset:0;z-index:99999;background:#fff;overflow-y:auto;padding:20px;font-family:'Sarabun',sans-serif;color:#1e293b}#${OVERLAY_ID} .pr-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px}#${OVERLAY_ID} .pr-close{cursor:pointer;font-size:20px;color:#64748b;background:none;border:none;padding:4px 8px}#${OVERLAY_ID} .pr-print{cursor:pointer;background:#4361ee;color:#fff;border:none;padding:8px 20px;border-radius:8px;font-size:14px;font-family:inherit}#${OVERLAY_ID} h1{font-size:18px;font-weight:700;margin-bottom:10px}#${OVERLAY_ID} .meta{font-size:13px;margin:3px 0}#${OVERLAY_ID} table{width:100%;border-collapse:collapse;margin-top:14px;font-size:12px}#${OVERLAY_ID} thead tr{background:#4361ee;color:#fff}#${OVERLAY_ID} th{padding:8px;font-weight:700;text-align:left}#${OVERLAY_ID} td{padding:7px 8px;border-bottom:1px solid #e2e8f0}#${OVERLAY_ID} tr:nth-child(even) td{background:#f5f7ff}`;
  document.head.appendChild(style);
  const overlay = document.createElement('div');
  overlay.id = OVERLAY_ID;
  overlay.innerHTML = `<div class="pr-header"><div><h1>Stock Count Report</h1><p class="meta">เลขที่เอกสาร : <b>${sub.docNo || sub.id}</b></p><p class="meta">วันที่ : <b>${dateStr}</b></p><p class="meta">พนักงาน : <b>${sub.counter}</b></p><p class="meta">รายการ : <b>${sub.itemCount}</b> &nbsp; รวม : <b>${sub.totalQty}</b></p>${sub.note ? `<p class="meta">หมายเหตุ : <b>${sub.note}</b></p>` : ''}</div><div style="display:flex;gap:8px;align-items:flex-start"><button class="pr-print" id="__pdf_print_btn__">🖨️ พิมพ์ / PDF</button><button class="pr-close" id="__pdf_close_btn__">✕</button></div></div><table><thead><tr><th>รหัสสินค้า</th><th>ชื่อสินค้า</th><th style="text-align:center">จำนวน</th><th style="text-align:center">หน่วย</th><th>Location</th><th style="text-align:right">ราคาขาย</th></tr></thead><tbody>${rows}</tbody></table><p style="margin-top:16px;font-size:11px;color:#64748b;text-align:right">พิมพ์เมื่อ ${new Date().toLocaleString('th-TH')}</p>`;
  document.body.appendChild(overlay);
  document.getElementById('__pdf_print_btn__')?.addEventListener('click', () => window.print());
  document.getElementById('__pdf_close_btn__')?.addEventListener('click', () => { document.getElementById(OVERLAY_ID)?.remove(); document.getElementById(STYLE_ID)?.remove(); });
}

const DRIVE_FOLDER_RECORDER     = '1ACXWxpekq69xJEEuiwOkZZa5KosPtXXa';
const DRIVE_FOLDER_STOCK_COMPARE = '1dc62dEDZ8VWCV8nUx_uQg9h-3PoKSioi';

async function generateDocNo(prefix = 'RC') {
  const today = new Date();
  const dateStr = today.getFullYear().toString() + String(today.getMonth()+1).padStart(2,'0') + String(today.getDate()).padStart(2,'0');
  const stateKey = `docNoState_${prefix}`;
  let seq = 1;
  try { const r = await storage.get(stateKey); if (r?.value) { const s = JSON.parse(r.value); if (s.date === dateStr) seq = s.seq + 1; } } catch {}
  try { await storage.set(stateKey, JSON.stringify({ date: dateStr, seq })); } catch {}
  return `${prefix}-${dateStr}${String(seq).padStart(4,'0')}`;
}

function defaultViewFor(user) {
  if (user?.role === 'manager') return 'staff';
  if (!user) return 'count';
  if (user.role === 'manager') return 'dashboard';
  if (user.feature === 'invoice') return 'invoice';
  return 'count';
}

export default function CombinedApp() {
  const [products, setProducts] = useState([]);
  const [countEntries, setCountEntries] = useState([]);
  const [submissions, setSubmissions] = useState([]);
  const [view, setView] = useState('count');
  const [loaded, setLoaded] = useState(false);
  const [currentUser, setCurrentUser] = useState(null);
  const [supabaseConfig, setSupabaseConfig] = useState({ url: '', anonKey: '', tableName: 'product_price', stockTableName: 'product_stock' });
  const [dataSource, setDataSource] = useState('none');
  const [lastSyncAt, setLastSyncAt] = useState(null);
  const [connectionStatus, setConnectionStatus] = useState('unknown');
  const [countDate, setCountDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [countDraft, setCountDraft] = useState({ barcode: '', qty: '', checkResult: null, error: '' });
  const [showReport, setShowReport] = useState(false);   // รายงานเปิดได้ก่อนเลือกชื่อ
  const [pickedAt, setPickedAt] = useState(null);   // เลือกชื่อไว้กี่โมง — โชว์ในหน้ายืนยันก่อนส่ง
  const [compareState, setCompareState] = useState({
    selectedSub: null, compareData: [], loading: false, loadProgress: '',
    error: '', compareAt: null, driveSaving: false, driveResult: null,
  });
  const updateDraft = useCallback((patch) => setCountDraft(prev => ({ ...prev, ...patch })), []);

  useEffect(() => {
    (async () => {
      const tryLoad = async (key, setter, parse = true) => {
        try { const r = await storage.get(key); if (r?.value) setter(parse ? JSON.parse(r.value) : r.value); } catch {}
      };
      await tryLoad('countEntries', setCountEntries);
      await tryLoad('submissions', setSubmissions);
      await tryLoad('lastSyncAt', setLastSyncAt, false);
      let cfg = { url: '', anonKey: '', tableName: 'product_price', stockTableName: 'product_stock' };
      try { const r = await storage.get('supabaseConfig'); if (r?.value) { cfg = JSON.parse(r.value); setSupabaseConfig(cfg); } } catch {}
      let src = 'none';
      try { const r = await storage.get('dataSource'); if (r?.value) src = r.value; } catch {}
      if (cfg.url && cfg.anonKey && src === 'none') src = 'supabase';
      setDataSource(src);
      // เครื่องรวม — ไม่จำคนล่าสุด ต้องเลือกชื่อทุกครั้งที่เปิด (หลักการข้อ 1)
      // ร่างยังอยู่ เพราะผูกกับ counterId ของคน ไม่ใช่เครื่อง
      setLoaded(true);
    })();
  }, []);

  useDebouncedStorage('countEntries', countEntries, loaded);
  useDebouncedStorage('submissions', submissions, loaded);

  const handleLogin = (user) => { setCurrentUser(user); setView(defaultViewFor(user)); setPickedAt(Date.now()); storage.set('currentUser', JSON.stringify(user)).catch(() => {}); };
  const handleLogout = () => { setCurrentUser(null); storage.delete('currentUser').catch(() => {}); };


  const checkBarcode = async (barcode) => {
    const trimmed = barcode.trim();
    if (!trimmed) return null;
    if (supabaseConfig.url && supabaseConfig.anonKey) {
      try {
        const row = await supabaseFindProduct(supabaseConfig, trimmed);
        if (row) { const product = mapSupabaseRow(row, trimmed); setProducts(prev => prev.some(p => p.id === product.id || p.barcode === product.barcode) ? prev : [...prev, product]); setConnectionStatus('ok'); return product; }
        setConnectionStatus('ok'); return null;
      } catch (e) { setConnectionStatus('error'); const cached = products.find(p => p.id === trimmed || p.barcode === trimmed); if (cached) return { ...cached, source: 'cached' }; throw new Error(`ค้นหาไม่ได้: ${e.message}`); }
    }
    if (dataSource === 'seed') { return null; }
    throw new Error('ยังไม่ได้ตั้งค่า Supabase');
  };

  const addCountEntry = (entry) => {
    const now = new Date();
    const todayStr = now.toISOString().slice(0, 10);
    const dateStr = entry.countDate || todayStr;
    // If countDate is overridden to a past date (test mode), use 00:00 so all day's movements fall in window
    const timeStr = (entry.countDate && entry.countDate !== todayStr) ? '00:00:00' : now.toTimeString().slice(0, 8);
    const timestamp = new Date(`${dateStr}T${timeStr}`).toISOString();
    setCountEntries(prev => [{
      id: `e${Date.now()}_${Math.random().toString(36).slice(2,7)}`,
      barcode: entry.barcode, productName: entry.productName, productId: entry.productId || '',
      unit: entry.unit || '', price: entry.price || 0, cost: entry.cost || 0,
      qty: parseInt(entry.qty) || 0, notFound: !!entry.notFound,
      location: entry.location || '',
      counter: currentUser?.name || 'พนักงาน', counterId: currentUser?.id || 'unknown',
      timestamp,
    }, ...prev]);
  };

  const deleteCountEntry = (id) => setCountEntries(prev => prev.filter(e => e.id !== id));
  const clearMyEntries = () => setCountEntries(prev => prev.filter(e => e.counterId !== currentUser?.id));

  const submitForReview = async (grouped, note) => {
    const prefix = currentUser?.feature === 'stock_compare' ? 'ST' : 'RC';
    const docNo = await generateDocNo(prefix);
    const now = new Date().toISOString();
    const sub = {
      id: `sub${Date.now()}_${Math.random().toString(36).slice(2,7)}`, docNo,
      counter: currentUser?.name || 'พนักงาน', counterId: currentUser?.id || 'unknown',
      featureType: currentUser?.feature || 'recorder',
      submittedAt: now,
      startedAt: grouped.reduce((min, g) => g.scannedAt && g.scannedAt < min ? g.scannedAt : min, now),
      note: note || '', status: 'pending', reviewedAt: null, reviewedBy: null, reviewNote: '',
      itemCount: grouped.length, totalQty: grouped.reduce((s, g) => s + g.qty, 0), data: grouped,
    };
    setSubmissions(prev => [sub, ...prev]);
    return sub;
  };

  const reviewSubmission = (id, status, reviewNote) => setSubmissions(prev => prev.map(s => s.id === id ? { ...s, status, reviewNote: reviewNote || '', reviewedAt: new Date().toISOString(), reviewedBy: currentUser?.name || 'ผู้จัดการ' } : s));
  const deleteSubmission = (id) => setSubmissions(prev => prev.filter(s => s.id !== id));

  const saveSupabaseConfig = async (cfg) => {
    setSupabaseConfig(cfg);
    await storage.set('supabaseConfig', JSON.stringify(cfg));
    if (cfg.url && cfg.anonKey) { setDataSource('supabase'); await storage.set('dataSource', 'supabase'); }
  };

  const testConnection = async (cfg) => {
    if (!cfg.url || !cfg.anonKey) throw new Error('ใส่ URL และ Anon Key ก่อน');
    const table = cfg.tableName || 'product_price';
    const res = await fetch(`${cfg.url.replace(/\/$/, '')}/rest/v1/${table}?select=count`, {
      headers: { apikey: cfg.anonKey, Authorization: `Bearer ${cfg.anonKey}`, Prefer: 'count=exact' }
    });
    if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
    const range = res.headers.get('content-range') || '';
    const match = range.match(/\/(\d+)/);
    return match ? parseInt(match[1]) : 0;
  };


  if (!loaded) return <div className="min-h-screen bg-[#F6F7F8] flex items-center justify-center"><div className="w-10 h-10 border-4 border-[#E4E6EA] border-t-[#0F172A] rounded-full animate-spin" /></div>;
    // รายงานเปิดดูได้เลย ไม่ต้องเลือกชื่อ/ไม่ต้องเป็นผู้จัดการ
  if (showReport) return (
    <div className="min-h-screen flex flex-col" style={{ background: '#F6F7F8' }}>
      <header className="bg-white border-b border-[#E4E6EA] px-4 py-3 sticky top-0 z-10 shadow-sm">
        <div className="max-w-6xl mx-auto flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <div className="text-white p-2 rounded-lg" style={{ background: '#0F172A' }}><FileSpreadsheet size={20} /></div>
            <div>
              <h1 className="font-bold text-slate-800">รายงาน</h1>
              <p className="text-xs text-slate-500">ดึงข้อมูลและส่งออกไฟล์ — ไม่ต้องเข้าสู่ระบบ</p>
            </div>
          </div>
          <button onClick={() => setShowReport(false)} className="flex items-center gap-1.5 text-xs font-bold text-slate-700 border border-[#E4E6EA] bg-white hover:bg-[#F6F7F8] px-3 py-2 rounded-lg">
            <ArrowRight size={14} className="rotate-180" />กลับหน้าแรก
          </button>
        </div>
      </header>
      <main className="flex-1 max-w-6xl w-full mx-auto p-4"><ReportView /></main>
    </div>
  );

  if (!currentUser) return <LoginScreen onLogin={handleLogin} onOpenReport={() => setShowReport(true)} />;

  const isManager = currentUser.role === 'manager';
  const feature = currentUser.feature || (isManager ? 'recorder' : 'recorder');
  const myEntries = countEntries.filter(e => e.counterId === currentUser.id);
  const isSupabaseReady = !!(supabaseConfig.url && supabaseConfig.anonKey);
  const pendingCount = submissions.filter(s => s.status === 'pending' && (s.featureType||'recorder') === feature).length;

  const FEATURE_LABEL = { recorder: 'Recorder', stock_compare: 'นับ+เปรียบเทียบ', invoice: 'สแกนบิล AI' };
  // ระบบสีตาม Color System.dc.html — Recorder เขียว · นับ+เปรียบเทียบ ม่วง · สแกนบิล ฟ้า
  // ผู้จัดการใช้สีของฟีเจอร์ที่กำลังดูอยู่ (ตามกติกา "หนึ่งหน้าจอเห็นสีฟีเจอร์ได้สีเดียว")
  const FEAT = {
    recorder:      { main: '#35706A', deep: '#2A5A55', soft: '#EAF1F0', line: '#B6D0CC' },
    stock_compare: { main: '#7658C9', deep: '#5F45A8', soft: '#F2EFFA', line: '#D5CAEE' },
    invoice:       { main: '#2F6E90', deep: '#255771', soft: '#EAF0F4', line: '#B9CFDC' },
  };
  const C = FEAT[feature] || FEAT.recorder;
  

  // Nav per role+feature
  const navItems = isManager
    ? feature === 'stock_compare'
      ? [{ id:'staff',label:'พนักงาน',icon:Users },{ id:'dashboard',label:'แดชบอร์ด',icon:Home },{ id:'inbox',label:'รีวิว',icon:Inbox,badge:pendingCount },{ id:'compare',label:'เปรียบเทียบ',icon:ArrowLeftRight },{ id:'report',label:'รายงาน',icon:FileSpreadsheet },{ id:'settings',label:'ตั้งค่า',icon:SettingsIcon }]
      : [{ id:'staff',label:'พนักงาน',icon:Users },{ id:'dashboard',label:'แดชบอร์ด',icon:Home },{ id:'inbox',label:'รีวิว',icon:Inbox,badge:pendingCount },{ id:'report',label:'รายงาน',icon:FileSpreadsheet },{ id:'settings',label:'ตั้งค่า',icon:SettingsIcon }]
    : feature === 'invoice'
      ? [{ id:'invoice',label:'สแกนบิล',icon:Receipt }]
      : feature === 'stock_compare'
        ? [{ id:'count',label:'นับสต็อก',icon:ScanLine },{ id:'review',label:'ตรวจสอบ',icon:ClipboardCheck,badge:myEntries.length },{ id:'my_submissions',label:'ที่ส่งแล้ว',icon:Send },{ id:'compare',label:'เปรียบเทียบ',icon:ArrowLeftRight }]
        : [{ id:'count',label:'นับสต็อก',icon:ScanLine },{ id:'review',label:'ตรวจสอบ',icon:ClipboardCheck,badge:myEntries.length },{ id:'my_submissions',label:'ที่ส่งแล้ว',icon:Send }];

  return (
    <div className="min-h-screen flex flex-col" style={{ background: '#F6F7F8' }}>
      <header className="bg-white border-b border-[#E4E6EA] px-4 py-3 sticky top-0 z-10 shadow-sm">
        <div className="max-w-6xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-2 min-w-0">
            {view !== navItems[0]?.id ? (
              <button onClick={() => setView(navItems[0].id)} title="กลับหน้าแรก"
                className="shrink-0 rounded-lg flex items-center justify-center border bg-white"
                style={{ width: 38, height: 38, borderColor: '#E4E6EA', color: C.main }}>
                <ArrowRight size={17} className="rotate-180" />
              </button>
            ) : (
              <div className="text-white p-2 rounded-lg shrink-0" style={{ background: C.main }}><Package size={20} /></div>
            )}
            <div>
              <h1 className="font-bold text-slate-800">KUUHOO</h1>
              <p className="text-xs text-slate-500">{isManager ? 'ผู้จัดการ' : 'พนักงาน'} • {currentUser.name} • {FEATURE_LABEL[feature] || feature}</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {isSupabaseReady && <div className={`flex items-center gap-1 px-2 py-1 rounded-full text-[10px] font-medium ${connectionStatus === 'ok' ? 'bg-[#EAF1F0] text-[#2A5A55]' : connectionStatus === 'error' ? 'bg-[#FEF2F2] text-[#B91C1C]' : 'bg-[#F6F7F8] text-slate-500'}`}><Cloud size={10} />Supabase</div>}
            {!isManager
              ? <button onClick={handleLogout} className="flex items-center gap-1.5 text-xs font-bold text-slate-700 border border-[#E4E6EA] bg-white hover:bg-[#F6F7F8] px-3 py-2 rounded-lg"><User size={14} />ไม่ใช่ฉัน</button>
              : <button onClick={handleLogout} className="flex items-center gap-1 text-xs text-slate-500 hover:text-slate-700 px-2 py-1 hover:bg-[#F6F7F8] rounded"><LogOut size={14} />ออก</button>}
          </div>
        </div>
      </header>

      {!isManager && (
        <div className="px-4 py-2 border-b" style={{ background: C.soft, borderColor: C.line }}>
          <div className="max-w-6xl mx-auto flex items-center gap-2 text-xs">
            <span className="font-bold" style={{ color: C.deep }}>
              กำลังทำงานในชื่อ
            </span>
            <span className="font-bold text-slate-800">{currentUser.name}</span>
            <span className="text-slate-400">·</span>
            <span className="text-slate-500">ถ้าไม่ใช่คุณ กด “ไม่ใช่ฉัน” มุมขวาบน</span>
          </div>
        </div>
      )}

      <main className="flex-1 max-w-6xl w-full mx-auto p-4" style={{ paddingBottom: 'calc(76px + env(safe-area-inset-bottom))' }}>
        {/* Counter - stock / stock_compare */}
        {!isManager && feature !== 'invoice' && view === 'count' && <CounterCountView entries={myEntries} addEntry={addCountEntry} deleteEntry={deleteCountEntry} checkBarcode={checkBarcode} setView={setView} products={products} isSupabaseReady={isSupabaseReady} connectionStatus={connectionStatus} countDate={countDate} setCountDate={setCountDate} draft={countDraft} updateDraft={updateDraft} />}
        {!isManager && feature !== 'invoice' && view === 'review' && <CounterReviewView entries={myEntries} setView={setView} submitForReview={submitForReview} clearMyEntries={clearMyEntries} currentUser={currentUser} pickedAt={pickedAt} />}
        {!isManager && feature !== 'invoice' && view === 'my_submissions' && <MySubmissionsView submissions={submissions.filter(s => s.counterId === currentUser.id && (s.featureType||'recorder') === feature)} setView={setView} />}
        {!isManager && feature === 'stock_compare' && view === 'compare' && <CompareStockView submissions={submissions.filter(s => s.counterId === currentUser.id && (s.featureType||'stock_compare') === 'stock_compare')} supabaseConfig={supabaseConfig} compareState={compareState} setCompareState={setCompareState} />}
        {/* Counter - invoice */}
        {!isManager && feature === 'invoice' && <InvoiceScannerModule supabaseConfig={supabaseConfig} />}
        {/* Manager */}
        {isManager && view === 'dashboard' && <Dashboard submissions={submissions.filter(s=>(s.featureType||'recorder')===feature)} products={products} setView={setView} isSupabaseReady={isSupabaseReady} lastSyncAt={lastSyncAt} pendingCount={pendingCount} />}
        {isManager && view === 'inbox' && <ManagerInboxView submissions={submissions.filter(s=>(s.featureType||'recorder')===feature)} onReview={reviewSubmission} onDelete={deleteSubmission} feature={feature} />}
        {isManager && feature === 'stock_compare' && view === 'compare' && <CompareStockView submissions={submissions.filter(s=>(s.featureType||'stock_compare')==='stock_compare')} supabaseConfig={supabaseConfig} compareState={compareState} setCompareState={setCompareState} />}
        {isManager && view === 'staff' && <StaffAdminView />}
        {isManager && view === 'report' && <ReportView />}
        {isManager && view === 'settings' && <SettingsView config={supabaseConfig} onSave={saveSupabaseConfig} onTestConnection={testConnection} dataSource={dataSource} lastSyncAt={lastSyncAt} productCount={products.length} />}
      </main>

      <nav className="fixed bottom-0 left-0 right-0 bg-white border-t border-[#E4E6EA]" style={{ paddingBottom: 'env(safe-area-inset-bottom)' }}>
        <div className="max-w-6xl mx-auto grid" style={{ gridTemplateColumns: `repeat(${navItems.length}, minmax(0, 1fr))` }}>
          {navItems.map(item => {
            const Icon = item.icon; const active = view === item.id;
            return (
              <button key={item.id} onClick={() => setView(item.id)}
                className="relative flex flex-col items-center justify-center gap-1 text-[10px] leading-none"
                style={{ minHeight: 56, paddingTop: 8, paddingBottom: 8, ...(active ? { color: C.main, background: C.soft } : { color: '#64748B', background: '#fff' }) }}>
                <Icon size={18} />
                <span className="font-semibold text-center px-0.5" style={{ fontSize: navItems.length > 4 ? 9.5 : 10.5 }}>{item.label}</span>
                {item.badge > 0 && (
                  <span className="absolute text-white text-[9px] font-bold rounded-full text-center"
                    style={{ top: 6, left: '50%', marginLeft: 6, minWidth: 15, padding: '0 3px', background: '#B91C1C' }}>{item.badge}</span>
                )}
              </button>
            );
          })}
        </div>
      </nav>
    </div>
  );
}

function LoginScreen({ onLogin, onOpenReport }) {
  const [role, setRole] = useState(null);
  const [feature, setFeature] = useState(null);
  const [name, setName] = useState('');
  const [pin, setPin] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const COUNTER_FEATURES = [
    { id: 'recorder',      label: 'Recorder',            desc: 'บันทึกและส่งข้อมูลให้ผู้จัดการ',              icon: ScanLine,      accent: 'recorder' },
    { id: 'stock_compare', label: 'นับ + เปรียบเทียบ',  desc: 'นับสต็อกและเปรียบเทียบกับยอดในระบบ',       icon: ArrowLeftRight, accent: 'compare'    },
    { id: 'invoice',       label: 'สแกนบิล AI',          desc: 'สแกนใบกำกับภาษีด้วย Claude AI',            icon: Receipt,       accent: 'invoice'  },
  ];
  const MANAGER_FEATURES = [
    { id: 'recorder',      label: 'Recorder',            desc: 'รีวิว อนุมัติ และจัดการผลการบันทึก',        icon: ScanLine,      accent: 'indigo'  },
    { id: 'stock_compare', label: 'นับ + เปรียบเทียบ',  desc: 'รีวิว อนุมัติ และเปรียบเทียบสต็อก',        icon: ArrowLeftRight, accent: 'indigo'  },
  ];

  const accentClass = (a, type) => {
    const map = {
      recorder: { border: 'hover:border-[#35706A] hover:bg-[#EAF1F0]', icon: 'bg-[#EAF1F0] text-[#2A5A55]', btn: 'bg-[#35706A] hover:bg-[#2A5A55]' },
      compare:  { border: 'hover:border-[#7658C9] hover:bg-[#F2EFFA]', icon: 'bg-[#F2EFFA] text-[#5F45A8]', btn: 'bg-[#7658C9] hover:bg-[#5F45A8]' },
      invoice:  { border: 'hover:border-[#2F6E90] hover:bg-[#EAF0F4]', icon: 'bg-[#EAF0F4] text-[#255771]', btn: 'bg-[#2F6E90] hover:bg-[#255771]' },
      indigo:   { border: 'hover:border-[#0F172A] hover:bg-[#F6F7F8]', icon: 'bg-[#F6F7F8] text-[#0F172A]', btn: 'bg-[#0F172A] hover:bg-[#0F172A]' },
    };
    return map[a]?.[type] || map.indigo[type];
  };

  const handleLogin = async () => {
    if (!name.trim()) return setError('กรุณาใส่ชื่อ');
    if (role === 'manager') {
      if (!pin) return setError('กรุณาใส่ PIN');
      setLoading(true); setError('');
      try {
        const r = await fetch('/api/auth', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ pin }),
        });
        const text = await r.text();
        let d;
        try { d = JSON.parse(text); } catch {
          setError('Server error — ตรวจสอบ Vercel logs: ' + text.slice(0, 100));
          setLoading(false); return;
        }
        if (!d.ok) { setError(d.msg || 'PIN ไม่ถูกต้อง'); setLoading(false); return; }
      } catch (e) { setError('เชื่อมต่อ server ไม่ได้: ' + e.message); setLoading(false); return; }
      setLoading(false);
    }
    const stableId = `${role}_${feature}_${name.trim().toLowerCase().replace(/\s+/g,'_')}`;
    onLogin({ id: stableId, name: name.trim(), role, feature, loginAt: new Date().toISOString() });
  };

  const features = role === 'manager' ? MANAGER_FEATURES : COUNTER_FEATURES;
  const selectedFeature = features.find(f => f.id === feature);
  const btnColor = role === 'manager' ? accentClass('indigo','btn') : selectedFeature ? accentClass(selectedFeature.accent,'btn') : 'bg-slate-600 hover:bg-slate-700';

  return (
    <div className="min-h-screen bg-[#F6F7F8] flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-xl w-full max-w-md p-6 space-y-5">
        <div className="text-center">
          <div className="bg-[#0F172A] text-white p-3 rounded-2xl inline-block mb-3"><Package size={28}/></div>
          <h1 className="text-2xl font-bold text-slate-800">KUUHOO</h1>
          <p className="text-sm text-slate-500">เข้าสู่ระบบ</p>
        </div>

        {/* Step 1: Role */}
        {!role && (
          <div className="space-y-3">
            <div className="text-sm font-semibold text-slate-700">เลือกบทบาท</div>
            <button onClick={() => setRole('counter')} className="w-full p-4 border-2 border-[#E4E6EA] hover:border-[#35706A] hover:bg-[#EAF1F0] rounded-xl transition-colors text-left flex items-center gap-3">
              <div className="bg-[#EAF1F0] text-[#2A5A55] p-3 rounded-lg"><User size={24}/></div>
              <div><div className="font-semibold text-slate-800">พนักงาน</div><div className="text-xs text-slate-500">นับสต็อก / สแกนบิล</div></div>
            </button>
            <button onClick={() => setRole('manager')} className="w-full p-4 border-2 border-[#E4E6EA] hover:border-[#0F172A] hover:bg-[#F6F7F8] rounded-xl transition-colors text-left flex items-center gap-3">
              <div className="bg-[#F6F7F8] text-[#0F172A] p-3 rounded-lg"><Shield size={24}/></div>
              <div><div className="font-semibold text-slate-800">ผู้จัดการ</div><div className="text-xs text-slate-500">รีวิวและอนุมัติผลการนับ</div></div>
            </button>
          <button onClick={onOpenReport} className="w-full flex items-center gap-3 p-3 rounded-xl border border-[#E4E6EA] bg-white hover:bg-[#F6F7F8] text-left">
            <div className="p-2 rounded-lg bg-[#F6F7F8] text-[#0F172A]"><FileSpreadsheet size={18} /></div>
            <div className="flex-1 min-w-0">
              <div className="text-sm font-bold text-slate-800">ดูรายงาน</div>
              <div className="text-[11.5px] text-slate-500">ไม่ต้องเลือกชื่อ — เข้าดูและส่งออกไฟล์ได้เลย</div>
            </div>
            <ArrowRight size={16} className="text-slate-400 shrink-0" />
          </button>
          </div>
        )}

        {/* Step 2: Feature */}
        {role && !feature && (
          <div className="space-y-3">
            <div className="flex items-center gap-2">
              <button onClick={() => { setRole(null); setError(''); }} className="text-xs text-slate-400 hover:text-slate-600">← กลับ</button>
              <div className="text-sm font-semibold text-slate-700">เลือก Feature</div>
            </div>
            <div className={`rounded-lg px-3 py-2 flex items-center gap-2 text-xs font-medium ${role === 'manager' ? 'bg-[#F6F7F8] text-[#0F172A]' : 'bg-[#EAF1F0] text-[#2A5A55]'}`}>
              {role === 'manager' ? <Shield size={13}/> : <User size={13}/>}
              {role === 'manager' ? 'ผู้จัดการ' : 'พนักงาน'}
            </div>
            {features.map(f => {
              const Icon = f.icon;
              return (
                <button key={f.id} onClick={() => setFeature(f.id)}
                  className={`w-full p-4 border-2 border-[#E4E6EA] ${accentClass(f.accent,'border')} rounded-xl transition-colors text-left flex items-center gap-3`}>
                  <div className={`${accentClass(f.accent,'icon')} p-3 rounded-lg`}><Icon size={22}/></div>
                  <div><div className="font-semibold text-slate-800">{f.label}</div><div className="text-xs text-slate-500">{f.desc}</div></div>
                </button>
              );
            })}
          </div>
        )}

        {/* Step 3: Name + PIN */}
        {role && feature && (
          <div className="space-y-3">
            <div className="flex items-center gap-2">
              <button onClick={() => { setFeature(null); setError(''); setPin(''); }} className="text-xs text-slate-400 hover:text-slate-600">← กลับ</button>
              <div className="text-sm font-semibold text-slate-700">ข้อมูลผู้ใช้</div>
            </div>
            <div className="flex gap-2">
              <div className={`rounded-lg px-3 py-2 flex items-center gap-1.5 text-xs font-medium ${role === 'manager' ? 'bg-[#F6F7F8] text-[#0F172A]' : 'bg-[#EAF1F0] text-[#2A5A55]'}`}>
                {role === 'manager' ? <Shield size={12}/> : <User size={12}/>}
                {role === 'manager' ? 'ผู้จัดการ' : 'พนักงาน'}
              </div>
              <div className={`rounded-lg px-3 py-2 flex items-center gap-1.5 text-xs font-medium ${accentClass(selectedFeature?.accent || 'indigo','icon').replace('bg-','').includes('emerald') ? 'bg-[#EAF1F0] text-[#2A5A55]' : selectedFeature?.accent === 'teal' ? 'bg-[#F2EFFA] text-[#5F45A8]' : selectedFeature?.accent === 'violet' ? 'bg-[#EAF0F4] text-[#255771]' : 'bg-[#F6F7F8] text-[#0F172A]'}`}>
                {selectedFeature && <selectedFeature.icon size={12}/>}
                {selectedFeature?.label}
              </div>
            </div>
            <div>
              <label className="text-sm font-medium text-slate-700 mb-1 block">ชื่อของคุณ</label>
              <input type="text" value={name} onChange={e => { setName(e.target.value); setError(''); }} placeholder="เช่น สมหญิง" className="w-full px-3 py-2.5 border border-[#E4E6EA] rounded-lg focus:ring-2 focus:ring-[#0F172A] outline-none" autoFocus/>
            </div>
            {role === 'manager' && (
              <div>
                <label className="text-sm font-medium text-slate-700 mb-1 block"><Lock size={12} className="inline mr-1"/>รหัส PIN</label>
                <input type="password" value={pin} onChange={e => { setPin(e.target.value); setError(''); }} placeholder="••••" maxLength={4} className="w-full px-3 py-2.5 border border-[#E4E6EA] rounded-lg focus:ring-2 focus:ring-[#0F172A] outline-none font-mono text-lg tracking-widest text-center"/>
              </div>
            )}
            {error && <div className="bg-[#FEF2F2] border border-[#FECACA] text-[#B91C1C] text-sm rounded-lg p-2">{error}</div>}
            <button onClick={handleLogin} disabled={loading} className={`w-full py-3 rounded-lg font-semibold text-white disabled:opacity-60 ${btnColor}`}>
              {loading ? 'กำลังตรวจสอบ...' : 'เข้าสู่ระบบ'}
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

function EntryRow({ e, deleteEntry, highlight }) {
  return (
    <div className={`flex items-center gap-2 p-3 border-b last:border-0 ${highlight ? 'border-[#FFFBEB] bg-[#FFFBEB]/30' : 'border-[#F6F7F8]'}`}>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-1.5 flex-wrap">
          <span className="text-sm font-medium text-slate-800 truncate">{e.productName}</span>
          {e.notFound && <span className="text-[9px] bg-[#B45309] text-white px-1.5 py-0.5 rounded-full flex-shrink-0">ไม่มีในระบบ</span>}
        </div>
        <div className="text-xs text-slate-500 font-mono flex items-center gap-1.5">
          {e.barcode}
          {e.location && <span className="text-[9px] bg-[#F6F7F8] text-slate-500 px-1.5 py-0.5 rounded">{e.location}</span>}
        </div>
      </div>
      <div className="text-right mr-1">
        <div className="font-bold text-slate-800">{e.qty}</div>
        <div className={`text-xs font-mono font-semibold ${e.notFound ? 'text-[#B45309]' : 'text-[#35706A]'}`}>{new Date(e.timestamp).toLocaleTimeString('th-TH',{hour:'2-digit',minute:'2-digit',second:'2-digit'})}</div>
        <div className="text-[10px] text-slate-400">{new Date(e.timestamp).toLocaleDateString('th-TH',{day:'2-digit',month:'short'})}</div>
      </div>
      <button onClick={() => deleteEntry(e.id)} className="p-1.5 hover:bg-[#FEF2F2] text-[#B91C1C] rounded"><Trash2 size={14}/></button>
    </div>
  );
}

function CounterCountView({ entries, addEntry, deleteEntry, checkBarcode, setView, products, isSupabaseReady, connectionStatus, countDate, setCountDate, draft, updateDraft }) {
  const [location, setLocation] = useState('');
  const [checking, setChecking] = useState(false);
  const [scanMode, setScanMode] = useState(false);
  const qtyInputRef = useRef(null); const barcodeInputRef = useRef(null);
  const { barcode, qty, checkResult, error } = draft;
  const setBarcode = v => updateDraft({ barcode: v });
  const setQty = v => updateDraft({ qty: v });
  const setCheckResult = v => updateDraft({ checkResult: v });
  const setError = v => updateDraft({ error: v });

  const handleCheck = async (raw) => {
    const code = (raw || barcode).trim(); if (!code) return;
    setChecking(true); setError(''); setCheckResult(null);
    try { const p = await checkBarcode(code); if (p) { setCheckResult(p); setBarcode(code); setTimeout(() => qtyInputRef.current?.focus(), 100); } else { setError(`ไม่พบรหัส "${code}"`); setBarcode(code); } }
    catch (e) { setError(e.message || 'เกิดข้อผิดพลาด'); }
    setChecking(false);
  };

  const handleAdd = () => {
    if (!checkResult || !qty || parseInt(qty) <= 0) return;
    addEntry({ barcode: checkResult.barcode, productName: checkResult.name, productId: checkResult.id, unit: checkResult.unit, price: checkResult.price || 0, cost: checkResult.cost || 0, qty: parseInt(qty), countDate, location: location.trim() });
    updateDraft({ barcode: '', qty: '', checkResult: null, error: '' });
    setTimeout(() => barcodeInputRef.current?.focus(), 100);
  };

  const handleAddManually = () => {
    if (!barcode.trim() || !qty || parseInt(qty) <= 0) return;
    addEntry({ barcode: barcode.trim(), productName: '(ไม่พบในระบบ)', productId: '', unit: '', qty: parseInt(qty), countDate, notFound: true, location: location.trim() });
    updateDraft({ barcode: '', qty: '', checkResult: null, error: '' });
    setTimeout(() => barcodeInputRef.current?.focus(), 100);
  };

  const totalQty = entries.reduce((s, e) => s + e.qty, 0);
  const uniqueBarcodes = new Set(entries.map(e => e.barcode)).size;
  const recent = [...entries].reverse();

  return (
    <div className="space-y-3">
      {/* รอบนับ + สถานะเซิร์ฟเวอร์ */}
      <div className="bg-white border rounded-xl overflow-hidden" style={{ borderColor: '#e4e6ea' }}>
        <div className="flex items-center gap-2 px-3 py-2.5 border-b" style={{ borderColor: '#eef0f3' }}>
          <div className="flex-1 min-w-0">
            <div className="text-[10px] font-bold tracking-wide" style={{ color: '#35706a' }}>รอบนับ</div>
            <input type="date" value={countDate} onChange={e => setCountDate(e.target.value)}
              className="mt-0.5 text-sm font-bold text-slate-800 bg-transparent outline-none" style={{ fontFamily: "'IBM Plex Mono', monospace" }} />
          </div>
          <div className={`shrink-0 text-[10px] font-semibold px-2 py-1 rounded-md ${!isSupabaseReady ? 'bg-[#FEF2F2] text-[#B91C1C]' : connectionStatus === 'error' ? 'bg-[#FFFBEB] text-[#B45309]' : 'bg-[#F6F7F8] text-slate-500'}`}>
            {!isSupabaseReady ? 'ยังไม่ตั้งค่าเซิร์ฟเวอร์' : connectionStatus === 'error' ? 'ต่อเซิร์ฟเวอร์ไม่ได้' : 'เซิร์ฟเวอร์พร้อม'}
          </div>
        </div>
        <div className="grid grid-cols-2 divide-x" style={{ borderColor: '#eef0f3' }}>
          <div className="px-3 py-2.5">
            <div className="text-[10.5px] text-slate-500">บาร์โค้ด</div>
            <div className="text-xl font-bold text-slate-800" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{uniqueBarcodes}</div>
          </div>
          <div className="px-3 py-2.5">
            <div className="text-[10.5px] text-slate-500">รวมจำนวน</div>
            <div className="text-xl font-bold text-slate-800" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{totalQty.toLocaleString('th-TH')}</div>
          </div>
        </div>
      </div>

      {/* ปุ่มสแกน */}
      <button onClick={() => setScanMode(true)}
        className="w-full text-white font-bold text-base flex items-center justify-center gap-2 rounded-xl"
        style={{ minHeight: 54, background: '#35706a', boxShadow: '0 2px 0 #2a5a55' }}>
        <Camera size={19} />{entries.length ? 'สแกนชิ้นต่อไป' : 'สแกนบาร์โค้ด'}
      </button>

      {/* พิมพ์เอง */}
      <div className="bg-white border rounded-xl p-3 space-y-2.5" style={{ borderColor: '#e2e8f0' }}>
        <div className="text-[11.5px] text-slate-500">หรือพิมพ์รหัสสินค้า / บาร์โค้ด</div>
        <div className="flex gap-2">
          <input ref={barcodeInputRef} type="text" value={barcode}
            onChange={e => { setBarcode(e.target.value); setCheckResult(null); setError(''); }}
            onKeyDown={e => e.key === 'Enter' && handleCheck()}
            placeholder="8850999320014"
            className="flex-1 px-3 py-2.5 border rounded-lg outline-none focus:border-[#7658C9] text-[15px]"
            style={{ borderColor: '#e2e8f0', fontFamily: "'IBM Plex Mono', monospace" }} />
          <button onClick={() => handleCheck()} disabled={!barcode.trim() || checking}
            className="px-4 rounded-lg text-sm font-semibold text-white disabled:opacity-40 flex items-center gap-1"
            style={{ background: '#0f172a' }}>
            {checking ? <RefreshCw size={14} className="animate-spin" /> : <Search size={14} />}ตรวจ
          </button>
        </div>
        <div>
          <div className="text-[11px] text-slate-500 mb-1 flex items-center gap-1"><MapPin size={11} />ตำแหน่ง (ไม่บังคับ)</div>
          <input type="text" value={location} onChange={e => setLocation(e.target.value)} placeholder="A-1"
            className="w-full px-3 py-2 border rounded-lg outline-none focus:border-[#7658C9] text-[13px]"
            style={{ borderColor: '#e2e8f0' }} />
        </div>
      </div>

      {/* การ์ดสินค้าที่พบ */}
      {checkResult && (
        <div className="bg-white rounded-xl overflow-hidden" style={{ border: '1.5px solid #b6d0cc' }}>
          <div className="px-3 py-2.5 border-b" style={{ background: '#eaf1f0', borderColor: '#d5e5e2' }}>
            <div className="text-[10.5px] font-semibold" style={{ color: '#2a5a55' }}>พบสินค้าในฐานข้อมูล</div>
            <div className="text-[15px] font-bold text-slate-800 mt-0.5">{checkResult.name}</div>
            <div className="flex gap-2.5 mt-1 text-[11px] text-slate-500">
              <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{checkResult.productCode}</span>
              {checkResult.unit && <span>หน่วย: {checkResult.unit}</span>}
              {checkResult.price ? <span>฿{checkResult.price}</span> : null}
            </div>
          </div>
          <div className="p-3 space-y-3">
            <div>
              <div className="text-xs font-semibold text-slate-600 mb-1.5">จำนวนที่นับได้</div>
              <div className="flex items-center gap-2.5">
                <button onClick={() => setQty(String(Math.max(0, (parseInt(qty) || 0) - 1)))}
                  className="shrink-0 rounded-xl border flex items-center justify-center text-2xl text-slate-700"
                  style={{ width: 52, height: 52, borderColor: '#e2e8f0', background: '#f8fafc' }}>−</button>
                <input ref={qtyInputRef} type="number" inputMode="numeric" value={qty}
                  onChange={e => setQty(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && handleAdd()} placeholder="0"
                  className="flex-1 text-center font-bold text-slate-800 border-0 outline-none"
                  style={{ fontSize: 30, fontFamily: "'IBM Plex Mono', monospace" }} />
                <button onClick={() => setQty(String((parseInt(qty) || 0) + 1))}
                  className="shrink-0 rounded-xl border flex items-center justify-center text-2xl"
                  style={{ width: 52, height: 52, borderColor: '#b6d0cc', background: '#eaf1f0', color: '#2a5a55' }}>+</button>
              </div>
            </div>
            <button onClick={handleAdd} disabled={!qty || parseInt(qty) <= 0}
              className="w-full text-white font-bold text-[15px] rounded-xl disabled:opacity-40"
              style={{ minHeight: 48, background: '#0f172a' }}>เพิ่มเข้ารายการ</button>
          </div>
        </div>
      )}

      {/* ไม่พบในระบบ — นับไว้ก่อน */}
      {error && (
        <div className="bg-white border rounded-xl p-3 space-y-2.5" style={{ borderColor: '#fde68a' }}>
          <div className="flex items-start gap-2 text-[12px] text-[#B45309]">
            <AlertCircle size={14} className="mt-0.5 shrink-0" />
            <div>{error} — ยังไม่มีในระบบ นับไว้ก่อนได้</div>
          </div>
          <div className="flex items-center gap-2.5">
            <button onClick={() => setQty(String(Math.max(0, (parseInt(qty) || 0) - 1)))}
              className="shrink-0 rounded-xl border flex items-center justify-center text-2xl text-slate-700"
              style={{ width: 48, height: 48, borderColor: '#e2e8f0', background: '#f8fafc' }}>−</button>
            <input type="number" inputMode="numeric" value={qty} onChange={e => setQty(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && handleAddManually()} placeholder="0"
              className="flex-1 text-center font-bold text-slate-800 border-0 outline-none"
              style={{ fontSize: 26, fontFamily: "'IBM Plex Mono', monospace" }} />
            <button onClick={() => setQty(String((parseInt(qty) || 0) + 1))}
              className="shrink-0 rounded-xl border flex items-center justify-center text-2xl text-[#B45309]"
              style={{ width: 48, height: 48, borderColor: '#fde68a', background: '#fffbeb' }}>+</button>
          </div>
          <button onClick={handleAddManually} disabled={!qty || parseInt(qty) <= 0}
            className="w-full text-white font-bold text-sm rounded-xl disabled:opacity-40"
            style={{ minHeight: 44, background: '#b45309' }}>นับไว้ก่อน (ยังไม่มีในระบบ)</button>
        </div>
      )}

      {/* ร่างของฉัน */}
      <div className="flex items-baseline justify-between px-0.5 pt-1">
        <div className="text-[13px] font-bold text-slate-700">ร่างของคุณ ({entries.length} ครั้ง)</div>
        <div className="text-[11.5px] text-slate-500">{uniqueBarcodes} บาร์โค้ด</div>
      </div>

      {entries.length === 0 ? (
        <div className="bg-white rounded-xl px-4 py-7 text-center text-[12.5px] text-slate-400"
          style={{ border: '1px dashed #cbd5e1' }}>ยังไม่ได้นับ — เริ่มจากสแกนบาร์โค้ด</div>
      ) : (
        <div className="space-y-2">
          {recent.slice(0, 30).map(e => (
            <div key={e.id} className="bg-white border rounded-xl px-3 py-2.5 flex items-center gap-2.5"
              style={{ borderColor: e.notFound ? '#fde68a' : '#e2e8f0' }}>
              <div className="flex-1 min-w-0">
                <div className="text-[13.5px] font-semibold text-slate-800 truncate">{e.productName}</div>
                <div className="text-[10.5px] text-slate-400 mt-0.5 truncate" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>
                  {e.barcode}{e.location ? ' · ' + e.location : ''}
                </div>
              </div>
              <div className="text-right shrink-0">
                <div className="text-base font-bold text-slate-800" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{e.qty}</div>
                <div className="text-[10.5px] text-slate-400">{e.unit || ''}</div>
              </div>
              <button onClick={() => deleteEntry(e.id)}
                className="shrink-0 rounded-lg flex items-center justify-center text-[#B91C1C]"
                style={{ width: 32, height: 32, background: '#fef2f2' }}><X size={14} /></button>
            </div>
          ))}
          {recent.length > 30 && <div className="text-center text-[11px] text-slate-400 py-1">แสดง 30 รายการล่าสุด จาก {recent.length}</div>}
        </div>
      )}

      {entries.length > 0 && (
        <button onClick={() => setView('review')}
          className="w-full text-white font-bold text-[15.5px] rounded-xl"
          style={{ minHeight: 52, background: '#0f172a' }}>ตรวจสอบ &amp; ส่ง →</button>
      )}

      <div className="bg-white border rounded-xl px-3 py-2.5 flex items-center gap-2" style={{ borderColor: '#e2e8f0' }}>
        <span className="shrink-0 rounded-full" style={{ width: 7, height: 7, background: '#35706a' }}></span>
        <div className="text-[11px] text-slate-500 leading-relaxed">เซฟทันทีที่กด — ปิดแอปหรือเน็ตหลุดก็กลับมานับต่อได้</div>
      </div>

      {scanMode && <ScannerModal products={products} onScan={(code) => { setScanMode(false); setBarcode(code); handleCheck(code); }} onClose={() => setScanMode(false)} />}
    </div>
  );
}

function GroupedRow({ g, highlight, onEditQty }) {
  const [editing, setEditing] = useState(false);
  const [editVal, setEditVal] = useState(String(g.qty));
  const confirmEdit = () => { const v = parseInt(editVal); if (!isNaN(v) && v >= 0) onEditQty && onEditQty(g.barcode, v); setEditing(false); };
  return (
    <div className="px-3 py-2.5 flex items-center gap-2.5" style={highlight ? { background: '#FFFBEB' } : undefined}>
      <div className="flex-1 min-w-0">
        <div className="text-[13.5px] font-semibold text-slate-800 truncate">{g.productName}</div>
        <div className="flex items-center gap-1.5 mt-0.5 flex-wrap">
          <span className="text-[10.5px] text-slate-400 tabular-nums">{g.barcode}</span>
          {g.location && <span className="text-[9.5px] px-1.5 py-0.5 rounded-full text-slate-600" style={{ background: '#F6F7F8' }}>{g.location}</span>}
          {g.scans > 1 && <span className="text-[9.5px] px-1.5 py-0.5 rounded-full text-slate-500" style={{ background: '#F6F7F8' }}>ยิง {g.scans} ครั้ง</span>}
          {g.overridden && <span className="text-[9.5px] px-1.5 py-0.5 rounded-full font-semibold" style={{ background: '#FFFBEB', color: '#B45309' }}>แก้จำนวนแล้ว</span>}
        </div>
      </div>
      {onEditQty && editing ? (
        <div className="flex items-center gap-1 shrink-0">
          <input autoFocus type="number" inputMode="numeric" value={editVal}
            onChange={e => setEditVal(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') confirmEdit(); if (e.key === 'Escape') setEditing(false); }}
            className="w-16 text-center font-bold border rounded-lg outline-none text-[15px] tabular-nums"
            style={{ height: 40, borderColor: '#35706A' }} />
          <button onClick={confirmEdit} className="rounded-lg text-white flex items-center justify-center"
            style={{ width: 40, height: 40, background: '#35706A' }}><Check size={16} /></button>
          <button onClick={() => setEditing(false)} className="rounded-lg text-slate-600 flex items-center justify-center"
            style={{ width: 40, height: 40, background: '#F6F7F8' }}><X size={16} /></button>
        </div>
      ) : (
        <div className="flex items-center gap-1.5 shrink-0">
          <div className="text-lg font-bold tabular-nums" style={{ color: highlight ? '#B45309' : '#0F172A' }}>{g.qty}</div>
          {g.unit && <div className="text-[10.5px] text-slate-400">{g.unit}</div>}
          {onEditQty && (
            <button onClick={() => { setEditVal(String(g.qty)); setEditing(true); }}
              className="rounded-lg flex items-center justify-center text-slate-400 hover:text-slate-700"
              style={{ width: 36, height: 36, background: '#F6F7F8' }}><Edit3 size={14} /></button>
          )}
        </div>
      )}
    </div>
  );
}

function CounterReviewView({ entries, setView, submitForReview, clearMyEntries, currentUser, pickedAt }) {
  const [note, setNote] = useState('');
  const [submitted, setSubmitted] = useState(null);
  const [confirming, setConfirming] = useState(false);
  const [qtyOverrides, setQtyOverrides] = useState({});
  const editQty = (barcode, newQty) => setQtyOverrides(prev => ({ ...prev, [barcode]: newQty }));
  const grouped = useMemo(() => {
    const map = new Map();
    [...entries].sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp)).forEach(e => {
      if (map.has(e.barcode)) { const ex = map.get(e.barcode); ex.qty += e.qty; ex.scans += 1; }
      else map.set(e.barcode, { barcode: e.barcode, productName: e.productName, unit: e.unit||'', price: e.price||0, cost: e.cost||0, qty: e.qty, scans: 1, scannedAt: e.timestamp, notFound: !!e.notFound, location: e.location||'' });
    });
    return Array.from(map.values()).map(g => ({ ...g, qty: qtyOverrides[g.barcode] !== undefined ? qtyOverrides[g.barcode] : g.qty, overridden: qtyOverrides[g.barcode] !== undefined })).sort((a, b) => a.barcode.localeCompare(b.barcode));
  }, [entries, qtyOverrides]);
  const totalItems = grouped.length, totalQty = grouped.reduce((s, g) => s + g.qty, 0);
  const handleSubmit = async () => { if (!confirming) { setConfirming(true); return; } const sub = await submitForReview(grouped, note); clearMyEntries(); setSubmitted(sub); };
  if (submitted) return (
    <div className="space-y-3">
      <div className="bg-white rounded-2xl overflow-hidden border" style={{ borderColor: '#B6D0CC' }}>
        <div className="px-4 py-5 text-center" style={{ background: '#EAF1F0' }}>
          <div className="text-white p-3 rounded-full inline-block mb-2" style={{ background: '#35706A' }}><Send size={26} /></div>
          <h2 className="text-lg font-bold" style={{ color: '#2A5A55' }}>ส่งเรียบร้อยแล้ว</h2>
          <p className="text-[12px] mt-1" style={{ color: '#2A5A55' }}>ผู้จัดการจะรีวิวและแจ้งผลกลับ</p>
        </div>
        <div className="p-3 space-y-2.5">
          {submitted.docNo && (
            <div className="rounded-xl px-3 py-2.5 text-center" style={{ background: '#F6F7F8' }}>
              <div className="text-[10px] text-slate-500">เลขที่เอกสาร</div>
              <div className="text-lg font-bold text-slate-800 tabular-nums mt-0.5">{submitted.docNo}</div>
            </div>
          )}
          <div className="flex gap-2">
            <div className="flex-1 rounded-xl px-2.5 py-2.5 text-center" style={{ background: '#F6F7F8' }}>
              <div className="text-[10px] text-slate-500">บาร์โค้ด</div>
              <div className="text-lg font-bold text-slate-800 tabular-nums">{submitted.itemCount}</div>
            </div>
            <div className="flex-1 rounded-xl px-2.5 py-2.5 text-center" style={{ background: '#F6F7F8' }}>
              <div className="text-[10px] text-slate-500">รวมจำนวน</div>
              <div className="text-lg font-bold text-slate-800 tabular-nums">{submitted.totalQty.toLocaleString('th-TH')}</div>
            </div>
          </div>
          <div className="text-[11px] text-slate-500 leading-relaxed text-center">ใบนี้แก้ไม่ได้แล้ว — ถ้ามีจุดผิดผู้จัดการจะส่งกลับมาให้นับใหม่</div>
        </div>
      </div>
      <button onClick={() => setView('count')}
        className="w-full text-white font-bold text-[15px] rounded-xl flex items-center justify-center gap-2"
        style={{ minHeight: 52, background: '#35706A', boxShadow: '0 2px 0 #2A5A55' }}><ScanLine size={18} />เริ่มนับรอบใหม่</button>
      <button onClick={() => setView('my_submissions')}
        className="w-full border bg-white hover:bg-[#F6F7F8] rounded-xl font-semibold text-slate-700 text-[13.5px]"
        style={{ minHeight: 46, borderColor: '#E4E6EA' }}>ดูสถานะใบที่ส่งแล้ว →</button>
    </div>
  );
  if (entries.length === 0) return (
    <div className="space-y-3">
      <h2 className="text-xl font-bold text-slate-800">ตรวจสอบและส่ง</h2>
      <div className="bg-white rounded-xl px-4 py-9 text-center" style={{ border: '1px dashed #CBD5E1' }}>
        <ClipboardCheck className="mx-auto text-[#E4E6EA] mb-2" size={38} />
        <div className="text-[13px] text-slate-500 mb-3">ยังไม่ได้นับสินค้า</div>
        <button onClick={() => setView('count')}
          className="text-white px-4 py-2.5 rounded-xl text-[13px] font-bold" style={{ background: '#35706A' }}>ไปหน้านับสต็อก</button>
      </div>
    </div>
  );
  const found = grouped.filter(g => !g.notFound), missing = grouped.filter(g => g.notFound);
  return (
    <div className="space-y-3">
      <div>
        <h2 className="text-xl font-bold text-slate-800">ตรวจสอบและส่ง</h2>
        <p className="text-[12px] text-slate-500 mt-0.5">บาร์โค้ดเดียวกันรวมเป็นบรรทัดเดียว — แก้จำนวนได้ที่นี่</p>
      </div>

      <div className="bg-white border rounded-xl grid grid-cols-2 divide-x" style={{ borderColor: '#E4E6EA' }}>
        <div className="px-3 py-2.5">
          <div className="text-[10.5px] text-slate-500">บาร์โค้ด</div>
          <div className="text-xl font-bold text-slate-800 tabular-nums">{totalItems}</div>
        </div>
        <div className="px-3 py-2.5" style={{ borderColor: '#EEF0F3' }}>
          <div className="text-[10.5px] text-slate-500">รวมจำนวน</div>
          <div className="text-xl font-bold tabular-nums" style={{ color: '#35706A' }}>{totalQty.toLocaleString('th-TH')}</div>
        </div>
      </div>

      {found.length > 0 && (
        <div className="bg-white rounded-xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
          <div className="px-3 py-2.5 border-b flex items-center gap-2" style={{ background: '#EAF1F0', borderColor: '#D5E5E2' }}>
            <CheckCircle2 size={13} style={{ color: '#2A5A55' }} className="shrink-0" />
            <span className="text-[11.5px] font-bold" style={{ color: '#2A5A55' }}>พบในระบบ · {found.length} รายการ</span>
            <button onClick={() => setView('count')} className="ml-auto text-[11px] font-semibold shrink-0" style={{ color: '#35706A' }}>← กลับไปนับ</button>
          </div>
          <div className="divide-y max-h-64 overflow-y-auto" style={{ borderColor: '#F6F7F8' }}>
            {found.map(g => <GroupedRow key={g.barcode} g={g} onEditQty={editQty} />)}
          </div>
        </div>
      )}

      {missing.length > 0 && (
        <div className="bg-white rounded-xl border overflow-hidden" style={{ borderColor: '#FDE68A' }}>
          <div className="px-3 py-2.5 border-b flex items-center gap-2" style={{ background: '#FFFBEB', borderColor: '#FDE68A' }}>
            <AlertCircle size={13} style={{ color: '#B45309' }} className="shrink-0" />
            <span className="text-[11.5px] font-bold" style={{ color: '#B45309' }}>ไม่พบในระบบ · {missing.length} รายการ</span>
          </div>
          <div className="divide-y max-h-56 overflow-y-auto" style={{ borderColor: '#FFFBEB' }}>
            {missing.map(g => <GroupedRow key={g.barcode} g={g} highlight onEditQty={editQty} />)}
          </div>
          <div className="px-3 py-2 text-[10.5px] leading-relaxed" style={{ background: '#FFFBEB', color: '#B45309' }}>
            ส่งไปได้ — ผู้จัดการจะเป็นคนจับคู่รหัสสินค้าให้
          </div>
        </div>
      )}

      <div className="bg-white rounded-xl border p-3" style={{ borderColor: '#E4E6EA' }}>
        <label className="text-[12px] font-semibold text-slate-700 mb-1.5 block">หมายเหตุถึงผู้จัดการ (ไม่บังคับ)</label>
        <textarea value={note} onChange={e => setNote(e.target.value)} rows={2}
          placeholder="เช่น นับโซน A ชั้น 1-3 เสร็จแล้ว"
          className="w-full px-3 py-2 border rounded-lg outline-none focus:border-[#35706A] text-[13px] resize-none"
          style={{ borderColor: '#E2E8F0' }} />
      </div>

      {!confirming ? (
        <button onClick={handleSubmit}
          className="w-full text-white font-bold text-[15.5px] rounded-xl flex items-center justify-center gap-2"
          style={{ minHeight: 54, background: '#35706A', boxShadow: '0 2px 0 #2A5A55' }}><Send size={19} />ส่งให้ผู้จัดการรีวิว</button>
      ) : (
        <div className="space-y-3">
          <div className="text-[11px] font-bold tracking-wide text-[#B45309]">ยืนยันก่อนส่ง</div>
          <div className="text-xl font-bold text-slate-800 leading-snug">
            ใบนี้จะออกในชื่อ {currentUser?.name || 'พนักงาน'} — ใช่ไหม
          </div>
          <div className="bg-white border-2 border-[#FDE68A] rounded-2xl p-4">
            <div className="flex items-center gap-3">
              <span className="w-11 h-11 shrink-0 rounded-full bg-[#EAF1F0] text-[#2A5A55] flex items-center justify-center text-lg font-bold">
                {(currentUser?.name || '?').trim().charAt(0)}
              </span>
              <div className="min-w-0 flex-1">
                <div className="text-base font-bold text-slate-800 truncate">{currentUser?.name || 'พนักงาน'}</div>
                <div className="text-[11px] text-slate-400 mt-0.5">
                  {pickedAt ? `เลือกชื่อไว้ ${new Date(pickedAt).toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' })}` : 'เลือกชื่อไว้เมื่อเปิดเครื่อง'}
                </div>
              </div>
            </div>
            <div className="h-px bg-[#F6F7F8] my-3" />
            <div className="flex gap-2">
              <div className="flex-1 bg-[#F6F7F8] rounded-lg p-2.5 text-center">
                <div className="text-[10px] text-slate-500">บาร์โค้ด</div>
                <div className="text-lg font-bold text-slate-800 tabular-nums">{totalItems}</div>
              </div>
              <div className="flex-1 bg-[#F6F7F8] rounded-lg p-2.5 text-center">
                <div className="text-[10px] text-slate-500">รวมหน่วย</div>
                <div className="text-lg font-bold text-slate-800 tabular-nums">{totalQty.toLocaleString('th-TH')}</div>
              </div>
            </div>
          </div>
          <div className="bg-[#FFFBEB] border border-[#FDE68A] rounded-xl p-3 text-[11.5px] text-[#B45309] leading-relaxed">
            เครื่องรวมไม่มี PIN — ถ้าคนก่อนหน้าลืมกด “ไม่ใช่ฉัน” ของที่คุณนับจะเข้าใบเขา จุดนี้คือด่านสุดท้ายที่จับได้
          </div>
          <button onClick={handleSubmit} className="w-full bg-[#35706A] hover:bg-[#2A5A55] text-white py-4 rounded-xl font-bold flex items-center justify-center gap-2 shadow-lg">
            <Send size={18}/>ใช่ ส่งในชื่อ {currentUser?.name || 'พนักงาน'}
          </button>
          <button onClick={()=>setConfirming(false)} className="w-full py-3 border border-[#E4E6EA] bg-white hover:bg-[#F6F7F8] rounded-xl font-medium text-sm text-slate-600">
            ‹ กลับไปแก้
          </button>
        </div>
      )}
    </div>
  );
}

function MySubmissionsView({ submissions, setView }) {
  const [expanded, setExpanded] = useState(null);
  const CFG = {
    pending:  { label: 'รอผู้จัดการรีวิว', soft: '#FFFBEB', line: '#FDE68A', ink: '#B45309', icon: Clock },
    approved: { label: 'อนุมัติแล้ว',      soft: '#F0FDF4', line: '#BBF7D0', ink: '#15803D', icon: ThumbsUp },
    rejected: { label: 'ส่งกลับแก้ไข',      soft: '#FEF2F2', line: '#FECACA', ink: '#B91C1C', icon: ThumbsDown },
  };
  const counts = submissions.reduce((a, s) => { a[s.status || 'pending'] = (a[s.status || 'pending'] || 0) + 1; return a; }, {});

  return (
    <div className="space-y-3">
      <div className="flex items-baseline justify-between">
        <div>
          <h2 className="text-xl font-bold text-slate-800">ใบที่ส่งแล้ว</h2>
          <p className="text-[12px] text-slate-500 mt-0.5">ส่งแล้วแก้ไม่ได้ — ถ้าผิดผู้จัดการจะส่งกลับมา</p>
        </div>
        <div className="text-[11.5px] text-slate-400 shrink-0">{submissions.length} ใบ</div>
      </div>

      {submissions.length > 0 && (
        <div className="grid grid-cols-3 gap-2">
          {['pending', 'approved', 'rejected'].map(k => (
            <div key={k} className="rounded-xl px-2.5 py-2 border" style={{ background: CFG[k].soft, borderColor: CFG[k].line }}>
              <div className="text-[10px] font-semibold leading-tight" style={{ color: CFG[k].ink }}>{CFG[k].label}</div>
              <div className="text-lg font-bold tabular-nums mt-0.5" style={{ color: CFG[k].ink }}>{counts[k] || 0}</div>
            </div>
          ))}
        </div>
      )}

      {submissions.length === 0 ? (
        <div className="bg-white rounded-xl px-4 py-9 text-center" style={{ border: '1px dashed #CBD5E1' }}>
          <Send className="mx-auto text-[#E4E6EA] mb-2" size={34} />
          <div className="text-[13px] text-slate-500 mb-3">ยังไม่มีใบที่ส่ง</div>
          <button onClick={() => setView('count')}
            className="text-white px-4 py-2.5 rounded-xl text-[13px] font-bold" style={{ background: '#35706A' }}>ไปนับสต็อก</button>
        </div>
      ) : (
        <div className="space-y-2">
          {submissions.map(s => {
            const cfg = CFG[s.status] || CFG.pending; const Icon = cfg.icon; const open = expanded === s.id;
            return (
              <div key={s.id} className="bg-white rounded-xl overflow-hidden border" style={{ borderColor: cfg.line }}>
                <div className="px-3 py-2 flex items-center gap-2 border-b" style={{ background: cfg.soft, borderColor: cfg.line }}>
                  <Icon size={14} style={{ color: cfg.ink }} className="shrink-0" />
                  <span className="text-[11.5px] font-bold" style={{ color: cfg.ink }}>{cfg.label}</span>
                  <span className="ml-auto text-[10.5px] tabular-nums text-slate-500 shrink-0">
                    {new Date(s.submittedAt).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' })}
                  </span>
                </div>

                <div className="p-3 space-y-2.5">
                  {s.docNo && <div className="text-[13.5px] font-bold text-slate-800 tabular-nums">{s.docNo}</div>}
                  <div className="flex gap-2">
                    <div className="flex-1 rounded-lg px-2.5 py-2 text-center" style={{ background: '#F6F7F8' }}>
                      <div className="text-[10px] text-slate-500">บาร์โค้ด</div>
                      <div className="text-base font-bold text-slate-800 tabular-nums">{s.itemCount}</div>
                    </div>
                    <div className="flex-1 rounded-lg px-2.5 py-2 text-center" style={{ background: '#F6F7F8' }}>
                      <div className="text-[10px] text-slate-500">รวมจำนวน</div>
                      <div className="text-base font-bold text-slate-800 tabular-nums">{s.totalQty.toLocaleString('th-TH')}</div>
                    </div>
                  </div>

                  {s.note && <div className="text-[11.5px] text-slate-500 leading-relaxed">หมายเหตุที่ส่งไป: “{s.note}”</div>}

                  {s.status !== 'pending' && s.reviewNote && (
                    <div className="rounded-lg p-2.5 border" style={{ background: cfg.soft, borderColor: cfg.line }}>
                      <div className="text-[10px] font-bold" style={{ color: cfg.ink }}>{s.reviewedBy || 'ผู้จัดการ'}</div>
                      <div className="text-[11.5px] mt-0.5 leading-relaxed" style={{ color: cfg.ink }}>{s.reviewNote}</div>
                    </div>
                  )}

                  {s.status === 'rejected' && (
                    <button onClick={() => setView('count')}
                      className="w-full text-white font-bold text-[13.5px] rounded-xl" style={{ minHeight: 44, background: '#B91C1C' }}>
                      นับใหม่ตามที่แจ้ง
                    </button>
                  )}

                  <button onClick={() => setExpanded(open ? null : s.id)}
                    className="w-full pt-2 border-t text-[11.5px] font-semibold text-slate-500 hover:text-slate-700"
                    style={{ borderColor: '#F6F7F8' }}>
                    {open ? 'ซ่อนรายการที่นับ' : `ดูรายการที่นับ (${s.data?.length || 0})`}
                  </button>

                  {open && (
                    <div className="rounded-lg divide-y max-h-56 overflow-y-auto" style={{ background: '#F6F7F8', borderColor: '#E4E6EA' }}>
                      {s.data.map((d, i) => (
                        <div key={i} className="flex items-center gap-2 px-2.5 py-2" style={{ borderColor: '#E4E6EA' }}>
                          <div className="flex-1 min-w-0">
                            <div className="text-[12px] font-semibold text-slate-800 truncate">{d.productName}</div>
                            <div className="text-[10px] text-slate-400 tabular-nums truncate">{d.barcode}{d.location ? ' · ' + d.location : ''}</div>
                          </div>
                          <div className="text-[13px] font-bold text-slate-800 tabular-nums shrink-0">{d.qty}</div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function PDFDownloadButton({ sub }) {
  return <button onClick={() => openPDFPrint(sub)} className="flex items-center gap-1 px-2 py-1.5 text-xs bg-[#FEF2F2] hover:bg-[#FECACA] rounded text-[#B91C1C] font-medium"><Download size={12}/>PDF</button>;
}

function ManagerInboxView({ submissions, onReview, onDelete, feature }) {
  const [selected, setSelected] = useState(null);
  const [reviewNote, setReviewNote] = useState('');
  const [tab, setTab] = useState('pending');
  const [driveSaving, setDriveSaving] = useState(null);
  const [driveResult, setDriveResult] = useState({});
  const [expandedId, setExpandedId] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [rejectError, setRejectError] = useState('');
  const filtered = submissions.filter(s => s.status === tab);
  const DRIVE_FOLDER = feature === 'stock_compare' ? DRIVE_FOLDER_STOCK_COMPARE : DRIVE_FOLDER_RECORDER;
  const handleApprove = () => { onReview(selected.id,'approved',reviewNote); setSelected(null); setReviewNote(''); setRejectError(''); };
  const handleReject = () => { if(!reviewNote.trim()){setRejectError('กรุณาใส่เหตุผลก่อนส่งกลับ');return;} onReview(selected.id,'rejected',reviewNote); setSelected(null); setReviewNote(''); setRejectError(''); };

  const uploadToDrive = async (sub, type) => {
    const key = `${sub.id}_${type}`;
    setDriveSaving(key);
    try {
      const base = sub.docNo || sub.id;
      let filename, mimeType, content, isBase64 = false;
      if (type === 'csv') {
        filename = `${base}.csv`; mimeType = 'text/csv';
        content = sub.data.map(d => `${d.barcode},${d.qty},${d.price||0},0`).join('\n');
      } else if (type === 'excel') {
        filename = `${base}.xlsx`; mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        const allRows = buildStockExcelRows(sub.data, sub.docNo, sub.startedAt||sub.submittedAt);
        const ws = XLSX.utils.aoa_to_sheet(allRows);
        ws['!cols'] = [{wch:20},{wch:30},{wch:10},{wch:12},{wch:12},{wch:10},{wch:10},{wch:20}];
        const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Stock');
        content = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' }); isBase64 = true;
      }
      const response = await fetch('/api/drive-upload', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify({ filename, mimeType, content, isBase64, folderId: DRIVE_FOLDER }) });
      const data = await response.json();
      if (!data.ok) throw new Error(data.error || 'Drive error');
      setDriveResult(prev => ({ ...prev, [key]: { ok: true, link: data.link, type } }));
    } catch (e) { setDriveResult(prev => ({ ...prev, [`${sub.id}_${type}`]: { ok: false, err: e.message, type } })); }
    setDriveSaving(null);
  };

  const STAT = {
    pending:  { label: 'รอรีวิว',    soft: '#FFFBEB', line: '#FDE68A', ink: '#B45309' },
    approved: { label: 'อนุมัติแล้ว', soft: '#F0FDF4', line: '#BBF7D0', ink: '#15803D' },
    rejected: { label: 'ส่งกลับ',    soft: '#FEF2F2', line: '#FECACA', ink: '#B91C1C' },
  };

  return (
    <div className="space-y-3">
      <div>
        <h2 className="text-xl font-bold text-slate-800">กล่องขาเข้า</h2>
        <p className="text-[12px] text-slate-500 mt-0.5">อนุมัติหรือส่งกลับ — ส่งกลับต้องบอกเหตุผล</p>
      </div>

      <div className="bg-white border rounded-xl flex overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
        {['pending', 'approved', 'rejected'].map(k => {
          const cnt = submissions.filter(s => s.status === k).length, on = tab === k;
          return (
            <button key={k} onClick={() => setTab(k)} className="flex-1 py-2.5 text-[12px] font-bold border-b-2"
              style={on ? { color: STAT[k].ink, background: STAT[k].soft, borderColor: STAT[k].ink }
                        : { color: '#64748B', borderColor: 'transparent' }}>
              {STAT[k].label}{cnt > 0 && <span className="tabular-nums"> {cnt}</span>}
            </button>
          );
        })}
      </div>

      {filtered.length === 0 ? (
        <div className="bg-white rounded-xl px-4 py-9 text-center" style={{ border: '1px dashed #CBD5E1' }}>
          <Inbox className="mx-auto text-[#E4E6EA] mb-2" size={34} />
          <div className="text-[13px] text-slate-500">ไม่มีใบ{STAT[tab].label}</div>
        </div>
      ) : (
        <div className="space-y-2">
          {filtered.map(s => {
            const st = STAT[s.status] || STAT.pending;
            const open = expandedId === s.id;
            return (
              <div key={s.id} className="bg-white rounded-xl border overflow-hidden" style={{ borderColor: st.line }}>
                <div className="px-3 py-2 flex items-center gap-2 border-b" style={{ background: st.soft, borderColor: st.line }}>
                  <span className="text-[11.5px] font-bold tabular-nums" style={{ color: st.ink }}>{s.docNo || '—'}</span>
                  <span className="ml-auto text-[11px] font-semibold text-slate-700 truncate">{s.counter}</span>
                </div>

                <div className="p-3 space-y-2.5">
                  <div className="flex gap-2">
                    <div className="flex-1 rounded-lg px-2.5 py-2 text-center" style={{ background: '#F6F7F8' }}>
                      <div className="text-[10px] text-slate-500">บาร์โค้ด</div>
                      <div className="text-base font-bold text-slate-800 tabular-nums">{s.itemCount}</div>
                    </div>
                    <div className="flex-1 rounded-lg px-2.5 py-2 text-center" style={{ background: '#EAF1F0' }}>
                      <div className="text-[10px]" style={{ color: '#2A5A55' }}>รวมจำนวน</div>
                      <div className="text-base font-bold tabular-nums" style={{ color: '#2A5A55' }}>{s.totalQty.toLocaleString('th-TH')}</div>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-2 rounded-lg px-2.5 py-2" style={{ background: '#F6F7F8' }}>
                    <div>
                      <div className="text-[9.5px] text-slate-400">เริ่มนับ</div>
                      <div className="text-[11px] font-semibold text-slate-700 tabular-nums">
                        {s.startedAt ? new Date(s.startedAt).toLocaleString('th-TH', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' }) : '—'}
                      </div>
                    </div>
                    <div>
                      <div className="text-[9.5px] text-slate-400">ส่งงาน</div>
                      <div className="text-[11px] font-semibold text-slate-700 tabular-nums">
                        {new Date(s.submittedAt).toLocaleString('th-TH', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })}
                      </div>
                    </div>
                  </div>

                  {s.note && <div className="text-[11.5px] text-slate-500 leading-relaxed">พนักงานแจ้ง: “{s.note}”</div>}

                  {s.status !== 'pending' && s.reviewNote && (
                    <div className="rounded-lg p-2.5 border" style={{ background: st.soft, borderColor: st.line }}>
                      <div className="text-[10px] font-bold" style={{ color: st.ink }}>หมายเหตุที่คุณเขียน</div>
                      <div className="text-[11.5px] mt-0.5 leading-relaxed" style={{ color: st.ink }}>{s.reviewNote}</div>
                    </div>
                  )}

                  {s.status === 'pending' && (
                    <button onClick={() => { setSelected(s); setReviewNote(''); }}
                      className="w-full text-white font-bold text-[14px] rounded-xl"
                      style={{ minHeight: 46, background: '#0F172A' }}>เปิดรีวิว →</button>
                  )}

                  <div className="flex flex-wrap gap-1.5 pt-2 border-t" style={{ borderColor: '#F6F7F8' }}>
                    <button onClick={() => setExpandedId(open ? null : s.id)}
                      className="flex items-center gap-1 px-2.5 rounded-lg text-[11px] font-semibold text-slate-600"
                      style={{ minHeight: 36, background: '#F6F7F8' }}>
                      <FileSpreadsheet size={12} />{open ? 'ซ่อนรายการ' : `รายการ (${s.data?.length || 0})`}
                    </button>
                    <PDFDownloadButton sub={s} />
                    {s.status === 'approved' && (() => {
                      const key = `${s.id}_csv`, res = driveResult[key], saving = driveSaving === key;
                      return (
                        <>
                          <button onClick={() => uploadToDrive(s, 'csv')} disabled={saving}
                            className="flex items-center gap-1 px-2.5 rounded-lg text-[11px] font-semibold disabled:opacity-60"
                            style={{ minHeight: 36, background: '#EAF0F4', color: '#255771' }}>
                            {saving ? <RefreshCw size={12} className="animate-spin" /> : <Upload size={12} />}
                            {saving ? 'กำลังอัปโหลด' : res?.ok ? 'ขึ้น Drive แล้ว' : 'ส่งขึ้น Drive'}
                          </button>
                          {res?.ok && res.link && <a href={res.link} target="_blank" rel="noopener noreferrer" className="text-[10.5px] font-semibold self-center underline" style={{ color: '#255771' }}>เปิด</a>}
                          {res && !res.ok && <span className="text-[10px] self-center" style={{ color: '#B91C1C' }}>ไม่สำเร็จ: {res.err}</span>}
                        </>
                      );
                    })()}
                    {confirmDelete === s.id ? (
                      <div className="flex items-center gap-1 ml-auto">
                        <span className="text-[10.5px] font-semibold" style={{ color: '#B91C1C' }}>ลบใบนี้?</span>
                        <button onClick={() => { onDelete(s.id); setConfirmDelete(null); }}
                          className="px-2.5 rounded-lg text-[10.5px] font-bold text-white" style={{ minHeight: 32, background: '#B91C1C' }}>ลบ</button>
                        <button onClick={() => setConfirmDelete(null)}
                          className="px-2.5 rounded-lg text-[10.5px] font-semibold text-slate-600" style={{ minHeight: 32, background: '#F6F7F8' }}>ยกเลิก</button>
                      </div>
                    ) : (
                      <button onClick={() => setConfirmDelete(s.id)} title="ลบใบนี้"
                        className="ml-auto rounded-lg flex items-center justify-center"
                        style={{ width: 36, height: 36, background: '#FEF2F2', color: '#B91C1C' }}><Trash2 size={14} /></button>
                    )}
                  </div>

                  {open && (
                    <div className="rounded-lg divide-y max-h-52 overflow-y-auto" style={{ background: '#F6F7F8', borderColor: '#E4E6EA' }}>
                      {s.data.map((d, i) => (
                        <div key={i} className="flex items-center gap-2 px-2.5 py-2" style={{ borderColor: '#E4E6EA' }}>
                          <div className="flex-1 min-w-0">
                            <div className="text-[12px] font-semibold text-slate-800 truncate">{d.productName}</div>
                            <div className="text-[10px] text-slate-400 tabular-nums truncate">{d.barcode}{d.location ? ' · ' + d.location : ''}</div>
                          </div>
                          <div className="text-[13px] font-bold text-slate-800 tabular-nums shrink-0">{d.qty}<span className="text-[10px] font-normal text-slate-400 ml-0.5">{d.unit}</span></div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {selected && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-end sm:items-center justify-center p-3">
          <div className="bg-white rounded-2xl w-full max-w-md max-h-[92vh] overflow-y-auto">
            <div className="px-4 py-3 border-b flex justify-between items-start gap-2 sticky top-0 bg-white" style={{ borderColor: '#E4E6EA' }}>
              <div className="min-w-0">
                <h3 className="font-bold text-slate-800 text-[15px]">รีวิวใบนับ</h3>
                <p className="text-[11px] text-slate-500 truncate">
                  {selected.docNo ? selected.docNo + ' · ' : ''}{selected.counter}
                </p>
              </div>
              <button onClick={() => setSelected(null)} className="shrink-0 rounded-lg flex items-center justify-center text-slate-500"
                style={{ width: 36, height: 36, background: '#F6F7F8' }}><X size={17} /></button>
            </div>
            <div className="p-4 space-y-3">
              <div className="flex gap-2">
                <div className="flex-1 rounded-xl px-3 py-2.5 text-center" style={{ background: '#F6F7F8' }}>
                  <div className="text-[10px] text-slate-500">บาร์โค้ด</div>
                  <div className="text-xl font-bold text-slate-800 tabular-nums">{selected.itemCount}</div>
                </div>
                <div className="flex-1 rounded-xl px-3 py-2.5 text-center" style={{ background: '#EAF1F0' }}>
                  <div className="text-[10px]" style={{ color: '#2A5A55' }}>รวมจำนวน</div>
                  <div className="text-xl font-bold tabular-nums" style={{ color: '#2A5A55' }}>{selected.totalQty.toLocaleString('th-TH')}</div>
                </div>
              </div>

              {selected.note && (
                <div className="rounded-xl p-3 border" style={{ background: '#EAF0F4', borderColor: '#B9CFDC' }}>
                  <div className="text-[10px] font-bold" style={{ color: '#255771' }}>พนักงานแจ้ง</div>
                  <div className="text-[12px] mt-0.5 leading-relaxed" style={{ color: '#255771' }}>{selected.note}</div>
                </div>
              )}

              <div className="bg-white border rounded-xl overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
                <div className="px-3 py-2 border-b text-[11.5px] font-bold text-slate-600" style={{ background: '#F6F7F8', borderColor: '#EEF0F3' }}>
                  รายการที่นับ · {selected.data.length}
                </div>
                <div className="max-h-52 overflow-y-auto divide-y" style={{ borderColor: '#F6F7F8' }}>
                  {selected.data.map((d, i) => (
                    <div key={i} className="px-3 py-2 flex items-center gap-2">
                      <div className="flex-1 min-w-0">
                        <div className="text-[12.5px] font-semibold text-slate-800 truncate">{d.productName}</div>
                        <div className="text-[10px] text-slate-400 tabular-nums truncate">
                          {d.barcode}{d.location ? ' · ' + d.location : ''}
                        </div>
                      </div>
                      <div className="text-[13.5px] font-bold text-slate-800 tabular-nums shrink-0">
                        {d.qty}<span className="text-[10px] font-normal text-slate-400 ml-0.5">{d.unit}</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div>
                <label className="text-[12px] font-semibold text-slate-700 mb-1.5 block">
                  หมายเหตุถึงพนักงาน <span style={{ color: '#B91C1C' }}>(บังคับถ้าส่งกลับ)</span>
                </label>
                <textarea value={reviewNote} onChange={e => { setReviewNote(e.target.value); setRejectError(''); }} rows={2}
                  placeholder="เช่น โซน A ยังไม่ครบ นับชั้นล่างด้วย"
                  className="w-full px-3 py-2 border rounded-lg outline-none text-[13px] resize-none"
                  style={rejectError ? { borderColor: '#B91C1C', background: '#FEF2F2' } : { borderColor: '#E2E8F0' }} />
                {rejectError && (
                  <div className="text-[11px] mt-1 flex items-center gap-1" style={{ color: '#B91C1C' }}>
                    <XCircle size={12} />{rejectError}
                  </div>
                )}
              </div>

              <div className="flex gap-2">
                <button onClick={handleReject} className="flex-1 text-white rounded-xl font-bold text-[13.5px] flex items-center justify-center gap-1.5"
                  style={{ minHeight: 50, background: '#B91C1C' }}><ThumbsDown size={16} />ส่งกลับแก้</button>
                <button onClick={handleApprove} className="flex-1 text-white rounded-xl font-bold text-[13.5px] flex items-center justify-center gap-1.5"
                  style={{ minHeight: 50, background: '#15803D' }}><ThumbsUp size={16} />อนุมัติ</button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
function Dashboard({ submissions, products, setView, isSupabaseReady, lastSyncAt, pendingCount }) {
  const pending = submissions.filter(s => s.status === 'pending');
  const approved = submissions.filter(s => s.status === 'approved');
  const rejected = submissions.filter(s => s.status === 'rejected');
  const today = new Date().toDateString();
  const todayCount = submissions.filter(s => new Date(s.submittedAt).toDateString() === today).length;
  const STAT = {
    pending:  { soft: '#FFFBEB', line: '#FDE68A', ink: '#B45309', label: 'รอรีวิว' },
    approved: { soft: '#F0FDF4', line: '#BBF7D0', ink: '#15803D', label: 'อนุมัติ' },
    rejected: { soft: '#FEF2F2', line: '#FECACA', ink: '#B91C1C', label: 'ส่งกลับ' },
  };

  return (
    <div className="space-y-3">
      <div className="flex items-start justify-between gap-2">
        <div>
          <h2 className="text-xl font-bold text-slate-800">แดชบอร์ด</h2>
          <p className="text-[12px] text-slate-500 mt-0.5">ภาพรวมใบนับที่พนักงานส่งเข้ามา</p>
        </div>
        <button onClick={() => setView('settings')}
          className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-[10.5px] font-semibold border shrink-0"
          style={isSupabaseReady
            ? { background: '#EAF1F0', borderColor: '#B6D0CC', color: '#2A5A55' }
            : { background: '#FEF2F2', borderColor: '#FECACA', color: '#B91C1C' }}>
          <Cloud size={11} />{isSupabaseReady ? 'เซิร์ฟเวอร์พร้อม' : 'ยังไม่ตั้งค่า'}
        </button>
      </div>

      {pendingCount > 0 && (
        <button onClick={() => setView('inbox')}
          className="w-full rounded-xl p-3.5 flex items-center gap-3 text-left border-2"
          style={{ background: '#FFFBEB', borderColor: '#FDE68A' }}>
          <div className="text-white p-2 rounded-lg shrink-0" style={{ background: '#B45309' }}><Inbox size={19} /></div>
          <div className="flex-1 min-w-0">
            <div className="text-[14px] font-bold" style={{ color: '#B45309' }}>{pendingCount} ใบรอคุณรีวิว</div>
            <div className="text-[11.5px] mt-0.5" style={{ color: '#B45309' }}>พนักงานรอผลอยู่ — กดเพื่อเปิดกล่องขาเข้า</div>
          </div>
          <ArrowRight size={17} className="shrink-0" style={{ color: '#B45309' }} />
        </button>
      )}

      <div className="grid grid-cols-3 gap-2">
        {[['pending', pending.length], ['approved', approved.length], ['rejected', rejected.length]].map(([k, v]) => (
          <div key={k} className="rounded-xl px-2.5 py-2.5 border" style={{ background: STAT[k].soft, borderColor: STAT[k].line }}>
            <div className="text-[10px] font-semibold" style={{ color: STAT[k].ink }}>{STAT[k].label}</div>
            <div className="text-xl font-bold tabular-nums mt-0.5" style={{ color: STAT[k].ink }}>{v}</div>
          </div>
        ))}
      </div>

      <div className="bg-white border rounded-xl grid grid-cols-2 divide-x" style={{ borderColor: '#E4E6EA' }}>
        <div className="px-3 py-2.5">
          <div className="text-[10.5px] text-slate-500">ส่งเข้ามาวันนี้</div>
          <div className="text-lg font-bold text-slate-800 tabular-nums">{todayCount}</div>
        </div>
        <div className="px-3 py-2.5" style={{ borderColor: '#EEF0F3' }}>
          <div className="text-[10.5px] text-slate-500">สินค้าในเครื่อง</div>
          <div className="text-lg font-bold text-slate-800 tabular-nums">{products.length.toLocaleString('th-TH')}</div>
        </div>
      </div>

      <div className="bg-white rounded-xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
        <div className="px-3 py-2.5 border-b flex items-center justify-between" style={{ background: '#F6F7F8', borderColor: '#EEF0F3' }}>
          <span className="text-[12px] font-bold text-slate-700">ส่งล่าสุด</span>
          {submissions.length > 0 && (
            <button onClick={() => setView('inbox')} className="text-[11px] font-semibold text-slate-500 hover:text-slate-800">ดูทั้งหมด →</button>
          )}
        </div>
        {submissions.length === 0 ? (
          <div className="px-4 py-8 text-center text-[12.5px] text-slate-400">ยังไม่มีใบนับส่งเข้ามา</div>
        ) : (
          <div className="divide-y" style={{ borderColor: '#F6F7F8' }}>
            {submissions.slice(0, 5).map(s => {
              const st = STAT[s.status] || STAT.pending;
              return (
                <button key={s.id} onClick={() => setView('inbox')} className="w-full px-3 py-2.5 flex items-center gap-2.5 text-left hover:bg-[#F6F7F8]">
                  <span className="w-2 h-2 rounded-full shrink-0" style={{ background: st.ink }} />
                  <div className="flex-1 min-w-0">
                    <div className="text-[13px] font-semibold text-slate-800 truncate">{s.counter}</div>
                    <div className="text-[10.5px] text-slate-400 tabular-nums">
                      {new Date(s.submittedAt).toLocaleString('th-TH', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })}
                    </div>
                  </div>
                  <div className="text-right shrink-0">
                    <div className="text-[12.5px] font-bold text-slate-800 tabular-nums">{s.itemCount} · {s.totalQty.toLocaleString('th-TH')}</div>
                    <div className="text-[10.5px] font-semibold" style={{ color: st.ink }}>{st.label}</div>
                  </div>
                </button>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

function CompareStockView({ submissions, supabaseConfig, compareState, setCompareState }) {
  const { selectedSub, compareData, loading, loadProgress, error, driveSaving, driveResult } = compareState;
  const set = (patch) => setCompareState(prev => ({ ...prev, ...patch }));
  const approvedSubs = submissions.filter(s => s.status === 'approved');
  const { url: sbUrl, anonKey: sbKey, stockTableName } = supabaseConfig;

  const fetchAndCompare = async (sub) => {
    const now = new Date().toISOString();
    set({ selectedSub: sub, loading: true, error: '', compareData: [], driveResult: null, compareAt: now });
    const groupedByBarcode = {};
    sub.data.forEach(d => {
      if (!groupedByBarcode[d.barcode]) {
        groupedByBarcode[d.barcode] = { barcode: d.barcode, productName: d.productName, unit: d.unit||'', qty: 0, scannedAt: d.scannedAt, locations: [], notFound: !!d.notFound };
      }
      const g = groupedByBarcode[d.barcode];
      g.qty += d.qty;
      if (d.scannedAt && (!g.scannedAt || d.scannedAt < g.scannedAt)) g.scannedAt = d.scannedAt;
      if (d.location && !g.locations.includes(d.location)) g.locations.push(d.location);
    });
    const groupedData = Object.values(groupedByBarcode);
    const codes = [...new Set(groupedData.map(d => d.barcode))];
    const submittedAt = sub.submittedAt;
    const minScannedAt = groupedData.reduce((min, d) => d.scannedAt && d.scannedAt < min ? d.scannedAt : min, submittedAt);
    const table = stockTableName || 'product_stock';
    const batchSize = 50;
    const qc = col => encodeURIComponent(`"${col}"`);

    try {
      let stockRows = [];
      for (let i = 0; i < codes.length; i += batchSize) {
        set({ loadProgress: `[1/3] ยอดปัจจุบัน ${Math.min(i+batchSize,codes.length)}/${codes.length}...` });
        const batch = codes.slice(i, i+batchSize);
        const inList = batch.map(c => encodeURIComponent(c)).join(',');
        const colCode = qc('รหัสสินค้า');
        const colSel = ['รหัสสินค้า','ชื่อสินค้า','หน่วย','รวม'].map(qc).join(',');
        const rows = await sbFetch(sbUrl, sbKey, table, `${colCode}=in.(${inList})&select=${colSel}`);
        stockRows = stockRows.concat(rows);
      }
      const sbMap = {};
      stockRows.forEach(r => {
        const code = String(r['รหัสสินค้า']||'');
        sbMap[code] = { name: String(r['ชื่อสินค้า']||''), unit: String(r['หน่วย']||'ชิ้น'), currentStock: parseInt(String(r['รวม']||'0').replace(/[^\d-]/g,''))||0 };
      });

      const saleMap = {};
      const startDate = minScannedAt.slice(0, 10);
      const endDate = submittedAt.slice(0, 10);
      for (let i = 0; i < codes.length; i += batchSize) {
        set({ loadProgress: `[2/3] ยอดขายระหว่างนับ ${Math.min(i+batchSize,codes.length)}/${codes.length}...` });
        const batch = codes.slice(i, i+batchSize);
        const inList = batch.map(c => encodeURIComponent(c)).join(',');
        const colSinc = qc('สินค้า');
        const colDate = qc('วันที่');
        const colTime = qc('เวลา');
        const colQty  = qc('จำนวน');
        const qs = `${colSinc}=in.(${inList})&${colDate}=gte.${startDate}&${colDate}=lte.${endDate}&select=${colSinc},${colDate},${colTime},${colQty}`;
        const rows = await sbFetch(sbUrl, sbKey, 'sale_item_with_time', qs);
        rows.forEach((r) => {
          const code = String(r['สินค้า']||'');
          const rawTime = r['เวลา'] ? String(r['เวลา']) : '00:00:00';
          const rawDate = r['วันที่'] ? String(r['วันที่']).slice(0,10) : startDate;
          const saleTime = new Date(`${rawDate}T${rawTime.slice(0,8)}`);
          const item = groupedData.find(d => d.barcode === code);
          const scannedAt = new Date(item?.scannedAt || minScannedAt);
          if (saleTime >= scannedAt && saleTime <= new Date(submittedAt)) {
            saleMap[code] = (saleMap[code]||0) + (parseFloat(r['จำนวน'])||0);
          }
        });
      }

      const receiveMap = {};
      for (let i = 0; i < codes.length; i += batchSize) {
        set({ loadProgress: `[3/3] ยอดรับสินค้าระหว่างนับ ${Math.min(i+batchSize,codes.length)}/${codes.length}...` });
        const batch = codes.slice(i, i+batchSize);
        const inList = batch.map(c => encodeURIComponent(c)).join(',');
        const isoStart = encodeURIComponent(minScannedAt);
        const isoEnd   = encodeURIComponent(submittedAt);
        const qs = `barcode=in.(${inList})&created_at=gte.${isoStart}&created_at=lte.${isoEnd}&select=barcode,qty,created_at`;
        const rows = await sbFetch(sbUrl, sbKey, 'imp_data', qs);
        rows.forEach(r => {
          const code = String(r.barcode||'');
          const recvTime = new Date(r.created_at);
          const item = groupedData.find(d => d.barcode === code);
          const scannedAt = new Date(item?.scannedAt || minScannedAt);
          if (recvTime >= scannedAt && recvTime <= new Date(submittedAt)) {
            receiveMap[code] = (receiveMap[code]||0) + (parseFloat(r.qty)||0);
          }
        });
      }

      const compared = groupedData.map(d => {
        const sb = sbMap[d.barcode]||null;
        const counted = d.qty;
        const sale = Math.round(saleMap[d.barcode]||0);
        const purchase = Math.round(receiveMap[d.barcode]||0);
        const stockAtSubmit = sb ? sb.currentStock : null;
        const adjustedCount = counted - sale + purchase;
        const adjustStock = stockAtSubmit !== null ? adjustedCount - stockAtSubmit : null;
        return { barcode: d.barcode, productName: sb?sb.name:d.productName, unit: sb?sb.unit:(d.unit||''), scannedAt: d.scannedAt||null, locations: d.locations||[], counted, sale, purchase, adjustedCount, stockAtSubmit, adjustStock, found: !!sb, notFound: !!d.notFound };
      });
      set({ compareData: compared });
    } catch (e) { set({ error: e.message }); }
    set({ loading: false, loadProgress: '' });
  };

  const buildSimpleCSV = () => compareData.map(d => `${d.barcode},${d.adjustStock??''}`).join('\n');
  const buildFullCSV = () => {
    const info = selectedSub ? [`# Counter: ${selectedSub.counter}`, `# time_submit: ${new Date(selectedSub.submittedAt).toLocaleString('th-TH')}`, `# compare_at: ${compareState.compareAt?new Date(compareState.compareAt).toLocaleString('th-TH'):'-'}`].join('\n') : '';
    const header = 'รหัสสินค้า,ชื่อสินค้า,หน่วย,location,นับได้,ขายระหว่างนับ,รับระหว่างนับ,Adjusted_count,stock_at_submit,Adjust_stock,พบในระบบ';
    const rows = compareData.map(d => `${d.barcode},"${d.productName}",${d.unit},"${(d.locations||[]).join('|')}",${d.counted},${d.sale},${d.purchase},${d.adjustedCount},${d.stockAtSubmit??'N/A'},${d.adjustStock??'N/A'},${d.found?'Y':'N'}`);
    return (info?info+'\n':'')+header+'\n'+rows.join('\n');
  };
  const getFilename = (ext='csv') => { const date = new Date().toISOString().slice(0,10); const counter = selectedSub?selectedSub.counter.replace(/[^a-zA-Z0-9ก-๙]/g,'_'):'compare'; return `compare_${counter}_${date}.${ext}`; };

  const downloadCompareCSV = () => downloadBlob(new Blob(['﻿'+buildFullCSV()],{type:'text/csv;charset=utf-8'}), getFilename('csv'));
  const downloadCompareExcel = () => {
    const rows = [['รหัสสินค้า','ชื่อสินค้า','หน่วย','location','นับได้(count)','ขายระหว่างนับ','รับระหว่างนับ','Adjusted_count','stock_at_submit','Adjust_stock','พบในระบบ'], ...compareData.map(d=>[d.barcode,d.productName,d.unit,(d.locations||[]).join('|'),d.counted,d.sale,d.purchase,d.adjustedCount,d.stockAtSubmit??'N/A',d.adjustStock??'N/A',d.found?'Y':'N'])];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{wch:18},{wch:28},{wch:8},{wch:14},{wch:10},{wch:12},{wch:12},{wch:14},{wch:14},{wch:12},{wch:8}];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Compare');
    const buf = XLSX.write(wb,{type:'array',bookType:'xlsx'});
    downloadBlob(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}), getFilename('xlsx'));
  };

  const saveToDrive = async () => {
    set({ driveSaving: true, driveResult: null });
    try {
      const content = buildSimpleCSV();
      const filename = getFilename('txt');
      const response = await fetch('/api/drive-upload', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ filename, mimeType:'text/csv', content, isBase64: false, folderId: DRIVE_FOLDER_STOCK_COMPARE }) });
      const data = await response.json();
      if (!data.ok) throw new Error(data.error||'Drive error');
      set({ driveResult: { ok: true, link: data.link } });
    } catch (e) { set({ driveResult: { ok: false, err: e.message } }); }
    set({ driveSaving: false });
  };

  if (!sbUrl || !sbKey) return (
    <div className="space-y-4"><div><h2 className="text-2xl font-bold text-slate-800">เปรียบเทียบสต็อก</h2></div>
      <div className="bg-[#FFFBEB] border border-[#FDE68A] rounded-xl p-4"><AlertCircle className="text-[#B45309] mb-2" size={24}/><div className="font-semibold text-[#B45309]">ยังไม่ได้ตั้งค่า Supabase</div><p className="text-sm text-[#B45309] mt-1">ไปที่ ตั้งค่า เพื่อใส่ URL และ Anon Key ก่อน</p></div>
    </div>
  );

  return (
    <div className="space-y-4">
      <div><h2 className="text-2xl font-bold text-slate-800">เปรียบเทียบสต็อก</h2><p className="text-sm text-slate-500">เปรียบเทียบยอดนับกับ Supabase ผ่าน REST API</p></div>
      {approvedSubs.length === 0 ? <div className="bg-white rounded-xl p-8 text-center border border-[#E4E6EA]"><ArrowLeftRight className="mx-auto text-[#E4E6EA] mb-2" size={40}/><div className="text-slate-500">ยังไม่มีรายการที่อนุมัติแล้ว</div></div> : (
        <div className="bg-white rounded-xl border border-[#E4E6EA] p-4 space-y-2">
          <div className="text-sm font-semibold text-slate-700 mb-2">เลือก submission ที่จะเปรียบเทียบ:</div>
          {approvedSubs.map(s => (
            <button key={s.id} onClick={() => fetchAndCompare(s)} disabled={loading} className={`w-full text-left p-3 rounded-lg border transition-colors ${selectedSub?.id===s.id?'border-[#94A3B8] bg-[#F6F7F8]':'border-[#E4E6EA] hover:border-[#E4E6EA] hover:bg-[#F6F7F8]'} disabled:opacity-50`}>
              <div className="flex items-center gap-2"><span className="font-mono text-xs font-bold text-[#0F172A] bg-[#F6F7F8] px-2 py-0.5 rounded">{s.docNo||'—'}</span><span className="font-semibold text-slate-800">{s.counter}</span></div>
              <div className="text-xs text-slate-500 mt-0.5">{s.itemCount} รายการ • {s.totalQty.toLocaleString()} ชิ้น</div>
              <div className="text-[10px] text-slate-500 mt-1 font-mono">
                <span className="text-blue-600">เริ่มนับ: {s.startedAt?new Date(s.startedAt).toLocaleString('th-TH',{day:'2-digit',month:'short',year:'2-digit',hour:'2-digit',minute:'2-digit'}):'—'}</span>
                {' → '}
                <span className="text-slate-600">ส่งงาน: {new Date(s.submittedAt).toLocaleString('th-TH',{day:'2-digit',month:'short',year:'2-digit',hour:'2-digit',minute:'2-digit'})}</span>
              </div>
            </button>
          ))}
        </div>
      )}
      {loading && <div className="bg-white rounded-xl border border-[#E4E6EA] p-6 text-center"><div className="w-8 h-8 border-4 border-[#E4E6EA] border-t-[#0F172A] rounded-full animate-spin mx-auto mb-3"/><div className="text-sm text-slate-600">{loadProgress || 'กำลังดึงข้อมูล...'}</div></div>}
      {error && <div className="bg-[#FEF2F2] border border-[#FECACA] rounded-xl p-4 text-[#B91C1C] text-sm"><strong>ข้อผิดพลาด:</strong> {error}</div>}
      {compareData.length > 0 && (
        <div className="space-y-3">
          {selectedSub && (
            <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-xs">
              <div className="font-semibold text-blue-900 mb-1">📅 Reconcile Window</div>
              <div className="font-mono text-blue-800">
                {selectedSub.startedAt ? new Date(selectedSub.startedAt).toLocaleString('th-TH',{day:'2-digit',month:'short',year:'2-digit',hour:'2-digit',minute:'2-digit'}) : '—'}
                {' → '}
                {new Date(selectedSub.submittedAt).toLocaleString('th-TH',{day:'2-digit',month:'short',year:'2-digit',hour:'2-digit',minute:'2-digit'})}
              </div>
              <div className="text-blue-700 mt-1 text-[10px]">ดึง sale/purchase ของช่วงนี้มาคำนวณ adjusted_count</div>
            </div>
          )}
          {compareState.debugInfo && null}
          <div className="flex flex-wrap gap-2">
            <button onClick={downloadCompareCSV} className="flex items-center gap-1 px-3 py-2 bg-[#F6F7F8] hover:bg-[#F6F7F8] text-[#0F172A] rounded-lg text-sm font-medium"><Download size={14}/>CSV เต็ม</button>
            <button onClick={downloadCompareExcel} className="flex items-center gap-1 px-3 py-2 bg-[#EAF0F4] hover:bg-[#EAF0F4] text-[#255771] rounded-lg text-sm font-medium"><FileSpreadsheet size={14}/>Excel</button>
            <button onClick={saveToDrive} disabled={driveSaving} className="flex items-center gap-1 px-3 py-2 bg-[#EAF1F0] hover:bg-[#EAF1F0] text-[#2A5A55] rounded-lg text-sm font-medium disabled:opacity-60">{driveSaving?<RefreshCw size={14} className="animate-spin"/>:<Upload size={14}/>}Drive (รหัส,ผลต่าง)</button>
            {driveResult?.ok && <span className="flex items-center gap-1 text-xs text-[#15803D]"><CheckCircle2 size={12}/>อัพโหลดแล้ว{driveResult.link&&<a href={driveResult.link} target="_blank" rel="noopener noreferrer" className="underline ml-1">เปิด</a>}</span>}
            {driveResult?.err && <span className="text-xs text-[#B91C1C]">Error: {driveResult.err}</span>}
          </div>
          <div className="grid grid-cols-3 gap-2 text-center text-xs">
            {[{label:'พบในระบบ',v:compareData.filter(d=>d.found).length,c:'emerald'},{label:'ไม่พบ',v:compareData.filter(d=>!d.found).length,c:'amber'},{label:'ส่วนต่าง≠0',v:compareData.filter(d=>d.adjustStock!==null&&d.adjustStock!==0).length,c:'red'}].map(x=>(
              <div key={x.label} className={`rounded-lg p-2 border ${x.c==='emerald'?'bg-[#EAF1F0] border-[#EAF1F0] text-[#2A5A55]':x.c==='amber'?'bg-[#FFFBEB] border-[#FFFBEB] text-[#B45309]':'bg-[#FEF2F2] border-[#FEF2F2] text-[#B91C1C]'}`}><div className="text-lg font-bold">{x.v}</div><div className="opacity-75">{x.label}</div></div>
            ))}
          </div>
          <div className="bg-white rounded-xl border border-[#E4E6EA] overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead className="bg-[#F6F7F8] border-b border-[#E4E6EA]">
                  <tr>{['รหัส','ชื่อสินค้า','นับได้','ขาย','รับ','Adj.Count','ยอดSB','Adj.Stock'].map(h=><th key={h} className="px-3 py-2 text-left font-semibold text-slate-600 whitespace-nowrap">{h}</th>)}</tr>
                </thead>
                <tbody className="divide-y divide-[#F6F7F8]">
                  {compareData.map((d,i) => {
                    const adjColor = d.adjustStock===null?'text-slate-400':d.adjustStock===0?'text-[#35706A]':d.adjustStock>0?'text-blue-600':'text-[#B91C1C]';
                    return (
                      <tr key={i} className={d.notFound?'bg-[#FFFBEB]/30':''}>
                        <td className="px-3 py-2 font-mono text-slate-700">{d.barcode}</td>
                        <td className="px-3 py-2 text-slate-800 max-w-[160px] truncate">{d.productName}{d.locations?.length>0&&<div className="text-[10px] text-slate-400">{d.locations.join(', ')}</div>}</td>
                        <td className="px-3 py-2 font-semibold text-slate-800">{d.counted}</td>
                        <td className="px-3 py-2 text-orange-600">{d.sale}</td>
                        <td className="px-3 py-2 text-blue-600">{d.purchase}</td>
                        <td className="px-3 py-2 font-semibold">{d.adjustedCount}</td>
                        <td className="px-3 py-2 text-slate-600">{d.stockAtSubmit??<span className="text-[#B45309]">N/A</span>}</td>
                        <td className={`px-3 py-2 font-bold ${adjColor}`}>{d.adjustStock!==null?d.adjustStock:<span className="text-slate-400">N/A</span>}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

const REPORT_TOPICS = [
  { id: 'count',   label: 'การนับ',        needDate: true  },
  { id: 'invoice', label: 'บิลซื้อ',        needDate: true  },
  { id: 'stock',   label: 'สินค้าคงเหลือ',  needDate: false },
  { id: 'in',      label: 'ซื้อเข้า',       needDate: true  },
  { id: 'out',     label: 'ขายออก',        needDate: true  },
];

// หัวคอลัมน์ภาษาไทย — คีย์ตรงกับที่ /api/report ส่งกลับมา
const COL_LABEL = {
  doc_no: 'เลขที่เอกสาร', counted_at: 'วันที่นับ', counter_name: 'ผู้นับ', zone: 'โซน',
  barcode: 'บาร์โค้ด', product_code: 'รหัสสินค้า', name: 'ชื่อสินค้า', unit: 'หน่วย',
  qty: 'จำนวน', status: 'สถานะ',
  file_name: 'ไฟล์', invoice_no: 'เลขที่บิล', invoice_date: 'วันที่บิล',
  vendor_name: 'ผู้ขาย', description: 'รายละเอียด', ea: 'ea', price_ea: 'ราคา/หน่วย',
  discount: 'ส่วนลด', amount: 'จำนวนเงิน', vat: 'ภาษี', total: 'รวม',
  on_hand: 'คงเหลือ', occurred_at: 'วันเวลา', kind: 'ประเภท', party: 'ผู้ขาย / ลูกค้า',
};

const isNumCol = (k) => ['qty','ea','price_ea','discount','amount','total','on_hand'].includes(k);
const isDateCol = (k) => ['counted_at','occurred_at','invoice_date'].includes(k);

function fmtCell(k, v) {
  if (v == null || v === '') return '';
  if (isDateCol(k)) {
    const d = new Date(v);
    if (isNaN(d)) return String(v);
    return k === 'invoice_date'
      ? d.toLocaleDateString('th-TH')
      : d.toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' });
  }
  if (isNumCol(k)) return Number(v).toLocaleString('th-TH', { maximumFractionDigits: 2 });
  return String(v);
}

const todayISO = () => new Date().toISOString().slice(0, 10);
const daysAgoISO = (n) => new Date(Date.now() - n * 86400000).toISOString().slice(0, 10);

function ReportView() {
  const [topic, setTopic] = useState('count');
  const [from, setFrom] = useState(daysAgoISO(30));
  const [to, setTo] = useState(todayISO());
  const [doc, setDoc] = useState('');
  const [barcode, setBarcode] = useState('');
  const [party, setParty] = useState('');
  const [rows, setRows] = useState(null);
  const [cols, setCols] = useState([]);
  const [colFilter, setColFilter] = useState({});   // กรองรายคอลัมน์
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState('');

  const conf = REPORT_TOPICS.find(t => t.id === topic);

  async function run() {
    setLoading(true); setErr(''); setRows(null); setColFilter({});
    try {
      const res = await fetch('/api/report', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          topic,
          from: conf.needDate ? from : null,
          to:   conf.needDate ? to   : null,
          doc: doc || null, barcode: barcode || null, party: party || null,
          limit: 5000,
        }),
      });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || 'ดึงข้อมูลไม่สำเร็จ');
      setRows(j.rows || []);
      setCols(j.columns || (j.rows?.[0] ? Object.keys(j.rows[0]) : []));
    } catch (e) {
      setErr(e.message);
    } finally {
      setLoading(false);
    }
  }

  // กรองรายคอลัมน์ — ทุกช่องต้องผ่านหมด (AND)
  const shown = useMemo(() => {
    if (!rows) return [];
    const active = Object.entries(colFilter).filter(([, v]) => v.trim() !== '');
    if (!active.length) return rows;
    return rows.filter(r => active.every(([k, v]) =>
      String(r[k] ?? '').toLowerCase().includes(v.trim().toLowerCase())));
  }, [rows, colFilter]);

  const sums = useMemo(() => {
    const s = {};
    for (const k of cols) if (isNumCol(k)) s[k] = shown.reduce((a, r) => a + (Number(r[k]) || 0), 0);
    return s;
  }, [shown, cols]);

  function exportXlsx() {
    const header = cols.map(k => COL_LABEL[k] || k);
    const body = shown.map(r => cols.map(k => (isNumCol(k) ? (Number(r[k]) || 0) : fmtCell(k, r[k]))));
    const meta = [
      [`รายงาน: ${conf.label}`],
      conf.needDate ? [`ช่วงวันที่: ${from} ถึง ${to}`] : ['ทั้งหมด (ไม่จำกัดวันที่)'],
      [`ดึงเมื่อ: ${new Date().toLocaleString('th-TH')} · ${shown.length} แถว`],
      [],
      header,
    ];
    const ws = XLSX.utils.aoa_to_sheet([...meta, ...body]);
    ws['!cols'] = cols.map(k => ({ wch: k === 'name' || k === 'description' ? 34 : 16 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, conf.label.slice(0, 30));
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    downloadBlob(new Blob([out], { type: 'application/octet-stream' }),
      `report-${topic}-${todayISO()}.xlsx`);
  }

  function exportCsv() {
    const esc = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
    const lines = [cols.map(k => esc(COL_LABEL[k] || k)).join(',')];
    for (const r of shown) lines.push(cols.map(k => esc(fmtCell(k, r[k]))).join(','));
    downloadBlob(new Blob(['\uFEFF' + lines.join('\n'), ], { type: 'text/csv;charset=utf-8' }),
      `report-${topic}-${todayISO()}.csv`);
  }

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-2xl font-bold text-slate-800">รายงาน</h2>
        <p className="text-sm text-slate-500">เลือกเรื่อง ใส่เงื่อนไข แล้วส่งออกเป็นไฟล์</p>
      </div>

      {/* เลือกเรื่อง */}
      <div className="bg-white rounded-xl border border-[#E4E6EA] p-4 space-y-4">
        <div className="flex flex-wrap gap-2">
          {REPORT_TOPICS.map(t => (
            <button key={t.id} onClick={() => { setTopic(t.id); setRows(null); }}
              className={`px-3 py-2 rounded-lg text-sm font-medium border transition-colors ${
                topic === t.id
                  ? 'bg-slate-800 border-slate-800 text-white'
                  : 'bg-white border-[#E4E6EA] text-slate-600 hover:bg-[#F6F7F8]'}`}>
              {t.label}
            </button>
          ))}
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          {conf.needDate && (
            <>
              <label className="block">
                <span className="text-xs font-medium text-slate-500">ตั้งแต่</span>
                <input type="date" value={from} onChange={e => setFrom(e.target.value)}
                  className="mt-1 w-full border border-[#E4E6EA] rounded-lg px-3 py-2 text-sm" />
              </label>
              <label className="block">
                <span className="text-xs font-medium text-slate-500">ถึง</span>
                <input type="date" value={to} onChange={e => setTo(e.target.value)}
                  className="mt-1 w-full border border-[#E4E6EA] rounded-lg px-3 py-2 text-sm" />
              </label>
            </>
          )}
          {conf.needDate && (
            <label className="block">
              <span className="text-xs font-medium text-slate-500">เลขที่เอกสาร</span>
              <input value={doc} onChange={e => setDoc(e.target.value)} placeholder="ทั้งหมด"
                className="mt-1 w-full border border-[#E4E6EA] rounded-lg px-3 py-2 text-sm" />
            </label>
          )}
          <label className="block">
            <span className="text-xs font-medium text-slate-500">บาร์โค้ด / ชื่อสินค้า</span>
            <input value={barcode} onChange={e => setBarcode(e.target.value)} placeholder="ทั้งหมด"
              className="mt-1 w-full border border-[#E4E6EA] rounded-lg px-3 py-2 text-sm" />
          </label>
          {(topic === 'invoice' || topic === 'in' || topic === 'out') && (
            <label className="block">
              <span className="text-xs font-medium text-slate-500">
                {topic === 'out' ? 'ลูกค้า' : 'ผู้ขาย'}
              </span>
              <input value={party} onChange={e => setParty(e.target.value)} placeholder="ทั้งหมด"
                className="mt-1 w-full border border-[#E4E6EA] rounded-lg px-3 py-2 text-sm" />
            </label>
          )}
          {topic === 'stock' && (
            <label className="block">
              <span className="text-xs font-medium text-slate-500">ประเภทสินค้า</span>
              <input value={party} onChange={e => setParty(e.target.value)} placeholder="ทั้งหมด"
                className="mt-1 w-full border border-[#E4E6EA] rounded-lg px-3 py-2 text-sm" />
            </label>
          )}
        </div>

        <button onClick={run} disabled={loading}
          className="w-full bg-[#0F172A] hover:bg-[#0F172A] disabled:bg-[#E4E6EA] text-white py-3 rounded-xl font-semibold flex items-center justify-center gap-2">
          {loading ? <><RefreshCw size={18} className="animate-spin" />กำลังดึง…</>
                   : <><Search size={18} />ดึงรายงาน</>}
        </button>

        {err && (
          <div className="bg-[#FEF2F2] border border-[#FECACA] text-[#B91C1C] rounded-lg px-3 py-2 text-sm flex items-start gap-2">
            <AlertCircle size={16} className="mt-0.5 shrink-0" />{err}
          </div>
        )}
      </div>

      {/* ผลลัพธ์ */}
      {rows && (
        <div className="bg-white rounded-xl border border-[#E4E6EA] overflow-hidden">
          <div className="px-4 py-3 border-b border-[#E4E6EA] bg-[#F6F7F8] flex flex-wrap items-center justify-between gap-2">
            <div className="text-sm font-semibold text-slate-800">
              {conf.label} · {shown.length.toLocaleString('th-TH')} แถว
              {shown.length !== rows.length && (
                <span className="text-slate-400 font-normal"> (จาก {rows.length.toLocaleString('th-TH')})</span>
              )}
            </div>
            <div className="flex gap-2">
              {Object.values(colFilter).some(v => v?.trim()) && (
                <button onClick={() => setColFilter({})}
                  className="px-3 py-1.5 rounded-lg text-xs font-medium border border-[#E4E6EA] text-slate-600 hover:bg-white flex items-center gap-1">
                  <X size={12} />ล้างตัวกรอง
                </button>
              )}
              <button onClick={exportCsv} disabled={!shown.length}
                className="px-3 py-1.5 rounded-lg text-xs font-medium border border-[#E4E6EA] text-slate-700 hover:bg-white disabled:opacity-40 flex items-center gap-1">
                <Download size={12} />CSV
              </button>
              <button onClick={exportXlsx} disabled={!shown.length}
                className="px-3 py-1.5 rounded-lg text-xs font-medium bg-[#35706A] hover:bg-[#2A5A55] text-white disabled:opacity-40 flex items-center gap-1">
                <FileSpreadsheet size={12} />Excel
              </button>
            </div>
          </div>

          {rows.length === 0 ? (
            <div className="p-8 text-center text-slate-500 text-sm">
              ไม่มีข้อมูลในเงื่อนไขนี้
            </div>
          ) : (
            <div className="overflow-auto max-h-[70vh]">
              <table className="w-full text-xs">
                <thead className="sticky top-0 bg-white shadow-sm">
                  <tr>
                    {cols.map(k => (
                      <th key={k} className={`px-3 pt-2.5 pb-1 font-semibold text-slate-600 whitespace-nowrap ${isNumCol(k) ? 'text-right' : 'text-left'}`}>
                        {COL_LABEL[k] || k}
                      </th>
                    ))}
                  </tr>
                  <tr>
                    {cols.map(k => (
                      <th key={k} className="px-2 pb-2 border-b border-[#E4E6EA]">
                        <input value={colFilter[k] || ''}
                          onChange={e => setColFilter(f => ({ ...f, [k]: e.target.value }))}
                          placeholder="กรอง"
                          className="w-full min-w-[70px] border border-[#E4E6EA] rounded px-1.5 py-1 text-[11px] font-normal focus:border-[#94A3B8] focus:outline-none" />
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {shown.map((r, i) => (
                    <tr key={i} className={i % 2 ? 'bg-[#F6F7F8]/60' : ''}>
                      {cols.map(k => (
                        <td key={k} className={`px-3 py-1.5 whitespace-nowrap ${
                          isNumCol(k) ? 'text-right tabular-nums font-medium text-slate-800' : 'text-slate-600'}`}>
                          {fmtCell(k, r[k])}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
                {Object.keys(sums).length > 0 && (
                  <tfoot className="sticky bottom-0 bg-white border-t-2 border-[#E4E6EA]">
                    <tr>
                      {cols.map((k, idx) => (
                        <td key={k} className={`px-3 py-2 font-bold text-slate-800 ${isNumCol(k) ? 'text-right tabular-nums' : ''}`}>
                          {idx === 0 ? 'รวม' : sums[k] != null
                            ? sums[k].toLocaleString('th-TH', { maximumFractionDigits: 2 }) : ''}
                        </td>
                      ))}
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function StaffAdminView() {
  const [staff, setStaff] = useState(null);
  const [tab, setTab] = useState('active');
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState('');
  const [mode, setMode] = useState(null);      // null | 'add' | id ของคนที่เปิดดู
  const [confirmOff, setConfirmOff] = useState(null);
  const [form, setForm] = useState({ name: '', dept: 'คลังสินค้า', allow_recorder: true, allow_compare: false, allow_invoice: false });
  const [saving, setSaving] = useState(false);

  const FEAT_META = [
    { key: 'allow_recorder', label: 'Recorder', sub: 'นับสต็อกส่งผู้จัดการ', soft: '#EAF1F0', line: '#B6D0CC', ink: '#2A5A55', main: '#35706A' },
    { key: 'allow_compare',  label: 'นับ+เปรียบเทียบ', sub: 'นับแล้วเทียบกับยอดในระบบ', soft: '#F2EFFA', line: '#D5CAEE', ink: '#5F45A8', main: '#7658C9' },
    { key: 'allow_invoice',  label: 'สแกนบิล', sub: 'คีย์บิลซื้อด้วยรูป', soft: '#EAF0F4', line: '#B9CFDC', ink: '#255771', main: '#2F6E90' },
  ];
  const DEPTS = ['คลังสินค้า', 'หน้าร้าน', 'บัญชี'];

  const load = async () => {
    setLoading(true); setErr('');
    try {
      const res = await fetch('/api/staff');
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || 'โหลดรายชื่อไม่สำเร็จ');
      setStaff(j.staff || []);
    } catch (e) { setErr(e.message); } finally { setLoading(false); }
  };
  useEffect(() => { load(); }, []);

  const post = async (body) => {
    setSaving(true); setErr('');
    try {
      const res = await fetch('/api/staff', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || 'บันทึกไม่สำเร็จ');
      await load();
      return true;
    } catch (e) { setErr(e.message); return false; } finally { setSaving(false); }
  };

  const shown = (staff || []).filter(s => (tab === 'active' ? s.active : !s.active));
  const opened = mode && mode !== 'add' ? (staff || []).find(s => s.id === mode) : null;

  // ── เพิ่มพนักงาน ────────────────────────────────────────────────
  if (mode === 'add') return (
    <div className="space-y-3">
      <button onClick={() => setMode(null)} className="text-[12px] font-semibold text-slate-500 hover:text-slate-800">‹ กลับรายชื่อ</button>
      <div>
        <h2 className="text-xl font-bold text-slate-800">เพิ่มพนักงาน</h2>
        <p className="text-[12px] text-slate-500 mt-0.5">ไม่ต้องตั้งรหัสผ่าน — ระบบไม่ใช้ล็อกอิน</p>
      </div>

      <div className="bg-white border rounded-xl p-3 space-y-3" style={{ borderColor: '#E4E6EA' }}>
        <div>
          <div className="text-[11.5px] font-semibold text-slate-600 mb-1.5">ชื่อที่ขึ้นหน้าเลือกชื่อ</div>
          <input value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))}
            placeholder="สมชาย พ."
            className="w-full px-3 py-2.5 border rounded-lg outline-none focus:border-[#35706A] text-[15px]"
            style={{ borderColor: '#E2E8F0' }} />
          <div className="text-[10.5px] text-slate-400 mt-1">ใช้ชื่อจริง + อักษรย่อนามสกุล เพราะเครื่องรวมกันหลายคน</div>
        </div>
        <div>
          <div className="text-[11.5px] font-semibold text-slate-600 mb-1.5">แผนก</div>
          <div className="flex gap-1.5">
            {DEPTS.map(d => (
              <button key={d} onClick={() => setForm(f => ({ ...f, dept: d }))}
                className="flex-1 rounded-lg text-[12px] font-semibold border"
                style={form.dept === d
                  ? { minHeight: 42, background: '#0F172A', borderColor: '#0F172A', color: '#fff' }
                  : { minHeight: 42, background: '#fff', borderColor: '#E2E8F0', color: '#475569' }}>{d}</button>
            ))}
          </div>
        </div>
      </div>

      <div className="bg-white border rounded-xl overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
        <div className="px-3 py-2.5 border-b" style={{ background: '#F6F7F8', borderColor: '#EEF0F3' }}>
          <div className="text-[12px] font-bold text-slate-700">ใช้ฟีเจอร์อะไรได้</div>
          <div className="text-[10.5px] text-slate-500 mt-0.5">เปิดเท่าที่ต้องใช้ — หน้าเลือกชื่อของฟีเจอร์อื่นจะไม่มีชื่อนี้</div>
        </div>
        <div className="divide-y" style={{ borderColor: '#F6F7F8' }}>
          {FEAT_META.map(f => {
            const on = form[f.key];
            return (
              <button key={f.key} onClick={() => setForm(v => ({ ...v, [f.key]: !v[f.key] }))}
                className="w-full px-3 py-2.5 flex items-center gap-2.5 text-left">
                <div className="flex-1 min-w-0">
                  <div className="text-[13px] font-semibold" style={{ color: on ? f.ink : '#334155' }}>{f.label}</div>
                  <div className="text-[10.5px] text-slate-400 truncate">{f.sub}</div>
                </div>
                <span className="shrink-0 rounded-full border-2 flex items-center justify-center"
                  style={{ width: 26, height: 26, background: on ? f.main : '#fff', borderColor: on ? f.main : '#CBD5E1' }}>
                  {on && <Check size={14} color="#fff" />}
                </span>
              </button>
            );
          })}
        </div>
      </div>

      {err && (
        <div className="rounded-xl px-3 py-2.5 text-[12px] flex items-start gap-2 border" style={{ background: '#FEF2F2', borderColor: '#FECACA', color: '#B91C1C' }}>
          <AlertCircle size={14} className="mt-0.5 shrink-0" />{err}
        </div>
      )}

      <button onClick={async () => { if (await post({ action: 'add', ...form })) { setMode(null); setForm({ name: '', dept: 'คลังสินค้า', allow_recorder: true, allow_compare: false, allow_invoice: false }); } }}
        disabled={!form.name.trim() || saving}
        className="w-full text-white font-bold text-[15px] rounded-xl disabled:opacity-40 flex items-center justify-center gap-2"
        style={{ minHeight: 52, background: '#0F172A' }}>
        {saving ? <><RefreshCw size={17} className="animate-spin" />กำลังบันทึก</> : <><Plus size={17} />เพิ่มเข้ารายชื่อ</>}
      </button>
    </div>
  );

  // ── เปิดดูรายคน ─────────────────────────────────────────────────
  if (opened) return (
    <div className="space-y-3">
      <button onClick={() => { setMode(null); setConfirmOff(null); }} className="text-[12px] font-semibold text-slate-500 hover:text-slate-800">‹ กลับรายชื่อ</button>

      <div className="bg-white border rounded-2xl overflow-hidden" style={{ borderColor: opened.active ? '#E4E6EA' : '#FDE68A' }}>
        <div className="px-4 py-4 flex items-center gap-3" style={{ background: opened.active ? '#F6F7F8' : '#FFFBEB' }}>
          <span className="shrink-0 rounded-full flex items-center justify-center text-xl font-bold text-white"
            style={{ width: 54, height: 54, background: opened.active ? '#35706A' : '#94A3B8' }}>
            {opened.initial || opened.name.charAt(0)}
          </span>
          <div className="min-w-0 flex-1">
            <div className="text-[17px] font-bold text-slate-800 truncate">{opened.name}</div>
            <div className="text-[11.5px] text-slate-500 mt-0.5">
              {opened.dept || 'ไม่ระบุแผนก'} · เข้าทำงาน {opened.joined_at ? new Date(opened.joined_at).toLocaleDateString('th-TH', { month: 'short', year: 'numeric' }) : '—'}
            </div>
          </div>
          {!opened.active && (
            <span className="shrink-0 text-[10px] font-bold px-2 py-1 rounded-lg" style={{ background: '#FDE68A', color: '#B45309' }}>ปิดใช้งาน</span>
          )}
        </div>
        <div className="p-3 grid grid-cols-2 gap-2">
          <div className="rounded-xl px-3 py-2.5" style={{ background: '#F6F7F8' }}>
            <div className="text-[10px] text-slate-500">ใบที่เคยส่ง</div>
            <div className="text-lg font-bold text-slate-800 tabular-nums">{opened.submission_count ?? 0}</div>
          </div>
          <div className="rounded-xl px-3 py-2.5" style={{ background: '#F6F7F8' }}>
            <div className="text-[10px] text-slate-500">ฟีเจอร์ที่เปิด</div>
            <div className="text-lg font-bold text-slate-800 tabular-nums">{FEAT_META.filter(f => opened[f.key]).length}</div>
          </div>
        </div>
      </div>

      <div className="bg-white border rounded-xl overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
        <div className="px-3 py-2.5 border-b text-[12px] font-bold text-slate-700" style={{ background: '#F6F7F8', borderColor: '#EEF0F3' }}>สิทธิ์การใช้งาน</div>
        <div className="divide-y" style={{ borderColor: '#F6F7F8' }}>
          {FEAT_META.map(f => {
            const on = !!opened[f.key];
            return (
              <button key={f.key} disabled={saving}
                onClick={() => post({ action: 'update', id: opened.id, [f.key]: !on })}
                className="w-full px-3 py-2.5 flex items-center gap-2.5 text-left disabled:opacity-60">
                <div className="flex-1 min-w-0">
                  <div className="text-[13px] font-semibold" style={{ color: on ? f.ink : '#334155' }}>{f.label}</div>
                  <div className="text-[10.5px] text-slate-400 truncate">{f.sub}</div>
                </div>
                <span className="shrink-0 rounded-full border-2 flex items-center justify-center"
                  style={{ width: 26, height: 26, background: on ? f.main : '#fff', borderColor: on ? f.main : '#CBD5E1' }}>
                  {on && <Check size={14} color="#fff" />}
                </span>
              </button>
            );
          })}
        </div>
      </div>

      {err && (
        <div className="rounded-xl px-3 py-2.5 text-[12px] flex items-start gap-2 border" style={{ background: '#FEF2F2', borderColor: '#FECACA', color: '#B91C1C' }}>
          <AlertCircle size={14} className="mt-0.5 shrink-0" />{err}
        </div>
      )}

      {opened.active ? (
        confirmOff === opened.id ? (
          <div className="space-y-2.5">
            <div className="rounded-2xl p-4 border-2" style={{ background: '#FFFBEB', borderColor: '#FDE68A' }}>
              <div className="text-[14.5px] font-bold" style={{ color: '#B45309' }}>ปิดการใช้งาน {opened.name}</div>
              <div className="text-[12px] mt-2 leading-relaxed" style={{ color: '#B45309' }}>
                • หายจากหน้าเลือกชื่อของทุกเครื่องทันที<br />
                • ใบ {opened.submission_count ?? 0} ใบที่เคยส่งยังอยู่ครบ ชื่อยังกำกับไว้<br />
                • เปิดกลับมาได้ทุกเมื่อ ไม่มีอะไรหาย
              </div>
            </div>
            <button onClick={async () => { if (await post({ action: 'setActive', id: opened.id, active: false })) setConfirmOff(null); }}
              disabled={saving}
              className="w-full text-white font-bold text-[14px] rounded-xl" style={{ minHeight: 50, background: '#B45309' }}>
              ยืนยันปิดการใช้งาน
            </button>
            <button onClick={() => setConfirmOff(null)}
              className="w-full border bg-white rounded-xl font-semibold text-[13.5px] text-slate-600"
              style={{ minHeight: 46, borderColor: '#E4E6EA' }}>ยกเลิก</button>
          </div>
        ) : (
          <button onClick={() => setConfirmOff(opened.id)}
            className="w-full border bg-white rounded-xl font-bold text-[13.5px]"
            style={{ minHeight: 48, borderColor: '#FDE68A', color: '#B45309' }}>ปิดการใช้งานคนนี้</button>
        )
      ) : (
        <button onClick={() => post({ action: 'setActive', id: opened.id, active: true })} disabled={saving}
          className="w-full text-white font-bold text-[14px] rounded-xl" style={{ minHeight: 50, background: '#35706A' }}>
          เปิดใช้งานกลับ
        </button>
      )}

      <div className="bg-white border rounded-xl p-3 text-[11px] text-slate-500 leading-relaxed" style={{ borderColor: '#E4E6EA' }}>
        ไม่มีปุ่มลบถาวร — ลบคนออกจะทำให้ใบเก่ากลายเป็นใบไม่มีเจ้าของ ตรวจย้อนหลังไม่ได้
      </div>
    </div>
  );

  // ── รายชื่อ ─────────────────────────────────────────────────────
  return (
    <div className="space-y-3">
      <div className="flex items-start justify-between gap-2">
        <div>
          <h2 className="text-xl font-bold text-slate-800">พนักงาน</h2>
          <p className="text-[12px] text-slate-500 mt-0.5">รายชื่อนี้คือหน้าเลือกชื่อที่พนักงานเห็น</p>
        </div>
        <button onClick={load} disabled={loading}
          className="shrink-0 rounded-lg flex items-center justify-center text-slate-500 border bg-white"
          style={{ width: 38, height: 38, borderColor: '#E4E6EA' }}>
          <RefreshCw size={15} className={loading ? 'animate-spin' : ''} />
        </button>
      </div>

      <div className="bg-white border rounded-xl flex overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
        {[['active', 'ใช้งานอยู่'], ['off', 'ปิดใช้งาน']].map(([k, label]) => {
          const cnt = (staff || []).filter(s => (k === 'active' ? s.active : !s.active)).length;
          const on = tab === k;
          return (
            <button key={k} onClick={() => setTab(k)} className="flex-1 py-2.5 text-[12.5px] font-bold border-b-2"
              style={on ? { color: '#0F172A', background: '#F6F7F8', borderColor: '#0F172A' } : { color: '#64748B', borderColor: 'transparent' }}>
              {label}{cnt > 0 && <span className="tabular-nums"> {cnt}</span>}
            </button>
          );
        })}
      </div>

      {err && (
        <div className="rounded-xl px-3 py-2.5 text-[12px] flex items-start gap-2 border" style={{ background: '#FEF2F2', borderColor: '#FECACA', color: '#B91C1C' }}>
          <AlertCircle size={14} className="mt-0.5 shrink-0" />{err}
        </div>
      )}

      {loading && !staff ? (
        <div className="bg-white rounded-xl px-4 py-9 text-center text-[13px] text-slate-400" style={{ border: '1px dashed #CBD5E1' }}>กำลังโหลดรายชื่อ…</div>
      ) : shown.length === 0 ? (
        <div className="bg-white rounded-xl px-4 py-9 text-center" style={{ border: '1px dashed #CBD5E1' }}>
          <User className="mx-auto text-[#E4E6EA] mb-2" size={32} />
          <div className="text-[13px] text-slate-500">{tab === 'active' ? 'ยังไม่มีพนักงานในรายชื่อ' : 'ไม่มีคนที่ปิดใช้งาน'}</div>
        </div>
      ) : (
        <div className="space-y-2">
          {shown.map(s => (
            <button key={s.id} onClick={() => setMode(s.id)}
              className="w-full bg-white border rounded-xl px-3 py-2.5 flex items-center gap-2.5 text-left hover:bg-[#F6F7F8]"
              style={{ borderColor: '#E4E6EA' }}>
              <span className="shrink-0 rounded-full flex items-center justify-center text-[15px] font-bold text-white"
                style={{ width: 42, height: 42, background: s.active ? '#35706A' : '#94A3B8' }}>
                {s.initial || s.name.charAt(0)}
              </span>
              <div className="flex-1 min-w-0">
                <div className="text-[14px] font-bold text-slate-800 truncate">{s.name}</div>
                <div className="flex items-center gap-1 mt-1 flex-wrap">
                  <span className="text-[10px] px-1.5 py-0.5 rounded-full text-slate-600" style={{ background: '#F6F7F8' }}>{s.dept || 'ไม่ระบุ'}</span>
                  {FEAT_META.filter(f => s[f.key]).map(f => (
                    <span key={f.key} className="text-[10px] px-1.5 py-0.5 rounded-full font-semibold"
                      style={{ background: f.soft, color: f.ink }}>{f.label}</span>
                  ))}
                </div>
              </div>
              <div className="text-right shrink-0">
                <div className="text-[13px] font-bold text-slate-800 tabular-nums">{s.submission_count ?? 0}</div>
                <div className="text-[9.5px] text-slate-400">ใบ</div>
              </div>
            </button>
          ))}
        </div>
      )}

      <button onClick={() => { setMode('add'); setErr(''); }}
        className="w-full text-white font-bold text-[15px] rounded-xl flex items-center justify-center gap-2"
        style={{ minHeight: 52, background: '#0F172A' }}><Plus size={18} />เพิ่มพนักงาน</button>

      <div className="bg-white border rounded-xl p-3 text-[11px] text-slate-500 leading-relaxed" style={{ borderColor: '#E4E6EA' }}>
        ชื่อที่ปิดใช้งานจะหายจากหน้าเลือกชื่อทันที แต่ใบเก่ายังมีชื่อกำกับไว้ครบ — ไม่มีปุ่มลบถาวร
      </div>
    </div>
  );
}

function SettingsView({ config, onSave, onTestConnection, dataSource, lastSyncAt, productCount }) {
  const [url, setUrl] = useState(config.url||'');
  const [anonKey, setAnonKey] = useState(config.anonKey||'');
  const [tableName, setTableName] = useState(config.tableName||'product_price');
  const [stockTableName, setStockTableName] = useState(config.stockTableName||'product_stock');
  const [showKey, setShowKey] = useState(false);
  const [testing, setTesting] = useState(false); const [testResult, setTestResult] = useState(null);
  const [saveMsg, setSaveMsg] = useState('');
  const cfg = { url: url.trim(), anonKey: anonKey.trim(), tableName: tableName.trim()||'product_price', stockTableName: stockTableName.trim()||'product_stock' };
  const handleSave = async () => { await onSave(cfg); setSaveMsg('บันทึกแล้ว'); setTimeout(()=>setSaveMsg(''), 2000); };
  const handleTest = async () => { setTesting(true); setTestResult(null); try { await handleSave(); const count = await onTestConnection(cfg); setTestResult({ok:true,msg:`เชื่อมต่อสำเร็จ! มี ${count.toLocaleString()} แถว ใน ${cfg.tableName}`}); } catch(e) { setTestResult({ok:false,msg:e.message}); } setTesting(false); };
  return (
    <div className="space-y-4">
      <div><h2 className="text-2xl font-bold text-slate-800">ตั้งค่า Supabase</h2></div>
      <div className={`rounded-xl p-4 border ${dataSource==='supabase'?'bg-[#EAF1F0] border-[#B6D0CC]':'bg-[#F6F7F8] border-[#E4E6EA]'}`}>
        <div className="flex items-center gap-3"><div className={`p-2 rounded-lg text-white ${dataSource==='supabase'?'bg-[#35706A]':'bg-slate-400'}`}><Cloud size={20}/></div><div><div className="font-semibold text-slate-800">{dataSource==='supabase'?'Supabase (เชื่อมต่อแล้ว)':'ยังไม่ได้เชื่อมต่อ'}</div><div className="text-xs text-slate-500">{productCount.toLocaleString()} รายการ cache{lastSyncAt&&dataSource==='supabase'&&` • ${new Date(lastSyncAt).toLocaleString('th-TH',{dateStyle:'short',timeStyle:'short'})}`}</div></div></div>
      </div>
      <div className="bg-white rounded-xl border border-[#E4E6EA] p-4 space-y-3">
        <div><label className="text-sm font-medium text-slate-700 mb-1 block">Supabase URL</label><input type="text" value={url} onChange={e=>{setUrl(e.target.value);setTestResult(null);}} placeholder="https://xxxxx.supabase.co" className="w-full px-3 py-2 border border-[#E4E6EA] rounded-lg outline-none focus:ring-2 focus:ring-[#35706A] font-mono text-sm"/></div>
        <div><label className="text-sm font-medium text-slate-700 mb-1 block">Anon Key</label><div className="relative"><input type={showKey?'text':'password'} value={anonKey} onChange={e=>{setAnonKey(e.target.value);setTestResult(null);}} placeholder="eyJ..." className="w-full px-3 py-2 pr-10 border border-[#E4E6EA] rounded-lg outline-none focus:ring-2 focus:ring-[#35706A] font-mono text-sm"/><button onClick={()=>setShowKey(!showKey)} className="absolute right-2 top-1/2 -translate-y-1/2 p-1">{showKey?<EyeOff size={16}/>:<Eye size={16}/>}</button></div></div>
        <div className="grid grid-cols-2 gap-3">
          <div><label className="text-sm font-medium text-slate-700 mb-1 block">ตาราง Product <span className="text-xs text-slate-400">(checkBarcode)</span></label><input type="text" value={tableName} onChange={e=>setTableName(e.target.value)} placeholder="product_price" className="w-full px-3 py-2 border border-[#E4E6EA] rounded-lg outline-none focus:ring-2 focus:ring-[#35706A] font-mono text-sm"/></div>
          <div><label className="text-sm font-medium text-slate-700 mb-1 block">ตาราง Stock <span className="text-xs text-slate-400">(compare)</span></label><input type="text" value={stockTableName} onChange={e=>setStockTableName(e.target.value)} placeholder="product_stock" className="w-full px-3 py-2 border border-[#E4E6EA] rounded-lg outline-none focus:ring-2 focus:ring-[#35706A] font-mono text-sm"/></div>
        </div>
        {testResult&&<div className={`rounded-lg p-3 text-sm flex items-start gap-2 ${testResult.ok?'bg-[#F0FDF4] border border-[#BBF7D0] text-[#15803D]':'bg-[#FEF2F2] border border-[#FECACA] text-[#B91C1C]'}`}>{testResult.ok?<CheckCircle2 size={16} className="flex-shrink-0 mt-0.5"/>:<XCircle size={16} className="flex-shrink-0 mt-0.5"/>}<span>{testResult.msg}</span></div>}
        <div className="flex gap-2">
          <button onClick={handleTest} disabled={testing||!url||!anonKey} className="flex-1 py-2.5 bg-slate-700 hover:bg-slate-800 disabled:bg-[#E4E6EA] text-white rounded-lg text-sm font-medium flex items-center justify-center gap-1.5">{testing?<RefreshCw size={14} className="animate-spin"/>:<Zap size={14}/>}{testing?'กำลังทดสอบ...':'ทดสอบการเชื่อมต่อ'}</button>
          <button onClick={handleSave} disabled={!url||!anonKey} className="flex-1 py-2.5 bg-[#35706A] hover:bg-[#2A5A55] disabled:bg-[#E4E6EA] text-white rounded-lg text-sm font-medium"><Save size={14} className="inline mr-1"/>บันทึก</button>
        </div>
        {saveMsg&&<div className="bg-[#F0FDF4] border border-[#BBF7D0] text-[#15803D] text-sm rounded-lg p-2 text-center">{saveMsg}</div>}
      </div>
    </div>
  );
}

function StepBar({ current }) {
  return (
    <div className="flex items-center justify-between mb-6">
      {STEPS.map((label, i) => {
        const n = i+1, done = current > n, active = current === n;
        return (
          <React.Fragment key={n}>
            <div className="flex flex-col items-center gap-1">
              <div className={`w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold border-2 ${done?'bg-[#0F172A] border-[#0F172A] text-white':active?'bg-white border-[#0F172A] text-[#0F172A]':'bg-white border-[#E4E6EA] text-slate-400'}`}>{done?'✓':n}</div>
              <span className={`text-[10px] whitespace-nowrap ${active?'text-[#0F172A] font-semibold':'text-slate-400'}`}>{label}</span>
            </div>
            {i < STEPS.length-1 && <div className={`flex-1 h-0.5 mx-1 mt-[-12px] ${current > n+1?'bg-[#0F172A]':current > n?'bg-[#E4E6EA]':'bg-[#E4E6EA]'}`}/>}
          </React.Fragment>
        );
      })}
    </div>
  );
}

function DropZone({ onFiles, multiple, accept, children }) {
  const [drag, setDrag] = useState(false); const ref = useRef();
  const handle = fs => { if (fs?.length) onFiles(Array.from(fs)); };
  return (
    <div ref={ref} onClick={() => ref.current.querySelector('input')?.click()}
      onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
      onDrop={e=>{e.preventDefault();setDrag(false);handle(e.dataTransfer.files);}}
      className={`border-2 border-dashed rounded-xl p-6 cursor-pointer transition-colors text-center ${drag?'border-[#94A3B8] bg-[#F6F7F8]':'border-[#E4E6EA] hover:border-[#94A3B8] hover:bg-[#F6F7F8]'}`}>
      <input type="file" multiple={multiple} accept={accept} className="hidden" onChange={e=>handle(e.target.files)}/>
      {children}
    </div>
  );
}

function InvFileThumb({ file }) {
  const [url, setUrl] = useState('');
  useEffect(() => { const u = URL.createObjectURL(file); setUrl(u); return () => URL.revokeObjectURL(u); }, [file]);
  return url ? <img src={url} alt="" style={{ width: 56, height: 56, objectFit: 'cover', borderRadius: 6, flexShrink: 0 }}/> : <div style={{ width: 56, height: 56, background: '#f1f5f9', borderRadius: 6 }}/>;
}

function Spinner({ size = 16 }) {
  return <span style={{ display:'inline-block', width:size, height:size, borderRadius:'50%', border:'2px solid #e5e7eb', borderTopColor:'#111', animation:'spin 0.7s linear infinite', flexShrink:0 }}/>;
}

function InvoiceScannerModule({ supabaseConfig }) {
  const w = useWinWidth(), mob = w < 600;
  const [model, setModel] = useState('claude-sonnet-4-6');
  const [step, setStep] = useState(1);
  const [invoices, setInvoices] = useState([]);
  const [productFiles, setPFiles] = useState([]);
  const [selPFileIds, setSelPFIds] = useState(new Set());
  const [barcodeMap, setBMap] = useState({});
  const [scanResults, setSRes] = useState([]);
  const [scanning, setScanning] = useState(false);
  const [fileName, setFileName] = useState('');
  const [driveStatus, setDSt] = useState(null);
  const [driveUrl, setDUrl] = useState(null);
  const [driveErr, setDErr] = useState(null);
  const [sbSt, setSbSt] = useState(null);
  const [sbErr, setSbErr] = useState(null);
  const [selectedPages, setSelPages] = useState({});
  const [gReady, setGReady] = useState(false);
  const [gToken, setGToken] = useState(null);
  const [gClientId, setGClientId] = useState(() => safeGet('g_client', '') || (typeof import.meta !== 'undefined' ? (import.meta.env?.VITE_GOOGLE_CLIENT_ID || '') : ''));
  const [driveFolder, setDriveFolder] = useState(() => safeGet('drive_folder', '') || (typeof import.meta !== 'undefined' ? (import.meta.env?.VITE_DRIVE_FOLDER_ID || DRIVE_FOLDER_DEFAULT) : DRIVE_FOLDER_DEFAULT));
  const [showInvCfg, setShowInvCfg] = useState(false);
  const [nameStatus, setNameStatus] = useState(null);

  const { url: sbUrl, anonKey: sbKey } = supabaseConfig;

  useEffect(() => {
    if (window.google?.accounts) { setGReady(true); return; }
    const s = document.createElement('script'); s.src = 'https://accounts.google.com/gsi/client'; s.async = true; s.defer = true; s.onload = () => setGReady(true); document.head.appendChild(s);
    return () => { try { document.head.removeChild(s); } catch {} };
  }, []);

  const connectGoogle = () => {
    if (!gClientId) { alert('กรุณาใส่ Google Client ID ก่อน'); setShowInvCfg(true); return; }
    if (!gReady || !window.google?.accounts) { alert('Google library ยังไม่โหลด'); return; }
    const client = window.google.accounts.oauth2.initTokenClient({ client_id: gClientId, scope: 'https://www.googleapis.com/auth/drive.file', callback: resp => { if (resp.access_token) setGToken(resp.access_token); else if (resp.error) setDErr('OAuth: '+resp.error); }, error_callback: err => setDErr('OAuth error: '+(err.message||err.type)) });
    client.requestAccessToken();
  };
  const disconnectGoogle = () => { if (gToken && window.google?.accounts) window.google.accounts.oauth2.revoke(gToken, () => {}); setGToken(null); };

  const uploadFileToDrive = async (filename, blob, mimeType) => {
    const metadata = { name: filename, mimeType };
    if (driveFolder) metadata.parents = [driveFolder];
    const form = new FormData(); form.append('metadata', new Blob([JSON.stringify(metadata)], {type:'application/json'})); form.append('file', blob);
    const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink,name', { method:'POST', headers:{Authorization:`Bearer ${gToken}`}, body:form });
    if (!res.ok) { const t = await res.text(); if (res.status === 401) { setGToken(null); throw new Error('Token หมดอายุ — กรุณา connect ใหม่'); } throw new Error(`Drive ${res.status}: ${t.slice(0,150)}`); }
    return res.json();
  };

  const HEADER_KEYS = ['invoice_no','invoice_date','vendor_name','vendor_tax_id','document_type','vendor_address','vendor_branch','vendor_no','price_type','_vendorFromDB'];
  const mergePageData = (pagesData) => {
    let merged = {}, allProducts = [];
    for (const pd of pagesData) { if (!pd) continue; for (const k of HEADER_KEYS) if (merged[k]==null&&pd[k]!=null) merged[k]=pd[k]; allProducts=[...allProducts,...(pd.products||[])]; }
    if (!allProducts.length && !Object.keys(merged).length) return null;
    const pt = merged.price_type ?? 'incl';
    merged.products = allProducts.map((p,i) => recalc({...p, no:i+1, _pt:pt}));
    if (merged.invoice_date) merged.invoice_date = toYMD(merged.invoice_date)||merged.invoice_date;
    return merged;
  };

  const readSinglePage = async (file) => {
    const { base64, mediaType } = await imgToBase64(file);
    const content = [{type:'image',source:{type:'base64',media_type:mediaType,data:base64}},{type:'text',text:INVOICE_PROMPT}];
    const res = await callClaude(content, {}, model);
    const tb = res.content?.find(b => b.type === 'text');
    let data;
    try { data = extractJSON(tb?.text||''); } catch { const r2 = await callClaude([content[0],{type:'text',text:INVOICE_PROMPT+'\n\nตอบ JSON เท่านั้น:'}],{},model); data = extractJSON(r2.content?.find(b=>b.type==='text')?.text||''); }
    if (data.invoice_date) data.invoice_date = toYMD(data.invoice_date)||null;
    data.vendor_no = null;
    if (data.vendor_name && sbUrl && sbKey) { try { const no = await lookupVendorREST(sbUrl, sbKey, data.vendor_name); if (no) { data.vendor_no=no; data._vendorFromDB=true; } } catch {} }
    return data;
  };

  const updateGroupPage = (gi, pi, pageData, status) => {
    setInvoices(prev => {
      const n=[...prev], g={...n[gi]};
      const pagesData=[...g.pagesData]; pagesData[pi]=pageData;
      const pageStatus=[...g.pageStatus]; pageStatus[pi]=status;
      const merged=mergePageData(pagesData);
      const allDone=pageStatus.every(s=>s==='done'||s==='error');
      n[gi]={...g,pagesData,pageStatus,data:merged,status:allDone?(merged?'done':'error'):'processing'};
      return n;
    });
  };

  const readPages = async (gi, files, pageIndices) => {
    setInvoices(prev => { const n=[...prev],g={...n[gi]},ps=[...g.pageStatus]; pageIndices.forEach(pi=>ps[pi]='processing'); n[gi]={...g,pageStatus:ps,status:'processing'}; return n; });
    await Promise.all(pageIndices.map(async pi => { try { const data=await readSinglePage(files[pi]); updateGroupPage(gi,pi,data,'done'); } catch { updateGroupPage(gi,pi,null,'error'); } }));
  };

  const makeGroup = (files) => ({ files, pagesData:files.map(()=>null), pageStatus:files.map(()=>'pending'), status:'pending', data:null, error:null });
  const addFiles = fs => { const arr=Array.from(fs); setInvoices(prev=>[...prev,makeGroup(arr)]); };

  const removeInv = i => { setInvoices(prev=>prev.filter((_,j)=>j!==i)); setSelPages(prev=>{const n={...prev};delete n[i];return n;}); };
  const removePage = (gi, pi) => {
    setInvoices(prev => {
      const n = [...prev];
      const inv = {...n[gi]};
      const files = inv.files.filter((_,i)=>i!==pi);
      if (files.length === 0) return n.filter((_,i)=>i!==gi);
      n[gi] = { ...inv, files, pagesData: inv.pagesData.filter((_,i)=>i!==pi), pageStatus: inv.pageStatus.filter((_,i)=>i!==pi) };
      return n;
    });
  };
  const updateData = (i, data) => setInvoices(prev => { const n=[...prev]; n[i]={...n[i],data}; return n; });
  const reprocessInvoice = async (gi) => { await readPages(gi, invoices[gi].files, invoices[gi].files.map((_,i)=>i)); };

  const processAll = async () => {
    setInvoices(prev=>prev.map(g=>({...g,status:g.pageStatus.some(s=>s==='pending')?'processing':g.status,pageStatus:g.pageStatus.map(s=>s==='pending'?'processing':s)})));
    await Promise.all(invoices.map(async (inv,gi) => { const pending=inv.pagesData.map((d,i)=>d===null?i:-1).filter(i=>i>=0); if(pending.length===0)return; await readPages(gi,inv.files,pending); }));
    setStep(2);
  };

  const mkPFileId = () => Math.random().toString(36).slice(2,10);
  const buildMap = (results) => { const map={}; results.forEach(r=>{if(r.match&&r.barcode)map[String(r.match).trim()]=r.barcode;}); return map; };
  const applyBarcodeMap = (newMap) => {
    const normMap={};
    Object.entries(newMap).forEach(([k,v])=>{if(k&&v)normMap[String(k).trim()]=v;});
    const mappedBarcodes = new Set(Object.values(normMap));
    setInvoices(prev=>prev.map(inv=>{
      if(!inv.data)return inv;
      const products=(inv.data.products||[]).map(p=>{
        const desc=String(p.description||'').trim();
        if(normMap[desc]) return {...p, barcode: normMap[desc]};
        // Clear barcode if it was previously assigned via map but now belongs to a different description
        if(p.barcode && mappedBarcodes.has(p.barcode)) return {...p, barcode: null};
        return p;
      });
      return{...inv,data:{...inv.data,products}};
    }));
    setBMap(normMap);
  };

  const scanProductItems = async (items) => {
    if (!items.length) return;
    const scanIds = new Set(items.map(it=>it.id));
    setPFiles(prev=>prev.map(it=>scanIds.has(it.id)?{...it,status:'processing'}:it));
    setScanning(true);
    const allP = invoices.filter(i=>i.status==='done'&&i.data?.products).flatMap(i=>i.data.products).filter((p,i,a)=>a.findIndex(x=>x.description===p.description)===i);
    const list = allP.map(p=>p.description).filter(Boolean).join('\n');
    // Keep rows from non-scanning files, AND user-override rows from scanning files
    const kept = scanResults.filter(r=>!scanIds.has(r._fileId) || r._userOverride);
    const newResults = await Promise.all(items.map(async it => {
      try {
        const { base64, mediaType } = await imgToBase64(it.file);
        const res = await callClaude([{type:'image',source:{type:'base64',media_type:mediaType,data:base64}},{type:'text',text:BARCODE_PROMPT(list)}],{},model);
        const tb = res.content?.find(b=>b.type==='text');
        const raw = extractJSON(tb?.text||'null');
        const arr = Array.isArray(raw)?raw:(raw?[raw]:[{barcode:null,match:null,description_image:null}]);
        const expanded = arr.flatMap(r=>{
          const barcodes=(r.barcode||'').split(/[,，\n]/).map(s=>s.trim()).filter(Boolean);
          const descs=(r.description_image||'').split(/[,，\n]/).map(s=>s.trim()).filter(Boolean);
          const matches=(r.match||'').split(/[,，\n]/).map(s=>s.trim()).filter(Boolean);
          if(barcodes.length<=1)return[r];
          return barcodes.map((bc,i)=>({barcode:bc||null,description_image:descs[i]||descs[0]||null,match:matches[i]||null}));
        });
        setPFiles(prev=>prev.map(p=>p.id===it.id?{...p,status:'done'}:p));
        return expanded.map(r=>({...r,_fileId:it.id}));
      } catch { setPFiles(prev=>prev.map(p=>p.id===it.id?{...p,status:'error'}:p)); return [{barcode:null,match:null,description_image:null,_fileId:it.id}]; }
    }));
    const combined = [...kept,...newResults.flat()];
    setSRes(combined); applyBarcodeMap(buildMap(combined)); setScanning(false);
  };

  const addProductFiles = (newFiles) => { const items=Array.from(newFiles).map(f=>({file:f,id:mkPFileId(),status:'pending'})); setPFiles(prev=>[...prev,...items]); };
  const scanPendingPFiles = async () => { const items=productFiles.filter(it=>it.status==='pending'); if(items.length===0)return; await scanProductItems(items); };
  const deleteProductFile = (id) => { setPFiles(prev=>prev.filter(it=>it.id!==id)); setSelPFIds(prev=>{const s=new Set(prev);s.delete(id);return s;}); const remaining=scanResults.filter(r=>r._fileId!==id); setSRes(remaining); applyBarcodeMap(buildMap(remaining)); };
  const rescanSelectedPFiles = async () => { const items=productFiles.filter(it=>selPFileIds.has(it.id)); setSelPFIds(new Set()); await scanProductItems(items); };

  const buildXLSXBlob = () => {
    if (!XLSX) return null;
    const wb = XLSX.utils.book_new();
    const done = invoices.filter(i=>i.status==='done'&&i.data);
    const str = v => v!=null?{t:'s',v:String(v).trim()}:null;

    // Sheet 1: bill_header (14 cols per spec)
    const ws1 = XLSX.utils.aoa_to_sheet([
      ['invoice_no','invoice_date','vendor_name','vendor_tax_id','document_type','vendor_address','total_amount','total_discount','net_total','excl_vat','vat_amount','vendor_branch','vendor_no','price_type'],
      ...done.map(inv=>{const d=inv.data,rawAmt=(d.products||[]).reduce((s,p)=>s+(+p.amount||0),0),vs=vatSummary(d.products);return[str(d.invoice_no),d.invoice_date??null,d.vendor_name??null,str(d.vendor_tax_id),d.document_type??null,d.vendor_address??null,+rawAmt.toFixed(2)||null,vs.sdTot,vs.netTotal,vs.excl,vs.vatAmt,str(d.vendor_branch),str(d.vendor_no),d.price_type??'incl'];})
    ]);
    XLSX.utils.book_append_sheet(wb, ws1, 'bill_header');

    // Sheet 2: invoice (17 cols per spec)
    const ws2 = XLSX.utils.aoa_to_sheet([
      ['invoice_no','no','description','carton_size','carton','ea','qty','price_ea','amount','special_discount','amount_sd','total','diff','excl_vat','vat_amt','vat','barcode'],
      ...done.flatMap(inv=>(inv.data.products||[]).map(p=>{
        return[
          str(inv.data.invoice_no), p.no??null, p.description??null,
          p.carton_size!=null?+p.carton_size:null,
          p.carton!=null?+p.carton:null,
          p.ea!=null?+p.ea:null,
          p.qty!=null?+p.qty:null,
          p.price_ea!=null?+p.price_ea:null,
          p.amount!=null?+p.amount:null,
          p.special_discount!=null?+p.special_discount:null,
          p.amount_sd!=null?+p.amount_sd:null,
          p.total!=null?+p.total:null,
          p.diff!=null?+p.diff:null,
          p.excl_vat!=null?+p.excl_vat:null,
          p.vat_amt!=null?+p.vat_amt:null,
          p.vat??null,
          str(p.barcode??barcodeMap[String(p.description||'').trim()]??null)
        ];
      }))
    ]);
    XLSX.utils.book_append_sheet(wb, ws2, 'invoice');

    // Sheet 3: product (only if scan results exist)
    if (scanResults.length > 0) {
      const ws3 = XLSX.utils.aoa_to_sheet([
        ['no','barcode','match','description_image'],
        ...scanResults.map((r,i)=>[i+1, str(r.barcode), str(r.match), str(r.description_image)])
      ]);
      XLSX.utils.book_append_sheet(wb, ws3, 'product');
    }

    const buf = XLSX.write(wb, {type:'array',bookType:'xlsx'});
    return new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  };

  const buildImpDataCSV = () => {
    const done = invoices.filter(i=>i.status==='done'&&i.data);
    const rows=[];
    for (const inv of done) { for (const p of (inv.data.products||[])) { const qty=p.qty!=null?+p.qty:0,pea=p.price_ea!=null?+p.price_ea:null,sd=p.special_discount!=null?+p.special_discount:0; const tot=(qty>0&&pea!=null)?+(qty*pea-sd).toFixed(2):null; const totPerQty=(tot!=null&&qty>0)?+(tot/qty).toFixed(4):0; const barcode=p.barcode??barcodeMap[String(p.description||'').trim()]??''; rows.push([barcode,qty,totPerQty,0]); } }
    const csv=rows.map(r=>r.join(',')).join('\r\n');
    return new Blob(['﻿'+csv],{type:'text/csv;charset=utf-8'});
  };

  const buildBillHeaderCSV = () => {
    const done = invoices.filter(i=>i.status==='done'&&i.data);
    const rows = done.map(inv=>{const d=inv.data,rawAmt=(d.products||[]).reduce((s,p)=>s+(+p.amount||0),0),vs=vatSummary(d.products);return[d.invoice_no??'',d.invoice_date??'',d.vendor_name??'',d.vendor_tax_id??'',+rawAmt.toFixed(2)||0,vs.netTotal,vs.vatAmt].join(',');});
    return new Blob(['﻿'+'invoice_no,invoice_date,vendor_name,vendor_tax_id,total_amount,net_total,vat_amount\r\n'+rows.join('\r\n')],{type:'text/csv;charset=utf-8'});
  };

  const uploadToDrive = async () => {
    if (!gToken) { alert('กรุณาเชื่อมต่อ Google Drive ก่อน'); return; }
    setDSt('uploading'); setDUrl(null); setDErr(null);
    try {
      const fname = fileName||'invoice';
      const r1 = await uploadFileToDrive(fname+'_bill_header.csv', buildBillHeaderCSV(), 'text/csv');
      await uploadFileToDrive(fname+'_imp_data.csv', buildImpDataCSV(), 'text/csv');
      setDUrl(r1.webViewLink); setDSt('done');
    } catch(e) { setDErr(e.message); setDSt('error'); }
  };

  const saveToSupabase = async () => {
    if (!sbUrl || !sbKey) { setSbErr('ยังไม่ได้ตั้งค่า Supabase — ไปที่ ตั้งค่า ก่อน'); setSbSt('error'); return; }
    setSbSt('saving'); setSbErr(null);
    const fn = fileName||null;
    const h = { apikey:sbKey, Authorization:`Bearer ${sbKey}`, 'Content-Type':'application/json' };
    try {
      const done = invoices.filter(i=>i.status==='done'&&i.data);
      for (const inv of done) {
        const d=inv.data, invNo=d.invoice_no??null;
        if (!invNo||!fn) continue;
        const rawAmt=(d.products||[]).reduce((s,p)=>s+(+p.amount||0),0), vs=vatSummary(d.products);
        const r1 = await fetch(`${sbUrl}/rest/v1/bill_header`,{method:'POST',headers:{...h,Prefer:'resolution=merge-duplicates,return=minimal'},body:JSON.stringify({invoice_no:invNo,file_name:fn,invoice_date:d.invoice_date??null,vendor_name:d.vendor_name??null,vendor_tax_id:d.vendor_tax_id?String(d.vendor_tax_id).replace(/\s/g,''):null,document_type:d.document_type??null,vendor_address:d.vendor_address??null,vendor_branch:d.vendor_branch?String(d.vendor_branch):null,vendor_no:d.vendor_no?String(d.vendor_no):null,price_type:d.price_type??'incl',total_amount:+rawAmt.toFixed(2)||null,net_total:vs.netTotal,excl_vat:vs.excl,vat_amount:vs.vatAmt})});
        if (!r1.ok) throw new Error('insert bill_header: '+(await r1.text()));
        for (const p of (d.products||[])) {
          const qty=p.qty!=null?+p.qty:null,pea=p.price_ea!=null?+p.price_ea:null,am=p.amount!=null?+p.amount:null,sd=p.special_discount!=null?+p.special_discount:0;
          const tot=(qty!=null&&pea!=null)?+(qty*pea-sd).toFixed(2):null;
          const vatV=p.vat==='v',pt=d.price_type??'incl';
          const exclV=tot!=null?(vatV?(pt==='incl'?+(tot/1.07).toFixed(2):tot):tot):null;
          const vatAmtV=tot!=null?(vatV?(pt==='incl'?+(tot-tot/1.07).toFixed(2):+(tot*0.07).toFixed(2)):0):null;
          const r2 = await fetch(`${sbUrl}/rest/v1/imp_data`,{method:'POST',headers:{...h,Prefer:'resolution=merge-duplicates,return=minimal'},body:JSON.stringify({invoice_no:invNo,file_name:fn,no:p.no??null,description:p.description??null,qty,price_ea:pea,amount:am,special_discount:sd||null,total:tot,excl_vat:exclV,vat_amt:vatAmtV,vat:p.vat??null,barcode:p.barcode??barcodeMap[String(p.description||'').trim()]??null})});
          if (!r2.ok) throw new Error('insert imp_data: '+(await r2.text()));
        }
      }
      setSbSt('done');
    } catch(e) { setSbErr(e.message); setSbSt('error'); }
  };

  const reset = () => { setStep(1);setInvoices([]);setPFiles([]);setBMap({});setSRes([]);setSelPFIds(new Set());setFileName('');setDSt(null);setDUrl(null);setDErr(null);setSbSt(null);setSbErr(null);setSelPages({}); };

  useEffect(() => {
    if (step===4&&!fileName&&sbUrl&&sbKey) {
      (async()=>{
        const d=new Date(),yy=String(d.getFullYear()).slice(-2),mm=String(d.getMonth()+1).padStart(2,'0'),dd=String(d.getDate()).padStart(2,'0'),dateKey=yy+mm+dd;
        try {
          const r=await fetch(`${sbUrl}/rest/v1/rpc/get_next_filename`,{method:'POST',headers:{apikey:sbKey,Authorization:`Bearer ${sbKey}`,'Content-Type':'application/json'},body:JSON.stringify({date_prefix:dateKey})});
          if(r.ok){const name=await r.json();if(typeof name==='string'&&name.length===10){setFileName(name);return;}}
        } catch {}
        const cur=parseInt(safeGet('filename_counter_'+dateKey,'0'))||0;
        safeSet('filename_counter_'+dateKey,String(cur+1));
        setFileName(dateKey+String(cur+1).padStart(4,'0'));
      })();
    }
  }, [step]);

  useEffect(() => {
    if (step!==4||!fileName){setNameStatus(null);return;}
    if (!sbUrl||!sbKey){return;}
    setNameStatus('checking');
    const t=setTimeout(async()=>{
      try{const r=await fetch(`${sbUrl}/rest/v1/rpc/check_filename_exists`,{method:'POST',headers:{apikey:sbKey,Authorization:`Bearer ${sbKey}`,'Content-Type':'application/json'},body:JSON.stringify({filename:fileName})});if(!r.ok){setNameStatus('error');return;}const exists=await r.json();setNameStatus(exists?'duplicate':'available');}catch{setNameStatus('error');}
    },500);
    return ()=>clearTimeout(t);
  },[fileName,step,sbUrl,sbKey]);

  const doneInvs = invoices.filter(i=>i.status==='done'&&i.data);
  const processing = invoices.some(i=>i.status==='processing');
  const allProds = doneInvs.flatMap(inv=>inv.data.products.map(p=>({...p,invoice_no:inv.data.invoice_no,_pt:inv.data.price_type??'incl'})));
  const grandTotal = allProds.reduce((s,p)=>s+(+p.amount||0),0);
  const allVs = vatSummary(allProds);

  return (
    <div style={{ maxWidth: 720, margin: '0 auto', padding: mob?8:0 }}>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:16, flexWrap:'wrap', gap:8 }}>
        <h2 style={{ fontWeight:700, fontSize:20, color:'#1e293b', margin:0 }}>สแกนใบกำกับภาษี</h2>
        <div style={{ display:'flex', gap:8, alignItems:'center', flexWrap:'wrap' }}>
          <select value={model} onChange={e=>setModel(e.target.value)} style={{ fontSize:12, padding:'4px 8px', borderRadius:6, border:'1px solid #cbd5e1', background:'#f8fafc', color:'#475569' }}>
            {MODELS.map(m=><option key={m.id} value={m.id}>{m.label}</option>)}
          </select>
          <button onClick={()=>setShowInvCfg(!showInvCfg)} style={{ fontSize:12, padding:'4px 10px', borderRadius:6, border:'1px solid #cbd5e1', background:'#f8fafc', cursor:'pointer' }}>⚙ ตั้งค่า</button>
          {step>1&&<button onClick={reset} style={{ fontSize:12, padding:'4px 10px', borderRadius:6, border:'1px solid #fca5a5', background:'#fef2f2', color:'#dc2626', cursor:'pointer' }}>↺ ใหม่</button>}
        </div>
      </div>

      {showInvCfg&&(
        <div style={{ background:'#f8fafc', border:'1px solid #e2e8f0', borderRadius:10, padding:16, marginBottom:16 }}>
          <div style={{ fontWeight:600, marginBottom:10 }}>ตั้งค่า Google Drive</div>
          <div style={{ marginBottom:8 }}><label style={{ fontSize:12, color:'#64748b' }}>Google Client ID</label><input value={gClientId} onChange={e=>setGClientId(e.target.value)} onBlur={()=>safeSet('g_client',gClientId)} placeholder="xxx.apps.googleusercontent.com" style={{ width:'100%', marginTop:4, padding:'6px 10px', borderRadius:6, border:'1px solid #cbd5e1', fontSize:12, fontFamily:'monospace', boxSizing:'border-box' }}/></div>
          <div style={{ marginBottom:8 }}><label style={{ fontSize:12, color:'#64748b' }}>Drive Folder ID</label><input value={driveFolder} onChange={e=>setDriveFolder(e.target.value)} onBlur={()=>safeSet('drive_folder',driveFolder)} placeholder={DRIVE_FOLDER_DEFAULT} style={{ width:'100%', marginTop:4, padding:'6px 10px', borderRadius:6, border:'1px solid #cbd5e1', fontSize:12, fontFamily:'monospace', boxSizing:'border-box' }}/></div>
          <button onClick={()=>setShowInvCfg(false)} style={{ fontSize:12, padding:'6px 16px', borderRadius:6, background:'#0f172a', color:'#fff', border:'none', cursor:'pointer' }}>บันทึก</button>
        </div>
      )}

      <StepBar current={step}/>

      {step===1&&(
        <div style={{ display:'flex', flexDirection:'column', gap:16 }}>
          <DropZone multiple accept="image/*,.pdf" onFiles={addFiles}>
            <div style={{ fontSize:32, marginBottom:8 }}>📄</div>
            <div style={{ fontWeight:600, color:'#374151', marginBottom:4 }}>ลากหรือคลิกเพื่ออัปโหลดใบกำกับภาษี</div>
            <div style={{ fontSize:12, color:'#9ca3af' }}>รองรับ JPG, PNG, PDF (หลายไฟล์)</div>
          </DropZone>
          {invoices.length>0&&(
            <div>
              {invoices.map((inv,gi)=>(
                <div key={gi} style={{ background:'#fff', border:'1px solid #e2e8f0', borderRadius:10, padding:12, marginBottom:10 }}>
                  <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:8 }}>
                    <div style={{ fontWeight:600, fontSize:13, color:'#374151' }}>กลุ่ม {gi+1} • {inv.files.length} หน้า</div>
                    <button onClick={()=>removeInv(gi)} style={{ color:'#ef4444', background:'none', border:'none', cursor:'pointer', fontSize:12 }}>✕ ลบ</button>
                  </div>
                  <div style={{ display:'flex', gap:8, flexWrap:'wrap' }}>
                    {inv.files.map((f,pi)=>(
                      <div key={pi} style={{ position:'relative' }}>
                        <InvFileThumb file={f}/>
                        <div style={{ position:'absolute', bottom:2, right:2, fontSize:9, background:'rgba(0,0,0,0.6)', color:'#fff', borderRadius:3, padding:'1px 4px' }}>
                          {inv.pageStatus[pi]==='processing'?'⏳':inv.pageStatus[pi]==='done'?'✓':inv.pageStatus[pi]==='error'?'✗':'⋯'}
                        </div>
                        <button onClick={()=>removePage(gi,pi)} style={{ position:'absolute', top:-6, left:-6, width:18, height:18, borderRadius:'50%', background:'#ef4444', color:'#fff', border:'none', cursor:'pointer', fontSize:10, display:'flex', alignItems:'center', justifyContent:'center', lineHeight:1, zIndex:10 }}>✕</button>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
              <button onClick={processAll} disabled={processing} style={{ width:'100%', padding:'12px', borderRadius:8, background:processing?'#9ca3af':'#4f46e5', color:'#fff', fontWeight:700, fontSize:14, border:'none', cursor:processing?'not-allowed':'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>
                {processing?<><Spinner size={18}/>กำลังประมวลผล...</>:<>🤖 ประมวลผลทุกใบ ({invoices.length} กลุ่ม)</>}
              </button>
            </div>
          )}
        </div>
      )}

      {step===2&&(
        <div style={{ display:'flex', flexDirection:'column', gap:16 }}>
          <div style={{ fontWeight:600, fontSize:15, color:'#374151' }}>สแกนรูปสินค้าเพื่อดึงบาร์โค้ด</div>
          <DropZone multiple accept="image/*" onFiles={addProductFiles}>
            <div style={{ fontSize:28, marginBottom:6 }}>📷</div>
            <div style={{ fontWeight:600, color:'#374151', fontSize:13 }}>ลากรูปสินค้าที่มีบาร์โค้ด</div>
            <div style={{ fontSize:11, color:'#9ca3af' }}>Claude จะอ่านบาร์โค้ดจากรูปและจับคู่กับรายการสินค้า</div>
          </DropZone>
          {productFiles.length>0&&(
            <div style={{ background:'#fff', border:'1px solid #e2e8f0', borderRadius:10, padding:12 }}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:10, gap:8, flexWrap:'wrap' }}>
                <div style={{ fontSize:13, fontWeight:600, color:'#374151' }}>รูปสินค้า ({productFiles.length} ไฟล์)</div>
                <div style={{ display:'flex', gap:6 }}>
                  {productFiles.some(it=>it.status==='pending')&&(
                    <button onClick={scanPendingPFiles} disabled={scanning} style={{ fontSize:11, padding:'4px 12px', borderRadius:6, background:'#059669', color:'#fff', border:'none', cursor:scanning?'wait':'pointer', fontWeight:600 }}>
                      {scanning?'กำลังสแกน...':`▶ เริ่ม scan (${productFiles.filter(it=>it.status==='pending').length})`}
                    </button>
                  )}
                  {selPFileIds.size>0&&<button onClick={rescanSelectedPFiles} disabled={scanning} style={{ fontSize:11, padding:'4px 10px', borderRadius:6, background:'#4f46e5', color:'#fff', border:'none', cursor:'pointer' }}>{scanning?'กำลังสแกน...':'Re-scan ที่เลือก'}</button>}
                </div>
              </div>
              <div style={{ display:'flex', flexWrap:'wrap', gap:8 }}>
                {productFiles.map(it=>(
                  <div key={it.id} style={{ position:'relative' }}>
                    <div style={{ cursor:'pointer' }} onClick={()=>setSelPFIds(prev=>{const s=new Set(prev);s.has(it.id)?s.delete(it.id):s.add(it.id);return s;})}>
                      <InvFileThumb file={it.file}/>
                      <div style={{ position:'absolute', top:2, right:2, fontSize:10, background:it.status==='done'?'#10b981':it.status==='error'?'#ef4444':it.status==='processing'?'#f59e0b':'#6b7280', color:'#fff', borderRadius:3, padding:'1px 4px' }}>{it.status==='done'?'✓':it.status==='error'?'✗':it.status==='processing'?'⏳':'⋯'}</div>
                      {selPFileIds.has(it.id)&&<div style={{ position:'absolute', inset:0, border:'2px solid #4f46e5', borderRadius:6, background:'rgba(79,70,229,0.1)' }}/>}
                    </div>
                    <button onClick={e=>{e.stopPropagation();deleteProductFile(it.id);}} style={{ position:'absolute', top:-6, left:-6, width:18, height:18, borderRadius:'50%', background:'#ef4444', color:'#fff', border:'none', cursor:'pointer', fontSize:10, display:'flex', alignItems:'center', justifyContent:'center', lineHeight:1, zIndex:10 }}>✕</button>
                  </div>
                ))}
              </div>
            </div>
          )}
          {scanResults.length>0&&(
            <div style={{ background:'#fff', border:'1px solid #e2e8f0', borderRadius:10, padding:12 }}>
              <div style={{ fontSize:13, fontWeight:600, color:'#374151', marginBottom:8 }}>ผลการจับคู่บาร์โค้ด ({scanResults.length} รายการ)</div>
              <div style={{ overflowX:'auto' }}>
                <table style={{ width:'100%', fontSize:11, borderCollapse:'collapse' }}>
                  <thead><tr style={{ background:'#f8fafc' }}>{['#','บาร์โค้ด','รูป (description_image)','จับคู่กับ'].map(h=><th key={h} style={{ padding:'6px 8px', textAlign:'left', fontWeight:600, color:'#64748b', whiteSpace:'nowrap' }}>{h}</th>)}</tr></thead>
                  <tbody>
                    {scanResults.map((r,i)=>{
                      const usedByOthers = new Set(scanResults.filter((_,j)=>j!==i).map(x=>x.match).filter(Boolean));
                      const allDescs = [...new Set(doneInvs.flatMap(inv=>(inv.data.products||[]).map(p=>p.description)).filter(Boolean))];
                      return (
                      <tr key={i} style={{ borderTop:'1px solid #f1f5f9', background: r._userOverride ? '#fefce8' : 'transparent' }}>
                        <td style={{ padding:'5px 8px', color:'#9ca3af' }}>{i+1}</td>
                        <td style={{ padding:'4px 8px' }}><input value={r.barcode??''} onChange={e=>{const u=[...scanResults];u[i]={...u[i],barcode:e.target.value||null};setSRes(u);applyBarcodeMap(buildMap(u));}} style={{ width:130, fontFamily:'monospace', fontSize:11, color:'#065f46', border:'1px solid #e2e8f0', borderRadius:4, padding:'2px 4px' }}/></td>
                        <td style={{ padding:'4px 8px' }}><input value={r.description_image??''} onChange={e=>{const u=[...scanResults];u[i]={...u[i],description_image:e.target.value||null};setSRes(u);}} style={{ width:180, fontSize:11, border:'1px solid #e2e8f0', borderRadius:4, padding:'2px 4px' }}/></td>
                        <td style={{ padding:'4px 8px' }}>
                          <select value={r.match??''} onChange={e=>{const u=[...scanResults];u[i]={...u[i],match:e.target.value||null,_userOverride:true};setSRes(u);applyBarcodeMap(buildMap(u));}} style={{ width:200, fontSize:11, border: r._userOverride?'2px solid #eab308':'1px solid #e2e8f0', borderRadius:4, padding:'2px 4px', background: r._userOverride?'#fefce8':'#fff' }}>
                            <option value="">— เลือก —</option>
                            {allDescs.map((desc,idx) => {
                              const isUsed = usedByOthers.has(desc);
                              return <option key={idx} value={desc} disabled={isUsed}>{(desc||'').slice(0,40)}{isUsed?' (ใช้แล้ว)':''}</option>;
                            })}
                          </select>
                        </td>
                      </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
          <div style={{ display:'flex', gap:8 }}>
            <button onClick={()=>setStep(1)} style={{ padding:'12px 20px', borderRadius:8, border:'1px solid #e2e8f0', background:'#fff', color:'#374151', fontWeight:600, fontSize:14, cursor:'pointer' }}>← กลับ</button>
            <button onClick={()=>setStep(3)} style={{ flex:1, padding:'12px', borderRadius:8, background:'#4f46e5', color:'#fff', fontWeight:700, fontSize:14, border:'none', cursor:'pointer' }}>ตรวจสอบข้อมูล →</button>
          </div>
        </div>
      )}

      {step===3&&(
        <div style={{ display:'flex', flexDirection:'column', gap:16 }}>
          <div style={{ fontWeight:600, fontSize:15, color:'#374151' }}>ตรวจสอบข้อมูลใบกำกับ</div>
          {doneInvs.map((inv,gi)=>{
            const d=inv.data, vs=vatSummary(d.products||[]);
            const upd = (patch) => updateData(gi, {...d, ...patch});
            const updP = (pi, patch) => { const prods=[...d.products]; prods[pi]=recalc({...prods[pi],...patch,_pt:d.price_type??'incl'}); updateData(gi,{...d,products:prods}); };
            return (
              <div key={gi} style={{ background:'#fff', border:'1px solid #e2e8f0', borderRadius:10, overflow:'hidden' }}>
                <div style={{ background:'#f8fafc', padding:'10px 14px', borderBottom:'1px solid #e2e8f0', display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                  <div style={{ fontWeight:700, fontSize:13, color:'#374151' }}>ใบที่ {gi+1}: {d.vendor_name||'(ไม่ระบุ)'}</div>
                  <button onClick={()=>reprocessInvoice(gi)} style={{ fontSize:11, padding:'3px 8px', borderRadius:5, background:'#e0e7ff', color:'#4f46e5', border:'none', cursor:'pointer' }}>อ่านใหม่</button>
                </div>
                <div style={{ padding:14 }}>
                  {/* Header fields */}
                  <div style={{ display:'grid', gridTemplateColumns:'1fr', gap:6, marginBottom:10, fontSize:12 }}>
                    <div><div style={{ color:'#9ca3af', fontSize:10, marginBottom:2 }}>ชื่อร้าน / บริษัท</div><input value={d.vendor_name??''} onChange={e=>upd({vendor_name:e.target.value||null})} style={{ width:'100%', padding:'5px 8px', border:'1px solid #e2e8f0', borderRadius:6, fontSize:12, boxSizing:'border-box' }}/></div>
                  </div>
                  <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:6, marginBottom:10, fontSize:12 }}>
                    {[['เลขที่ใบกำกับ','invoice_no'],['วันที่','invoice_date'],['เลขภาษี','vendor_tax_id'],['ประเภทเอกสาร','document_type'],['สาขา','vendor_branch']].map(([label,key])=>(
                      <div key={key}><div style={{ color:'#9ca3af', fontSize:10, marginBottom:2 }}>{label}</div><input value={d[key]??''} onChange={e=>upd({[key]:e.target.value||null})} style={{ width:'100%', padding:'5px 8px', border:'1px solid #e2e8f0', borderRadius:6, fontSize:12, boxSizing:'border-box' }}/></div>
                    ))}
                    <div>
                      <div style={{ color:'#9ca3af', fontSize:10, marginBottom:2, display:'flex', alignItems:'center', gap:4 }}>
                        รหัสผู้ขาย (vendor_no)
                        {d._vendorFromDB && d.vendor_no && <span style={{ background:'#d1fae5', color:'#065f46', fontSize:9, fontWeight:700, padding:'1px 5px', borderRadius:4 }}>● DB</span>}
                      </div>
                      <input value={d.vendor_no??''} onChange={e=>upd({vendor_no:e.target.value||null, _vendorFromDB:false})} style={{ width:'100%', padding:'5px 8px', border:'1px solid #e2e8f0', borderRadius:6, fontSize:12, boxSizing:'border-box' }}/>
                    </div>
                    <div><div style={{ color:'#9ca3af', fontSize:10, marginBottom:2 }}>ราคา</div>
                      <select value={d.price_type??'incl'} onChange={e=>{ const pt=e.target.value; const prods=(d.products||[]).map(p=>recalc({...p,_pt:pt})); updateData(gi,{...d,price_type:pt,products:prods}); }} style={{ width:'100%', padding:'5px 8px', border:'1px solid #e2e8f0', borderRadius:6, fontSize:12 }}>
                        <option value="incl">รวม VAT แล้ว (incl)</option>
                        <option value="excl">ยังไม่รวม VAT (excl)</option>
                      </select>
                    </div>
                  </div>
                  <div style={{ marginBottom:10, fontSize:12 }}>
                    <div style={{ color:'#9ca3af', fontSize:10, marginBottom:2 }}>ที่อยู่</div>
                    <textarea value={d.vendor_address??''} onChange={e=>upd({vendor_address:e.target.value||null})} rows={2} style={{ width:'100%', padding:'5px 8px', border:'1px solid #e2e8f0', borderRadius:6, fontSize:12, resize:'vertical', boxSizing:'border-box' }}/>
                  </div>
                  {/* Product table */}
                  <div style={{ overflowX:'auto' }}>
                    <table style={{ width:'100%', fontSize:11, borderCollapse:'collapse', minWidth:1100 }}>
                      <thead>
                        <tr style={{ background:'#f8fafc' }}>
                          {['#','สินค้า','ขนาดลัง','ลัง','หน่วยลัง','ชิ้น','หน่วย','รวม','ราคา/หน่วย','ยอดตามใบ','ส่วนลด','ยอดสุทธิ','diff','VAT','บาร์โค้ด'].map(h=>(
                            <th key={h} style={{ padding:'6px 8px', textAlign:'left', fontWeight:600, color:'#64748b', whiteSpace:'nowrap' }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {(d.products||[]).map((p,pi)=>{
                          const bc=p.barcode??barcodeMap[String(p.description||'').trim()]??null;
                          const inp = (field, w=50, type='number') => (
                            <input type={type} value={p[field]??''} onChange={e=>updP(pi,{[field]:e.target.value===''?null:e.target.value})}
                              style={{ width:w, fontSize:11, border:'1px solid #e2e8f0', borderRadius:4, padding:'2px 4px', textAlign:'center' }}/>
                          );
                          const inpText = (field, w=55) => (
                            <input type="text" value={p[field]??''} onChange={e=>updP(pi,{[field]:e.target.value||null})}
                              style={{ width:w, fontSize:11, border:'1px solid #e2e8f0', borderRadius:4, padding:'2px 4px', textAlign:'center' }}/>
                          );
                          const diffColor = p.diff == null ? '#9ca3af'
                            : Math.abs(p.diff) < 0.01 ? '#059669'
                            : p.diff > 0 ? '#d97706' : '#dc2626';
                          return (
                            <tr key={pi} style={{ borderTop:'1px solid #f1f5f9' }}>
                              <td style={{ padding:'4px 8px', color:'#9ca3af' }}>{p.no??pi+1}</td>
                              <td style={{ padding:'4px 8px', minWidth:160 }}>
                                <input value={p.description??''} onChange={e=>{const prods=[...d.products];prods[pi]={...prods[pi],description:e.target.value};updateData(gi,{...d,products:prods});}}
                                  style={{ width:'100%', fontSize:11, border:'1px solid #e2e8f0', borderRadius:4, padding:'2px 4px' }}/>
                              </td>
                              <td style={{ padding:'4px 8px' }}>{inp('carton_size',50)}</td>
                              <td style={{ padding:'4px 8px' }}>{inp('carton',45)}</td>
                              <td style={{ padding:'4px 8px' }}>{inpText('carton_unit',55)}</td>
                              <td style={{ padding:'4px 8px' }}>{inp('ea',45)}</td>
                              <td style={{ padding:'4px 8px' }}>{inpText('ea_unit',55)}</td>
                              <td style={{ padding:'4px 8px' }}><span style={{ fontWeight:600, color:'#374151' }}>{p.qty??'-'}</span></td>
                              <td style={{ padding:'4px 8px' }}>{inp('price_ea',70)}</td>
                              <td style={{ padding:'4px 8px' }}>{inp('amount',70)}</td>
                              <td style={{ padding:'4px 8px' }}>{inp('special_discount',60)}</td>
                              <td style={{ padding:'4px 8px' }}><span style={{ fontWeight:600, color:'#059669' }}>{p.total!=null?Number(p.total).toLocaleString():'-'}</span></td>
                              <td style={{ padding:'4px 8px' }}><span style={{ fontWeight:600, color:diffColor }}>{p.diff!=null?Number(p.diff).toLocaleString():'-'}</span></td>
                              <td style={{ padding:'4px 8px' }}>
                                <select value={p.vat??'v'} onChange={e=>updP(pi,{vat:e.target.value})}
                                  style={{ fontSize:11, border:'1px solid #e2e8f0', borderRadius:4, padding:'2px 4px' }}>
                                  <option value="v">7%</option><option value="n">0%</option>
                                </select>
                              </td>
                              <td style={{ padding:'4px 8px' }}>
                                <input value={bc??''} onChange={e=>{const prods=[...d.products];prods[pi]={...prods[pi],barcode:e.target.value||null};updateData(gi,{...d,products:prods});}}
                                  style={{ width:130, fontFamily:'monospace', fontSize:11, border:'1px solid #e2e8f0', borderRadius:4, padding:'2px 4px', color:'#065f46' }}/>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                  <div style={{ display:'flex', justifyContent:'flex-end', gap:16, marginTop:8, fontSize:12, color:'#64748b', flexWrap:'wrap' }}>
                    <span>ส่วนลดรวม: <strong style={{ color:'#dc2626' }}>-฿{vs.sdTot.toLocaleString()}</strong></span>
                    <span>ไม่รวม VAT: <strong style={{ color:'#374151' }}>฿{vs.excl.toLocaleString()}</strong></span>
                    <span>VAT 7%: <strong style={{ color:'#374151' }}>฿{vs.vatAmt.toLocaleString()}</strong></span>
                    <span>ยอดสุทธิ: <strong style={{ color:'#059669' }}>฿{vs.netTotal.toLocaleString()}</strong></span>
                  </div>
                </div>
              </div>
            );
          })}
          <div style={{ display:'flex', gap:8 }}>
            <button onClick={()=>setStep(2)} style={{ padding:'10px 20px', borderRadius:8, border:'1px solid #e2e8f0', background:'#fff', color:'#374151', fontWeight:600, fontSize:13, cursor:'pointer' }}>← กลับ</button>
            <button onClick={()=>setStep(4)} disabled={doneInvs.length===0} style={{ flex:1, padding:'10px', borderRadius:8, background:doneInvs.length===0?'#9ca3af':'#4f46e5', color:'#fff', fontWeight:700, fontSize:13, border:'none', cursor:doneInvs.length===0?'not-allowed':'pointer' }}>สรุปและบันทึก →</button>
          </div>
        </div>
      )}

      {step===4&&(
        <div style={{ display:'flex', flexDirection:'column', gap:16 }}>
          <div style={{ background:'#fff', border:'1px solid #e2e8f0', borderRadius:10, padding:16 }}>
            <div style={{ fontWeight:700, fontSize:15, color:'#374151', marginBottom:12 }}>สรุปภาพรวม</div>
            <div style={{ display:'grid', gridTemplateColumns:'repeat(3,1fr)', gap:12, textAlign:'center', marginBottom:16 }}>
              {[['ใบกำกับ',`${doneInvs.length} ใบ`],['สินค้า',`${allProds.length} รายการ`],['ยอดรวม',`฿${grandTotal.toLocaleString()}`]].map(([k,v])=>(
                <div key={k} style={{ background:'#f8fafc', borderRadius:8, padding:'10px' }}><div style={{ fontSize:11, color:'#9ca3af', marginBottom:4 }}>{k}</div><div style={{ fontWeight:700, fontSize:16, color:'#374151' }}>{v}</div></div>
              ))}
            </div>
            {(() => {
              const issues = [];
              for (const inv of doneInvs) {
                for (const p of (inv.data.products||[])) {
                  const miss = [];
                  const bc = p.barcode ?? barcodeMap[String(p.description||'').trim()] ?? '';
                  if (!bc) miss.push('barcode');
                  const qty = p.qty != null ? +p.qty : 0;
                  if (qty <= 0) miss.push('qty');
                  if (p.price_ea == null && qty > 0) miss.push('total/qty');
                  if (miss.length > 0) issues.push({ invoice_no: inv.data.invoice_no, no: p.no, description: p.description, missing: miss });
                }
              }
              if (issues.length === 0) return null;
              return (
                <div style={{ background:'#fef2f2', border:'1.5px solid #fecaca', borderRadius:8, padding:10, marginBottom:12 }}>
                  <div style={{ fontWeight:700, fontSize:12, color:'#991b1b', marginBottom:6 }}>⚠ พบ {issues.length} รายการ ที่ข้อมูลไม่ครบ (จะส่งออกไปเป็นค่าว่าง / 0)</div>
                  <div style={{ maxHeight:120, overflowY:'auto', fontSize:11, color:'#7f1d1d' }}>
                    {issues.slice(0, 30).map((iss, idx) => (
                      <div key={idx} style={{ marginBottom:2 }}>
                        • {iss.invoice_no} #{iss.no} <span style={{ color:'#9ca3af' }}>{(iss.description||'').slice(0,30)}</span> — ขาด: <strong>{iss.missing.join(', ')}</strong>
                      </div>
                    ))}
                    {issues.length > 30 && <div style={{ color:'#9ca3af' }}>... และอีก {issues.length - 30} รายการ</div>}
                  </div>
                </div>
              );
            })()}
            <div style={{ marginBottom:12 }}>
              <label style={{ fontSize:12, fontWeight:600, color:'#374151', display:'block', marginBottom:4 }}>ชื่อไฟล์</label>
              <div style={{ display:'flex', gap:8, alignItems:'center' }}>
                <input value={fileName} onChange={e=>setFileName(e.target.value)} placeholder="260517xxxx" style={{ flex:1, padding:'8px 12px', border:`1.5px solid ${nameStatus==='duplicate'?'#ef4444':nameStatus==='available'?'#10b981':'#e2e8f0'}`, borderRadius:8, fontSize:13, fontFamily:'monospace' }}/>
                {nameStatus==='checking'&&<Spinner size={16}/>}
                {nameStatus==='duplicate'&&<span style={{ color:'#ef4444', fontSize:11 }}>ซ้ำ!</span>}
                {nameStatus==='available'&&<span style={{ color:'#10b981', fontSize:11 }}>✓</span>}
              </div>
            </div>
            <div style={{ borderTop:'1px solid #f1f5f9', paddingTop:12, marginTop:12 }}>
              <div style={{ fontWeight:600, fontSize:13, color:'#374151', marginBottom:8 }}>อัปโหลดไปยัง Google Drive</div>
              <div style={{ display:'flex', gap:8, flexWrap:'wrap' }}>
                {!gToken ? <button onClick={connectGoogle} style={{ padding:'8px 16px', borderRadius:8, background:'#4285f4', color:'#fff', border:'none', fontSize:13, fontWeight:600, cursor:'pointer' }}>🔗 Connect Google Drive</button>
                : <><button onClick={uploadToDrive} disabled={driveStatus==='uploading'} style={{ padding:'8px 16px', borderRadius:8, background:driveStatus==='uploading'?'#9ca3af':'#4285f4', color:'#fff', border:'none', fontSize:13, fontWeight:600, cursor:driveStatus==='uploading'?'not-allowed':'pointer', display:'flex', alignItems:'center', gap:8 }}>{driveStatus==='uploading'?<><Spinner size={16}/>กำลังอัปโหลด...</>:'📤 อัปโหลด CSV → Drive'}</button><button onClick={disconnectGoogle} style={{ padding:'8px 12px', borderRadius:8, background:'#fee2e2', color:'#dc2626', border:'none', fontSize:12, cursor:'pointer' }}>ยกเลิก Google</button></>}
              </div>
              {driveStatus==='done'&&driveUrl&&<div style={{ marginTop:8, fontSize:12, color:'#059669' }}>✓ อัปโหลดแล้ว <a href={driveUrl} target="_blank" rel="noopener noreferrer" style={{ color:'#4f46e5', textDecoration:'underline' }}>เปิดใน Drive →</a></div>}
              {driveStatus==='error'&&driveErr&&<div style={{ marginTop:8, fontSize:12, color:'#dc2626' }}>✗ {driveErr}</div>}
            </div>
            <div style={{ borderTop:'1px solid #f1f5f9', paddingTop:12, marginTop:12 }}>
              <div style={{ fontWeight:600, fontSize:13, color:'#374151', marginBottom:8 }}>บันทึกลง Supabase</div>
              {!sbUrl||!sbKey ? <div style={{ fontSize:12, color:'#f59e0b' }}>⚠ ยังไม่ได้ตั้งค่า Supabase — ไปที่ ตั้งค่า ก่อน</div> : (
                <div style={{ display:'flex', gap:8, flexWrap:'wrap', alignItems:'center' }}>
                  <button onClick={saveToSupabase} disabled={sbSt==='saving'||nameStatus==='duplicate'} style={{ padding:'8px 16px', borderRadius:8, background:sbSt==='saving'||nameStatus==='duplicate'?'#9ca3af':'#0f172a', color:'#fff', border:'none', fontSize:13, fontWeight:600, cursor:sbSt==='saving'||nameStatus==='duplicate'?'not-allowed':'pointer', display:'flex', alignItems:'center', gap:8 }}>{sbSt==='saving'?<><Spinner size={16}/>กำลังบันทึก...</>:'💾 บันทึก bill_header + imp_data'}</button>
                  {sbSt==='done'&&<span style={{ fontSize:12, color:'#059669' }}>✓ บันทึกแล้ว!</span>}
                  {sbSt==='error'&&sbErr&&<span style={{ fontSize:12, color:'#dc2626' }}>✗ {sbErr}</span>}
                </div>
              )}
            </div>
            <div style={{ borderTop:'1px solid #f1f5f9', paddingTop:12, marginTop:12 }}>
              <button onClick={()=>{const b=buildXLSXBlob();if(b)downloadBlob(b,(fileName||'invoice')+'.xlsx');}} style={{ padding:'8px 16px', borderRadius:8, background:'#059669', color:'#fff', border:'none', fontSize:13, fontWeight:600, cursor:'pointer' }}>⬇ ดาวน์โหลด Excel</button>
            </div>
          </div>
          <button onClick={()=>setStep(3)} style={{ padding:'10px', borderRadius:8, border:'1px solid #e2e8f0', background:'#fff', color:'#374151', fontWeight:600, fontSize:13, cursor:'pointer' }}>← แก้ไขข้อมูล</button>
        </div>
      )}
    </div>
  );
}

function ScannerModal({ products, onScan, onClose }) {
  const [tab, setTab] = useState('camera');
  const [manualBarcode, setManualBarcode] = useState('');
  const [scanError, setScanError] = useState('');
  const [scanning, setScanning] = useState(false);
  const fileInputRef = useRef(null);
  const html5QrRef = useRef(null);
  const cameraStarted = useRef(false);

  useEffect(() => {
    if (tab !== 'camera') {
      if (html5QrRef.current) {
        html5QrRef.current.stop().catch(() => {});
        html5QrRef.current = null;
        cameraStarted.current = false;
      }
      return;
    }
    let cancelled = false;
    (async () => {
      try {
        const { Html5Qrcode, Html5QrcodeSupportedFormats } = await import('html5-qrcode');
        if (cancelled) return;
        const formatsToSupport = [
          Html5QrcodeSupportedFormats.EAN_13,
          Html5QrcodeSupportedFormats.EAN_8,
          Html5QrcodeSupportedFormats.CODE_128,
          Html5QrcodeSupportedFormats.CODE_39,
          Html5QrcodeSupportedFormats.UPC_A,
          Html5QrcodeSupportedFormats.UPC_E,
          Html5QrcodeSupportedFormats.QR_CODE,
          Html5QrcodeSupportedFormats.ITF,
        ];
        html5QrRef.current = new Html5Qrcode('__qr_reader__', { formatsToSupport, verbose: false });
        await html5QrRef.current.start(
          { facingMode: 'environment' },
          { fps: 10, qrbox: { width: 240, height: 240 } },
          (text) => { if (!cancelled) onScan(text); }
        );
        cameraStarted.current = true;
      } catch (e) {
        if (!cancelled) setScanError('เปิดกล้องไม่ได้: ' + (e.message || e));
      }
    })();
    return () => {
      cancelled = true;
      if (html5QrRef.current) {
        html5QrRef.current.stop().catch(() => {});
        html5QrRef.current = null;
        cameraStarted.current = false;
      }
    };
  }, [tab]);

  const preprocessImage = (file, deg = 0) => new Promise((resolve) => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      const sw = deg === 90 || deg === 270;
      const c = document.createElement('canvas');
      c.width  = sw ? img.height : img.width;
      c.height = sw ? img.width  : img.height;
      const ctx = c.getContext('2d');
      if (deg !== 0) {
        ctx.translate(c.width / 2, c.height / 2);
        ctx.rotate((deg * Math.PI) / 180);
        ctx.drawImage(img, -img.width / 2, -img.height / 2);
        ctx.resetTransform();
      } else {
        ctx.drawImage(img, 0, 0);
      }
      const d = ctx.getImageData(0, 0, c.width, c.height);
      for (let i = 0; i < d.data.length; i += 4) {
        const gray = Math.round(0.299 * d.data[i] + 0.587 * d.data[i+1] + 0.114 * d.data[i+2]);
        const boosted = Math.min(255, Math.max(0, ((gray - 128) * 1.8) + 128));
        d.data[i] = d.data[i+1] = d.data[i+2] = boosted;
      }
      ctx.putImageData(d, 0, 0);
      URL.revokeObjectURL(url);
      c.toBlob(blob => resolve(new File([blob], 'scan.jpg', { type: 'image/jpeg' })), 'image/jpeg', 0.95);
    };
    img.src = url;
  });

  const decodeWithQuagga = (dataUrl) => new Promise((resolve, reject) => {
    import('@ericblade/quagga2').then(({ default: Quagga }) => {
      Quagga.decodeSingle({
        decoder: {
          readers: ['ean_reader','ean_8_reader','code_128_reader','code_39_reader','upc_reader','upc_e_reader'],
          multiple: false,
        },
        locate: true,
        src: dataUrl,
      }, (result) => {
        if (result?.codeResult?.code) resolve(result.codeResult.code);
        else reject(new Error('not found'));
      });
    }).catch(reject);
  });

  const handleFileSelect = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setScanError(''); setScanning(true);
    try {
      for (const deg of [0, 180, 90, 270]) {
        try {
          const processed = await preprocessImage(file, deg);
          const dataUrl = await new Promise(res => {
            const fr = new FileReader();
            fr.onload = e => res(e.target.result);
            fr.readAsDataURL(processed);
          });
          const code = await decodeWithQuagga(dataUrl);
          onScan(code); setScanning(false); return;
        } catch (_) {}
      }
      setScanError('ไม่พบบาร์โค้ด — ลองถ่ายให้ชัด บาร์โค้ดอยู่เต็มกรอบ หรือพิมพ์รหัสในแท็บ "พิมพ์"');
    } catch (err) {
      setScanError('เกิดข้อผิดพลาด: ' + err.message);
    }
    setScanning(false);
  };

  return (
    <div className="fixed inset-0 bg-black/80 z-50 flex items-end sm:items-center justify-center p-0 sm:p-4">
      <div className="bg-white rounded-t-2xl sm:rounded-2xl w-full max-w-md max-h-[95vh] overflow-y-auto">
        <div className="p-4 border-b border-[#E4E6EA] flex justify-between items-center sticky top-0 bg-white z-10">
          <h3 className="font-semibold">สแกนบาร์โค้ด</h3>
          <button onClick={onClose}><X size={20} /></button>
        </div>
        <div className="flex border-b border-[#E4E6EA] sticky top-[57px] bg-white z-10">
          {[
            { id: 'camera', label: 'กล้อง',    icon: Camera    },
            { id: 'upload', label: 'อัพโหลด',  icon: ImageIcon },
            { id: 'manual', label: 'พิมพ์',    icon: Edit3     },
          ].map(t => {
            const Icon = t.icon;
            return (
              <button
                key={t.id}
                onClick={() => { setScanError(''); setTab(t.id); }}
                className={`flex-1 py-2.5 text-sm font-medium flex items-center justify-center gap-1.5 ${tab === t.id ? 'text-[#35706A] border-b-2 border-[#35706A]' : 'text-slate-500'}`}
              >
                <Icon size={16} />{t.label}
              </button>
            );
          })}
        </div>

        {tab === 'camera' && (
          <div className="p-2">
            <div id="__qr_reader__" className="w-full rounded-lg overflow-hidden" />
            {scanError && (
              <div className="mt-2 bg-[#FEF2F2] border border-[#FECACA] rounded-lg p-3 text-sm text-[#B91C1C]">{scanError}</div>
            )}
            <p className="text-center text-xs text-slate-400 mt-2">จ่อบาร์โค้ดให้อยู่ในกรอบ</p>
          </div>
        )}

        {tab === 'upload' && (
          <div className="p-4 space-y-3">
            <div id="__qr_file_reader__" style={{ display: 'none' }} />
            <button
              onClick={() => fileInputRef.current?.click()}
              className="w-full border-2 border-dashed border-[#E4E6EA] hover:border-[#35706A] hover:bg-[#EAF1F0] rounded-xl p-8 flex flex-col items-center gap-2"
            >
              <div className="bg-[#EAF1F0] p-3 rounded-full">
                <Upload className="text-[#35706A]" size={24} />
              </div>
              <div className="text-sm font-medium text-slate-700">{scanning ? 'กำลังอ่าน...' : 'เลือกรูปบาร์โค้ด'}</div>
              <div className="text-xs text-slate-400">รองรับทุกเบราว์เซอร์</div>
            </button>
            {scanError && (
              <div className="bg-[#FEF2F2] border border-[#FECACA] rounded-lg p-3 text-sm text-[#B91C1C]">{scanError}</div>
            )}
            <input ref={fileInputRef} type="file" accept="image/*" onChange={handleFileSelect} className="hidden" />
          </div>
        )}

        {tab === 'manual' && (
          <div className="p-4 space-y-3">
            <input
              type="text"
              value={manualBarcode}
              onChange={(e) => setManualBarcode(e.target.value)}
              placeholder="พิมพ์รหัสสินค้า..."
              className="w-full px-3 py-3 border border-[#E4E6EA] rounded-lg outline-none focus:ring-2 focus:ring-[#35706A] text-lg font-mono"
              autoFocus
            />
            <button
              onClick={() => manualBarcode && onScan(manualBarcode)}
              disabled={!manualBarcode}
              className="w-full bg-[#35706A] hover:bg-[#2A5A55] disabled:bg-[#E4E6EA] text-white py-2.5 rounded-lg font-medium"
            >
              ใช้รหัสนี้
            </button>
            {products.length > 0 && (
              <div>
                <div className="text-xs text-slate-500 mb-2">ตัวอย่าง:</div>
                <div className="grid grid-cols-2 gap-2">
                  {products.slice(0, 4).map(p => (
                    <button
                      key={p.id}
                      onClick={() => onScan(p.id)}
                      className="text-left p-2 bg-[#F6F7F8] hover:bg-[#F6F7F8] rounded-lg text-xs"
                    >
                      <div className="font-medium text-slate-700 truncate">{p.name}</div>
                      <div className="text-slate-500 font-mono">{p.id}</div>
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
