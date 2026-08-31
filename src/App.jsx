import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import {
  Package, Search, ScanLine, Plus, Minus, Download, Home, List, X, Check,
  Camera, AlertCircle, Trash2, Edit3, Save, Image as ImageIcon, Upload, User,
  Shield, Eye, EyeOff, ClipboardCheck, Lock, LogOut, Database, Cloud, ChevronRight,
  RefreshCw, Settings as SettingsIcon, CheckCircle2, XCircle, Layers,
  FileSpreadsheet, ArrowRight, FileCheck, WifiOff, Zap, Send, Clock,
  ThumbsUp, ThumbsDown, Inbox, ArrowLeftRight, Receipt, MapPin, Users, Menu, Sparkles,
} from 'lucide-react';

const INVOICE_API = "/api/claude";
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

// จับคู่ด้วยชื่อที่อ่านจากรูปเท่านั้น — ไม่ต้องส่งรูปซ้ำ
// (ผลค้น product_price ใช้บอกแค่ว่าบาร์โค้ดนี้มีในระบบไหม ไม่เอาไปใช้จับคู่)
const REMATCH_PROMPT = (list, items) => `จับคู่ชื่อสินค้าที่อ่านจากรูป กับรายการในใบกำกับ
ชื่อบนกล่องมักย่อหรือสลับคำ ให้ดูขนาด/ปริมาณ/รส ประกอบ
ถ้าไม่มีอันไหนตรงจริง ให้ match เป็น null — ห้ามเดา

ตอบเป็น JSON array เท่านั้น รักษาค่า i เดิม:
[{"i":0,"match":"ชื่อในใบกำกับที่ตรง"}]

รายการในใบกำกับ:
${list}

ชื่อที่อ่านจากรูป:
${items}`;

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

// บาร์โค้ดที่ยิงมาอาจเป็น: รหัสสินค้า · บาร์โค้ดหลัก · บาร์โค้ด-1..5 · บาร์โค้ดผู้ขาย · บาร์โค้ดแพ็ค
// product_barcode (view) รวมทุกช่องไว้แล้ว — ค้นที่นี่ก่อน ได้รหัสสินค้าจริงกลับมา
// product_pack = บาร์โค้ดแพ็ค → รหัสจริง + ยิง 1 ครั้งตัดกี่หน่วย
// 🔴 แพ็ค/ลังเป็นแถวแยกใน product_master แต่ "ไม่มีสต็อกของตัวเอง"
// สต็อกอยู่ที่ "รหัสหลัก" (xProduct) รหัสเดียว แล้วใช้ "ตัวคูณสต็อก" แปลงเป็นชิ้น
// ~11% ของสินค้าเป็นแถวแพ็ค — ถ้าเก็บด้วยรหัสแพ็ค สต็อกจะแตกเป็น 2 ก้อน ไม่ตรง POS
async function resolveBarcode(base, h, scanned, branch) {
  const brFilter = stockBranchFilter(branch);
  const sel = encodeURIComponent('"รหัสสินค้า","ชื่อสินค้า","ราคา-1","ทุนเฉลี่ย","หน่วยนับ","ตัวคูณสต็อก","รหัสหลัก","เป็นแพ็ค","คงเหลือ"');
  const pick = (d) => ({
    code: String(d['รหัสสินค้า'] || ''),
    // รหัสที่ใช้เก็บ/นับสต็อกเสมอ
    masterCode: String(d['รหัสหลัก'] || d['รหัสสินค้า'] || ''),
    ratio: Number(d['ตัวคูณสต็อก']) || 1,
    isPack: !!d['เป็นแพ็ค'],
    onHandPieces: d['คงเหลือ'] == null ? null : Number(d['คงเหลือ']),
    name: d['ชื่อสินค้า'] || '',
    unit: d['หน่วยนับ'] || '',
    price: Number(d['ราคา-1']) || 0,
    cost: Number(d['ทุนเฉลี่ย']) || 0,
  });
  try {
    const bcol = encodeURIComponent('"บาร์โค้ด"');
    const r = await fetch(`${base}/rest/v1/product_barcode?${bcol}=eq.${encodeURIComponent(scanned)}${brFilter}&select=${sel}&limit=2`, { headers: h });
    if (r.ok) {
      const d = await r.json();
      // ชนกันหลายสินค้า = เดาไม่ได้ เตือนคน ไม่หยิบตัวแรกมั่ว
      if (d.length) return { ...pick(d[0]), conflict: d.length > 1 ? d.length : null };
    }
  } catch {}
  // ยิงด้วยรหัสสินค้าตรง ๆ (ไม่ใช่บาร์โค้ด) — view ค้นด้วยรหัสไม่ได้ ต้องมาทางนี้
  try {
    const ccol = encodeURIComponent('"รหัสสินค้า"');
    const r = await fetch(`${base}/rest/v1/product_barcode?${ccol}=eq.${encodeURIComponent(scanned)}${brFilter}&select=${sel}&limit=1`, { headers: h });
    if (r.ok) { const d = await r.json(); if (d.length) return pick(d[0]); }
  } catch {}
  return null;
}

async function supabaseFindProduct(cfg, code, branch = '') {
  const { url, anonKey, tableName, stockTableName } = cfg;
  if (!url || !anonKey) return null;
  const base = url.replace(/\/$/, '');
  const h = { apikey: anonKey, Authorization: `Bearer ${anonKey}` };
  const col = encodeURIComponent('"รหัสสินค้า"');
  const priceTable = tableName || 'product_price';
  const stockTable = stockTableName || 'product_stock';

  // แปลงบาร์โค้ดเป็นรหัสสินค้าก่อน — ยิงบาร์โค้ดสำรองหรือบาร์โค้ดแพ็คก็หาเจอ
  const hit = await resolveBarcode(base, h, code, branch);
  const lookupCode = hit?.code || code;            // ราคาผูกกับหน่วยที่ยิง
  const stockCode = hit?.masterCode || lookupCode;  // สต็อกผูกกับรหัสหลักเสมอ
  const extra = hit ? {
    _scannedBarcode: code, _masterCode: hit.masterCode, _stockRatio: hit.ratio,
    _isPack: hit.isPack, _onHandPieces: hit.onHandPieces, _barcodeConflict: hit.conflict || null,
    _unit: hit.unit, _name: hit.name, _price: hit.price, _cost: hit.cost,
  } : {};

  // Query both tables in parallel
  const [priceRes, stockRes] = await Promise.all([
    fetch(`${base}/rest/v1/${priceTable}?${col}=eq.${encodeURIComponent(lookupCode)}${stockBranchFilter(branch)}&limit=1`, { headers: h }),
    fetch(`${base}/rest/v1/${stockTable}?${col}=eq.${encodeURIComponent(stockCode)}${stockBranchFilter(branch)}&limit=1`, { headers: h }),
  ]);

  let priceRow = null, stockRow = null;
  if (priceRes.ok) { const arr = await priceRes.json(); if (arr.length) priceRow = arr[0]; }
  if (stockRes.ok) { const arr = await stockRes.json(); if (arr.length) stockRow = arr[0]; }

  if (!priceRow && !stockRow) {
    if (!priceRes.ok && !stockRes.ok) throw new Error(`Supabase ${priceRes.status}/${stockRes.status}`);
    return null;
  }
  // Merge: prefer price row fields, fill missing from stock row
  return { ...(stockRow||{}), ...(priceRow||{}), ...extra };
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
  const num = (v) => parseFloat(String(v ?? '0').replace(/[^\d.-]/g, '')) || 0;
  return {
    source: 'supabase', id,
    barcode: String(row._scannedBarcode || get('บาร์โค้ด', 'barcode') || id || fallbackCode || ''),
    productCode: id,
    // หน่วย/ราคา/ชื่อ ต้องเป็นของหน่วยที่ยิง (1x6, 1x36) ไม่ใช่ของรหัสหลัก
    name: String(row._name || get('ชื่อสินค้า', 'product_name', 'name') || '(ไม่มีชื่อ)'),
    category: String(get('ประเภท', 'category') || 'อื่นๆ'),
    unit: String(row._unit || get('หน่วยนับ', 'หน่วย', 'unit') || 'ชิ้น'),
    price: row._price || num(get('ราคา', 'price', 'ราคาขาย')),
    cost:  row._cost  || num(get('ทุนเฉลี่ย', 'ต้นทุน', 'cost', 'ราคาทุน')),
    // ยิงบาร์โค้ดสำรอง/แพ็ค — เก็บตัวที่ยิงจริงไว้ ไม่ให้กลายเป็นบาร์โค้ดหลักเงียบ ๆ
    scannedBarcode: row._scannedBarcode || '',
    masterCode: row._masterCode || id,          // รหัสที่สต็อกอยู่จริง
    stockRatio: Number(row._stockRatio) || 1,   // 1 หน่วยนี้ = กี่ชิ้น
    isPack: !!row._isPack,
    onHandPieces: row._onHandPieces ?? null,
    barcodeConflict: row._barcodeConflict || null,
  };
}

// ค้นชื่อจริงในระบบจากบาร์โค้ด — ทีเดียวหลายตัว ลดจำนวนรอบเรียก
// ลองทั้งคอลัมน์ "บาร์โค้ด" และ "รหัสสินค้า" เพราะบางสินค้าใช้รหัสร้านเป็นบาร์โค้ด
async function lookupBarcodes(cfg, codes) {
  const out = new Map();
  const list = [...new Set(codes.map(c => String(c || '').trim()).filter(Boolean))];
  if (!cfg?.url || !cfg?.anonKey || !list.length) return out;
  const base = cfg.url.replace(/\/$/, '');
  const h = { apikey: cfg.anonKey, Authorization: `Bearer ${cfg.anonKey}` };
  const table = cfg.tableName || 'product_price';
  const inList = list.map(c => `"${c.replace(/"/g, '')}"`).join(',');
  const qs = `or=(${encodeURIComponent('"บาร์โค้ด"')}.in.(${inList}),${encodeURIComponent('"รหัสสินค้า"')}.in.(${inList}))&limit=500`;
  try {
    const res = await fetch(`${base}/rest/v1/${table}?${qs}`, { headers: h });
    if (!res.ok) return out;
    for (const row of await res.json()) {
      const mapped = mapSupabaseRow(row);
      for (const k of [row['บาร์โค้ด'], row['รหัสสินค้า'], mapped.barcode, mapped.productCode]) {
        const key = String(k || '').trim();
        if (key && !out.has(key)) out.set(key, mapped);
      }
    }
  } catch {}
  return out;
}

async function sbFetch(url, key, table, rawQS) {
  const res = await fetch(`${url.replace(/\/$/, '')}/rest/v1/${table}?${rawQS}`, {
    headers: { apikey: key, Authorization: `Bearer ${key}` }
  });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  return res.json();
}

// คำที่ไม่ใช่ชื่อเฉพาะ — ถ้าเอาไปค้นจะเจอผู้ขายรายไหนก็ได้
const VENDOR_NOISE = /(บริษัท|บมจ\.?|บจก\.?|หจก\.?|หสน\.?|ห้างหุ้นส่วน(จำกัด|สามัญ)?|ร้าน|จำกัด|มหาชน|สำนักงานใหญ่|สาขา\s*\S*|co\.?,?\s*ltd\.?|company|limited|part\.?)/gi;
const vendorCore = (name) => String(name || '').replace(VENDOR_NOISE, ' ').replace(/[()ฯ.,\-]/g, ' ').replace(/\s+/g, ' ').trim();

// ลำดับความแม่น: เลขภาษี 13 หลัก → ชื่อตรงเป๊ะ → ชื่อเฉพาะ (คำที่ยาวสุด)
// เจอมากกว่า 1 ราย = เดาไม่ได้ ไม่เติมให้ ปล่อยให้คนกรอก
async function lookupVendorREST(sbUrl, sbKey, vendorName, taxId) {
  if (!sbUrl || !sbKey) return null;
  const h = { apikey: sbKey, Authorization: `Bearer ${sbKey}` };
  const sel = encodeURIComponent('"รหัส"');
  const q = async (col, op, kw) => {
    const val = op === 'ilike' ? `*${encodeURIComponent(kw)}*` : encodeURIComponent(kw);
    const r = await fetch(`${sbUrl}/rest/v1/vendor_info?select=${sel}&${encodeURIComponent('"' + col + '"')}=${op}.${val}&limit=2`, { headers: h });
    if (!r.ok) return null;
    const d = await r.json();
    return d.length === 1 ? (d[0]['รหัส'] ?? null) : null;   // กำกวมก็ไม่เดา
  };

  const tax = String(taxId || '').replace(/\D/g, '');
  if (tax.length === 13) {
    for (const col of ['เลขประจำตัวผู้เสียภาษี', 'เลขผู้เสียภาษี', 'tax_id']) {
      try { const hit = await q(col, 'eq', tax); if (hit) return { no: hit, by: 'tax' }; } catch {}
    }
  }
  if (!vendorName) return null;

  const exact = await q('ชื่อ-นามสกุล', 'eq', vendorName.trim());
  if (exact) return { no: exact, by: 'name' };

  const core = vendorCore(vendorName);
  if (core) {
    const hit = await q('ชื่อ-นามสกุล', 'ilike', core);
    if (hit) return { no: hit, by: 'near' };
    // คำที่ยาวสุดมักเป็นชื่อเฉพาะจริง ต่างจาก "คำแรก" ที่มักเป็น "บริษัท"
    const longest = core.split(' ').filter(w => w.length >= 3).sort((a, b) => b.length - a.length)[0];
    if (longest) { const h2 = await q('ชื่อ-นามสกุล', 'ilike', longest); if (h2) return { no: h2, by: 'near' }; }
  }
  return null;
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
  const header = ['รหัสสินค้า', 'ชื่อสินค้า', 'หน่วย', 'จำนวนนับ', 'Location', 'ราคาขาย', 'มูลค่ารวม', 'ทุนเฉลี่ย', 'เวลาสแกน'];
  const rows = data.map(d => [
    d.barcode, d.productName, d.unit || '', d.qty,
    d.location || '', Number(d.price || 0), Number(d.price || 0) * Number(d.qty || 0), Number(d.cost || 0),
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
  ws['!cols'] = [{ wch: 20 }, { wch: 30 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 20 }];
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
  const rows = sub.data.map(d => `<tr><td>${d.barcode}</td><td>${d.productName || ''}</td><td style="text-align:center">${d.qty}</td><td style="text-align:center">${d.unit || ''}</td>${d.location ? `<td>${d.location}</td>` : '<td>-</td>'}<td style="text-align:right">${Number(d.price||0).toLocaleString()}</td><td style="text-align:right">${(Number(d.price||0)*Number(d.qty||0)).toLocaleString()}</td></tr>`).join('');
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
  overlay.innerHTML = `<div class="pr-header"><div><h1>Stock Count Report</h1><p class="meta">เลขที่เอกสาร : <b>${sub.docNo || sub.id}</b></p><p class="meta">วันที่ : <b>${dateStr}</b></p><p class="meta">พนักงาน : <b>${sub.counter}</b></p><p class="meta">รายการ : <b>${sub.itemCount}</b> &nbsp; รวม : <b>${sub.totalQty}</b></p>${sub.note ? `<p class="meta">หมายเหตุ : <b>${sub.note}</b></p>` : ''}</div><div style="display:flex;gap:8px;align-items:flex-start"><button class="pr-print" id="__pdf_print_btn__">🖨️ พิมพ์ / PDF</button><button class="pr-close" id="__pdf_close_btn__">✕</button></div></div><table><thead><tr><th>รหัสสินค้า</th><th>ชื่อสินค้า</th><th style="text-align:center">จำนวน</th><th style="text-align:center">หน่วย</th><th>Location</th><th style="text-align:right">ราคาขาย</th><th style="text-align:right">มูลค่ารวม</th></tr></thead><tbody>${rows}</tbody></table><p style="margin-top:16px;font-size:11px;color:#64748b;text-align:right">พิมพ์เมื่อ ${new Date().toLocaleString('th-TH')}</p>`;
  document.body.appendChild(overlay);
  document.getElementById('__pdf_print_btn__')?.addEventListener('click', () => window.print());
  document.getElementById('__pdf_close_btn__')?.addEventListener('click', () => { document.getElementById(OVERLAY_ID)?.remove(); document.getElementById(STYLE_ID)?.remove(); });
}

// ค่ากลางจาก /api/config — ตั้งครั้งเดียว ทุกเครื่องใช้ร่วม ค่าใน const เป็นค่าตั้งต้น
let CLOUD_SETTINGS = {};
async function loadCloudSettings() {
  try {
    const r = await fetch('/api/config');
    const j = await r.json();
    if (j?.settings) CLOUD_SETTINGS = j.settings;
  } catch {}
  return CLOUD_SETTINGS;
}
const setting = (key, fallback = '') => CLOUD_SETTINGS[key] || fallback;

// 2 สาขา — ชื่อแก้ได้ในหน้าตั้งค่า ไม่ต้อง redeploy
const BRANCHES = () => [
  { id: '1', name: setting('branch_1_name', 'สาขา 1') },
  { id: '2', name: setting('branch_2_name', 'สาขา 2') },
];
const branchName = (id) => id === 'all' ? 'ทุกสาขา' : (BRANCHES().find(b => b.id === String(id))?.name || `สาขา ${id}`);
// สีประจำสาขา — ให้เห็นแวบเดียวว่ากำลังทำงานสาขาไหน กันคีย์ผิดสาขา
const BRANCH_INK = { '1': '#255771', '2': '#8A5A1C', all: '#475569' };
const BRANCH_SOFT = { '1': '#EAF0F4', '2': '#FDF6EC', all: '#F6F7F8' };

// ค่าที่อยู่ในคอลัมน์สาขาของ product_stock — ต่างจากรหัสสาขาในแอปได้
const stockBranchFilter = (branch) => {
  const col = setting('branch_stock_column', '');
  if (!col || !branch || branch === 'all') return '';   // ไม่ตั้ง = อ่านรวมทุกสาขาเหมือนเดิม
  const val = setting(`branch_${branch}_value`, branch);
  return `&${encodeURIComponent('"' + col + '"')}=eq.${encodeURIComponent(val)}`;
};

// โฟลเดอร์ Drive ไม่ hardcode — มาจาก /api/config (env → app_settings)

/* ---- ส่งขึ้น Drive: จำว่าไฟล์ไหนขึ้นแล้ว กันส่งซ้ำ + กันชื่อไทยทำ base64 พัง ---- */
const UPLOAD_LOG_KEY = 'drive_upload_log';
const uploadKey = (subId, type) => `${subId}::${type}`;
function readUploadLog() { try { return JSON.parse(localStorage.getItem(UPLOAD_LOG_KEY) || '{}'); } catch { return {}; } }
function writeUploadLog(log) {
  try { localStorage.setItem(UPLOAD_LOG_KEY, JSON.stringify(log)); } catch {}
  try { window.dispatchEvent(new Event('drive-log')); } catch {}
}
function useUploadLog() {
  const [log, setLog] = useState(readUploadLog);
  useEffect(() => {
    const h = () => setLog(readUploadLog());
    window.addEventListener('drive-log', h);
    return () => window.removeEventListener('drive-log', h);
  }, []);
  return log;
}
// btoa พังทันทีเมื่อเจออักขระไทย → "string did not match the expected pattern"
function utf8ToBase64(str) {
  const bytes = new TextEncoder().encode(str);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}
// ชื่อไฟล์: ตัดอักขระที่ Drive ไม่รับ แต่เก็บภาษาไทยไว้
const safeFilename = (name) => String(name ?? 'file').replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').replace(/\s+/g, ' ').trim().slice(0, 120);

// สถานะจริงอยู่ใน DB (คอลัมน์ drive_* ของใบ) — localStorage เป็นแค่แคชตอนเน็ตหลุด
const driveTypeOf = (sub) => (sub?._kind === 'invoice' || sub?.invoiceNo !== undefined) ? 'csv_header'
  : (sub?.featureType || sub?._kind) === 'stock_compare' ? 'csv_adjust' : 'csv_count';
function driveEntryOf(sub, cache = {}, type) {
  type = type || driveTypeOf(sub);
  const d = sub?.drive;
  if (d?.status === 'ok')     return { ok: true,  link: d.url, filename: d.filename, at: d.uploadedAt, tries: d.tries, fromDb: true };
  if (d?.status === 'failed') return { ok: false, err: d.error, filename: d.filename, at: d.uploadedAt, tries: d.tries, fromDb: true };
  return cache[uploadKey(sub.id, type)] || null;
}
async function recordDriveResult(sub, entry, by, endpoint = '/api/submission') {
  try {
    const res = await fetch(endpoint, {
      method: 'PATCH', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id: sub.id, drive: { ok: !!entry.ok, url: entry.link || '', filename: entry.filename || '', error: entry.err || '' } }),
    });
    const j = await res.json().catch(() => null);
    return j?.submission || null;
  } catch { return null; }   // บันทึกสถานะไม่ได้ก็ยังมีแคชในเครื่อง
}

const uploadInFlight = new Set();
// ส่งจริงครั้งเดียวต่อ (ใบ, ชนิดไฟล์) — ต้อง force ถ้าจะส่งซ้ำ
async function driveUpload({ subId, type, filename, mimeType, content, isBase64 = false, bom = false, folderId, force = false }) {
  if (!folderId) return { ok: false, err: 'ยังไม่ได้ตั้งโฟลเดอร์ Drive — ตั้งที่ บันทึกแล้ว → ตั้งค่าโฟลเดอร์ Drive' };
  const key = uploadKey(subId, type);
  const log = readUploadLog();
  if (!force && log[key]?.ok) return { ...log[key], skipped: true };
  if (uploadInFlight.has(key)) return { ok: false, err: 'กำลังส่งอยู่ รอสักครู่' };
  uploadInFlight.add(key);
  const name = safeFilename(filename);
  try {
    const clean = isBase64 ? content : String(content).replace(/^\uFEFF/, '').replace(/[\r\n]+$/, '');
    const body = isBase64
      ? { content, isBase64: true }
      : { content: utf8ToBase64((bom ? '\uFEFF' : '') + clean), isBase64: true };
    const res = await fetch('/api/drive-upload', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: name, mimeType, folderId, ...body }),
    });
    const data = await res.json().catch(() => ({ ok: false, error: 'อ่านคำตอบจากเซิร์ฟเวอร์ไม่ได้' }));
    if (!res.ok || !data.ok) throw new Error(data.error || `HTTP ${res.status}`);
    const entry = { ok: true, link: data.link, filename: name, at: new Date().toISOString(), type, tries: (log[key]?.tries || 0) + 1 };
    writeUploadLog({ ...readUploadLog(), [key]: entry });
    return entry;
  } catch (e) {
    const entry = { ok: false, err: e.message, filename: name, at: new Date().toISOString(), type, tries: (log[key]?.tries || 0) + 1 };
    writeUploadLog({ ...readUploadLog(), [key]: entry });
    return entry;
  } finally { uploadInFlight.delete(key); }
}
// ใบนับ → บันทึกมือ · เทียบยอด → ปรับยอด
function subFolderId(sub) {
  return (sub.featureType || 'recorder') === 'stock_compare'
    ? setting('drive_folder_stock_adjust')
    : setting('drive_folder_manual');
}

// ไฟล์บิลซื้อ: 2 ชีต bill_header + invoice ตามสเปคเดิม
function buildInvoiceXlsxBase64(inv) {
  const h = inv.header || {};
  const lines = Array.isArray(inv.lines) ? inv.lines : [];
  const vs = vatSummary(lines) || {};
  const rawAmt = lines.reduce((a, p) => a + (Number(p.amount) || 0), 0);
  const wb = XLSX.utils.book_new();

  const ws1 = XLSX.utils.aoa_to_sheet([
    ['invoice_no','invoice_date','vendor_name','vendor_tax_id','document_type','vendor_address',
     'total_amount','total_discount','net_total','excl_vat','vat_amount','vendor_branch','vendor_no','price_type'],
    [inv.invoiceNo || h.invoice_no || '', inv.invoiceDate || h.invoice_date || '', inv.vendorName || h.vendor_name || '',
     h.vendor_tax_id || '', h.document_type || '', h.vendor_address || '',
     rawAmt, vs.sdTot ?? 0, Number(inv.netTotal) || 0, vs.excl ?? 0, vs.vatAmt ?? 0,
     h.vendor_branch || '', h.vendor_no || '', h.price_type || 'incl'],
  ]);
  ws1['!cols'] = [{wch:16},{wch:12},{wch:28},{wch:16},{wch:12},{wch:34},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:10}];
  XLSX.utils.book_append_sheet(wb, ws1, 'bill_header');

  const ws2 = XLSX.utils.aoa_to_sheet([
    ['invoice_no','no','description','carton_size','carton','ea','qty','price_ea','amount',
     'special_discount','amount_sd','total','excl_vat','vat_amt','vat','barcode'],
    ...lines.map((d, i) => [inv.invoiceNo || h.invoice_no || '', i + 1, d.description || d.name || '',
      d.carton_size ?? '', d.carton ?? '', d.ea ?? '', d.qty ?? '', d.price_ea ?? '', d.amount ?? '',
      d.special_discount ?? '', d.amount_sd ?? '', d.total ?? '', d.excl_vat ?? '', d.vat_amt ?? '', d.vat ?? '', d.barcode || '']),
  ]);
  ws2['!cols'] = [{wch:16},{wch:6},{wch:34},{wch:11},{wch:8},{wch:8},{wch:8},{wch:10},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:10},{wch:8},{wch:16}];
  XLSX.utils.book_append_sheet(wb, ws2, 'invoice');

  return XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
}

// CSV สำหรับ import เข้า POS — 2 ไฟล์ต่อบิล ชื่อไฟล์ = doc_no
// bill_header: 1 บรรทัดต่อบิล (มีหัวคอลัมน์)
// imp_data:    barcode,qty,ราคาต่อหน่วยหลังส่วนลด,0 — ไม่มีหัวคอลัมน์ ตามที่ POS รับ
// RC — บันทึกมือ: barcode,qty,price,0 (ไม่มีหัวคอลัมน์)
function buildRecorderCsv(rows) {
  return (rows || []).map(d => [
    // รหัสหลัก + จำนวนชิ้น — POS เก็บสต็อกที่รหัสหลักรหัสเดียว
    d.masterCode || d.barcode || '',
    d.pieces != null ? Number(d.pieces) : (Number(d.qty) || 0) * (Number(d.stockRatio) || 1),
    Number(d.price) || 0, 0,
  ].join(',')).join('\r\n');
}

// ST — ปรับยอด: barcode,adjust_stock (+/- ให้ตรงสต็อกจริง · ว่างถ้าไม่พบในระบบ)
async function buildAdjustCsv(rows, sbUrl, sbKey, stockTable = 'product_stock', branch = '1') {
  const codes = [...new Set((rows || []).map(d => String(d.masterCode || d.barcode || '')).filter(Boolean))];
  const onHand = {};
  if (sbUrl && sbKey && codes.length) {
    const h = { apikey: sbKey, Authorization: `Bearer ${sbKey}` };
    const qc = (c) => encodeURIComponent(`"${c}"`);
    for (let i = 0; i < codes.length; i += 50) {
      const inList = codes.slice(i, i + 50).map(encodeURIComponent).join(',');
      const url = `${sbUrl}/rest/v1/${stockTable}?${qc('รหัสสินค้า')}=in.(${inList})&select=${qc('รหัสสินค้า')},${qc('รวม')}${stockBranchFilter(branch)}`;
      const r = await fetch(url, { headers: h });
      if (!r.ok) throw new Error('อ่านยอดคงเหลือไม่ได้: ' + (await r.text()).slice(0, 120));
      for (const row of await r.json()) {
        onHand[String(row['รหัสสินค้า'] || '')] = parseInt(String(row['รวม'] ?? '0').replace(/[^\d-]/g, '')) || 0;
      }
    }
  }
  // รวมบาร์โค้ดซ้ำก่อน — นับหลายรอบต้องเป็นยอดเดียว
  const counted = {};
  for (const d of rows || []) {
    const bc = String(d.masterCode || d.barcode || '');
    if (!bc) continue;
    const pieces = d.pieces != null ? Number(d.pieces) : (Number(d.qty) || 0) * (Number(d.stockRatio) || 1);
    counted[bc] = (counted[bc] || 0) + pieces;
  }
  const lines = Object.keys(counted).map(bc => {
    const cur = onHand[bc];
    if (cur == null) return `${bc},`;                     // ไม่พบในระบบ → ปล่อยว่าง
    const adj = counted[bc] - cur;
    return `${bc},${adj > 0 ? '+' + adj : adj}`;
  });
  const missing = Object.keys(counted).filter(bc => onHand[bc] == null).length;
  return { csv: lines.join('\r\n'), missing, total: lines.length };
}

function buildInvoiceCsvPair(inv) {
  const h = inv.header || {};
  const lines = Array.isArray(inv.lines) ? inv.lines : [];
  const vs = vatSummary(lines) || {};
  const rawAmt = lines.reduce((a, p) => a + (Number(p.amount) || 0), 0);
  const esc = (v) => { const t = v == null ? '' : String(v); return /[",\r\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t; };

  const headerCsv =
    'invoice_no,invoice_date,vendor_name,vendor_tax_id,document_type,vendor_address,' +
    'total_amount,total_discount,net_total,excl_vat,vat_amount,vendor_branch,vendor_no,price_type\r\n' +
    [inv.invoiceNo || h.invoice_no || '', inv.invoiceDate || h.invoice_date || '',
     h.vendor_name ?? inv.vendorName ?? '', h.vendor_tax_id ?? '', h.document_type ?? '', h.vendor_address ?? '',
     +rawAmt.toFixed(2) || 0, vs.sdTot ?? 0, Number(inv.netTotal) || vs.netTotal || 0,
     vs.excl ?? 0, vs.vatAmt ?? 0, h.vendor_branch ?? '', h.vendor_no ?? '', h.price_type ?? 'incl'].map(esc).join(',');

  // ราคาต่อชิ้นหลังหักส่วนลด 4 ทศนิยม · ไม่มีบาร์โค้ดปล่อยว่าง
  const impCsv = lines.map(d => {
    const qty = d.qty != null ? +d.qty : 0;
    const pea = d.price_ea != null ? +d.price_ea : null;
    const sd  = d.special_discount != null ? +d.special_discount : 0;
    const tot = (qty > 0 && pea != null) ? qty * pea - sd : null;
    const per = (tot != null && qty > 0) ? (tot / qty) : 0;
    return [d.barcode || '', qty, per.toFixed(4), 0].join(',');
  }).join('\r\n');

  return { headerCsv, impCsv };
}

// เลขจริงออกจากเซิร์ฟเวอร์ (next_doc_no) — ที่นี่แค่ป้ายชั่วคราวสำหรับใบที่ยังส่งไม่ขึ้น
function provisionalDocNo(prefix = 'RC') {
  const d = new Date();
  const dateStr = d.getFullYear() + String(d.getMonth()+1).padStart(2,'0') + String(d.getDate()).padStart(2,'0');
  return `${prefix}-${dateStr} (รอเลข)`;
}

// สีประจำแต่ละหน้าในเมนูล่าง — แตะแล้วเห็นชัดว่าเปลี่ยนหน้าแล้ว
const NAV_TONE = {
  count:          { main: '#35706A', soft: '#EAF1F0' },
  review:         { main: '#B45309', soft: '#FFFBEB' },
  my_submissions: { main: '#2F6E90', soft: '#EAF0F4' },
  compare:        { main: '#7C4DFF', soft: '#F3EDFF' },
  invoice:        { main: '#2F6E90', soft: '#EAF0F4' },
  dashboard:      { main: '#0F172A', soft: '#F6F7F8' },
  inbox:          { main: '#B45309', soft: '#FFFBEB' },
};

function defaultViewFor(user) {
  if (user?.role === 'manager') return 'inbox';
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
  const [view, setView] = useState(() => safeGet('lastView', 'count') || 'count');
  const [loaded, setLoaded] = useState(false);
  const [currentUser, setCurrentUser] = useState(() => {
    try {
      const v = localStorage.getItem('currentUser');
      const u = v ? JSON.parse(v) : null;
      // ผู้ใช้ที่ค้างจากก่อนมี 2 สาขา ไม่รู้ว่าทำงานสาขาไหน — ให้เลือกใหม่ ดีกว่าเดาผิดสาขา
      if (u && u.role !== 'manager' && !u.branch) return null;
      return u;
    } catch { return null; }
  });
  const [supabaseConfig, setSupabaseConfig] = useState({ url: '', anonKey: '', tableName: 'product_price', stockTableName: 'product_stock' });
  const [dataSource, setDataSource] = useState('none');
  const [lastSyncAt, setLastSyncAt] = useState(null);
  const [connectionStatus, setConnectionStatus] = useState('unknown');
  const [countDate, setCountDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [countDraft, setCountDraft] = useState({ barcode: '', qty: '', checkResult: null, error: '' });
  const [publicPage, setPublicPage] = useState(() => safeGet('lastPublicPage', '') || null);   // 'staff' | 'report' | 'settings'
  const [pickedAt, setPickedAt] = useState(null);   // เลือกชื่อไว้กี่โมง — โชว์ในหน้ายืนยันก่อนส่ง
  // ผู้จัดการสลับดูได้ (all/1/2) · พนักงานล็อกไว้ที่สาขาที่เลือกตอนเข้าใช้
  const [viewBranch, setViewBranch] = useState(() => {
    try { const u = JSON.parse(localStorage.getItem('currentUser') || 'null');
      return u && u.role !== 'manager' && u.branch ? String(u.branch) : 'all'; } catch { return 'all'; }
  });
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

      // ยังไม่ได้ตั้งค่าในเครื่องนี้ → ดึงค่ากลางจากเซิร์ฟเวอร์ ทุกเครื่องพร้อมใช้ทันที
      if (!cfg.url || !cfg.anonKey) {
        try {
          const res = await fetch('/api/config');
          const j = await res.json();
          if (j?.configured && j.url && j.anonKey) {
            cfg = {
              url: j.url, anonKey: j.anonKey,
              tableName: cfg.tableName || j.tableName || 'product_price',
              stockTableName: cfg.stockTableName || j.stockTableName || 'product_stock',
            };
            setSupabaseConfig(cfg);
            storage.set('supabaseConfig', JSON.stringify(cfg)).catch(() => {});
          }
        } catch {}
      }
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

  // ร่างขึ้นเซิร์ฟเวอร์ — นับเครื่องหนึ่ง ไปต่อเครื่องอื่นได้
  const [draftSync, setDraftSync] = useState({ busy: '', msg: '', err: '' });

  useEffect(() => { loadCloudSettings(); }, []);

  const pushDraft = async () => {
    if (!currentUser) return;
    setDraftSync({ busy: 'up', msg: '', err: '' });
    try {
      const res = await fetch('/api/draft', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          counter_id: currentUser.id, counter_name: currentUser.name,
          feature: currentUser.feature || 'recorder',
          branch: String(currentUser.branch || '1'),
          device_id: safeGet('deviceId', '') || null,
          entries: countEntries.filter(e => e.counterId === currentUser.id),
        }),
      });
      const j = await res.json();
      if (!res.ok || j.error) throw new Error(j.error || 'HTTP ' + res.status);
      setDraftSync({ busy: '', msg: `บันทึกขึ้นเซิร์ฟเวอร์แล้ว ${j.saved} บรรทัด`, err: '' });
    } catch (e) {
      setDraftSync({ busy: '', msg: '', err: e.message });   // ของยังอยู่ในเครื่อง ไม่หาย
    }
  };

  const pullDraft = async () => {
    if (!currentUser) return;
    setDraftSync({ busy: 'down', msg: '', err: '' });
    try {
      const res = await fetch(`/api/draft?counter_id=${encodeURIComponent(currentUser.id)}&feature=${encodeURIComponent(currentUser.feature || 'recorder')}&branch=${encodeURIComponent(String(currentUser.branch || '1'))}`);
      const j = await res.json();
      if (!res.ok || j.error) throw new Error(j.error || 'HTTP ' + res.status);
      const mine = (j.entries || []).map(e => ({ ...e, counterId: currentUser.id, counterName: currentUser.name, id: e.id || `${Date.now()}_${Math.random()}` }));
      // ทับร่างของคนนี้ในเครื่องนี้ ของคนอื่นไม่แตะ
      setCountEntries(prev => [...prev.filter(e => e.counterId !== currentUser.id), ...mine]);
      setDraftSync({ busy: '', msg: mine.length ? `ดึงร่างมา ${mine.length} บรรทัด` : 'ไม่มีร่างบนเซิร์ฟเวอร์', err: '' });
    } catch (e) { setDraftSync({ busy: '', msg: '', err: e.message }); }
  };

  const handleLogin = (user) => { setCurrentUser(user); setView(defaultViewFor(user)); setPickedAt(Date.now());
    setViewBranch(user.role === 'manager' ? 'all' : String(user.branch || '1')); storage.set('currentUser', JSON.stringify(user)).catch(() => {}); };
  // สลับฟีเจอร์โดยไม่ต้องออกจากระบบ — ร่างแยกตามฟีเจอร์ ไม่หาย
  const switchFeature = (f) => {
    if (!currentUser || f === currentUser.feature) return;
    const next = { ...currentUser, feature: f };
    setCurrentUser(next);
    try { localStorage.setItem('currentUser', JSON.stringify(next)); } catch {}
    setView(f === 'invoice' ? 'invoice' : 'count');
    setMenuOpen(false);
  };

  const handleLogout = () => {
    setCurrentUser(null);
    storage.delete('currentUser').catch(() => {});
    try { localStorage.removeItem('lastView'); localStorage.removeItem('lastPublicPage'); } catch {}
  };

  // จำหน้าที่เปิดอยู่ กด refresh แล้วไม่เด้งกลับหน้าแรก
  useEffect(() => { safeSet('lastView', view); }, [view]);
  useEffect(() => {
    if (publicPage) safeSet('lastPublicPage', publicPage);
    else { try { localStorage.removeItem('lastPublicPage'); } catch {} }
  }, [publicPage]);
  useEffect(() => {
    if (currentUser) { try { localStorage.setItem('currentUser', JSON.stringify(currentUser)); } catch {} }
  }, [currentUser]);


  const checkBarcode = async (barcode) => {
    const trimmed = barcode.trim();
    if (!trimmed) return null;
    if (supabaseConfig.url && supabaseConfig.anonKey) {
      try {
        const row = await supabaseFindProduct(supabaseConfig, trimmed, String(currentUser?.branch || ''));
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
      featureType: currentUser?.feature || 'recorder',
      barcode: entry.barcode, productName: entry.productName, productId: entry.productId || '',
      // สต็อกอยู่ที่รหัสหลัก · ตัวคูณแปลงหน่วยที่ยิงเป็นชิ้น (แพ็ค 1x6 → 6)
      masterCode: entry.masterCode || entry.productId || entry.barcode,
      stockRatio: Number(entry.stockRatio) || 1,
      isPack: !!entry.isPack,
      unit: entry.unit || '', price: entry.price || 0, cost: entry.cost || 0,
      qty: parseInt(entry.qty) || 0, notFound: !!entry.notFound,
      location: entry.location || '',
      counter: currentUser?.name || 'พนักงาน', counterId: currentUser?.id || 'unknown',
      timestamp,
    }, ...prev]);
  };

  const deleteCountEntry = (id) => setCountEntries(prev => prev.filter(e => e.id !== id));
  const clearMyEntries = () => setCountEntries(prev => prev.filter(e => e.counterId !== currentUser?.id));

  const submittingRef = useRef(false);
  const submitForReview = async (grouped, note) => {
    if (submittingRef.current) return null;   // กันส่งซ้อน
    submittingRef.current = true;
    try {
    const prefix = currentUser?.feature === 'stock_compare' ? 'ST' : 'RC';
    const docNo = provisionalDocNo(prefix);   // เซิร์ฟเวอร์จะเปลี่ยนเป็นเลขจริงตอนบันทึก
    const now = new Date().toISOString();
    const sub = {
      id: `sub${Date.now()}_${Math.random().toString(36).slice(2,7)}`, docNo,
      counter: currentUser?.name || 'พนักงาน', counterId: currentUser?.id || 'unknown',
      featureType: currentUser?.feature || 'recorder',
      branch: String(currentUser?.branch || '1'),
      submittedAt: now,
      startedAt: grouped.reduce((min, g) => g.scannedAt && g.scannedAt < min ? g.scannedAt : min, now),
      note: note || '', status: 'pending', reviewedAt: null, reviewedBy: null, reviewNote: '',
      itemCount: grouped.length, totalQty: grouped.reduce((s, g) => s + g.qty, 0), data: grouped,
      reviseOf: reviseCount?.id || null,
    };
    setSubmissions(prev => [sub, ...prev]);

    // ขึ้นเซิร์ฟเวอร์ — ผู้จัดการเปิดจากเครื่องไหนก็เห็น
    let saved = sub;
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const res = await fetch('/api/submission', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ submission: { ...saved, deviceId: safeGet('deviceId', '') || null } }),
        });
        const j = await res.json();
        if (res.status === 409 && j.duplicate && attempt === 0) {
          continue;   // เซิร์ฟเวอร์ออกเลขใหม่ให้เอง
        }
        if (!res.ok || j.error) throw new Error(j.error || 'HTTP ' + res.status);
        // ใช้แถวจากเซิร์ฟเวอร์เป็นตัวจริง (id ตรงกันทุกเครื่อง)
        // deduped = เซิร์ฟเวอร์เจอใบเดิมของคนเดิม ไม่สร้างใบใหม่ให้
        setSubmissions(prev => {
          const others = prev.filter(s => s.id !== sub.id && s.docNo !== j.submission.docNo);
          return [{ ...j.submission, synced: true }, ...others];
        });
        saved = j.submission;
        break;
      } catch (e) {
        setSubmissions(prev => prev.map(s => s.id === sub.id ? { ...s, syncError: e.message || 'ส่งไม่สำเร็จ' } : s));
        break;
      }
    }

    clearReviseCount();   // ส่งรอบแก้แล้ว ออกจากโหมดแก้
    // ส่งเป็นใบแล้ว ร่างบนเซิร์ฟเวอร์ต้องหาย ไม่ให้เครื่องอื่นดึงของเก่ากลับมา
    try {
      await fetch(`/api/draft?counter_id=${encodeURIComponent(currentUser?.id || '')}&feature=${encodeURIComponent(currentUser?.feature || 'recorder')}&branch=${encodeURIComponent(String(currentUser?.branch || '1'))}`, { method: 'DELETE' });
    } catch {}
    return saved;
    } finally { submittingRef.current = false; }
  };

  // ใบที่ขึ้นเซิร์ฟเวอร์ไม่สำเร็จ ค้างอยู่ในเครื่องและยังไม่มีเลข — ส่งซ้ำได้
  const resendSubmission = async (id) => {
    const sub = submissions.find(x => x.id === id);
    if (!sub) return;
    setSubmissions(prev => prev.map(x => x.id === id ? { ...x, syncError: '', sending: true } : x));
    try {
      const res = await fetch('/api/submission', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ submission: { ...sub, syncError: undefined, sending: undefined, deviceId: safeGet('deviceId', '') || null } }),
      });
      const j = await res.json();
      if (!res.ok || j.error) throw new Error(j.error || 'HTTP ' + res.status);
      setSubmissions(prev => [{ ...j.submission, synced: true }, ...prev.filter(x => x.id !== id && x.docNo !== j.submission.docNo)]);
    } catch (e) {
      setSubmissions(prev => prev.map(x => x.id === id ? { ...x, sending: false, syncError: e.message || 'ส่งไม่สำเร็จ' } : x));
    }
  };

  const reviewSubmission = (id, status, reviewNote) => {
    setSubmissions(prev => prev.map(s => s.id === id ? { ...s, status, reviewNote: reviewNote || '', reviewedAt: new Date().toISOString(), reviewedBy: currentUser?.name || 'ผู้จัดการ' } : s));
    fetch('/api/submission', {
      method: 'PATCH', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id, status, review_note: reviewNote || '', reviewed_by: currentUser?.name || 'ผู้จัดการ' }),
    }).catch(() => {});
  };

  const deleteSubmission = (id) => {
    setSubmissions(prev => prev.filter(s => s.id !== id));
    fetch('/api/submission?id=' + encodeURIComponent(id), { method: 'DELETE' }).catch(() => {});
  };

  // ดึงใบจากเซิร์ฟเวอร์ — ผู้จัดการ/พนักงานเปิดเครื่องไหนก็เห็นของเดียวกัน
  const [subSync, setSubSync] = useState({ busy: false, at: null, err: '' });
  const [invSubs, setInvSubs] = useState([]);   // บิลรออนุมัติ
  const [reviseTarget, setReviseTarget] = useState(null);   // บิลที่ถูกส่งกลับและกำลังแก้
  // ใบนับที่ถูกส่งกลับและกำลังแก้ — อยู่ต่อข้ามการปิดแอป
  const [reviseCount, setReviseCount] = useState(() => {
    try { const v = localStorage.getItem('reviseCount'); return v ? JSON.parse(v) : null; } catch { return null; }
  });
  const startReviseCount = (t) => {
    setReviseCount(t);
    try { localStorage.setItem('reviseCount', JSON.stringify(t)); } catch {}
    // ยกรายการที่นับไว้เดิมกลับมาเป็นของคนนี้ แก้ต่อได้เลย ไม่ต้องสแกนใหม่ทั้งใบ
    setCountEntries(prev => [
      ...(t.data || []).map((d, i) => ({
        id: `rv${Date.now()}_${i}`,
        featureType: t.featureType || currentUser?.feature || 'recorder',
        barcode: d.barcode, productName: d.productName, productId: d.productId || '',
        unit: d.unit || '', price: d.price || 0, cost: d.cost || 0,
        qty: Number(d.qty) || 0, notFound: !!d.notFound,
        location: d.location || '', scannedAt: d.scannedAt || t.submittedAt,
        counterId: currentUser?.id || 'unknown', counterName: currentUser?.name || '',
      })),
      ...prev.filter(e => e.counterId !== currentUser?.id),
    ]);
    setView('count');
  };
  const clearReviseCount = () => {
    setReviseCount(null);
    try { localStorage.removeItem('reviseCount'); } catch {}
  };
  const [menuOpen, setMenuOpen] = useState(false);   // เมนูข้างของผู้จัดการ
  const driveLog = useUploadLog();                   // ใบที่ยังไม่ขึ้น Drive → ตัวเลขบนเมนู
  const pullSubmissions = useCallback(async (feat) => {
    setSubSync(s => ({ ...s, busy: true, err: '' }));
    try {
      const res = await fetch('/api/submission?feature=' + encodeURIComponent(feat || 'recorder'));
      const j = await res.json();
      if (!res.ok || j.error) throw new Error(j.error || 'HTTP ' + res.status);
      const server = j.submissions || [];
      // ใบจากเซิร์ฟเวอร์ชนะ ใบในเครื่องที่ยังไม่ขึ้นเก็บไว้ต่อ
      setSubmissions(prev => {
        const docs = new Set(server.map(s => s.docNo));
        return [...server, ...prev.filter(s => !docs.has(s.docNo))];
      });
      setSubSync({ busy: false, at: new Date().toISOString(), err: '' });
    } catch (e) { setSubSync({ busy: false, at: null, err: e.message }); }
  }, []);

  // บิลรออนุมัติ — ผู้จัดการเห็นในกล่องขาเข้าเดียวกัน
  const pullInvSubs = useCallback(async (me) => {
    const get = async (qs) => {
      const res = await fetch('/api/invoice-submission' + (qs || ''));
      const j = await res.json();
      if (!res.ok || j.error) throw new Error(j.error || 'HTTP ' + res.status);
      return j.submissions || [];
    };
    try {
      if (!me?.id) { setInvSubs(await get()); return; }   // ผู้จัดการ — เห็นทุกใบ
      let mine = await get('?keyed_by_id=' + encodeURIComponent(me.id));
      // id เปลี่ยน (เลือกชื่อใหม่ / เครื่องใหม่) → ใบเก่าผูกกับ id เดิม ค้นด้วยชื่อแทน
      if (!mine.length && me.name) {
        const all = await get();
        mine = all.filter(x => (x.keyedBy || '').trim() === me.name.trim());
      }
      setInvSubs(mine);
    } catch (e) { setInvSubs([]); }
  }, []);

  const reviewInvoice = async (id, status, note, patch = {}) => {
    const res = await fetch('/api/invoice-submission', {
      method: 'PATCH', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id, status, review_note: note || '', reviewed_by: currentUser?.name || 'ผู้จัดการ', ...patch }),
    });
    const j = await res.json();
    if (!res.ok || j.error) throw new Error(j.error || 'HTTP ' + res.status);
    setInvSubs(prev => prev.map(s => s.id === id ? j.submission : s));
    return j.submission;
  };

  const deleteInvoiceSub = (id) => {
    setInvSubs(prev => prev.filter(s => s.id !== id));
    fetch('/api/invoice-submission?id=' + encodeURIComponent(id), { method: 'DELETE' }).catch(() => {});
  };

  // เข้าชื่อ/สลับฟีเจอร์ → ดึงใบล่าสุดจากเซิร์ฟเวอร์
  useEffect(() => {
    if (!currentUser) return;
    if (currentUser.role === 'manager') {
      // ผู้จัดการเห็นทุกฟีเจอร์ในหน้าเดียว
      pullSubmissions('recorder');
      pullSubmissions('stock_compare');
      pullInvSubs();
    } else if ((currentUser.feature || 'recorder') === 'invoice') {
      pullInvSubs(currentUser);         // เห็นเฉพาะบิลของตัวเอง — รู้ว่าถูกส่งกลับแก้
      // ผู้จัดการรีวิวตอนไหนก็ได้ — เช็คซ้ำเป็นระยะ พนักงานไม่ต้องกดรีเฟรชเอง
      const t = setInterval(() => pullInvSubs(currentUser), 60000);
      return () => clearInterval(t);
    } else {
      pullSubmissions(currentUser.feature || 'recorder');
    }
  }, [currentUser, pullSubmissions, pullInvSubs]);

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
  const isSupabaseReady = !!(supabaseConfig.url && supabaseConfig.anonKey);

  // พนักงาน + รายงาน + ตั้งค่า เปิดดูได้เลย ไม่ต้องเลือกชื่อ/ไม่ต้องเป็นผู้จัดการ
  if (publicPage) {
    const PAGES = {
      staff:    { icon: Users,           title: 'จัดการพนักงาน', body: <StaffAdminView /> },
      report:   { icon: FileSpreadsheet, title: 'รายงาน',      body: <ReportView /> },
      settings: { icon: Cloud,           title: 'ตั้งค่าเซิร์ฟเวอร์',
                  body: <SettingsView config={supabaseConfig} onSave={saveSupabaseConfig} onTestConnection={testConnection} dataSource={dataSource} lastSyncAt={lastSyncAt} productCount={products.length} /> },
    };
    const P = PAGES[publicPage] || PAGES.report;
    const PIcon = P.icon;
    return (
      <div className="min-h-screen flex flex-col" style={{ background: '#F6F7F8' }}>
        <header className="bg-white border-b border-[#E4E6EA] px-4 py-3 sticky top-0 z-10 shadow-sm">
          <div className="max-w-6xl mx-auto flex items-center justify-between gap-2">
            <div className="flex items-center gap-2 min-w-0">
              <button onClick={() => setPublicPage(null)} title="กลับหน้าแรก"
                className="shrink-0 rounded-lg flex items-center justify-center border bg-white"
                style={{ width: 38, height: 38, borderColor: '#E4E6EA', color: '#0F172A' }}>
                <ArrowRight size={17} className="rotate-180" />
              </button>
              <div className="min-w-0">
                <h1 className="font-bold text-slate-800">{P.title}</h1>
              </div>
            </div>
            <div className="text-white p-2 rounded-lg shrink-0" style={{ background: '#0F172A' }}><PIcon size={18} /></div>
          </div>
        </header>
        <main className="flex-1 max-w-6xl w-full mx-auto p-4">{P.body}</main>
      </div>
    );
  }

  if (!currentUser) return <LoginScreen onLogin={handleLogin} onOpenPage={setPublicPage} serverReady={isSupabaseReady}
    productCount={products.length} onPrefillView={setView} onTestDb={testConnection}
    db={{ ...supabaseConfig, connection: connectionStatus, dataSource, lastSyncAt, productCount: products.length }} />;

  const isManager = currentUser.role === 'manager';
  const myBranch = String(currentUser.branch || '1');   // พนักงานทำงานสาขานี้เท่านั้น
  // ใบเก่าก่อนมี 2 สาขา ไม่มี branch → นับเป็นสาขา 1
  const inBranch = (x) => viewBranch === 'all' || String(x.branch || '1') === viewBranch;
  const bSubs = isManager ? submissions.filter(inBranch) : submissions;
  const bInvSubs = isManager ? invSubs.filter(inBranch) : invSubs;
  const feature = currentUser.feature || (isManager ? 'recorder' : 'recorder');
  const myEntries = countEntries.filter(e =>
    e.counterId === currentUser.id && (e.featureType || 'recorder') === (currentUser.feature || 'recorder'));
  const pendingCount = isManager
    ? submissions.filter(s => s.status === 'pending' && inBranch(s)).length + invSubs.filter(s => s.status === 'pending' && inBranch(s)).length
    : submissions.filter(s => s.status === 'pending' && (s.featureType||'recorder') === feature).length;
  // ใบที่อนุมัติแล้วแต่ยังไม่ขึ้น Drive
  const pendingUploadCount = isManager
    ? submissions.filter(s => s.status === 'approved' && !driveEntryOf(s, driveLog)?.ok).length
      + invSubs.filter(s => s.status === 'approved' && !driveEntryOf(s, driveLog)?.ok).length
    : 0;

  const FEATURE_LABEL = { recorder: 'นับสินค้า', stock_compare: 'นับเทียบยอด', invoice: 'บันทึกบิล' };
  // ระบบสีตาม Color System.dc.html — นับสินค้า เขียว · นับเทียบยอด ม่วง · บันทึกบิล ฟ้า
  // ผู้จัดการใช้สีของฟีเจอร์ที่กำลังดูอยู่ (ตามกติกา "หนึ่งหน้าจอเห็นสีฟีเจอร์ได้สีเดียว")
  const FEAT = {
    recorder:      { main: '#35706A', deep: '#2A5A55', soft: '#EAF1F0', line: '#B6D0CC' },
    stock_compare: { main: '#7C4DFF', deep: '#5B21B6', soft: '#F3EDFF', line: '#D6C6FF' },
    invoice:       { main: '#2F6E90', deep: '#255771', soft: '#EAF0F4', line: '#B9CFDC' },
  };
  // ผู้จัดการใช้สีกลาง (หมึก) — สีฟีเจอร์ไปอยู่บนป้ายของแต่ละใบ
  // ฟีเจอร์ที่พนักงานคนนี้มีสิทธิ์ — ใช้ในลิ้นชักสลับฟีเจอร์
  const ALL_FEATURES = [
    { id: 'recorder',      key: 'allow_recorder', label: 'นับสินค้า',   icon: ScanLine,       main: '#35706A', soft: '#EAF1F0', ink: '#2A5A55' },
    { id: 'stock_compare', key: 'allow_compare',  label: 'นับเทียบยอด', icon: ArrowLeftRight, main: '#7C4DFF', soft: '#F3EDFF', ink: '#5B21B6' },
    { id: 'invoice',       key: 'allow_invoice',  label: 'บันทึกบิล',   icon: Receipt,        main: '#2F6E90', soft: '#EAF0F4', ink: '#255771' },
  ];
  const myFeatures = ALL_FEATURES
    .filter(f => currentUser?.allow ? currentUser.allow[f.key] : f.id === currentUser?.feature)
    .map(f => ({ ...f, pending: countEntries.filter(e => e.counterId === currentUser?.id && (e.featureType || 'recorder') === f.id).length }));

  const C = isManager
    ? { main: '#0F172A', deep: '#0F172A', soft: '#F6F7F8', line: '#E4E6EA' }
    : (FEAT[feature] || FEAT.recorder);
  

  // Nav per role+feature
  const navItems = isManager
    ? [{ id:'inbox',label:'รีวิวและอนุมัติ',icon:Inbox,badge:pendingCount },{ id:'saved',label:'บันทึกแล้ว',icon:Upload,badge:pendingUploadCount },{ id:'compare',label:'เทียบยอด',icon:ArrowLeftRight },{ id:'data_sync',label:'สถานะข้อมูล POS',icon:Database },{ id:'drive_cfg',label:'ตั้งค่า',icon:SettingsIcon }]
    : feature === 'invoice'
      ? [{ id:'invoice',label:'บันทึกบิล',icon:Receipt }]
      : [{ id:'count',label:'นับสต็อก',icon:ScanLine },{ id:'review',label:'ตรวจสอบ',icon:ClipboardCheck,badge:myEntries.length },{ id:'my_submissions',label:'ที่ส่งแล้ว',icon:Send }];

  // หน้าที่จำไว้ไม่มีในเมนูของบทบาทนี้ (เช่นสลับฟีเจอร์) → ใช้หน้าแรกของเมนูแทน
  // ต้องเป็นค่าคำนวณ ไม่ใช่ hook เพราะอยู่หลังจุด return ของหน้าเลือกชื่อ
  // my_bills ไม่ได้อยู่ในเมนูล่าง (เปิดจากปุ่มบนหัว) แต่ต้องไม่ถูกเด้งกลับ
  const EXTRA_VIEWS = ['my_bills'];
  const activeView = (navItems.some(i => i.id === view) || EXTRA_VIEWS.includes(view))
    ? view : (navItems[0]?.id || view);

  return (
    <div className="min-h-screen flex flex-col" style={{ background: '#F6F7F8' }}>
      <header className="bg-white border-b border-[#E4E6EA] px-4 py-3 sticky top-0 z-10 shadow-sm">
        <div className="max-w-6xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-2 min-w-0">
            {isManager ? (
              <button onClick={() => setMenuOpen(true)} title="เมนู"
                className="shrink-0 rounded-lg flex items-center justify-center border bg-white relative"
                style={{ width: 40, height: 40, borderColor: '#E4E6EA', color: '#0F172A' }}>
                <span style={{ display:'flex', flexDirection:'column', gap:3 }}>
                  <span style={{ width:16, height:2, background:'#0F172A', borderRadius:2, display:'block' }}></span>
                  <span style={{ width:16, height:2, background:'#0F172A', borderRadius:2, display:'block' }}></span>
                  <span style={{ width:16, height:2, background:'#0F172A', borderRadius:2, display:'block' }}></span>
                </span>
                {pendingCount > 0 && (
                  <span className="absolute text-white text-[9px] font-bold rounded-full text-center"
                    style={{ top: -5, right: -5, minWidth: 17, padding: '1px 4px', background: '#B45309' }}>{pendingCount}</span>
                )}
              </button>
            ) : (
              <button onClick={() => setMenuOpen(true)} title="เมนู"
                className="shrink-0 rounded-lg flex items-center justify-center border bg-white"
                style={{ width: 40, height: 40, borderColor: '#E4E6EA' }}>
                <span style={{ display:'flex', flexDirection:'column', gap:3.5, alignItems:'stretch', width:17 }}>
                  <span style={{ height:2, background:C.main, borderRadius:2 }}></span>
                  <span style={{ height:2, background:C.main, borderRadius:2 }}></span>
                  <span style={{ height:2, background:C.main, borderRadius:2 }}></span>
                </span>
              </button>
            )}
            <div className="min-w-0">
              <h1 className="font-bold text-slate-800">
                {activeView === 'my_bills' ? 'บิลที่ส่งแล้ว' : (navItems.find(i => i.id === activeView)?.label || 'KUUHOO')}
              </h1>
              <p className="text-xs text-slate-500 truncate">
                {isManager ? 'ผู้จัดการ' : (FEATURE_LABEL[feature] || feature)}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {isManager ? (
              <div className="flex rounded-lg border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
                {[{ id: 'all', label: 'รวม' }, ...BRANCHES().map(b => ({ id: b.id, label: b.name }))].map((b, i) => {
                  const on = viewBranch === b.id;
                  return (
                    <button key={b.id} onClick={() => setViewBranch(b.id)}
                      className="font-bold" title={`ดู${b.label}`}
                      style={{ minHeight: 34, padding: '0 9px', fontSize: 11.5, borderLeft: i ? '1px solid #E4E6EA' : 'none',
                               background: on ? BRANCH_SOFT[b.id] : '#fff', color: on ? BRANCH_INK[b.id] : '#94A3B8' }}>
                      {b.label}
                    </button>
                  );
                })}
              </div>
            ) : (
              <span className="rounded-lg px-2.5 font-bold flex items-center shrink-0"
                style={{ minHeight: 32, fontSize: 11.5, background: BRANCH_SOFT[myBranch], color: BRANCH_INK[myBranch] }}>
                {branchName(myBranch)}
              </span>
            )}
            <button onClick={handleLogout} className="flex items-center gap-1 text-xs text-slate-500 hover:text-slate-700 px-2 py-1 hover:bg-[#F6F7F8] rounded"><LogOut size={14} />ออก</button>
          </div>
        </div>
      </header>


      {menuOpen && (
        <div className="fixed inset-0 z-50 flex" onClick={() => setMenuOpen(false)}>
          <div className="bg-white h-full flex flex-col" style={{ width: 268, maxWidth: '84vw', boxShadow: '2px 0 16px rgba(15,23,42,.18)' }}
            onClick={e => e.stopPropagation()}>
            <div className="px-4 py-4 border-b flex items-center gap-3" style={{ borderColor: '#E4E6EA', background: '#F6F7F8' }}>
              <span className="rounded-full flex items-center justify-center text-[15px] font-bold text-white shrink-0"
                style={{ width: 42, height: 42, background: '#0F172A' }}>{(currentUser.name || '?').charAt(0)}</span>
              <div className="min-w-0">
                <div className="text-[14px] font-bold text-slate-800 truncate">{currentUser.name}</div>
                <div className="text-[11px] text-slate-500">{isManager ? 'ผู้จัดการ' : 'พนักงาน'}</div>
              </div>
            </div>

            <div className="flex-1 overflow-y-auto py-2">
              {!isManager && myFeatures.length > 1 && (
                <>
                  <div className="px-4 py-2 text-[10px] font-bold tracking-wide text-slate-400">ฟีเจอร์</div>
                  {myFeatures.map(f => {
                    const on = feature === f.id;
                    const FIcon = f.icon;
                    return (
                      <button key={f.id} onClick={() => switchFeature(f.id)}
                        className="w-full flex items-center gap-3 px-4 text-left"
                        style={{ minHeight: 52, background: on ? '#F6F7F8' : '#fff',
                                 borderLeft: on ? '3px solid #0F172A' : '3px solid transparent' }}>
                        <FIcon size={18} style={{ color: on ? f.main : '#64748B' }} className="shrink-0" />
                        <span className="flex-1 text-[14px] font-semibold" style={{ color: on ? '#0F172A' : '#334155' }}>{f.label}</span>
                        {f.pending > 0 && (
                          <span className="text-[10px] font-bold rounded-full text-center shrink-0"
                            style={{ minWidth: 20, padding: '2px 6px', background: f.soft, color: f.ink }}>{f.pending}</span>
                        )}
                      </button>
                    );
                  })}
                </>
              )}
              {isManager && (
                <>
                  <div className="px-4 py-2 mt-1 text-[10px] font-bold tracking-wide text-slate-400">งานประจำวัน</div>
                  {navItems.map(item => {
                const Icon = item.icon; const on = activeView === item.id;
                return (
                  <button key={item.id} onClick={() => { setView(item.id); setMenuOpen(false); }}
                    className="w-full flex items-center gap-3 px-4 text-left"
                    style={{ minHeight: 52, background: on ? '#F6F7F8' : '#fff',
                             borderLeft: on ? '3px solid #0F172A' : '3px solid transparent' }}>
                    <Icon size={18} style={{ color: on ? '#0F172A' : '#64748B' }} className="shrink-0" />
                    <span className="flex-1 text-[14px] font-semibold" style={{ color: on ? '#0F172A' : '#334155' }}>{item.label}</span>
                    {item.badge > 0 && (
                      <span className="text-white text-[10px] font-bold rounded-full text-center shrink-0"
                        style={{ minWidth: 20, padding: '2px 6px', background: '#B45309' }}>{item.badge}</span>
                    )}
                  </button>
                );
                  })}
                </>
              )}
            </div>

            <div className="border-t p-3" style={{ borderColor: '#E4E6EA' }}>
              <button onClick={() => { setMenuOpen(false); handleLogout(); }}
                className="w-full flex items-center justify-center gap-2 rounded-xl font-bold text-[13.5px] border"
                style={{ minHeight: 46, borderColor: '#E4E6EA', color: '#475569', background: '#fff' }}>
                <LogOut size={15} />ออกจากระบบ
              </button>
            </div>
          </div>
          <div className="flex-1" style={{ background: 'rgba(15,23,42,.45)' }}></div>
        </div>
      )}

      <main className="flex-1 max-w-6xl w-full mx-auto p-4" style={{ paddingBottom: isManager ? 24 : 'calc(76px + env(safe-area-inset-bottom))' }}>
        {/* Counter - stock / stock_compare */}
        {!isManager && feature !== 'invoice' && activeView === 'count' && <CounterCountView entries={myEntries} addEntry={addCountEntry} deleteEntry={deleteCountEntry} checkBarcode={checkBarcode} setView={setView} products={products} isSupabaseReady={isSupabaseReady} connectionStatus={connectionStatus} countDate={countDate} setCountDate={setCountDate} draft={countDraft} updateDraft={updateDraft} pushDraft={pushDraft} pullDraft={pullDraft} draftSync={draftSync} tone={C} revise={reviseCount} onCancelRevise={() => { clearMyEntries(); clearReviseCount(); }} />}
        {!isManager && feature !== 'invoice' && activeView === 'review' && <CounterReviewView tone={C} entries={myEntries} setView={setView} submitForReview={submitForReview} clearMyEntries={clearMyEntries} currentUser={currentUser} pickedAt={pickedAt} revise={reviseCount} />}
        {!isManager && feature !== 'invoice' && activeView === 'my_submissions' && <MySubmissionsView onRefresh={() => pullSubmissions(feature)} onResend={resendSubmission} subSync={subSync} submissions={submissions.filter(s => s.counterId === currentUser.id && (s.featureType||'recorder') === feature)} setView={setView} onRevise={startReviseCount} />}
        {/* Counter - invoice */}
        {!isManager && feature === 'invoice' && <ErrorBox><InvoiceScannerModule supabaseConfig={supabaseConfig} currentUser={currentUser}
            onOpenSent={() => setView('my_bills')} onCloseSent={() => setView('invoice')}
            sentActive={activeView === 'my_bills'}
            sentBadge={(invSubs||[]).filter(x=>x.status==='rejected').length}
            reviseTarget={reviseTarget} onClearRevise={() => setReviseTarget(null)}
            sentView={<MyBillsView invSubs={invSubs} setView={setView} onRefresh={() => pullInvSubs(currentUser)}
              onRevise={(t) => { setReviseTarget(t); setView('invoice'); }} />} /></ErrorBox>}
        {/* Manager */}
                {isManager && activeView === 'compare' && <CompareStockView branch={viewBranch} submissions={bSubs} supabaseConfig={supabaseConfig} compareState={compareState} setCompareState={setCompareState} />}
        {isManager && activeView === 'data_sync' && <ErrorBox><DataSyncView /></ErrorBox>}
        {isManager && activeView === 'drive_cfg' && <ErrorBox><DriveSettingsView currentUser={currentUser} /></ErrorBox>}
        {isManager && activeView === 'saved' && <ErrorBox><SavedUploadsView submissions={bSubs} invSubs={bInvSubs} subSync={subSync} currentUser={currentUser} supabaseConfig={supabaseConfig}
            onRefresh={() => { pullSubmissions('recorder'); pullSubmissions('stock_compare'); pullInvSubs(); }}
            onPatched={u => setSubmissions(prev => prev.map(x => x.id === u.id ? { ...x, ...u } : x))}
            onInvPatched={u => setInvSubs(prev => prev.map(x => x.id === u.id ? { ...x, ...u } : x))} /></ErrorBox>}
        {isManager && activeView === 'inbox' && <ManagerInboxView branch={viewBranch} invSubs={bInvSubs} onReviewInvoice={reviewInvoice} onDeleteInvoice={deleteInvoiceSub} onRefresh={() => { pullSubmissions('recorder'); pullSubmissions('stock_compare'); pullInvSubs(); }} subSync={subSync} submissions={bSubs} onReview={reviewSubmission} onDelete={deleteSubmission} feature={feature} />}
        {isManager && feature === 'stock_compare' && activeView === 'compare' && <CompareStockView branch={viewBranch} submissions={bSubs.filter(s=>(s.featureType||'stock_compare')==='stock_compare')} supabaseConfig={supabaseConfig} compareState={compareState} setCompareState={setCompareState} />}
      </main>

      {!isManager && navItems.length > 1 && (
      <nav className="fixed bottom-0 left-0 right-0 bg-white border-t border-[#E4E6EA]" style={{ paddingBottom: 'env(safe-area-inset-bottom)' }}>
        <div className="max-w-6xl mx-auto grid" style={{ gridTemplateColumns: `repeat(${navItems.length}, minmax(0, 1fr))` }}>
          {navItems.map(item => {
            const Icon = item.icon; const active = activeView === item.id;
            const tone = NAV_TONE[item.id] || { main: C.main, soft: C.soft };
            return (
              <button key={item.id} onClick={() => setView(item.id)}
                className="relative flex flex-col items-center justify-center gap-1 text-[10px] leading-none"
                style={{ minHeight: 56, paddingTop: 8, paddingBottom: 8,
                         borderTop: active ? `3px solid ${tone.main}` : '3px solid transparent',
                         ...(active ? { color: tone.main, background: tone.soft } : { color: '#64748B', background: '#fff' }) }}>
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
      )}
    </div>
  );
}

function LoginScreen({ onLogin, onOpenPage, serverReady, productCount = 0, db, onTestDb, onPrefillView }) {
  const [role, setRole] = useState(null);
  const [branch, setBranch] = useState(null);   // เลือกทุกครั้งที่เปิด — เครื่องรวมอาจย้ายสาขา
  const [feature, setFeature] = useState(null);
  const [name, setName] = useState('');
  const [pin, setPin] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [staffList, setStaffList] = useState(undefined);   // undefined = ยังไม่โหลด · null = โหลดไม่ได้
  const [staffLoading, setStaffLoading] = useState(false);
  const [staffErr, setStaffErr] = useState('');            // เหตุผลจริงที่โหลดไม่ได้

  // โทนสีของแต่ละฟีเจอร์ ใช้กับป้ายบอกฟีเจอร์ที่เลือก
  const FEAT_TONE = {
    recorder: { soft: '#EAF1F0', ink: '#2A5A55' },
    compare:  { soft: '#F3EDFF', ink: '#5B21B6' },
    invoice:  { soft: '#EAF0F4', ink: '#255771' },
    indigo:   { soft: '#F6F7F8', ink: '#0F172A' },
  };

  // เลือกชื่อได้ทันทีจากหน้าแรก — ไม่ต้องเลือกบทบาท/ฟีเจอร์ก่อน
  useEffect(() => {
    if (role === 'manager') { setStaffList(undefined); return; }
    let alive = true;
    setStaffLoading(true); setStaffErr('');
    fetch('/api/staff')
      .then(async r => {
        const text = await r.text();
        let j; try { j = JSON.parse(text); } catch { throw new Error('เซิร์ฟเวอร์ตอบไม่ใช่ JSON (' + r.status + '): ' + text.slice(0, 120)); }
        if (!r.ok || j.error) throw new Error(j.error || ('HTTP ' + r.status));
        return j;
      })
      .then(j => {
        if (!alive) return;
        const list = Array.isArray(j.staff) ? j.staff : null;
        // เฉพาะคนที่เปิดใช้งานและมีสิทธิ์อย่างน้อยหนึ่งฟีเจอร์
        setStaffList(list ? list.filter(s => s.active !== false && (s.allow_recorder || s.allow_compare || s.allow_invoice)) : null);
      })
      .catch(e => { if (alive) { setStaffList(null); setStaffErr(e.message || 'ไม่ทราบสาเหตุ'); } })
      .finally(() => { if (alive) setStaffLoading(false); });
    return () => { alive = false; };
  }, [role]);

  const COUNTER_FEATURES = [
    { id: 'recorder',      label: 'นับสินค้า',            desc: 'นับแล้วส่งให้ผู้จัดการตรวจ',              icon: ScanLine,      accent: 'recorder' },
    { id: 'stock_compare', label: 'นับเทียบยอด',        desc: 'นับแล้วเทียบกับยอดในระบบ',       icon: ArrowLeftRight, accent: 'compare'    },
    { id: 'invoice',       label: 'บันทึกบิล',            desc: 'ถ่ายบิลซื้อ ไม่ต้องพิมพ์เอง',            icon: Receipt,       accent: 'invoice'  },
  ];
  const MANAGER_FEATURES = [
    { id: 'recorder',      label: 'นับสินค้า',            desc: 'ตรวจและอนุมัติใบนับ',        icon: ScanLine,      accent: 'indigo'  },
    { id: 'stock_compare', label: 'นับเทียบยอด',        desc: 'ตรวจใบนับและดูส่วนต่าง',        icon: ArrowLeftRight, accent: 'indigo'  },
  ];

  const accentClass = (a, type) => {
    const map = {
      recorder: { border: 'hover:border-[#35706A] hover:bg-[#EAF1F0]', icon: 'bg-[#EAF1F0] text-[#2A5A55]', btn: 'bg-[#35706A] hover:bg-[#2A5A55]' },
      compare:  { border: 'hover:border-[#7C4DFF] hover:bg-[#F3EDFF]', icon: 'bg-[#F3EDFF] text-[#5B21B6]', btn: 'bg-[#7C4DFF] hover:bg-[#5B21B6]' },
      invoice:  { border: 'hover:border-[#2F6E90] hover:bg-[#EAF0F4]', icon: 'bg-[#EAF0F4] text-[#255771]', btn: 'bg-[#2F6E90] hover:bg-[#255771]' },
      indigo:   { border: 'hover:border-[#0F172A] hover:bg-[#F6F7F8]', icon: 'bg-[#F6F7F8] text-[#0F172A]', btn: 'bg-[#0F172A] hover:bg-[#0F172A]' },
    };
    return map[a]?.[type] || map.indigo[type];
  };

  // ผู้จัดการไม่ต้องเลือกฟีเจอร์ — งานคือรีวิวทุกฟีเจอร์ในหน้าเดียว
  useEffect(() => { if (role === 'manager' && !feature) setFeature('all'); }, [role, feature]);

  const handleLogin = async (picked) => {
    const useName = picked?.name || name.trim();
    if (!useName) return setError('กรุณาเลือกหรือใส่ชื่อ');
    if (role === 'counter' && !branch) return setError('กรุณาเลือกสาขา');
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
    // เลือกจากรายชื่อ → ใช้ id ของ staff เพื่อให้ร่างเดินตามคนข้ามเครื่อง
    const allow = picked ? {
      allow_recorder: !!picked.allow_recorder,
      allow_compare:  !!picked.allow_compare,
      allow_invoice:  !!picked.allow_invoice,
    } : { allow_recorder: true, allow_compare: false, allow_invoice: false };
    const firstFeature = role === 'manager' ? 'all'
      : allow.allow_recorder ? 'recorder' : allow.allow_compare ? 'stock_compare' : allow.allow_invoice ? 'invoice' : 'recorder';
    const useFeature = role === 'manager' ? 'all' : firstFeature;
    const stableId = picked?.id || `${role}_${useName.toLowerCase().replace(/\s+/g,'_')}`;
    // ผู้จัดการเห็นทั้ง 2 สาขา สลับดูได้จากหัวหน้าจอ
    onLogin({ id: stableId, name: useName, role, feature: useFeature, allow,
      branch: role === 'manager' ? 'all' : branch, loginAt: new Date().toISOString() });
  };

  const features = role === 'manager' ? MANAGER_FEATURES : COUNTER_FEATURES;
  const selectedFeature = features.find(f => f.id === feature);

  return (
    <div className="min-h-screen flex items-center justify-center p-4" style={{ background: '#F6F7F8' }}>
      {!role ? (
      <div className="w-full" style={{ maxWidth: 460 }}>
        <div className="flex items-center gap-3 mb-4">
          <div className="text-white rounded-2xl flex items-center justify-center shrink-0"
            style={{ width: 52, height: 52, background: '#0F172A' }}><Package size={26} /></div>
          <div className="min-w-0">
            <h1 className="text-[26px] font-bold text-slate-800 leading-none">KUUHOO</h1>
            <p className="text-[12.5px] text-slate-500 mt-1">ระบบนับสต็อกและคีย์บิลหน้าร้าน</p>
          </div>
        </div>

        {!serverReady && (
          <div className="w-full flex items-center gap-2.5 px-3 py-2.5 rounded-xl border mb-4"
            style={{ background: '#FFFBEB', borderColor: '#FDE68A' }}>
            <span className="rounded-full shrink-0" style={{ width: 8, height: 8, background: '#B45309' }} />
            <div className="flex-1 min-w-0">
              <div className="text-[12.5px] font-bold" style={{ color: '#B45309' }}>ยังไม่ได้ตั้งค่าเซิร์ฟเวอร์</div>
              <div className="text-[11px] mt-0.5 truncate" style={{ color: '#B45309' }}>ต้องใส่ URL และ Anon Key ก่อนเริ่มนับ</div>
            </div>
          </div>
        )}

        <div className="text-[11px] font-bold tracking-wide text-slate-400 mb-2">เริ่มทำงาน</div>
        <div className="grid grid-cols-2 gap-2.5">
          <button onClick={() => setRole('counter')}
            className="bg-white rounded-2xl border-2 p-4 text-left flex flex-col gap-2.5 hover:border-[#35706A]"
            style={{ borderColor: '#E4E6EA', minHeight: 132 }}>
            <div className="rounded-xl flex items-center justify-center shrink-0"
              style={{ width: 42, height: 42, background: '#EAF1F0', color: '#2A5A55' }}><User size={21} /></div>
            <div className="mt-auto">
              <div className="text-[15px] font-bold text-slate-800">พนักงาน</div>
              <div className="text-[11.5px] text-slate-500 mt-0.5 leading-snug">เลือกชื่อแล้วเริ่มงาน</div>
            </div>
          </button>
          <button onClick={() => { setRole('manager'); setFeature('all'); }}
            className="bg-white rounded-2xl border-2 p-4 text-left flex flex-col gap-2.5 hover:border-[#0F172A]"
            style={{ borderColor: '#E4E6EA', minHeight: 132 }}>
            <div className="rounded-xl flex items-center justify-center shrink-0"
              style={{ width: 42, height: 42, background: '#F6F7F8', color: '#0F172A' }}><Shield size={21} /></div>
            <div className="mt-auto">
              <div className="text-[15px] font-bold text-slate-800">ผู้จัดการ</div>
              <div className="text-[11.5px] text-slate-500 mt-0.5 leading-snug">รีวิว อนุมัติ เทียบยอด</div>
            </div>
          </button>
        </div>

        {error && (
          <div className="mt-2.5 rounded-xl px-3 py-2.5 text-[12px] border" style={{ background: '#FEF2F2', borderColor: '#FECACA', color: '#B91C1C' }}>{error}</div>
        )}

        <div className="mt-5">
          <LandingStatus
            onOpenManager={() => { onPrefillView?.('saved'); setRole('manager'); setFeature('all'); }}
            onOpenSync={() => { onPrefillView?.('data_sync'); setRole('manager'); setFeature('all'); }}
            db={db} onTestDb={onTestDb}
            onOpenDbSettings={() => onOpenPage('settings')} />
        </div>

        <div className="text-[11px] font-bold tracking-wide text-slate-400 mt-5 mb-2">เครื่องมือ · เปิดได้เลย</div>
        <div className="bg-white rounded-2xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
          {[
            { id: 'staff',     icon: Users,           title: 'จัดการพนักงาน',    sub: 'เพิ่มคน ตั้งสิทธิ์ ปิดการใช้งาน' },
            { id: 'report',    icon: FileSpreadsheet, title: 'รายงาน',           sub: 'ดึงข้อมูลและส่งออก Excel / CSV' },
          ].map((p, i) => {
            const PIcon = p.icon;
            return (
              <button key={p.id} onClick={() => onOpenPage(p.id)}
                className="w-full flex items-center gap-3 px-3.5 text-left hover:bg-[#F6F7F8]"
                style={{ minHeight: 62, borderTop: i ? '1px solid #F1F3F5' : 'none' }}>
                <div className="rounded-lg flex items-center justify-center shrink-0"
                  style={{ width: 36, height: 36, background: '#F6F7F8', color: '#0F172A' }}><PIcon size={17} /></div>
                <div className="flex-1 min-w-0">
                  <div className="text-[13.5px] font-bold text-slate-800">{p.title}</div>
                  <div className="text-[11px] text-slate-500 truncate">{p.sub}</div>
                </div>
                <ArrowRight size={15} className="text-slate-300 shrink-0" />
              </button>
            );
          })}
        </div>
      </div>
      ) : (
      <div className="bg-white rounded-2xl shadow-xl w-full max-w-md p-6 space-y-5">
        {/* Step 2: Feature */}
        {/* พนักงาน — เลือกสาขาก่อน กันคีย์ผิดสาขา */}
        {role === 'counter' && !branch && (
          <div className="space-y-3">
            <div className="flex items-center gap-2">
              <button onClick={() => { setRole(null); setError(''); }} className="text-xs text-slate-400 hover:text-slate-600">← กลับ</button>
              <div className="text-sm font-semibold text-slate-700">วันนี้ทำงานสาขาไหน</div>
            </div>
            <div className="grid grid-cols-2 gap-2.5">
              {BRANCHES().map(b => (
                <button key={b.id} onClick={() => { setBranch(b.id); setError(''); }}
                  className="rounded-2xl border-2 p-4 text-left flex flex-col gap-2 bg-white"
                  style={{ borderColor: '#E4E6EA', minHeight: 112 }}>
                  <span className="rounded-xl flex items-center justify-center font-bold shrink-0"
                    style={{ width: 42, height: 42, fontSize: 19, background: BRANCH_SOFT[b.id], color: BRANCH_INK[b.id] }}>{b.id}</span>
                  <span className="text-[15px] font-bold text-slate-800 mt-auto leading-tight">{b.name}</span>
                </button>
              ))}
            </div>
            <div className="text-[11px] text-slate-400 leading-relaxed">
              เลือกสาขาทุกครั้งที่เปิดแอป — ยอดคงเหลือและเลขเอกสารแยกกันคนละสาขา
            </div>
          </div>
        )}

        {/* พนักงาน — เลือกชื่อ */}
        {role === 'counter' && branch && (
          <div className="space-y-3">
            <div className="flex items-center gap-2">
              <button onClick={() => { setBranch(null); setError(''); setName(''); }} className="text-xs text-slate-400 hover:text-slate-600">← กลับ</button>
              <div className="text-sm font-semibold text-slate-700">เลือกชื่อของคุณ</div>
              <span className="ml-auto rounded-lg px-2 py-1 text-[11px] font-bold"
                style={{ background: BRANCH_SOFT[branch], color: BRANCH_INK[branch] }}>{branchName(branch)}</span>
            </div>

            {staffLoading || staffList === undefined ? (
              <div className="py-8 text-center text-[13px] text-slate-400 flex items-center justify-center gap-2">
                <RefreshCw size={15} className="animate-spin" />กำลังโหลดรายชื่อ…
              </div>
            ) : staffList === null ? (
              <div className="space-y-2.5">
                <div className="rounded-xl p-3 border text-[12px] leading-relaxed" style={{ background: '#FFFBEB', borderColor: '#FDE68A', color: '#B45309' }}>
                  <div className="font-bold">โหลดรายชื่อไม่ได้ — พิมพ์ชื่อเองไปก่อนได้</div>
                  {staffErr && <div className="mt-1.5 font-mono text-[10.5px] break-all" style={{ opacity: .85 }}>{staffErr}</div>}
                </div>
                <input type="text" value={name} onChange={e => { setName(e.target.value); setError(''); }} placeholder="ชื่อของคุณ"
                  className="w-full px-3 py-3 border rounded-xl outline-none focus:border-[#35706A] text-[15px]" style={{ borderColor: '#E2E8F0' }} />
                <button onClick={() => handleLogin()} disabled={!name.trim()}
                  className="w-full text-white font-bold text-[15px] rounded-xl disabled:opacity-40"
                  style={{ minHeight: 52, background: '#35706A' }}>เริ่มงาน</button>
              </div>
            ) : staffList.length === 0 ? (
              <div className="px-4 py-8 text-center">
                <User className="mx-auto text-[#E4E6EA] mb-2" size={32} />
                <div className="text-[13px] text-slate-500">ยังไม่มีพนักงานในรายชื่อ</div>
                <button onClick={() => onOpenPage('staff')}
                  className="mt-3 text-white px-4 rounded-xl text-[13px] font-bold" style={{ minHeight: 44, background: '#0F172A' }}>ไปเพิ่มพนักงาน</button>
              </div>
            ) : (
              <div className="rounded-xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
                {staffList.map((s, i) => {
                  const tags = [
                    s.allow_recorder && { label: 'นับสินค้า', soft: '#EAF1F0', ink: '#2A5A55' },
                    s.allow_compare && { label: 'นับเทียบยอด', soft: '#F3EDFF', ink: '#5B21B6' },
                    s.allow_invoice && { label: 'บันทึกบิล', soft: '#EAF0F4', ink: '#255771' },
                  ].filter(Boolean);
                  return (
                    <button key={s.id} onClick={() => handleLogin(s)}
                      className="w-full flex items-center gap-3 px-3.5 text-left hover:bg-[#F6F7F8]"
                      style={{ minHeight: 68, borderTop: i ? '1px solid #F1F3F5' : 'none' }}>
                      <span className="rounded-full flex items-center justify-center text-[16px] font-bold text-white shrink-0"
                        style={{ width: 44, height: 44, background: '#35706A' }}>
                        {s.initial || (s.name || '?').charAt(0)}
                      </span>
                      <div className="flex-1 min-w-0">
                        <div className="text-[15px] font-bold text-slate-800 truncate">{s.name}</div>
                        <div className="flex items-center gap-1 mt-1 flex-wrap">
                          {tags.map(t => (
                            <span key={t.label} className="text-[9.5px] font-semibold px-1.5 py-0.5 rounded-full"
                              style={{ background: t.soft, color: t.ink }}>{t.label}</span>
                          ))}
                        </div>
                      </div>
                      <ArrowRight size={15} className="text-slate-300 shrink-0" />
                    </button>
                  );
                })}
              </div>
            )}

            {error && <div className="rounded-lg px-3 py-2 text-[12px] border" style={{ background: '#FEF2F2', borderColor: '#FECACA', color: '#B91C1C' }}>{error}</div>}
          </div>
        )}

        {role === 'manager' && (
          <div className="space-y-3">
            <div className="flex items-center gap-2">
              <button onClick={() => { setError(''); setPin(''); setName(''); if (role === 'manager') { setRole(null); setFeature(null); } else setFeature(null); }} className="text-xs text-slate-400 hover:text-slate-600">← กลับ</button>
              <div className="text-sm font-semibold text-slate-700">{role === 'manager' ? 'ข้อมูลผู้ใช้' : 'เลือกชื่อของคุณ'}</div>
            </div>

            <div className="flex gap-2 flex-wrap">
              <div className="rounded-lg px-3 py-2 flex items-center gap-1.5 text-xs font-medium"
                style={role === 'manager' ? { background: '#F6F7F8', color: '#0F172A' } : { background: '#EAF1F0', color: '#2A5A55' }}>
                {role === 'manager' ? <Shield size={12} /> : <User size={12} />}
                {role === 'manager' ? 'ผู้จัดการ' : 'พนักงาน'}
              </div>
              {selectedFeature && (
                <div className="rounded-lg px-3 py-2 flex items-center gap-1.5 text-xs font-medium"
                  style={{ background: FEAT_TONE[selectedFeature.accent]?.soft || '#F6F7F8', color: FEAT_TONE[selectedFeature.accent]?.ink || '#0F172A' }}>
                  <selectedFeature.icon size={12} />{selectedFeature.label}
                </div>
              )}
            </div>

            {role === 'manager' ? (
              <>
                <div>
                  <label className="text-sm font-medium text-slate-700 mb-1 block">ชื่อของคุณ</label>
                  <input type="text" value={name} onChange={e => { setName(e.target.value); setError(''); }} placeholder="เช่น สมหญิง"
                    className="w-full px-3 py-2.5 border border-[#E4E6EA] rounded-lg focus:ring-2 focus:ring-[#0F172A] outline-none" autoFocus />
                </div>
                <div>
                  <label className="text-sm font-medium text-slate-700 mb-1 block"><Lock size={12} className="inline mr-1" />รหัส PIN</label>
                  <input type="password" value={pin} onChange={e => { setPin(e.target.value); setError(''); }} placeholder="••••" maxLength={4}
                    className="w-full px-3 py-2.5 border border-[#E4E6EA] rounded-lg focus:ring-2 focus:ring-[#0F172A] outline-none font-mono text-lg tracking-widest text-center" />
                </div>
                {error && <div className="bg-[#FEF2F2] border border-[#FECACA] text-[#B91C1C] text-sm rounded-lg p-2">{error}</div>}
                <button onClick={() => handleLogin()} disabled={loading}
                  className="w-full py-3 rounded-xl font-bold text-white disabled:opacity-60" style={{ background: '#0F172A' }}>
                  {loading ? 'กำลังตรวจสอบ...' : 'เข้าสู่ระบบ'}
                </button>
              </>
            ) : (
              <>
                {staffLoading || staffList === undefined ? (
                  <div className="py-8 text-center text-[13px] text-slate-400 flex items-center justify-center gap-2">
                    <RefreshCw size={15} className="animate-spin" />กำลังโหลดรายชื่อ…
                  </div>
                ) : staffList === null ? (
                  <div className="rounded-xl p-3 border text-[12px] leading-relaxed" style={{ background: '#FFFBEB', borderColor: '#FDE68A', color: '#B45309' }}>
                    <div className="font-bold">โหลดรายชื่อไม่ได้ — พิมพ์ชื่อเองไปก่อนได้</div>
                    {staffErr && <div className="mt-1.5 font-mono text-[10.5px] break-all" style={{ opacity: .85 }}>{staffErr}</div>}
                  </div>
                ) : staffList.length === 0 ? (
                  <div className="rounded-xl p-3 border text-[12px] leading-relaxed" style={{ background: '#FFFBEB', borderColor: '#FDE68A', color: '#B45309' }}>
                    ยังไม่มีชื่อที่เปิดสิทธิ์ฟีเจอร์นี้ — ไปเพิ่มที่ “จัดการพนักงาน” หน้าแรก
                  </div>
                ) : (
                  <div className="rounded-xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
                    {staffList.map((s, i) => (
                      <button key={s.id} onClick={() => handleLogin(s)}
                        className="w-full flex items-center gap-3 px-3 text-left hover:bg-[#F6F7F8]"
                        style={{ minHeight: 58, borderTop: i ? '1px solid #F1F3F5' : 'none' }}>
                        <span className="rounded-full flex items-center justify-center text-[15px] font-bold text-white shrink-0"
                          style={{ width: 38, height: 38, background: '#35706A' }}>
                          {s.initial || (s.name || '?').charAt(0)}
                        </span>
                        <div className="flex-1 min-w-0">
                          <div className="text-[14.5px] font-bold text-slate-800 truncate">{s.name}</div>
                          {s.dept && <div className="text-[11px] text-slate-500 truncate">{s.dept}</div>}
                        </div>
                        <ArrowRight size={15} className="text-slate-300 shrink-0" />
                      </button>
                    ))}
                  </div>
                )}

                {(staffList === null || staffList?.length === 0) && (
                  <>
                    <div>
                      <label className="text-sm font-medium text-slate-700 mb-1 block">พิมพ์ชื่อเอง</label>
                      <input type="text" value={name} onChange={e => { setName(e.target.value); setError(''); }} placeholder="เช่น สมหญิง"
                        className="w-full px-3 py-2.5 border border-[#E4E6EA] rounded-lg focus:ring-2 focus:ring-[#35706A] outline-none" />
                    </div>
                    <button onClick={() => handleLogin()} disabled={!name.trim()}
                      className="w-full py-3 rounded-xl font-bold text-white disabled:opacity-40" style={{ background: '#35706A' }}>
                      เริ่มทำงาน
                    </button>
                  </>
                )}

                {error && <div className="bg-[#FEF2F2] border border-[#FECACA] text-[#B91C1C] text-sm rounded-lg p-2">{error}</div>}
              </>
            )}
          </div>
        )}
      </div>
      )}
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

function CounterCountView({ revise, onCancelRevise, entries, addEntry, deleteEntry, checkBarcode, setView, products, isSupabaseReady, connectionStatus, countDate, setCountDate, draft, updateDraft, pushDraft, pullDraft, draftSync = {}, tone }) {
  const T = tone || { main: '#35706A', deep: '#2A5A55', soft: '#EAF1F0', line: '#B6D0CC' };
  const [location, setLocation] = useState('');
  const [manualPrice, setManualPrice] = useState('');   // ราคาขายของรายการที่ไม่พบในระบบ
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
    addEntry({ barcode: checkResult.barcode, productName: checkResult.name, productId: checkResult.id,
      masterCode: checkResult.masterCode || checkResult.id, stockRatio: checkResult.stockRatio || 1, isPack: !!checkResult.isPack,
      unit: checkResult.unit, price: checkResult.price || 0, cost: checkResult.cost || 0, qty: parseInt(qty), countDate, location: location.trim() });
    updateDraft({ barcode: '', qty: '', checkResult: null, error: '' });
    setTimeout(() => barcodeInputRef.current?.focus(), 100);
  };

  const handleAddManually = () => {
    if (!barcode.trim() || !qty || parseInt(qty) <= 0) return;
    addEntry({ barcode: barcode.trim(), productName: '(ไม่พบในระบบ)', productId: '', unit: '', qty: parseInt(qty),
               price: parseFloat(manualPrice) || 0, cost: 0, countDate, notFound: true, location: location.trim() });
    setManualPrice('');
    updateDraft({ barcode: '', qty: '', checkResult: null, error: '' });
    setTimeout(() => barcodeInputRef.current?.focus(), 100);
  };

  const totalQty = entries.reduce((s, e) => s + e.qty, 0);
  const uniqueBarcodes = new Set(entries.map(e => e.barcode)).size;
  const recent = [...entries].reverse();

  return (
    <div className="space-y-3">
      {revise && (
        <div className="rounded-xl border px-3 py-2.5 flex items-center gap-2.5" style={{ background: '#FEF2F2', borderColor: '#FECACA' }}>
          <div className="flex-1 min-w-0">
            <div className="text-[12px] font-bold" style={{ color: '#B91C1C' }}>
              กำลังแก้ {revise.docNo} → {String(revise.docNo || '').replace(/R\d+$/i, '')}R{(revise.reviseNo || 0) + 1}
            </div>
            {revise.note && <div className="text-[11px] mt-0.5 truncate" style={{ color: '#B91C1C' }}>{revise.note}</div>}
          </div>
          <button onClick={onCancelRevise}
            className="shrink-0 rounded-lg font-bold text-[11.5px] border"
            style={{ minHeight: 32, padding: '0 10px', background: '#fff', borderColor: '#FECACA', color: '#B91C1C' }}>ยกเลิก</button>
        </div>
      )}
      {/* รอบนับ + สถานะเซิร์ฟเวอร์ */}
      <div className="bg-white border rounded-xl overflow-hidden" style={{ borderColor: '#e4e6ea' }}>
        <div className="flex items-center gap-2 px-3 py-2.5 border-b" style={{ borderColor: '#eef0f3' }}>
          <div className="flex-1 min-w-0">
            <div className="text-[10px] font-bold tracking-wide" style={{ color: T.main }}>รอบนับ</div>
            <input type="date" value={countDate} onChange={e => setCountDate(e.target.value)}
              className="mt-0.5 text-sm font-bold text-slate-800 bg-transparent outline-none" style={{ fontFamily: "'IBM Plex Mono', monospace" }} />
          </div>
          {(!isSupabaseReady || connectionStatus === 'error') && (
            <div className={`shrink-0 text-[10px] font-semibold px-2 py-1 rounded-md ${!isSupabaseReady ? 'bg-[#FEF2F2] text-[#B91C1C]' : 'bg-[#FFFBEB] text-[#B45309]'}`}>
              {!isSupabaseReady ? 'ยังไม่ตั้งค่าเซิร์ฟเวอร์' : 'ต่อเซิร์ฟเวอร์ไม่ได้'}
            </div>
          )}
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
        style={{ minHeight: 54, background: T.main, boxShadow: `0 2px 0 ${T.deep}` }}>
        <Camera size={19} />{entries.length ? 'สแกนชิ้นต่อไป' : 'สแกนบาร์โค้ด'}
      </button>

      {/* พิมพ์เอง */}
      <div className="bg-white border rounded-xl p-3 space-y-2.5" style={{ borderColor: '#e2e8f0' }}>
        <div className="flex gap-2">
          <input ref={barcodeInputRef} type="text" value={barcode}
            onChange={e => { setBarcode(e.target.value); setCheckResult(null); setError(''); }}
            onKeyDown={e => e.key === 'Enter' && handleCheck()}
            placeholder="8850999320014"
            className="flex-1 px-3 py-2.5 border rounded-lg outline-none focus:border-[#7C4DFF] text-[15px]"
            style={{ borderColor: '#e2e8f0', fontFamily: "'IBM Plex Mono', monospace" }} />
          <button onClick={() => handleCheck()} disabled={!barcode.trim() || checking}
            className="px-4 rounded-lg text-sm font-semibold text-white disabled:opacity-40 flex items-center gap-1"
            style={{ background: '#0f172a' }}>
            {checking ? <RefreshCw size={14} className="animate-spin" /> : <Search size={14} />}ตรวจ
          </button>
        </div>
        <div>
          <div className="text-[11px] text-slate-500 mb-1 flex items-center gap-1"><MapPin size={11} />ตำแหน่ง</div>
          <input type="text" value={location} onChange={e => setLocation(e.target.value)} placeholder="A-1"
            className="w-full px-3 py-2 border rounded-lg outline-none focus:border-[#7C4DFF] text-[13px]"
            style={{ borderColor: '#e2e8f0' }} />
        </div>
      </div>

      {/* การ์ดสินค้าที่พบ */}
      {checkResult && (
        <div className="bg-white rounded-xl overflow-hidden" style={{ border: '1.5px solid #b6d0cc' }}>
          <div className="px-3 py-2.5 border-b" style={{ background: T.soft, borderColor: '#d5e5e2' }}>
            <div className="text-[10.5px] font-semibold" style={{ color: T.deep }}>พบสินค้าในฐานข้อมูล</div>
            <div className="text-[15px] font-bold text-slate-800 mt-0.5">{checkResult.name}</div>
            <div className="flex gap-2.5 mt-1 text-[11px] text-slate-500 flex-wrap">
              <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{checkResult.productCode}</span>
              {checkResult.unit && <span>หน่วย: {checkResult.unit}</span>}
              {checkResult.price ? <span>฿{checkResult.price}</span> : null}
            </div>
            {/* ยิงบาร์โค้ดสำรอง — บอกว่าเป็นสินค้าตัวเดียวกัน ไม่ให้คิดว่ายิงผิด */}
            {checkResult.scannedBarcode && checkResult.scannedBarcode !== checkResult.productCode && (
              <div className="text-[10.5px] mt-1" style={{ color: T.deep }}>
                ยิง <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{checkResult.scannedBarcode}</span> → รหัส {checkResult.productCode}
              </div>
            )}
            {checkResult.stockRatio > 1 && (
              <div className="text-[11px] font-bold mt-1 rounded-lg px-2 py-1" style={{ background: '#FFFBEB', color: '#B45309' }}>
                {checkResult.unit || 'แพ็ค'} · นับ 1 = {checkResult.stockRatio} ชิ้น
                {checkResult.masterCode && checkResult.masterCode !== checkResult.productCode && (
                  <span> · สต็อกอยู่ที่รหัส {checkResult.masterCode}</span>
                )}
              </div>
            )}
            {checkResult.barcodeConflict > 1 && (
              <div className="text-[11px] font-bold mt-1 rounded-lg px-2 py-1" style={{ background: '#FEF2F2', color: '#B91C1C' }}>
                บาร์โค้ดนี้ผูกกับสินค้า {checkResult.barcodeConflict} รายการ — เช็คให้ตรงก่อนนับ
              </div>
            )}
          </div>
          <div className="p-3 space-y-3">
            <div>
              <div className="text-xs font-semibold text-slate-600 mb-1.5">จำนวนที่นับได้</div>
              <div className="grid items-center gap-2.5" style={{ gridTemplateColumns: '52px minmax(0,1fr) 52px' }}>
                <button onClick={() => setQty(String(Math.max(0, (parseInt(qty) || 0) - 1)))}
                  className="rounded-xl border flex items-center justify-center text-2xl text-slate-700"
                  style={{ width: 52, height: 52, borderColor: '#e2e8f0', background: '#f8fafc' }}>−</button>
                <input ref={qtyInputRef} type="number" inputMode="numeric" value={qty}
                  onChange={e => setQty(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && handleAdd()} placeholder="0"
                  className="w-full min-w-0 text-center font-bold text-slate-800 border-0 outline-none"
                  style={{ fontSize: 30, fontFamily: "'IBM Plex Mono', monospace" }} />
                <button onClick={() => setQty(String((parseInt(qty) || 0) + 1))}
                  className="rounded-xl border flex items-center justify-center text-2xl"
                  style={{ width: 52, height: 52, borderColor: T.line, background: T.soft, color: T.deep }}>+</button>
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
          <div className="grid items-center gap-2.5" style={{ gridTemplateColumns: '48px minmax(0,1fr) 48px' }}>
            <button onClick={() => setQty(String(Math.max(0, (parseInt(qty) || 0) - 1)))}
              className="rounded-xl border flex items-center justify-center text-2xl text-slate-700"
              style={{ width: 48, height: 48, borderColor: '#e2e8f0', background: '#f8fafc' }}>−</button>
            <input type="number" inputMode="numeric" value={qty} onChange={e => setQty(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && handleAddManually()} placeholder="0"
              className="w-full min-w-0 text-center font-bold text-slate-800 border-0 outline-none"
              style={{ fontSize: 26, fontFamily: "'IBM Plex Mono', monospace" }} />
            <button onClick={() => setQty(String((parseInt(qty) || 0) + 1))}
              className="rounded-xl border flex items-center justify-center text-2xl text-[#B45309]"
              style={{ width: 48, height: 48, borderColor: '#fde68a', background: '#fffbeb' }}>+</button>
          </div>
          <div className="flex items-center gap-2 border rounded-xl px-3" style={{ borderColor: '#fde68a', background: '#fffbeb', minHeight: 48 }}>
            <span className="text-[11.5px] font-bold shrink-0" style={{ color: '#b45309' }}>ราคาขาย</span>
            <input type="number" inputMode="decimal" value={manualPrice} onChange={e => setManualPrice(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && handleAddManually()} placeholder="ใส่ถ้ารู้"
              className="flex-1 min-w-0 text-right font-bold text-slate-800 bg-transparent border-0 outline-none"
              style={{ fontSize: 16, fontFamily: "'IBM Plex Mono', monospace" }} />
            <span className="text-[11.5px] text-slate-400 shrink-0">บาท</span>
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

      {(pushDraft || pullDraft) && (
        <div className="bg-white border rounded-xl p-2.5 space-y-2" style={{ borderColor: '#e2e8f0' }}>
          <div className="grid grid-cols-2 gap-2">
            <button onClick={pushDraft} disabled={!!draftSync.busy || entries.length === 0}
              className="rounded-xl font-bold text-[12.5px] text-white flex items-center justify-center gap-1.5 disabled:opacity-40"
              style={{ minHeight: 46, background: T.main }}>
              {draftSync.busy === 'up' ? <RefreshCw size={14} className="animate-spin" /> : <Upload size={14} />}บันทึกขึ้นเซิร์ฟเวอร์
            </button>
            <button onClick={pullDraft} disabled={!!draftSync.busy}
              className="rounded-xl font-bold text-[12.5px] border flex items-center justify-center gap-1.5 disabled:opacity-40"
              style={{ minHeight: 46, borderColor: T.line, background: T.soft, color: T.deep }}>
              {draftSync.busy === 'down' ? <RefreshCw size={14} className="animate-spin" /> : <Download size={14} />}ดึงร่างจากเซิร์ฟเวอร์
            </button>
          </div>
          {draftSync.msg && <div className="text-[11px] font-semibold" style={{ color: T.deep }}>{draftSync.msg}</div>}
          {draftSync.err && <div className="text-[11px] font-semibold" style={{ color: '#B91C1C' }}>{draftSync.err}</div>}
        </div>
      )}

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
                style={{ width: 40, height: 40, background: '#fef2f2' }}><X size={14} /></button>
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
              style={{ width: 40, height: 40, background: '#F6F7F8' }}><Edit3 size={14} /></button>
          )}
        </div>
      )}
    </div>
  );
}

function CounterReviewView({ entries, setView, submitForReview, clearMyEntries, currentUser, pickedAt, tone, revise }) {
  const T = tone || { main: '#35706A', deep: '#2A5A55', soft: '#EAF1F0', line: '#B6D0CC' };
  const [note, setNote] = useState('');
  const [submitted, setSubmitted] = useState(null);
  const [confirming, setConfirming] = useState(false);
  const [sending, setSending] = useState(false);      // กดส่งซ้ำระหว่างกำลังส่ง = ได้ 2 ใบ
  const [foundOpen, setFoundOpen] = useState(false);  // รายการที่พบในระบบ เริ่มต้นพับไว้
  const [qtyOverrides, setQtyOverrides] = useState({});
  const editQty = (barcode, newQty) => setQtyOverrides(prev => ({ ...prev, [barcode]: newQty }));
  const grouped = useMemo(() => {
    const map = new Map();
    [...entries].sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp)).forEach(e => {
      // แพ็คกับชิ้นของตระกูลเดียวกันต้องรวมเป็นก้อนเดียว — สต็อกใน POS มีก้อนเดียว
      const key = e.masterCode || e.barcode;
      const pieces = (Number(e.qty) || 0) * (Number(e.stockRatio) || 1);
      if (map.has(key)) {
        const ex = map.get(key);
        ex.qty += e.qty; ex.pieces += pieces; ex.scans += 1;
        if (e.isPack) ex.hasPack = true; else ex.hasEach = true;
      } else map.set(key, {
        barcode: e.barcode, masterCode: key, stockRatio: Number(e.stockRatio) || 1,
        pieces, hasPack: !!e.isPack, hasEach: !e.isPack,
        productName: e.productName, unit: e.unit||'', price: e.price||0, cost: e.cost||0,
        qty: e.qty, scans: 1, scannedAt: e.timestamp, notFound: !!e.notFound, location: e.location||'',
      });
    });
    return Array.from(map.values()).map(g => {
      const ov = qtyOverrides[g.barcode];
      const qty = ov !== undefined ? ov : g.qty;
      // แก้จำนวนมือแล้ว ชิ้นต้องคิดใหม่จากตัวคูณ ไม่ใช้ยอดที่รวมไว้
      return { ...g, qty, pieces: ov !== undefined ? qty * (g.stockRatio || 1) : g.pieces, overridden: ov !== undefined };
    }).sort((a, b) => a.barcode.localeCompare(b.barcode));
  }, [entries, qtyOverrides]);
  const totalItems = grouped.length, totalQty = grouped.reduce((s, g) => s + g.qty, 0);
  const handleSubmit = async () => {
    if (!confirming) { setConfirming(true); return; }
    if (sending || submitted) return;                 // กันกดซ้ำ
    setSending(true);
    try {
      const sub = await submitForReview(grouped, note);
      clearMyEntries();
      setSubmitted(sub);
    } finally { setSending(false); }
  };
  if (submitted) return (
    <div className="space-y-3">
      <div className="bg-white rounded-2xl overflow-hidden border" style={{ borderColor: T.line }}>
        <div className="px-4 py-5 text-center" style={{ background: T.soft }}>
          <div className="text-white p-3 rounded-full inline-block mb-2" style={{ background: T.main }}><Send size={26} /></div>
          <h2 className="text-lg font-bold" style={{ color: T.deep }}>ส่งเรียบร้อยแล้ว</h2>
        </div>
        <div className="p-3 space-y-2.5">
          {submitted.docNo && (
            <div className="rounded-xl px-3 py-2.5 text-center" style={{ background: '#F6F7F8' }}>
              <div className="text-[10px] text-slate-500">เลขที่เอกสาร</div>
              <div className="text-lg font-bold text-slate-800 tabular-nums mt-0.5">{submitted.docNo}</div>
            </div>
          )}
          {submitted.startedAt && (
            <div className="rounded-xl px-3 py-2.5" style={{ background: '#EAF0F4' }}>
              <div className="text-[10px] font-bold" style={{ color: '#255771' }}>ช่วงเวลาที่นำมาคิด</div>
              <div className="text-[12px] tabular-nums mt-0.5" style={{ color: '#255771' }}>
                {new Date(submitted.startedAt).toLocaleString('th-TH', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })}
                {' → '}
                {new Date(submitted.submittedAt).toLocaleString('th-TH', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })}
              </div>
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
        </div>
      </div>
      <button onClick={() => setView('count')}
        className="w-full text-white font-bold text-[15px] rounded-xl flex items-center justify-center gap-2"
        style={{ minHeight: 52, background: T.main, boxShadow: `0 2px 0 ${T.deep}` }}><ScanLine size={18} />เริ่มนับรอบใหม่</button>
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
          className="text-white px-4 py-2.5 rounded-xl text-[13px] font-bold" style={{ background: T.main }}>ไปหน้านับสต็อก</button>
      </div>
    </div>
  );
  const found = grouped.filter(g => !g.notFound), missing = grouped.filter(g => g.notFound);
  return (
    <div className="space-y-3">
      <div>
        <h2 className="text-xl font-bold text-slate-800">ตรวจสอบและส่ง</h2>
      </div>

      <div className="bg-white border rounded-xl grid grid-cols-2 divide-x" style={{ borderColor: '#E4E6EA' }}>
        <div className="px-3 py-2.5">
          <div className="text-[10.5px] text-slate-500">บาร์โค้ด</div>
          <div className="text-xl font-bold text-slate-800 tabular-nums">{totalItems}</div>
        </div>
        <div className="px-3 py-2.5" style={{ borderColor: '#EEF0F3' }}>
          <div className="text-[10.5px] text-slate-500">รวมจำนวน</div>
          <div className="text-xl font-bold tabular-nums" style={{ color: T.main }}>{totalQty.toLocaleString('th-TH')}</div>
        </div>
      </div>

      {found.length > 0 && (
        <div className="bg-white rounded-xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
          <button onClick={() => setFoundOpen(o => !o)}
            className="w-full text-left px-3 py-2.5 border-b flex items-center gap-2"
            style={{ background: T.soft, borderColor: '#D5E5E2', minHeight: 48 }}>
            <CheckCircle2 size={13} style={{ color: T.deep }} className="shrink-0" />
            <span className="text-[11.5px] font-bold" style={{ color: T.deep }}>พบในระบบ · {found.length} รายการ</span>
            <span className="ml-auto text-[10.5px] font-semibold shrink-0" style={{ color: T.main }}>{foundOpen ? 'ซ่อน' : 'ดูรายการ'}</span>
          </button>
          {foundOpen && (
            <div className="divide-y max-h-64 overflow-y-auto" style={{ borderColor: '#F6F7F8' }}>
              {found.map(g => <GroupedRow key={g.barcode} g={g} onEditQty={editQty} />)}
            </div>
          )}
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
        </div>
      )}

      <div className="bg-white rounded-xl border p-3" style={{ borderColor: '#E4E6EA' }}>
        <label className="text-[12px] font-semibold text-slate-700 mb-1.5 block">หมายเหตุถึงผู้จัดการ</label>
        <textarea value={note} onChange={e => setNote(e.target.value)} rows={2}
          placeholder=""
          className="w-full px-3 py-2 border rounded-lg outline-none  text-[13px] resize-none"
          style={{ borderColor: '#E2E8F0' }} />
      </div>

      {!confirming ? (
        <button onClick={handleSubmit} disabled={sending}
          className="w-full text-white font-bold text-[15.5px] rounded-xl flex items-center justify-center gap-2 disabled:opacity-60"
          style={{ minHeight: 54, background: revise ? '#B91C1C' : T.main, boxShadow: `0 2px 0 ${revise ? '#7F1D1D' : T.deep}` }}><Send size={19} />
          {revise ? `ส่งรอบแก้ (${String(revise.docNo || '').replace(/R\d+$/i, '')}R${(revise.reviseNo || 0) + 1})` : 'ส่งให้ผู้จัดการรีวิว'}</button>
      ) : (
        <div className="space-y-3">
          <div className="text-[11px] font-bold tracking-wide text-[#B45309]">ยืนยันก่อนส่ง</div>
          <div className="text-xl font-bold text-slate-800 leading-snug">
            ใบนี้จะออกในชื่อ {currentUser?.name || 'พนักงาน'} — ใช่ไหม
          </div>
          <div className="bg-white border-2 border-[#FDE68A] rounded-2xl p-4">
            <div className="flex items-center gap-3">
              <span className="w-11 h-11 shrink-0 rounded-full   flex items-center justify-center text-lg font-bold">
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
          <button onClick={handleSubmit} className="w-full text-white py-4 rounded-xl font-bold flex items-center justify-center gap-2 shadow-lg" style={{ background: T.main }}>
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

// บิลที่พนักงานส่งไปแล้ว — ตาราง 1 บรรทัด 1 ใบ แตะเข้าดูรายละเอียดอีกชั้น
const TH_MONTHS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
// dd-mmm-yyyy — ไม่สับสนระหว่างวันกับเดือน
const thDocDate = (v) => {
  if (!v) return '—';
  const iso = String(v).slice(0, 10);
  const [y, m, d] = iso.split('-');
  if (!y || !m || !d) return String(v);
  return `${+d}-${TH_MONTHS[+m - 1] || m}-${y}`;
};
const bahtFmt = (n) => (Number(n) || 0).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function MyBillsView({ invSubs = [], setView, onRefresh, onRevise }) {
  const [openId, setOpenId] = useState(null);
  const [q, setQ] = useState('');
  const [bc, setBc] = useState('');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const [status, setStatus] = useState('all');
  const [adv, setAdv] = useState(true);
  useEffect(() => { onRefresh?.(); }, []);   // เปิดหน้าแล้วเอาสถานะล่าสุดเลย

  const counts = invSubs.reduce((a, s2) => { const k = s2.status || 'pending'; a[k] = (a[k] || 0) + 1; return a; }, {});
  const CFG = {
    rejected:   { label: 'ต้องแก้',       soft: '#FEF2F2', line: '#FECACA', ink: '#B91C1C', edge: '#B91C1C', bg: '#FEF2F2' },
    pending:    { label: 'รอรีวิว',       soft: '#FFFBEB', line: '#FDE68A', ink: '#B45309', edge: '#FDE68A', bg: '#fff' },
    approved:   { label: 'อนุมัติแล้ว',    soft: '#F0FDF4', line: '#BBF7D0', ink: '#15803D', edge: '#BBF7D0', bg: '#fff' },
    superseded: { label: 'แก้แล้วส่งใหม่', soft: '#F6F7F8', line: '#E4E6EA', ink: '#64748B', edge: '#E4E6EA', bg: '#fff' },
  };
  const RANK = { rejected: 0, pending: 1, approved: 2, superseded: 3 };
  const linesOf = (s) => Array.isArray(s.lines) ? s.lines : [];
  const bcOf = (d) => String(d.barcode || d.product_code || '');
  const qtyOf = (d) => Number(d.quantity ?? d.qty ?? 0) || 0;

  const bb = bc.trim(), qq = q.trim().toLowerCase();
  const list = invSubs.filter(s => {
    const st = s.status || 'pending';
    if (status !== 'all' && st !== status) return false;
    const dt = String(s.invoiceDate || '').slice(0, 10);
    if (from && (!dt || dt < from)) return false;
    if (to && (!dt || dt > to)) return false;
    if (bb && !linesOf(s).some(d => bcOf(d).includes(bb))) return false;
    if (qq && ![s.docNo, s.invoiceNo, s.vendorName].join(' ').toLowerCase().includes(qq)) return false;
    return true;
  }).sort((a, b2) => (RANK[a.status] ?? 9) - (RANK[b2.status] ?? 9)
    || new Date(b2.submittedAt) - new Date(a.submittedAt));

  const cur = openId ? invSubs.find(s => s.id === openId) : null;
  const dirty = !!(q || bc || from || to || status !== 'all');
  const COLS = 'minmax(124px,1.1fr) minmax(96px,1.6fr) minmax(78px,.8fr) minmax(86px,.9fr) minmax(40px,.5fr) minmax(84px,.9fr) 18px';

  // ── ชั้นรายละเอียด ──
  if (cur) {
    const cfg = CFG[cur.status] || CFG.pending;
    const lines = linesOf(cur);
    const idx = list.findIndex(s => s.id === cur.id);
    return (
      <div className="space-y-2.5">
        <div className="flex items-center gap-2">
          <button onClick={() => setOpenId(null)}
            className="shrink-0 rounded-lg border bg-white font-bold text-[12.5px] text-slate-600"
            style={{ minHeight: 38, padding: '0 12px', borderColor: '#E4E6EA' }}>‹ กลับ</button>
          <div className="flex-1 min-w-0">
            <div className="text-[15px] font-bold text-slate-800 tabular-nums truncate">{cur.docNo || '—'}</div>
            <div className="text-[11px] text-slate-400">
              {idx >= 0 ? `ใบที่ ${idx + 1} จาก ${list.length} ในรายการที่กรองไว้` : 'ไม่อยู่ในผลกรองปัจจุบัน'}
            </div>
          </div>
          <span className="shrink-0 rounded-full border text-[11.5px] font-bold"
            style={{ padding: '6px 12px', background: cfg.soft, borderColor: cfg.line, color: cfg.ink }}>
            {cfg.label}{cur.reviseNo > 0 ? ` · แก้รอบ ${cur.reviseNo}` : ''}
          </span>
        </div>

        {cur.reviewNote && cur.status !== 'pending' && (
          <div className="rounded-xl border p-3" style={{ background: cfg.soft, borderColor: cfg.line }}>
            <div className="text-[10.5px] font-bold" style={{ color: cfg.ink }}>{cur.reviewedBy || 'ผู้จัดการ'}</div>
            <div className="text-[12.5px] mt-1 leading-relaxed" style={{ color: cfg.ink }}>{cur.reviewNote}</div>
          </div>
        )}

        <div className="bg-white rounded-xl border p-3 space-y-2.5" style={{ borderColor: '#E4E6EA' }}>
          <div className="text-[15px] font-bold text-slate-800">{cur.vendorName || 'ไม่ระบุผู้ขาย'}</div>
          <div className="grid grid-cols-3 gap-2">
            {[['เลขที่บิล', cur.invoiceNo || '—'], ['วันที่บิล', thDocDate(cur.invoiceDate)],
              ['ส่งเมื่อ', cur.submittedAt ? new Date(cur.submittedAt).toLocaleString('th-TH', { day:'numeric', month:'short', hour:'2-digit', minute:'2-digit' }) : '—']].map(([k, v]) => (
              <div key={k} className="rounded-xl px-2.5 py-2" style={{ background: '#F6F7F8' }}>
                <div className="text-[10px] text-slate-500">{k}</div>
                <div className="text-[12.5px] font-bold text-slate-800 mt-0.5 tabular-nums truncate">{v}</div>
              </div>
            ))}
          </div>
          <div className="flex items-baseline gap-2 border-t pt-2.5" style={{ borderColor: '#F1F3F5' }}>
            <span className="text-[12px] text-slate-500">{lines.length} รายการ</span>
            <span className="flex-1" />
            <span className="text-[11px] text-slate-500">ยอดสุทธิ</span>
            <span className="text-[19px] font-bold text-slate-800 tabular-nums">{bahtFmt(cur.netTotal)}</span>
          </div>
        </div>

        {lines.length > 0 && (
          <div className="bg-white rounded-xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
            <div className="grid gap-2 px-3 py-2 border-b" style={{ gridTemplateColumns: 'minmax(0,1fr) 46px 74px 88px', background: '#F8FAFC', borderColor: '#E4E6EA' }}>
              {['สินค้า / บาร์โค้ด', 'จำนวน', 'ราคา/หน่วย', 'รวม'].map((h, i) => (
                <span key={h} className="text-[10.5px] font-bold text-slate-500" style={{ textAlign: i ? 'right' : 'left' }}>{h}</span>
              ))}
            </div>
            {lines.map((d, i) => {
              const hit = bb && bcOf(d).includes(bb);
              return (
                <div key={i} className="grid gap-2 items-center px-3 py-2 border-b"
                  style={{ gridTemplateColumns: 'minmax(0,1fr) 46px 74px 88px', borderColor: '#F6F7F8', background: hit ? '#FFFBEB' : '#fff' }}>
                  <div className="min-w-0">
                    <div className="text-[12px] font-semibold text-slate-800 truncate">{d.product_name || d.productName || d.description || '—'}</div>
                    <div className="text-[10px] text-slate-400 tabular-nums">{bcOf(d) || '—'}</div>
                  </div>
                  <span className="text-[12.5px] text-slate-600 text-right tabular-nums">×{qtyOf(d)}</span>
                  <span className="text-[12px] text-slate-500 text-right tabular-nums">{d.price_ea != null ? bahtFmt(d.price_ea) : '—'}</span>
                  <span className="text-[12.5px] font-semibold text-slate-800 text-right tabular-nums">{d.total != null ? bahtFmt(d.total) : '—'}</span>
                </div>
              );
            })}
          </div>
        )}

        {cur.status === 'rejected' && (
          <button onClick={() => onRevise?.({
              id: cur.id, docNo: cur.docNo, invoiceNo: cur.invoiceNo, note: cur.reviewNote,
              reviseNo: cur.reviseNo || 0,
              // ข้อมูลใบเดิมทั้งหมด — แก้ต่อได้เลย ไม่ต้องถ่ายรูปใหม่
              data: { ...(cur.header || {}),
                      invoice_no: cur.invoiceNo || cur.header?.invoice_no || '',
                      invoice_date: cur.invoiceDate || cur.header?.invoice_date || '',
                      vendor_name: cur.vendorName || cur.header?.vendor_name || '',
                      products: cur.lines || [] },
              fileName: cur.fileName || '',
            })}
            className="w-full text-white font-bold text-[14px] rounded-xl" style={{ minHeight: 48, background: '#B91C1C' }}>
            แก้ใบนี้ส่งใหม่ ({cur.docNo ? cur.docNo.replace(/R\d+$/i, '') + 'R' + ((cur.reviseNo || 0) + 1) : 'รอบใหม่'})
          </button>
        )}
      </div>
    );
  }

  // ── ชั้นตาราง ──
  return (
    <div className="space-y-2.5">
      <div className="flex items-baseline gap-2">
        <h2 className="text-[19px] font-bold text-slate-800">บิลที่ส่งแล้ว</h2>
        <div className="flex-1" />
        <div className="text-[11.5px] text-slate-400 whitespace-nowrap">
          {dirty ? `${list.length} จาก ${invSubs.length} ใบ` : `${invSubs.length} ใบ`}
        </div>
        <button onClick={onRefresh} title="ดึงสถานะล่าสุด"
          className="shrink-0 rounded-lg flex items-center justify-center border bg-white text-slate-500"
          style={{ width: 30, height: 30, borderColor: '#E4E6EA' }}><RefreshCw size={13} /></button>
      </div>

      <div className="bg-white border rounded-xl p-2.5 space-y-2" style={{ borderColor: '#E4E6EA' }}>
        <div className="flex gap-1.5">
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="ค้นเลขที่เอกสาร บิล หรือผู้ขาย"
            className="flex-1 min-w-0 rounded-lg border text-[12px] text-slate-800 px-2.5"
            style={{ height: 30, borderColor: '#E4E6EA', boxSizing: 'border-box' }} />
          <button onClick={() => setAdv(!adv)}
            className="shrink-0 rounded-lg font-bold text-[11.5px] border"
            style={{ minWidth: 46, height: 30,
                     background: adv ? '#0F172A' : '#fff', color: adv ? '#fff' : '#475569',
                     borderColor: adv ? '#0F172A' : '#E4E6EA' }}>{adv ? 'ปิด' : 'กรอง'}</button>
        </div>

        {adv && (
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-1.5 pt-0.5 border-t" style={{ borderColor: '#F1F3F5' }}>
            <label className="flex flex-col gap-1">
              <span className="text-[10.5px] font-semibold text-slate-500">สถานะ</span>
              <select value={status} onChange={e => setStatus(e.target.value)}
                className="rounded-lg border text-[12px] text-slate-800 bg-white px-1.5"
                style={{ height: 34, borderColor: '#E4E6EA', boxSizing: 'border-box' }}>
                <option value="all">ทุกสถานะ ({invSubs.length})</option>
                {['rejected', 'pending', 'approved', 'superseded'].map(k => (
                  <option key={k} value={k}>{CFG[k].label} ({counts[k] || 0})</option>
                ))}
              </select>
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-[10.5px] font-semibold text-slate-500">บิลจาก</span>
              <input type="date" value={from} onChange={e => setFrom(e.target.value)}
                className="rounded-lg border text-[12px] text-slate-800 px-2"
                style={{ height: 34, borderColor: '#E4E6EA', boxSizing: 'border-box' }} />
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-[10.5px] font-semibold text-slate-500">ถึง</span>
              <input type="date" value={to} onChange={e => setTo(e.target.value)}
                className="rounded-lg border text-[12px] text-slate-800 px-2"
                style={{ height: 34, borderColor: '#E4E6EA', boxSizing: 'border-box' }} />
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-[10.5px] font-semibold text-slate-500">บาร์โค้ดในบิล</span>
              <input value={bc} onChange={e => setBc(e.target.value)} inputMode="numeric" placeholder="8851753098835"
                className="rounded-lg border text-[12px] text-slate-800 px-2 tabular-nums"
                style={{ height: 34, borderColor: '#E4E6EA', boxSizing: 'border-box' }} />
            </label>
          </div>
        )}

        {dirty && (
          <div className="flex items-center gap-2">
            <span className="flex-1 text-[11px] text-slate-500 truncate">
              {[status !== 'all' && CFG[status].label, (from || to) && `${from || 'เริ่มต้น'} → ${to || 'ล่าสุด'}`,
                bc && 'บาร์โค้ด ' + bc, q && `“${q}”`].filter(Boolean).join(' · ')}
            </span>
            <button onClick={() => { setQ(''); setBc(''); setFrom(''); setTo(''); setStatus('all'); }}
              className="shrink-0 rounded-lg border bg-white text-[11.5px] font-bold text-slate-600"
              style={{ height: 30, padding: '0 10px', borderColor: '#E4E6EA' }}>ล้างตัวกรอง</button>
          </div>
        )}
      </div>

      {invSubs.length === 0 ? (
        <div className="bg-white rounded-xl px-4 py-9 text-center" style={{ border: '1px dashed #CBD5E1' }}>
          <Receipt className="mx-auto text-[#E4E6EA] mb-2" size={34} />
          <div className="text-[13px] text-slate-500 mb-3">ยังไม่มีบิลที่ส่ง</div>
          <button onClick={() => setView('invoice')}
            className="text-white px-4 py-2.5 rounded-xl text-[13px] font-bold" style={{ background: '#2F6E90' }}>ไปบันทึกบิล</button>
        </div>
      ) : (
        <div className="bg-white border rounded-xl overflow-x-auto" style={{ borderColor: '#E4E6EA' }}>
          <div className="grid gap-2 px-3 py-2 border-b" style={{ gridTemplateColumns: COLS, minWidth: 600, background: '#F8FAFC', borderColor: '#E4E6EA' }}>
            {[['เลขที่เอกสาร', 0], ['ผู้ขาย', 0], ['วันที่บิล', 0], ['สถานะ', 0], ['จำนวน', 1], ['ยอดสุทธิ', 1], ['', 0]].map(([h, right], i) => (
              <span key={i} className="text-[10.5px] font-bold text-slate-500" style={{ textAlign: right ? 'right' : 'left' }}>{h}</span>
            ))}
          </div>

          {list.map(s => {
            const cfg = CFG[s.status] || CFG.pending;
            const lines = linesOf(s);
            return (
              <button key={s.id} onClick={() => setOpenId(s.id)}
                className="w-full grid gap-2 items-center px-3 py-2.5 text-left border-b"
                style={{ gridTemplateColumns: COLS, minWidth: 600, boxSizing: 'border-box',
                         borderColor: '#F1F3F5', borderLeft: `3px solid ${cfg.edge}`, background: cfg.bg }}>
                <span className="min-w-0 text-[11.5px] font-bold text-slate-800 tabular-nums truncate">{s.docNo || '—'}</span>
                <span className="min-w-0 text-[12.5px] font-semibold text-slate-800 truncate">{s.vendorName || 'ไม่ระบุผู้ขาย'}</span>
                <span className="text-[11.5px] text-slate-500 tabular-nums">{thDocDate(s.invoiceDate)}</span>
                <span className="justify-self-start rounded-full text-[10px] font-bold whitespace-nowrap"
                  style={{ padding: '3px 8px', background: cfg.soft, color: cfg.ink }}>
                  {cfg.label}{s.reviseNo > 0 ? ` R${s.reviseNo}` : ''}
                </span>
                <span className="text-[12.5px] text-slate-600 text-right tabular-nums">
                  {lines.reduce((a, d) => a + qtyOf(d), 0) || s.itemCount || 0}
                </span>
                <span className="text-[13px] font-bold text-slate-800 text-right tabular-nums">{bahtFmt(s.netTotal)}</span>
                <span className="text-[13px] text-slate-300 text-right">›</span>
              </button>
            );
          })}

          {list.length === 0 && (
            <div className="px-4 py-8 text-center text-[13px] text-slate-500">ไม่มีเอกสารตรงตามที่กรอง</div>
          )}
        </div>
      )}
    </div>
  );
}

// ใบนับที่พนักงานส่งไปแล้ว — ตาราง 1 บรรทัด 1 ใบ แตะเข้าดูรายละเอียดอีกชั้น (โครงเดียวกับบิลที่ส่งแล้ว)
function MySubmissionsView({ submissions, setView, onRefresh, subSync = {}, onRevise, onResend }) {
  const [openId, setOpenId] = useState(null);
  const [q, setQ] = useState('');
  const [bc, setBc] = useState('');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const [status, setStatus] = useState('all');
  const [adv, setAdv] = useState(true);

  const counts = submissions.reduce((a, s2) => { const k = s2.status || 'pending'; a[k] = (a[k] || 0) + 1; return a; }, {});
  const CFG = {
    rejected:   { label: 'ต้องแก้',       soft: '#FEF2F2', line: '#FECACA', ink: '#B91C1C', edge: '#B91C1C', bg: '#FEF2F2' },
    pending:    { label: 'รอรีวิว',       soft: '#FFFBEB', line: '#FDE68A', ink: '#B45309', edge: '#FDE68A', bg: '#fff' },
    approved:   { label: 'อนุมัติแล้ว',    soft: '#F0FDF4', line: '#BBF7D0', ink: '#15803D', edge: '#BBF7D0', bg: '#fff' },
    superseded: { label: 'แก้แล้วส่งใหม่', soft: '#F6F7F8', line: '#E4E6EA', ink: '#64748B', edge: '#E4E6EA', bg: '#fff' },
  };
  const RANK = { rejected: 0, pending: 1, approved: 2, superseded: 3 };
  const rowsOf = (s2) => Array.isArray(s2.data) ? s2.data : [];

  const bb = bc.trim(), qq = q.trim().toLowerCase();
  const list = submissions.filter(s2 => {
    const st = s2.status || 'pending';
    if (status !== 'all' && st !== status) return false;
    const dt = String(s2.submittedAt || '').slice(0, 10);   // ใบนับไม่มีวันที่บิล ใช้วันที่ส่ง
    if (from && (!dt || dt < from)) return false;
    if (to && (!dt || dt > to)) return false;
    if (bb && !rowsOf(s2).some(d => String(d.barcode || '').includes(bb))) return false;
    if (qq && ![s2.docNo, s2.note].join(' ').toLowerCase().includes(qq)) return false;
    return true;
  }).sort((a, b2) => (RANK[a.status] ?? 9) - (RANK[b2.status] ?? 9)
    || new Date(b2.submittedAt) - new Date(a.submittedAt));

  const cur = openId ? submissions.find(s2 => s2.id === openId) : null;
  const dirty = !!(q || bc || from || to || status !== 'all');
  const COLS = 'minmax(126px,1.3fr) minmax(96px,1fr) minmax(86px,.9fr) minmax(46px,.6fr) minmax(56px,.7fr) 18px';
  const dt = (v, opt) => v ? new Date(v).toLocaleString('th-TH', opt || { day:'numeric', month:'short', hour:'2-digit', minute:'2-digit' }) : '—';

  // ── ชั้นรายละเอียด ──
  if (cur) {
    const cfg = CFG[cur.status] || CFG.pending;
    const rows = rowsOf(cur);
    const idx = list.findIndex(s2 => s2.id === cur.id);
    return (
      <div className="space-y-2.5">
        <div className="flex items-center gap-2">
          <button onClick={() => setOpenId(null)}
            className="shrink-0 rounded-lg border bg-white font-bold text-[12.5px] text-slate-600"
            style={{ minHeight: 38, padding: '0 12px', borderColor: '#E4E6EA' }}>‹ กลับ</button>
          <div className="flex-1 min-w-0">
            <div className="text-[15px] font-bold text-slate-800 tabular-nums truncate">{cur.docNo || '—'}</div>
            <div className="text-[11px] text-slate-400">
              {idx >= 0 ? `ใบที่ ${idx + 1} จาก ${list.length} ในรายการที่กรองไว้` : 'ไม่อยู่ในผลกรองปัจจุบัน'}
            </div>
          </div>
          <span className="shrink-0 rounded-full border text-[11.5px] font-bold"
            style={{ padding: '6px 12px', background: cfg.soft, borderColor: cfg.line, color: cfg.ink }}>
            {cfg.label}{cur.reviseNo > 0 ? ` · แก้รอบ ${cur.reviseNo}` : ''}
          </span>
        </div>

        {cur.reviewNote && cur.status !== 'pending' && (
          <div className="rounded-xl border p-3" style={{ background: cfg.soft, borderColor: cfg.line }}>
            <div className="text-[10.5px] font-bold" style={{ color: cfg.ink }}>{cur.reviewedBy || 'ผู้จัดการ'}</div>
            <div className="text-[12.5px] mt-1 leading-relaxed" style={{ color: cfg.ink }}>{cur.reviewNote}</div>
          </div>
        )}

        <div className="bg-white rounded-xl border p-3 space-y-2.5" style={{ borderColor: '#E4E6EA' }}>
          <div className="rounded-xl px-2.5 py-2" style={{ background: '#EAF0F4' }}>
            <div className="text-[10px] font-bold" style={{ color: '#255771' }}>ช่วงเวลาที่นำมาคิด</div>
            <div className="text-[12px] tabular-nums mt-0.5" style={{ color: '#255771' }}>
              {dt(cur.startedAt)} → {dt(cur.submittedAt)}
            </div>
          </div>
          <div className="grid grid-cols-2 gap-2">
            <div className="rounded-xl px-2.5 py-2" style={{ background: '#F6F7F8' }}>
              <div className="text-[10px] text-slate-500">บาร์โค้ด</div>
              <div className="text-[15px] font-bold text-slate-800 tabular-nums mt-0.5">{cur.itemCount ?? rows.length}</div>
            </div>
            <div className="rounded-xl px-2.5 py-2" style={{ background: '#F6F7F8' }}>
              <div className="text-[10px] text-slate-500">รวมจำนวน</div>
              <div className="text-[15px] font-bold text-slate-800 tabular-nums mt-0.5">{(cur.totalQty || 0).toLocaleString('th-TH')}</div>
            </div>
          </div>
          {cur.note && <div className="text-[11.5px] text-slate-500 leading-relaxed">หมายเหตุที่ส่งไป: “{cur.note}”</div>}
        </div>

        {rows.length > 0 && (
          <div className="bg-white rounded-xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
            <div className="grid gap-2 px-3 py-2 border-b" style={{ gridTemplateColumns: 'minmax(0,1fr) 74px 52px', background: '#F8FAFC', borderColor: '#E4E6EA' }}>
              {['สินค้า / บาร์โค้ด', 'จุดเก็บ', 'จำนวน'].map((h, i) => (
                <span key={h} className="text-[10.5px] font-bold text-slate-500" style={{ textAlign: i === 2 ? 'right' : 'left' }}>{h}</span>
              ))}
            </div>
            {rows.map((d, i) => {
              const hit = bb && String(d.barcode || '').includes(bb);
              return (
                <div key={i} className="grid gap-2 items-center px-3 py-2 border-b"
                  style={{ gridTemplateColumns: 'minmax(0,1fr) 74px 52px', borderColor: '#F6F7F8', background: hit ? '#FFFBEB' : '#fff' }}>
                  <div className="min-w-0">
                    <div className="text-[12px] font-semibold text-slate-800 truncate">{d.productName || '—'}</div>
                    <div className="text-[10px] text-slate-400 tabular-nums">{d.barcode || '—'}</div>
                  </div>
                  <span className="text-[11px] text-slate-500 truncate">{d.location || '—'}</span>
                  <span className="text-[12.5px] font-semibold text-slate-800 text-right tabular-nums">{d.qty}</span>
                </div>
              );
            })}
          </div>
        )}

        {cur.status === 'rejected' && (
          <button onClick={() => onRevise?.({
              id: cur.id, docNo: cur.docNo, note: cur.reviewNote,
              reviseNo: cur.reviseNo || 0, featureType: cur.featureType,
              data: cur.data || [],   // ยกรายการเดิมมาแก้ ไม่ต้องสแกนใหม่
            })}
            className="w-full text-white font-bold text-[14px] rounded-xl" style={{ minHeight: 48, background: '#B91C1C' }}>
            แก้ใบนี้ส่งใหม่ ({cur.docNo ? cur.docNo.replace(/R\d+$/i, '') + 'R' + ((cur.reviseNo || 0) + 1) : 'รอบใหม่'})
          </button>
        )}
      </div>
    );
  }

  // ── ชั้นตาราง ──
  return (
    <div className="space-y-2.5">
      <div className="flex items-baseline gap-2">
        <h2 className="text-[19px] font-bold text-slate-800">ใบที่ส่งแล้ว</h2>
        <div className="flex-1" />
        <div className="text-[11.5px] text-slate-400 whitespace-nowrap">
          {dirty ? `${list.length} จาก ${submissions.length} ใบ` : `${submissions.length} ใบ`}
        </div>
        {onRefresh && (
          <button onClick={onRefresh} disabled={subSync.busy} title="ดึงสถานะล่าสุดจากเซิร์ฟเวอร์"
            className="shrink-0 rounded-lg flex items-center justify-center border bg-white text-slate-500"
            style={{ width: 30, height: 30, borderColor: '#E4E6EA' }}>
            <RefreshCw size={13} className={subSync.busy ? 'animate-spin' : ''} />
          </button>
        )}
      </div>

      <div className="bg-white border rounded-xl p-2.5 space-y-2" style={{ borderColor: '#E4E6EA' }}>
        <div className="flex gap-1.5">
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="ค้นเลขที่เอกสาร หรือหมายเหตุ"
            className="flex-1 min-w-0 rounded-lg border text-[12px] text-slate-800 px-2.5"
            style={{ height: 30, borderColor: '#E4E6EA', boxSizing: 'border-box' }} />
          <button onClick={() => setAdv(!adv)}
            className="shrink-0 rounded-lg font-bold text-[11.5px] border"
            style={{ minWidth: 46, height: 30,
                     background: adv ? '#0F172A' : '#fff', color: adv ? '#fff' : '#475569',
                     borderColor: adv ? '#0F172A' : '#E4E6EA' }}>{adv ? 'ปิด' : 'กรอง'}</button>
        </div>

        {adv && (
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-1.5 pt-0.5 border-t" style={{ borderColor: '#F1F3F5' }}>
            <label className="flex flex-col gap-1">
              <span className="text-[10.5px] font-semibold text-slate-500">สถานะ</span>
              <select value={status} onChange={e => setStatus(e.target.value)}
                className="rounded-lg border text-[12px] text-slate-800 bg-white px-1.5"
                style={{ height: 34, borderColor: '#E4E6EA', boxSizing: 'border-box' }}>
                <option value="all">ทุกสถานะ ({submissions.length})</option>
                {['rejected', 'pending', 'approved', 'superseded'].map(k => (
                  <option key={k} value={k}>{CFG[k].label} ({counts[k] || 0})</option>
                ))}
              </select>
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-[10.5px] font-semibold text-slate-500">ส่งจาก</span>
              <input type="date" value={from} onChange={e => setFrom(e.target.value)}
                className="rounded-lg border text-[12px] text-slate-800 px-2"
                style={{ height: 34, borderColor: '#E4E6EA', boxSizing: 'border-box' }} />
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-[10.5px] font-semibold text-slate-500">ถึง</span>
              <input type="date" value={to} onChange={e => setTo(e.target.value)}
                className="rounded-lg border text-[12px] text-slate-800 px-2"
                style={{ height: 34, borderColor: '#E4E6EA', boxSizing: 'border-box' }} />
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-[10.5px] font-semibold text-slate-500">บาร์โค้ดในใบ</span>
              <input value={bc} onChange={e => setBc(e.target.value)} inputMode="numeric" placeholder="8851753098835"
                className="rounded-lg border text-[12px] text-slate-800 px-2 tabular-nums"
                style={{ height: 34, borderColor: '#E4E6EA', boxSizing: 'border-box' }} />
            </label>
          </div>
        )}

        {dirty && (
          <div className="flex items-center gap-2">
            <span className="flex-1 text-[11px] text-slate-500 truncate">
              {[status !== 'all' && CFG[status].label, (from || to) && `${from || 'เริ่มต้น'} → ${to || 'ล่าสุด'}`,
                bc && 'บาร์โค้ด ' + bc, q && `“${q}”`].filter(Boolean).join(' · ')}
            </span>
            <button onClick={() => { setQ(''); setBc(''); setFrom(''); setTo(''); setStatus('all'); }}
              className="shrink-0 rounded-lg border bg-white text-[11.5px] font-bold text-slate-600"
              style={{ height: 30, padding: '0 10px', borderColor: '#E4E6EA' }}>ล้างตัวกรอง</button>
          </div>
        )}
      </div>

      {submissions.length === 0 ? (
        <div className="bg-white rounded-xl px-4 py-9 text-center" style={{ border: '1px dashed #CBD5E1' }}>
          <Send className="mx-auto text-[#E4E6EA] mb-2" size={34} />
          <div className="text-[13px] text-slate-500 mb-3">ยังไม่มีใบที่ส่ง</div>
          <button onClick={() => setView('count')}
            className="text-white px-4 py-2.5 rounded-xl text-[13px] font-bold" style={{ background: '#35706A' }}>ไปนับสต็อก</button>
        </div>
      ) : (
        <div className="bg-white border rounded-xl overflow-x-auto" style={{ borderColor: '#E4E6EA' }}>
          <div className="grid gap-2 px-3 py-2 border-b" style={{ gridTemplateColumns: COLS, minWidth: 600, background: '#F8FAFC', borderColor: '#E4E6EA' }}>
            {[['เลขที่เอกสาร', 0], ['ส่งเมื่อ', 0], ['สถานะ', 0], ['บาร์โค้ด', 1], ['รวมจำนวน', 1], ['', 0]].map(([h, right], i) => (
              <span key={i} className="text-[10.5px] font-bold text-slate-500" style={{ textAlign: right ? 'right' : 'left' }}>{h}</span>
            ))}
          </div>

          {list.map(s2 => {
            const cfg = CFG[s2.status] || CFG.pending;
            if (s2.syncError || s2.sending) return (
              <div key={s2.id} className="px-3 py-2.5 border-b" style={{ borderColor: '#F1F3F5', borderLeft: '3px solid #B91C1C', background: '#FDF2F2' }}>
                <div className="flex items-center gap-2">
                  <span className="text-[11.5px] font-bold tabular-nums" style={{ color: '#B91C1C' }}>{s2.docNo || '—'}</span>
                  <span className="text-[11px] text-slate-500">{s2.itemCount ?? rowsOf(s2).length} รายการ · {(s2.totalQty || 0).toLocaleString('th-TH')}</span>
                  <span className="flex-1" />
                  <button onClick={() => onResend?.(s2.id)} disabled={s2.sending}
                    className="shrink-0 rounded-lg text-[11.5px] font-bold text-white disabled:opacity-60"
                    style={{ minHeight: 30, padding: '0 10px', background: '#B91C1C' }}>
                    {s2.sending ? 'กำลังส่ง…' : 'ส่งซ้ำ'}
                  </button>
                </div>
                <div className="text-[11px] mt-1 break-words" style={{ color: '#B91C1C' }}>
                  ยังไม่ขึ้นเซิร์ฟเวอร์ — {s2.syncError || 'กำลังลองใหม่'}
                </div>
              </div>
            );
            return (
              <button key={s2.id} onClick={() => setOpenId(s2.id)}
                className="w-full grid gap-2 items-center px-3 py-2.5 text-left border-b"
                style={{ gridTemplateColumns: COLS, minWidth: 600, boxSizing: 'border-box',
                         borderColor: '#F1F3F5', borderLeft: `3px solid ${cfg.edge}`, background: cfg.bg }}>
                <span className="min-w-0 text-[11.5px] font-bold text-slate-800 tabular-nums truncate">{s2.docNo || '—'}</span>
                <span className="text-[11.5px] text-slate-500 tabular-nums truncate">{dt(s2.submittedAt)}</span>
                <span className="justify-self-start rounded-full text-[10px] font-bold whitespace-nowrap"
                  style={{ padding: '3px 8px', background: cfg.soft, color: cfg.ink }}>
                  {cfg.label}{s2.reviseNo > 0 ? ` R${s2.reviseNo}` : ''}
                </span>
                <span className="text-[12.5px] text-slate-600 text-right tabular-nums">{s2.itemCount ?? rowsOf(s2).length}</span>
                <span className="text-[13px] font-bold text-slate-800 text-right tabular-nums">{(s2.totalQty || 0).toLocaleString('th-TH')}</span>
                <span className="text-[13px] text-slate-300 text-right">›</span>
              </button>
            );
          })}

          {list.length === 0 && (
            <div className="px-4 py-8 text-center text-[13px] text-slate-500">ไม่มีเอกสารตรงตามที่กรอง</div>
          )}
        </div>
      )}
    </div>
  );
}

function PDFDownloadButton({ sub }) {
  return <button onClick={() => openPDFPrint(sub)} className="flex items-center gap-1 px-2 py-1.5 text-xs bg-[#FEF2F2] hover:bg-[#FECACA] rounded text-[#B91C1C] font-medium"><Download size={12}/>PDF</button>;
}

// สถานะระบบบนหน้าแรก — แยกเป็น 2 การ์ด: ส่งขึ้น Drive / ซิงก์จาก POS
// ใช้หน้าตาเดียวกับแถวในกล่อง "เครื่องมือ" แตะเพื่อกางรายละเอียด
function LandingStatus({ onOpenManager, onOpenSync, db, onOpenDbSettings, onTestDb }) {
  const [st, setSt] = useState({ busy: true, err: '', d: null });
  const [open, setOpen] = useState('');
  const [dbTest, setDbTest] = useState({ busy: false, msg: '', ok: null });

  const load = useCallback(async () => {
    setSt(v => ({ ...v, busy: true, err: '' }));
    try {
      const res = await fetch('/api/sync-status');
      const j = await res.json();
      if (!res.ok || j.error) throw new Error(j.error || `HTTP ${res.status}`);
      setSt({ busy: false, err: '', d: j });
    } catch (e) { setSt({ busy: false, err: e.message, d: null }); }
  }, []);
  useEffect(() => { load(); }, [load]);

  const INK = { ok: '#2F5D50', stale: '#B45309', late: '#B91C1C', unknown: '#94A3B8' };
  const d = st.d;
  const up = d?.uploads;
  const reports = d?.reports || [];
  const fmt = (iso) => { try { return new Date(iso).toLocaleString('th-TH', { day:'numeric', month:'short', hour:'2-digit', minute:'2-digit' }); } catch { return '—'; } };
  const ago = (h) => h === null ? '—' : h < 1 ? 'เพิ่งดึง' : h < 24 ? `${h} ชม.ก่อน` : `${Math.floor(h / 24)} วันก่อน`;
  // ไม่มีวันที่ (บาง view ไม่ส่งมา) อย่าโชว์ Invalid Date
  const thDate = (v) => {
    if (!v) return '—';
    const d = new Date(String(v).slice(0, 10) + 'T00:00:00+07:00');
    return isNaN(d) ? '—' : d.toLocaleDateString('th-TH', { day:'numeric', month:'short' });
  };


  const shortRows = reports.filter(r => r.missing > 0).length;
  const driveState = st.err ? 'unknown' : !up ? 'unknown' : up.failed ? 'late' : up.pending ? 'stale' : 'ok';
  const syncState  = st.err ? 'unknown'
    : shortRows || reports.some(r => r.state === 'late') ? 'late'
    : reports.some(r => r.state === 'stale') ? 'stale'
    : reports.length ? 'ok' : 'unknown';

  const dbReady = !!(db?.url && db?.anonKey);
  const dbState = !dbReady ? 'late' : db.connection === 'error' ? 'late' : 'ok';

  const loading = st.busy && !d;
  const cards = [
    {
      id: 'db', icon: Cloud, title: 'เชื่อมต่อ Supabase', state: dbState,
      sub: !dbReady ? 'ยังไม่ได้ตั้งค่า — แอปทำงานไม่ได้'
        : db.connection === 'error' ? 'ต่อไม่ได้ — เช็คคีย์หรือชื่อตาราง'
        : db.productCount > 0 ? `ต่ออยู่ · สินค้า ${db.productCount.toLocaleString()} รายการ`
        : 'ต่ออยู่',
      badge: 0,
    },
    {
      id: 'drive', icon: Upload, title: 'ส่งขึ้น Drive', state: driveState,
      sub: st.err ? 'เช็คสถานะไม่ได้'
        : !up ? 'ยังไม่ได้เปิดใช้ — รัน sql/13 ก่อน'
        : up.failed ? `ส่งไม่สำเร็จ ${up.failed} ใบ`
        : up.pending ? `รอส่ง ${up.pending} ใบ`
        : 'ส่งครบแล้ว',
      badge: up ? (up.failed || up.pending || 0) : 0,
    },
    {
      id: 'sync', icon: Database, title: 'ซิงก์จาก POS', state: syncState,
      sub: st.err ? 'เช็คสถานะไม่ได้'
        : !reports.length ? 'ยังไม่มีรอบซิงก์'
        : shortRows ? `ขึ้นไม่ครบ ${shortRows} รายงาน`
        : reports.some(r => r.state !== 'ok') ? 'มีรายงานที่ค้าง'
        : `ครบทั้ง ${reports.length} รายงาน`,
      badge: shortRows || reports.filter(r => r.state !== 'ok').length,
    },
  ];

  return (
    <>
      <div className="flex items-center justify-between gap-2 mt-1 mb-2">
        <div className="text-[11px] font-bold tracking-wide text-slate-400">สถานะระบบ</div>
        <button onClick={load} disabled={st.busy}
          className="flex items-center gap-1 text-[11px] font-semibold text-slate-400 disabled:opacity-50">
          <RefreshCw size={11} className={st.busy ? 'animate-spin' : ''} />
          {loading ? 'กำลังเช็ค…' : d ? fmt(d.checkedAt) : 'เช็คใหม่'}
        </button>
      </div>

      <div className="bg-white rounded-2xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
        {cards.map((c, i) => {
          const busyCard = loading && c.id !== 'db';   // สถานะ Supabase รู้จากในเครื่อง ไม่ต้องรอ API
          const state = busyCard ? 'unknown' : c.state;
          const CIcon = busyCard ? RefreshCw : c.icon;
          const on = open === c.id;
          const ink = INK[state];
          return (
            <div key={c.id} style={{ borderTop: i ? '1px solid #F1F3F5' : 'none' }}>
              <button onClick={() => setOpen(on ? '' : c.id)}
                className="w-full flex items-center gap-3 px-3.5 text-left hover:bg-[#F6F7F8]"
                style={{ minHeight: 62 }}>
                <div className="rounded-lg flex items-center justify-center shrink-0"
                  style={{ width: 36, height: 36, background: `${ink}14`, color: ink }}>
                  <CIcon size={17} className={busyCard ? 'animate-spin' : ''} />
                </div>
                <div className="flex-1 min-w-0">
                  <div className="text-[13.5px] font-bold text-slate-800">{c.title}</div>
                  <div className="text-[11px] truncate" style={{ color: ink }}>{busyCard ? 'กำลังเช็ค…' : c.sub}</div>
                </div>
                {state === 'ok' && !busyCard && <CheckCircle2 size={15} className="shrink-0" style={{ color: ink }} />}
                {!busyCard && c.badge > 0 && (
                  <span className="shrink-0 rounded-full text-[11px] font-bold text-white tabular-nums px-2 py-0.5"
                    style={{ background: ink }}>{c.badge}</span>
                )}
                <ChevronRight size={15} className="shrink-0 text-slate-300"
                  style={{ transform: on ? 'rotate(90deg)' : 'none', transition: 'transform .15s' }} />
              </button>

              {on && c.id === 'db' && (
                <div className="px-3.5 pb-3.5 pt-1 space-y-2">
                  <div className="rounded-xl px-3 py-2.5 space-y-1.5" style={{ background: '#F6F7F8' }}>
                    {[['เซิร์ฟเวอร์', dbReady ? String(db.url).replace(/^https?:\/\//, '') : 'ยังไม่ได้ตั้ง'],
                      ['ตารางสินค้า', db?.tableName || '—'],
                      ['ตารางสต็อก', db?.stockTableName || '—'],
                      ['แหล่งข้อมูล', db?.dataSource === 'supabase' ? 'Supabase' : db?.dataSource === 'seed' ? 'ข้อมูลตัวอย่าง' : 'ยังไม่โหลด']].map(([k, v]) => (
                      <div key={k} className="flex items-center gap-2">
                        <span className="text-[11px] text-slate-500 shrink-0" style={{ width: 74 }}>{k}</span>
                        <span className="flex-1 min-w-0 truncate font-mono text-[11.5px] font-semibold text-slate-800">{v}</span>
                      </div>
                    ))}
                  </div>
                  {db?.lastSyncAt && <div className="text-[11px] text-slate-500">โหลดสินค้าล่าสุด {fmt(db.lastSyncAt)}</div>}
                  {dbTest.msg && (
                    <div className="rounded-xl px-3 py-2.5 text-[11.5px] leading-relaxed"
                      style={{ background: dbTest.ok ? '#F0F7F4' : '#FDF2F2', color: dbTest.ok ? '#2A5A55' : '#B91C1C' }}>{dbTest.msg}</div>
                  )}
                  <div className="flex gap-2">
                    <button disabled={dbTest.busy || !dbReady}
                      onClick={async () => {
                        setDbTest({ busy: true, msg: '', ok: null });
                        try { const r = await onTestDb?.(db); setDbTest({ busy: false, ok: true, msg: r || 'ต่อได้ อ่านข้อมูลได้ปกติ' }); }
                        catch (e) { setDbTest({ busy: false, ok: false, msg: e.message || 'ต่อไม่ได้' }); }
                      }}
                      className="flex-1 rounded-xl text-white text-[12.5px] font-bold disabled:opacity-50"
                      style={{ minHeight: 42, background: '#0F172A' }}>{dbTest.busy ? 'กำลังทดสอบ…' : 'ทดสอบการเชื่อมต่อ'}</button>
                    <button onClick={onOpenDbSettings}
                      className="rounded-xl border bg-white text-[12.5px] font-semibold text-slate-600 px-3"
                      style={{ minHeight: 42, borderColor: '#E4E6EA' }}>แก้ด้วยมือ</button>
                  </div>
                </div>
              )}

              {on && c.id === 'drive' && (
                <div className="px-3.5 pb-3.5 pt-1">
                  {st.err ? (
                    <div className="rounded-xl px-3 py-2.5 text-[11.5px] leading-relaxed" style={{ background: '#FDF2F2', color: '#B91C1C' }}>{st.err}</div>
                  ) : !up ? (
                    <div className="text-[11.5px] text-slate-500 leading-relaxed">ต้องรัน sql/13-drive-status.sql ก่อน จึงจะนับใบที่รอส่งได้</div>
                  ) : (
                    <>
                      <div className="grid grid-cols-3 gap-2">
                        {[{ k:'รอส่ง', v: up.pending, ink: up.pending ? '#B45309' : '#475569' },
                          { k:'ส่งแล้ว', v: up.uploaded ?? (up.count.uploaded + up.invoice.uploaded), ink: '#2A5A55' },
                          { k:'ไม่สำเร็จ', v: up.failed, ink: up.failed ? '#B91C1C' : '#475569' }].map(x => (
                          <div key={x.k} className="rounded-xl px-2.5 py-2" style={{ background: '#F6F7F8' }}>
                            <div className="text-[10px] text-slate-500">{x.k}</div>
                            <div className="text-[15px] font-bold tabular-nums" style={{ color: x.ink }}>{Number(x.v || 0).toLocaleString()}</div>
                          </div>
                        ))}
                      </div>
                      {up.lastAt && <div className="text-[11px] text-slate-500 mt-2">ส่งครั้งล่าสุด {fmt(up.lastAt)}</div>}
                      {(up.pending || up.failed) ? (
                        <button onClick={onOpenManager}
                          className="w-full rounded-xl text-white text-[12.5px] font-bold mt-2.5"
                          style={{ minHeight: 42, background: '#0F172A' }}>จัดการใน "บันทึกแล้ว"</button>
                      ) : null}
                    </>
                  )}
                </div>
              )}

              {on && c.id === 'sync' && (
                <div className="px-3.5 pb-3.5 pt-1">
                  {st.err ? (
                    <div className="rounded-xl px-3 py-2.5 text-[11.5px] leading-relaxed" style={{ background: '#FDF2F2', color: '#B91C1C' }}>{st.err}</div>
                  ) : !reports.length ? (
                    <div className="text-[11.5px] text-slate-500 leading-relaxed">ยังไม่มีรอบซิงก์ใน sync_log</div>
                  ) : (
                    <>
                      {Object.entries(reports.reduce((acc, r) => {
                        const b = String(r.branch || '1').replace(/^0+/, '') || '1';
                        (acc[b] = acc[b] || []).push(r); return acc;
                      }, {})).sort(([a], [b]) => a.localeCompare(b)).map(([br, list]) => (
                      <div key={br}>
                      <div className="flex items-center gap-2 mt-1 mb-0.5">
                        <span className="rounded px-1.5 py-0.5 text-[10px] font-bold"
                          style={{ background: BRANCH_SOFT[br] || '#F6F7F8', color: BRANCH_INK[br] || '#475569' }}>{branchName(br)}</span>
                        <span className="text-[10.5px] text-slate-400">{list.length} รายงาน</span>
                        {list.some(x => x.state !== 'ok') && (
                          <span className="text-[10.5px] font-bold" style={{ color: '#B45309' }}>· {list.filter(x => x.state !== 'ok').length} ต้องดู</span>
                        )}
                      </div>
                      <div className="divide-y" style={{ borderColor: '#F1F3F5' }}>
                        {list.map((r, ri) => {
                          const short = r.missing > 0;
                          return (
                            <div key={r.report + '|' + br + ri} className="py-2">
                              <div className="flex items-center gap-2">
                                <span className="rounded-full shrink-0" style={{ width: 6, height: 6, background: INK[r.state] || INK.unknown }} />
                                <span className="min-w-0 truncate text-[12px] font-semibold text-slate-800">{r.label || r.report}</span>
                                <span className="flex-1" />
                                <span className="text-[11.5px] tabular-nums shrink-0" style={{ color: short ? '#B91C1C' : '#2A5A55' }}>
                                  {r.rowsCsv !== null && r.rowsCsv !== r.rowsSent
                                    ? `${Number(r.rowsSent || 0).toLocaleString()} / ${r.rowsCsv.toLocaleString()} แถว`
                                    : `${Number(r.rowsSent ?? r.rows ?? 0).toLocaleString()} แถว`}
                                </span>
                              </div>
                              <div className="flex items-center gap-2 pl-3.5 mt-0.5 text-[11px] text-slate-500">
                                <span className="shrink-0">{r.dataDate ? 'ข้อมูลวันที่ ' + thDate(r.dataDate) : 'ภาพรวมทั้งร้าน'}</span>
                                {r.mode === 'auto' && <span className="shrink-0 rounded px-1.5 py-0.5 text-[10px] font-bold" style={{ background: '#F0F7F4', color: '#2A5A55' }}>auto</span>}
                                {r.mode === 'manual' && <span className="shrink-0 rounded px-1.5 py-0.5 text-[10px] font-bold" style={{ background: '#FFFBEB', color: '#B45309' }}>คนสั่ง</span>}
                                <span className="flex-1" />
                                <span className="shrink-0">ดึงเมื่อ {ago(r.hoursAgo)}</span>
                              </div>
                              {short && (
                                <div className="pl-3.5 mt-0.5 text-[11px] font-semibold" style={{ color: '#B91C1C' }}>
                                  ขาด {r.missing.toLocaleString()} แถว
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                      </div>
                      ))}
                      <button onClick={onOpenSync}
                        className="w-full rounded-xl border bg-white text-[12.5px] font-semibold text-slate-700 mt-2.5"
                        style={{ minHeight: 42, borderColor: '#E4E6EA' }}>ดูประวัติการซิงก์</button>
                    </>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </>
  );
}

const REPORT_LABEL_TH = {
  product_price: 'ราคาสินค้า', product_stock: 'สินค้าคงคลัง', vendor_info: 'ข้อมูลผู้ขาย',
  sale_report_bill: 'บิลขาย', sale_item: 'รายการสินค้าขาย',
};

function DataSyncView() {
  const [st, setSt] = useState({ busy: true, err: '', d: null });
  const [tab, setTab] = useState('now');

  const load = useCallback(async () => {
    setSt(v => ({ ...v, busy: true, err: '' }));
    try {
      const res = await fetch('/api/sync-status?runs=40');
      const j = await res.json();
      if (!res.ok || j.error) throw new Error(j.error || `HTTP ${res.status}`);
      setSt({ busy: false, err: '', d: j });
    } catch (e) { setSt({ busy: false, err: e.message, d: null }); }
  }, []);
  useEffect(() => {
    load();
    const t = setInterval(load, 60000);
    return () => clearInterval(t);
  }, [load]);

  const TONE = {
    ok:      { label: 'ปกติ',                   ink: '#2A5A55', soft: '#F0F7F4', line: '#DBE9E3' },
    stale:   { label: 'ไม่ได้ซิงก์นานผิดปกติ',  ink: '#B45309', soft: '#FFFBEB', line: '#F3E3C3' },
    late:    { label: 'ล้มเหลว',                ink: '#B91C1C', soft: '#FDF2F2', line: '#F3D5D5' },
    unknown: { label: 'ไม่มีข้อมูล', ink: '#475569', soft: '#F6F7F8', line: '#E4E6EA' },
  };
  const d = st.d;
  const head = TONE[d?.overall] || TONE.unknown;
  const fmtTime = (iso) => { try { return new Date(iso).toLocaleString('th-TH', { day:'numeric', month:'short', hour:'2-digit', minute:'2-digit' }); } catch { return '—'; } };
  const fmtDate = (v) => {
    if (!v) return '—';
    const d = new Date(String(v).slice(0, 10) + 'T00:00:00+07:00');
    return isNaN(d) ? '—' : d.toLocaleDateString('th-TH', { day:'numeric', month:'short', year:'2-digit' });
  };
  const ago = (h) => h === null ? '—' : h < 1 ? 'เพิ่งซิงก์' : h < 24 ? `${h} ชม.ก่อน` : `${Math.floor(h / 24)} วันก่อน`;

  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between gap-2">
        <div className="min-w-0">
          <h2 className="text-xl font-bold text-slate-800">สถานะข้อมูล POS</h2>
          <p className="text-[11.5px] text-slate-500">จาก sync_log — ที่ export ขึ้น Supabase ครบไหม</p>
        </div>
        <button onClick={load} disabled={st.busy}
          className="shrink-0 rounded-lg flex items-center justify-center border bg-white disabled:opacity-60"
          style={{ width: 40, height: 40, borderColor: '#E4E6EA', color: '#0F172A' }}>
          <RefreshCw size={16} className={st.busy ? 'animate-spin' : ''} />
        </button>
      </div>

      {st.err ? (
        <div className="rounded-2xl p-4 border" style={{ background: '#FDF2F2', borderColor: '#F3D5D5' }}>
          <div className="flex items-center gap-2 mb-1">
            <XCircle size={16} style={{ color: '#B91C1C' }} />
            <span className="text-[13px] font-bold" style={{ color: '#B91C1C' }}>เช็คสถานะไม่ได้</span>
          </div>
          <div className="text-[11.5px] leading-relaxed" style={{ color: '#B91C1C' }}>{st.err}</div>
        </div>
      ) : !d ? (
        <div className="bg-white rounded-2xl border border-[#E4E6EA] p-8 text-center text-[13px] text-slate-500">กำลังเช็ค…</div>
      ) : (
        <>
          <div className="rounded-2xl px-4 py-3.5 border" style={{ background: head.soft, borderColor: head.line }}>
            <div className="flex items-center gap-2">
              {d.overall === 'ok' ? <CheckCircle2 size={17} style={{ color: head.ink }} /> : <XCircle size={17} style={{ color: head.ink }} />}
              <span className="text-[14.5px] font-bold" style={{ color: head.ink }}>
                {d.overall === 'ok' ? 'ทุกรีพอร์ตซิงก์ปกติ'
                  : d.failedRuns ? `มีรอบที่ล้มเหลว ${d.failedRuns} รอบ`
                  : 'มีรีพอร์ตที่ต้องดู'}
              </span>
            </div>
            <div className="text-[11px] text-slate-500 mt-1 font-mono">เช็คเมื่อ {fmtTime(d.checkedAt)}</div>
          </div>

          <div className="flex gap-2 flex-wrap">
            {[{ id:'now', label:'รายรายงาน', n: d.reports.length },
              { id:'odd', label:'บิลที่ควรตรวจ', n: d.anomalies?.count ?? 0 },
              { id:'old', label:'สินค้าค้าง', n: d.stale?.count ?? 0 },
              { id:'runs', label:'ประวัติ', n: d.runs.length }].map(t => {
              const on = tab === t.id;
              return (
                <button key={t.id} onClick={() => setTab(t.id)}
                  className="rounded-xl flex items-center justify-center gap-1.5 text-[12px] font-bold border px-3"
                  style={{ minHeight: 42, flex: '1 1 44%', background: on ? '#0F172A' : '#fff', color: on ? '#fff' : '#475569', borderColor: on ? '#0F172A' : '#E4E6EA' }}>
                  {t.label}
                  <span className="rounded-full text-[11px] font-bold px-2 py-0.5"
                    style={{ background: on ? '#fff' : '#F6F7F8', color: on ? '#0F172A' : '#64748B' }}>{t.n}</span>
                </button>
              );
            })}
          </div>

          {tab === 'now' ? (
            d.reports.length === 0 ? (
              <div className="bg-white rounded-2xl border border-[#E4E6EA] p-8 text-center text-[13px] text-slate-500">ยังไม่มีรอบซิงก์ใน sync_log</div>
            ) : Object.entries(d.reports.reduce((acc, r) => {
              const b = String(r.branch || '1').replace(/^0+/, '') || '1';
              (acc[b] = acc[b] || []).push(r); return acc;
            }, {})).sort(([a], [b]) => a.localeCompare(b)).map(([br, list]) => (
            <div key={br} className="space-y-2">
              <div className="flex items-center gap-2 pt-1">
                <span className="rounded-lg px-2 py-1 text-[11.5px] font-bold"
                  style={{ background: BRANCH_SOFT[br] || '#F6F7F8', color: BRANCH_INK[br] || '#475569' }}>{branchName(br)}</span>
                <span className="text-[11px] text-slate-400">{list.length} รายงาน</span>
                {list.some(x => x.state !== 'ok') && (
                  <span className="text-[11px] font-bold" style={{ color: '#B45309' }}>· {list.filter(x => x.state !== 'ok').length} ต้องดู</span>
                )}
              </div>
              {list.map(r => {
              const t = TONE[r.state] || TONE.unknown;
              const short = r.rowsCsv !== null && r.missing > 0;
              return (
                <div key={r.report + '|' + br} className="bg-white rounded-2xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
                  <div className="flex items-center gap-2 px-3 py-2.5 border-b" style={{ background: t.soft, borderColor: '#ECEEF0' }}>
                    <span className="flex-1 min-w-0 truncate text-[13px] font-bold text-slate-900">{r.label || r.report}</span>
                    <span className="shrink-0 text-[10.5px] font-bold rounded-full px-2 py-1 bg-white" style={{ color: t.ink }}>{r.summary || t.label}</span>
                  </div>
                  <div className="p-3">
                    <div className="grid grid-cols-3 gap-2">
                      <div className="rounded-xl px-2.5 py-2" style={{ background: '#F6F7F8' }}>
                        <div className="text-[10px] text-slate-500">ข้อมูลวันที่</div>
                        <div className="text-[12.5px] font-bold text-slate-800">{r.dataDate ? fmtDate(r.dataDate) : 'ทั้งร้าน'}</div>
                      </div>
                      <div className="rounded-xl px-2.5 py-2" style={{ background: '#F6F7F8' }}>
                        <div className="text-[10px] text-slate-500">ซิงก์ล่าสุด</div>
                        <div className="text-[12.5px] font-bold" style={{ color: r.state === 'ok' ? '#0F172A' : t.ink }}>{ago(r.hoursAgo)}</div>
                      </div>
                      <div className="rounded-xl px-2.5 py-2" style={{ background: short ? '#FDF2F2' : '#F6F7F8' }}>
                        <div className="text-[10px]" style={{ color: short ? '#B91C1C' : '#64748B' }}>แถวที่เข้า</div>
                        <div className="text-[13px] font-bold tabular-nums" style={{ color: short ? '#B91C1C' : '#0F172A' }}>
                          {(r.rowsSent ?? r.rows ?? 0).toLocaleString()}
                          {r.rowsCsv !== null && <span className="text-[11px] font-semibold text-slate-400"> / {r.rowsCsv.toLocaleString()}</span>}
                        </div>
                      </div>
                    </div>

                    {short ? (
                      <div className="rounded-xl px-3 py-2.5 mt-2 text-[11.5px] leading-relaxed" style={{ background: '#FDF2F2', color: '#B91C1C' }}>
                        ขึ้นไม่ครบ — ไฟล์มี {r.rowsCsv.toLocaleString()} แถว เข้า Supabase {r.rowsSent.toLocaleString()} แถว ขาด {r.missing.toLocaleString()}
                      </div>
                    ) : r.rowsCsv !== null && (
                      <div className="flex items-center gap-1.5 mt-2 text-[11.5px]" style={{ color: '#2A5A55' }}>
                        <CheckCircle2 size={13} />ครบตามไฟล์ ({r.rowsCsv.toLocaleString()} แถว)
                      </div>
                    )}
                    {r.error && (
                      <div className="rounded-xl px-3 py-2.5 mt-2 text-[11.5px] leading-relaxed break-words" style={{ background: '#FDF2F2', color: '#B91C1C' }}>{r.error}</div>
                    )}
                    <div className="flex items-center gap-2 mt-2">
                      <span className="text-[11px] text-slate-400 font-mono">{fmtTime(r.lastRunAt)}</span>
                      {r.durationSec !== null && <span className="text-[11px] text-slate-400 tabular-nums">· {r.durationSec} วิ</span>}
                      <span className="flex-1" />
                      {r.mode && (
                        <span className="rounded px-1.5 py-0.5 text-[10px] font-bold"
                          style={{ background: r.mode === 'auto' ? '#F0F7F4' : '#FFFBEB', color: r.mode === 'auto' ? '#2A5A55' : '#B45309' }}>
                          {r.mode === 'auto' ? 'รันอัตโนมัติ' : 'คนสั่งเอง'}
                        </span>
                      )}
                    </div>
                  </div>
                </div>
              );
              })}
            </div>
            ))
          ) : tab === 'odd' ? (
            !d.anomalies ? (
              <div className="bg-white rounded-2xl border border-[#E4E6EA] p-8 text-center text-[13px] text-slate-500">ยังไม่ได้เปิดใช้การตรวจบิล</div>
            ) : d.anomalies.count === 0 ? (
              <div className="bg-white rounded-2xl border border-[#E4E6EA] p-8 text-center">
                <CheckCircle2 className="mx-auto mb-2" size={30} style={{ color: '#DBE9E3' }} />
                <div className="text-[13px] text-slate-500">ไม่มีบิลที่ต้องตรวจ</div>
              </div>
            ) : (
              <div className="bg-white rounded-2xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
                <div className="px-3 py-2.5 border-b text-[11.5px] text-slate-500 leading-relaxed" style={{ background: '#F6F7F8', borderColor: '#ECEEF0' }}>
                  ตัวเลขผิดปกติ — ระบบไม่แก้อะไรให้ แค่ชี้ให้ไปดูเอง
                </div>
                <div className="divide-y" style={{ borderColor: '#F6F7F8' }}>
                  {d.anomalies.rows.map((a, i) => (
                    <div key={i} className="px-3 py-2.5">
                      <div className="flex items-center gap-2">
                        <span className="flex-1 min-w-0 truncate font-mono text-[12px] font-bold text-slate-800">{a.ref || '—'}</span>
                        <span className="text-[12.5px] font-bold tabular-nums shrink-0" style={{ color: '#B91C1C' }}>{Number(a.value || 0).toLocaleString()}</span>
                      </div>
                      <div className="text-[11px] text-slate-500 mt-0.5">{a.reason} · {a.date} · {a.table}</div>
                    </div>
                  ))}
                </div>
              </div>
            )
          ) : tab === 'old' ? (
            !d.stale ? (
              <div className="bg-white rounded-2xl border border-[#E4E6EA] p-8 text-center text-[13px] text-slate-500">ยังไม่ได้เปิดใช้การตรวจสินค้าค้าง</div>
            ) : d.stale.count === 0 ? (
              <div className="bg-white rounded-2xl border border-[#E4E6EA] p-8 text-center">
                <div className="text-[13px] text-slate-500 leading-relaxed">ยังไม่มีข้อมูล<br /><span className="text-[11.5px] text-slate-400">ต้องเก็บสถิติอย่างน้อย 30 วันก่อน</span></div>
              </div>
            ) : (
              <div className="bg-white rounded-2xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
                <div className="px-3 py-2.5 border-b text-[11.5px] text-slate-500 leading-relaxed" style={{ background: '#F6F7F8', borderColor: '#ECEEF0' }}>
                  ไม่ถูกอัปเดตเกิน 30 วัน — อาจเลิกขายแล้วแต่ยังค้างในระบบ
                </div>
                <div className="divide-y" style={{ borderColor: '#F6F7F8' }}>
                  {d.stale.rows.map((x, i) => (
                    <div key={i} className="px-3 py-2.5 flex items-center gap-2">
                      <div className="flex-1 min-w-0">
                        <div className="text-[12.5px] font-semibold text-slate-800 truncate">{x.name || '(ไม่มีชื่อ)'}</div>
                        <div className="font-mono text-[10.5px] text-slate-400 truncate">{x.code} · {x.table}</div>
                      </div>
                      <span className="text-[11.5px] tabular-nums shrink-0" style={{ color: '#B45309' }}>{x.days} วัน</span>
                    </div>
                  ))}
                </div>
              </div>
            )
          ) : (
            <div className="bg-white rounded-2xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
              <div className="divide-y" style={{ borderColor: '#F6F7F8' }}>
                {d.runs.map(r => {
                  const bad = r.status && r.status !== 'ok';
                  const short = r.missing > 0;
                  return (
                    <div key={r.id} className="px-3 py-2.5" style={{ background: bad ? '#FDF2F2' : '#fff' }}>
                      <div className="flex items-center gap-2">
                        <span className="rounded-full shrink-0" style={{ width: 6, height: 6, background: bad || short ? '#B91C1C' : '#2F5D50' }} />
                        <span className="flex-1 min-w-0 truncate text-[12px] font-bold text-slate-800">{REPORT_LABEL_TH[r.report] || r.report}</span>
                        <span className="text-[11px] text-slate-500 tabular-nums shrink-0">{(r.rowsSent ?? 0).toLocaleString()} แถว</span>
                      </div>
                      <div className="flex items-center gap-2 mt-1 pl-3.5">
                        <span className="text-[11px] text-slate-500 font-mono">{fmtTime(r.runAt)}</span>
                        {r.durationSec !== null && <span className="text-[11px] text-slate-400 tabular-nums">{r.durationSec} วิ</span>}
                        {short && <span className="text-[11px] font-semibold" style={{ color: '#B91C1C' }}>ขาด {r.missing.toLocaleString()}</span>}
                        {bad && <span className="text-[11px] font-bold" style={{ color: '#B91C1C' }}>{r.status}</span>}
                      </div>
                      {r.error && <div className="text-[11px] mt-1 pl-3.5 leading-relaxed break-words" style={{ color: '#B91C1C' }}>{r.error}</div>}
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}

// หน้าตั้งค่าโฟลเดอร์ Drive — เมนูของผู้จัดการ
function DriveSettingsView({ currentUser }) {
  const [v, setV] = useState(() => ({ ...CLOUD_SETTINGS }));
  const [saved, setSaved] = useState('');

  useEffect(() => { loadCloudSettings().then(cs => setV({ ...cs })); }, []);

  const FIELDS = [
    { k: 'drive_folder_stock_count',  label: 'นับสต็อก (stock_count)' },
    { k: 'drive_folder_stock_adjust', label: 'ปรับยอด (stock_adjustment)' },
    { k: 'drive_folder_manual',       label: 'บันทึกมือ (manual_record)' },
    { k: 'drive_folder_purchase',     label: 'บิลซื้อ (purchase_bill)' },
  ];

  const save = async () => {
    setSaved('saving');
    try {
      const r = await fetch('/api/config', {
        method: 'PATCH', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ settings: v, by: currentUser?.name || '' }),
      });
      const j = await r.json();
      if (!r.ok || j.error) throw new Error(j.error);
      CLOUD_SETTINGS = { ...CLOUD_SETTINGS, ...v };
      setSaved('ok');
    } catch { setSaved('err'); }
  };

  const missing = FIELDS.filter(f => !v[f.k]).length;

  return (
    <div className="space-y-3">
      <div>
        <h2 className="text-xl font-bold text-slate-800">ตั้งค่า</h2>
        <p className="text-[11.5px] text-slate-500">สาขาและโฟลเดอร์ Drive — ตั้งครั้งเดียวทุกเครื่องใช้ร่วม</p>
      </div>

      <div className="bg-white rounded-2xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
        <div className="px-3.5 py-2.5 border-b text-[12px] font-bold text-slate-700" style={{ borderColor: '#E4E6EA', background: '#F8FAFC' }}>สาขา</div>
        <div className="p-3.5 space-y-2.5">
          <div className="grid grid-cols-2 gap-2.5">
            {['1', '2'].map(id => (
              <div key={id}>
                <label className="text-[11px] font-semibold text-slate-500">ชื่อสาขา {id}</label>
                <input value={v[`branch_${id}_name`] || ''} placeholder={`สาขา ${id}`}
                  onChange={e => { setV(o => ({ ...o, [`branch_${id}_name`]: e.target.value })); setSaved(''); }}
                  className="w-full mt-1 rounded-xl border px-3 text-[13px] text-slate-800 outline-none"
                  style={{ minHeight: 44, borderColor: '#E4E6EA', background: '#FAFBFB' }} />
              </div>
            ))}
          </div>
          <div>
            <label className="text-[11px] font-semibold text-slate-500">คอลัมน์สาขาในตาราง product_stock</label>
            <input value={v.branch_stock_column || ''} placeholder="เช่น สาขา — เว้นว่าง = อ่านยอดรวมทุกสาขา"
              onChange={e => { setV(o => ({ ...o, branch_stock_column: e.target.value })); setSaved(''); }}
              className="w-full mt-1 rounded-xl border px-3 font-mono text-[12px] text-slate-800 outline-none"
              style={{ minHeight: 44, borderColor: '#E4E6EA', background: '#FAFBFB' }} />
            <div className="text-[11px] text-slate-400 mt-1 leading-relaxed">
              ต้องตั้งให้ตรงกับชื่อคอลัมน์จริง ไม่งั้นการเทียบยอดจะใช้ยอดรวมทุกสาขา ทำให้ตัวเลขปรับผิด
            </div>
          </div>
          {v.branch_stock_column && (
            <div className="grid grid-cols-2 gap-2.5">
              {['1', '2'].map(id => (
                <div key={id}>
                  <label className="text-[11px] font-semibold text-slate-500">ค่าในคอลัมน์นั้น = สาขา {id}</label>
                  <input value={v[`branch_${id}_value`] || ''} placeholder={id}
                    onChange={e => { setV(o => ({ ...o, [`branch_${id}_value`]: e.target.value })); setSaved(''); }}
                    className="w-full mt-1 rounded-xl border px-3 font-mono text-[12px] text-slate-800 outline-none"
                    style={{ minHeight: 44, borderColor: '#E4E6EA', background: '#FAFBFB' }} />
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="bg-white rounded-2xl border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
        <div className="p-3.5 space-y-2.5">
          {missing > 0 && (
            <div className="rounded-xl px-3 py-2.5 text-[11.5px] leading-relaxed" style={{ background: '#FFFBEB', color: '#B45309' }}>
              ยังไม่ได้ตั้ง {missing} โฟลเดอร์ — ฟีเจอร์นั้นจะส่งไฟล์ขึ้น Drive ไม่ได้
            </div>
          )}
          {FIELDS.map(f => (
            <div key={f.k}>
              <label className="text-[11px] font-semibold text-slate-500">{f.label}</label>
              <input value={v[f.k] || ''} placeholder="วาง Folder ID"
                onChange={e => { setV(o => ({ ...o, [f.k]: e.target.value })); setSaved(''); }}
                className="w-full mt-1 rounded-xl border px-3 font-mono text-[12px] text-slate-800 outline-none"
                style={{ minHeight: 44, borderColor: '#E4E6EA', background: '#FAFBFB' }} />
            </div>
          ))}
          <div className="text-[11px] text-slate-400 leading-relaxed">
            Folder ID คือส่วนท้าย URL ของโฟลเดอร์ใน Drive — หลัง /folders/ ตัด ?usp=... ออก · เว้นว่าง = ฟีเจอร์นั้นส่งไฟล์ไม่ได้
          </div>
          <div className="flex items-center gap-2.5">
            <button onClick={save} disabled={saved === 'saving'}
              className="rounded-xl text-white text-[12.5px] font-bold px-4 disabled:opacity-60"
              style={{ minHeight: 42, background: '#0F172A' }}>{saved === 'saving' ? 'กำลังบันทึก…' : 'บันทึกค่ากลาง'}</button>
            <span className="text-[11.5px] font-semibold"
              style={{ color: saved === 'err' ? '#B91C1C' : saved === 'ok' ? '#2A5A55' : '#94A3B8' }}>
              {saved === 'ok' ? 'บันทึกแล้ว — ทุกเครื่องใช้ค่านี้' : saved === 'err' ? 'บันทึกไม่ได้ — ยังไม่ได้รัน sql/14' : ''}
            </span>
          </div>
        </div>
      </div>
    </div>
  );
}

function SavedUploadsView({ submissions, invSubs = [], onRefresh, subSync = {}, currentUser, supabaseConfig, onPatched, onInvPatched }) {
  const cache = useUploadLog();
  const [tab, setTab] = useState('pending');
  const [busyId, setBusyId] = useState(null);
  const [confirmId, setConfirmId] = useState(null);

  const KIND = {
    recorder:      { label: 'นับสินค้า', ink: '#2A5A55' },
    stock_compare: { label: 'เทียบยอด',  ink: '#5B3FA8' },
    invoice:       { label: 'บิลซื้อ',    ink: '#255771' },
  };
  const entryOf = (sub) => driveEntryOf(sub, cache);
  // ใบนับกับบิลซื้อขึ้นรายการเดียวกัน — ผู้จัดการดูที่เดียวจบ
  const approved = [
    ...submissions.filter(s => s.status === 'approved').map(s => ({ ...s, _kind: s.featureType || 'recorder' })),
    ...invSubs.filter(s => s.status === 'approved').map(s => ({ ...s, _kind: 'invoice' })),
  ];
  const isInv = (sub) => sub._kind === 'invoice';
  const done = approved.filter(s => entryOf(s)?.ok);
  const todo = approved.filter(s => !entryOf(s)?.ok);
  const list = (tab === 'pending' ? todo : done)
    .slice().sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt));

  const send = async (sub, force = false) => {
    setBusyId(sub.id); setConfirmId(null);
    const inv = isInv(sub);
    let entry;
    if (inv) {
      // บิลซื้อ → CSV สำหรับ import เข้า POS 2 ไฟล์ ชื่อไฟล์ = doc_no
      const base = sub.docNo || sub.id;
      const { headerCsv, impCsv } = buildInvoiceCsvPair(sub);
      const folderId = setting('drive_folder_purchase');
      const e1 = await driveUpload({ subId: sub.id, type: 'csv_header',
        filename: `${base}_bill_header.csv`, mimeType: 'text/csv',
        content: headerCsv, bom: true, folderId, force });
      const e2 = await driveUpload({ subId: sub.id, type: 'csv_imp',
        filename: `${base}_imp_data.csv`, mimeType: 'text/csv',
        content: impCsv, folderId, force });
      // ทั้งคู่ต้องขึ้นครบ ถือว่าสำเร็จ — ขาดไฟล์ใดไฟล์หนึ่ง import ไม่ได้
      entry = (e1.ok && e2.ok) ? { ...e1, filename: `${base}_bill_header.csv + _imp_data.csv` }
            : (e1.ok ? e2 : e1);
      if (e1.skipped && e2.skipped) entry.skipped = true;
    } else if (sub._kind === 'stock_compare') {
      // ปรับยอด → ต้องรู้ยอดคงเหลือปัจจุบันก่อน ถึงคำนวณ +/- ได้
      try {
        const { csv, missing, total } = await buildAdjustCsv(
          sub.data, supabaseConfig?.url, supabaseConfig?.anonKey, supabaseConfig?.stockTableName, sub.branch || '1');
        entry = await driveUpload({
          subId: sub.id, type: 'csv_adjust',
          filename: `${sub.docNo || sub.id}.txt`, mimeType: 'text/csv',
          content: csv, folderId: subFolderId(sub), force,
        });
        if (entry.ok && missing) entry.warn = `${missing} จาก ${total} รายการไม่พบในระบบ ปล่อยค่าปรับว่างไว้`;
      } catch (e) {
        entry = { ok: false, err: e.message || 'สร้างไฟล์ปรับยอดไม่ได้' };
      }
    } else {
      entry = await driveUpload({
        subId: sub.id, type: 'csv_count',
        filename: `${sub.docNo || sub.id}.csv`, mimeType: 'text/csv',
        content: buildRecorderCsv(sub.data), folderId: subFolderId(sub), force,
      });
    }
    if (!entry.skipped) {
      const updated = await recordDriveResult(sub, entry, currentUser?.name,
        inv ? '/api/invoice-submission' : '/api/submission');
      if (updated) (inv ? onInvPatched : onPatched)?.(updated);
    }
    setBusyId(null);
  };

  const fmt = (iso) => { try { return new Date(iso).toLocaleString('th-TH', { day:'numeric', month:'short', hour:'2-digit', minute:'2-digit' }); } catch { return ''; } };

  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between gap-2">
        <div className="min-w-0">
          <h2 className="text-xl font-bold text-slate-800">บันทึกแล้ว</h2>
          <p className="text-[11.5px] text-slate-500">ส่งขึ้น Drive ที่นี่ที่เดียว — ระบบจำว่าใบไหนขึ้นแล้ว</p>
        </div>
        <button onClick={onRefresh} disabled={subSync.busy}
          className="shrink-0 rounded-lg flex items-center justify-center border bg-white disabled:opacity-60"
          style={{ width: 40, height: 40, borderColor: '#E4E6EA', color: '#0F172A' }}>
          <RefreshCw size={16} className={subSync.busy ? 'animate-spin' : ''} />
        </button>
      </div>

      <div className="flex gap-2">
        {[{ id:'pending', label:'ยังไม่ขึ้น', n: todo.length }, { id:'uploaded', label:'ขึ้นแล้ว', n: done.length }].map(t => {
          const on = tab === t.id;
          return (
            <button key={t.id} onClick={() => setTab(t.id)}
              className="flex-1 rounded-xl flex items-center justify-center gap-2 text-[12.5px] font-bold border"
              style={{ minHeight: 42, background: on ? '#0F172A' : '#fff', color: on ? '#fff' : '#475569', borderColor: on ? '#0F172A' : '#E4E6EA' }}>
              {t.label}
              <span className="rounded-full text-[11px] font-bold px-2 py-0.5"
                style={{ background: on ? '#fff' : '#F6F7F8', color: on ? '#0F172A' : '#64748B' }}>{t.n}</span>
            </button>
          );
        })}
      </div>

      {list.length === 0 ? (
        <div className="bg-white rounded-xl border border-[#E4E6EA] p-8 text-center">
          <Upload className="mx-auto text-[#E4E6EA] mb-2" size={34} />
          <div className="text-[13px] text-slate-500">{tab === 'pending' ? 'ขึ้น Drive ครบทุกใบแล้ว' : 'ยังไม่มีใบที่ขึ้น Drive'}</div>
        </div>
      ) : list.map(sub => {
        const k = KIND[sub._kind] || KIND.recorder;
        const res = entryOf(sub);
        const busy = busyId === sub.id;
        const failed = res && !res.ok;
        return (
          <div key={sub.id} className="bg-white rounded-2xl overflow-hidden border"
            style={{ borderColor: failed ? '#F3D5D5' : res?.ok ? '#E4E6EA' : '#CBD5E1', borderWidth: res?.ok ? 1 : 1.5 }}>
            <div className="flex items-center gap-2 px-3 py-2.5 border-b"
              style={{ background: failed ? '#FDF2F2' : '#F6F7F8', borderColor: '#ECEEF0' }}>
              <span className="text-[10.5px] font-bold bg-white rounded-md px-2 py-1" style={{ color: k.ink }}>{k.label}</span>
              <span className="flex-1 min-w-0 truncate font-mono font-bold text-[13.5px] text-slate-900">{sub.docNo || sub.id}</span>
              <span className="text-[10px] font-bold px-1.5 py-0.5 rounded shrink-0"
                style={{ background: BRANCH_SOFT[String(sub.branch || '1')], color: BRANCH_INK[String(sub.branch || '1')] }}>
                {branchName(sub.branch || '1')}
              </span>
            </div>

            <div className="p-3">
              {!res || (!res.ok && !failed) ? null : null}
              {res?.ok ? (
                <>
                  <div className="flex items-center gap-2 mb-2.5">
                    <CheckCircle2 size={15} style={{ color: '#2F5D50' }} />
                    <span className="text-[12px] font-bold" style={{ color: '#2A5A55' }}>ขึ้น Drive แล้ว</span>
                    <span className="font-mono text-[11px] text-slate-400">{fmt(res.at)}</span>
                  </div>
                  <div className="rounded-xl px-3 py-2.5 text-[11.5px] leading-relaxed border"
                    style={{ background: '#F0F7F4', borderColor: '#DBE9E3', color: '#2A5A55' }}>
                    <div className="font-semibold break-all">{res.filename}</div>
                    <div className="text-[11px] text-slate-500 mt-0.5">{isInv(sub) ? `${sub.vendorName || 'ไม่ระบุผู้ขาย'} · ฿${Number(sub.netTotal||0).toLocaleString()}` : `${sub.counter || sub.staffName || 'ทีมงาน'} · ${sub.data?.length || 0} รายการ`}</div>
                  </div>
                  <div className="flex gap-2 mt-2.5">
                    {res.link && (
                      <a href={res.link} target="_blank" rel="noopener noreferrer"
                        className="flex-1 rounded-xl border bg-white flex items-center justify-center text-[12.5px] font-semibold text-slate-700"
                        style={{ minHeight: 42, borderColor: '#E4E6EA' }}>เปิดใน Drive</a>
                    )}
                    {confirmId === sub.id ? (
                      <div className="flex items-center gap-1.5">
                        <button onClick={() => send(sub, true)}
                          className="rounded-xl text-white text-[12px] font-bold px-3" style={{ minHeight: 42, background: '#0F172A' }}>ส่งซ้ำจริง</button>
                        <button onClick={() => setConfirmId(null)}
                          className="rounded-xl text-[12px] font-semibold text-slate-600 px-3" style={{ minHeight: 42, background: '#F6F7F8' }}>ยกเลิก</button>
                      </div>
                    ) : (
                      <button onClick={() => setConfirmId(sub.id)} disabled={busy}
                        className="rounded-xl text-[12px] font-semibold px-3 disabled:opacity-60"
                        style={{ minHeight: 42, minWidth: 96, border: '1px dashed #CBD5E1', background: '#FAFBFB', color: '#94A3B8' }}>
                        {busy ? 'กำลังส่ง…' : 'ส่งซ้ำ'}
                      </button>
                    )}
                  </div>
                  {confirmId === sub.id && (
                    <div className="text-[11px] mt-2 leading-relaxed" style={{ color: '#B45309' }}>
                      ใบนี้มีไฟล์อยู่บน Drive แล้ว — ส่งซ้ำจะได้ไฟล์ชื่อเดิมเพิ่มอีกใบ
                    </div>
                  )}
                </>
              ) : failed ? (
                <>
                  <div className="flex items-center gap-2 mb-2">
                    <XCircle size={15} style={{ color: '#B91C1C' }} />
                    <span className="text-[12px] font-bold" style={{ color: '#B91C1C' }}>ส่งไม่สำเร็จ</span>
                    <span className="text-[11px] text-slate-400">ลองแล้ว {res.tries || 1} ครั้ง</span>
                  </div>
                  <div className="rounded-xl px-3 py-2.5 text-[11.5px] leading-relaxed" style={{ background: '#FDF2F2', color: '#B91C1C' }}>
                    ยังไม่มีไฟล์ขึ้นไป — {res.err || 'เชื่อมต่อ Drive ไม่ได้'}
                  </div>
                  <button onClick={() => send(sub)} disabled={busy}
                    className="w-full rounded-xl text-white text-[14px] font-bold mt-2.5 flex items-center justify-center gap-2 disabled:opacity-60"
                    style={{ minHeight: 46, background: '#0F172A' }}>
                    {busy ? <RefreshCw size={15} className="animate-spin" /> : <Upload size={15} />}
                    {busy ? 'กำลังส่ง…' : 'ลองส่งอีกครั้ง'}
                  </button>
                </>
              ) : (
                <>
                  <div className="flex items-center gap-2 mb-2.5">
                    <span className="w-2 h-2 rounded-full" style={{ background: '#94A3B8' }} />
                    <span className="text-[12px] font-bold text-slate-600">ยังไม่ขึ้น Drive</span>
                    <span className="text-[11.5px] text-slate-400">{isInv(sub) ? `${sub.vendorName || 'ไม่ระบุผู้ขาย'} · ${sub.itemCount || 0} รายการ` : `${sub.counter || sub.staffName || 'ทีมงาน'} · ${sub.data?.length || 0} รายการ`}</span>
                  </div>
                  <div className="flex gap-2">
                    <button onClick={() => send(sub)} disabled={busy}
                      className="flex-1 rounded-xl text-white text-[14px] font-bold flex items-center justify-center gap-2 disabled:opacity-60"
                      style={{ minHeight: 46, background: '#2F5D50' }}>
                      {busy ? <RefreshCw size={15} className="animate-spin" /> : <Upload size={15} />}
                      {busy ? 'กำลังส่ง…' : 'ส่งขึ้น Drive'}
                    </button>
                    {!isInv(sub) && (
                      <button onClick={() => openPDFPrint(sub)}
                        className="rounded-xl text-[12.5px] font-bold border"
                        style={{ minHeight: 46, minWidth: 70, background: '#FDF2F2', borderColor: '#F3D5D5', color: '#B91C1C' }}>PDF</button>
                    )}
                  </div>
                </>
              )}
            </div>
          </div>
        );
      })}

    </div>
  );
}

function ManagerInboxView({ submissions, onReview, onDelete, feature, onRefresh, subSync = {}, invSubs = [], onReviewInvoice, onDeleteInvoice, branch = 'all' }) {
  const [selected, setSelected] = useState(null);
  const [reviewNote, setReviewNote] = useState('');
  const [tab, setTab] = useState('pending');
  const [expandedId, setExpandedId] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [rejectError, setRejectError] = useState('');
  const [invSel, setInvSel] = useState(null);
  const [invNote, setInvNote] = useState('');
  const [invBusy, setInvBusy] = useState(false);
  const [invErr, setInvErr] = useState('');
  const [invOpenId, setInvOpenId] = useState(null);

  // ป้ายบอกที่มาของใบ
  const KIND = {
    recorder:      { label: 'นับสินค้า',  soft: '#EAF1F0', line: '#B6D0CC', ink: '#2A5A55' },
    stock_compare: { label: 'นับเทียบยอด', soft: '#F3EDFF', line: '#D6C6FF', ink: '#5B21B6' },
    invoice:       { label: 'บันทึกบิล',   soft: '#EAF0F4', line: '#B9CFDC', ink: '#255771' },
  };

  // รวมใบนับ + บิล เรียงตามเวลาส่ง
  const items = [
    ...submissions.map(s => ({ kind: s.featureType || 'recorder', at: s.submittedAt, status: s.status, s })),
    ...invSubs.map(s => ({ kind: 'invoice', at: s.submittedAt, status: s.status, s })),
  ].filter(it => it.status === tab)
   .sort((a, b) => new Date(b.at) - new Date(a.at));

  const counts = (st) => [
    ...submissions.filter(s => s.status === st),
    ...invSubs.filter(s => s.status === st),
  ].length;

  const filtered = submissions.filter(s => s.status === tab);
  const handleApprove = () => { onReview(selected.id,'approved',reviewNote); setSelected(null); setReviewNote(''); setRejectError(''); };
  const handleReject = () => { if(!reviewNote.trim()){setRejectError('กรุณาใส่เหตุผลก่อนส่งกลับ');return;} onReview(selected.id,'rejected',reviewNote); setSelected(null); setReviewNote(''); setRejectError(''); };

  const STAT = {
    pending:  { label: 'รอรีวิว',    soft: '#FFFBEB', line: '#FDE68A', ink: '#B45309' },
    approved: { label: 'อนุมัติแล้ว', soft: '#F0FDF4', line: '#BBF7D0', ink: '#15803D' },
    rejected: { label: 'ส่งกลับ',    soft: '#FEF2F2', line: '#FECACA', ink: '#B91C1C' },
  };

  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between gap-2">
        <h2 className="text-xl font-bold text-slate-800">กล่องขาเข้า</h2>
        {onRefresh && (
          <button onClick={onRefresh} disabled={subSync.busy} title="ดึงใบล่าสุดจากเซิร์ฟเวอร์"
            className="shrink-0 rounded-lg flex items-center justify-center border bg-white text-slate-500"
            style={{ width: 38, height: 38, borderColor: '#E4E6EA' }}>
            <RefreshCw size={15} className={subSync.busy ? 'animate-spin' : ''} />
          </button>
        )}
      </div>
      {subSync.err && <div className="text-[11px] font-semibold" style={{ color: '#B91C1C' }}>ดึงใบไม่สำเร็จ: {subSync.err}</div>}



      <div className="bg-white border rounded-xl flex overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
        {['pending', 'approved', 'rejected'].map(k => {
          const cnt = counts(k), on = tab === k;
          return (
            <button key={k} onClick={() => setTab(k)} className="flex-1 py-2.5 text-[12px] font-bold border-b-2"
              style={on ? { color: STAT[k].ink, background: STAT[k].soft, borderColor: STAT[k].ink }
                        : { color: '#64748B', borderColor: 'transparent' }}>
              {STAT[k].label}{cnt > 0 && <span className="tabular-nums"> {cnt}</span>}
            </button>
          );
        })}
      </div>

      {items.length === 0 ? (
        <div className="bg-white rounded-xl px-4 py-9 text-center" style={{ border: '1px dashed #CBD5E1' }}>
          <Inbox className="mx-auto text-[#E4E6EA] mb-2" size={34} />
          <div className="text-[13px] text-slate-500">ไม่มีใบ{STAT[tab].label}</div>
        </div>
      ) : (
        <div className="space-y-2">
          {items.map(it => {
            const s = it.s;
            const kd = KIND[it.kind] || KIND.recorder;
            const st = STAT[s.status] || STAT.pending;

            // ── บิลรออนุมัติ ────────────────────────────────
            if (it.kind === 'invoice') {
              const open = invOpenId === s.id;
              return (
                <div key={s.id} className="bg-white rounded-xl border overflow-hidden" style={{ borderColor: st.line }}>
                  <div className="px-3 py-2 flex items-center gap-2 border-b" style={{ background: st.soft, borderColor: st.line }}>
                    <span className="text-[10px] font-bold px-2 py-0.5 rounded-full shrink-0"
                      style={{ background: kd.soft, color: kd.ink }}>{kd.label}</span>
                    <span className="text-[11.5px] font-bold tabular-nums" style={{ color: st.ink }}>{s.docNo}</span>
                    <span className="ml-auto text-[11px] font-semibold text-slate-700 truncate">{s.keyedBy}</span>
                  </div>

                  <div className="p-3 space-y-2.5">
                    <div>
                      <div className="text-[13.5px] font-bold text-slate-800">{s.vendorName || '(ไม่ระบุผู้ขาย)'}</div>
                      <div className="text-[10.5px] text-slate-400 tabular-nums mt-0.5">
                        บิลเลขที่ {s.invoiceNo || '—'} · {s.invoiceDate || '—'}
                      </div>
                    </div>

                    <div className="flex gap-2">
                      <div className="flex-1 rounded-lg px-2.5 py-2 text-center" style={{ background: '#F6F7F8' }}>
                        <div className="text-[10px] text-slate-500">รายการ</div>
                        <div className="text-base font-bold text-slate-800 tabular-nums">{s.itemCount}</div>
                      </div>
                      <div className="flex-1 rounded-lg px-2.5 py-2 text-center" style={{ background: '#EAF0F4' }}>
                        <div className="text-[10px]" style={{ color: '#255771' }}>ยอดสุทธิ</div>
                        <div className="text-base font-bold tabular-nums" style={{ color: '#255771' }}>฿{Number(s.netTotal||0).toLocaleString()}</div>
                      </div>
                    </div>

                    {s.status !== 'pending' && s.reviewNote && (
                      <div className="rounded-lg p-2.5 border" style={{ background: st.soft, borderColor: st.line }}>
                        <div className="text-[10px] font-bold" style={{ color: st.ink }}>หมายเหตุที่คุณเขียน</div>
                        <div className="text-[11.5px] mt-0.5 leading-relaxed" style={{ color: st.ink }}>{s.reviewNote}</div>
                      </div>
                    )}
                    {s.status === 'approved' && s.postedAt && (
                      <div className="text-[11px] font-semibold" style={{ color: '#15803D' }}>บันทึกเข้าระบบแล้ว</div>
                    )}

                    {s.status === 'pending' && (
                      <button onClick={() => { setInvSel(s); setInvNote(''); setInvErr(''); }}
                        className="w-full text-white font-bold text-[14px] rounded-xl"
                        style={{ minHeight: 46, background: '#0F172A' }}>เปิดรีวิว →</button>
                    )}

                    <div className="flex flex-wrap gap-1.5 pt-2 border-t" style={{ borderColor: '#F6F7F8' }}>
                      <button onClick={() => setInvOpenId(open ? null : s.id)}
                        className="flex items-center gap-1 px-2.5 rounded-lg text-[11px] font-semibold text-slate-600"
                        style={{ minHeight: 40, background: '#F6F7F8' }}>
                        <FileSpreadsheet size={12} />{open ? 'ซ่อนรายการ' : `รายการ (${s.lines?.length || 0})`}
                      </button>
                      {onDeleteInvoice && (
                        confirmDelete === s.id ? (
                          <div className="flex items-center gap-1 ml-auto">
                            <span className="text-[10.5px] font-semibold" style={{ color: '#B91C1C' }}>ลบใบนี้?</span>
                            <button onClick={() => { onDeleteInvoice(s.id); setConfirmDelete(null); }}
                              className="px-2.5 rounded-lg text-[10.5px] font-bold text-white" style={{ minHeight: 40, background: '#B91C1C' }}>ลบ</button>
                            <button onClick={() => setConfirmDelete(null)}
                              className="px-2.5 rounded-lg text-[10.5px] font-semibold text-slate-600" style={{ minHeight: 40, background: '#F6F7F8' }}>ยกเลิก</button>
                          </div>
                        ) : (
                          <button onClick={() => setConfirmDelete(s.id)} title="ลบใบนี้"
                            className="ml-auto rounded-lg flex items-center justify-center"
                            style={{ width: 40, height: 40, background: '#FEF2F2', color: '#B91C1C' }}><Trash2 size={14} /></button>
                        )
                      )}
                    </div>

                    {open && (
                      <div className="rounded-lg divide-y max-h-56 overflow-y-auto" style={{ background: '#F6F7F8', borderColor: '#E4E6EA' }}>
                        {(s.lines || []).map((d, i) => (
                          <div key={i} className="flex items-center gap-2 px-2.5 py-2" style={{ borderColor: '#E4E6EA' }}>
                            <div className="flex-1 min-w-0">
                              <div className="text-[12px] font-semibold text-slate-800 truncate">{d.description || '(ไม่มีชื่อ)'}</div>
                              <div className="text-[10px] text-slate-400 tabular-nums truncate">{d.barcode || 'ไม่มีบาร์โค้ด'}</div>
                            </div>
                            <div className="text-[13px] font-bold text-slate-800 tabular-nums shrink-0">
                              {d.qty ?? '—'}<span className="text-[10px] font-normal text-slate-400 ml-1">×{Number(d.price_ea||0).toLocaleString()}</span>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              );
            }

            // ── ใบนับ ────────────────────────────────────────
            const open = expandedId === s.id;
            return (
              <div key={s.id} className="bg-white rounded-xl border overflow-hidden" style={{ borderColor: st.line }}>
                <div className="px-3 py-2 flex items-center gap-2 border-b" style={{ background: st.soft, borderColor: st.line }}>
                  <span className="text-[10px] font-bold px-2 py-0.5 rounded-full shrink-0"
                    style={{ background: kd.soft, color: kd.ink }}>{kd.label}</span>
                  <span className="text-[11.5px] font-bold tabular-nums" style={{ color: st.ink }}>{s.docNo || '—'}</span>
                  <span className="text-[10px] font-bold px-1.5 py-0.5 rounded shrink-0"
                    style={{ background: BRANCH_SOFT[String(s.branch || '1')], color: BRANCH_INK[String(s.branch || '1')] }}>
                    {branchName(s.branch || '1')}
                  </span>
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
                      style={{ minHeight: 40, background: '#F6F7F8' }}>
                      <FileSpreadsheet size={12} />{open ? 'ซ่อนรายการ' : `รายการ (${s.data?.length || 0})`}
                    </button>
                    <PDFDownloadButton sub={s} />
                    {confirmDelete === s.id ? (
                      <div className="flex items-center gap-1 ml-auto">
                        <span className="text-[10.5px] font-semibold" style={{ color: '#B91C1C' }}>ลบใบนี้?</span>
                        <button onClick={() => { onDelete(s.id); setConfirmDelete(null); }}
                          className="px-2.5 rounded-lg text-[10.5px] font-bold text-white" style={{ minHeight: 40, background: '#B91C1C' }}>ลบ</button>
                        <button onClick={() => setConfirmDelete(null)}
                          className="px-2.5 rounded-lg text-[10.5px] font-semibold text-slate-600" style={{ minHeight: 40, background: '#F6F7F8' }}>ยกเลิก</button>
                      </div>
                    ) : (
                      <button onClick={() => setConfirmDelete(s.id)} title="ลบใบนี้"
                        className="ml-auto rounded-lg flex items-center justify-center"
                        style={{ width: 40, height: 40, background: '#FEF2F2', color: '#B91C1C' }}><Trash2 size={14} /></button>
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

      {invSel && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-end sm:items-center justify-center p-3">
          <div className="bg-white rounded-2xl w-full max-w-md max-h-[92dvh] overflow-y-auto">
            <div className="px-4 py-3 border-b flex justify-between items-start gap-2 sticky top-0 bg-white" style={{ borderColor: '#E4E6EA' }}>
              <div className="min-w-0">
                <h3 className="font-bold text-slate-800 text-[15px]">รีวิวบิลซื้อ</h3>
                <p className="text-[11px] text-slate-500 truncate">{invSel.docNo} · {invSel.keyedBy}</p>
              </div>
              <button onClick={() => setInvSel(null)} className="shrink-0 rounded-lg flex items-center justify-center text-slate-500"
                style={{ width: 40, height: 40, background: '#F6F7F8' }}><X size={17} /></button>
            </div>

            <div className="p-4 space-y-3">
              <div className="rounded-xl p-3" style={{ background: '#F6F7F8' }}>
                <div className="text-[13.5px] font-bold text-slate-800">{invSel.vendorName || '(ไม่ระบุผู้ขาย)'}</div>
                <div className="text-[11px] text-slate-500 tabular-nums mt-1">บิลเลขที่ {invSel.invoiceNo || '—'} · {invSel.invoiceDate || '—'}</div>
              </div>

              <div className="flex gap-2">
                <div className="flex-1 rounded-xl px-3 py-2.5 text-center" style={{ background: '#F6F7F8' }}>
                  <div className="text-[10px] text-slate-500">รายการ</div>
                  <div className="text-xl font-bold text-slate-800 tabular-nums">{invSel.itemCount}</div>
                </div>
                <div className="flex-1 rounded-xl px-3 py-2.5 text-center" style={{ background: '#EAF0F4' }}>
                  <div className="text-[10px]" style={{ color: '#255771' }}>ยอดสุทธิ</div>
                  <div className="text-xl font-bold tabular-nums" style={{ color: '#255771' }}>฿{Number(invSel.netTotal||0).toLocaleString()}</div>
                </div>
              </div>

              {(() => {
                const miss = (invSel.lines||[]).filter(d => !d.barcode).length;
                if (!miss) return null;
                return (
                  <div className="rounded-xl p-3 border text-[11.5px] leading-relaxed" style={{ background: '#FFFBEB', borderColor: '#FDE68A', color: '#B45309' }}>
                    {miss} รายการยังไม่มีบาร์โค้ด — อนุมัติได้ แต่ช่องบาร์โค้ดจะว่างในระบบ
                  </div>
                );
              })()}

              <div className="bg-white border rounded-xl overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
                <div className="px-3 py-2 border-b text-[11.5px] font-bold text-slate-600" style={{ background: '#F6F7F8', borderColor: '#EEF0F3' }}>
                  รายการในบิล · {(invSel.lines||[]).length}
                </div>
                <div className="max-h-56 overflow-y-auto divide-y" style={{ borderColor: '#F6F7F8' }}>
                  {(invSel.lines||[]).map((d,i) => (
                    <div key={i} className="px-3 py-2 flex items-center gap-2">
                      <div className="flex-1 min-w-0">
                        <div className="text-[12.5px] font-semibold text-slate-800 truncate">{d.description || '(ไม่มีชื่อ)'}</div>
                        <div className="text-[10px] tabular-nums truncate" style={{ color: d.barcode ? '#94A3B8' : '#B45309' }}>{d.barcode || 'ไม่มีบาร์โค้ด'}</div>
                      </div>
                      <div className="text-right shrink-0">
                        <div className="text-[13px] font-bold text-slate-800 tabular-nums">{d.qty ?? '—'}</div>
                        <div className="text-[10px] text-slate-400 tabular-nums">฿{Number(d.total||d.amount||0).toLocaleString()}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div>
                <label className="text-[12px] font-semibold text-slate-700 mb-1.5 block">
                  หมายเหตุถึงพนักงาน <span style={{ color: '#B91C1C' }}>(บังคับถ้าส่งกลับ)</span>
                </label>
                <textarea value={invNote} onChange={e => { setInvNote(e.target.value); setInvErr(''); }} rows={2}
                  className="w-full px-3 py-2 border rounded-lg outline-none text-[13px] resize-none"
                  style={invErr ? { borderColor: '#B91C1C', background: '#FEF2F2' } : { borderColor: '#E2E8F0' }} />
                {invErr && (
                  <div className="text-[11px] mt-1 flex items-center gap-1" style={{ color: '#B91C1C' }}>
                    <XCircle size={12} />{invErr}
                  </div>
                )}
              </div>

              <div className="flex gap-2">
                <button disabled={invBusy}
                  onClick={async () => {
                    if (!invNote.trim()) { setInvErr('กรุณาใส่เหตุผลก่อนส่งกลับ'); return; }
                    setInvBusy(true);
                    try { await onReviewInvoice(invSel.id, 'rejected', invNote); setInvSel(null); setInvNote(''); }
                    catch (e) { setInvErr(e.message); }
                    setInvBusy(false);
                  }}
                  className="flex-1 text-white rounded-xl font-bold text-[13.5px] flex items-center justify-center gap-1.5 disabled:opacity-60"
                  style={{ minHeight: 50, background: '#B91C1C' }}><ThumbsDown size={16} />ส่งกลับแก้</button>
                <button disabled={invBusy}
                  onClick={async () => {
                    setInvBusy(true);
                    try { await onReviewInvoice(invSel.id, 'approved', invNote); setInvSel(null); setInvNote(''); }
                    catch (e) { setInvErr(e.message); }
                    setInvBusy(false);
                  }}
                  className="flex-1 text-white rounded-xl font-bold text-[13.5px] flex items-center justify-center gap-1.5 disabled:opacity-60"
                  style={{ minHeight: 50, background: '#15803D' }}>
                  {invBusy ? <RefreshCw size={16} className="animate-spin" /> : <ThumbsUp size={16} />}อนุมัติ + บันทึก
                </button>
              </div>
              <div className="text-[10.5px] text-slate-400 text-center leading-relaxed">อนุมัติแล้วระบบจะเขียนลง bill_header และ imp_data ให้ทันที</div>
            </div>
          </div>
        </div>
      )}

      {selected && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-end sm:items-center justify-center p-3">
          <div className="bg-white rounded-2xl w-full max-w-md max-h-[92dvh] overflow-y-auto">
            <div className="px-4 py-3 border-b flex justify-between items-start gap-2 sticky top-0 bg-white" style={{ borderColor: '#E4E6EA' }}>
              <div className="min-w-0">
                <h3 className="font-bold text-slate-800 text-[15px]">รีวิวใบนับ</h3>
                <p className="text-[11px] text-slate-500 truncate">
                  {selected.docNo ? selected.docNo + ' · ' : ''}{selected.counter}
                </p>
              </div>
              <button onClick={() => setSelected(null)} className="shrink-0 rounded-lg flex items-center justify-center text-slate-500"
                style={{ width: 40, height: 40, background: '#F6F7F8' }}><X size={17} /></button>
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
                  หมายเหตุถึงพนักงาน
                </label>
                <textarea value={reviewNote} onChange={e => { setReviewNote(e.target.value); setRejectError(''); }} rows={2}
                  placeholder=""
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
        </div>
        <div className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-[10.5px] font-semibold border shrink-0"
          style={isSupabaseReady
            ? { background: '#EAF1F0', borderColor: '#B6D0CC', color: '#2A5A55' }
            : { background: '#FEF2F2', borderColor: '#FECACA', color: '#B91C1C' }}>
          <Cloud size={11} />{isSupabaseReady ? 'เซิร์ฟเวอร์พร้อม' : 'ยังไม่ตั้งค่า'}
        </div>
      </div>

      {pendingCount > 0 && (
        <button onClick={() => setView('inbox')}
          className="w-full rounded-xl p-3.5 flex items-center gap-3 text-left border-2"
          style={{ background: '#FFFBEB', borderColor: '#FDE68A' }}>
          <div className="text-white p-2 rounded-lg shrink-0" style={{ background: '#B45309' }}><Inbox size={19} /></div>
          <div className="flex-1 min-w-0">
            <div className="text-[14px] font-bold" style={{ color: '#B45309' }}>{pendingCount} ใบรอคุณรีวิว</div>
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

function CompareStockView({ submissions, supabaseConfig, compareState, setCompareState, branch = 'all' }) {
  const [cmpOpen, setCmpOpen] = useState({});   // การ์ดเปรียบเทียบ เริ่มต้นพับไว้
  const { selectedSub, compareData, loading, loadProgress, error, driveSaving, driveResult } = compareState;
  const set = (patch) => setCompareState(prev => ({ ...prev, ...patch }));
  const approvedSubs = submissions.filter(s => s.status === 'approved');
  const { url: sbUrl, anonKey: sbKey, stockTableName } = supabaseConfig;

  const fetchAndCompare = async (sub) => {
    const now = new Date().toISOString();
    set({ selectedSub: sub, loading: true, error: '', compareData: [], driveResult: null, compareAt: now });
    // เทียบยอดต้องคิดเป็น "ชิ้น" ที่รหัสหลัก — สต็อกใน POS มีก้อนเดียวต่อตระกูล
    // ใบเก่าก่อนมีตัวคูณ ไม่มี masterCode/pieces → ถือว่าเป็นชิ้น ตัวคูณ 1
    const groupedByBarcode = {};
    sub.data.forEach(d => {
      const key = d.masterCode || d.barcode;
      const pieces = d.pieces != null ? Number(d.pieces) : (Number(d.qty) || 0) * (Number(d.stockRatio) || 1);
      if (!groupedByBarcode[key]) {
        groupedByBarcode[key] = { barcode: key, scannedCode: d.barcode, productName: d.productName,
          unit: (Number(d.stockRatio) || 1) > 1 ? 'ชิ้น' : (d.unit || ''),
          qty: 0, scannedAt: d.scannedAt, locations: [], notFound: !!d.notFound };
      }
      const g = groupedByBarcode[key];
      g.qty += pieces;
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
        // ใบไหนนับสาขาไหน อ่านยอดของสาขานั้น (ถ้ายังไม่ตั้งคอลัมน์สาขา จะอ่านรวมเหมือนเดิม)
        const rows = await sbFetch(sbUrl, sbKey, table, `${colCode}=in.(${inList})&select=${colSel}${stockBranchFilter(sub.branch || '1')}`);   // codes = รหัสหลักแล้ว
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

  // barcode,adjust_stock — ต้องมี + ให้ชัด · ไม่พบในระบบปล่อยว่าง
  const buildSimpleCSV = () => compareData.map(d => {
    const a = d.adjustStock;
    return `${d.barcode},${a == null ? '' : a > 0 ? '+' + a : a}`;
  }).join('\r\n');
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

  const saveToDrive = async (force = false) => {
    set({ driveSaving: true, driveResult: null });
    const res = await driveUpload({
      subId: selectedSub.id, type: 'csv_adjust',
      filename: `${selectedSub.docNo || selectedSub.id}.txt`, mimeType: 'text/csv',
      content: buildSimpleCSV(), isBase64: false,
      folderId: setting('drive_folder_stock_adjust'), force,
    });
    set({ driveResult: res.ok ? { ok: true, link: res.link, skipped: res.skipped } : { ok: false, err: res.err }, driveSaving: false });
  };


  if (!sbUrl || !sbKey) return (
    <div className="space-y-4"><div><h2 className="text-2xl font-bold text-slate-800">เปรียบเทียบสต็อก</h2></div>
      <div className="bg-[#FFFBEB] border border-[#FDE68A] rounded-xl p-4"><AlertCircle className="text-[#B45309] mb-2" size={24}/><div className="font-semibold text-[#B45309]">ยังไม่ได้เชื่อมต่อฐานข้อมูล</div></div>
    </div>
  );

  return (
    <div className="space-y-4">
      <div><h2 className="text-2xl font-bold text-slate-800">เปรียบเทียบสต็อก</h2></div>
      {approvedSubs.length === 0 ? <div className="bg-white rounded-xl p-8 text-center border border-[#E4E6EA]"><ArrowLeftRight className="mx-auto text-[#E4E6EA] mb-2" size={40}/><div className="text-slate-500">ยังไม่มีรายการที่อนุมัติแล้ว</div></div> : (
        <div className="bg-white rounded-xl border border-[#E4E6EA] p-4 space-y-2">
          <div className="text-sm font-semibold text-slate-700 mb-2">เลือกใบที่จะเปรียบเทียบ</div>
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
            <div className="rounded-xl p-3 text-xs" style={{ background: '#EAF0F4', border: '1px solid #B9CFDC' }}>
              <div className="font-semibold mb-1" style={{ color: '#255771' }}>ช่วงเวลาที่นำมาคิด</div>
              <div className="font-mono" style={{ color: '#255771' }}>
                {selectedSub.startedAt ? new Date(selectedSub.startedAt).toLocaleString('th-TH',{day:'2-digit',month:'short',year:'2-digit',hour:'2-digit',minute:'2-digit'}) : '—'}
                {' → '}
                {new Date(selectedSub.submittedAt).toLocaleString('th-TH',{day:'2-digit',month:'short',year:'2-digit',hour:'2-digit',minute:'2-digit'})}
              </div>
            </div>
          )}
          {compareState.debugInfo && null}
          <div className="flex flex-wrap gap-2">
            <button onClick={downloadCompareCSV} className="flex items-center gap-1 px-3 py-2 bg-[#F6F7F8] hover:bg-[#F6F7F8] text-[#0F172A] rounded-lg text-sm font-medium"><Download size={14}/>CSV</button>
            <button onClick={downloadCompareExcel} className="flex items-center gap-1 px-3 py-2 bg-[#EAF0F4] hover:bg-[#EAF0F4] text-[#255771] rounded-lg text-sm font-medium"><FileSpreadsheet size={14}/>Excel</button>
            <button onClick={() => saveToDrive()} disabled={driveSaving} className="flex items-center gap-1 px-3 py-2 bg-[#EAF1F0] hover:bg-[#EAF1F0] text-[#2A5A55] rounded-lg text-sm font-medium disabled:opacity-60">{driveSaving?<RefreshCw size={14} className="animate-spin"/>:<Upload size={14}/>}ขึ้น Drive</button>
            {driveResult?.ok && <span className="flex items-center gap-1 text-xs text-[#15803D]"><CheckCircle2 size={12}/>อัพโหลดแล้ว{driveResult.link&&<a href={driveResult.link} target="_blank" rel="noopener noreferrer" className="underline ml-1">เปิด</a>}</span>}
            {driveResult?.err && <span className="text-xs text-[#B91C1C]">Error: {driveResult.err}</span>}
          </div>
          <div className="grid grid-cols-3 gap-2 text-center text-xs">
            {[{label:'พบในระบบ',v:compareData.filter(d=>d.found).length,c:'emerald'},{label:'ไม่พบ',v:compareData.filter(d=>!d.found).length,c:'amber'},{label:'ส่วนต่าง≠0',v:compareData.filter(d=>d.adjustStock!==null&&d.adjustStock!==0).length,c:'red'}].map(x=>(
              <div key={x.label} className={`rounded-lg p-2 border ${x.c==='emerald'?'bg-[#EAF1F0] border-[#EAF1F0] text-[#2A5A55]':x.c==='amber'?'bg-[#FFFBEB] border-[#FFFBEB] text-[#B45309]':'bg-[#FEF2F2] border-[#FEF2F2] text-[#B91C1C]'}`}><div className="text-lg font-bold">{x.v}</div><div className="opacity-75">{x.label}</div></div>
            ))}
          </div>
          <div className="space-y-2">
            {compareData.map((d,i) => {
              const openRow = !!cmpOpen[i];
              const zero = d.adjustStock === 0;
              const na = d.adjustStock === null;
              const tone = na
                ? { soft:'#F6F7F8', line:'#E4E6EA', ink:'#64748B', label:'ไม่มียอดระบบ' }
                : zero
                  ? { soft:'#EAF1F0', line:'#B6D0CC', ink:'#2A5A55', label:'ตรงกัน' }
                  : d.adjustStock > 0
                    ? { soft:'#EAF0F4', line:'#B9CFDC', ink:'#255771', label:'นับได้มากกว่า' }
                    : { soft:'#FEF2F2', line:'#FECACA', ink:'#B91C1C', label:'นับได้น้อยกว่า' };
              return (
                <div key={i} className="bg-white rounded-xl border overflow-hidden" style={{ borderColor: d.notFound ? '#FDE68A' : tone.line }}>
                  <button onClick={() => setCmpOpen(prev => ({ ...prev, [i]: !prev[i] }))}
                    className="w-full text-left px-3 py-2.5 flex items-center gap-2.5"
                    style={{ background: d.notFound ? '#FFFBEB' : tone.soft, minHeight: 56 }}>
                    <div className="flex-1 min-w-0">
                      <div className="text-[13.5px] font-semibold text-slate-800 truncate">{d.productName}</div>
                      <div className="text-[10px] font-bold mt-0.5" style={{ color: d.notFound ? '#B45309' : tone.ink }}>
                        {d.notFound ? 'ไม่พบในระบบ' : tone.label}
                      </div>
                    </div>
                    <div className="text-right shrink-0">
                      <div className="text-[15px] font-bold tabular-nums" style={{ color: d.notFound ? '#B45309' : tone.ink }}>
                        {na ? '—' : (d.adjustStock > 0 ? '+' : '') + d.adjustStock}
                      </div>
                      <div className="text-[9.5px] text-slate-400">{openRow ? 'ซ่อน' : 'ดูรายละเอียด'}</div>
                    </div>
                  </button>

                  {openRow && (
                  <div className="p-3 space-y-2.5 border-t" style={{ borderColor: d.notFound ? '#FDE68A' : tone.line }}>
                    <div className="text-[10.5px] text-slate-400 tabular-nums" style={{ overflowWrap: 'anywhere' }}>
                      {d.barcode}{d.locations?.length > 0 ? ' · ' + d.locations.join(', ') : ''}
                    </div>

                    <div className="grid grid-cols-3 gap-2">
                      {[['นับได้', d.counted, '#0F172A'], ['ขาย', d.sale, '#B45309'], ['รับ', d.purchase, '#255771']].map(([k,v,c]) => (
                        <div key={k} className="rounded-lg px-2 py-1.5 text-center" style={{ background: '#F6F7F8' }}>
                          <div className="text-[10px] text-slate-500">{k}</div>
                          <div className="text-[15px] font-bold tabular-nums" style={{ color: c }}>{v}</div>
                        </div>
                      ))}
                    </div>

                    <div className="grid grid-cols-3 gap-2">
                      {[['ควรเหลือ', d.adjustedCount, '#0F172A'],
                        ['ยอดระบบ', d.stockAtSubmit ?? '—', d.stockAtSubmit == null ? '#94A3B8' : '#0F172A'],
                        ['ส่วนต่าง', na ? '—' : d.adjustStock, na ? '#94A3B8' : tone.ink]].map(([k,v,c]) => (
                        <div key={k} className="rounded-lg px-2 py-1.5 text-center border" style={{ background: '#fff', borderColor: '#EEF0F3' }}>
                          <div className="text-[10px] text-slate-500">{k}</div>
                          <div className="text-[15px] font-bold tabular-nums" style={{ color: c }}>{v}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

const REPORT_GROUPS = [
  { key: 'app', label: 'บันทึกในแอป' },
  { key: 'pos', label: 'ดึงจาก POS' },
];

const REPORT_TOPICS = [
  { id: 'count',   label: 'การนับ',        needDate: true,  group: 'app', edge: '#22C55E', hint: 'เลขที่ใบ · ผู้นับ · วันที่' },
  { id: 'invoice', label: 'บิลซื้อ',        needDate: true,  group: 'app', edge: '#22C55E', hint: 'เลขที่บิล · ผู้ขาย' },
  { id: 'stock',   label: 'สินค้าคงเหลือ',  needDate: false, group: 'pos', edge: '#F59E0B', hint: 'รหัสสินค้า · ประเภท' },
  { id: 'price',   label: 'ราคาสินค้า',     needDate: false, group: 'pos', edge: '#F59E0B', hint: 'ราคาขาย · ทุน · กำไร' },
  { id: 'in',      label: 'ซื้อเข้า',       needDate: true,  group: 'pos', edge: '#F59E0B', hint: 'เลขที่ใบรับ · ผู้ขาย' },
  { id: 'out',     label: 'ขายออก',        needDate: true,  group: 'pos', edge: '#F59E0B', hint: 'เลขที่บิลขาย · ลูกค้า' },
];

// หัวคอลัมน์ภาษาไทย — คีย์ตรงกับที่ /api/report ส่งกลับมา
const COL_LABEL = {
  branch: 'สาขา', barcode: 'บาร์โค้ด', master_code: 'รหัสหลัก (สต็อก)',
  doc_no: 'เลขที่เอกสาร', counted_at: 'วันที่นับ', counter_name: 'ผู้นับ', zone: 'โซน',
  product_code: 'รหัสสินค้า', name: 'ชื่อสินค้า', unit: 'หน่วย',
  qty: 'จำนวน', status: 'สถานะ',
  price: 'ราคาขาย', cost: 'ทุนเฉลี่ย', margin: 'กำไร/หน่วย', margin_pct: '% กำไร', category: 'ประเภท',
  file_name: 'ไฟล์', invoice_no: 'เลขที่บิล', invoice_date: 'วันที่บิล',
  vendor_name: 'ผู้ขาย', description: 'รายละเอียด', ea: 'ea', price_ea: 'ราคา/หน่วย',
  discount: 'ส่วนลด', amount: 'จำนวนเงิน', vat: 'ภาษี', total: 'รวม',
  on_hand: 'คงเหลือ', occurred_at: 'วันเวลา', kind: 'ประเภท', party: 'ผู้ขาย / ลูกค้า',
};

const isNumCol = (k) => ['qty','ea','price_ea','discount','amount','total','on_hand','price','cost','margin','margin_pct'].includes(k);
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
  const [branch, setBranch] = useState('all');   // รายงานการนับ/บิลซื้อ กรองสาขาได้
  const [truncated, setTruncated] = useState(false);   // ผลชนเพดาน 50,000 แถว
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
          branch: branch === 'all' ? null : branch,
          limit: 50000,
        }),
      });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || 'ดึงข้อมูลไม่สำเร็จ');
      setRows(j.rows || []);
      setCols(j.columns || (j.rows?.[0] ? Object.keys(j.rows[0]) : []));
      setTruncated(!!j.truncated);
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

  // ราคาต่อหน่วย/ทุน/กำไร/% เอามาบวกกันไม่ได้ — รวมได้แต่จำนวนและยอดเงิน
  const NO_SUM = ['price', 'cost', 'margin', 'margin_pct', 'price_ea'];
  const sums = useMemo(() => {
    const s = {};
    for (const k of cols) if (isNumCol(k) && !NO_SUM.includes(k)) s[k] = shown.reduce((a, r) => a + (Number(r[k]) || 0), 0);
    return s;
  }, [shown, cols]);

  function exportXlsx() {
    const header = cols.map(k => COL_LABEL[k] || k);
    const body = shown.map(r => cols.map(k => (isNumCol(k) ? (Number(r[k]) || 0) : fmtCell(k, r[k]))));
    const meta = [
      [`รายงาน: ${conf.label}`],
      conf.needDate ? [`ช่วงวันที่: ${from} ถึง ${to}`] : ['ทั้งหมด (ไม่จำกัดวันที่)'],
      [`สาขา: ${branch === 'all' ? 'ทุกสาขา' : branchName(branch)}`],
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
      `report-${topic}${branch === 'all' ? '' : '-สาขา' + branch}-${todayISO()}.xlsx`);
  }

  // แยกไฟล์ต่อสาขา — ส่งให้แต่ละสาขาดูของตัวเองได้ ไม่ต้องมาคัดเอง
  function exportPerBranch() {
    for (const b of BRANCHES()) {
      const part = shown.filter(r => String(r.branch || '1').replace(/^0+/, '') === b.id);
      if (!part.length) continue;
      const header = cols.map(k => COL_LABEL[k] || k);
      const body = part.map(r => cols.map(k => (isNumCol(k) ? (Number(r[k]) || 0) : fmtCell(k, r[k]))));
      const ws = XLSX.utils.aoa_to_sheet([
        [`รายงาน: ${conf.label} — ${b.name}`],
        conf.needDate ? [`ช่วงวันที่: ${from} ถึง ${to}`] : ['ทั้งหมด (ไม่จำกัดวันที่)'],
        [`ดึงเมื่อ: ${new Date().toLocaleString('th-TH')} · ${part.length} แถว`],
        [], header, ...body,
      ]);
      ws['!cols'] = cols.map(k => ({ wch: k === 'name' || k === 'description' ? 34 : 16 }));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, b.name.slice(0, 30));
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      downloadBlob(new Blob([out], { type: 'application/octet-stream' }),
        `report-${topic}-${b.name}-${todayISO()}.xlsx`);
    }
  }

  const [dl, setDl] = useState('');
  // ดาวน์โหลดทั้งหมดโดยไม่ผ่านตารางบนจอ — ใช้เมื่อข้อมูลเกินเพดานที่แสดงได้
  async function exportAll(kind) {
    setDl(kind); 
    try {
      const res = await fetch('/api/report', {
        method: 'POST', headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          topic, from: conf.needDate ? from : null, to: conf.needDate ? to : null,
          doc: doc || null, barcode: barcode || null, party: party || null,
          branch: branch === 'all' ? null : branch,
          limit: 0,   // ทั้งหมด
        }),
      });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || 'ดึงข้อมูลไม่สำเร็จ');
      const all = j.rows || [];
      const ks = j.columns || (all[0] ? Object.keys(all[0]) : []);
      const head = ks.map(k => COL_LABEL[k] || k);
      if (kind === 'csv') {
        const esc = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
        const lines = [head.map(esc).join(',')];
        for (const rw of all) lines.push(ks.map(k => esc(fmtCell(k, rw[k]))).join(','));
        downloadBlob(new Blob(['\uFEFF' + lines.join('\n')], { type: 'text/csv;charset=utf-8' }),
          `report-${topic}-ทั้งหมด-${todayISO()}.csv`);
      } else {
        const ws = XLSX.utils.aoa_to_sheet([
          [`รายงาน: ${conf.label}${branch === 'all' ? '' : ' — ' + branchName(branch)}`],
          conf.needDate ? [`ช่วงวันที่: ${from} ถึง ${to}`] : ['ทั้งหมด (ไม่จำกัดวันที่)'],
          [`ดึงเมื่อ: ${new Date().toLocaleString('th-TH')} · ${all.length} แถว`],
          [], head,
          ...all.map(rw => ks.map(k => (isNumCol(k) ? (Number(rw[k]) || 0) : fmtCell(k, rw[k])))),
        ]);
        ws['!cols'] = ks.map(k => ({ wch: k === 'name' || k === 'description' ? 34 : 16 }));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, conf.label.slice(0, 30));
        const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        downloadBlob(new Blob([out], { type: 'application/octet-stream' }),
          `report-${topic}-ทั้งหมด-${todayISO()}.xlsx`);
      }
    } catch (e) { setErr(e.message); }
    setDl('');
  }

  function exportCsv() {
    const esc = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
    const lines = [cols.map(k => esc(COL_LABEL[k] || k)).join(',')];
    for (const r of shown) lines.push(cols.map(k => esc(fmtCell(k, r[k]))).join(','));
    downloadBlob(new Blob(['\uFEFF' + lines.join('\n'), ], { type: 'text/csv;charset=utf-8' }),
      `report-${topic}${branch === 'all' ? '' : '-สาขา' + branch}-${todayISO()}.csv`);
  }

  const [menuOpen, setMenuOpen] = useState(false);

  return (
    <div className="space-y-4">

      {/* เมนูรายงาน — ปุ่ม 3 ขีด + แผงเลื่อนออกมา เหมือนเมนูฟีเจอร์อื่น */}
      {menuOpen && (
        <div className="fixed inset-0 z-50 flex" onClick={() => setMenuOpen(false)}>
          <div className="bg-white h-full flex flex-col" style={{ width: 268, maxWidth: '84vw', boxShadow: '2px 0 16px rgba(15,23,42,.18)' }}
            onClick={e => e.stopPropagation()}>
            <div className="px-4 py-4 border-b" style={{ borderColor: '#E4E6EA', background: '#F6F7F8' }}>
              <div className="text-[14px] font-bold text-slate-800">รายงาน</div>
              <div className="text-[11px] text-slate-500 mt-0.5">เลือกเรื่อง ใส่เงื่อนไข ส่งออกไฟล์</div>
            </div>
            <div className="flex-1 overflow-y-auto py-2">
              {REPORT_GROUPS.map(g => (
                <div key={g.key}>
                  <div className="px-4 py-2 text-[10px] font-bold tracking-wide text-slate-400">{g.label}</div>
                  {REPORT_TOPICS.filter(t => t.group === g.key).map(t => {
                    const on = topic === t.id;
                    return (
                      <button key={t.id} onClick={() => { setTopic(t.id); setRows(null); setColFilter({}); setMenuOpen(false); }}
                        className="w-full flex items-center gap-3 px-4 text-left"
                        style={{ minHeight: 52, background: on ? '#F6F7F8' : '#fff',
                                 borderLeft: on ? '3px solid #0F172A' : '3px solid transparent' }}>
                        <span className="shrink-0 rounded-sm" style={{ width: 4, height: 22, background: t.edge }} />
                        <span className="flex-1 min-w-0">
                          <span className="block text-[14px] font-semibold" style={{ color: on ? '#0F172A' : '#334155' }}>{t.label}</span>
                          <span className="block text-[10.5px] text-slate-400 truncate">{t.hint}</span>
                        </span>
                      </button>
                    );
                  })}
                </div>
              ))}
            </div>
          </div>
          <div className="flex-1" style={{ background: 'rgba(15,23,42,.45)' }}></div>
        </div>
      )}

      <div className="bg-white rounded-xl border border-[#E4E6EA] overflow-hidden">
        <div className="flex items-center gap-3 px-4 py-3 border-b border-[#E4E6EA]" style={{ background: '#F1F3F2' }}>
          <button onClick={() => setMenuOpen(true)} title="เมนูรายงาน"
            className="shrink-0 rounded-lg flex items-center justify-center border bg-white"
            style={{ width: 40, height: 40, borderColor: '#E4E6EA', color: '#0F172A' }}>
            <Menu size={20} />
          </button>
          <h2 className="text-xl font-bold text-slate-800 min-w-0 truncate flex-1">{conf.label}</h2>
          <span className={`text-[10.5px] font-bold rounded-md px-2 py-1 shrink-0 ${
            conf.group === 'app' ? 'text-[#15803D] bg-[#F0FDF4]' : 'text-[#92400E] bg-[#FFFBEB]'}`}>
            {conf.group === 'app' ? 'บันทึกในแอป' : 'จาก POS'}
          </span>
        </div>

        <div className="p-4 space-y-4">
        {conf.needDate && (
          <div className="grid grid-cols-2 gap-4">
            <label className="block min-w-0">
              <span className="text-xs font-medium text-slate-500">ตั้งแต่</span>
              <input type="date" value={from} onChange={e => setFrom(e.target.value)}
                className="mt-1 block w-full min-w-0 max-w-full border border-[#E4E6EA] rounded-lg px-2.5 py-2 text-[13px] appearance-none bg-white" />
            </label>
            <label className="block min-w-0">
              <span className="text-xs font-medium text-slate-500">ถึง</span>
              <input type="date" value={to} onChange={e => setTo(e.target.value)}
                className="mt-1 block w-full min-w-0 max-w-full border border-[#E4E6EA] rounded-lg px-2.5 py-2 text-[13px] appearance-none bg-white" />
            </label>
          </div>
        )}

        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          {conf.needDate && (
            <label className="block">
              <span className="text-xs font-medium text-slate-500">เลขที่เอกสาร</span>
              <input value={doc} onChange={e => setDoc(e.target.value)} placeholder="ทั้งหมด"
                className="mt-1 w-full border border-[#E4E6EA] rounded-lg px-3 py-2 text-sm" />
            </label>
          )}
          {(
            <label className="block">
              <span className="text-xs font-medium text-slate-500">สาขา</span>
              <div className="mt-1 flex rounded-lg border overflow-hidden" style={{ borderColor: '#E4E6EA' }}>
                {[{ id: 'all', label: 'ทุกสาขา' }, ...BRANCHES().map(b => ({ id: b.id, label: b.name }))].map((b, i) => {
                  const on = branch === b.id;
                  return (
                    <button key={b.id} onClick={() => setBranch(b.id)} className="flex-1 font-semibold text-[12px]"
                      style={{ minHeight: 38, borderLeft: i ? '1px solid #E4E6EA' : 'none',
                               background: on ? BRANCH_SOFT[b.id] : '#fff', color: on ? BRANCH_INK[b.id] : '#94A3B8' }}>
                      {b.label}
                    </button>
                  );
                })}
              </div>
            </label>
          )}
          <label className="block">
            <span className="text-xs font-medium text-slate-500">รหัสสินค้า / ชื่อสินค้า</span>
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
          {(topic === 'stock' || topic === 'price') && (
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
              {truncated && (
                <>
                  <span className="text-[11px] font-bold rounded-lg px-2 py-1" style={{ background: '#FFFBEB', color: '#B45309' }}>
                    แสดงได้ 50,000 แถว — มีมากกว่านี้
                  </span>
                  <button onClick={() => exportAll('xlsx')} disabled={!!dl}
                    className="px-3 py-1.5 rounded-lg text-xs font-bold text-white disabled:opacity-50 flex items-center gap-1"
                    style={{ background: '#B45309' }}>
                    {dl === 'xlsx' ? <RefreshCw size={12} className="animate-spin" /> : <Download size={12} />}
                    {dl === 'xlsx' ? 'กำลังดึงทั้งหมด…' : 'ดาวน์โหลดทั้งหมด'}
                  </button>
                </>
              )}
              <button onClick={exportCsv} disabled={!shown.length}
                className="px-3 py-1.5 rounded-lg text-xs font-medium border border-[#E4E6EA] text-slate-700 hover:bg-white disabled:opacity-40 flex items-center gap-1">
                <Download size={12} />CSV
              </button>
              {branch === 'all' && cols.includes('branch')
                && new Set(shown.map(r => String(r.branch || '1'))).size > 1 && (
                <button onClick={exportPerBranch}
                  className="px-3 py-1.5 rounded-lg text-xs font-medium border border-[#E4E6EA] text-slate-700 hover:bg-white flex items-center gap-1">
                  <FileSpreadsheet size={12} />Excel แยกสาขา
                </button>
              )}
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
            <div className="overflow-auto max-h-[70dvh]" style={{ WebkitOverflowScrolling: 'touch' }}>
              <table className="text-xs" style={{ minWidth: '100%' }}>
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
    { key: 'allow_recorder', label: 'นับสินค้า', soft: '#EAF1F0', line: '#B6D0CC', ink: '#2A5A55', main: '#35706A' },
    { key: 'allow_compare',  label: 'นับเทียบยอด', soft: '#F3EDFF', line: '#D6C6FF', ink: '#5B21B6', main: '#7C4DFF' },
    { key: 'allow_invoice',  label: 'บันทึกบิล', soft: '#EAF0F4', line: '#B9CFDC', ink: '#255771', main: '#2F6E90' },
  ];
  const DEPTS = ['คลังสินค้า', 'หน้าร้าน', 'บัญชี'];

  const load = async () => {
    setLoading(true); setErr('');
    try {
      const res = await fetch('/api/staff');
      const text = await res.text();
      let j; try { j = JSON.parse(text); }
      catch { throw new Error(`เซิร์ฟเวอร์ตอบไม่ใช่ JSON (HTTP ${res.status}) — ฟังก์ชัน /api/staff พัง ดู Logs บน Vercel · ${text.slice(0, 100)}`); }
      if (!res.ok || j.error) throw new Error(j.error || `HTTP ${res.status}`);
      setStaff(j.staff || []);
    } catch (e) { setErr(e.message); } finally { setLoading(false); }
  };
  useEffect(() => { load(); }, []);

  const post = async (body) => {
    setSaving(true); setErr('');
    try {
      const res = await fetch('/api/staff', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) });
      const text = await res.text();
      let j; try { j = JSON.parse(text); }
      catch { throw new Error(`เซิร์ฟเวอร์ตอบไม่ใช่ JSON (HTTP ${res.status}) · ${text.slice(0, 100)}`); }
      if (!res.ok || j.error) throw new Error(j.error || 'บันทึกไม่สำเร็จ');
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
      </div>

      <div className="bg-white border rounded-xl p-3 space-y-3" style={{ borderColor: '#E4E6EA' }}>
        <div>
          <div className="text-[11.5px] font-semibold text-slate-600 mb-1.5">ชื่อที่ขึ้นหน้าเลือกชื่อ</div>
          <input value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))}
            placeholder="สมชาย พ."
            className="w-full px-3 py-2.5 border rounded-lg outline-none focus:border-[#35706A] text-[15px]"
            style={{ borderColor: '#E2E8F0' }} />
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

    </div>
  );

  // ── รายชื่อ ─────────────────────────────────────────────────────
  return (
    <div className="space-y-3">
      <div className="flex items-start justify-between gap-2">
        <div>
          <h2 className="text-xl font-bold text-slate-800">พนักงาน</h2>
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
  const [gate, setGate] = useState('');
  const [gateErr, setGateErr] = useState('');
  const [gateOk, setGateOk] = useState(false);
  const [gateBusy, setGateBusy] = useState(false);

  const checkGate = async () => {
    if (!gate) return setGateErr('ใส่รหัสผ่านก่อน');
    setGateBusy(true); setGateErr('');
    try {
      const r = await fetch('/api/auth', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ pin: gate }),
      });
      const d = await r.json();
      if (!d.ok) { setGateErr(d.msg || 'รหัสผ่านไม่ถูกต้อง'); setGateBusy(false); return; }
      setGateOk(true); setGate('');
    } catch (e) { setGateErr('เชื่อมต่อไม่ได้: ' + e.message); }
    setGateBusy(false);
  };

  // ตั้งค่าเซิร์ฟเวอร์ต้องใส่รหัสผ่านก่อน — คนนอกแก้การเชื่อมต่อไม่ได้
  if (!gateOk) return (
    <div className="max-w-sm mx-auto pt-6">
      <div className="bg-white rounded-2xl border p-5 space-y-4" style={{ borderColor: '#E4E6EA' }}>
        <div className="flex items-center gap-3">
          <div className="rounded-xl flex items-center justify-center text-white shrink-0"
            style={{ width: 44, height: 44, background: '#0F172A' }}><Lock size={20} /></div>
          <div className="min-w-0">
            <div className="text-[16px] font-bold text-slate-800">ต้องใส่รหัสผ่าน</div>
          </div>
        </div>
        <input type="password" value={gate} autoFocus
          onChange={e => { setGate(e.target.value); setGateErr(''); }}
          onKeyDown={e => e.key === 'Enter' && checkGate()}
          placeholder="••••" maxLength={4}
          className="w-full px-3 py-3 border rounded-xl outline-none font-mono text-xl tracking-[0.4em] text-center"
          style={{ borderColor: gateErr ? '#B91C1C' : '#E4E6EA', background: gateErr ? '#FEF2F2' : '#fff' }} />
        {gateErr && (
          <div className="text-[11.5px] flex items-center gap-1.5" style={{ color: '#B91C1C' }}>
            <XCircle size={13} />{gateErr}
          </div>
        )}
        <button onClick={checkGate} disabled={gateBusy || !gate}
          className="w-full text-white font-bold text-[14.5px] rounded-xl disabled:opacity-40 flex items-center justify-center gap-2"
          style={{ minHeight: 50, background: '#0F172A' }}>
          {gateBusy ? <><RefreshCw size={16} className="animate-spin" />กำลังตรวจสอบ</> : 'เข้าหน้าตั้งค่า'}
        </button>
      </div>
    </div>
  );

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

function StepBar({ current, onGo, maxStep, onOpenSent, sentBadge = 0, sentActive = false }) {
  return (
    <div className="fixed bottom-0 left-0 right-0 bg-white border-t border-[#E4E6EA] z-20"
      style={{ paddingBottom: 'env(safe-area-inset-bottom)' }}>
      <div className="max-w-6xl mx-auto grid" style={{ gridTemplateColumns: onOpenSent ? 'repeat(4, minmax(0,1fr)) 1.15fr' : 'repeat(4, minmax(0,1fr))' }}>
        {STEPS.map((label, i) => {
          const n = i+1, done = current > n, active = !sentActive && current === n, reachable = n <= (maxStep ?? current);
          return (
            <button key={n} onClick={() => reachable && onGo && onGo(n)} disabled={!reachable}
              className="flex flex-col items-center justify-center gap-1 leading-none"
              style={{ minHeight: 58, paddingTop: 8, paddingBottom: 8,
                       borderTop: active ? '3px solid #2F6E90' : '3px solid transparent',
                       background: active ? '#EAF0F4' : '#fff',
                       color: active ? '#255771' : done ? '#2F6E90' : '#94A3B8',
                       opacity: reachable ? 1 : 0.5, cursor: reachable ? 'pointer' : 'default' }}>
              <span className="rounded-full flex items-center justify-center text-[11px] font-bold"
                style={{ width: 22, height: 22,
                         background: done ? '#2F6E90' : active ? '#fff' : '#F6F7F8',
                         border: active ? '2px solid #2F6E90' : done ? 'none' : '1px solid #E4E6EA',
                         color: done ? '#fff' : active ? '#255771' : '#94A3B8' }}>{done ? '✓' : n}</span>
              <span className="text-[10px] font-semibold whitespace-nowrap">{label}</span>
            </button>
          );
        })}
        {onOpenSent && (
          <button onClick={onOpenSent}
            className="relative flex flex-col items-center justify-center gap-1 leading-none border-l"
            style={{ minHeight: 58, paddingTop: 8, paddingBottom: 8, borderColor: '#E4E6EA',
                     borderTop: sentActive ? '3px solid #2F6E90' : '3px solid transparent',
                     background: sentActive ? '#EAF0F4' : '#fff',
                     color: sentActive ? '#255771' : '#64748B', cursor: 'pointer' }}>
            <Send size={17} />
            <span className="text-[10px] font-semibold whitespace-nowrap">บิลที่ส่งแล้ว</span>
            {sentBadge > 0 && (
              <span className="absolute text-white font-bold rounded-full text-center"
                style={{ top: 6, right: 10, minWidth: 16, height: 16, fontSize: 9.5, lineHeight: '16px', background: '#B91C1C' }}>{sentBadge}</span>
            )}
          </button>
        )}
      </div>
    </div>
  );
}

function DropZone({ onFiles, multiple, accept, children, capture, bare }) {
  const [drag, setDrag] = useState(false); const ref = useRef();
  const handle = fs => { if (fs?.length) onFiles(Array.from(fs)); };
  return (
    <div ref={ref} onClick={() => ref.current.querySelector('input')?.click()}
      onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
      onDrop={e=>{e.preventDefault();setDrag(false);handle(e.dataTransfer.files);}}
      className="cursor-pointer transition-colors text-center"
      style={bare
        ? { borderRadius: 12 }
        : { borderRadius: 12, padding: '30px 16px', background: '#fff',
            border: drag ? '1px dashed #2F6E90' : '1px dashed #CBD5E1',
            boxShadow: drag ? 'inset 0 0 0 3px #EAF0F4' : 'none' }}>
      <input type="file" multiple={multiple} accept={accept} {...(capture ? { capture: 'environment' } : {})} className="hidden" onChange={e=>handle(e.target.files)}/>
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

class ErrorBox extends React.Component {
  constructor(p) { super(p); this.state = { err: null }; }
  static getDerivedStateFromError(err) { return { err }; }
  render() {
    if (!this.state.err) return this.props.children;
    return (
      <div style={{ padding: 16 }}>
        <div style={{ background:'#fff', border:'1px solid #fecaca', borderRadius:12, padding:14, display:'flex', flexDirection:'column', gap:10 }}>
          <div style={{ fontSize:14.5, fontWeight:700, color:'#b91c1c' }}>หน้านี้แสดงไม่ได้</div>
          <div style={{ fontSize:11.5, color:'#b91c1c', fontFamily:"'IBM Plex Mono', monospace", wordBreak:'break-word', lineHeight:1.5 }}>
            {String(this.state.err?.message || this.state.err)}
          </div>
          <button onClick={() => this.setState({ err: null })}
            style={{ minHeight:46, borderRadius:11, border:'none', background:'#2f6e90', color:'#fff', fontFamily:'inherit', fontSize:14, fontWeight:700, cursor:'pointer' }}>ลองแสดงอีกครั้ง</button>
          <div style={{ fontSize:11, color:'#94a3b8', lineHeight:1.5 }}>ข้อมูลที่คีย์ไว้ยังอยู่ — กดปุ่มด้านบน หรือถ่ายภาพข้อความนี้ส่งให้ผู้ดูแล</div>
        </div>
      </div>
    );
  }
}

function InvoiceScannerModule({ supabaseConfig, currentUser, onOpenSent, onCloseSent, sentBadge = 0, sentActive = false, sentView = null, reviseTarget = null, onClearRevise }) {
  const [invDraft, setInvDraft] = useState({ busy: '', msg: '', err: '' });
  const [invOpen, setInvOpen] = useState({});    // พับ/กางรายการสินค้าต่อใบ
  const [prodOpen, setProdOpen] = useState({});  // พับ/กางสินค้าแต่ละตัว
  const maxStep = 4;   // เดินดูได้ทุกขั้นตั้งแต่ยังไม่อัปรูป — แต่ละขั้นบอกเองว่าต้องมีอะไรก่อน
  const setMaxStep = () => {};
  const [sendSt, setSendSt] = useState({ busy:false, done:null, err:'' });   // ส่งบิลให้ผู้จัดการ
  // มาจากปุ่ม "แก้ใบนี้ส่งใหม่" — ส่งแล้วได้เลขเดิม + R1
  const revise = reviseTarget;
  const [reviseLoaded, setReviseLoaded] = useState(null);   // เก็บ id ใบที่ยกข้อมูลมาแล้ว
  const clearRevise = () => { setReviseLoaded(null); onClearRevise?.(); };

  // แก้ใบเดิม: ยกข้อมูลเก่ามาวางในขั้นตรวจสอบเลย — แก้บาร์โค้ด/จำนวน/ราคาได้ทันที
  useEffect(() => {
    if (!revise?.data || reviseLoaded === revise.id) return;
    setReviseLoaded(revise.id);
    setInvoices([{ files: [], pagesData: [revise.data], pageStatus: ['done'], status: 'done', data: revise.data, error: null }]);
    if (revise.fileName) setFileName(revise.fileName);
    setStep(3);
  }, [revise, reviseLoaded]);

  // พนักงานไม่บันทึกลงระบบเอง — ส่งให้ผู้จัดการอนุมัติก่อน
  const sendInvoicesForReview = async () => {
    if (!currentUser) return setSendSt({ busy:false, done:null, err:'ยังไม่ได้เลือกชื่อ' });
    const ready = doneInvs.filter(inv => (inv.data?.products||[]).length > 0);
    if (!ready.length) return setSendSt({ busy:false, done:null, err:'ไม่มีใบที่อ่านสำเร็จ' });
    setSendSt({ busy:true, done:null, err:'' });
    const sent = [];
    try {
      for (const inv of ready) {
        const d = inv.data;
        const vs = vatSummary(d.products||[]);
        const res = await fetch('/api/invoice-submission', {
          method:'POST', headers:{ 'Content-Type':'application/json' },
          body: JSON.stringify({ submission: {
            keyedById: currentUser.id, keyedBy: currentUser.name,
            branch: String(currentUser.branch || '1'),
            reviseOf: revise?.id || null,
            deviceId: safeGet('deviceId','') || null,
            invoiceNo: d.invoice_no || null, invoiceDate: d.invoice_date || null,
            vendorName: d.vendor_name || null,
            fileName: fileName || d.invoice_no || null,
            header: d, lines: d.products || [], netTotal: vs.netTotal,
          }}),
        });
        const j = await res.json();
        if (!res.ok || j.error) throw new Error(j.error || ('HTTP ' + res.status));
        sent.push(j.submission?.docNo || '');
      }
      // ส่งแล้ว ร่างบนเซิร์ฟเวอร์ต้องหาย
      try { await fetch('/api/invoice-draft?keyed_by_id=' + encodeURIComponent(currentUser.id), { method:'DELETE' }); } catch {}
      setSendSt({ busy:false, done:sent, err:'' });
      clearRevise();   // ส่งแก้แล้ว รอบถัดไปเป็นบิลใหม่ตามปกติ
    } catch (e) { setSendSt({ busy:false, done:null, err:e.message }); }
  };
  const w = useWinWidth(), mob = w < 600;
  const [model, setModel] = useState('claude-sonnet-4-6');
  const [step, setStep] = useState(1);
  const [invoices, setInvoices] = useState([]);
  const [productFiles, setPFiles] = useState([]);
  const [selPFileIds, setSelPFIds] = useState(new Set());
  const [barcodeMap, setBMap] = useState({});
  const [scanResults, setSRes] = useState([]);
  const [scanning, setScanning] = useState(false);
  const [rematching, setRematching] = useState(false);
  const [fileName, setFileName] = useState('');
  const [sbSt, setSbSt] = useState(null);
  const [sbErr, setSbErr] = useState(null);
  const [selectedPages, setSelPages] = useState({});
  const [cfgSaved, setCfgSaved] = useState('');
  // ค่ากลางมาช้ากว่า mount — ทับค่าในเครื่องเมื่อโหลดเสร็จ
  const saveCentral = async (key, value) => {
    setCfgSaved('saving');
    try {
      const r = await fetch('/api/config', {
        method: 'PATCH', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key, value, by: currentUser?.name || '' }),
      });
      const j = await r.json();
      if (!r.ok || j.error) throw new Error(j.error);
      CLOUD_SETTINGS = { ...CLOUD_SETTINGS, [key]: value };
      setCfgSaved('ok');
    } catch { setCfgSaved('err'); }
  };
  const [nameStatus, setNameStatus] = useState(null);

  const { url: sbUrl, anonKey: sbKey } = supabaseConfig;


  // ร่างบิลขึ้นเซิร์ฟเวอร์ — คีย์ค้างที่เครื่องหนึ่ง ไปต่อเครื่องอื่นได้ (เก็บผลที่อ่านได้ ไม่เก็บรูป)
  const pushInvDraft = async () => {
    if (!currentUser) return;
    setInvDraft({ busy: 'up', msg: '', err: '' });
    try {
      const res = await fetch('/api/invoice-draft', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          keyed_by_id: currentUser.id, keyed_by: currentUser.name,
          device_id: safeGet('deviceId', '') || null, invoices,
        }),
      });
      const j = await res.json();
      if (!res.ok || j.error) throw new Error(j.error || 'HTTP ' + res.status);
      setInvDraft({ busy: '', msg: `บันทึกขึ้นเซิร์ฟเวอร์แล้ว ${j.saved} ใบ`, err: '' });
    } catch (e) { setInvDraft({ busy: '', msg: '', err: e.message }); }
  };

  const pullInvDraft = async () => {
    if (!currentUser) return;
    setInvDraft({ busy: 'down', msg: '', err: '' });
    try {
      const res = await fetch('/api/invoice-draft?keyed_by_id=' + encodeURIComponent(currentUser.id));
      const j = await res.json();
      if (!res.ok || j.error) throw new Error(j.error || 'HTTP ' + res.status);
      const list = j.invoices || [];
      if (list.length) { setInvoices(list); setStep(2); }
      setInvDraft({ busy: '', msg: list.length ? `ดึงร่างมา ${list.length} ใบ` : 'ไม่มีร่างบนเซิร์ฟเวอร์', err: '' });
    } catch (e) { setInvDraft({ busy: '', msg: '', err: e.message }); }
  };


  const HEADER_KEYS = ['invoice_no','invoice_date','vendor_name','vendor_tax_id','document_type','vendor_address','vendor_branch','vendor_no','price_type','_vendorFromDB','_vendorMatchBy'];
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
    if (sbUrl && sbKey) { try {
      const hit = await lookupVendorREST(sbUrl, sbKey, data.vendor_name, data.vendor_tax_id);
      if (hit) { data.vendor_no = hit.no; data._vendorFromDB = true; data._vendorMatchBy = hit.by; }
    } catch {} }
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
    let combined = [...kept,...newResults.flat()];
    // อ่านบาร์โค้ดได้แล้ว → ค้น product_price เอาชื่อจริงมาใช้จับคู่ (แม่นกว่าชื่อบนกล่อง)
    combined = await withSystemNames(combined);
    setSRes(combined); applyBarcodeMap(buildMap(combined)); setScanning(false);
  };

  // เติมชื่อจริงจากระบบให้ทุกแถวที่มีบาร์โค้ด
  const withSystemNames = async (rows) => {
    const need = rows.filter(r => r.barcode && r._sysName === undefined);
    if (!need.length) return rows;
    const found = await lookupBarcodes(supabaseConfig, need.map(r => r.barcode));
    return rows.map(r => {
      if (!r.barcode) return r;
      const hit = found.get(String(r.barcode).trim());
      // เก็บแค่ "เจอ/ไม่เจอ" — ชื่อในระบบไม่เอาไปจับคู่ (ชื่อบิลกับชื่อในระบบเขียนไม่เหมือนกัน)
      return { ...r, _inSystem: !!hit, _sysName: hit ? hit.name : null };
    });
  };

  const addProductFiles = (newFiles) => { const items=Array.from(newFiles).map(f=>({file:f,id:mkPFileId(),status:'pending'})); setPFiles(prev=>[...prev,...items]); };
  const pendingPCount = productFiles.filter(it=>it.status==='pending').length;
  const scanPendingPFiles = async () => { const items=productFiles.filter(it=>it.status==='pending'); if(items.length===0)return; await scanProductItems(items); };
  const deleteProductFile = (id) => { setPFiles(prev=>prev.filter(it=>it.id!==id)); setSelPFIds(prev=>{const s=new Set(prev);s.delete(id);return s;}); const remaining=scanResults.filter(r=>r._fileId!==id); setSRes(remaining); applyBarcodeMap(buildMap(remaining)); };
  // จับคู่ใหม่: ใช้ชื่อที่อ่านจากรูปไว้แล้ว เทียบกับรายการในบิลปัจจุบัน — ไม่ส่งรูปซ้ำ
  const rematchNames = async (onlyBlank) => {
    setRematching(true);
    // ยังไม่รู้ชื่อจริง → ค้นระบบก่อน แล้วจับคู่ด้วยชื่อจริง
    let base = await withSystemNames(scanResults);
    setSRes(base);
    const targets = base.filter(r => r.description_image && (!onlyBlank || !r.match));
    if (!targets.length) { setRematching(false); return; }
    try {
      const allP = invoices.filter(i=>i.status==='done'&&i.data?.products).flatMap(i=>i.data.products)
        .filter((x,i,a)=>a.findIndex(y=>y.description===x.description)===i);
      const list = allP.map(x=>x.description).filter(Boolean).join('\n');
      const payload = targets.map((r,i)=>({ i, barcode: r.barcode || '', name: r.description_image }));
      const res = await callClaude([{ type:'text', text: REMATCH_PROMPT(list, JSON.stringify(payload, null, 1)) }], {}, model);
      const raw = extractJSON(res.content?.find(b=>b.type==='text')?.text || 'null');
      const arr = Array.isArray(raw) ? raw : [];
      const byIdx = new Map(arr.filter(x=>x && x.i != null).map(x=>[Number(x.i), x]));
      const idOf = new Map(targets.map((r,i)=>[r, i]));
      const next = base.map(r => {
        if (!idOf.has(r)) return r;
        const got = byIdx.get(idOf.get(r));
        if (!got) return r;
        return { ...r, match: got.match || null, _userOverride: false };
      });
      setSRes(next); applyBarcodeMap(buildMap(next));
    } catch {}
    setRematching(false);
  };

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
      // บันทึกเป็นบิลจริงแล้ว ร่างบนเซิร์ฟเวอร์ต้องหาย
      if (currentUser?.id) {
        try { await fetch('/api/invoice-draft?keyed_by_id=' + encodeURIComponent(currentUser.id), { method: 'DELETE' }); } catch {}
      }
    } catch(e) { setSbErr(e.message); setSbSt('error'); }
  };

  const reset = () => { setStep(1);setInvoices([]);setPFiles([]);setBMap({});setSRes([]);setSelPFIds(new Set());setFileName('');setSbSt(null);setSbErr(null);setSelPages({});
    setSendSt({ busy:false, done:null, err:'' });
    clearRevise();   // เริ่มใหม่แล้วต้องเป็นบิลใหม่จริง ๆ ไม่ใช่รอบแก้ของใบเดิม
  };

  useEffect(() => {
    if (step===4&&!fileName&&doneInvs.length>0) {
      (async()=>{
        // ชื่อไฟล์ = เลขเอกสารจากตัวนับกลาง เช่น IV-202608220001 — หลายคนคีย์พร้อมกันก็ไม่ชน
        try {
          const r=await fetch('/api/config?next=IV');
          const name=await r.json();
          if(r.ok&&typeof name==='string'&&name){setFileName(name);return;}
        } catch {}
        // เซิร์ฟเวอร์ล่ม: ใส่เวลา+สุ่มท้าย ไม่ใช้ตัวนับในเครื่อง (นั่นคือต้นเหตุที่ชื่อชน)
        const d=new Date(),p=n=>String(n).padStart(2,'0');
        setFileName('IV-'+d.getFullYear()+p(d.getMonth()+1)+p(d.getDate())+'-'+p(d.getHours())+p(d.getMinutes())+'-'+Math.random().toString(36).slice(2,5));
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
    <div style={{ maxWidth: 720, margin: '0 auto', padding: mob?8:0, paddingBottom: 'calc(78px + env(safe-area-inset-bottom))' }}>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      {sentActive && sentView}
      {/* ซ่อนไว้ ไม่ unmount — กลับมาแล้วงานที่คีย์ไว้ยังอยู่ */}
      <div style={{ display: sentActive ? 'none' : 'block' }}>
      {revise && (
        <div style={{ background:'#FEF2F2', border:'1px solid #FECACA', borderRadius:12, padding:'10px 12px', marginBottom:12, display:'flex', flexDirection:'column', gap:6 }}>
          <div style={{ display:'flex', alignItems:'center', gap:8 }}>
            <span style={{ flex:1, minWidth:0, fontSize:12.5, fontWeight:700, color:'#B91C1C', fontFamily:"'IBM Plex Mono', monospace" }}>
              {String(revise.docNo || '').replace(/R\d+$/i,'')}R{(revise.reviseNo || 0) + 1}
            </span>
            <button onClick={reset} title="ยกเลิกการแก้ เริ่มบิลใหม่"
              style={{ flex:'none', minHeight:30, padding:'0 11px', borderRadius:8, border:'1px solid #FECACA', background:'#fff', color:'#B91C1C', fontFamily:'inherit', fontSize:11.5, fontWeight:700, cursor:'pointer' }}>
              ยกเลิก
            </button>
          </div>
          {revise.note && <div style={{ fontSize:11.5, color:'#B91C1C', lineHeight:1.45 }}>ผู้จัดการแจ้ง: “{revise.note}”</div>}
        </div>
      )}
      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:16, flexWrap:'wrap', gap:8 }}>
        <h2 style={{ fontWeight:700, fontSize:20, color:'#0f172a', margin:0 }}>{revise ? 'แก้บิลที่ถูกส่งกลับ' : 'บันทึกบิลซื้อ'}</h2>
        <div style={{ display:'flex', gap:8, alignItems:'center', flexWrap:'wrap' }}>
          <select value={model} onChange={e=>setModel(e.target.value)} style={{ fontSize:12, padding:'4px 8px', borderRadius:9, border:'1px solid #cbd5e1', background:'#f8fafc', color:'#475569' }}>
            {MODELS.map(m=><option key={m.id} value={m.id}>{m.label}</option>)}
          </select>

          {step>1&&!revise&&<button onClick={reset}
            style={{ fontSize:12, padding:'4px 10px', borderRadius:9, border:'1px solid #fca5a5', background:'#fef2f2', color:'#b91c1c', cursor:'pointer' }}>↺ ใหม่</button>}
        </div>
      </div>


      {currentUser && (
        <div style={{ background:'#fff', border:'1px solid #e2e8f0', borderRadius:12, padding:10, marginBottom:12, display:'flex', flexDirection:'column', gap:8 }}>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:8 }}>
            <button onClick={pushInvDraft} disabled={!!invDraft.busy || invoices.length===0}
              style={{ minHeight:44, borderRadius:10, border:'none', background:'#2f6e90', color:'#fff', fontWeight:700, fontSize:12.5, fontFamily:'inherit', cursor:'pointer', opacity:(!!invDraft.busy||invoices.length===0)?0.4:1 }}>
              {invDraft.busy==='up' ? 'กำลังบันทึก…' : 'บันทึกร่างขึ้นเซิร์ฟเวอร์'}
            </button>
            <button onClick={pullInvDraft} disabled={!!invDraft.busy}
              style={{ minHeight:44, borderRadius:10, border:'1px solid #b9cfdc', background:'#eaf0f4', color:'#255771', fontWeight:700, fontSize:12.5, fontFamily:'inherit', cursor:'pointer', opacity:invDraft.busy?0.4:1 }}>
              {invDraft.busy==='down' ? 'กำลังดึง…' : 'ดึงร่างจากเซิร์ฟเวอร์'}
            </button>
          </div>
          {invDraft.msg && <div style={{ fontSize:11, fontWeight:600, color:'#255771' }}>{invDraft.msg}</div>}
          {invDraft.err && <div style={{ fontSize:11, fontWeight:600, color:'#B91C1C' }}>{invDraft.err}</div>}
        </div>
      )}



      {step>1 && invoices.length===0 && (
        <div style={{ background:'#fff', border:'1px dashed #cbd5e1', borderRadius:12, padding:'30px 16px', textAlign:'center', display:'flex', flexDirection:'column', gap:12 }}>
          <div style={{ fontSize:13, color:'#64748b' }}>ยังไม่มีบิลในขั้นนี้</div>
          <button onClick={()=>setStep(1)}
            style={{ minHeight:48, borderRadius:11, border:'none', background:'#2f6e90', color:'#fff', fontFamily:'inherit', fontSize:14.5, fontWeight:700, cursor:'pointer' }}>กลับไปถ่ายรูปบิล</button>
        </div>
      )}

      {step===1&&(
        <div style={{ display:'flex', flexDirection:'column', gap:12 }}>
          <div style={{ fontSize:12.5, color:'#64748b', lineHeight:1.6 }}>ถ่ายบิลทีละหน้า หลายหน้าที่เป็นใบเดียวกันให้อยู่กลุ่มเดียวกัน — AI จะอ่านต่อกันเป็นใบเดียว</div>

          <DropZone bare multiple accept="image/*" capture onFiles={addFiles}>
            <div style={{ minHeight:54, borderRadius:12, background:'#2f6e90', color:'#fff', fontSize:16, fontWeight:700, display:'flex', alignItems:'center', justifyContent:'center', gap:9, boxShadow:'0 2px 0 #255771' }}>
              <span style={{ width:18, height:15, border:'2px solid #fff', borderRadius:4, display:'inline-block' }}></span>ถ่ายรูปบิล
            </div>
          </DropZone>

          <DropZone bare multiple accept="image/*,.pdf" onFiles={addFiles}>
            <div style={{ minHeight:46, borderRadius:11, background:'#fff', border:'1px solid #e4e6ea', color:'#475569', fontSize:13.5, fontWeight:600, display:'flex', alignItems:'center', justifyContent:'center' }}>
              เลือกจากคลังรูป / ไฟล์ PDF
            </div>
          </DropZone>

          {invoices.length===0 ? (
            <div style={{ background:'#fff', border:'1px dashed #cbd5e1', borderRadius:12, padding:'34px 16px', textAlign:'center', color:'#94a3b8', fontSize:12.5 }}>
              ยังไม่มีรูป — เริ่มจากถ่ายรูปบิล
            </div>
          ) : (
            <div style={{ display:'flex', flexDirection:'column', gap:12 }}>
              {invoices.map((inv,gi)=>(
                <div key={gi} style={{ background:'#fff', border:'1px solid #e4e6ea', borderRadius:12, padding:12 }}>
                  <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', gap:10 }}>
                    <div style={{ fontSize:12.5, fontWeight:700, color:'#334155', whiteSpace:'nowrap' }}>กลุ่ม {gi+1} · {inv.files.length} หน้า</div>
                    <div style={{ display:'flex', alignItems:'center', gap:8, flex:'none' }}>
                      <span style={{ fontSize:11, color:'#94a3b8', whiteSpace:'nowrap' }}>= 1 ใบกำกับ</span>
                      <button onClick={()=>removeInv(gi)} style={{ minHeight:34, padding:'0 10px', borderRadius:9, background:'#fef2f2', color:'#b91c1c', border:'none', cursor:'pointer', fontSize:11.5, fontWeight:700, fontFamily:'inherit' }}>ลบกลุ่ม</button>
                    </div>
                  </div>

                  <div style={{ display:'flex', gap:8, marginTop:11, flexWrap:'wrap' }}>
                    {inv.files.map((f,pi)=>(
                      <div key={pi} style={{ position:'relative', width:70, height:92, flex:'none' }}>
                        <div style={{ width:70, height:92, borderRadius:8, border:'1px solid #e4e6ea', background:'#f8fafc', overflow:'hidden' }}>
                          <InvFileThumb file={f}/>
                        </div>
                        <div style={{ position:'absolute', bottom:3, right:3, fontSize:9, fontWeight:600, fontFamily:"'IBM Plex Mono', monospace", background:'rgba(15,23,42,.72)', color:'#fff', borderRadius:3, padding:'1px 4px' }}>
                          {inv.pageStatus[pi]==='processing'?'…':inv.pageStatus[pi]==='done'?'✓':inv.pageStatus[pi]==='error'?'✗':pi+1}
                        </div>
                        <button onClick={()=>removePage(gi,pi)} aria-label="ลบหน้านี้"
                          style={{ position:'absolute', top:-7, left:-7, width:26, height:26, borderRadius:'50%', background:'#b91c1c', color:'#fff', border:'2px solid #fff', cursor:'pointer', fontSize:12, display:'flex', alignItems:'center', justifyContent:'center', lineHeight:1, zIndex:10, padding:0 }}>✕</button>
                      </div>
                    ))}
                  </div>

                  <div style={{ marginTop:11, fontSize:11, color:'#94a3b8', lineHeight:1.55 }}>ถ่ายให้เห็นครบทั้งหัวบิลและบรรทัดล่างสุด · แสงสะท้อนบนพลาสติกใสทำให้อ่านตัวเลขพลาดบ่อยที่สุด</div>
                </div>
              ))}

              <button onClick={processAll} disabled={processing}
                style={{ width:'100%', minHeight:54, borderRadius:12, background:processing?'#94a3b8':'#2f6e90', color:'#fff', fontWeight:700, fontSize:15.5, border:'none', fontFamily:'inherit', cursor:processing?'not-allowed':'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>
                {processing?<><Spinner size={18}/>AI กำลังอ่านบิล…</>:<>อ่านด้วย AI ({invoices.reduce((s,i)=>s+i.files.length,0)} หน้า) →</>}
              </button>
            </div>
          )}
        </div>
      )}

      {step===2&&(
        <div style={{ display:'flex', flexDirection:'column', gap:12 }}>
          {invoices.length===0 && (
            <div style={{ background:'#fff', border:'1px dashed #cbd5e1', borderRadius:12, padding:'26px 16px', textAlign:'center' }}>
              <div style={{ fontSize:13, fontWeight:700, color:'#475569' }}>ยังไม่มีบิลให้อ่าน</div>
              <div style={{ fontSize:11.5, color:'#94a3b8', marginTop:4, lineHeight:1.5 }}>อัปโหลดรูปบิลที่ขั้น 1 ก่อน แล้วกลับมาที่นี่เพื่อให้ AI อ่าน</div>
              <button onClick={()=>setStep(1)} style={{ marginTop:12, minHeight:40, padding:'0 14px', borderRadius:10, border:'none', background:'#2f6e90', color:'#fff', fontFamily:'inherit', fontSize:12.5, fontWeight:700, cursor:'pointer' }}>ไปขั้นอัปโหลดรูป</button>
            </div>
          )}
          <div style={{ background:'#eaf0f4', border:'1px solid #b9cfdc', borderRadius:12, padding:12, display:'flex', gap:9, alignItems:'center' }}>
            <span style={{ fontSize:15, color:'#255771', flex:'none' }}>✓</span>
            <div style={{ flex:1, minWidth:0, fontSize:11.5, color:'#255771', lineHeight:1.5 }}>
              อ่านได้ {doneInvs.reduce((s,inv)=>s+((inv.data?.products||[]).length),0)} รายการจาก {doneInvs.length} ใบ — ขั้นนี้เป็นงาน<strong>เสริม</strong> ข้ามไปตรวจข้อมูลได้เลย
            </div>
          </div>

          <div style={{ fontSize:12.5, color:'#64748b', lineHeight:1.6 }}>บิลบางใบไม่มีบาร์โค้ด มีแต่ชื่อสินค้า ถ่ายรูปตัวสินค้าเพิ่มเพื่อให้ AI ดึงบาร์โค้ดมาจับคู่ให้</div>

          {/* capture = เปิดกล้องเลย / ไม่ใส่ = เลือกจากคลังรูปหรือลากไฟล์มาวางได้ */}
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:8 }}>
            <DropZone bare multiple accept="image/*" capture onFiles={addProductFiles}>
              <div style={{ minHeight:50, borderRadius:12, background:'#2f6e90', color:'#fff', fontSize:14.5, fontWeight:700, display:'flex', alignItems:'center', justifyContent:'center', gap:7, boxShadow:'0 2px 0 #255771' }}>
                <Camera size={17} />ถ่ายรูปสินค้า
              </div>
            </DropZone>
            <DropZone bare multiple accept="image/*" onFiles={addProductFiles}>
              <div style={{ minHeight:50, borderRadius:12, background:'#fff', border:'1px solid #e4e6ea', color:'#334155', fontSize:14.5, fontWeight:700, display:'flex', alignItems:'center', justifyContent:'center', gap:7 }}>
                <ImageIcon size={17} />อัพโหลดรูป
              </div>
            </DropZone>
          </div>

          {productFiles.length>0&&(
            <div style={{ background:'#fff', border:'1px solid #e4e6ea', borderRadius:12, padding:12 }}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', gap:8, flexWrap:'wrap' }}>
                <div style={{ fontSize:12.5, fontWeight:700, color:'#334155' }}>รูปสินค้า · {productFiles.length} ไฟล์</div>
                <div style={{ display:'flex', gap:6 }}>
                  {productFiles.some(it=>it.status==='pending')&&(
                    <button onClick={scanPendingPFiles} disabled={scanning}
                      style={{ minHeight:36, padding:'0 12px', borderRadius:9, background:'#2f6e90', color:'#fff', border:'none', cursor:scanning?'wait':'pointer', fontWeight:700, fontSize:11.5, fontFamily:'inherit' }}>
                      {scanning?'กำลังอ่าน…':`เริ่มอ่าน (${productFiles.filter(it=>it.status==='pending').length})`}
                    </button>
                  )}
                  {selPFileIds.size>0&&(
                    <button onClick={rescanSelectedPFiles} disabled={scanning}
                      style={{ minHeight:36, padding:'0 12px', borderRadius:9, background:'#eaf0f4', color:'#255771', border:'1px solid #b9cfdc', cursor:'pointer', fontWeight:700, fontSize:11.5, fontFamily:'inherit' }}>
                      {scanning?'กำลังอ่าน…':`อ่านซ้ำ ${selPFileIds.size} รูป`}
                    </button>
                  )}
                </div>
              </div>
              <div style={{ display:'flex', flexWrap:'wrap', gap:8, marginTop:11 }}>
                {productFiles.map(it=>(
                  <div key={it.id} style={{ position:'relative', width:70, height:92, flex:'none' }}>
                    <div style={{ cursor:'pointer', width:70, height:92, borderRadius:8, border:'1px solid #e4e6ea', background:'#f8fafc', overflow:'hidden' }}
                      onClick={()=>setSelPFIds(prev=>{const s=new Set(prev);s.has(it.id)?s.delete(it.id):s.add(it.id);return s;})}>
                      <InvFileThumb file={it.file}/>
                    </div>
                    <div style={{ position:'absolute', bottom:3, right:3, fontSize:9, fontWeight:700, fontFamily:"'IBM Plex Mono', monospace", color:'#fff', borderRadius:3, padding:'1px 4px',
                      background: it.status==='done'?'#15803d':it.status==='error'?'#b91c1c':it.status==='processing'?'#b45309':'rgba(15,23,42,.72)' }}>
                      {it.status==='done'?'✓':it.status==='error'?'✗':it.status==='processing'?'…':'⋯'}
                    </div>
                    {selPFileIds.has(it.id)&&<div style={{ position:'absolute', inset:0, border:'2px solid #2f6e90', borderRadius:8, background:'rgba(47,110,144,.12)', pointerEvents:'none' }}/>}
                    <button onClick={e=>{e.stopPropagation();deleteProductFile(it.id);}} aria-label="ลบรูปนี้"
                      style={{ position:'absolute', top:-7, left:-7, width:26, height:26, borderRadius:'50%', background:'#b91c1c', color:'#fff', border:'2px solid #fff', cursor:'pointer', fontSize:12, display:'flex', alignItems:'center', justifyContent:'center', lineHeight:1, zIndex:10, padding:0 }}>✕</button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {scanResults.length>0&&(
            <div style={{ background:'#fff', border:'1px solid #e4e6ea', borderRadius:12, overflow:'hidden' }}>
              <div style={{ padding:'10px 12px', background:'#f8fafc', borderBottom:'1px solid #e4e6ea', display:'flex', flexDirection:'column', gap:8 }}>
                <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                  <span style={{ fontSize:11, fontWeight:700, color:'#64748b' }}>ผลการจับคู่บาร์โค้ด · {scanResults.length} รูป</span>
                  {scanResults.filter(r=>r._inSystem).length>0 && (
                    <span style={{ fontSize:10, fontWeight:700, color:'#15803d', background:'#f0fdf4', borderRadius:99, padding:'3px 8px' }}>
                      มีในระบบ {scanResults.filter(r=>r._inSystem).length}
                    </span>
                  )}
                  {scanResults.filter(r=>!r.match).length>0 && (
                    <span style={{ fontSize:10, fontWeight:700, color:'#b91c1c', background:'#fef2f2', borderRadius:99, padding:'3px 8px' }}>
                      ยังไม่ได้คู่ {scanResults.filter(r=>!r.match).length}
                    </span>
                  )}
                </div>
                <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:7 }}>
                  <button onClick={()=>rematchNames(true)} disabled={rematching||scanning}
                    style={{ minHeight:38, borderRadius:9, border:'1px solid #e4e6ea', background:'#fff', fontFamily:'inherit', fontSize:12, fontWeight:700, color:'#334155', cursor:'pointer', opacity:(rematching||scanning)?0.6:1 }}>
                    {rematching?'กำลังจับคู่…':'จับคู่เฉพาะที่ว่าง'}
                  </button>
                  <button onClick={()=>rematchNames(false)} disabled={rematching||scanning}
                    title="ทับผลเดิมทั้งหมด รวมที่แก้เองไว้"
                    style={{ minHeight:38, borderRadius:9, border:'none', background:'#2f6e90', color:'#fff', fontFamily:'inherit', fontSize:12, fontWeight:700, cursor:'pointer', opacity:(rematching||scanning)?0.6:1 }}>
                    {rematching?'กำลังจับคู่…':'จับคู่ใหม่ทั้งหมด'}
                  </button>
                </div>
                <div style={{ fontSize:10.5, color:'#94a3b8', lineHeight:1.5 }}>
                  จับคู่จากชื่อที่อ่านจากรูป · เช็คด้วยว่าบาร์โค้ดมีในระบบไหม — ไม่ต้องถ่ายรูปใหม่
                </div>
              </div>
              {scanResults.map((r,i)=>{
                const usedByOthers = new Set(scanResults.filter((_,j)=>j!==i).map(x=>x.match).filter(Boolean));
                const allDescs = [...new Set(doneInvs.flatMap(inv=>(inv.data.products||[]).map(p=>p.description)).filter(Boolean))];
                return (
                  <div key={i} style={{ padding:'11px 12px', borderBottom:'1px solid #f6f7f8', display:'flex', flexDirection:'column', gap:7, background: r._userOverride ? '#fffbeb' : 'transparent' }}>
                    <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                      <span style={{ fontSize:10.5, fontWeight:700, color:'#94a3b8', fontFamily:"'IBM Plex Mono', monospace", flex:'none' }}>{i+1}</span>
                      <input value={r.barcode??''} placeholder="บาร์โค้ด"
                        onChange={e=>{const u=[...scanResults];u[i]={...u[i],barcode:e.target.value||null};setSRes(u);applyBarcodeMap(buildMap(u));}}
                        style={{ flex:1, minWidth:0, minHeight:38, fontFamily:"'IBM Plex Mono', monospace", fontSize:12.5, fontWeight:600, color:'#15803d', border:'1px solid #e4e6ea', borderRadius:9, padding:'0 9px' }}/>
                      {r._userOverride
                      ? <span style={{ flex:'none', fontSize:10, fontWeight:700, color:'#b45309', background:'#fffbeb', borderRadius:99, padding:'4px 8px', whiteSpace:'nowrap' }}>แก้เอง</span>
                      : r._inSystem === true
                        ? <span style={{ flex:'none', fontSize:10, fontWeight:700, color:'#15803d', background:'#f0fdf4', borderRadius:99, padding:'4px 8px', whiteSpace:'nowrap' }}>มีในระบบ</span>
                        : r._inSystem === false && r.barcode
                          ? <span style={{ flex:'none', fontSize:10, fontWeight:700, color:'#b91c1c', background:'#fef2f2', borderRadius:99, padding:'4px 8px', whiteSpace:'nowrap' }}>ไม่มีในระบบ</span>
                          : null}
                    </div>
                    <input value={r.description_image??''} placeholder="ชื่อที่อ่านจากรูป"
                      onChange={e=>{const u=[...scanResults];u[i]={...u[i],description_image:e.target.value||null};setSRes(u);}}
                      style={{ width:'100%', minHeight:38, fontSize:12.5, color:'#334155', border:'1px solid #e4e6ea', borderRadius:9, padding:'0 9px' }}/>
                    {r._inSystem && r._sysName && (
                      <div style={{ fontSize:10.5, color:'#15803d', lineHeight:1.4 }}>ในระบบชื่อ “{r._sysName}”</div>
                    )}
                    {r._inSystem === false && r.barcode && (
                      <div style={{ fontSize:10.5, color:'#b91c1c', lineHeight:1.4 }}>บาร์โค้ดนี้ไม่มีในระบบ — เช็คว่าอ่านถูกไหม หรือเป็นสินค้าใหม่</div>
                    )}
                    <select value={r.match??''}
                      onChange={e=>{const u=[...scanResults];u[i]={...u[i],match:e.target.value||null,_userOverride:true};setSRes(u);applyBarcodeMap(buildMap(u));}}
                      style={{ width:'100%', minHeight:40, fontSize:12.5, fontFamily:'inherit', borderRadius:9, padding:'0 9px',
                               border: r._userOverride?'1px solid #b45309':'1px solid #e4e6ea', background: r._userOverride?'#fffbeb':'#fff' }}>
                      <option value="">— จับคู่กับรายการในบิล —</option>
                      {allDescs.map((desc,idx) => {
                        const isUsed = usedByOthers.has(desc);
                        return <option key={idx} value={desc} disabled={isUsed}>{(desc||'').slice(0,40)}{isUsed?' (ใช้แล้ว)':''}</option>;
                      })}
                    </select>
                  </div>
                );
              })}
              <div style={{ padding:'11px 12px', fontSize:11, color:'#94a3b8', lineHeight:1.55 }}>แถวสีเหลืองคือแก้เอง — ระบบจำการแก้ไว้ใช้กับบิลใบต่อไป</div>
            </div>
          )}

          <div style={{ display:'grid', gridTemplateColumns:'auto minmax(0,1fr)', gap:8, alignItems:'stretch' }}>
            <button onClick={()=>setStep(1)} aria-label="ย้อนกลับ"
              style={{ minWidth:56, minHeight:54, borderRadius:12, border:'1px solid #e4e6ea', background:'#fff', color:'#475569', fontFamily:'inherit', fontSize:20, fontWeight:700, cursor:'pointer' }}>‹</button>
            {(() => {
              const unread = productFiles.filter(it=>it.status==='pending').length;
              if (unread > 0 || scanning) return (
                <button onClick={scanPendingPFiles} disabled={scanning}
                  style={{ minHeight:54, borderRadius:12, background:'#2f6e90', color:'#fff', fontWeight:700, fontSize:15.5, border:'none', fontFamily:'inherit', cursor:scanning?'wait':'pointer', boxShadow:'0 2px 0 #255771', display:'flex', alignItems:'center', justifyContent:'center', gap:8, opacity:scanning?0.75:1 }}>
                  <Sparkles size={18} />{scanning ? 'กำลังอ่าน…' : `อ่านด้วย AI (${unread})`}
                </button>
              );
              return (
                <button onClick={()=>setStep(3)}
                  style={{ minHeight:54, borderRadius:12, background:'#2f6e90', color:'#fff', fontWeight:700, fontSize:15.5, border:'none', fontFamily:'inherit', cursor:'pointer', boxShadow:'0 2px 0 #255771', display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>
                  ตรวจสอบข้อมูล<span style={{ fontSize:18 }}>›</span>
                </button>
              );
            })()}
          </div>
        </div>
      )}

      {step===3&&(
        <div style={{ display:'flex', flexDirection:'column', gap:12 }}>
          {invoices.length===0 && (
            <div style={{ background:'#fff', border:'1px dashed #cbd5e1', borderRadius:12, padding:'26px 16px', textAlign:'center' }}>
              <div style={{ fontSize:13, fontWeight:700, color:'#475569' }}>ยังไม่มีรายการให้ตรวจ</div>
              <div style={{ fontSize:11.5, color:'#94a3b8', marginTop:4, lineHeight:1.5 }}>ต้องอ่านบิลด้วย AI ที่ขั้น 2 ก่อน</div>
              <button onClick={()=>setStep(1)} style={{ marginTop:12, minHeight:40, padding:'0 14px', borderRadius:10, border:'none', background:'#2f6e90', color:'#fff', fontFamily:'inherit', fontSize:12.5, fontWeight:700, cursor:'pointer' }}>ไปขั้นอัปโหลดรูป</button>
            </div>
          )}
          <div style={{ fontSize:12.5, color:'#64748b', lineHeight:1.6 }}>AI อ่านมาให้แล้ว — ตรวจทีละรายการ แตะแก้ตัวเลขที่ผิดได้เลย ยอดสุทธิคิดใหม่ให้ทันที</div>

          {doneInvs.length===0 && invoices.length>0 && (
            <div style={{ background:'#fffbeb', border:'1px solid #fde68a', borderRadius:12, padding:14, display:'flex', flexDirection:'column', gap:10 }}>
              <div style={{ fontSize:12.5, color:'#b45309', lineHeight:1.6 }}>ยังไม่มีใบที่อ่านสำเร็จ — กลับไปกด “อ่านด้วย AI” ก่อน</div>
              <button onClick={()=>setStep(1)}
                style={{ minHeight:46, borderRadius:11, border:'1px solid #fde68a', background:'#fff', color:'#b45309', fontFamily:'inherit', fontSize:13.5, fontWeight:700, cursor:'pointer' }}>กลับไปหน้าอัปโหลด</button>
            </div>
          )}
          {doneInvs.map((inv,gi)=>{
            const d=inv.data, vs=vatSummary(d.products||[]);
            const upd = (patch) => updateData(gi, {...d, ...patch});
            const updP = (pi, patch) => { const prods=[...d.products]; prods[pi]=recalc({...prods[pi],...patch,_pt:d.price_type??'incl'}); updateData(gi,{...d,products:prods}); };
            const open = invOpen[gi] ?? false;   // เริ่มต้นพับไว้
            const F = ({ label, k, wide, tag }) => ((
              <div style={{ gridColumn: wide ? '1 / -1' : 'auto', minWidth:0 }}>
                <div style={{ color:'#94a3b8', fontSize:10, marginBottom:3, display:'flex', alignItems:'center', gap:4 }}>{label}{tag}</div>
                <input value={d[k]??''} onChange={e=>upd({[k]:e.target.value||null})}
                  style={{ width:'100%', minHeight:40, padding:'0 9px', border:'1px solid #e4e6ea', borderRadius:9, fontSize:13, boxSizing:'border-box' }}/>
              </div>
            ));
            return (
              <div key={gi} style={{ background:'#fff', border:'1px solid #e4e6ea', borderRadius:12, overflow:'hidden' }}>
                <div style={{ background:'#eaf0f4', padding:'10px 12px', borderBottom:'1px solid #b9cfdc', display:'flex', alignItems:'center', gap:8 }}>
                  <div style={{ flex:1, minWidth:0 }}>
                    <div style={{ fontSize:10, fontWeight:700, color:'#255771' }}>ใบที่ {gi+1} · {(d.products||[]).length} รายการ</div>
                    <div style={{ fontSize:13.5, fontWeight:700, color:'#0f172a', marginTop:1, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{d.vendor_name||'(ไม่ระบุชื่อร้าน)'}</div>
                  </div>
                  <button onClick={()=>reprocessInvoice(gi)}
                    style={{ flex:'none', minHeight:36, padding:'0 10px', borderRadius:9, background:'#fff', color:'#255771', border:'1px solid #b9cfdc', cursor:'pointer', fontSize:11.5, fontWeight:700, fontFamily:'inherit' }}>อ่านใหม่</button>
                </div>

                <div style={{ padding:12, display:'flex', flexDirection:'column', gap:10 }}>
                  {/* หัวบิล */}
                  <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:8 }}>
                    {F({ label:'ชื่อร้าน / บริษัท', k:'vendor_name', wide:true })}
                    {F({ label:'เลขที่ใบกำกับ', k:'invoice_no' })}
                    {F({ label:'วันที่', k:'invoice_date' })}
                    {F({ label:'เลขภาษี', k:'vendor_tax_id' })}
                    {F({ label:'สาขา', k:'vendor_branch' })}
                    {F({ label:'ประเภทเอกสาร', k:'document_type' })}
                    {F({ label:'รหัสผู้ขาย', k:'vendor_no',
                      tag: d._vendorFromDB && d.vendor_no ? (
                        d._vendorMatchBy === 'near'
                          ? <span style={{ background:'#fffbeb', color:'#b45309', fontSize:9, fontWeight:700, padding:'1px 5px', borderRadius:4 }}>ชื่อใกล้เคียง เช็คซ้ำ</span>
                          : <span style={{ background:'#f0fdf4', color:'#15803d', fontSize:9, fontWeight:700, padding:'1px 5px', borderRadius:4 }}>{d._vendorMatchBy === 'tax' ? 'ตรงเลขภาษี' : 'จากระบบ'}</span>
                      ) : null })}
                    <div style={{ gridColumn:'1 / -1', minWidth:0 }}>
                      <div style={{ color:'#94a3b8', fontSize:10, marginBottom:3 }}>ราคาบนบิล</div>
                      <select value={d.price_type??'incl'}
                        onChange={e=>{ const pt=e.target.value; const prods=(d.products||[]).map(p=>recalc({...p,_pt:pt})); updateData(gi,{...d,price_type:pt,products:prods}); }}
                        style={{ width:'100%', minHeight:42, padding:'0 9px', border:'1px solid #e4e6ea', borderRadius:9, fontSize:13, fontFamily:'inherit' }}>
                        <option value="incl">รวม VAT แล้ว</option>
                        <option value="excl">ยังไม่รวม VAT</option>
                      </select>
                    </div>
                    <div style={{ gridColumn:'1 / -1', minWidth:0 }}>
                      <div style={{ color:'#94a3b8', fontSize:10, marginBottom:3 }}>ที่อยู่</div>
                      <textarea value={d.vendor_address??''} onChange={e=>upd({vendor_address:e.target.value||null})} rows={2}
                        style={{ width:'100%', padding:'8px 9px', border:'1px solid #e4e6ea', borderRadius:9, fontSize:13, resize:'vertical', boxSizing:'border-box', fontFamily:'inherit' }}/>
                    </div>
                  </div>

                  {/* ยอดรวมของใบ */}
                  <div style={{ background:'#f8fafc', borderRadius:10, padding:10, display:'grid', gridTemplateColumns:'1fr 1fr', gap:8 }}>
                    {[['ส่วนลดรวม', '-฿'+vs.sdTot.toLocaleString(), '#b91c1c'],
                      ['ไม่รวม VAT', '฿'+vs.excl.toLocaleString(), '#334155'],
                      ['VAT 7%', '฿'+vs.vatAmt.toLocaleString(), '#334155'],
                      ['ยอดสุทธิ', '฿'+vs.netTotal.toLocaleString(), '#15803d']].map(([k,v,c])=>(
                      <div key={k}>
                        <div style={{ fontSize:10, color:'#94a3b8' }}>{k}</div>
                        <div style={{ fontSize:14, fontWeight:700, color:c, fontFamily:"'IBM Plex Mono', monospace", wordBreak:'break-all' }}>{v}</div>
                      </div>
                    ))}
                  </div>

                  <button onClick={()=>setInvOpen(prev=>({ ...prev, [gi]: !open }))}
                    style={{ width:'100%', minHeight:42, borderRadius:10, border:'1px solid #e4e6ea', background:'#fff', color:'#475569', fontFamily:'inherit', fontSize:12.5, fontWeight:700, cursor:'pointer' }}>
                    {open ? 'ซ่อนรายการสินค้า' : `ดูรายการสินค้า (${(d.products||[]).length})`}
                  </button>

                  {/* รายการสินค้า — การ์ดต่อรายการ */}
                  {open && (d.products||[]).map((p,pi)=>{
                    const bc=p.barcode??barcodeMap[String(p.description||'').trim()]??null;
                    const diffColor = p.diff == null ? '#94a3b8'
                      : Math.abs(p.diff) < 0.01 ? '#15803d'
                      : p.diff > 0 ? '#b45309' : '#b91c1c';
                    const num = (label, field) => (
                      <div style={{ minWidth:0 }}>
                        <div style={{ fontSize:9.5, color:'#94a3b8', marginBottom:2 }}>{label}</div>
                        <input type="number" inputMode="decimal" value={p[field]??''} onChange={e=>updP(pi,{[field]:e.target.value===''?null:e.target.value})}
                          style={{ width:'100%', minHeight:40, border:'1px solid #e4e6ea', borderRadius:9, padding:'0 6px', textAlign:'center', fontSize:13.5, fontFamily:"'IBM Plex Mono', monospace", boxSizing:'border-box' }}/>
                      </div>
                    );
                    const txt = (label, field) => (
                      <div style={{ minWidth:0 }}>
                        <div style={{ fontSize:9.5, color:'#94a3b8', marginBottom:2 }}>{label}</div>
                        <input type="text" value={p[field]??''} onChange={e=>updP(pi,{[field]:e.target.value||null})}
                          style={{ width:'100%', minHeight:40, border:'1px solid #e4e6ea', borderRadius:9, padding:'0 6px', textAlign:'center', fontSize:13, boxSizing:'border-box' }}/>
                      </div>
                    );
                    const pKey = gi + '_' + pi;
                    const pOpen = !!prodOpen[pKey];
                    return (
                      <div key={pi} style={{ border:'1px solid #e4e6ea', borderRadius:11, display:'flex', flexDirection:'column', background: bc ? '#fff' : '#fffbeb', overflow:'hidden' }}>
                        <button onClick={()=>setProdOpen(prev=>({ ...prev, [pKey]: !prev[pKey] }))}
                          style={{ width:'100%', textAlign:'left', minHeight:56, padding:'8px 10px', display:'flex', alignItems:'center', gap:9, background:'transparent', border:'none', cursor:'pointer', fontFamily:'inherit' }}>
                          <span style={{ flex:'none', fontSize:11, fontWeight:700, color:'#94a3b8', fontFamily:"'IBM Plex Mono', monospace" }}>{p.no??pi+1}</span>
                          <span style={{ flex:1, minWidth:0 }}>
                            <span style={{ display:'block', fontSize:13.5, fontWeight:600, color:'#0f172a', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{p.description||'(ไม่มีชื่อ)'}</span>
                            <span style={{ display:'block', fontSize:10, color: bc?'#94a3b8':'#b45309', marginTop:2, fontFamily:"'IBM Plex Mono', monospace" }}>{bc || 'ยังไม่มีบาร์โค้ด'}</span>
                          </span>
                          <span style={{ flex:'none', textAlign:'right' }}>
                            <span style={{ display:'block', fontSize:14, fontWeight:700, color:'#15803d', fontFamily:"'IBM Plex Mono', monospace" }}>{p.total!=null?Number(p.total).toLocaleString():'—'}</span>
                            <span style={{ display:'block', fontSize:9.5, color:'#94a3b8' }}>{pOpen?'ซ่อน':'แตะแก้'}</span>
                          </span>
                        </button>

                        {pOpen && (
                        <div style={{ padding:'0 10px 10px', display:'flex', flexDirection:'column', gap:8 }}>
                        <div>
                          <div style={{ fontSize:9.5, color:'#94a3b8', marginBottom:2 }}>ชื่อสินค้าบนบิล</div>
                          <input value={p.description??''} placeholder="ชื่อสินค้าบนบิล"
                            onChange={e=>{const prods=[...d.products];prods[pi]={...prods[pi],description:e.target.value};updateData(gi,{...d,products:prods});}}
                            style={{ width:'100%', minHeight:42, border:'1px solid #e4e6ea', borderRadius:9, padding:'0 9px', fontSize:13.5, fontWeight:600, color:'#0f172a', boxSizing:'border-box' }}/>
                        </div>

                        <div>
                          <div style={{ fontSize:9.5, color:'#94a3b8', marginBottom:2 }}>บาร์โค้ด{!bc && ' · ยังไม่มี'}</div>
                          <input value={bc??''} placeholder="ยังไม่มีบาร์โค้ด"
                            onChange={e=>{const prods=[...d.products];prods[pi]={...prods[pi],barcode:e.target.value||null};updateData(gi,{...d,products:prods});}}
                            style={{ width:'100%', minHeight:42, fontFamily:"'IBM Plex Mono', monospace", fontSize:13.5, fontWeight:600, border:'1px solid #e4e6ea', borderRadius:9, padding:'0 9px', color: bc?'#15803d':'#b45309', background:'#fff', boxSizing:'border-box' }}/>
                        </div>

                        <div style={{ display:'grid', gridTemplateColumns:'repeat(3,minmax(0,1fr))', gap:6 }}>
                          {num('ขนาดลัง','carton_size')}
                          {num('ลัง','carton')}
                          {txt('หน่วยลัง','carton_unit')}
                          {num('ชิ้น','ea')}
                          {txt('หน่วย','ea_unit')}
                          <div style={{ minWidth:0 }}>
                            <div style={{ fontSize:9.5, color:'#94a3b8', marginBottom:2 }}>รวมจำนวน</div>
                            <div style={{ minHeight:40, borderRadius:9, background:'#f8fafc', display:'flex', alignItems:'center', justifyContent:'center', fontSize:14, fontWeight:700, color:'#0f172a', fontFamily:"'IBM Plex Mono', monospace" }}>{p.qty??'—'}</div>
                          </div>
                        </div>

                        <div style={{ display:'grid', gridTemplateColumns:'repeat(3,minmax(0,1fr))', gap:6 }}>
                          {num('ราคา/หน่วย','price_ea')}
                          {num('ยอดตามใบ','amount')}
                          {num('ส่วนลด','special_discount')}
                        </div>

                        <div style={{ display:'grid', gridTemplateColumns:'repeat(3,minmax(0,1fr))', gap:6, alignItems:'end' }}>
                          <div style={{ minWidth:0 }}>
                            <div style={{ fontSize:9.5, color:'#94a3b8', marginBottom:2 }}>ยอดสุทธิ</div>
                            <div style={{ minHeight:40, borderRadius:9, background:'#f0fdf4', display:'flex', alignItems:'center', justifyContent:'center', fontSize:14, fontWeight:700, color:'#15803d', fontFamily:"'IBM Plex Mono', monospace" }}>{p.total!=null?Number(p.total).toLocaleString():'—'}</div>
                          </div>
                          <div style={{ minWidth:0 }}>
                            <div style={{ fontSize:9.5, color:'#94a3b8', marginBottom:2 }}>ต่างจากใบ</div>
                            <div style={{ minHeight:40, borderRadius:9, background:'#f8fafc', display:'flex', alignItems:'center', justifyContent:'center', fontSize:13.5, fontWeight:700, color:diffColor, fontFamily:"'IBM Plex Mono', monospace" }}>{p.diff!=null?Number(p.diff).toLocaleString():'—'}</div>
                          </div>
                          <div style={{ minWidth:0 }}>
                            <div style={{ fontSize:9.5, color:'#94a3b8', marginBottom:2 }}>VAT</div>
                            <select value={p.vat??'v'} onChange={e=>updP(pi,{vat:e.target.value})}
                              style={{ width:'100%', minHeight:40, border:'1px solid #e4e6ea', borderRadius:9, padding:'0 6px', fontSize:13, fontFamily:'inherit' }}>
                              <option value="v">7%</option><option value="n">0%</option>
                            </select>
                          </div>
                        </div>
                        </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}

          <div style={{ display:'grid', gridTemplateColumns:'auto minmax(0,1fr)', gap:8, alignItems:'stretch' }}>
            <button onClick={()=>setStep(2)} aria-label="ย้อนกลับ"
              style={{ minWidth:56, minHeight:54, borderRadius:12, border:'1px solid #e4e6ea', background:'#fff', color:'#475569', fontFamily:'inherit', fontSize:20, fontWeight:700, cursor:'pointer' }}>‹</button>
            <button onClick={()=>setStep(4)} disabled={doneInvs.length===0}
              style={{ minHeight:54, borderRadius:12, border:'none', fontFamily:'inherit', fontSize:15.5, fontWeight:700, color:'#fff',
                       background: doneInvs.length===0?'#94a3b8':'#2f6e90',
                       boxShadow: doneInvs.length===0?'none':'0 2px 0 #255771',
                       cursor: doneInvs.length===0?'not-allowed':'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>สรุปและบันทึก<span style={{ fontSize:18 }}>›</span></button>
          </div>
        </div>
      )}

      {step===4&&(
        <div style={{ display:'flex', flexDirection:'column', gap:12 }}>
          {invoices.length===0 && (
            <div style={{ background:'#fff', border:'1px dashed #cbd5e1', borderRadius:12, padding:'26px 16px', textAlign:'center' }}>
              <div style={{ fontSize:13, fontWeight:700, color:'#475569' }}>ยังไม่มีบิลให้ส่ง</div>
              <div style={{ fontSize:11.5, color:'#94a3b8', marginTop:4, lineHeight:1.5 }}>ตรวจรายการที่ขั้น 3 ให้เสร็จก่อน</div>
              <button onClick={()=>setStep(1)} style={{ marginTop:12, minHeight:40, padding:'0 14px', borderRadius:10, border:'none', background:'#2f6e90', color:'#fff', fontFamily:'inherit', fontSize:12.5, fontWeight:700, cursor:'pointer' }}>ไปขั้นอัปโหลดรูป</button>
            </div>
          )}
          <div style={{ display:'grid', gridTemplateColumns:'repeat(3,minmax(0,1fr))', gap:8 }}>
            {[['ใบกำกับ',`${doneInvs.length}`,'ใบ'],['สินค้า',`${allProds.length}`,'รายการ'],['ยอดรวม',`฿${grandTotal.toLocaleString()}`,'']].map(([k,v,u])=>(
              <div key={k} style={{ background:'#fff', border:'1px solid #e4e6ea', borderRadius:12, padding:'10px 8px', textAlign:'center' }}>
                <div style={{ fontSize:10, color:'#94a3b8' }}>{k}</div>
                <div style={{ fontWeight:700, fontSize:16, color:'#0f172a', fontFamily:"'IBM Plex Mono', monospace", marginTop:2, wordBreak:'break-all' }}>{v}</div>
                {u&&<div style={{ fontSize:9.5, color:'#94a3b8' }}>{u}</div>}
              </div>
            ))}
          </div>

          {(() => {
            const issues = [];
            for (const inv of doneInvs) {
              for (const p of (inv.data.products||[])) {
                const miss = [];
                const bc = p.barcode ?? barcodeMap[String(p.description||'').trim()] ?? '';
                if (!bc) miss.push('บาร์โค้ด');
                const qty = p.qty != null ? +p.qty : 0;
                if (qty <= 0) miss.push('จำนวน');
                if (p.price_ea == null && qty > 0) miss.push('ราคา');
                if (miss.length > 0) issues.push({ invoice_no: inv.data.invoice_no, no: p.no, description: p.description, missing: miss });
              }
            }
            if (issues.length === 0) return (
              <div style={{ background:'#f0fdf4', border:'1px solid #bbf7d0', borderRadius:12, padding:12, display:'flex', gap:9, alignItems:'center' }}>
                <span style={{ fontSize:15, color:'#15803d', flex:'none' }}>✓</span>
                <div style={{ flex:1, minWidth:0, fontSize:11.5, color:'#15803d', lineHeight:1.5 }}>ข้อมูลครบทุกรายการ พร้อมบันทึก</div>
              </div>
            );
            return (
              <div style={{ background:'#fffbeb', border:'1px solid #fde68a', borderRadius:12, overflow:'hidden' }}>
                <div style={{ padding:'10px 12px', fontSize:11.5, fontWeight:700, color:'#b45309' }}>
                  {issues.length} รายการข้อมูลไม่ครบ — บันทึกได้ แต่ช่องที่ขาดจะเป็นค่าว่าง
                </div>
                <div style={{ maxHeight:150, overflowY:'auto' }}>
                  {issues.slice(0, 30).map((iss, idx) => (
                    <div key={idx} style={{ padding:'8px 12px', borderTop:'1px solid #fde68a', fontSize:11, color:'#b45309', lineHeight:1.5 }}>
                      <span style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{iss.invoice_no} #{iss.no}</span>{' '}
                      <span style={{ color:'#94a3b8' }}>{(iss.description||'').slice(0,30)}</span><br/>
                      ขาด: <strong>{iss.missing.join(', ')}</strong>
                    </div>
                  ))}
                  {issues.length > 30 && <div style={{ padding:'8px 12px', fontSize:11, color:'#94a3b8' }}>และอีก {issues.length - 30} รายการ</div>}
                </div>
              </div>
            );
          })()}

          <div style={{ background:'#fff', border:'1px solid #e4e6ea', borderRadius:12, padding:12 }}>
            <div style={{ fontSize:12, fontWeight:700, color:'#334155', marginBottom:7 }}>ชื่อไฟล์</div>
            <div style={{ display:'flex', gap:8, alignItems:'center' }}>
              <input value={fileName} onChange={e=>setFileName(e.target.value)} placeholder="260517xxxx"
                style={{ flex:1, minWidth:0, minHeight:44, padding:'0 12px', borderRadius:10, fontSize:14, fontFamily:"'IBM Plex Mono', monospace",
                         border:`1px solid ${nameStatus==='duplicate'?'#b91c1c':nameStatus==='available'?'#15803d':'#e4e6ea'}`,
                         background: nameStatus==='duplicate'?'#fef2f2':'#fff' }}/>
              {nameStatus==='checking'&&<Spinner size={16}/>}
              {nameStatus==='duplicate'&&<span style={{ flex:'none', color:'#b91c1c', fontSize:11, fontWeight:700 }}>ซ้ำ</span>}
              {nameStatus==='available'&&<span style={{ flex:'none', color:'#15803d', fontSize:13, fontWeight:700 }}>✓</span>}
            </div>
          </div>

          {/* ส่งให้ผู้จัดการอนุมัติ — งานหลักของขั้นนี้ */}
          <div style={{ background:'#fff', border:'1px solid #e4e6ea', borderRadius:12, padding:12, display:'flex', flexDirection:'column', gap:9 }}>
            <div style={{ fontSize:12, fontWeight:700, color:'#334155' }}>ส่งให้ผู้จัดการอนุมัติ</div>
            {sendSt.done ? (
              <div style={{ background:'#f0fdf4', border:'1px solid #bbf7d0', borderRadius:10, padding:12, display:'flex', flexDirection:'column', gap:7 }}>
                <div style={{ fontSize:13, fontWeight:700, color:'#15803d' }}>ส่งเรียบร้อย {sendSt.done.length} ใบ</div>
                {sendSt.done.filter(Boolean).map(no => (
                  <div key={no} style={{ fontSize:12, fontWeight:700, color:'#15803d', fontFamily:"'IBM Plex Mono', monospace" }}>{no}</div>
                ))}
                <div style={{ fontSize:11, color:'#15803d', lineHeight:1.5 }}>ผู้จัดการจะตรวจแล้วบันทึกเข้าระบบให้ · ถ้ามีจุดผิดจะส่งกลับมา</div>
                <button onClick={reset}
                  style={{ minHeight:46, borderRadius:11, border:'none', background:'#2f6e90', color:'#fff', fontFamily:'inherit', fontSize:14, fontWeight:700, cursor:'pointer' }}>เริ่มบิลใบใหม่</button>
              </div>
            ) : (
              <>
                <button onClick={sendInvoicesForReview} disabled={sendSt.busy||doneInvs.length===0}
                  style={{ width:'100%', minHeight:54, borderRadius:12, border:'none', fontFamily:'inherit', fontSize:15.5, fontWeight:700, color:'#fff',
                           background: (sendSt.busy||doneInvs.length===0)?'#94a3b8':'#2f6e90',
                           boxShadow: (sendSt.busy||doneInvs.length===0)?'none':'0 2px 0 #255771',
                           cursor:(sendSt.busy||doneInvs.length===0)?'not-allowed':'pointer',
                           display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>
                  {sendSt.busy?<><Spinner size={18}/>กำลังส่ง…</>:`ส่งให้ผู้จัดการ (${doneInvs.length} ใบ)`}
                </button>
                {sendSt.err&&(
                  <div style={{ background:'#fef2f2', border:'1px solid #fecaca', borderRadius:10, padding:10, fontSize:11.5, color:'#b91c1c', lineHeight:1.5 }}>{sendSt.err}</div>
                )}
                <div style={{ fontSize:11, color:'#94a3b8', lineHeight:1.5 }}>ส่งแล้วแก้ไม่ได้ — ผู้จัดการเป็นคนบันทึกเข้าระบบ</div>
              </>
            )}
          </div>

          {/* ส่งออกไฟล์ */}
          <div style={{ background:'#fff', border:'1px solid #e4e6ea', borderRadius:12, padding:12, display:'flex', flexDirection:'column', gap:9 }}>
            <div style={{ fontSize:12, fontWeight:700, color:'#334155' }}>เก็บไฟล์ไว้ดูเอง (ไม่บังคับ)</div>
            <button onClick={()=>{const b=buildXLSXBlob();if(b)downloadBlob(b,(fileName||'invoice')+'.xlsx');}}
              style={{ width:'100%', minHeight:48, borderRadius:11, border:'1px solid #e4e6ea', background:'#fff', color:'#334155', fontFamily:'inherit', fontSize:14, fontWeight:700, cursor:'pointer' }}>
              ดาวน์โหลด Excel
            </button>
            <div style={{ fontSize:11, color:'#94a3b8', lineHeight:1.55 }}>
              ไฟล์เข้า Drive อัตโนมัติหลังผู้จัดการอนุมัติ — พนักงานไม่ต้องอัปโหลดเอง
            </div>
          </div>

          <button onClick={()=>setStep(3)}
            style={{ alignSelf:'flex-start', minHeight:44, padding:'0 14px', borderRadius:11, border:'1px solid #e4e6ea', background:'#fff', color:'#475569', fontFamily:'inherit', fontWeight:600, fontSize:13, cursor:'pointer' }}>‹ กลับไปแก้ข้อมูล</button>
        </div>
      )}
      </div>
      <StepBar current={step} maxStep={maxStep} sentBadge={sentBadge} sentActive={sentActive}
        onGo={(n)=>{ setStep(n); onCloseSent?.(); }}
        onOpenSent={onOpenSent}/>
    </div>
  );
}

function ScannerModal({ products, onScan, onClose }) {
  const [tab, setTab] = useState('camera');
  const [manualBarcode, setManualBarcode] = useState('');
  const [scanError, setScanError] = useState('');
  const [scanning, setScanning] = useState(false);
  const [rematching, setRematching] = useState(false);
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
      <div className="bg-white rounded-t-2xl sm:rounded-2xl w-full max-w-md max-h-[95dvh] overflow-y-auto">
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
