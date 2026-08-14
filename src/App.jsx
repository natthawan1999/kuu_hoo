// ═══════════════════════════════════════════════════════════════════
// KUUHOO — หน้ารายงาน (เพิ่มลง src/App.jsx)
// ═══════════════════════════════════════════════════════════════════
//
// ทำ 3 ขั้น ตามลำดับ:
//
// ─── ขั้น 1: เพิ่มเมนู ─────────────────────────────────────────────
// บรรทัด 524-525 — เพิ่ม report เข้าไปก่อน settings ทั้ง 2 บรรทัด
// (ของเดิม)  ...,{ id:'settings',label:'ตั้งค่า',icon:SettingsIcon }]
// (ของใหม่)  ...,{ id:'report',label:'รายงาน',icon:FileSpreadsheet },{ id:'settings',label:'ตั้งค่า',icon:SettingsIcon }]
//
// ─── ขั้น 2: เพิ่มบรรทัดเรียกหน้าจอ ────────────────────────────────
// บรรทัด 562 — แทรกบรรทัดนี้ไว้เหนือบรรทัด view === 'settings'
//        {isManager && view === 'report' && <ReportView />}
//
// ─── ขั้น 3: วาง component ข้างล่างทั้งก้อนนี้ ──────────────────────
// วางไว้ก่อน  function SettingsView(...)  (บรรทัด ~1321)
//
// ไม่ต้อง import อะไรเพิ่ม — ใช้ไอคอนกับ XLSX ที่ import ไว้แล้ว
// ═══════════════════════════════════════════════════════════════════

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
      <div className="bg-white rounded-xl border border-slate-200 p-4 space-y-4">
        <div className="flex flex-wrap gap-2">
          {REPORT_TOPICS.map(t => (
            <button key={t.id} onClick={() => { setTopic(t.id); setRows(null); }}
              className={`px-3 py-2 rounded-lg text-sm font-medium border transition-colors ${
                topic === t.id
                  ? 'bg-slate-800 border-slate-800 text-white'
                  : 'bg-white border-slate-300 text-slate-600 hover:bg-slate-50'}`}>
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
                  className="mt-1 w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
              </label>
              <label className="block">
                <span className="text-xs font-medium text-slate-500">ถึง</span>
                <input type="date" value={to} onChange={e => setTo(e.target.value)}
                  className="mt-1 w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
              </label>
            </>
          )}
          {conf.needDate && (
            <label className="block">
              <span className="text-xs font-medium text-slate-500">เลขที่เอกสาร</span>
              <input value={doc} onChange={e => setDoc(e.target.value)} placeholder="ทั้งหมด"
                className="mt-1 w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
            </label>
          )}
          <label className="block">
            <span className="text-xs font-medium text-slate-500">บาร์โค้ด / ชื่อสินค้า</span>
            <input value={barcode} onChange={e => setBarcode(e.target.value)} placeholder="ทั้งหมด"
              className="mt-1 w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
          </label>
          {(topic === 'invoice' || topic === 'in' || topic === 'out') && (
            <label className="block">
              <span className="text-xs font-medium text-slate-500">
                {topic === 'out' ? 'ลูกค้า' : 'ผู้ขาย'}
              </span>
              <input value={party} onChange={e => setParty(e.target.value)} placeholder="ทั้งหมด"
                className="mt-1 w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
            </label>
          )}
          {topic === 'stock' && (
            <label className="block">
              <span className="text-xs font-medium text-slate-500">ประเภทสินค้า</span>
              <input value={party} onChange={e => setParty(e.target.value)} placeholder="ทั้งหมด"
                className="mt-1 w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
            </label>
          )}
        </div>

        <button onClick={run} disabled={loading}
          className="w-full bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-300 text-white py-3 rounded-xl font-semibold flex items-center justify-center gap-2">
          {loading ? <><RefreshCw size={18} className="animate-spin" />กำลังดึง…</>
                   : <><Search size={18} />ดึงรายงาน</>}
        </button>

        {err && (
          <div className="bg-red-50 border border-red-200 text-red-700 rounded-lg px-3 py-2 text-sm flex items-start gap-2">
            <AlertCircle size={16} className="mt-0.5 shrink-0" />{err}
          </div>
        )}
      </div>

      {/* ผลลัพธ์ */}
      {rows && (
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <div className="px-4 py-3 border-b border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-2">
            <div className="text-sm font-semibold text-slate-800">
              {conf.label} · {shown.length.toLocaleString('th-TH')} แถว
              {shown.length !== rows.length && (
                <span className="text-slate-400 font-normal"> (จาก {rows.length.toLocaleString('th-TH')})</span>
              )}
            </div>
            <div className="flex gap-2">
              {Object.values(colFilter).some(v => v?.trim()) && (
                <button onClick={() => setColFilter({})}
                  className="px-3 py-1.5 rounded-lg text-xs font-medium border border-slate-300 text-slate-600 hover:bg-white flex items-center gap-1">
                  <X size={12} />ล้างตัวกรอง
                </button>
              )}
              <button onClick={exportCsv} disabled={!shown.length}
                className="px-3 py-1.5 rounded-lg text-xs font-medium border border-slate-300 text-slate-700 hover:bg-white disabled:opacity-40 flex items-center gap-1">
                <Download size={12} />CSV
              </button>
              <button onClick={exportXlsx} disabled={!shown.length}
                className="px-3 py-1.5 rounded-lg text-xs font-medium bg-emerald-600 hover:bg-emerald-700 text-white disabled:opacity-40 flex items-center gap-1">
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
                      <th key={k} className="px-2 pb-2 border-b border-slate-200">
                        <input value={colFilter[k] || ''}
                          onChange={e => setColFilter(f => ({ ...f, [k]: e.target.value }))}
                          placeholder="กรอง"
                          className="w-full min-w-[70px] border border-slate-200 rounded px-1.5 py-1 text-[11px] font-normal focus:border-indigo-400 focus:outline-none" />
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {shown.map((r, i) => (
                    <tr key={i} className={i % 2 ? 'bg-slate-50/60' : ''}>
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
                  <tfoot className="sticky bottom-0 bg-white border-t-2 border-slate-300">
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
