import { useState, useRef, useCallback, useMemo } from "react";

// ── Palette ──────────────────────────────────────────────────────────────────
const C = {
  bg: "#f1f5f9",
  surface: "#ffffff",
  surfaceHi: "#f8fafc",
  border: "#e2e8f0",
  accent: "#4f6ef7",
  accentLo: "#eff2fe",
  green: "#16a34a",
  greenLo: "#dcfce7",
  amber: "#d97706",
  amberLo: "#fef3c7",
  red: "#dc2626",
  redLo: "#fee2e2",
  text: "#0f172a",
  muted: "#64748b",
  mono: "'Courier New', monospace",
};

// ── Seed Data ────────────────────────────────────────────────────────────────
const PRODUCTS = [
  { barcode: "8851234567001", name: "น้ำดื่มสิงห์ 600ml",    unit: "ขวด",    price: 10,  cost: 7,   stock: 245 },
  { barcode: "8851234567002", name: "น้ำดื่มคริสตัล 600ml", unit: "ขวด",    price: 9,   cost: 6,   stock: 180 },
  { barcode: "8851234567003", name: "โค้ก 325ml",            unit: "กระป๋อง",price: 17,  cost: 12,  stock: 120 },
  { barcode: "8851234567004", name: "เป๊ปซี่ 325ml",          unit: "กระป๋อง",price: 17,  cost: 12,  stock: 95  },
  { barcode: "8851234567005", name: "มาม่าหมูสับ",            unit: "ซอง",    price: 6,   cost: 4,   stock: 320 },
  { barcode: "8851234567006", name: "มาม่าต้มยำกุ้ง",         unit: "ซอง",    price: 6,   cost: 4,   stock: 280 },
  { barcode: "8851234567007", name: "เลย์รสต้นตำรับ 50g",    unit: "ซอง",    price: 20,  cost: 14,  stock: 88  },
  { barcode: "8851234567008", name: "แชมพูแพนทีน 340ml",     unit: "ขวด",    price: 159, cost: 110, stock: 45  },
  { barcode: "8851234567009", name: "สบู่โปรเทคส์ 65g",      unit: "ก้อน",   price: 15,  cost: 10,  stock: 156 },
  { barcode: "8851234567010", name: "นมเมจิ 225ml",           unit: "กล่อง",  price: 17,  cost: 12,  stock: 200 },
];

const DEMO_SUBMISSIONS = [
  {
    id: "sub001", docNo: "RC-202505170001", counter: "สมหญิง",
    submittedAt: new Date(Date.now() - 3600000 * 2).toISOString(),
    status: "pending", itemCount: 5, totalQty: 142, note: "นับโซน A เสร็จแล้ว",
    data: [
      { barcode: "8851234567001", productName: "น้ำดื่มสิงห์ 600ml",  qty: 40, unit: "ขวด",    location: "A1" },
      { barcode: "8851234567003", productName: "โค้ก 325ml",           qty: 28, unit: "กระป๋อง",location: "A2" },
      { barcode: "8851234567005", productName: "มาม่าหมูสับ",          qty: 35, unit: "ซอง",    location: "A3" },
      { barcode: "8851234567006", productName: "มาม่าต้มยำกุ้ง",      qty: 29, unit: "ซอง",    location: "A3" },
      { barcode: "8851234567009", productName: "สบู่โปรเทคส์ 65g",    qty: 10, unit: "ก้อน",   location: "A4" },
    ],
  },
  {
    id: "sub002", docNo: "RC-202505170002", counter: "สมชาย",
    submittedAt: new Date(Date.now() - 3600000 * 5).toISOString(),
    status: "approved", itemCount: 3, totalQty: 77, note: "",
    data: [
      { barcode: "8851234567002", productName: "น้ำดื่มคริสตัล 600ml",qty: 30, unit: "ขวด",    location: "B1" },
      { barcode: "8851234567007", productName: "เลย์รสต้นตำรับ 50g",  qty: 20, unit: "ซอง",    location: "B2" },
      { barcode: "8851234567010", productName: "นมเมจิ 225ml",         qty: 27, unit: "กล่อง",  location: "B3" },
    ],
  },
];

// ── Tiny helpers ──────────────────────────────────────────────────────────────
const fmt = n => Number(n).toLocaleString("th-TH");
const timeStr = iso => new Date(iso).toLocaleString("th-TH", { hour: "2-digit", minute: "2-digit", day: "2-digit", month: "short" });

function Badge({ color = "accent", children }) {
  const map = { accent: [C.accentLo, C.accent], green: [C.greenLo, C.green], amber: [C.amberLo, C.amber], red: [C.redLo, C.red], muted: ["#1e2433", C.muted] };
  const [bg, fg] = map[color] || map.accent;
  return <span style={{ background: bg, color: fg, fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: 99, letterSpacing: "0.05em", textTransform: "uppercase" }}>{children}</span>;
}

function Card({ children, style = {} }) {
  return <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, padding: 16, ...style }}>{children}</div>;
}

function Btn({ children, onClick, disabled, color = "accent", size = "md", full = false, style = {} }) {
  const bg = { accent: C.accent, green: "#16a34a", red: "#dc2626", ghost: C.surfaceHi }[color];
  const pad = size === "sm" ? "6px 14px" : size === "lg" ? "14px 28px" : "10px 20px";
  return (
    <button onClick={onClick} disabled={disabled}
      style={{ background: disabled ? C.border : bg, color: disabled ? C.muted : "#fff", border: "none", borderRadius: 8, padding: pad, fontWeight: 700, fontSize: size === "sm" ? 12 : 14, cursor: disabled ? "not-allowed" : "pointer", width: full ? "100%" : undefined, opacity: disabled ? 0.6 : 1, transition: "all .15s", ...style }}>
      {children}
    </button>
  );
}

function Input({ value, onChange, placeholder, type = "text", mono = false, style = {} }) {
  return <input type={type} value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder}
    style={{ background: C.surfaceHi, border: `1px solid ${C.border}`, color: C.text, borderRadius: 8, padding: "10px 14px", fontSize: 13, width: "100%", outline: "none", fontFamily: mono ? C.mono : undefined, boxSizing: "border-box", ...style }} />;
}

// ── LOGIN — 3-step flow ──────────────────────────────────────────────────────
const FEATURES = [
  {
    id: "stock_compare",
    icon: "⚖️",
    label: "นับ Stock & Compare",
    sub: "นับสต็อก + เปรียบเทียบกับฐานข้อมูล",
    tags: ["นับสต็อก", "เปรียบเทียบ", "รีวิว"],
    accentColor: "#4f6ef7",
    roles: ["counter", "manager"],
  },
  {
    id: "stock_only",
    icon: "📦",
    label: "Record Helper",
    sub: "บันทึกและส่งให้ผู้จัดการ ไม่เปรียบเทียบ",
    tags: ["นับสต็อก", "รีวิว"],
    accentColor: "#22c55e",
    roles: ["counter", "manager"],
  },
  {
    id: "invoice",
    icon: "🧾",
    label: "Invoice Scanner (AI)",
    sub: "อ่านใบกำกับภาษีด้วย Claude AI",
    tags: ["AI", "OCR", "Export"],
    accentColor: "#a855f7",
    roles: ["counter", "manager"],
  },
];

function Login({ onLogin }) {
  const [step, setStep] = useState(1);   // 1=role, 2=feature, 3=name
  const [role, setRole] = useState(null);
  const [feature, setFeature] = useState(null);
  const [name, setName] = useState("");

  const roleCards = [
    { id: "counter", emoji: "👷", label: "พนักงาน",  sub: "นับสต็อก · Invoice · ส่งรายการ", color: "#22c55e" },
    { id: "manager", emoji: "👔", label: "ผู้จัดการ", sub: "รีวิว · เปรียบเทียบ · Export",   color: "#4f6ef7" },
  ];

  const availableFeatures = FEATURES.filter(f => f.roles.includes(role));
  const selectedCard = roleCards.find(r => r.id === role);
  const accentCol = feature?.accentColor || selectedCard?.color || C.accent;

  const goStep1 = () => { setStep(1); setRole(null); setFeature(null); setName(""); };
  const goStep2 = (r) => { setRole(r); setFeature(null); setStep(2); };
  const goStep3 = (f) => { setFeature(f); setStep(3); };

  const dots = [1, 2, 3];

  return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ width: "100%", maxWidth: 400 }}>

        {/* Logo */}
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{
            display: "inline-flex", alignItems: "center", justifyContent: "center",
            width: 64, height: 64, borderRadius: 18,
            background: "linear-gradient(135deg, #4f6ef7 0%, #a855f7 100%)",
            fontSize: 30, marginBottom: 12, boxShadow: "0 0 32px #4f6ef740",
          }}>📦</div>
          <h1 style={{ color: C.text, fontSize: 34, fontWeight: 900, letterSpacing: "-0.04em", margin: "0 0 4px" }}>KUUHOO</h1>
          <p style={{ color: C.muted, fontSize: 12, margin: 0 }}>Inventory Management System · Demo</p>
        </div>

        {/* Step dots */}
        <div style={{ display: "flex", justifyContent: "center", gap: 6, marginBottom: 28 }}>
          {dots.map(s => (
            <div key={s} style={{
              width: s === step ? 22 : 7, height: 7, borderRadius: 99,
              background: s < step ? C.green : s === step ? accentCol : C.border,
              transition: "all .3s",
            }} />
          ))}
        </div>

        {/* Step 1 — Role */}
        {step === 1 && (
          <div>
            <p style={{ color: C.muted, fontSize: 12, textAlign: "center", marginBottom: 16 }}>เลือกบทบาทของคุณ</p>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              {roleCards.map(r => (
                <button key={r.id} onClick={() => goStep2(r.id)}
                  style={{
                    background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16,
                    padding: "28px 16px", cursor: "pointer", textAlign: "center",
                    display: "flex", flexDirection: "column", alignItems: "center", gap: 12,
                    transition: "all .2s",
                  }}
                  onMouseEnter={e => { e.currentTarget.style.borderColor = r.color; e.currentTarget.style.background = C.surfaceHi; }}
                  onMouseLeave={e => { e.currentTarget.style.borderColor = C.border; e.currentTarget.style.background = C.surface; }}>
                  <div style={{
                    width: 56, height: 56, borderRadius: 16, fontSize: 28,
                    display: "flex", alignItems: "center", justifyContent: "center",
                    background: r.color + "20", border: `2px solid ${r.color}40`,
                  }}>{r.emoji}</div>
                  <div>
                    <div style={{ color: C.text, fontWeight: 800, fontSize: 17 }}>{r.label}</div>
                    <div style={{ color: C.muted, fontSize: 10, marginTop: 4, lineHeight: 1.5 }}>{r.sub}</div>
                  </div>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Step 2 — Feature */}
        {step === 2 && (
          <div>
            <button onClick={goStep1}
              style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 12, marginBottom: 14, padding: 0 }}>
              ← เปลี่ยนบทบาท
            </button>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
              <span style={{ fontSize: 18 }}>{selectedCard?.emoji}</span>
              <p style={{ color: C.muted, fontSize: 12, margin: 0 }}>
                <b style={{ color: C.text }}>{selectedCard?.label}</b> — เลือก Feature
              </p>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {availableFeatures.map(f => (
                <button key={f.id} onClick={() => goStep3(f)}
                  style={{
                    background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14,
                    padding: "16px 18px", cursor: "pointer", textAlign: "left",
                    display: "flex", alignItems: "center", gap: 14, transition: "all .2s",
                  }}
                  onMouseEnter={e => { e.currentTarget.style.borderColor = f.accentColor; e.currentTarget.style.background = C.surfaceHi; }}
                  onMouseLeave={e => { e.currentTarget.style.borderColor = C.border; e.currentTarget.style.background = C.surface; }}>
                  <div style={{
                    width: 46, height: 46, borderRadius: 12, fontSize: 22,
                    display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                    background: f.accentColor + "20", border: `1.5px solid ${f.accentColor}40`,
                  }}>{f.icon}</div>
                  <div style={{ flex: 1 }}>
                    <div style={{ color: C.text, fontWeight: 700, fontSize: 14 }}>{f.label}</div>
                    <div style={{ color: C.muted, fontSize: 11, marginTop: 3 }}>{f.sub}</div>
                    <div style={{ display: "flex", gap: 4, marginTop: 6, flexWrap: "wrap" }}>
                      {f.tags.map(t => (
                        <span key={t} style={{ background: f.accentColor + "20", color: f.accentColor, fontSize: 9, fontWeight: 700, padding: "2px 7px", borderRadius: 99 }}>{t}</span>
                      ))}
                    </div>
                  </div>
                  <span style={{ color: C.muted, fontSize: 20 }}>›</span>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Step 3 — Name */}
        {step === 3 && feature && (
          <div>
            <button onClick={() => setStep(2)}
              style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 12, marginBottom: 16, padding: 0 }}>
              ← เปลี่ยน Feature
            </button>

            {/* Summary badge */}
            <div style={{
              display: "flex", alignItems: "center", gap: 12, marginBottom: 20,
              background: accentCol + "12", border: `1px solid ${accentCol}30`,
              borderRadius: 12, padding: "12px 16px",
            }}>
              <span style={{ fontSize: 24 }}>{feature.icon}</span>
              <div>
                <div style={{ color: C.text, fontWeight: 700, fontSize: 13 }}>{feature.label}</div>
                <div style={{ color: C.muted, fontSize: 11 }}>{selectedCard?.emoji} {selectedCard?.label}</div>
              </div>
            </div>

            <div style={{ marginBottom: 16 }}>
              <label style={{ color: C.muted, fontSize: 11, display: "block", marginBottom: 6 }}>ชื่อของคุณ</label>
              <Input value={name} onChange={setName} placeholder="เช่น สมหญิง"
                style={{ fontSize: 15, padding: "12px 14px" }} />
            </div>

            <button
              onClick={() => name.trim() && onLogin({ name: name.trim(), role, feature: feature.id })}
              disabled={!name.trim()}
              style={{
                width: "100%", padding: "14px", borderRadius: 12, border: "none",
                background: name.trim() ? `linear-gradient(135deg, ${accentCol}, ${accentCol}bb)` : C.border,
                color: name.trim() ? "#fff" : C.muted,
                fontWeight: 800, fontSize: 15, cursor: name.trim() ? "pointer" : "not-allowed",
                boxShadow: name.trim() ? `0 4px 20px ${accentCol}40` : "none",
                transition: "all .2s",
              }}>
              เข้าสู่ระบบ →
            </button>
          </div>
        )}

      </div>
    </div>
  );
}

// ── COUNTER: Count View ───────────────────────────────────────────────────────
function CountView({ user, entries, onAdd, onDelete, onNext }) {
  const [barcode, setBarcode] = useState("");
  const [qty, setQty] = useState("");
  const [loc, setLoc] = useState("");
  const [found, setFound] = useState(null);
  const [notFound, setNotFound] = useState(false);
  const [err, setErr] = useState("");

  const lookup = () => {
    const code = barcode.trim();
    if (!code) return;
    const p = PRODUCTS.find(x => x.barcode === code);
    if (p) { setFound(p); setNotFound(false); setErr(""); }
    else { setFound(null); setNotFound(true); setErr(`ไม่พบ "${code}" ในระบบ`); }
  };

  const addEntry = () => {
    if (!loc.trim()) { setErr("กรุณาระบุ Location"); return; }
    if (!qty || parseInt(qty) <= 0) return;
    onAdd({
      id: `e${Date.now()}`,
      barcode: found ? found.barcode : barcode.trim(),
      productName: found ? found.name : "(ไม่พบในระบบ)",
      unit: found?.unit || "", price: found?.price || 0, cost: found?.cost || 0,
      qty: parseInt(qty), location: loc.trim(), notFound: !found,
      timestamp: new Date().toISOString(),
    });
    setBarcode(""); setQty(""); setFound(null); setNotFound(false); setErr("");
  };

  const quickFill = (p) => { setBarcode(p.barcode); setFound(p); setNotFound(false); setErr(""); };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      {/* Stats */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Card style={{ textAlign: "center", padding: 12 }}>
          <div style={{ color: C.muted, fontSize: 10, marginBottom: 4 }}>รายการ</div>
          <div style={{ color: C.text, fontSize: 28, fontWeight: 900 }}>{entries.length}</div>
        </Card>
        <Card style={{ textAlign: "center", padding: 12 }}>
          <div style={{ color: C.muted, fontSize: 10, marginBottom: 4 }}>จำนวนรวม</div>
          <div style={{ color: C.green, fontSize: 28, fontWeight: 900 }}>{fmt(entries.reduce((s, e) => s + e.qty, 0))}</div>
        </Card>
      </div>

      {/* Quick pick */}
      <Card>
        <p style={{ color: C.muted, fontSize: 11, marginBottom: 8 }}>⚡ Demo — เลือกสินค้าตัวอย่าง</p>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
          {PRODUCTS.slice(0, 6).map(p => (
            <button key={p.barcode} onClick={() => quickFill(p)}
              style={{ background: C.surfaceHi, border: `1px solid ${C.border}`, color: C.text, borderRadius: 6, padding: "4px 10px", fontSize: 11, cursor: "pointer" }}>
              {p.name.slice(0, 14)}
            </button>
          ))}
        </div>
      </Card>

      {/* Input */}
      <Card>
        <div style={{ marginBottom: 12 }}>
          <label style={{ color: C.muted, fontSize: 11, display: "block", marginBottom: 6 }}>บาร์โค้ด</label>
          <div style={{ display: "flex", gap: 8 }}>
            <Input value={barcode} onChange={v => { setBarcode(v); setFound(null); setNotFound(false); setErr(""); }} placeholder="8851234567001..." mono />
            <Btn onClick={lookup} size="sm">ตรวจ</Btn>
          </div>
        </div>

        {found && (
          <div style={{ background: C.greenLo, border: `1px solid ${C.green}30`, borderRadius: 8, padding: 10, marginBottom: 12 }}>
            <div style={{ color: C.green, fontSize: 11, fontWeight: 700 }}>✓ พบสินค้า</div>
            <div style={{ color: C.text, fontWeight: 700, marginTop: 2 }}>{found.name}</div>
            <div style={{ color: C.muted, fontSize: 11 }}>{found.unit} · ราคา ฿{found.price}</div>
          </div>
        )}

        {err && (
          <div style={{ background: C.redLo, border: `1px solid ${C.red}30`, borderRadius: 8, padding: 10, marginBottom: 12, color: C.red, fontSize: 12 }}>{err}</div>
        )}

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={{ color: C.muted, fontSize: 11, display: "block", marginBottom: 6 }}>Location *</label>
            <Input value={loc} onChange={setLoc} placeholder="A1, ชั้น 2..." />
          </div>
          <div>
            <label style={{ color: C.muted, fontSize: 11, display: "block", marginBottom: 6 }}>จำนวน</label>
            <Input value={qty} onChange={setQty} placeholder="0" type="number" />
          </div>
        </div>

        <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
          {[1, 5, 10, 20, 50].map(n => (
            <button key={n} onClick={() => setQty(String(n))}
              style={{ flex: 1, background: C.surfaceHi, border: `1px solid ${C.border}`, color: C.text, borderRadius: 6, padding: "6px 0", fontSize: 12, cursor: "pointer" }}>{n}</button>
          ))}
        </div>

        <Btn full onClick={addEntry} disabled={(!found && !notFound) || !qty || parseInt(qty) <= 0} color={notFound ? "ghost" : "green"}>
          {notFound ? "+ นับไว้ก่อน (ไม่มีในระบบ)" : "+ เพิ่มในรายการ"}
        </Btn>
      </Card>

      {/* List */}
      {entries.length > 0 && (
        <Card style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "10px 14px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ color: C.text, fontSize: 13, fontWeight: 700 }}>รายการที่นับ</span>
            <Btn size="sm" color="accent" onClick={onNext}>ตรวจสอบ & ส่ง →</Btn>
          </div>
          {entries.map(e => (
            <div key={e.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", borderBottom: `1px solid ${C.border}20` }}>
              <div style={{ flex: 1 }}>
                <div style={{ color: C.text, fontSize: 12, fontWeight: 600 }}>{e.productName}</div>
                <div style={{ color: C.muted, fontSize: 10, fontFamily: C.mono }}>{e.barcode} · {e.location}</div>
              </div>
              <div style={{ color: e.notFound ? C.amber : C.green, fontWeight: 900, fontSize: 18 }}>{e.qty}</div>
              <button onClick={() => onDelete(e.id)} style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 14 }}>✕</button>
            </div>
          ))}
        </Card>
      )}
    </div>
  );
}

// ── COUNTER: Review & Submit ──────────────────────────────────────────────────
function ReviewView({ entries, user, onSubmit, onBack }) {
  const [note, setNote] = useState("");
  const [done, setDone] = useState(null);

  const grouped = useMemo(() => {
    const map = {};
    entries.forEach(e => {
      if (map[e.barcode]) { map[e.barcode].qty += e.qty; map[e.barcode].scans++; }
      else map[e.barcode] = { ...e, scans: 1 };
    });
    return Object.values(map);
  }, [entries]);

  const totalQty = grouped.reduce((s, g) => s + g.qty, 0);

  const submit = () => {
    const sub = {
      id: `sub${Date.now()}`, docNo: `RC-${new Date().toISOString().slice(0,10).replace(/-/g,"")}${String(Math.floor(Math.random()*9000)+1000)}`,
      counter: user.name, submittedAt: new Date().toISOString(),
      status: "pending", itemCount: grouped.length, totalQty, note, data: grouped,
    };
    onSubmit(sub);
    setDone(sub);
  };

  if (done) return (
    <Card style={{ textAlign: "center", padding: 32 }}>
      <div style={{ fontSize: 48, marginBottom: 12 }}>✅</div>
      <h2 style={{ color: C.green, fontSize: 20, fontWeight: 900, margin: "0 0 8px" }}>ส่งเรียบร้อย!</h2>
      <div style={{ color: C.muted, fontSize: 12, marginBottom: 16 }}>เลขที่ <span style={{ color: C.accent, fontFamily: C.mono }}>{done.docNo}</span></div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 20 }}>
        <div style={{ background: C.surfaceHi, borderRadius: 8, padding: 12 }}>
          <div style={{ color: C.muted, fontSize: 10 }}>บาร์โค้ด</div>
          <div style={{ color: C.text, fontSize: 22, fontWeight: 900 }}>{done.itemCount}</div>
        </div>
        <div style={{ background: C.surfaceHi, borderRadius: 8, padding: 12 }}>
          <div style={{ color: C.muted, fontSize: 10 }}>จำนวนรวม</div>
          <div style={{ color: C.green, fontSize: 22, fontWeight: 900 }}>{done.totalQty}</div>
        </div>
      </div>
      <Btn full color="accent" onClick={onBack}>นับรอบใหม่</Btn>
    </Card>
  );

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Card style={{ textAlign: "center", padding: 12 }}>
          <div style={{ color: C.muted, fontSize: 10 }}>บาร์โค้ด</div>
          <div style={{ color: C.text, fontSize: 26, fontWeight: 900 }}>{grouped.length}</div>
        </Card>
        <Card style={{ textAlign: "center", padding: 12 }}>
          <div style={{ color: C.muted, fontSize: 10 }}>จำนวนรวม</div>
          <div style={{ color: C.green, fontSize: 26, fontWeight: 900 }}>{fmt(totalQty)}</div>
        </Card>
      </div>

      <Card style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "10px 14px", borderBottom: `1px solid ${C.border}` }}>
          <span style={{ color: C.text, fontSize: 13, fontWeight: 700 }}>รายการรวม</span>
        </div>
        {grouped.map(g => (
          <div key={g.barcode} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", borderBottom: `1px solid ${C.border}20` }}>
            <div style={{ flex: 1 }}>
              <div style={{ color: C.text, fontSize: 12, fontWeight: 600 }}>{g.productName}</div>
              <div style={{ color: C.muted, fontSize: 10, fontFamily: C.mono }}>{g.barcode}</div>
              {g.scans > 1 && <Badge color="muted">{g.scans} ครั้ง</Badge>}
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ color: g.notFound ? C.amber : C.green, fontWeight: 900, fontSize: 20 }}>{g.qty}</div>
              <div style={{ color: C.muted, fontSize: 10 }}>{g.unit}</div>
            </div>
          </div>
        ))}
      </Card>

      <Card>
        <label style={{ color: C.muted, fontSize: 11, display: "block", marginBottom: 6 }}>หมายเหตุถึงผู้จัดการ</label>
        <textarea value={note} onChange={e => setNote(e.target.value)} rows={2}
          style={{ background: C.surfaceHi, border: `1px solid ${C.border}`, color: C.text, borderRadius: 8, padding: "10px 14px", fontSize: 13, width: "100%", resize: "none", outline: "none", boxSizing: "border-box" }}
          placeholder="นับโซน A เสร็จแล้ว..." />
      </Card>

      <Btn full color="green" size="lg" onClick={submit}>📤 ส่งให้ผู้จัดการ</Btn>
      <Btn full color="ghost" onClick={onBack}>← กลับแก้ไข</Btn>
    </div>
  );
}

// ── MANAGER: Dashboard ────────────────────────────────────────────────────────
function DashboardView({ submissions, setView }) {
  const pending = submissions.filter(s => s.status === "pending").length;
  const approved = submissions.filter(s => s.status === "approved").length;
  const totalItems = submissions.reduce((s, sub) => s + sub.itemCount, 0);
  const totalQty = submissions.reduce((s, sub) => s + sub.totalQty, 0);

  const stats = [
    { label: "รอรีวิว", value: pending, color: "amber" },
    { label: "อนุมัติแล้ว", value: approved, color: "green" },
    { label: "บาร์โค้ดรวม", value: totalItems, color: "accent" },
    { label: "จำนวนรวม", value: fmt(totalQty), color: "muted" },
  ];

  const colorMap = { amber: C.amber, green: C.green, accent: C.accent, muted: C.text };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      {pending > 0 && (
        <button onClick={() => setView("inbox")}
          style={{ background: C.amberLo, border: `1px solid ${C.amber}40`, borderRadius: 12, padding: "14px 16px", display: "flex", alignItems: "center", gap: 12, cursor: "pointer", textAlign: "left" }}>
          <span style={{ fontSize: 24 }}>📬</span>
          <div style={{ flex: 1 }}>
            <div style={{ color: C.amber, fontWeight: 700, fontSize: 14 }}>มี {pending} รายการรอรีวิว</div>
            <div style={{ color: C.muted, fontSize: 11 }}>คลิกเพื่อตรวจสอบ</div>
          </div>
          <span style={{ color: C.amber }}>→</span>
        </button>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        {stats.map(s => (
          <Card key={s.label} style={{ padding: 14 }}>
            <div style={{ color: C.muted, fontSize: 10, marginBottom: 4 }}>{s.label}</div>
            <div style={{ color: colorMap[s.color], fontSize: 26, fontWeight: 900 }}>{s.value}</div>
          </Card>
        ))}
      </div>

      <Card>
        <div style={{ color: C.text, fontWeight: 700, fontSize: 13, marginBottom: 12 }}>กิจกรรมล่าสุด</div>
        {submissions.slice(0, 5).map(s => (
          <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 10, paddingBottom: 10, marginBottom: 10, borderBottom: `1px solid ${C.border}40` }}>
            <div style={{ width: 8, height: 8, borderRadius: "50%", background: s.status === "pending" ? C.amber : s.status === "approved" ? C.green : C.red, flexShrink: 0 }} />
            <div style={{ flex: 1 }}>
              <div style={{ color: C.text, fontSize: 12, fontWeight: 600 }}>{s.counter}</div>
              <div style={{ color: C.muted, fontSize: 10 }}>{timeStr(s.submittedAt)}</div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ color: C.text, fontSize: 12, fontWeight: 700 }}>{s.itemCount} · {s.totalQty}</div>
              <Badge color={s.status === "pending" ? "amber" : s.status === "approved" ? "green" : "red"}>
                {s.status === "pending" ? "รอ" : s.status === "approved" ? "อนุมัติ" : "ส่งกลับ"}
              </Badge>
            </div>
          </div>
        ))}
      </Card>
    </div>
  );
}

// ── MANAGER: Inbox ────────────────────────────────────────────────────────────
function InboxView({ submissions, onReview }) {
  const [selected, setSelected] = useState(null);
  const [note, setNote] = useState("");
  const [tab, setTab] = useState("pending");
  const [err, setErr] = useState("");

  const filtered = submissions.filter(s => s.status === tab);

  const approve = () => { onReview(selected.id, "approved", note); setSelected(null); setNote(""); setErr(""); };
  const reject = () => {
    if (!note.trim()) { setErr("กรุณาใส่เหตุผลก่อนส่งกลับ"); return; }
    onReview(selected.id, "rejected", note); setSelected(null); setNote(""); setErr("");
  };

  const tabs = [
    { id: "pending", label: "รอรีวิว", color: "amber" },
    { id: "approved", label: "อนุมัติ", color: "green" },
    { id: "rejected", label: "ส่งกลับ", color: "red" },
  ];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "flex", borderBottom: `1px solid ${C.border}`, marginBottom: 4 }}>
        {tabs.map(t => {
          const cnt = submissions.filter(s => s.status === t.id).length;
          const cmap = { amber: C.amber, green: C.green, red: C.red };
          return (
            <button key={t.id} onClick={() => setTab(t.id)}
              style={{ flex: 1, padding: "10px", background: "none", border: "none", borderBottom: `2px solid ${tab === t.id ? cmap[t.color] : "transparent"}`, color: tab === t.id ? cmap[t.color] : C.muted, fontWeight: 700, fontSize: 12, cursor: "pointer" }}>
              {t.label} {cnt > 0 && <span style={{ color: cmap[t.color] }}>({cnt})</span>}
            </button>
          );
        })}
      </div>

      {filtered.length === 0 ? (
        <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>ไม่มีรายการ</div>
      ) : filtered.map(s => (
        <Card key={s.id}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
            <div>
              <span style={{ color: C.accent, fontFamily: C.mono, fontSize: 11, fontWeight: 700 }}>{s.docNo}</span>
              <div style={{ color: C.text, fontWeight: 700, fontSize: 14 }}>{s.counter}</div>
            </div>
            <Badge color={s.status === "pending" ? "amber" : s.status === "approved" ? "green" : "red"}>
              {s.status === "pending" ? "รอรีวิว" : s.status === "approved" ? "อนุมัติ" : "ส่งกลับ"}
            </Badge>
          </div>
          <div style={{ display: "flex", gap: 16, marginBottom: 8 }}>
            <span style={{ color: C.muted, fontSize: 12 }}><b style={{ color: C.text }}>{s.itemCount}</b> บาร์โค้ด</span>
            <span style={{ color: C.muted, fontSize: 12 }}><b style={{ color: C.green }}>{fmt(s.totalQty)}</b> ชิ้น</span>
            <span style={{ color: C.muted, fontSize: 11 }}>{timeStr(s.submittedAt)}</span>
          </div>
          {s.note && <div style={{ background: C.surfaceHi, borderRadius: 6, padding: "6px 10px", fontSize: 11, color: C.muted, marginBottom: 8, fontStyle: "italic" }}>"{s.note}"</div>}

          {/* Items preview */}
          <div style={{ background: C.surfaceHi, borderRadius: 8, padding: 8, marginBottom: 10, fontSize: 11 }}>
            {s.data.slice(0, 3).map((d, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", color: C.muted, paddingBottom: 4 }}>
                <span>{d.productName}</span>
                <span style={{ color: C.green, fontWeight: 700 }}>{d.qty} {d.unit}</span>
              </div>
            ))}
            {s.data.length > 3 && <div style={{ color: C.muted, fontSize: 10 }}>+{s.data.length - 3} รายการ...</div>}
          </div>

          {s.status === "pending" && (
            <Btn size="sm" color="accent" onClick={() => { setSelected(s); setNote(""); setErr(""); }}>รีวิว →</Btn>
          )}
        </Card>
      ))}

      {/* Review Modal */}
      {selected && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.8)", zIndex: 100, display: "flex", alignItems: "flex-end", padding: 16 }}>
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, width: "100%", maxWidth: 480, margin: "0 auto", overflow: "hidden" }}>
            <div style={{ padding: "14px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between" }}>
              <div>
                <div style={{ color: C.text, fontWeight: 700 }}>รีวิว: {selected.counter}</div>
                <div style={{ color: C.muted, fontSize: 11 }}>{selected.itemCount} รายการ · {fmt(selected.totalQty)} ชิ้น</div>
              </div>
              <button onClick={() => setSelected(null)} style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 18 }}>✕</button>
            </div>
            <div style={{ padding: 16, maxHeight: 300, overflowY: "auto" }}>
              {selected.data.map((d, i) => (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: `1px solid ${C.border}30`, fontSize: 12 }}>
                  <div><div style={{ color: C.text }}>{d.productName}</div><div style={{ color: C.muted, fontFamily: C.mono, fontSize: 10 }}>{d.barcode} · {d.location}</div></div>
                  <div style={{ color: C.green, fontWeight: 900 }}>{d.qty} {d.unit}</div>
                </div>
              ))}
            </div>
            <div style={{ padding: "12px 16px", borderTop: `1px solid ${C.border}` }}>
              <textarea value={note} onChange={e => { setNote(e.target.value); setErr(""); }} rows={2}
                style={{ background: C.surfaceHi, border: `1px solid ${err ? C.red : C.border}`, color: C.text, borderRadius: 8, padding: "8px 12px", fontSize: 12, width: "100%", resize: "none", outline: "none", boxSizing: "border-box", marginBottom: 4 }}
                placeholder="หมายเหตุ / เหตุผล (บังคับถ้าส่งกลับ)..." />
              {err && <div style={{ color: C.red, fontSize: 11, marginBottom: 8 }}>{err}</div>}
              <div style={{ display: "flex", gap: 8 }}>
                <Btn full color="red" onClick={reject}>✕ ส่งกลับ</Btn>
                <Btn full color="green" onClick={approve}>✓ อนุมัติ</Btn>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── MANAGER: Compare (simplified) ────────────────────────────────────────────
function CompareView({ submissions }) {
  const [sel, setSel] = useState(null);
  const approved = submissions.filter(s => s.status === "approved");

  const compareData = useMemo(() => {
    if (!sel) return [];
    return sel.data.map(d => {
      const master = PRODUCTS.find(p => p.barcode === d.barcode);
      const sale = Math.floor(Math.random() * 5);
      const purchase = 0;
      const adjCount = d.qty - sale + purchase;
      const stockNow = master?.stock ?? null;
      const adjStock = stockNow !== null ? adjCount - stockNow : null;
      return { ...d, sale, purchase, adjCount, stockNow, adjStock, found: !!master };
    });
  }, [sel]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <Card>
        <div style={{ color: C.muted, fontSize: 11, marginBottom: 8 }}>เลือก submission ที่อนุมัติแล้ว</div>
        {approved.length === 0 ? (
          <div style={{ color: C.muted, fontSize: 12, textAlign: "center", padding: 20 }}>ยังไม่มีรายการที่อนุมัติ — ไปรีวิวก่อน</div>
        ) : approved.map(s => (
          <button key={s.id} onClick={() => setSel(s)}
            style={{ display: "block", width: "100%", textAlign: "left", background: sel?.id === s.id ? C.accentLo : C.surfaceHi, border: `1px solid ${sel?.id === s.id ? C.accent : C.border}`, borderRadius: 8, padding: "10px 12px", marginBottom: 6, cursor: "pointer" }}>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span style={{ color: C.accent, fontFamily: C.mono, fontSize: 11 }}>{s.docNo}</span>
              <span style={{ color: C.muted, fontSize: 11 }}>{s.counter}</span>
            </div>
            <div style={{ color: C.muted, fontSize: 11 }}>{s.itemCount} รายการ · {fmt(s.totalQty)} ชิ้น</div>
          </button>
        ))}
      </Card>

      {compareData.length > 0 && (
        <>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
            {[
              { label: "พบในระบบ", v: compareData.filter(d => d.found).length, color: "green" },
              { label: "ส่วนต่าง≠0", v: compareData.filter(d => d.adjStock !== null && d.adjStock !== 0).length, color: "amber" },
              { label: "ไม่พบ", v: compareData.filter(d => !d.found).length, color: "red" },
            ].map(x => (
              <Card key={x.label} style={{ textAlign: "center", padding: 10 }}>
                <div style={{ color: C.muted, fontSize: 9 }}>{x.label}</div>
                <div style={{ color: { green: C.green, amber: C.amber, red: C.red }[x.color], fontSize: 22, fontWeight: 900 }}>{x.v}</div>
              </Card>
            ))}
          </div>

          <Card style={{ padding: 0, overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
              <thead>
                <tr style={{ background: C.surfaceHi }}>
                  {["สินค้า", "นับได้", "ขาย", "Adj.Count", "ยอดSB", "Adj.Stock"].map(h => (
                    <th key={h} style={{ padding: "8px 10px", textAlign: "left", color: C.muted, fontWeight: 700, whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {compareData.map((d, i) => {
                  const adjColor = d.adjStock === null ? C.muted : d.adjStock === 0 ? C.green : d.adjStock > 0 ? C.accent : C.red;
                  return (
                    <tr key={i} style={{ borderTop: `1px solid ${C.border}30` }}>
                      <td style={{ padding: "8px 10px", color: C.text }}>{d.productName.slice(0, 16)}</td>
                      <td style={{ padding: "8px 10px", color: C.text, fontWeight: 700 }}>{d.qty}</td>
                      <td style={{ padding: "8px 10px", color: C.amber }}>{d.sale}</td>
                      <td style={{ padding: "8px 10px", color: C.text, fontWeight: 700 }}>{d.adjCount}</td>
                      <td style={{ padding: "8px 10px", color: C.muted }}>{d.stockNow ?? "N/A"}</td>
                      <td style={{ padding: "8px 10px", color: adjColor, fontWeight: 900 }}>{d.adjStock !== null ? (d.adjStock > 0 ? "+" : "") + d.adjStock : "N/A"}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </Card>
        </>
      )}
    </div>
  );
}


// ── COUNTER: My Submissions ───────────────────────────────────────────────────
function MySubsView({ submissions }) {
  const [expanded, setExpanded] = useState(null);
  const statusCfg = {
    pending:  { label: "รอรีวิว",      color: C.amber, bg: C.amberLo },
    approved: { label: "อนุมัติแล้ว",  color: C.green, bg: C.greenLo },
    rejected: { label: "ส่งกลับแก้ไข", color: C.red,   bg: C.redLo  },
  };

  if (submissions.length === 0) return (
    <Card style={{ textAlign: "center", padding: 40 }}>
      <div style={{ fontSize: 40, marginBottom: 12 }}>📭</div>
      <div style={{ color: C.muted, fontSize: 13 }}>ยังไม่มีรายการที่ส่ง</div>
    </Card>
  );

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <div style={{ color: C.text, fontWeight: 700, fontSize: 15 }}>รายการที่ส่งแล้ว</div>
      {submissions.map(s => {
        const cfg = statusCfg[s.status] || statusCfg.pending;
        const isOpen = expanded === s.id;
        return (
          <Card key={s.id}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
              <div>
                <span style={{ color: C.accent, fontFamily: C.mono, fontSize: 11 }}>{s.docNo}</span>
                <div style={{ color: C.text, fontSize: 13, fontWeight: 700, marginTop: 2 }}>{timeStr(s.submittedAt)}</div>
              </div>
              <span style={{ background: cfg.bg, color: cfg.color, fontSize: 11, fontWeight: 700, padding: "3px 10px", borderRadius: 99 }}>{cfg.label}</span>
            </div>
            <div style={{ display: "flex", gap: 16, marginBottom: 8 }}>
              <span style={{ color: C.muted, fontSize: 12 }}><b style={{ color: C.text }}>{s.itemCount}</b> บาร์โค้ด</span>
              <span style={{ color: C.muted, fontSize: 12 }}><b style={{ color: C.green }}>{fmt(s.totalQty)}</b> ชิ้น</span>
            </div>
            {s.note && <div style={{ background: C.surfaceHi, borderRadius: 6, padding: "6px 10px", fontSize: 11, color: C.muted, marginBottom: 8, fontStyle: "italic" }}>"{s.note}"</div>}
            {s.status !== "pending" && s.reviewNote && (
              <div style={{ background: cfg.bg, border: `1px solid ${cfg.color}30`, borderRadius: 8, padding: "8px 10px", fontSize: 11, color: cfg.color, marginBottom: 8 }}>
                <b>{s.reviewedBy}:</b> {s.reviewNote}
              </div>
            )}
            <button onClick={() => setExpanded(isOpen ? null : s.id)}
              style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 11, padding: 0 }}>
              {isOpen ? "▲ ซ่อนรายการ" : "▼ ดูรายการ"}
            </button>
            {isOpen && (
              <div style={{ marginTop: 8, background: C.surfaceHi, borderRadius: 8, padding: 8, fontSize: 11 }}>
                {s.data.map((d, i) => (
                  <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", color: C.muted }}>
                    <span>{d.productName}</span>
                    <span style={{ color: C.green, fontWeight: 700 }}>{d.qty} {d.unit}</span>
                  </div>
                ))}
              </div>
            )}
          </Card>
        );
      })}
    </div>
  );
}

// ── MANAGER: Invoice Demo ─────────────────────────────────────────────────────
const DEMO_INVOICE = {
  invoice_no: "INV-2025-0582",
  invoice_date: "2025-05-17",
  vendor_name: "บริษัท เบเวอเรจ ซัพพลาย จำกัด",
  vendor_tax_id: "0105555012345",
  document_type: "invoice",
  price_type: "excl",
  products: [
    { no: 1, description: "น้ำดื่มสิงห์ 600ml",    qty: 240, price_ea: 6.5,  amount: 1560, vat: "v", barcode: "8851234567001" },
    { no: 2, description: "โค้ก 325ml",             qty: 120, price_ea: 11,   amount: 1320, vat: "v", barcode: "8851234567003" },
    { no: 3, description: "เป๊ปซี่ 325ml",           qty: 96,  price_ea: 11,   amount: 1056, vat: "v", barcode: "8851234567004" },
    { no: 4, description: "มาม่าหมูสับ",             qty: 480, price_ea: 3.8,  amount: 1824, vat: "v", barcode: "8851234567005" },
    { no: 5, description: "มาม่าต้มยำกุ้ง",          qty: 480, price_ea: 3.8,  amount: 1824, vat: "v", barcode: "8851234567006" },
  ],
};

function InvoiceDemoView() {
  const [step, setStep] = useState(1);
  const [scanning, setScanning] = useState(false);
  const [data, setData] = useState(null);
  const [saved, setSaved] = useState(false);

  const simulateScan = () => {
    setScanning(true);
    setTimeout(() => { setData(DEMO_INVOICE); setScanning(false); setStep(2); }, 2200);
  };

  const totalAmt = data ? data.products.reduce((s, p) => s + p.amount, 0) : 0;
  const vatAmt = +(totalAmt * 0.07).toFixed(2);
  const net = +(totalAmt + vatAmt).toFixed(2);

  const STEPS = ["อัปโหลด", "ตรวจสอบ", "บันทึก"];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      {/* Step bar */}
      <div style={{ display: "flex", alignItems: "center", gap: 0 }}>
        {STEPS.map((label, i) => {
          const n = i + 1, done = step > n, active = step === n;
          const col = done ? C.green : active ? C.accent : C.muted;
          return (
            <div key={n} style={{ display: "flex", alignItems: "center", flex: i < STEPS.length - 1 ? 1 : undefined }}>
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 3 }}>
                <div style={{ width: 28, height: 28, borderRadius: "50%", border: `2px solid ${col}`, background: done ? C.green : active ? C.accentLo : "transparent", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 900, color: done ? "#fff" : col }}>
                  {done ? "✓" : n}
                </div>
                <span style={{ fontSize: 9, color: col, whiteSpace: "nowrap" }}>{label}</span>
              </div>
              {i < STEPS.length - 1 && <div style={{ flex: 1, height: 2, background: step > n + 1 ? C.green : step > n ? C.accent : C.border, margin: "0 6px", marginBottom: 14 }} />}
            </div>
          );
        })}
      </div>

      {/* Step 1 */}
      {step === 1 && (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <Card style={{ textAlign: "center", padding: 32, border: `2px dashed ${C.border}` }}>
            <div style={{ fontSize: 40, marginBottom: 8 }}>🧾</div>
            <div style={{ color: C.text, fontWeight: 700, fontSize: 14, marginBottom: 4 }}>อัปโหลดใบกำกับภาษี</div>
            <div style={{ color: C.muted, fontSize: 12 }}>รองรับ JPG, PNG, PDF</div>
          </Card>
          <div style={{ background: C.amberLo, border: `1px solid ${C.amber}30`, borderRadius: 10, padding: 12, fontSize: 12, color: C.amber }}>
            💡 Demo mode — คลิก "จำลองการสแกน" เพื่อดูผลลัพธ์ AI
          </div>
          <Btn full color="accent" size="lg" onClick={simulateScan} disabled={scanning}>
            {scanning ? "🤖 Claude กำลังอ่านใบกำกับ..." : "🤖 จำลองการสแกน AI"}
          </Btn>
          {scanning && (
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {["กำลังอ่านรูปภาพ...", "วิเคราะห์ header...", "สกัดรายการสินค้า..."].map((t, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: C.muted }}>
                  <div style={{ width: 6, height: 6, borderRadius: "50%", background: C.accent, animation: "pulse 1s infinite", animationDelay: `${i * 0.3}s` }} />
                  {t}
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Step 2 */}
      {step === 2 && data && (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <Card>
            <div style={{ color: C.muted, fontSize: 10, marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ color: C.green }}>✓</span> AI อ่านสำเร็จ — ตรวจสอบและแก้ไขได้
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              {[["เลขที่", data.invoice_no], ["วันที่", data.invoice_date], ["Vendor", data.vendor_name], ["เลขภาษี", data.vendor_tax_id]].map(([k, v]) => (
                <div key={k}>
                  <div style={{ color: C.muted, fontSize: 9, marginBottom: 3 }}>{k}</div>
                  <div style={{ background: C.surfaceHi, border: `1px solid ${C.border}`, borderRadius: 6, padding: "6px 10px", fontSize: 11, color: C.text, fontWeight: 600 }}>{v}</div>
                </div>
              ))}
            </div>
          </Card>

          <Card style={{ padding: 0, overflow: "hidden" }}>
            <div style={{ padding: "10px 14px", borderBottom: `1px solid ${C.border}`, color: C.text, fontWeight: 700, fontSize: 12 }}>
              รายการสินค้า ({data.products.length} รายการ)
            </div>
            {data.products.map((p, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 14px", borderBottom: `1px solid ${C.border}20`, fontSize: 11 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ color: C.text, fontWeight: 600 }}>{p.description}</div>
                  <div style={{ color: C.muted }}>{p.qty} × ฿{p.price_ea}</div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ color: C.green, fontWeight: 700 }}>฿{fmt(p.amount)}</div>
                  {p.barcode && <div style={{ color: C.accent, fontSize: 9, fontFamily: C.mono }}>{p.barcode}</div>}
                </div>
              </div>
            ))}
            <div style={{ padding: "10px 14px", background: C.surfaceHi, display: "flex", flexDirection: "column", gap: 4, fontSize: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", color: C.muted }}><span>ยอดก่อน VAT</span><span>฿{fmt(totalAmt)}</span></div>
              <div style={{ display: "flex", justifyContent: "space-between", color: C.muted }}><span>VAT 7%</span><span>฿{fmt(vatAmt)}</span></div>
              <div style={{ display: "flex", justifyContent: "space-between", color: C.text, fontWeight: 700, fontSize: 14 }}><span>ยอดรวม</span><span style={{ color: C.green }}>฿{fmt(net)}</span></div>
            </div>
          </Card>

          <Btn full color="green" onClick={() => setStep(3)}>บันทึก & ส่งออก →</Btn>
        </div>
      )}

      {/* Step 3 */}
      {step === 3 && (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <Card style={{ textAlign: "center", padding: 24 }}>
            <div style={{ fontSize: 36, marginBottom: 8 }}>📊</div>
            <div style={{ color: C.text, fontWeight: 700, fontSize: 14, marginBottom: 12 }}>พร้อม Export</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 16 }}>
              {[["ใบกำกับ", "1 ใบ"], ["สินค้า", `${DEMO_INVOICE.products.length} รายการ`], ["ยอดรวม", `฿${fmt(net)}`], ["VAT", `฿${fmt(vatAmt)}`]].map(([k, v]) => (
                <div key={k} style={{ background: C.surfaceHi, borderRadius: 8, padding: 10 }}>
                  <div style={{ color: C.muted, fontSize: 9 }}>{k}</div>
                  <div style={{ color: C.text, fontWeight: 700, fontSize: 14 }}>{v}</div>
                </div>
              ))}
            </div>
          </Card>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            {[
              { label: "📥 Excel", color: "#16a34a", sub: "bill_header + invoice" },
              { label: "📋 CSV",   color: C.accent,  sub: "imp_data format" },
              { label: "☁️ Drive",  color: "#4285f4", sub: "Google Drive" },
              { label: "💾 Supabase", color: "#3ecf8e", sub: "bill_header + imp_data" },
            ].map(b => (
              <button key={b.label} onClick={() => setSaved(true)}
                style={{ background: C.surfaceHi, border: `1px solid ${C.border}`, borderRadius: 10, padding: "14px 10px", cursor: "pointer", textAlign: "center" }}>
                <div style={{ fontSize: 20, marginBottom: 4 }}>{b.label.split(" ")[0]}</div>
                <div style={{ color: b.color, fontWeight: 700, fontSize: 12 }}>{b.label.split(" ").slice(1).join(" ")}</div>
                <div style={{ color: C.muted, fontSize: 10, marginTop: 2 }}>{b.sub}</div>
              </button>
            ))}
          </div>
          {saved && (
            <div style={{ background: C.greenLo, border: `1px solid ${C.green}30`, borderRadius: 10, padding: 12, display: "flex", alignItems: "center", gap: 10, fontSize: 12 }}>
              <span style={{ fontSize: 20 }}>✅</span>
              <div style={{ color: C.green }}>บันทึกสำเร็จ! (Demo mode — ไม่ได้ส่งจริง)</div>
            </div>
          )}
          <Btn full color="ghost" onClick={() => { setStep(1); setData(null); setSaved(false); }}>↺ เริ่มใหม่</Btn>
        </div>
      )}
    </div>
  );
}

// ── MAIN APP ──────────────────────────────────────────────────────────────────
// Nav config per feature
const NAV_MAP = {
  // counter features
  stock_compare_counter: [
    { id: "count",    label: "นับสต็อก",   icon: "📦" },
    { id: "review",   label: "ตรวจสอบ",    icon: "📋" },
    { id: "my_subs",  label: "ที่ส่งแล้ว", icon: "📤" },
  ],
  stock_only_counter: [
    { id: "count",    label: "นับสต็อก",   icon: "📦" },
    { id: "review",   label: "ตรวจสอบ",    icon: "📋" },
    { id: "my_subs",  label: "ที่ส่งแล้ว", icon: "📤" },
  ],
  // manager features
  stock_compare_manager: [
    { id: "dashboard", label: "แดชบอร์ด",    icon: "🏠" },
    { id: "inbox",     label: "รีวิว",        icon: "📬" },
    { id: "compare",   label: "เปรียบเทียบ",  icon: "⚖️" },
  ],
  stock_only_manager: [
    { id: "dashboard", label: "แดชบอร์ด",    icon: "🏠" },
    { id: "inbox",     label: "รีวิว",        icon: "📬" },
  ],
  invoice_counter: [
    { id: "invoice",   label: "สแกนบิล",     icon: "🧾" },
  ],
  invoice_manager: [
    { id: "invoice",   label: "สแกนบิล",     icon: "🧾" },
  ],
};

const FEATURE_ACCENT = {
  stock_compare: "#4f6ef7",
  stock_only:    "#22c55e",
  invoice:       "#a855f7",
};

export default function App() {
  const [user, setUser] = useState(null);
  const [view, setView] = useState("count");
  const [entries, setEntries] = useState([]);
  const [submissions, setSubmissions] = useState(DEMO_SUBMISSIONS);

  const onLogin = u => {
    setUser(u);
    if (u.feature === "invoice") setView("invoice");
    else if (u.role === "manager") setView("dashboard");
    else setView("count");
  };
  const onLogout = () => { setUser(null); setView("count"); setEntries([]); };

  const addEntry = e => setEntries(p => [e, ...p]);
  const deleteEntry = id => setEntries(p => p.filter(e => e.id !== id));
  const onSubmit = sub => setSubmissions(p => [sub, ...p]);
  const onReview = (id, status, note) => setSubmissions(p =>
    p.map(s => s.id === id ? { ...s, status, reviewNote: note, reviewedAt: new Date().toISOString(), reviewedBy: user.name } : s)
  );

  if (!user) return <Login onLogin={onLogin} />;

  const isManager = user.role === "manager";
  const feat = user.feature || "stock_compare";
  const pendingCount = submissions.filter(s => s.status === "pending").length;

  // nav tabs per feature
  const NAV_BY_FEATURE = {
    stock_compare: isManager
      ? [{ id:"dashboard",label:"แดชบอร์ด",icon:"🏠"},{id:"inbox",label:"รีวิว",icon:"📬"},{id:"compare",label:"เปรียบเทียบ",icon:"⚖️"}]
      : [{id:"count",label:"นับสต็อก",icon:"📦"},{id:"review",label:"ตรวจสอบ",icon:"📋"},{id:"my_subs",label:"ที่ส่งแล้ว",icon:"📤"}],
    stock_only: isManager
      ? [{id:"dashboard",label:"แดชบอร์ด",icon:"🏠"},{id:"inbox",label:"รีวิว",icon:"📬"}]
      : [{id:"count",label:"นับสต็อก",icon:"📦"},{id:"review",label:"ตรวจสอบ",icon:"📋"},{id:"my_subs",label:"ที่ส่งแล้ว",icon:"📤"}],
    invoice: [{id:"invoice",label:"สแกนบิล",icon:"🧾"}],
  };

  const navBase = NAV_BY_FEATURE[feat] || NAV_BY_FEATURE.stock_compare;
  const nav = navBase.map(item => ({
    ...item,
    badge: item.id === "inbox" ? pendingCount : item.id === "review" ? entries.length : 0,
  }));

  const ACCENT_BY_FEATURE = { stock_compare: "#4f6ef7", stock_only: "#22c55e", invoice: "#a855f7" };
  const accentCol = ACCENT_BY_FEATURE[feat] || C.accent;
  const featureLabel = FEATURES.find(f => f.id === feat)?.label || "";

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.text, fontFamily: "'Segoe UI', system-ui, sans-serif" }}>
      {/* Header */}
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: "12px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", position: "sticky", top: 0, zIndex: 50 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 34, height: 34, borderRadius: 9, background: `linear-gradient(135deg, ${accentCol}, ${accentCol}99)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>📦</div>
          <div>
            <div style={{ fontWeight: 900, fontSize: 15, letterSpacing: "-0.02em" }}>KUUHOO</div>
            <div style={{ color: C.muted, fontSize: 9 }}>{isManager ? "👔" : "👷"} {user.name} · <span style={{ color: accentCol }}>{featureLabel}</span></div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <Badge color="accent">Demo</Badge>
          <button onClick={onLogout} style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 12 }}>ออก</button>
        </div>
      </div>

      {/* Content */}
      <div style={{ padding: "16px", paddingBottom: nav.length > 0 ? 80 : 16, maxWidth: 560, margin: "0 auto" }}>
        {!isManager && view === "count"     && <CountView    user={user} entries={entries} onAdd={addEntry} onDelete={deleteEntry} onNext={() => setView("review")} />}
        {!isManager && view === "review"    && <ReviewView   user={user} entries={entries} onSubmit={onSubmit} onBack={() => setView("count")} />}
        {!isManager && view === "my_subs"   && <MySubsView   submissions={submissions.filter(s => s.counter === user.name)} />}
        {isManager  && view === "dashboard" && <DashboardView submissions={submissions} setView={setView} />}
        {isManager  && view === "inbox"     && <InboxView    submissions={submissions} onReview={onReview} />}
        {isManager  && view === "compare"   && <CompareView  submissions={submissions} />}
        {isManager  && view === "invoice"   && <InvoiceDemoView />}
        {/* invoice for counter if ever needed */}
        {!isManager && view === "invoice"   && <InvoiceDemoView />}
      </div>

      {/* Bottom Nav */}
      {nav.length > 0 && (
        <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: C.surface, borderTop: `1px solid ${C.border}`, display: "flex" }}>
          {nav.map(item => {
            const active = view === item.id;
            return (
              <button key={item.id} onClick={() => setView(item.id)}
                style={{ flex: 1, padding: "10px 4px 12px", background: "none", border: "none", borderTop: `2px solid ${active ? accentCol : "transparent"}`, color: active ? accentCol : C.muted, cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 3, position: "relative" }}>
                <span style={{ fontSize: 18 }}>{item.icon}</span>
                <span style={{ fontSize: 10, fontWeight: 700 }}>{item.label}</span>
                {item.badge > 0 && (
                  <span style={{ position: "absolute", top: 6, left: "55%", background: C.red, color: "#fff", fontSize: 9, fontWeight: 900, padding: "1px 5px", borderRadius: 99 }}>{item.badge}</span>
                )}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
