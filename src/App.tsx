import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

const INITIAL_STOCK = [
  { barcode: "8851234567890", product: "สินค้า A", stock: 120, location: "A01-01" },
  { barcode: "8851234567891", product: "สินค้า B", stock: 80, location: "A01-02" },
  { barcode: "8851234567892", product: "สินค้า C", stock: 60, location: "A01-03" },
  { barcode: "8851234567893", product: "สินค้า D", stock: 40, location: "A02-01" },
  { barcode: "8851234567894", product: "สินค้า E", stock: 200, location: "B01-01" },
];

const INITIAL_ORDERS = [];

function generateId() {
  return "PO" + Date.now().toString().slice(-6) + Math.floor(Math.random() * 100);
}

function now() {
  const d = new Date();
  const pad = n => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()+543} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}


function Badge({ children, color = "gray" }) {
  const colors = {
    pending: { bg: "#fff7e6", text: "#d46b08", border: "#ffd591" },
    complete: { bg: "#f6ffed", text: "#389e0d", border: "#b7eb8f" },
    over3h: { bg: "#fff1f0", text: "#cf1322", border: "#ffa39e" },
    gray: { bg: "#f5f5f5", text: "#595959", border: "#d9d9d9" },
  };
  const c = colors[color] || colors.gray;
  return (
    <span style={{ background: c.bg, color: c.text, border: `1px solid ${c.border}`, borderRadius: 6, padding: "2px 10px", fontSize: 12, fontWeight: 600 }}>
      {children}
    </span>
  );
}

function StatCard({ icon, value, label, color = "#1677ff" }) {
  return (
    <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #f0f0f0", padding: "18px 20px", flex: 1, minWidth: 140, boxShadow: "0 1px 4px rgba(0,0,0,0.06)" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        <div style={{ width: 44, height: 44, borderRadius: 10, background: color + "18", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22 }}>{icon}</div>
        <div>
          <div style={{ fontSize: 26, fontWeight: 700, color, lineHeight: 1.1 }}>{value}</div>
          <div style={{ fontSize: 13, color: "#8c8c8c", marginTop: 2 }}>{label}</div>
        </div>
      </div>
    </div>
  );
}

function Modal({ open, onClose, title, children }) {
  if (!open) return null;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 14, padding: "28px 32px", minWidth: 380, maxWidth: 520, width: "90%", boxShadow: "0 8px 40px rgba(0,0,0,0.18)" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <span style={{ fontWeight: 700, fontSize: 17 }}>{title}</span>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 20, color: "#8c8c8c" }}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}

const inputStyle = {
  border: "1px solid #d9d9d9", borderRadius: 8, padding: "8px 12px", fontSize: 14, width: "100%", outline: "none", boxSizing: "border-box",
  fontFamily: "inherit", transition: "border 0.2s",
};

const btnPrimary = {
  background: "#1677ff", color: "#fff", border: "none", borderRadius: 8, padding: "9px 18px",
  fontSize: 14, fontWeight: 600, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6,
};
const btnGreen = { ...btnPrimary, background: "#52c41a" };
const btnOrange = { ...btnPrimary, background: "#fa8c16" };
const btnRed = { ...btnPrimary, background: "#ff4d4f" };
const btnGray = { ...btnPrimary, background: "#fff", color: "#595959", border: "1px solid #d9d9d9" };

export default function App() {
  const [stock, setStock] = useState(() => {
    try { return JSON.parse(localStorage.getItem("wh_stock")) || INITIAL_STOCK; } catch { return INITIAL_STOCK; }
  });
  const [orders, setOrders] = useState(() => {
    try { return JSON.parse(localStorage.getItem("wh_orders")) || INITIAL_ORDERS; } catch { return INITIAL_ORDERS; }
  });

  const [pickNo, setPickNo] = useState("");
  const [barcode, setBarcode] = useState("");
  const [location, setLocation] = useState("");
  const [qty, setQty] = useState(1);
  const [scanStatus, setScanStatus] = useState("พร้อมสแกน");
  const [scanStatusColor, setScanStatusColor] = useState("#52c41a");
  const [autoConfirm, setAutoConfirm] = useState(true);
  const [searchOrder, setSearchOrder] = useState("");
  const [showCompleted, setShowCompleted] = useState(false);
  const [tab, setTab] = useState("main");
  const [stockSearch, setStockSearch] = useState("");

  const [addStockForm, setAddStockForm] = useState({ barcode: "", product: "", stock: "", location: "" });
  const [addOrderForm, setAddOrderForm] = useState({ pickNo: "", customer: "", barcode: "", product: "", location: "", required: 1 });
  const [editStockItem, setEditStockItem] = useState(null);
  const [editOrderItem, setEditOrderItem] = useState(null);
  const [confirmClear, setConfirmClear] = useState(false);
  const [alert, setAlert] = useState(null);

  const barcodeRef = useRef();
  const pickNoRef = useRef();

  useEffect(() => { localStorage.setItem("wh_stock", JSON.stringify(stock)); }, [stock]);
  useEffect(() => { localStorage.setItem("wh_orders", JSON.stringify(orders)); }, [orders]);

  const showAlert = useCallback((msg, type = "success") => {
    setAlert({ msg, type });
    setTimeout(() => setAlert(null), 2800);
  }, []);

  const pendingOrders = orders.filter(o => o.status !== "COMPLETE");
  const completedOrders = orders.filter(o => o.status === "COMPLETE");
  const displayOrders = showCompleted ? completedOrders : pendingOrders;

  const filteredOrders = displayOrders.filter(o => {
    const q = searchOrder.toLowerCase();
    return !q || o.pickNo?.toLowerCase().includes(q) || o.customer?.toLowerCase().includes(q) ||
      o.product?.toLowerCase().includes(q) || o.barcode?.toLowerCase().includes(q);
  });

  const filteredStock = stock.filter(s => {
    const q = stockSearch.toLowerCase();
    return !q || s.barcode.includes(q) || s.product.toLowerCase().includes(q) || s.location.toLowerCase().includes(q);
  });

  function getOrderStatus(o) {
    if (o.status === "COMPLETE") return "complete";
    const created = new Date(o.createdAtRaw);
    if ((Date.now() - created) > 3 * 60 * 60 * 1000) return "over3h";
    return "pending";
  }

  function confirmPick() {
    if (!barcode.trim()) { setScanStatus("❌ กรุณากรอก Barcode"); setScanStatusColor("#ff4d4f"); return; }
    const stockItem = stock.find(s => s.barcode === barcode.trim());
    if (!stockItem) { setScanStatus("❌ ไม่พบสินค้าในระบบ"); setScanStatusColor("#ff4d4f"); return; }

    let matchedOrder = null;
    if (pickNo.trim()) {
      matchedOrder = orders.find(o => o.pickNo === pickNo.trim() && o.barcode === barcode.trim() && o.status !== "COMPLETE");
    } else {
      matchedOrder = orders.find(o => o.barcode === barcode.trim() && o.status !== "COMPLETE");
    }

    const pickQty = Number(qty) || 1;
    if (stockItem.stock < pickQty) { setScanStatus("❌ สต็อกไม่เพียงพอ"); setScanStatusColor("#ff4d4f"); return; }

    setStock(prev => prev.map(s => s.barcode === barcode.trim() ? { ...s, stock: s.stock - pickQty } : s));

    if (matchedOrder) {
      setOrders(prev => prev.map(o => o.id === matchedOrder.id
        ? { ...o, picked: (o.picked || 0) + pickQty, status: (o.picked || 0) + pickQty >= o.required ? "COMPLETE" : o.status, completedAt: (o.picked || 0) + pickQty >= o.required ? now() : o.completedAt, completedAtRaw: (o.picked || 0) + pickQty >= o.required ? new Date().toISOString() : o.completedAtRaw }
        : o));
      setScanStatus(`✅ หยิบสำเร็จ: ${stockItem.product} x${pickQty}`);
    } else {
      setScanStatus(`✅ หยิบสำเร็จ: ${stockItem.product} x${pickQty} (ไม่มี Order)`);
    }
    setScanStatusColor("#52c41a");

    if (location.trim()) setLocation("");
    setBarcode("");
    setTimeout(() => { setScanStatus("พร้อมสแกน"); setScanStatusColor("#52c41a"); barcodeRef.current?.focus(); }, 1500);
  }

  function handleBarcodeKeyDown(e) {
    if (e.key === "Enter") {
      const found = stock.find(s => s.barcode === barcode.trim());
      if (found && !location) setLocation(found.location);
      if (autoConfirm) { setTimeout(confirmPick, 50); }
    }
  }

  function exportExcel() {
    const rows = [["Pick No", "Customer", "Barcode", "Product", "Location", "Required", "Picked", "Status", "Created At", "Completed At", "ใช้เวลา"]];
    orders.forEach(o => {
      let duration = "";
      if (o.completedAtRaw && o.createdAtRaw) {
        const secs = Math.round((new Date(o.completedAtRaw) - new Date(o.createdAtRaw)) / 1000);
        if (secs < 60) duration = secs + " วินาที";
        else if (secs < 3600) duration = Math.round(secs/60) + " นาที";
        else { const h = Math.floor(secs/3600), m = Math.round((secs%3600)/60); duration = h + "ชม. " + m + "น."; }
      }
      rows.push([o.pickNo, o.customer, o.barcode, o.product, o.location, o.required, o.picked || 0, o.status, o.createdAt, o.completedAt || "", duration]);
    });
    const csv = rows.map(r => r.map(c => `"${c}"`).join(",")).join("\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = "pick_orders.csv"; a.click();
  }

  function backupData() {
    const data = { stock, orders, backupAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = `warehouse_backup_${Date.now()}.json`; a.click();
    showAlert("Backup สำเร็จ");
  }

  function restoreData(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const data = JSON.parse(ev.target.result);
        if (data.stock) setStock(data.stock);
        if (data.orders) setOrders(data.orders);
        showAlert("Restore สำเร็จ");
      } catch { showAlert("ไฟล์ไม่ถูกต้อง", "error"); }
    };
    reader.readAsText(file);
    e.target.value = "";
  }

  function handleAddStock() {
    const { barcode: bc, product, stock: st, location: loc } = addStockForm;
    if (!bc || !product || !st || !loc) { showAlert("กรุณากรอกข้อมูลให้ครบ", "error"); return; }
    const existing = stock.find(s => s.barcode === bc);
    if (existing) {
      setStock(prev => prev.map(s => s.barcode === bc ? { ...s, stock: s.stock + Number(st), location: loc } : s));
      showAlert("อัพเดทสต็อกสำเร็จ");
    } else {
      setStock(prev => [...prev, { barcode: bc, product, stock: Number(st), location: loc }]);
      showAlert("เพิ่มสินค้าสำเร็จ");
    }
    setAddStockForm({ barcode: "", product: "", stock: "", location: "" });
  }

  function handleAddOrder() {
    const { customer, barcode: bc, product, location: loc, required } = addOrderForm;
    if (!bc || !product) { showAlert("กรุณากรอก Barcode และชื่อสินค้า", "error"); return; }
    const id = generateId();
    const newOrder = {
      id, pickNo: addOrderForm.pickNo || id, customer, barcode: bc, product,
      location: loc, required: Number(required) || 1, picked: 0,
      status: "PENDING", createdAt: now(), createdAtRaw: new Date().toISOString(),
    };
    setOrders(prev => [...prev, newOrder]);
    setAddOrderForm({ pickNo: "", customer: "", barcode: "", product: "", location: "", required: 1 });
    showAlert("เพิ่ม Order สำเร็จ");
  }

  function handleEditStock() {
    setStock(prev => prev.map(s => s.barcode === editStockItem.barcode ? editStockItem : s));
    setEditStockItem(null); showAlert("แก้ไขสำเร็จ");
  }

  function handleEditOrder() {
    setOrders(prev => prev.map(o => o.id === editOrderItem.id ? editOrderItem : o));
    setEditOrderItem(null); showAlert("แก้ไขสำเร็จ");
  }

  // ---- Excel Upload helpers ----
  const [xlsxPreview, setXlsxPreview] = useState(null); // { type: 'stock'|'order', rows: [], errors: [] }

  function downloadStockTemplate() {
    const ws = XLSX.utils.aoa_to_sheet([
      ["barcode", "product", "stock", "location"],
      ["8851234567890", "สินค้า A", 100, "A01-01"],
      ["8851234567891", "สินค้า B", 50, "A01-02"],
    ]);
    ws["!cols"] = [{ wch: 20 }, { wch: 20 }, { wch: 10 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Stock");
    XLSX.writeFile(wb, "stock_template.xlsx");
  }

  function downloadOrderTemplate() {
    const ws = XLSX.utils.aoa_to_sheet([
      ["pickNo", "customer", "barcode", "product", "location", "required"],
      ["PO001", "ลูกค้า A", "8851234567890", "สินค้า A", "A01-01", 5],
      ["PO002", "ลูกค้า B", "8851234567891", "สินค้า B", "A01-02", 3],
    ]);
    ws["!cols"] = [{ wch: 14 }, { wch: 16 }, { wch: 20 }, { wch: 20 }, { wch: 14 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Orders");
    XLSX.writeFile(wb, "order_template.xlsx");
  }

  function handleStockXlsx(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const errors = [];
        const rows = raw.map((r, i) => {
          const row = {
            barcode: String(r.barcode || r.Barcode || r["บาร์โค้ด"] || "").trim(),
            product: String(r.product || r.Product || r["ชื่อสินค้า"] || "").trim(),
            stock: Number(r.stock || r.Stock || r["จำนวน"] || 0),
            location: String(r.location || r.Location || r["ที่เก็บ"] || "").trim(),
          };
          if (!row.barcode) errors.push(`แถว ${i + 2}: ไม่มี Barcode`);
          if (!row.product) errors.push(`แถว ${i + 2}: ไม่มีชื่อสินค้า`);
          return row;
        }).filter(r => r.barcode);
        setXlsxPreview({ type: "stock", rows, errors });
      } catch { showAlert("ไม่สามารถอ่านไฟล์ได้", "error"); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }

  function handleOrderXlsx(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const errors = [];
        const rows = raw.map((r, i) => {
          const row = {
            pickNo: String(r.pickNo || r["Pick No"] || r.pick_no || "").trim(),
            customer: String(r.customer || r.Customer || r["ลูกค้า"] || "").trim(),
            barcode: String(r.barcode || r.Barcode || r["บาร์โค้ด"] || "").trim(),
            product: String(r.product || r.Product || r["สินค้า"] || r["ชื่อสินค้า"] || "").trim(),
            location: String(r.location || r.Location || r["ที่เก็บ"] || "").trim(),
            required: Number(r.required || r.Required || r["จำนวน"] || 1),
          };
          if (!row.barcode) errors.push(`แถว ${i + 2}: ไม่มี Barcode`);
          if (!row.product) errors.push(`แถว ${i + 2}: ไม่มีชื่อสินค้า`);
          return row;
        }).filter(r => r.barcode);
        setXlsxPreview({ type: "order", rows, errors });
      } catch { showAlert("ไม่สามารถอ่านไฟล์ได้", "error"); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }

  function confirmXlsxImport() {
    if (!xlsxPreview) return;
    if (xlsxPreview.type === "stock") {
      setStock(prev => {
        const updated = [...prev];
        xlsxPreview.rows.forEach(r => {
          const idx = updated.findIndex(s => s.barcode === r.barcode);
          if (idx >= 0) updated[idx] = { ...updated[idx], stock: updated[idx].stock + r.stock, location: r.location || updated[idx].location, product: r.product || updated[idx].product };
          else updated.push(r);
        });
        return updated;
      });
      showAlert(`นำเข้าสต็อกสำเร็จ ${xlsxPreview.rows.length} รายการ`);
    } else {
      const newOrders = xlsxPreview.rows.map(r => {
        const id = generateId();
        return { id, pickNo: r.pickNo || id, customer: r.customer, barcode: r.barcode, product: r.product, location: r.location, required: r.required, picked: 0, status: "PENDING", createdAt: now(), createdAtRaw: new Date().toISOString() };
      });
      setOrders(prev => [...prev, ...newOrders]);
      showAlert(`นำเข้า Order สำเร็จ ${newOrders.length} รายการ`);
    }
    setXlsxPreview(null);
  }

  const sectionCard = { background: "#fff", borderRadius: 14, border: "1px solid #f0f0f0", padding: "22px 24px", marginBottom: 20, boxShadow: "0 1px 4px rgba(0,0,0,0.05)" };
  const labelStyle = { fontSize: 13, color: "#595959", marginBottom: 4, display: "block" };

  return (
    <div style={{ background: "#f7f8fa", minHeight: "100vh", fontFamily: "'Noto Sans Thai', sans-serif", color: "#262626" }}>
      {alert && (
        <div style={{ position: "fixed", top: 20, right: 20, zIndex: 2000, background: alert.type === "error" ? "#ff4d4f" : "#52c41a", color: "#fff", borderRadius: 10, padding: "12px 22px", fontWeight: 600, fontSize: 14, boxShadow: "0 4px 16px rgba(0,0,0,0.18)", animation: "fadein 0.3s" }}>
          {alert.type === "error" ? "❌ " : "✅ "}{alert.msg}
        </div>
      )}

      {/* Header */}
      <div style={{ background: "#fff", borderBottom: "1px solid #f0f0f0", padding: "0 32px", display: "flex", alignItems: "center", gap: 16, position: "sticky", top: 0, zIndex: 100, height: 58 }}>
        <span style={{ fontSize: 22, marginRight: 4 }}>📦</span>
        <span style={{ fontWeight: 800, fontSize: 18, letterSpacing: -0.5 }}>Warehouse Picking System</span>
        <div style={{ flex: 1 }} />
        <input type="file" accept=".json" id="restore-file" style={{ display: "none" }} onChange={restoreData} />
        <button style={btnGray} onClick={backupData}>💾 Backup</button>
        <button style={btnGray} onClick={() => document.getElementById("restore-file").click()}>🔄 Restore</button>
        <button style={btnRed} onClick={() => setConfirmClear(true)}>🗑️ ล้างข้อมูลวันนี้</button>
      </div>

      <div style={{ maxWidth: 1200, margin: "0 auto", padding: "24px 20px" }}>
        {/* Stat Cards */}
        <div style={{ display: "flex", gap: 16, marginBottom: 20, flexWrap: "wrap" }}>
          <StatCard icon="📋" value={orders.length} label="รายการทั้งหมด" color="#1677ff" />
          <StatCard icon="✅" value={completedOrders.length} label="เสร็จสิ้น" color="#52c41a" />
          <StatCard icon="⏳" value={pendingOrders.length} label="รอดำเนินการ" color="#fa8c16" />
          <StatCard icon="🎁" value={stock.reduce((a, s) => a + s.stock, 0)} label="สินค้าคงเหลือรวม" color="#722ed1" />
        </div>

        {/* Scanner Section */}
        <div style={sectionCard}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
            <span style={{ fontWeight: 700, fontSize: 16 }}>📱 สแกนหยิบสินค้า (Mobile Scanner)</span>
            <button onClick={() => setAutoConfirm(v => !v)} style={{ ...btnPrimary, background: autoConfirm ? "#52c41a" : "#8c8c8c", fontSize: 12, padding: "4px 12px" }}>
              ⚡ Auto Confirm {autoConfirm ? "เปิดอยู่" : "ปิดอยู่"}
            </button>
          </div>
          <div style={{ background: "#fffbe6", border: "1px solid #ffe58f", borderRadius: 8, padding: "8px 14px", marginBottom: 16, fontSize: 13, color: "#614700" }}>
            💡 วิธีใช้: กรอก Pick No → สแกน/กรอก Barcode แล้วกด Enter — ระบบจะบันทึกการหยิบอัตโนมัติทันที
          </div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
            <div style={{ flex: "1 1 180px" }}>
              <label style={labelStyle}>Pick No</label>
              <input ref={pickNoRef} style={inputStyle} value={pickNo} onChange={e => setPickNo(e.target.value)} placeholder="สแกนหรือกรอก Pick No" onKeyDown={e => e.key === "Enter" && barcodeRef.current?.focus()} />
            </div>
            <div style={{ flex: "1 1 220px" }}>
              <label style={labelStyle}>Barcode <span style={{ color: "#8c8c8c", fontWeight: 400 }}>(กด Enter เพื่อบันทึก)</span></label>
              <input ref={barcodeRef} style={{ ...inputStyle, border: "2px solid #40a9ff", background: "#f0f9ff" }} value={barcode} onChange={e => setBarcode(e.target.value)} placeholder="สแกนหรือกรอก Barcode ที่นี่" onKeyDown={handleBarcodeKeyDown} />
            </div>
            <div style={{ flex: "1 1 160px" }}>
              <label style={labelStyle}>Location <span style={{ color: "#8c8c8c", fontWeight: 400 }}>(ระบุเพื่อแยก barcode ซ้ำ)</span></label>
              <input style={inputStyle} value={location} onChange={e => setLocation(e.target.value)} placeholder="เช่น Q14-028-40" />
            </div>
            <div style={{ flex: "0 0 80px" }}>
              <label style={labelStyle}>Qty</label>
              <input style={{ ...inputStyle, textAlign: "center" }} type="number" min="1" value={qty} onChange={e => setQty(e.target.value)} />
            </div>
            <button style={{ ...btnGreen, flexShrink: 0 }} onClick={confirmPick}>✅ ยืนยันการหยิบสินค้า</button>
          </div>
          <div style={{ marginTop: 10, fontSize: 14 }}>
            สถานะ: <span style={{ color: scanStatusColor, fontWeight: 600 }}>{scanStatus}</span>
          </div>
        </div>

        {/* Pick Order Table */}
        <div style={sectionCard}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16, flexWrap: "wrap" }}>
            <span style={{ fontWeight: 700, fontSize: 16 }}>📋 รายการ Pick Order</span>
            <Badge color="pending">⏳ PENDING {pendingOrders.length} รายการ</Badge>
            <div style={{ flex: 1, minWidth: 220 }}>
              <input style={{ ...inputStyle, maxWidth: 320 }} value={searchOrder} onChange={e => setSearchOrder(e.target.value)} placeholder="🔍 ค้นหา Pick No, ลูกค้า, สินค้า, Barcode" />
            </div>
            <button style={{ ...btnGreen, position: "relative" }} onClick={() => setShowCompleted(v => !v)}>
              ✅ {showCompleted ? "ดูรายการค้าง" : "ดูรายการเสร็จแล้ว"}
              <span style={{ background: "#fff", color: "#52c41a", borderRadius: 99, padding: "1px 7px", fontSize: 12, marginLeft: 6 }}>{showCompleted ? pendingOrders.length : completedOrders.length}</span>
            </button>
            <button style={btnOrange} onClick={exportExcel}>⬇️ Export Excel</button>
          </div>

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: "#fafafa", borderBottom: "1px solid #f0f0f0" }}>
                  {["Pick No", "Customer", "Barcode", "Product", "Location", "Required", "Picked", "Status", "Alert", "Created At", "Completed At", "ใช้เวลา", ""].map(h => (
                    <th key={h} style={{ padding: "10px 12px", textAlign: "left", fontWeight: 600, color: "#595959", whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredOrders.length === 0 ? (
                  <tr><td colSpan={13} style={{ textAlign: "center", padding: 40, color: "#8c8c8c" }}>
                    <div style={{ fontSize: 32, marginBottom: 8 }}>🎉</div>
                    ไม่มีรายการค้างดำเนินการ<br />
                    <span style={{ fontSize: 12 }}>รายการทั้งหมดเสร็จสิ้นแล้ว กดปุ่ม "ดูรายการเสร็จแล้ว" เพื่อดู</span>
                  </td></tr>
                ) : filteredOrders.map(o => {
                  const st = getOrderStatus(o);
                  return (
                    <tr key={o.id} style={{ borderBottom: "1px solid #f5f5f5" }}>
                      <td style={{ padding: "10px 12px", fontWeight: 600, color: "#1677ff" }}>{o.pickNo}</td>
                      <td style={{ padding: "10px 12px" }}>{o.customer || "-"}</td>
                      <td style={{ padding: "10px 12px", fontFamily: "monospace" }}>{o.barcode}</td>
                      <td style={{ padding: "10px 12px" }}>{o.product}</td>
                      <td style={{ padding: "10px 12px" }}>{o.location || "-"}</td>
                      <td style={{ padding: "10px 12px", textAlign: "center" }}>{o.required}</td>
                      <td style={{ padding: "10px 12px", textAlign: "center", color: (o.picked || 0) >= o.required ? "#52c41a" : "#fa8c16", fontWeight: 600 }}>{o.picked || 0}</td>
                      <td style={{ padding: "10px 12px" }}>
                        <Badge color={st === "complete" ? "complete" : st === "over3h" ? "over3h" : "pending"}>
                          {o.status}
                        </Badge>
                      </td>
                      <td style={{ padding: "10px 12px" }}>{st === "over3h" ? <span style={{ color: "#ff4d4f", fontWeight: 600 }}>⚠️ เกิน 3ชม.</span> : "-"}</td>
                      {/* Created At */}
                      <td style={{ padding: "10px 12px", whiteSpace: "nowrap" }}>
                        {o.createdAt ? (
                          <div>
                            <div style={{ fontSize: 12, color: "#262626", fontWeight: 500 }}>{o.createdAt.split(" ")[0]}</div>
                            <div style={{ fontSize: 12, color: "#1677ff", fontWeight: 600 }}>{o.createdAt.split(" ")[1]}</div>
                          </div>
                        ) : "-"}
                      </td>
                      {/* Completed At */}
                      <td style={{ padding: "10px 12px", whiteSpace: "nowrap" }}>
                        {o.completedAt ? (
                          <div>
                            <div style={{ fontSize: 12, color: "#262626", fontWeight: 500 }}>{o.completedAt.split(" ")[0]}</div>
                            <div style={{ fontSize: 12, color: "#52c41a", fontWeight: 600 }}>{o.completedAt.split(" ")[1]}</div>
                          </div>
                        ) : <span style={{ color: "#d9d9d9", fontSize: 12 }}>-</span>}
                      </td>
                      {/* Duration */}
                      <td style={{ padding: "10px 12px", textAlign: "center" }}>
                        {o.completedAtRaw && o.createdAtRaw ? (() => {
                          const mins = Math.round((new Date(o.completedAtRaw) - new Date(o.createdAtRaw)) / 60000);
                          const secs = Math.round((new Date(o.completedAtRaw) - new Date(o.createdAtRaw)) / 1000);
                          if (secs < 60) return <span style={{ color: "#722ed1", fontWeight: 700, fontSize: 13 }}>{secs} วิ</span>;
                          if (mins < 60) return <span style={{ color: "#52c41a", fontWeight: 700, fontSize: 13 }}>{mins} นาที</span>;
                          const h = Math.floor(mins / 60), m = mins % 60;
                          return <span style={{ color: "#fa8c16", fontWeight: 700, fontSize: 13 }}>{h}ชม. {m}น.</span>;
                        })() : <span style={{ color: "#d9d9d9", fontSize: 12 }}>-</span>}
                      </td>
                      <td style={{ padding: "10px 12px" }}>
                        <div style={{ display: "flex", gap: 6 }}>
                          <button onClick={() => setEditOrderItem({ ...o })} style={{ ...btnGray, padding: "4px 10px", fontSize: 12 }}>✏️</button>
                          <button onClick={() => { setOrders(prev => prev.filter(x => x.id !== o.id)); showAlert("ลบ Order แล้ว"); }} style={{ ...btnRed, padding: "4px 10px", fontSize: 12 }}>🗑️</button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ marginTop: 12, fontSize: 12, color: "#8c8c8c", display: "flex", gap: 20, flexWrap: "wrap" }}>
            <span>แสดงเฉพาะรายการที่ยังไม่เสร็จ:</span>
            <Badge color="pending">PENDING = รอดำเนินการ</Badge>
            <Badge color="over3h">OVER 3 HOURS = ค้างเกิน 3 ชั่วโมง</Badge>
            <span style={{ color: "#595959" }}>NORMAL = ปกติ</span>
            <span>✅ รายการ COMPLETE อยู่ในปุ่ม "ดูรายการเสร็จแล้ว"</span>
          </div>
        </div>

        {/* Add Stock & Add Order */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 20 }}>
          <div style={sectionCard}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
              <div style={{ fontWeight: 700, fontSize: 16 }}>📦 Add Stock</div>
              <div style={{ display: "flex", gap: 8 }}>
                <button style={{ ...btnGray, fontSize: 12, padding: "5px 12px" }} onClick={downloadStockTemplate}>📄 Template</button>
                <input type="file" accept=".xlsx,.xls,.csv" id="stock-xlsx" style={{ display: "none" }} onChange={handleStockXlsx} />
                <button style={{ ...btnGreen, fontSize: 12, padding: "5px 12px" }} onClick={() => document.getElementById("stock-xlsx").click()}>📤 Upload Excel</button>
              </div>
            </div>
            {/* Drop zone hint */}
            <div
              style={{ border: "2px dashed #91d5ff", borderRadius: 8, padding: "10px 14px", marginBottom: 14, background: "#f0f9ff", fontSize: 12, color: "#0958d9", textAlign: "center", cursor: "pointer" }}
              onClick={() => document.getElementById("stock-xlsx").click()}
              onDragOver={e => e.preventDefault()}
              onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) { const inp = document.getElementById("stock-xlsx"); const dt = new DataTransfer(); dt.items.add(f); inp.files = dt.files; handleStockXlsx({ target: inp }); } }}
            >
              📂 คลิกหรือลากไฟล์ Excel / CSV มาวางที่นี่<br />
              <span style={{ color: "#8c8c8c" }}>รองรับคอลัมน์: barcode, product, stock, location</span>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {[["barcode", "Barcode"], ["product", "ชื่อสินค้า"], ["stock", "จำนวน", "number"], ["location", "Location"]].map(([key, label, type]) => (
                <div key={key}>
                  <label style={labelStyle}>{label}</label>
                  <input style={inputStyle} type={type || "text"} value={addStockForm[key]} onChange={e => setAddStockForm(f => ({ ...f, [key]: e.target.value }))} placeholder={label} />
                </div>
              ))}
              <button style={{ ...btnPrimary, marginTop: 4, width: "100%", justifyContent: "center" }} onClick={handleAddStock}>➕ เพิ่มสต็อก</button>
            </div>
          </div>

          <div style={sectionCard}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
              <div style={{ fontWeight: 700, fontSize: 16 }}>📋 Add Pick Order</div>
              <div style={{ display: "flex", gap: 8 }}>
                <button style={{ ...btnGray, fontSize: 12, padding: "5px 12px" }} onClick={downloadOrderTemplate}>📄 Template</button>
                <input type="file" accept=".xlsx,.xls,.csv" id="order-xlsx" style={{ display: "none" }} onChange={handleOrderXlsx} />
                <button style={{ ...btnOrange, fontSize: 12, padding: "5px 12px" }} onClick={() => document.getElementById("order-xlsx").click()}>📤 Upload Excel</button>
              </div>
            </div>
            {/* Drop zone hint */}
            <div
              style={{ border: "2px dashed #ffd591", borderRadius: 8, padding: "10px 14px", marginBottom: 14, background: "#fffbe6", fontSize: 12, color: "#874d00", textAlign: "center", cursor: "pointer" }}
              onClick={() => document.getElementById("order-xlsx").click()}
              onDragOver={e => e.preventDefault()}
              onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) { const inp = document.getElementById("order-xlsx"); const dt = new DataTransfer(); dt.items.add(f); inp.files = dt.files; handleOrderXlsx({ target: inp }); } }}
            >
              📂 คลิกหรือลากไฟล์ Excel / CSV มาวางที่นี่<br />
              <span style={{ color: "#8c8c8c" }}>รองรับคอลัมน์: pickNo, customer, barcode, product, location, required</span>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {/* Pick No */}
              <div>
                <label style={labelStyle}>Pick No (ไม่ต้องใส่ก็ได้)</label>
                <input style={inputStyle} value={addOrderForm.pickNo} onChange={e => setAddOrderForm(f => ({ ...f, pickNo: e.target.value }))} placeholder="Pick No (ไม่ต้องใส่ก็ได้)" />
              </div>
              {/* Customer */}
              <div>
                <label style={labelStyle}>ชื่อลูกค้า</label>
                <input style={inputStyle} value={addOrderForm.customer} onChange={e => setAddOrderForm(f => ({ ...f, customer: e.target.value }))} placeholder="ชื่อลูกค้า" />
              </div>
              {/* Barcode — triggers auto-fill */}
              <div>
                <label style={labelStyle}>Barcode *</label>
                <input
                  style={{ ...inputStyle, border: addOrderForm.barcode && stock.find(s => s.barcode === addOrderForm.barcode.trim()) ? "2px solid #52c41a" : inputStyle.border }}
                  value={addOrderForm.barcode}
                  onChange={e => {
                    const bc = e.target.value;
                    const found = stock.find(s => s.barcode === bc.trim());
                    setAddOrderForm(f => ({
                      ...f,
                      barcode: bc,
                      product: found ? found.product : f.product,
                      location: found ? found.location : f.location,
                    }));
                  }}
                  placeholder="Barcode *"
                />
                {addOrderForm.barcode && (() => {
                  const found = stock.find(s => s.barcode === addOrderForm.barcode.trim());
                  return found
                    ? <div style={{ marginTop: 4, fontSize: 12, color: "#52c41a", fontWeight: 600 }}>✅ พบสินค้า: {found.product} | Stock: {found.stock} | {found.location}</div>
                    : addOrderForm.barcode.length > 3
                      ? <div style={{ marginTop: 4, fontSize: 12, color: "#fa8c16" }}>⚠️ ไม่พบ Barcode นี้ใน Stock — กรอกชื่อสินค้าด้านล่างเอง</div>
                      : null;
                })()}
              </div>
              {/* Product — auto or manual */}
              <div>
                <label style={labelStyle}>ชื่อสินค้า *</label>
                <input style={inputStyle} value={addOrderForm.product} onChange={e => setAddOrderForm(f => ({ ...f, product: e.target.value }))} placeholder="ชื่อสินค้า *" />
              </div>
              {/* Location — auto or manual */}
              <div>
                <label style={labelStyle}>Location</label>
                <input style={inputStyle} value={addOrderForm.location} onChange={e => setAddOrderForm(f => ({ ...f, location: e.target.value }))} placeholder="Location" />
              </div>
              {/* Required qty */}
              <div>
                <label style={labelStyle}>จำนวนที่ต้องหยิบ</label>
                <input style={inputStyle} type="number" value={addOrderForm.required} onChange={e => setAddOrderForm(f => ({ ...f, required: e.target.value }))} placeholder="จำนวนที่ต้องหยิบ" />
              </div>
              <button style={{ ...btnOrange, marginTop: 4, width: "100%", justifyContent: "center" }} onClick={handleAddOrder}>➕ เพิ่ม Order</button>
            </div>
          </div>
        </div>

        {/* Stock Table */}
        <div style={sectionCard}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
            <span style={{ fontWeight: 700, fontSize: 16 }}>📦 Stock คงเหลือ</span>
            <div style={{ flex: 1 }}>
              <input style={{ ...inputStyle, maxWidth: 280 }} value={stockSearch} onChange={e => setStockSearch(e.target.value)} placeholder="ค้นหา Stock..." />
            </div>
            <button style={btnRed} onClick={() => { if (window.confirm("ลบ Stock ทั้งหมด?")) { setStock([]); showAlert("ลบ Stock ทั้งหมดแล้ว"); } }}>🗑️ ลบทั้งหมด</button>
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 14 }}>
              <thead>
                <tr style={{ background: "#fafafa", borderBottom: "1px solid #f0f0f0" }}>
                  {["Barcode", "Product", "Stock", "Location", "แก้ไข"].map(h => (
                    <th key={h} style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, color: "#595959" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredStock.map(s => (
                  <tr key={s.barcode} style={{ borderBottom: "1px solid #f5f5f5" }}>
                    <td style={{ padding: "10px 14px", fontFamily: "monospace" }}>{s.barcode}</td>
                    <td style={{ padding: "10px 14px" }}>{s.product}</td>
                    <td style={{ padding: "10px 14px", color: s.stock > 0 ? "#52c41a" : "#ff4d4f", fontWeight: 700 }}>{s.stock}</td>
                    <td style={{ padding: "10px 14px" }}>{s.location}</td>
                    <td style={{ padding: "10px 14px" }}>
                      <div style={{ display: "flex", gap: 8 }}>
                        <button onClick={() => setEditStockItem({ ...s })} style={{ ...btnPrimary, padding: "5px 14px", fontSize: 13 }}>✏️ แก้ไข</button>
                        <button onClick={() => { setStock(prev => prev.filter(x => x.barcode !== s.barcode)); showAlert("ลบสินค้าแล้ว"); }} style={{ ...btnRed, padding: "5px 14px", fontSize: 13 }}>🗑️ ลบ</button>
                      </div>
                    </td>
                  </tr>
                ))}
                {filteredStock.length === 0 && (
                  <tr><td colSpan={5} style={{ textAlign: "center", padding: 30, color: "#8c8c8c" }}>ไม่พบสินค้า</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Edit Stock Modal */}
      <Modal open={!!editStockItem} onClose={() => setEditStockItem(null)} title="แก้ไขสต็อก">
        {editStockItem && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            {[["barcode", "Barcode"], ["product", "ชื่อสินค้า"], ["stock", "จำนวน", "number"], ["location", "Location"]].map(([key, label, type]) => (
              <div key={key}>
                <label style={labelStyle}>{label}</label>
                <input style={inputStyle} type={type || "text"} value={editStockItem[key]} onChange={e => setEditStockItem(f => ({ ...f, [key]: type === "number" ? Number(e.target.value) : e.target.value }))} />
              </div>
            ))}
            <div style={{ display: "flex", gap: 10, marginTop: 8 }}>
              <button style={{ ...btnPrimary, flex: 1 }} onClick={handleEditStock}>💾 บันทึก</button>
              <button style={{ ...btnGray, flex: 1 }} onClick={() => setEditStockItem(null)}>ยกเลิก</button>
            </div>
          </div>
        )}
      </Modal>

      {/* Edit Order Modal */}
      <Modal open={!!editOrderItem} onClose={() => setEditOrderItem(null)} title="แก้ไข Order">
        {editOrderItem && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            {[["pickNo", "Pick No"], ["customer", "ลูกค้า"], ["barcode", "Barcode"], ["product", "สินค้า"], ["location", "Location"], ["required", "จำนวนที่ต้องหยิบ", "number"], ["picked", "หยิบแล้ว", "number"]].map(([key, label, type]) => (
              <div key={key}>
                <label style={labelStyle}>{label}</label>
                <input style={inputStyle} type={type || "text"} value={editOrderItem[key] ?? ""} onChange={e => setEditOrderItem(f => ({ ...f, [key]: type === "number" ? Number(e.target.value) : e.target.value }))} />
              </div>
            ))}
            <div>
              <label style={labelStyle}>Status</label>
              <select style={inputStyle} value={editOrderItem.status} onChange={e => setEditOrderItem(f => ({ ...f, status: e.target.value }))}>
                <option>PENDING</option>
                <option>COMPLETE</option>
              </select>
            </div>
            <div style={{ display: "flex", gap: 10, marginTop: 8 }}>
              <button style={{ ...btnPrimary, flex: 1 }} onClick={handleEditOrder}>💾 บันทึก</button>
              <button style={{ ...btnGray, flex: 1 }} onClick={() => setEditOrderItem(null)}>ยกเลิก</button>
            </div>
          </div>
        )}
      </Modal>

      {/* Excel Preview Modal */}
      <Modal open={!!xlsxPreview} onClose={() => setXlsxPreview(null)} title={xlsxPreview?.type === "stock" ? "📤 ตรวจสอบข้อมูล Stock ก่อนนำเข้า" : "📤 ตรวจสอบข้อมูล Order ก่อนนำเข้า"}>
        {xlsxPreview && (
          <div>
            {xlsxPreview.errors.length > 0 && (
              <div style={{ background: "#fff1f0", border: "1px solid #ffa39e", borderRadius: 8, padding: "10px 14px", marginBottom: 14, fontSize: 13, color: "#a8071a" }}>
                ⚠️ พบข้อผิดพลาด {xlsxPreview.errors.length} รายการ:<br />
                {xlsxPreview.errors.slice(0, 5).map((e, i) => <div key={i}>• {e}</div>)}
                {xlsxPreview.errors.length > 5 && <div>...และอีก {xlsxPreview.errors.length - 5} รายการ</div>}
              </div>
            )}
            <div style={{ background: "#f6ffed", border: "1px solid #b7eb8f", borderRadius: 8, padding: "10px 14px", marginBottom: 14, fontSize: 13, color: "#237804" }}>
              ✅ พบข้อมูลที่นำเข้าได้ <strong>{xlsxPreview.rows.length}</strong> แถว
            </div>
            <div style={{ maxHeight: 240, overflowY: "auto", border: "1px solid #f0f0f0", borderRadius: 8, marginBottom: 16 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: "#fafafa", position: "sticky", top: 0 }}>
                    {xlsxPreview.type === "stock"
                      ? ["Barcode", "Product", "Stock", "Location"].map(h => <th key={h} style={{ padding: "7px 10px", textAlign: "left", fontWeight: 600, color: "#595959" }}>{h}</th>)
                      : ["Pick No", "Customer", "Barcode", "Product", "Location", "Required"].map(h => <th key={h} style={{ padding: "7px 10px", textAlign: "left", fontWeight: 600, color: "#595959" }}>{h}</th>)
                    }
                  </tr>
                </thead>
                <tbody>
                  {xlsxPreview.rows.slice(0, 50).map((r, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #f5f5f5" }}>
                      {xlsxPreview.type === "stock"
                        ? [r.barcode, r.product, r.stock, r.location].map((v, j) => <td key={j} style={{ padding: "6px 10px", fontFamily: j === 0 ? "monospace" : "inherit" }}>{v}</td>)
                        : [r.pickNo, r.customer, r.barcode, r.product, r.location, r.required].map((v, j) => <td key={j} style={{ padding: "6px 10px", fontFamily: j === 2 ? "monospace" : "inherit" }}>{v}</td>)
                      }
                    </tr>
                  ))}
                  {xlsxPreview.rows.length > 50 && (
                    <tr><td colSpan={6} style={{ padding: "8px 10px", color: "#8c8c8c", fontSize: 12, textAlign: "center" }}>...แสดง 50 แถวแรก จากทั้งหมด {xlsxPreview.rows.length} แถว</td></tr>
                  )}
                </tbody>
              </table>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button style={{ ...btnGreen, flex: 1, justifyContent: "center" }} onClick={confirmXlsxImport}>✅ ยืนยันนำเข้า {xlsxPreview.rows.length} รายการ</button>
              <button style={{ ...btnGray, flex: 1, justifyContent: "center" }} onClick={() => setXlsxPreview(null)}>ยกเลิก</button>
            </div>
          </div>
        )}
      </Modal>
      <Modal open={confirmClear} onClose={() => setConfirmClear(false)} title="⚠️ ล้างข้อมูลวันนี้">
        <p style={{ color: "#595959", marginBottom: 20 }}>ต้องการล้างข้อมูล Orders วันนี้ทั้งหมดหรือไม่? (Stock จะไม่ถูกลบ)</p>
        <div style={{ display: "flex", gap: 10 }}>
          <button style={{ ...btnRed, flex: 1 }} onClick={() => {
            const today = new Date().toLocaleDateString("th-TH");
            setOrders(prev => prev.filter(o => new Date(o.createdAtRaw).toLocaleDateString("th-TH") !== today));
            setConfirmClear(false); showAlert("ล้างข้อมูลวันนี้แล้ว");
          }}>ยืนยันล้างข้อมูล</button>
          <button style={{ ...btnGray, flex: 1 }} onClick={() => setConfirmClear(false)}>ยกเลิก</button>
        </div>
      </Modal>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@400;600;700;800&display=swap');
        * { box-sizing: border-box; }
        input:focus, select:focus { border-color: #40a9ff !important; outline: none; box-shadow: 0 0 0 2px rgba(24,144,255,0.15); }
        @keyframes fadein { from { opacity: 0; transform: translateY(-8px); } to { opacity: 1; transform: translateY(0); } }
      `}</style>
    </div>
  );
}
