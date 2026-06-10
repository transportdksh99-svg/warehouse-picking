import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getDatabase, ref as dbRef, onValue, set } from "firebase/database";

const firebaseConfig = {
  apiKey: "AIzaSyDbxbu1w9VSapEWT2ApUvvPBz5xUdj29vU",
  authDomain: "warehouse-app-52001.firebaseapp.com",
  databaseURL: "https://warehouse-app-52001-default-rtdb.firebaseio.com",
  projectId: "warehouse-app-52001",
  storageBucket: "warehouse-app-52001.firebasestorage.app",
  messagingSenderId: "72442819990",
  appId: "1:72442819990:web:b349bfe2bb1e08679259b7",
};

const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

function generateId() {
  return (
    "PO" + Date.now().toString().slice(-6) + Math.floor(Math.random() * 100)
  );
}
function now() {
  const d = new Date(),
    pad = (n) => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${
    d.getFullYear() + 543
  } ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}
function cleanObj(obj) {
  return Object.fromEntries(
    Object.entries(obj).filter(([, v]) => v !== undefined && v !== null)
  );
}

// ── UI atoms ──────────────────────────────────────────────────────────────────
function Badge({ children, color = "gray" }) {
  const C = {
    pending: { bg: "#fff7e6", text: "#d46b08", border: "#ffd591" },
    complete: { bg: "#f6ffed", text: "#389e0d", border: "#b7eb8f" },
    over3h: { bg: "#fff1f0", text: "#cf1322", border: "#ffa39e" },
    gray: { bg: "#f5f5f5", text: "#595959", border: "#d9d9d9" },
  };
  const c = C[color] || C.gray;
  return (
    <span
      style={{
        background: c.bg,
        color: c.text,
        border: `1px solid ${c.border}`,
        borderRadius: 6,
        padding: "2px 10px",
        fontSize: 12,
        fontWeight: 600,
      }}
    >
      {children}
    </span>
  );
}
function StatCard({ icon, value, label, color = "#1677ff" }) {
  return (
    <div
      style={{
        background: "#fff",
        borderRadius: 12,
        border: "1px solid #f0f0f0",
        padding: "18px 20px",
        flex: 1,
        minWidth: 140,
        boxShadow: "0 1px 4px rgba(0,0,0,0.06)",
      }}
    >
      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        <div
          style={{
            width: 44,
            height: 44,
            borderRadius: 10,
            background: color + "18",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: 22,
          }}
        >
          {icon}
        </div>
        <div>
          <div
            style={{ fontSize: 26, fontWeight: 700, color, lineHeight: 1.1 }}
          >
            {value}
          </div>
          <div style={{ fontSize: 13, color: "#8c8c8c", marginTop: 2 }}>
            {label}
          </div>
        </div>
      </div>
    </div>
  );
}
function Modal({ open, onClose, title, children }) {
  if (!open) return null;
  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,0.45)",
        zIndex: 1000,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      }}
      onClick={onClose}
    >
      <div
        style={{
          background: "#fff",
          borderRadius: 14,
          padding: "28px 32px",
          minWidth: 380,
          maxWidth: 520,
          width: "90%",
          boxShadow: "0 8px 40px rgba(0,0,0,0.18)",
          maxHeight: "90vh",
          overflowY: "auto",
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            marginBottom: 20,
          }}
        >
          <span style={{ fontWeight: 700, fontSize: 17 }}>{title}</span>
          <button
            onClick={onClose}
            style={{
              background: "none",
              border: "none",
              cursor: "pointer",
              fontSize: 20,
              color: "#8c8c8c",
            }}
          >
            ✕
          </button>
        </div>
        {children}
      </div>
    </div>
  );
}

// ── Style tokens ──────────────────────────────────────────────────────────────
const inp = {
  border: "1px solid #d9d9d9",
  borderRadius: 8,
  padding: "8px 12px",
  fontSize: 14,
  width: "100%",
  outline: "none",
  boxSizing: "border-box",
  fontFamily: "inherit",
  transition: "border 0.2s",
};
const B = {
  base: {
    border: "none",
    borderRadius: 8,
    padding: "9px 18px",
    fontSize: 14,
    fontWeight: 600,
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
  },
};
const bP = { ...B.base, background: "#1677ff", color: "#fff" };
const bG = { ...B.base, background: "#52c41a", color: "#fff" };
const bO = { ...B.base, background: "#fa8c16", color: "#fff" };
const bR = { ...B.base, background: "#ff4d4f", color: "#fff" };
const bGr = {
  ...B.base,
  background: "#fff",
  color: "#595959",
  border: "1px solid #d9d9d9",
};
const lbl = {
  fontSize: 13,
  color: "#595959",
  marginBottom: 4,
  display: "block",
};
const card = {
  background: "#fff",
  borderRadius: 14,
  border: "1px solid #f0f0f0",
  padding: "22px 24px",
  marginBottom: 20,
  boxShadow: "0 1px 4px rgba(0,0,0,0.05)",
};

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [stock, setStock] = useState([]);
  const [orders, setOrders] = useState([]);
  const [dbReady, setDbReady] = useState(false);

  const [pickNo, setPickNo] = useState("");
  const [barcode, setBarcode] = useState("");
  const [location, setLocation] = useState("");
  const [qty, setQty] = useState(1);
  const [scanStatus, setScanStatus] = useState("พร้อมสแกน");
  const [scanStatusColor, setScanStatusColor] = useState("#52c41a");
  const [autoConfirm, setAutoConfirm] = useState(true);
  const [searchOrder, setSearchOrder] = useState("");
  const [showCompleted, setShowCompleted] = useState(false);

  const [addStockForm, setAddStockForm] = useState({
    barcode: "",
    product: "",
    stock: "",
    location: "",
  });
  const [addOrderForm, setAddOrderForm] = useState({
    pickNo: "",
    customer: "",
    barcode: "",
    product: "",
    location: "",
    required: 1,
  });
  const [editStockItem, setEditStockItem] = useState(null);
  const [editOrderItem, setEditOrderItem] = useState(null);
  const [confirmClear, setConfirmClear] = useState(false);
  const [alertMsg, setAlertMsg] = useState(null);
  const [xlsxPreview, setXlsxPreview] = useState(null);
  const [stockSearch, setStockSearch] = useState("");

  const barcodeRef = useRef();

  // ── Firebase ────────────────────────────────────────────────────────────────
  useEffect(() => {
    let sL = false,
      oL = false;
    const check = () => {
      if (sL && oL) setDbReady(true);
    };
    const u1 = onValue(dbRef(db, "stock"), (s) => {
      setStock(s.val() ? Object.values(s.val()) : []);
      sL = true;
      check();
    });
    const u2 = onValue(dbRef(db, "orders"), (s) => {
      setOrders(s.val() ? Object.values(s.val()) : []);
      oL = true;
      check();
    });
    return () => {
      u1();
      u2();
    };
  }, []);

  const updateStock = useCallback((fn) => {
    setStock((prev) => {
      const next = typeof fn === "function" ? fn(prev) : fn;
      const obj = {};
      next.forEach((s) => {
        obj[String(s.barcode).replace(/[^a-zA-Z0-9_-]/g, "_")] = cleanObj(s);
      });
      set(dbRef(db, "stock"), next.length ? obj : null);
      return next;
    });
  }, []);

  const updateOrders = useCallback((fn) => {
    setOrders((prev) => {
      const next = typeof fn === "function" ? fn(prev) : fn;
      const obj = {};
      next.forEach((o) => {
        obj[o.id] = cleanObj(o);
      });
      set(dbRef(db, "orders"), next.length ? obj : null);
      return next;
    });
  }, []);

  // ── Alert (success / warning / error) ──────────────────────────────────────
  const showAlert = useCallback((msg, type = "success") => {
    setAlertMsg({ msg, type });
    setTimeout(() => setAlertMsg(null), 3000);
  }, []);

  // ── Derived ─────────────────────────────────────────────────────────────────
  const pendingOrders = orders.filter((o) => o.status !== "COMPLETE");
  const completedOrders = orders.filter((o) => o.status === "COMPLETE");
  const displayOrders = showCompleted ? completedOrders : pendingOrders;

  const filteredOrders = displayOrders
    .filter((o) => {
      const q = searchOrder.toLowerCase();
      return (
        !q ||
        o.pickNo?.toLowerCase().includes(q) ||
        o.customer?.toLowerCase().includes(q) ||
        o.product?.toLowerCase().includes(q) ||
        o.barcode?.toLowerCase().includes(q)
      );
    })
    .sort((a, b) =>
      showCompleted
        ? new Date(b.completedAtRaw || 0) - new Date(a.completedAtRaw || 0)
        : new Date(b.createdAtRaw || 0) - new Date(a.createdAtRaw || 0)
    );

  const filteredStock = stock.filter((s) => {
    const q = stockSearch.toLowerCase();
    return (
      !q ||
      s.barcode.includes(q) ||
      s.product.toLowerCase().includes(q) ||
      s.location.toLowerCase().includes(q)
    );
  });

  function getOrderStatus(o) {
    if (o.status === "COMPLETE") return "complete";
    if (Date.now() - new Date(o.createdAtRaw) > 3 * 60 * 60 * 1000)
      return "over3h";
    return "pending";
  }

  // ── Sound ────────────────────────────────────────────────────────────────────
  function playSound(type) {
    try {
      const ctx = new (window.AudioContext || window.webkitAudioContext)();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(ctx.destination);
      if (type === "success") {
        osc.frequency.setValueAtTime(880, ctx.currentTime);
        osc.frequency.setValueAtTime(1100, ctx.currentTime + 0.1);
        gain.gain.setValueAtTime(0.3, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.3);
        osc.start(ctx.currentTime);
        osc.stop(ctx.currentTime + 0.3);
      } else if (type === "warning") {
        osc.frequency.setValueAtTime(600, ctx.currentTime);
        osc.frequency.setValueAtTime(500, ctx.currentTime + 0.1);
        osc.frequency.setValueAtTime(600, ctx.currentTime + 0.2);
        gain.gain.setValueAtTime(0.3, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.35);
        osc.start(ctx.currentTime);
        osc.stop(ctx.currentTime + 0.35);
      } else {
        osc.frequency.setValueAtTime(300, ctx.currentTime);
        osc.frequency.setValueAtTime(200, ctx.currentTime + 0.15);
        gain.gain.setValueAtTime(0.4, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.4);
        osc.start(ctx.currentTime);
        osc.stop(ctx.currentTime + 0.4);
      }
    } catch {}
  }

  // ── Pick ─────────────────────────────────────────────────────────────────────
  function executePick(pickQty, stockItem, matchedOrder) {
    // ลด stock
    updateStock((prev) =>
      prev.map((s) =>
        s.barcode === stockItem.barcode ? { ...s, stock: s.stock - pickQty } : s
      )
    );

    const newPicked = (matchedOrder.picked || 0) + pickQty;
    const isComplete = newPicked >= matchedOrder.required;
    const remaining = matchedOrder.required - newPicked;
    const completedAt = isComplete ? now() : undefined;
    const completedAtRaw = isComplete ? new Date().toISOString() : undefined;

    updateOrders((prev) =>
      prev.map((o) =>
        o.id === matchedOrder.id
          ? cleanObj({
              ...o,
              picked: newPicked,
              status: isComplete ? "COMPLETE" : o.status,
              completedAt,
              completedAtRaw,
            })
          : o
      )
    );

    if (isComplete) {
      // ✅ หยิบครบ
      playSound("success");
      setScanStatus(`✅ หยิบครบ: ${stockItem.product} x${pickQty}`);
      setScanStatusColor("#52c41a");
    } else {
      // ⚠️ หยิบบางส่วน — แจ้งเตือนเฉยๆ ไม่ block
      playSound("warning");
      setScanStatus(
        `⚠️ หยิบบางส่วน: ${stockItem.product} x${pickQty} — ยังต้องหยิบอีก ${remaining} ชิ้น`
      );
      setScanStatusColor("#fa8c16");
      showAlert(
        `⚠️ Order ${matchedOrder.pickNo} — หยิบแล้ว ${newPicked}/${matchedOrder.required} ยังเหลืออีก ${remaining} ชิ้น`,
        "warning"
      );
    }

    setBarcode("");
    setLocation("");
    setQty(1);
    setTimeout(() => {
      setScanStatus("พร้อมสแกน");
      setScanStatusColor("#52c41a");
      barcodeRef.current?.focus();
    }, 2200);
  }

  function confirmPick() {
    if (!barcode.trim()) {
      playSound("error");
      setScanStatus("❌ กรุณากรอก Barcode");
      setScanStatusColor("#ff4d4f");
      return;
    }
    const pickQty = parseInt(String(qty), 10);
    if (isNaN(pickQty) || pickQty <= 0) {
      playSound("error");
      setScanStatus("❌ กรุณาระบุจำนวน (Qty) ที่ถูกต้อง");
      setScanStatusColor("#ff4d4f");
      return;
    }
    const stockItem = stock.find((s) => s.barcode === barcode.trim());
    if (!stockItem) {
      playSound("error");
      setScanStatus("❌ ไม่พบสินค้าในระบบ");
      setScanStatusColor("#ff4d4f");
      return;
    }
    if (stockItem.stock < pickQty) {
      playSound("error");
      setScanStatus(`❌ สต็อกไม่เพียงพอ (มี ${stockItem.stock} ชิ้น)`);
      setScanStatusColor("#ff4d4f");
      return;
    }

    // หา Order ที่ตรงกัน
    let matchedOrder = null;
    if (pickNo.trim()) {
      matchedOrder = orders.find(
        (o) =>
          o.pickNo === pickNo.trim() &&
          o.barcode === barcode.trim() &&
          o.status !== "COMPLETE"
      );
    } else {
      matchedOrder = orders.find(
        (o) => o.barcode === barcode.trim() && o.status !== "COMPLETE"
      );
    }

    // ✅ ต้องมี Order — ห้ามตัด Stock โดยไม่มี Order
    if (!matchedOrder) {
      playSound("error");
      setScanStatus("❌ ไม่พบ Order ที่ตรงกับ Barcode นี้");
      setScanStatusColor("#ff4d4f");
      return;
    }

    // ✅ หยิบได้เลยทุกกรณี (ทั้งครบและบางส่วน) — แค่แจ้งเตือนถ้าไม่ครบ
    executePick(pickQty, stockItem, matchedOrder);
  }

  function handleBarcodeKeyDown(e) {
    if (e.key === "Enter") {
      const found = stock.find((s) => s.barcode === barcode.trim());
      if (found && !location) setLocation(found.location);
      if (autoConfirm) setTimeout(() => confirmPick(), 50);
    }
  }

  // ── Edit Order (fix: ตรวจ id + handle completedAt) ──────────────────────────
  function handleEditOrder() {
    const original = orders.find((o) => o.id === editOrderItem.id);
    if (!original) {
      showAlert("ไม่พบ Order นี้ในระบบ", "error");
      return;
    }

    const isNowComplete =
      editOrderItem.status === "COMPLETE" && original.status !== "COMPLETE";
    const updated = cleanObj({
      ...editOrderItem,
      completedAt: isNowComplete
        ? now()
        : editOrderItem.completedAt ?? original.completedAt,
      completedAtRaw: isNowComplete
        ? new Date().toISOString()
        : editOrderItem.completedAtRaw ?? original.completedAtRaw,
    });
    updateOrders((prev) =>
      prev.map((o) => (o.id === updated.id ? updated : o))
    );
    setEditOrderItem(null);
    showAlert("แก้ไข Order สำเร็จ");
  }

  // ── Edit Stock ───────────────────────────────────────────────────────────────
  function handleEditStock() {
    updateStock((prev) =>
      prev.map((s) => {
        const key = editStockItem._originalBarcode || editStockItem.barcode;
        if (s.barcode === key) {
          const { _originalBarcode, ...rest } = editStockItem;
          return rest;
        }
        return s;
      })
    );
    setEditStockItem(null);
    showAlert("แก้ไข Stock สำเร็จ");
  }

  // ── Add Stock / Order ────────────────────────────────────────────────────────
  function handleAddStock() {
    const { barcode: bc, product, stock: st, location: loc } = addStockForm;
    if (!bc || !product || !st || !loc) {
      showAlert("กรุณากรอกข้อมูลให้ครบ", "error");
      return;
    }
    if (stock.find((s) => s.barcode === bc)) {
      updateStock((prev) =>
        prev.map((s) =>
          s.barcode === bc
            ? { ...s, stock: s.stock + Number(st), location: loc }
            : s
        )
      );
      showAlert("อัพเดทสต็อกสำเร็จ");
    } else {
      updateStock((prev) => [
        ...prev,
        { barcode: bc, product, stock: Number(st), location: loc },
      ]);
      showAlert("เพิ่มสินค้าสำเร็จ");
    }
    setAddStockForm({ barcode: "", product: "", stock: "", location: "" });
  }

  function handleAddOrder() {
    const {
      customer,
      barcode: bc,
      product,
      location: loc,
      required,
    } = addOrderForm;
    if (!bc || !product) {
      showAlert("กรุณากรอก Barcode และชื่อสินค้า", "error");
      return;
    }
    const id = generateId();
    updateOrders((prev) => [
      ...prev,
      cleanObj({
        id,
        pickNo: addOrderForm.pickNo || id,
        customer,
        barcode: bc,
        product,
        location: loc,
        required: Number(required) || 1,
        picked: 0,
        status: "PENDING",
        createdAt: now(),
        createdAtRaw: new Date().toISOString(),
      }),
    ]);
    setAddOrderForm({
      pickNo: "",
      customer: "",
      barcode: "",
      product: "",
      location: "",
      required: 1,
    });
    showAlert("เพิ่ม Order สำเร็จ");
  }

  // ── Export ───────────────────────────────────────────────────────────────────
  function exportExcel() {
    const rows = [
      [
        "Pick No",
        "Customer",
        "Barcode",
        "Product",
        "Location",
        "Required",
        "Picked",
        "Status",
        "Created At",
        "Completed At",
        "ใช้เวลา",
      ],
    ];
    orders.forEach((o) => {
      let dur = "";
      if (o.completedAtRaw && o.createdAtRaw) {
        const s = Math.round(
          (new Date(o.completedAtRaw) - new Date(o.createdAtRaw)) / 1000
        );
        dur =
          s < 60
            ? s + " วินาที"
            : s < 3600
            ? Math.round(s / 60) + " นาที"
            : Math.floor(s / 3600) +
              "ชม. " +
              Math.round((s % 3600) / 60) +
              "น.";
      }
      rows.push([
        o.pickNo,
        o.customer,
        o.barcode,
        o.product,
        o.location,
        o.required,
        o.picked || 0,
        o.status,
        o.createdAt,
        o.completedAt || "",
        dur,
      ]);
    });
    const csv = rows.map((r) => r.map((c) => `"${c}"`).join(",")).join("\n");
    const blob = new Blob(["\uFEFF" + csv], {
      type: "text/csv;charset=utf-8;",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "pick_orders.csv";
    a.click();
  }

  // ── Backup / Restore ─────────────────────────────────────────────────────────
  function backupData() {
    const blob = new Blob(
      [
        JSON.stringify(
          { stock, orders, backupAt: new Date().toISOString() },
          null,
          2
        ),
      ],
      { type: "application/json" }
    );
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `warehouse_backup_${Date.now()}.json`;
    a.click();
    showAlert("Backup สำเร็จ");
  }
  function restoreData(e) {
    const file = e.target.files[0];
    if (!file) return;
    const r = new FileReader();
    r.onload = (ev) => {
      try {
        const d = JSON.parse(ev.target.result);
        if (d.stock) updateStock(d.stock);
        if (d.orders) updateOrders(d.orders);
        showAlert("Restore สำเร็จ");
      } catch {
        showAlert("ไฟล์ไม่ถูกต้อง", "error");
      }
    };
    r.readAsText(file);
    e.target.value = "";
  }

  // ── Excel import ─────────────────────────────────────────────────────────────
  function downloadStockTemplate() {
    const ws = XLSX.utils.aoa_to_sheet([
      ["barcode", "product", "stock", "location"],
      ["8851234567890", "สินค้า A", 100, "A01-01"],
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
    ]);
    ws["!cols"] = [
      { wch: 14 },
      { wch: 16 },
      { wch: 20 },
      { wch: 20 },
      { wch: 14 },
      { wch: 12 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Orders");
    XLSX.writeFile(wb, "order_template.xlsx");
  }
  function parseStockXlsx(file) {
    const r = new FileReader();
    r.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const errors = [];
        const rows = raw
          .map((x, i) => {
            const row = {
              barcode: String(
                x.barcode || x.Barcode || x["บาร์โค้ด"] || ""
              ).trim(),
              product: String(
                x.product || x.Product || x["ชื่อสินค้า"] || ""
              ).trim(),
              stock: Number(x.stock || x.Stock || x["จำนวน"] || 0),
              location: String(
                x.location || x.Location || x["ที่เก็บ"] || ""
              ).trim(),
            };
            if (!row.barcode) errors.push(`แถว ${i + 2}: ไม่มี Barcode`);
            if (!row.product) errors.push(`แถว ${i + 2}: ไม่มีชื่อสินค้า`);
            return row;
          })
          .filter((x) => x.barcode);
        setXlsxPreview({ type: "stock", rows, errors });
      } catch {
        showAlert("ไม่สามารถอ่านไฟล์ได้", "error");
      }
    };
    r.readAsArrayBuffer(file);
  }
  function parseOrderXlsx(file) {
    const r = new FileReader();
    r.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const errors = [];
        const rows = raw
          .map((x, i) => {
            const row = {
              pickNo: String(
                x.pickNo || x["Pick No"] || x.pick_no || ""
              ).trim(),
              customer: String(
                x.customer || x.Customer || x["ลูกค้า"] || ""
              ).trim(),
              barcode: String(
                x.barcode || x.Barcode || x["บาร์โค้ด"] || ""
              ).trim(),
              product: String(
                x.product || x.Product || x["สินค้า"] || x["ชื่อสินค้า"] || ""
              ).trim(),
              location: String(
                x.location || x.Location || x["ที่เก็บ"] || ""
              ).trim(),
              required: Number(x.required || x.Required || x["จำนวน"] || 1),
            };
            if (!row.barcode) errors.push(`แถว ${i + 2}: ไม่มี Barcode`);
            if (!row.product) errors.push(`แถว ${i + 2}: ไม่มีชื่อสินค้า`);
            return row;
          })
          .filter((x) => x.barcode);
        setXlsxPreview({ type: "order", rows, errors });
      } catch {
        showAlert("ไม่สามารถอ่านไฟล์ได้", "error");
      }
    };
    r.readAsArrayBuffer(file);
  }
  function handleStockXlsx(e) {
    const f = e.target.files[0];
    if (f) parseStockXlsx(f);
    e.target.value = "";
  }
  function handleOrderXlsx(e) {
    const f = e.target.files[0];
    if (f) parseOrderXlsx(f);
    e.target.value = "";
  }
  function confirmXlsxImport() {
    if (!xlsxPreview) return;
    if (xlsxPreview.type === "stock") {
      updateStock((prev) => {
        const u = [...prev];
        xlsxPreview.rows.forEach((r) => {
          const i = u.findIndex((s) => s.barcode === r.barcode);
          if (i >= 0)
            u[i] = {
              ...u[i],
              stock: u[i].stock + r.stock,
              location: r.location || u[i].location,
              product: r.product || u[i].product,
            };
          else u.push(r);
        });
        return u;
      });
      showAlert(`นำเข้าสต็อกสำเร็จ ${xlsxPreview.rows.length} รายการ`);
    } else {
      updateOrders((prev) => [
        ...prev,
        ...xlsxPreview.rows.map((r) => {
          const id = generateId();
          return cleanObj({
            id,
            pickNo: r.pickNo || id,
            customer: r.customer,
            barcode: r.barcode,
            product: r.product,
            location: r.location,
            required: r.required,
            picked: 0,
            status: "PENDING",
            createdAt: now(),
            createdAtRaw: new Date().toISOString(),
          });
        }),
      ]);
      showAlert(`นำเข้า Order สำเร็จ ${xlsxPreview.rows.length} รายการ`);
    }
    setXlsxPreview(null);
  }

  // ── Duration ─────────────────────────────────────────────────────────────────
  function fmtDur(o) {
    if (!o.completedAtRaw || !o.createdAtRaw) return null;
    const s = Math.round(
        (new Date(o.completedAtRaw) - new Date(o.createdAtRaw)) / 1000
      ),
      m = Math.floor(s / 60);
    if (s < 60)
      return (
        <span style={{ color: "#722ed1", fontWeight: 700, fontSize: 13 }}>
          {s} วิ
        </span>
      );
    if (m < 60)
      return (
        <span style={{ color: "#52c41a", fontWeight: 700, fontSize: 13 }}>
          {m} นาที
        </span>
      );
    return (
      <span style={{ color: "#fa8c16", fontWeight: 700, fontSize: 13 }}>
        {Math.floor(m / 60)}ชม. {m % 60}น.
      </span>
    );
  }

  // ── Loading ──────────────────────────────────────────────────────────────────
  if (!dbReady)
    return (
      <div
        style={{
          background: "#f7f8fa",
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          flexDirection: "column",
          gap: 16,
          fontFamily: "'Noto Sans Thai',sans-serif",
        }}
      >
        <div style={{ fontSize: 40 }}>🔥</div>
        <div style={{ fontSize: 18, fontWeight: 700, color: "#065A82" }}>
          กำลังเชื่อมต่อ Firebase...
        </div>
        <div style={{ fontSize: 13, color: "#8c8c8c" }}>
          โหลดข้อมูล real-time จากฐานข้อมูลกลาง
        </div>
      </div>
    );

  // ── Render ───────────────────────────────────────────────────────────────────
  const alertBg =
    alertMsg?.type === "error"
      ? "#ff4d4f"
      : alertMsg?.type === "warning"
      ? "#fa8c16"
      : "#52c41a";

  return (
    <div
      style={{
        background: "#f7f8fa",
        minHeight: "100vh",
        fontFamily: "'Noto Sans Thai',sans-serif",
        color: "#262626",
      }}
    >
      {/* Toast */}
      {alertMsg && (
        <div
          style={{
            position: "fixed",
            top: 20,
            right: 20,
            zIndex: 2000,
            background: alertBg,
            color: "#fff",
            borderRadius: 10,
            padding: "12px 22px",
            fontWeight: 600,
            fontSize: 14,
            boxShadow: "0 4px 16px rgba(0,0,0,0.18)",
            animation: "fadein 0.3s",
            maxWidth: 360,
          }}
        >
          {alertMsg.type === "error"
            ? "❌ "
            : alertMsg.type === "warning"
            ? "⚠️ "
            : "✅ "}
          {alertMsg.msg}
        </div>
      )}

      {/* Header */}
      <div
        style={{
          background: "#fff",
          borderBottom: "1px solid #f0f0f0",
          padding: "0 32px",
          display: "flex",
          alignItems: "center",
          gap: 16,
          position: "sticky",
          top: 0,
          zIndex: 100,
          height: 58,
        }}
      >
        <span style={{ fontSize: 22 }}>📦</span>
        <span style={{ fontWeight: 800, fontSize: 18, letterSpacing: -0.5 }}>
          Warehouse Picking System
        </span>
        <div style={{ flex: 1 }} />
        <input
          type="file"
          accept=".json"
          id="restore-file"
          style={{ display: "none" }}
          onChange={restoreData}
        />
        <button style={bGr} onClick={backupData}>
          💾 Backup
        </button>
        <button
          style={bGr}
          onClick={() => document.getElementById("restore-file").click()}
        >
          🔄 Restore
        </button>
        <button style={bR} onClick={() => setConfirmClear(true)}>
          🗑️ ล้างข้อมูลวันนี้
        </button>
      </div>

      <div style={{ maxWidth: 1200, margin: "0 auto", padding: "24px 20px" }}>
        {/* Stats */}
        <div
          style={{
            display: "flex",
            gap: 16,
            marginBottom: 20,
            flexWrap: "wrap",
          }}
        >
          <StatCard
            icon="📋"
            value={orders.length}
            label="รายการทั้งหมด"
            color="#1677ff"
          />
          <StatCard
            icon="✅"
            value={completedOrders.length}
            label="เสร็จสิ้น"
            color="#52c41a"
          />
          <StatCard
            icon="⏳"
            value={pendingOrders.length}
            label="รอดำเนินการ"
            color="#fa8c16"
          />
          <StatCard
            icon="🎁"
            value={stock.reduce((a, s) => a + s.stock, 0)}
            label="สินค้าคงเหลือรวม"
            color="#722ed1"
          />
        </div>

        {/* Scanner */}
        <div style={card}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 12,
              marginBottom: 16,
            }}
          >
            <span style={{ fontWeight: 700, fontSize: 16 }}>
              📱 สแกนหยิบสินค้า (Mobile Scanner)
            </span>
            <button
              onClick={() => setAutoConfirm((v) => !v)}
              style={{
                ...bP,
                background: autoConfirm ? "#52c41a" : "#8c8c8c",
                fontSize: 12,
                padding: "4px 12px",
              }}
            >
              ⚡ Auto Confirm {autoConfirm ? "เปิดอยู่" : "ปิดอยู่"}
            </button>
          </div>

          {/* Info bar */}
          <div
            style={{
              background: "#fffbe6",
              border: "1px solid #ffe58f",
              borderRadius: 8,
              padding: "8px 14px",
              marginBottom: 16,
              fontSize: 13,
              color: "#614700",
            }}
          >
            💡 วิธีใช้: กรอก Pick No → สแกน/กรอก Barcode แล้วกด Enter —
            ระบบจะบันทึกการหยิบอัตโนมัติทันที
            <br />
            <span style={{ color: "#cf1322", fontWeight: 600 }}>
              ⚠️ ต้องมี Order ที่ตรงกันก่อนจึงจะหยิบสินค้าได้ — หยิบบางส่วนได้
              ระบบจะแจ้งเตือนยอดคงเหลือ
            </span>
          </div>

          <div
            style={{
              display: "flex",
              gap: 12,
              flexWrap: "wrap",
              alignItems: "flex-end",
            }}
          >
            <div style={{ flex: "1 1 180px" }}>
              <label style={lbl}>Pick No</label>
              <input
                style={inp}
                value={pickNo}
                onChange={(e) => setPickNo(e.target.value)}
                placeholder="สแกนหรือกรอก Pick No"
                onKeyDown={(e) =>
                  e.key === "Enter" && barcodeRef.current?.focus()
                }
              />
            </div>
            <div style={{ flex: "1 1 220px" }}>
              <label style={lbl}>
                Barcode{" "}
                <span style={{ color: "#8c8c8c", fontWeight: 400 }}>
                  (กด Enter เพื่อบันทึก)
                </span>
              </label>
              <input
                ref={barcodeRef}
                style={{
                  ...inp,
                  border: "2px solid #40a9ff",
                  background: "#f0f9ff",
                }}
                value={barcode}
                onChange={(e) => setBarcode(e.target.value)}
                placeholder="สแกนหรือกรอก Barcode ที่นี่"
                onKeyDown={handleBarcodeKeyDown}
              />
            </div>
            <div style={{ flex: "1 1 160px" }}>
              <label style={lbl}>
                Location{" "}
                <span style={{ color: "#8c8c8c", fontWeight: 400 }}>
                  (ระบุเพื่อแยก barcode ซ้ำ)
                </span>
              </label>
              <input
                style={inp}
                value={location}
                onChange={(e) => setLocation(e.target.value)}
                placeholder="เช่น Q14-028-40"
              />
            </div>
            <div style={{ flex: "0 0 100px" }}>
              <label style={lbl}>Qty</label>
              <input
                style={{ ...inp, textAlign: "center" }}
                type="number"
                min="1"
                value={qty}
                onChange={(e) => {
                  const v = parseInt(e.target.value, 10);
                  setQty(isNaN(v) || v < 1 ? 1 : v);
                }}
              />
            </div>
            <button style={{ ...bG, flexShrink: 0 }} onClick={confirmPick}>
              ✅ ยืนยันการหยิบสินค้า
            </button>
          </div>

          <div style={{ marginTop: 12, fontSize: 14 }}>
            สถานะ:{" "}
            <span style={{ color: scanStatusColor, fontWeight: 600 }}>
              {scanStatus}
            </span>
          </div>
        </div>

        {/* Order table */}
        <div style={card}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 12,
              marginBottom: 16,
              flexWrap: "wrap",
            }}
          >
            <span style={{ fontWeight: 700, fontSize: 16 }}>
              📋 รายการ Pick Order
            </span>
            <Badge color="pending">
              ⏳ PENDING {pendingOrders.length} รายการ
            </Badge>
            <div style={{ flex: 1, minWidth: 220 }}>
              <input
                style={{ ...inp, maxWidth: 320 }}
                value={searchOrder}
                onChange={(e) => setSearchOrder(e.target.value)}
                placeholder="🔍 ค้นหา Pick No, ลูกค้า, สินค้า, Barcode"
              />
            </div>
            <button style={bG} onClick={() => setShowCompleted((v) => !v)}>
              ✅ {showCompleted ? "ดูรายการค้าง" : "ดูรายการเสร็จแล้ว"}
              <span
                style={{
                  background: "#fff",
                  color: "#52c41a",
                  borderRadius: 99,
                  padding: "1px 7px",
                  fontSize: 12,
                  marginLeft: 6,
                }}
              >
                {showCompleted ? pendingOrders.length : completedOrders.length}
              </span>
            </button>
            <button style={bO} onClick={exportExcel}>
              ⬇️ Export Excel
            </button>
          </div>
          <div style={{ overflowX: "auto" }}>
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                fontSize: 13,
              }}
            >
              <thead>
                <tr
                  style={{
                    background: "#fafafa",
                    borderBottom: "1px solid #f0f0f0",
                  }}
                >
                  {[
                    "Pick No",
                    "Customer",
                    "Barcode",
                    "Product",
                    "Location",
                    "Required",
                    "Picked",
                    "Status",
                    "Alert",
                    "Created At",
                    "Completed At",
                    "ใช้เวลา",
                    "",
                  ].map((h) => (
                    <th
                      key={h}
                      style={{
                        padding: "10px 12px",
                        textAlign: "left",
                        fontWeight: 600,
                        color: "#595959",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredOrders.length === 0 ? (
                  <tr>
                    <td
                      colSpan={13}
                      style={{
                        textAlign: "center",
                        padding: 40,
                        color: "#8c8c8c",
                      }}
                    >
                      <div style={{ fontSize: 32, marginBottom: 8 }}>🎉</div>
                      ไม่มีรายการค้างดำเนินการ
                      <br />
                      <span style={{ fontSize: 12 }}>
                        กดปุ่ม "ดูรายการเสร็จแล้ว" เพื่อดูรายการที่เสร็จสิ้น
                      </span>
                    </td>
                  </tr>
                ) : (
                  filteredOrders.map((o) => {
                    const st = getOrderStatus(o),
                      picked = o.picked || 0;
                    return (
                      <tr
                        key={o.id}
                        style={{ borderBottom: "1px solid #f5f5f5" }}
                      >
                        <td
                          style={{
                            padding: "10px 12px",
                            fontWeight: 600,
                            color: "#1677ff",
                          }}
                        >
                          {o.pickNo}
                        </td>
                        <td style={{ padding: "10px 12px" }}>
                          {o.customer || "-"}
                        </td>
                        <td
                          style={{
                            padding: "10px 12px",
                            fontFamily: "monospace",
                          }}
                        >
                          {o.barcode}
                        </td>
                        <td style={{ padding: "10px 12px" }}>{o.product}</td>
                        <td style={{ padding: "10px 12px" }}>
                          {o.location || "-"}
                        </td>
                        <td
                          style={{ padding: "10px 12px", textAlign: "center" }}
                        >
                          {o.required}
                        </td>
                        <td
                          style={{
                            padding: "10px 12px",
                            textAlign: "center",
                            color: picked >= o.required ? "#52c41a" : "#fa8c16",
                            fontWeight: 600,
                          }}
                        >
                          {picked}
                        </td>
                        <td style={{ padding: "10px 12px" }}>
                          <Badge
                            color={
                              st === "complete"
                                ? "complete"
                                : st === "over3h"
                                ? "over3h"
                                : "pending"
                            }
                          >
                            {o.status}
                          </Badge>
                        </td>
                        <td style={{ padding: "10px 12px" }}>
                          {st === "over3h" ? (
                            <span style={{ color: "#ff4d4f", fontWeight: 600 }}>
                              ⚠️ เกิน 3ชม.
                            </span>
                          ) : (
                            "-"
                          )}
                        </td>
                        <td
                          style={{ padding: "10px 12px", whiteSpace: "nowrap" }}
                        >
                          {o.createdAt ? (
                            <div>
                              <div
                                style={{
                                  fontSize: 12,
                                  color: "#262626",
                                  fontWeight: 500,
                                }}
                              >
                                {o.createdAt.split(" ")[0]}
                              </div>
                              <div
                                style={{
                                  fontSize: 12,
                                  color: "#1677ff",
                                  fontWeight: 600,
                                }}
                              >
                                {o.createdAt.split(" ")[1]}
                              </div>
                            </div>
                          ) : (
                            "-"
                          )}
                        </td>
                        <td
                          style={{ padding: "10px 12px", whiteSpace: "nowrap" }}
                        >
                          {o.completedAt ? (
                            <div>
                              <div
                                style={{
                                  fontSize: 12,
                                  color: "#262626",
                                  fontWeight: 500,
                                }}
                              >
                                {o.completedAt.split(" ")[0]}
                              </div>
                              <div
                                style={{
                                  fontSize: 12,
                                  color: "#52c41a",
                                  fontWeight: 600,
                                }}
                              >
                                {o.completedAt.split(" ")[1]}
                              </div>
                            </div>
                          ) : (
                            <span style={{ color: "#d9d9d9", fontSize: 12 }}>
                              -
                            </span>
                          )}
                        </td>
                        <td
                          style={{ padding: "10px 12px", textAlign: "center" }}
                        >
                          {fmtDur(o) ?? (
                            <span style={{ color: "#d9d9d9", fontSize: 12 }}>
                              -
                            </span>
                          )}
                        </td>
                        <td style={{ padding: "10px 12px" }}>
                          <div style={{ display: "flex", gap: 6 }}>
                            <button
                              onClick={() => setEditOrderItem({ ...o })}
                              style={{
                                ...bGr,
                                padding: "4px 10px",
                                fontSize: 12,
                              }}
                            >
                              ✏️
                            </button>
                            <button
                              onClick={() => {
                                updateOrders((prev) =>
                                  prev.filter((x) => x.id !== o.id)
                                );
                                showAlert("ลบ Order แล้ว");
                              }}
                              style={{
                                ...bR,
                                padding: "4px 10px",
                                fontSize: 12,
                              }}
                            >
                              🗑️
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
          <div
            style={{
              marginTop: 12,
              fontSize: 12,
              color: "#8c8c8c",
              display: "flex",
              gap: 16,
              flexWrap: "wrap",
            }}
          >
            <Badge color="pending">PENDING = รอดำเนินการ</Badge>
            <Badge color="over3h">OVER 3H = ค้างเกิน 3 ชั่วโมง</Badge>
            <Badge color="complete">COMPLETE = เสร็จสิ้น</Badge>
          </div>
        </div>

        {/* Add Stock / Add Order */}
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: 20,
            marginBottom: 20,
          }}
        >
          {/* Add Stock */}
          <div style={card}>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                marginBottom: 16,
              }}
            >
              <div style={{ fontWeight: 700, fontSize: 16 }}>📦 Add Stock</div>
              <div style={{ display: "flex", gap: 8 }}>
                <button
                  style={{ ...bGr, fontSize: 12, padding: "5px 12px" }}
                  onClick={downloadStockTemplate}
                >
                  📄 Template
                </button>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  id="stock-xlsx"
                  style={{ display: "none" }}
                  onChange={handleStockXlsx}
                />
                <button
                  style={{ ...bG, fontSize: 12, padding: "5px 12px" }}
                  onClick={() => document.getElementById("stock-xlsx").click()}
                >
                  📤 Upload Excel
                </button>
              </div>
            </div>
            <div
              style={{
                border: "2px dashed #91d5ff",
                borderRadius: 8,
                padding: "10px 14px",
                marginBottom: 14,
                background: "#f0f9ff",
                fontSize: 12,
                color: "#0958d9",
                textAlign: "center",
                cursor: "pointer",
              }}
              onClick={() => document.getElementById("stock-xlsx").click()}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                const f = e.dataTransfer.files[0];
                if (f) parseStockXlsx(f);
              }}
            >
              📂 คลิกหรือลากไฟล์ Excel / CSV มาวางที่นี่
              <br />
              <span style={{ color: "#8c8c8c" }}>
                รองรับคอลัมน์: barcode, product, stock, location
              </span>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {[
                ["barcode", "Barcode"],
                ["product", "ชื่อสินค้า"],
                ["stock", "จำนวน", "number"],
                ["location", "Location"],
              ].map(([k, lb, tp]) => (
                <div key={k}>
                  <label style={lbl}>{lb}</label>
                  <input
                    style={inp}
                    type={tp || "text"}
                    value={addStockForm[k]}
                    onChange={(e) =>
                      setAddStockForm((f) => ({ ...f, [k]: e.target.value }))
                    }
                    placeholder={lb}
                  />
                </div>
              ))}
              <button
                style={{
                  ...bP,
                  marginTop: 4,
                  width: "100%",
                  justifyContent: "center",
                }}
                onClick={handleAddStock}
              >
                ➕ เพิ่มสต็อก
              </button>
            </div>
          </div>

          {/* Add Order */}
          <div style={card}>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                marginBottom: 16,
              }}
            >
              <div style={{ fontWeight: 700, fontSize: 16 }}>
                📋 Add Pick Order
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button
                  style={{ ...bGr, fontSize: 12, padding: "5px 12px" }}
                  onClick={downloadOrderTemplate}
                >
                  📄 Template
                </button>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  id="order-xlsx"
                  style={{ display: "none" }}
                  onChange={handleOrderXlsx}
                />
                <button
                  style={{ ...bO, fontSize: 12, padding: "5px 12px" }}
                  onClick={() => document.getElementById("order-xlsx").click()}
                >
                  📤 Upload Excel
                </button>
              </div>
            </div>
            <div
              style={{
                border: "2px dashed #ffd591",
                borderRadius: 8,
                padding: "10px 14px",
                marginBottom: 14,
                background: "#fffbe6",
                fontSize: 12,
                color: "#874d00",
                textAlign: "center",
                cursor: "pointer",
              }}
              onClick={() => document.getElementById("order-xlsx").click()}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                const f = e.dataTransfer.files[0];
                if (f) parseOrderXlsx(f);
              }}
            >
              📂 คลิกหรือลากไฟล์ Excel / CSV มาวางที่นี่
              <br />
              <span style={{ color: "#8c8c8c" }}>
                รองรับคอลัมน์: pickNo, customer, barcode, product, location,
                required
              </span>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              <div>
                <label style={lbl}>Pick No (ไม่ต้องใส่ก็ได้)</label>
                <input
                  style={inp}
                  value={addOrderForm.pickNo}
                  onChange={(e) =>
                    setAddOrderForm((f) => ({ ...f, pickNo: e.target.value }))
                  }
                  placeholder="Pick No"
                />
              </div>
              <div>
                <label style={lbl}>ชื่อลูกค้า</label>
                <input
                  style={inp}
                  value={addOrderForm.customer}
                  onChange={(e) =>
                    setAddOrderForm((f) => ({ ...f, customer: e.target.value }))
                  }
                  placeholder="ชื่อลูกค้า"
                />
              </div>
              <div>
                <label style={lbl}>Barcode *</label>
                <input
                  style={{
                    ...inp,
                    border:
                      addOrderForm.barcode &&
                      stock.find(
                        (s) => s.barcode === addOrderForm.barcode.trim()
                      )
                        ? "2px solid #52c41a"
                        : inp.border,
                  }}
                  value={addOrderForm.barcode}
                  onChange={(e) => {
                    const bc = e.target.value;
                    const found = stock.find((s) => s.barcode === bc.trim());
                    setAddOrderForm((f) => ({
                      ...f,
                      barcode: bc,
                      product: found ? found.product : f.product,
                      location: found ? found.location : f.location,
                    }));
                  }}
                  placeholder="Barcode *"
                />
                {addOrderForm.barcode &&
                  (() => {
                    const found = stock.find(
                      (s) => s.barcode === addOrderForm.barcode.trim()
                    );
                    return found ? (
                      <div
                        style={{
                          marginTop: 4,
                          fontSize: 12,
                          color: "#52c41a",
                          fontWeight: 600,
                        }}
                      >
                        ✅ พบสินค้า: {found.product} | Stock: {found.stock} |{" "}
                        {found.location}
                      </div>
                    ) : addOrderForm.barcode.length > 3 ? (
                      <div
                        style={{ marginTop: 4, fontSize: 12, color: "#fa8c16" }}
                      >
                        ⚠️ ไม่พบ Barcode นี้ใน Stock
                      </div>
                    ) : null;
                  })()}
              </div>
              <div>
                <label style={lbl}>ชื่อสินค้า *</label>
                <input
                  style={inp}
                  value={addOrderForm.product}
                  onChange={(e) =>
                    setAddOrderForm((f) => ({ ...f, product: e.target.value }))
                  }
                  placeholder="ชื่อสินค้า *"
                />
              </div>
              <div>
                <label style={lbl}>Location</label>
                <input
                  style={inp}
                  value={addOrderForm.location}
                  onChange={(e) =>
                    setAddOrderForm((f) => ({ ...f, location: e.target.value }))
                  }
                  placeholder="Location"
                />
              </div>
              <div>
                <label style={lbl}>จำนวนที่ต้องหยิบ</label>
                <input
                  style={inp}
                  type="number"
                  min="1"
                  value={addOrderForm.required}
                  onChange={(e) =>
                    setAddOrderForm((f) => ({ ...f, required: e.target.value }))
                  }
                  placeholder="จำนวน"
                />
              </div>
              <button
                style={{
                  ...bO,
                  marginTop: 4,
                  width: "100%",
                  justifyContent: "center",
                }}
                onClick={handleAddOrder}
              >
                ➕ เพิ่ม Order
              </button>
            </div>
          </div>
        </div>

        {/* Stock table */}
        <div style={card}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 12,
              marginBottom: 16,
            }}
          >
            <span style={{ fontWeight: 700, fontSize: 16 }}>
              📦 Stock คงเหลือ
            </span>
            <div style={{ flex: 1 }}>
              <input
                style={{ ...inp, maxWidth: 280 }}
                value={stockSearch}
                onChange={(e) => setStockSearch(e.target.value)}
                placeholder="ค้นหา Stock..."
              />
            </div>
            <button
              style={bR}
              onClick={() => {
                if (window.confirm("ลบ Stock ทั้งหมด?")) {
                  updateStock([]);
                  showAlert("ลบ Stock ทั้งหมดแล้ว");
                }
              }}
            >
              🗑️ ลบทั้งหมด
            </button>
          </div>
          <div style={{ overflowX: "auto" }}>
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                fontSize: 14,
              }}
            >
              <thead>
                <tr
                  style={{
                    background: "#fafafa",
                    borderBottom: "1px solid #f0f0f0",
                  }}
                >
                  {["Barcode", "Product", "Stock", "Location", "แก้ไข"].map(
                    (h) => (
                      <th
                        key={h}
                        style={{
                          padding: "10px 14px",
                          textAlign: "left",
                          fontWeight: 600,
                          color: "#595959",
                        }}
                      >
                        {h}
                      </th>
                    )
                  )}
                </tr>
              </thead>
              <tbody>
                {filteredStock.map((s) => (
                  <tr
                    key={s.barcode}
                    style={{ borderBottom: "1px solid #f5f5f5" }}
                  >
                    <td
                      style={{ padding: "10px 14px", fontFamily: "monospace" }}
                    >
                      {s.barcode}
                    </td>
                    <td style={{ padding: "10px 14px" }}>{s.product}</td>
                    <td
                      style={{
                        padding: "10px 14px",
                        color: s.stock > 0 ? "#52c41a" : "#ff4d4f",
                        fontWeight: 700,
                      }}
                    >
                      {s.stock}
                    </td>
                    <td style={{ padding: "10px 14px" }}>{s.location}</td>
                    <td style={{ padding: "10px 14px" }}>
                      <div style={{ display: "flex", gap: 8 }}>
                        <button
                          onClick={() =>
                            setEditStockItem({
                              ...s,
                              _originalBarcode: s.barcode,
                            })
                          }
                          style={{ ...bP, padding: "5px 14px", fontSize: 13 }}
                        >
                          ✏️ แก้ไข
                        </button>
                        <button
                          onClick={() => {
                            updateStock((prev) =>
                              prev.filter((x) => x.barcode !== s.barcode)
                            );
                            showAlert("ลบสินค้าแล้ว");
                          }}
                          style={{ ...bR, padding: "5px 14px", fontSize: 13 }}
                        >
                          🗑️ ลบ
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
                {filteredStock.length === 0 && (
                  <tr>
                    <td
                      colSpan={5}
                      style={{
                        textAlign: "center",
                        padding: 30,
                        color: "#8c8c8c",
                      }}
                    >
                      ไม่พบสินค้า
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Edit Stock Modal */}
      <Modal
        open={!!editStockItem}
        onClose={() => setEditStockItem(null)}
        title="แก้ไขสต็อก"
      >
        {editStockItem && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            {[
              ["barcode", "Barcode"],
              ["product", "ชื่อสินค้า"],
              ["stock", "จำนวน", "number"],
              ["location", "Location"],
            ].map(([k, lb, tp]) => (
              <div key={k}>
                <label style={lbl}>{lb}</label>
                <input
                  style={inp}
                  type={tp || "text"}
                  value={editStockItem[k]}
                  onChange={(e) =>
                    setEditStockItem((f) => ({
                      ...f,
                      [k]:
                        tp === "number"
                          ? Number(e.target.value)
                          : e.target.value,
                    }))
                  }
                />
              </div>
            ))}
            <div style={{ display: "flex", gap: 10, marginTop: 8 }}>
              <button style={{ ...bP, flex: 1 }} onClick={handleEditStock}>
                💾 บันทึก
              </button>
              <button
                style={{ ...bGr, flex: 1 }}
                onClick={() => setEditStockItem(null)}
              >
                ยกเลิก
              </button>
            </div>
          </div>
        )}
      </Modal>

      {/* Edit Order Modal */}
      <Modal
        open={!!editOrderItem}
        onClose={() => setEditOrderItem(null)}
        title="แก้ไข Order"
      >
        {editOrderItem && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            {[
              ["pickNo", "Pick No"],
              ["customer", "ลูกค้า"],
              ["barcode", "Barcode"],
              ["product", "สินค้า"],
              ["location", "Location"],
              ["required", "จำนวนที่ต้องหยิบ", "number"],
              ["picked", "หยิบแล้ว", "number"],
            ].map(([k, lb, tp]) => (
              <div key={k}>
                <label style={lbl}>{lb}</label>
                <input
                  style={inp}
                  type={tp || "text"}
                  value={editOrderItem[k] ?? ""}
                  onChange={(e) =>
                    setEditOrderItem((f) => ({
                      ...f,
                      [k]:
                        tp === "number"
                          ? Number(e.target.value)
                          : e.target.value,
                    }))
                  }
                />
              </div>
            ))}
            <div>
              <label style={lbl}>Status</label>
              <select
                style={inp}
                value={editOrderItem.status}
                onChange={(e) =>
                  setEditOrderItem((f) => ({ ...f, status: e.target.value }))
                }
              >
                <option>PENDING</option>
                <option>COMPLETE</option>
              </select>
            </div>
            {editOrderItem.status === "COMPLETE" &&
              !editOrderItem.completedAt && (
                <div
                  style={{
                    background: "#f6ffed",
                    border: "1px solid #b7eb8f",
                    borderRadius: 8,
                    padding: "8px 12px",
                    fontSize: 12,
                    color: "#389e0d",
                  }}
                >
                  ✅ ระบบจะบันทึกเวลา Completed At อัตโนมัติเมื่อกดบันทึก
                </div>
              )}
            <div style={{ display: "flex", gap: 10, marginTop: 8 }}>
              <button style={{ ...bP, flex: 1 }} onClick={handleEditOrder}>
                💾 บันทึก
              </button>
              <button
                style={{ ...bGr, flex: 1 }}
                onClick={() => setEditOrderItem(null)}
              >
                ยกเลิก
              </button>
            </div>
          </div>
        )}
      </Modal>

      {/* Excel Preview Modal */}
      <Modal
        open={!!xlsxPreview}
        onClose={() => setXlsxPreview(null)}
        title={
          xlsxPreview?.type === "stock"
            ? "📤 ตรวจสอบข้อมูล Stock ก่อนนำเข้า"
            : "📤 ตรวจสอบข้อมูล Order ก่อนนำเข้า"
        }
      >
        {xlsxPreview && (
          <div>
            {xlsxPreview.errors.length > 0 && (
              <div
                style={{
                  background: "#fff1f0",
                  border: "1px solid #ffa39e",
                  borderRadius: 8,
                  padding: "10px 14px",
                  marginBottom: 14,
                  fontSize: 13,
                  color: "#a8071a",
                }}
              >
                ⚠️ พบข้อผิดพลาด {xlsxPreview.errors.length} รายการ:
                <br />
                {xlsxPreview.errors.slice(0, 5).map((e, i) => (
                  <div key={i}>• {e}</div>
                ))}
                {xlsxPreview.errors.length > 5 && (
                  <div>...และอีก {xlsxPreview.errors.length - 5} รายการ</div>
                )}
              </div>
            )}
            <div
              style={{
                background: "#f6ffed",
                border: "1px solid #b7eb8f",
                borderRadius: 8,
                padding: "10px 14px",
                marginBottom: 14,
                fontSize: 13,
                color: "#237804",
              }}
            >
              ✅ พบข้อมูลที่นำเข้าได้ <strong>{xlsxPreview.rows.length}</strong>{" "}
              แถว
            </div>
            <div
              style={{
                maxHeight: 240,
                overflowY: "auto",
                border: "1px solid #f0f0f0",
                borderRadius: 8,
                marginBottom: 16,
              }}
            >
              <table
                style={{
                  width: "100%",
                  borderCollapse: "collapse",
                  fontSize: 12,
                }}
              >
                <thead>
                  <tr
                    style={{
                      background: "#fafafa",
                      position: "sticky",
                      top: 0,
                    }}
                  >
                    {(xlsxPreview.type === "stock"
                      ? ["Barcode", "Product", "Stock", "Location"]
                      : [
                          "Pick No",
                          "Customer",
                          "Barcode",
                          "Product",
                          "Location",
                          "Required",
                        ]
                    ).map((h) => (
                      <th
                        key={h}
                        style={{
                          padding: "7px 10px",
                          textAlign: "left",
                          fontWeight: 600,
                          color: "#595959",
                        }}
                      >
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {xlsxPreview.rows.slice(0, 50).map((r, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #f5f5f5" }}>
                      {(xlsxPreview.type === "stock"
                        ? [r.barcode, r.product, r.stock, r.location]
                        : [
                            r.pickNo,
                            r.customer,
                            r.barcode,
                            r.product,
                            r.location,
                            r.required,
                          ]
                      ).map((v, j) => (
                        <td key={j} style={{ padding: "6px 10px" }}>
                          {v}
                        </td>
                      ))}
                    </tr>
                  ))}
                  {xlsxPreview.rows.length > 50 && (
                    <tr>
                      <td
                        colSpan={6}
                        style={{
                          padding: "8px 10px",
                          color: "#8c8c8c",
                          fontSize: 12,
                          textAlign: "center",
                        }}
                      >
                        ...แสดง 50 แถวแรก จากทั้งหมด {xlsxPreview.rows.length}{" "}
                        แถว
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button
                style={{ ...bG, flex: 1, justifyContent: "center" }}
                onClick={confirmXlsxImport}
              >
                ✅ ยืนยันนำเข้า {xlsxPreview.rows.length} รายการ
              </button>
              <button
                style={{ ...bGr, flex: 1, justifyContent: "center" }}
                onClick={() => setXlsxPreview(null)}
              >
                ยกเลิก
              </button>
            </div>
          </div>
        )}
      </Modal>

      {/* Confirm Clear Modal */}
      <Modal
        open={confirmClear}
        onClose={() => setConfirmClear(false)}
        title="⚠️ ล้างข้อมูลวันนี้"
      >
        <p style={{ color: "#595959", marginBottom: 20 }}>
          ต้องการล้างข้อมูล Orders วันนี้ทั้งหมดหรือไม่? (Stock จะไม่ถูกลบ)
        </p>
        <div style={{ display: "flex", gap: 10 }}>
          <button
            style={{ ...bR, flex: 1 }}
            onClick={() => {
              const today = new Date().toLocaleDateString("th-TH");
              updateOrders((prev) =>
                prev.filter(
                  (o) =>
                    new Date(o.createdAtRaw).toLocaleDateString("th-TH") !==
                    today
                )
              );
              setConfirmClear(false);
              showAlert("ล้างข้อมูลวันนี้แล้ว");
            }}
          >
            ยืนยันล้างข้อมูล
          </button>
          <button
            style={{ ...bGr, flex: 1 }}
            onClick={() => setConfirmClear(false)}
          >
            ยกเลิก
          </button>
        </div>
      </Modal>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@400;600;700;800&display=swap');
        * { box-sizing:border-box; }
        input:focus,select:focus { border-color:#40a9ff !important; outline:none; box-shadow:0 0 0 2px rgba(24,144,255,0.15); }
        @keyframes fadein { from{opacity:0;transform:translateY(-8px)} to{opacity:1;transform:translateY(0)} }
      `}</style>
    </div>
  );
}
