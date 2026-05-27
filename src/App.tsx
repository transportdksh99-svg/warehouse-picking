import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getDatabase, ref, onValue, set } from "firebase/database";

// ── Firebase Config ───────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey: "AIzaSyDbxbu1w9VSapEWT2ApUvvPBz5xUdj29vU",
  authDomain: "warehouse-app-52001.firebaseapp.com",
  databaseURL: "https://warehouse-app-52001-default-rtdb.firebaseio.com",
  projectId: "warehouse-app-52001",
  storageBucket: "warehouse-app-52001.firebasestorage.app",
  messagingSenderId: "72442819990",
  appId: "1:72442819990:web:b349bfe2bb1e08679259b7",
};
const firebaseApp = initializeApp(firebaseConfig);
const db = getDatabase(firebaseApp);

// ── Utility ──────────────────────────────────────────────────────────────────
const now = () => new Date().toLocaleString("th-TH", { hour12: false });
const fmtDate = (d) =>
  d
    ? new Date(d).toLocaleString("th-TH", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
    : "-";

const INIT_STOCK = [
  {
    barcode: "8851234567890",
    product: "สินค้า A",
    stock: 120,
    location: "A01-01",
  },
  {
    barcode: "8851234567891",
    product: "สินค้า B",
    stock: 80,
    location: "A01-02",
  },
  {
    barcode: "8851234567892",
    product: "สินค้า C",
    stock: 60,
    location: "A01-03",
  },
  {
    barcode: "8851234567893",
    product: "สินค้า D",
    stock: 40,
    location: "A02-01",
  },
  {
    barcode: "8851234567894",
    product: "สินค้า E",
    stock: 25,
    location: "A02-02",
  },
];

const INIT_ORDERS = [
  {
    id: "PO-240515-001",
    customer: "บริษัท เอ จำกัด",
    barcode: "8851234567890",
    product: "สินค้า A",
    location: "A01-01",
    required: 10,
    picked: 10,
    status: "COMPLETE",
    createdAt: "2024-05-15T09:15:00",
    completedAt: "2024-05-15T10:05:00",
    autoTime: "50 นาที",
  },
  {
    id: "PO-240515-002",
    customer: "บริษัท บี จำกัด",
    barcode: "8851234567891",
    product: "สินค้า B",
    location: "A01-02",
    required: 20,
    picked: 12,
    status: "PENDING",
    createdAt: "2024-05-15T08:00:00",
    completedAt: null,
    autoTime: null,
  },
  {
    id: "PO-240515-003",
    customer: "บริษัท ซี จำกัด",
    barcode: "8851234567892",
    product: "สินค้า C",
    location: "A01-03",
    required: 5,
    picked: 2,
    status: "PENDING",
    createdAt: "2024-05-15T10:30:00",
    completedAt: null,
    autoTime: null,
  },
  {
    id: "PO-240515-004",
    customer: "บริษัท ดี จำกัด",
    barcode: "8851234567893",
    product: "สินค้า D",
    location: "A02-01",
    required: 15,
    picked: 15,
    status: "COMPLETE",
    createdAt: "2024-05-15T11:00:00",
    completedAt: "2024-05-15T11:40:00",
    autoTime: "40 นาที",
  },
  {
    id: "PO-240515-005",
    customer: "บริษัท อี จำกัด",
    barcode: "8851234567894",
    product: "สินค้า E",
    location: "A02-02",
    required: 8,
    picked: 0,
    status: "PENDING",
    createdAt: "2024-05-15T11:20:00",
    completedAt: null,
    autoTime: null,
  },
];

const alertLabel = (order) => {
  if (order.status === "COMPLETE") return { text: "DONE", cls: "done" };
  const mins = (Date.now() - new Date(order.createdAt)) / 60000;
  if (mins > 180) return { text: "OVER 3 HOURS", cls: "over3" };
  return { text: "NORMAL", cls: "normal" };
};

// ── Styles ────────────────────────────────────────────────────────────────────
const S = {
  app: {
    fontFamily: "'Sarabun', 'Noto Sans Thai', sans-serif",
    background: "#f4f6fa",
    minHeight: "100vh",
    fontSize: 14,
    color: "#1a2233",
  },
  header: {
    background: "#fff",
    borderBottom: "1px solid #e5e9f0",
    padding: "0 28px",
    height: 56,
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    boxShadow: "0 1px 4px rgba(0,0,0,.06)",
    position: "sticky",
    top: 0,
    zIndex: 100,
  },
  logo: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    fontWeight: 700,
    fontSize: 18,
    color: "#1a2233",
  },
  btnGroup: { display: "flex", gap: 8 },
  body: { padding: "20px 28px", maxWidth: 1300, margin: "0 auto" },
  statsRow: {
    display: "grid",
    gridTemplateColumns: "repeat(4,1fr)",
    gap: 16,
    marginBottom: 20,
  },
  statCard: {
    background: "#fff",
    borderRadius: 12,
    padding: "18px 20px",
    display: "flex",
    alignItems: "center",
    gap: 16,
    boxShadow: "0 1px 4px rgba(0,0,0,.06)",
  },
  statIcon: {
    width: 48,
    height: 48,
    borderRadius: 10,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: 22,
  },
  statNum: { fontSize: 28, fontWeight: 800, lineHeight: 1.1 },
  statLabel: { fontSize: 12, color: "#7a8599", marginTop: 2 },
  card: {
    background: "#fff",
    borderRadius: 12,
    padding: "20px 22px",
    boxShadow: "0 1px 4px rgba(0,0,0,.06)",
    marginBottom: 18,
  },
  sectionTitle: {
    fontWeight: 700,
    fontSize: 16,
    display: "flex",
    alignItems: "center",
    gap: 8,
    marginBottom: 16,
  },
  row3: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr auto auto",
    gap: 10,
    alignItems: "end",
    marginBottom: 8,
  },
  input: {
    border: "1px solid #d9dde8",
    borderRadius: 7,
    padding: "8px 12px",
    fontSize: 14,
    outline: "none",
    width: "100%",
    boxSizing: "border-box",
  },
  btn: {
    borderRadius: 7,
    border: "none",
    cursor: "pointer",
    padding: "9px 18px",
    fontWeight: 600,
    fontSize: 13,
    display: "flex",
    alignItems: "center",
    gap: 6,
  },
  btnBlue: { background: "#2563eb", color: "#fff" },
  btnGreen: { background: "#16a34a", color: "#fff" },
  btnRed: { background: "#ef4444", color: "#fff" },
  btnGray: {
    background: "#f1f3f8",
    color: "#4a5568",
    border: "1px solid #d9dde8",
  },
  btnExport: { background: "#f59e0b", color: "#fff" },
  twoCol: {
    display: "grid",
    gridTemplateColumns: "1fr auto",
    gap: 16,
    marginBottom: 18,
  },
  twoColMain: {
    display: "grid",
    gridTemplateColumns: "1fr 340px",
    gap: 18,
    alignItems: "start",
  },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th: {
    textAlign: "left",
    padding: "8px 10px",
    borderBottom: "2px solid #e5e9f0",
    color: "#7a8599",
    fontWeight: 600,
    whiteSpace: "nowrap",
  },
  td: {
    padding: "9px 10px",
    borderBottom: "1px solid #f0f2f7",
    verticalAlign: "middle",
  },
  badge: {
    display: "inline-block",
    padding: "2px 10px",
    borderRadius: 20,
    fontSize: 12,
    fontWeight: 700,
  },
  bottomRow: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18 },
  formGrid2: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 10,
    marginBottom: 10,
  },
  formGrid4: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr 1fr",
    gap: 10,
    marginBottom: 10,
  },
  footer: {
    background: "#1a2233",
    color: "#9aa5bb",
    padding: "14px 28px",
    display: "flex",
    justifyContent: "space-between",
    fontSize: 12,
    marginTop: 30,
  },
  tag: {
    display: "inline-block",
    padding: "1px 8px",
    borderRadius: 4,
    fontSize: 12,
    fontWeight: 700,
  },
  statusTag: {
    COMPLETE: { background: "#dcfce7", color: "#16a34a" },
    PENDING: { background: "#fef9c3", color: "#b45309" },
  },
  alertTag: {
    done: { background: "#dbeafe", color: "#1d4ed8" },
    over3: { background: "#fee2e2", color: "#dc2626" },
    normal: { background: "#f1f5f9", color: "#475569" },
  },
  toast: {
    position: "fixed",
    bottom: 30,
    right: 30,
    background: "#1a2233",
    color: "#fff",
    padding: "12px 20px",
    borderRadius: 10,
    zIndex: 9999,
    fontSize: 14,
    boxShadow: "0 4px 20px rgba(0,0,0,.2)",
  },
};

// ── Toast ─────────────────────────────────────────────────────────────────────
function Toast({ msg, onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 2500);
    return () => clearTimeout(t);
  }, [onClose]);
  return <div style={S.toast}>{msg}</div>;
}

// ── Firebase helpers ─────────────────────────────────────────────────────────
const saveDB = (path, val) => set(ref(db, path), val).catch(console.error);

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [stock, setStock] = useState(INIT_STOCK);
  const [orders, setOrders] = useState(INIT_ORDERS);
  const [scanPickNo, setScanPickNo] = useState("");
  const [dbReady, setDbReady] = useState(false);
  const [scanBarcode, setScanBarcode] = useState("");
  const [scanLocation, setScanLocation] = useState("");
  const [scanQty, setScanQty] = useState(1);
  const [scanStatus, setScanStatus] = useState("พร้อมสแกน");
  const [orderSearch, setOrderSearch] = useState("");
  const [stockSearch, setStockSearch] = useState("");
  const [newStock, setNewStock] = useState({
    barcode: "",
    product: "",
    stock: "",
    location: "",
  });
  const [newOrder, setNewOrder] = useState({
    id: "",
    customer: "",
    barcode: "",
    qty: "",
    location: "",
  });
  const [editStockIdx, setEditStockIdx] = useState(null);
  const [editStockData, setEditStockData] = useState({});
  const [toast, setToast] = useState(null);
  const [cameraOpen, setCameraOpen] = useState(false);
  const [cameraError, setCameraError] = useState("");
  const [scanning, setScanning] = useState(false);
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const animFrameRef = useRef(null);
  const canvasRef = useRef(null);
  const scanBarcodeInputRef = useRef(null);
  const isEditingStock = editStockIdx !== null;

  const showToast = (msg) => setToast(msg);

  // ── Firebase: โหลดข้อมูลจาก Firebase เมื่อเริ่มต้น (real-time sync)
  useEffect(() => {
    // ฟัง stock
    const unsubStock = onValue(ref(db, "stock"), (snap) => {
      if (snap.exists()) setStock(snap.val());
      setDbReady(true);
    });
    // ฟัง orders
    const unsubOrders = onValue(ref(db, "orders"), (snap) => {
      if (snap.exists()) setOrders(snap.val());
    });
    return () => { unsubStock(); unsubOrders(); };
  }, []);

  // ── Firebase: บันทึก stock ทุกครั้งที่เปลี่ยน (หลัง DB ready)
  useEffect(() => { if (dbReady) saveDB("stock", stock); }, [stock]);
  useEffect(() => { if (dbReady) saveDB("orders", orders); }, [orders]);

  // Focus กลับที่ช่อง Barcode เมื่อออกจาก edit mode
  useEffect(() => {
    if (editStockIdx === null && scanBarcodeInputRef.current) {
      scanBarcodeInputRef.current.focus();
    }
  }, [editStockIdx]);

  // ── Camera / BarcodeDetector ──────────────────────────────────────────────
  const stopCamera = () => {
    if (animFrameRef.current) cancelAnimationFrame(animFrameRef.current);
    if (streamRef.current)
      streamRef.current.getTracks().forEach((t) => t.stop());
    streamRef.current = null;
    setScanning(false);
  };

  const closeCamera = () => {
    stopCamera();
    setCameraOpen(false);
    setCameraError("");
  };

  const openCamera = async () => {
    setCameraError("");
    setCameraOpen(true);
    setScanning(true);
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: "environment" },
      });
      streamRef.current = stream;
      setTimeout(() => {
        if (videoRef.current) {
          videoRef.current.srcObject = stream;
          videoRef.current.play();
          startDetect();
        }
      }, 300);
    } catch (err) {
      setCameraError("❌ ไม่สามารถเข้าถึงกล้องได้: " + err.message);
      setScanning(false);
    }
  };

  const startDetect = () => {
    if ("BarcodeDetector" in window) {
      const detector = new window.BarcodeDetector({
        formats: [
          "ean_13", "ean_8", "code_128", "code_39",
          "qr_code", "upc_a", "upc_e",
        ],
      });
      const tick = async () => {
        if (!videoRef.current || videoRef.current.readyState < 2) {
          animFrameRef.current = requestAnimationFrame(tick);
          return;
        }
        try {
          const codes = await detector.detect(videoRef.current);
          if (codes.length > 0) {
            const val = codes[0].rawValue;
            setScanBarcode(val);
            setScanStatus(`📷 สแกนได้: ${val}`);
            showToast(`✅ สแกน Barcode: ${val}`);
            closeCamera();
            return;
          }
        } catch (_) {}
        animFrameRef.current = requestAnimationFrame(tick);
      };
      animFrameRef.current = requestAnimationFrame(tick);
    } else {
      setCameraError(
        "⚠️ browser นี้ไม่รองรับ BarcodeDetector — กรุณากรอก Barcode ด้านล่าง หรือใช้ Chrome บน Android/Desktop"
      );
      setScanning(false);
    }
  };

  // Stats
  const totalOrders = orders.length;
  const completed = orders.filter((o) => o.status === "COMPLETE").length;
  const pending = orders.filter((o) => o.status === "PENDING").length;
  const totalStock = stock.reduce((a, s) => a + Number(s.stock), 0);

  // ── Scanner confirm ───────────────────────────────────────────────────────
  const handleScanConfirm = () => {
    if (!scanPickNo || !scanBarcode) {
      setScanStatus("⚠️ กรุณากรอก Pick No และ Barcode");
      return;
    }
    // จับคู่ด้วย id + barcode + location (ถ้ากรอก location ให้ตรงด้วย)
    const orderIdx = orders.findIndex((o) => {
      const matchBase = o.id === scanPickNo && o.barcode === scanBarcode;
      if (!scanLocation.trim()) return matchBase && o.status !== "COMPLETE";
      return matchBase && o.location.trim() === scanLocation.trim();
    });
    if (orderIdx === -1) {
      // ถ้าหาไม่เจอ ลองดูว่ามี order ที่ barcode+id ตรงแต่ complete หมดแล้วหรือเปล่า
      const allDone = orders.filter(
        (o) => o.id === scanPickNo && o.barcode === scanBarcode
      ).every((o) => o.status === "COMPLETE");
      if (allDone && orders.some((o) => o.id === scanPickNo && o.barcode === scanBarcode)) {
        setScanStatus("✅ Order นี้เสร็จสิ้นแล้ว");
      } else if (scanLocation.trim()) {
        setScanStatus(`❌ ไม่พบ Order: Pick No=${scanPickNo}, Barcode=${scanBarcode}, Location=${scanLocation}`);
      } else {
        setScanStatus("❌ ไม่พบ Order ที่ตรงกัน");
      }
      return;
    }
    const order = orders[orderIdx];
    if (order.status === "COMPLETE") {
      setScanStatus("✅ Order นี้เสร็จสิ้นแล้ว");
      return;
    }
    // จับคู่ stock ด้วย barcode + location ของ order
    const stockIdx = stock.findIndex(
      (s) => s.barcode === scanBarcode && s.location.trim() === order.location.trim()
    ) !== -1
      ? stock.findIndex((s) => s.barcode === scanBarcode && s.location.trim() === order.location.trim())
      : stock.findIndex((s) => s.barcode === scanBarcode);
    if (stockIdx === -1) {
      setScanStatus("❌ ไม่พบสินค้าใน Stock");
      return;
    }
    const qty = Number(scanQty);
    if (qty <= 0) {
      setScanStatus("⚠️ จำนวนต้องมากกว่า 0");
      return;
    }

    const remaining = order.required - order.picked;
    if (qty > remaining) {
      setScanStatus(
        `❌ หยิบเกิน! ต้องการอีกแค่ ${remaining} ชิ้น แต่กรอก ${qty} ชิ้น`
      );
      showToast(`❌ หยิบเกิน! Order นี้ต้องการอีกแค่ ${remaining} ชิ้น`);
      return;
    }

    const currentStock = stock[stockIdx].stock;
    if (qty > currentStock) {
      setScanStatus(
        `❌ Stock ไม่พอ! มีสินค้าคงเหลือ ${currentStock} ชิ้น แต่ต้องการ ${qty} ชิ้น`
      );
      showToast(`❌ Stock ไม่พอ! คงเหลือ ${currentStock} ชิ้น`);
      return;
    }

    const newPicked = order.picked + qty;
    const newOrders = [...orders];
    const newStockArr = [...stock];
    newStockArr[stockIdx] = {
      ...newStockArr[stockIdx],
      stock: newStockArr[stockIdx].stock - qty,
    };
    newOrders[orderIdx] = {
      ...order,
      picked: newPicked,
      status: newPicked >= order.required ? "COMPLETE" : "PENDING",
      completedAt: newPicked >= order.required ? new Date().toISOString() : null,
      autoTime:
        newPicked >= order.required
          ? `${Math.round((Date.now() - new Date(order.createdAt)) / 60000)} นาที`
          : null,
    };
    setOrders(newOrders);
    setStock(newStockArr);
    setScanStatus(`✅ หยิบ ${qty} ชิ้น สำเร็จ! (${newPicked}/${order.required})`);
    setScanBarcode("");
    setScanLocation("");
    setScanQty(1);
    showToast(`หยิบสินค้าสำเร็จ: ${order.product} x${qty} [${order.location}]`);

    // ✅ AUTO-FOCUS กลับที่ช่อง Barcode เพื่อสแกนรายการถัดไปทันที
    setTimeout(() => scanBarcodeInputRef.current?.focus(), 50);
  };

  // Add stock
  const handleAddStock = () => {
    if (!newStock.barcode || !newStock.product || !newStock.stock || !newStock.location) {
      showToast("⚠️ กรุณากรอกข้อมูลให้ครบ");
      return;
    }
    const idx = stock.findIndex((s) => s.barcode === newStock.barcode);
    if (idx !== -1) {
      const updated = [...stock];
      updated[idx] = { ...updated[idx], stock: updated[idx].stock + Number(newStock.stock) };
      setStock(updated);
      showToast("อัปเดต Stock สำเร็จ");
    } else {
      setStock([...stock, { ...newStock, stock: Number(newStock.stock) }]);
      showToast("เพิ่ม Stock สำเร็จ");
    }
    setNewStock({ barcode: "", product: "", stock: "", location: "" });
  };

  // Add order
  const handleAddOrder = () => {
    if (!newOrder.id || !newOrder.customer || !newOrder.barcode || !newOrder.qty) {
      showToast("⚠️ กรุณากรอกข้อมูลให้ครบ");
      return;
    }
    const stockItem = stock.find((s) => s.barcode === newOrder.barcode);
    if (!stockItem) {
      showToast("❌ ไม่พบ Barcode ใน Stock");
      return;
    }
    setOrders([
      ...orders,
      {
        id: newOrder.id,
        customer: newOrder.customer,
        barcode: newOrder.barcode,
        product: stockItem.product,
        location: newOrder.location.trim() || stockItem.location,
        required: Number(newOrder.qty),
        picked: 0,
        status: "PENDING",
        createdAt: new Date().toISOString(),
        completedAt: null,
        autoTime: null,
      },
    ]);
    showToast("เพิ่ม Pick Order สำเร็จ");
    setNewOrder({ id: "", customer: "", barcode: "", qty: "", location: "" });
  };

  // Edit stock inline
  const startEditStock = (idx) => {
    if (scanBarcodeInputRef.current) scanBarcodeInputRef.current.blur();
    setEditStockIdx(idx);
    setEditStockData({ ...stock[idx] });
  };
  const saveEditStock = () => {
    const updated = [...stock];
    updated[editStockIdx] = { ...editStockData, stock: Number(editStockData.stock) };
    setStock(updated);
    setEditStockIdx(null);
    showToast("แก้ไข Stock สำเร็จ");
  };

  const deleteStock = (idx) => {
    setStock(stock.filter((_, i) => i !== idx));
    showToast("ลบ Stock สำเร็จ");
  };

  const deleteOrder = (idx) => {
    setOrders(orders.filter((_, i) => i !== idx));
    showToast("ลบ Order สำเร็จ");
  };

  // Export CSV
  const handleExport = () => {
    const rows = orders.map((o) => ({
      "Pick No": "'" + String(o.id),
      "Customer": String(o.customer || ""),
      "Barcode": "'" + String(o.barcode),
      "Product": String(o.product || ""),
      "Location": String(o.location || ""),
      "Required": o.required,
      "Picked": o.picked,
      "Status": o.status,
      "Created At": fmtDate(o.createdAt),
      "Completed At": fmtDate(o.completedAt),
      "Auto Time": o.autoTime || "-",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [
      { wch: 20 }, { wch: 45 }, { wch: 18 }, { wch: 32 },
      { wch: 16 }, { wch: 10 }, { wch: 10 }, { wch: 12 },
      { wch: 22 }, { wch: 22 }, { wch: 12 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pick Orders");
    // เขียนเป็น csv ด้วย BOM เพื่อรองรับภาษาไทย
    const csvData = XLSX.utils.sheet_to_csv(ws);
    const bom = "\uFEFF";
    const blob = new Blob([bom + csvData], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "pick_orders.csv";
    a.click();
    URL.revokeObjectURL(url);
  };

  // Backup/Restore
  const handleBackup = () => {
    const data = JSON.stringify({ stock, orders });
    const blob = new Blob([data], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "warehouse_backup.json";
    a.click();
    showToast("Backup สำเร็จ");
  };
  const restoreRef = useRef();
  const handleRestore = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const r = new FileReader();
    r.onload = (ev) => {
      try {
        const d = JSON.parse(ev.target.result);
        setStock(d.stock || []);
        setOrders(d.orders || []);
        showToast("Restore สำเร็จ");
      } catch {
        showToast("❌ ไฟล์ไม่ถูกต้อง");
      }
    };
    r.readAsText(file);
  };
  const handleClearAll = () => {
    if (window.confirm("ยืนยันลบข้อมูลทั้งหมด?")) {
      setStock([]);
      setOrders([]);
      saveDB("stock", []);
      saveDB("orders", []);
      showToast("🗑️ ลบข้อมูลทั้งหมดแล้ว");
    }
  };

  const filteredOrders = orders.filter(
    (o) =>
      o.id.includes(orderSearch) ||
      o.customer.includes(orderSearch) ||
      o.product.includes(orderSearch) ||
      o.barcode.includes(orderSearch)
  );
  const filteredStock = stock.filter(
    (s) =>
      s.barcode.includes(stockSearch) ||
      s.product.includes(stockSearch) ||
      s.location.includes(stockSearch)
  );

  return (
    <div style={S.app}>
      {/* HEADER */}
      <header style={S.header}>
        <div style={S.logo}>
          <span style={{ fontSize: 22 }}>📦</span>
          <span>Warehouse Picking System</span>
        </div>
        <div style={S.btnGroup}>
          <button style={{ ...S.btn, ...S.btnGray }} onClick={handleBackup}>
            ☁️ Backup
          </button>
          <button style={{ ...S.btn, ...S.btnGray }} onClick={() => restoreRef.current.click()}>
            🔄 Restore
          </button>
          <input
            ref={restoreRef}
            type="file"
            accept=".json"
            style={{ display: "none" }}
            onChange={handleRestore}
          />
          <button style={{ ...S.btn, ...S.btnRed }} onClick={handleClearAll}>
            🗑️ ล้างข้อมูลวันนี้
          </button>
        </div>
      </header>

      <div style={S.body}>
        {/* STATS */}
        <div style={S.statsRow}>
          {[
            { icon: "📋", label: "Total Orders", labelTh: "รายการทั้งหมด", val: totalOrders, color: "#dbeafe", num: "#2563eb" },
            { icon: "✅", label: "Completed", labelTh: "เสร็จสิ้น", val: completed, color: "#dcfce7", num: "#16a34a" },
            { icon: "⏳", label: "Pending", labelTh: "รอดำเนินการ", val: pending, color: "#fef9c3", num: "#d97706" },
            { icon: "🎁", label: "Total Stock", labelTh: "สินค้าคงเหลือรวม", val: totalStock.toLocaleString(), color: "#f3e8ff", num: "#9333ea" },
          ].map((s, i) => (
            <div key={i} style={S.statCard}>
              <div style={{ ...S.statIcon, background: s.color }}>{s.icon}</div>
              <div>
                <div style={{ fontSize: 12, color: "#7a8599" }}>{s.label}</div>
                <div style={{ ...S.statNum, color: s.num }}>{s.val}</div>
                <div style={{ fontSize: 11, color: "#9aa5bb" }}>{s.labelTh}</div>
              </div>
            </div>
          ))}
        </div>

        {/* SCANNER */}
        <div style={S.card}>
          <div style={S.sectionTitle}>
            <span>📲</span> สแกนหยิบสินค้า (Mobile Scanner)
            {/* ── badge แสดงว่า auto-confirm เปิดอยู่ ── */}
            <span style={{
              marginLeft: 8,
              background: "#dcfce7",
              color: "#16a34a",
              fontSize: 11,
              fontWeight: 700,
              padding: "2px 10px",
              borderRadius: 20,
              border: "1px solid #bbf7d0",
            }}>
              ⚡ Auto Confirm เปิดอยู่
            </span>
          </div>

          {/* hint */}
          <div style={{ fontSize: 12, color: "#7a8599", marginBottom: 10, lineHeight: 1.6 }}>
            💡 <strong>วิธีใช้:</strong> กรอก Pick No → สแกน/กรอก Barcode แล้ว<strong>กด Enter</strong> — ระบบจะบันทึกการหยิบอัตโนมัติทันที
          </div>

          <div style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr 1fr 100px auto auto",
            gap: 10,
            alignItems: "end",
          }}>
            <div>
              <div style={{ fontSize: 12, color: "#7a8599", marginBottom: 4 }}>Pick No</div>
              <input
                style={S.input}
                placeholder="สแกนหรือกรอก Pick No"
                value={scanPickNo}
                onChange={(e) => setScanPickNo(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    scanBarcodeInputRef.current?.focus();
                  }
                }}
              />
            </div>
            <div>
              <div style={{ fontSize: 12, color: "#7a8599", marginBottom: 4 }}>
                Barcode{" "}
                <span style={{ color: "#16a34a", fontSize: 11 }}>(กด Enter เพื่อบันทึก)</span>
              </div>
              <input
                ref={scanBarcodeInputRef}
                style={{
                  ...S.input,
                  borderColor: "#16a34a",
                  boxShadow: "0 0 0 2px rgba(22,163,74,0.15)",
                }}
                placeholder="สแกนหรือกรอก Barcode แล้วกด Enter"
                value={scanBarcode}
                onChange={(e) => setScanBarcode(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    handleScanConfirm();
                  }
                }}
              />
            </div>
            <div>
              <div style={{ fontSize: 12, color: "#7a8599", marginBottom: 4 }}>
                Location{" "}
                <span style={{ color: "#f59e0b", fontSize: 11 }}>(ระบุเพื่อแยก barcode ซ้ำ)</span>
              </div>
              <input
                style={{
                  ...S.input,
                  borderColor: "#f59e0b",
                  boxShadow: "0 0 0 2px rgba(245,158,11,0.12)",
                }}
                placeholder="เช่น Q14-028-40"
                value={scanLocation}
                onChange={(e) => setScanLocation(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    handleScanConfirm();
                  }
                }}
              />
            </div>
            <div>
              <div style={{ fontSize: 12, color: "#7a8599", marginBottom: 4 }}>Qty</div>
              <input
                style={S.input}
                type="number"
                min={1}
                value={scanQty}
                onChange={(e) => setScanQty(e.target.value)}
              />
            </div>
            <button
              style={{ ...S.btn, background: "#0ea5e9", color: "#fff" }}
              onClick={openCamera}
            >
              📷 สแกนกล้องมือถือ
            </button>
            <button
              style={{ ...S.btn, ...S.btnGreen }}
              onClick={handleScanConfirm}
            >
              ✅ ยืนยันการหยิบสินค้า
            </button>
          </div>
          <div style={{ marginTop: 8, fontSize: 13 }}>
            สถานะ:{" "}
            <span style={{
              color: scanStatus.startsWith("✅")
                ? "#16a34a"
                : scanStatus.startsWith("❌") || scanStatus.startsWith("⚠️")
                ? "#ef4444"
                : "#2563eb",
              fontWeight: 600,
            }}>
              {scanStatus}
            </span>
          </div>
        </div>

        {/* ORDERS TABLE */}
        <div style={S.card}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <div style={S.sectionTitle}>
              <span>📋</span> รายการ Pick Order
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <input
                style={{ ...S.input, width: 200 }}
                placeholder="ค้นหา Pick Order..."
                value={orderSearch}
                onChange={(e) => setOrderSearch(e.target.value)}
              />
              <button style={{ ...S.btn, ...S.btnExport }} onClick={handleExport}>
                ⬇️ Export Excel
              </button>
            </div>
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={S.table}>
              <thead>
                <tr>
                  {["Pick No", "Customer", "Barcode", "Product", "Location", "Required", "Picked", "Status", "Alert", "Created At", "Completed At", "Auto Time", ""].map((h) => (
                    <th key={h} style={S.th}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredOrders.map((o, i) => {
                  const al = alertLabel(o);
                  return (
                    <tr key={o.id} style={{ background: o.status === "PENDING" && al.cls === "over3" ? "#fff7ed" : "" }}>
                      <td style={S.td}>{o.id}</td>
                      <td style={S.td}>{o.customer}</td>
                      <td style={S.td}>{o.barcode}</td>
                      <td style={S.td}>{o.product}</td>
                      <td style={S.td}>{o.location}</td>
                      <td style={S.td}>{o.required}</td>
                      <td style={S.td}>{o.picked}</td>
                      <td style={S.td}>
                        <span style={{ ...S.badge, ...(S.statusTag[o.status] || { background: "#e5e7eb", color: "#374151" }) }}>
                          {o.status}
                        </span>
                      </td>
                      <td style={S.td}>
                        <span style={{ ...S.badge, ...(S.alertTag[al.cls] || {}) }}>
                          {al.text}
                        </span>
                      </td>
                      <td style={S.td} nowrap="true">{fmtDate(o.createdAt)}</td>
                      <td style={S.td} nowrap="true">{fmtDate(o.completedAt)}</td>
                      <td style={S.td}>{o.autoTime || "-"}</td>
                      <td style={S.td}>
                        <button
                          style={{ ...S.btn, ...S.btnRed, padding: "4px 10px", fontSize: 12 }}
                          onClick={() => deleteOrder(i)}
                        >
                          🗑️
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ marginTop: 10, fontSize: 12, color: "#9aa5bb", display: "flex", gap: 18, flexWrap: "wrap" }}>
            <span><span style={{ color: "#16a34a", fontWeight: 700 }}>COMPLETE</span> = เสร็จสิ้น</span>
            <span style={{ marginLeft: 12 }}><span style={{ color: "#d97706", fontWeight: 700 }}>PENDING</span> = รอดำเนินการ</span>
            <span style={{ marginLeft: 12 }}><span style={{ color: "#dc2626", fontWeight: 700 }}>OVER 3 HOURS</span> = ค้างเกิน 3 ชั่วโมง</span>
            <span style={{ marginLeft: 12 }}><span style={{ color: "#475569", fontWeight: 700 }}>NORMAL</span> = ปกติ</span>
            <span style={{ marginLeft: 12 }}><span style={{ color: "#1d4ed8", fontWeight: 700 }}>DONE</span> = ดำเนินการเสร็จแล้ว</span>
          </div>
        </div>

        {/* ADD STOCK & ADD ORDER */}
        <div style={S.bottomRow}>
          <div style={S.card}>
            <div style={S.sectionTitle}><span>📦</span> Add Stock</div>
            <div style={S.formGrid4}>
              <input style={S.input} placeholder="Barcode" value={newStock.barcode} onChange={(e) => setNewStock({ ...newStock, barcode: e.target.value })} onKeyDown={(e) => { if (e.key === "Enter") handleAddStock(); }} />
              <input style={S.input} placeholder="Product Name" value={newStock.product} onChange={(e) => setNewStock({ ...newStock, product: e.target.value })} onKeyDown={(e) => { if (e.key === "Enter") handleAddStock(); }} />
              <input style={S.input} placeholder="Stock" type="number" value={newStock.stock} onChange={(e) => setNewStock({ ...newStock, stock: e.target.value })} onKeyDown={(e) => { if (e.key === "Enter") handleAddStock(); }} />
            </div>
            <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
              <input style={{ ...S.input, flex: 1 }} placeholder="Location" value={newStock.location} onChange={(e) => setNewStock({ ...newStock, location: e.target.value })} onKeyDown={(e) => { if (e.key === "Enter") handleAddStock(); }} />
              <button style={{ ...S.btn, ...S.btnBlue }} onClick={handleAddStock}>+ Add Stock</button>
            </div>
            <hr style={{ margin: "16px 0", border: "none", borderTop: "1px solid #f0f2f7" }} />
            <div style={S.sectionTitle}><span>📁</span> Upload Excel Stock</div>
            <div style={{ fontSize: 11, color: "#9aa5bb", marginBottom: 6 }}>
              รองรับชื่อคอลัมน์: <strong>barcode/บาร์โค้ด, product/สินค้า/ชื่อสินค้า, stock/จำนวน/คงเหลือ, location/ที่เก็บ/ตำแหน่ง</strong>
            </div>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={(e) => {
                const file = e.target.files[0];
                if (!file) return;
                const reader = new FileReader();
                reader.onload = (ev) => {
                  try {
                    const wb = XLSX.read(ev.target.result, { type: "array" });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
                    if (!rows.length) { showToast("❌ ไม่พบข้อมูลในไฟล์"); return; }
                    const keys = Object.keys(rows[0]);
                    const findCol = (aliases) => keys.find((k) => aliases.some((a) => k.toLowerCase().trim() === a.toLowerCase().trim())) || "";
                    const colBarcode = findCol(["barcode", "บาร์โค้ด", "รหัสสินค้า", "รหัส"]);
                    const colProduct = findCol(["product", "product name", "ชื่อสินค้า", "สินค้า", "ชื่อ", "รายการ"]);
                    const colStock = findCol(["stock", "จำนวน", "คงเหลือ", "qty", "quantity", "จำนวนคงเหลือ"]);
                    const colLocation = findCol(["location", "ที่เก็บ", "ตำแหน่ง", "ชั้น", "โลเคชั่น", "loc"]);
                    if (!colBarcode) { showToast(`❌ ไม่พบคอลัมน์ barcode — คอลัมน์ในไฟล์: ${keys.join(", ")}`); return; }
                    const mapped = rows.map((r) => ({
                      barcode: String(r[colBarcode] || "").trim(),
                      product: String(colProduct ? r[colProduct] : "").trim(),
                      stock: Number(colStock ? r[colStock] : 0),
                      location: String(colLocation ? r[colLocation] : "").trim(),
                    })).filter((r) => r.barcode);
                    if (!mapped.length) { showToast("❌ ไม่พบข้อมูลที่มี barcode"); return; }
                    setStock((prev) => {
                      const updated = [...prev];
                      mapped.forEach((row) => {
                        const idx = updated.findIndex((s) => s.barcode === row.barcode);
                        if (idx !== -1) updated[idx] = { ...updated[idx], ...row, stock: updated[idx].stock + row.stock };
                        else updated.push(row);
                      });
                      return updated;
                    });
                    showToast(`✅ นำเข้า Stock สำเร็จ ${mapped.length} รายการ`);
                  } catch (err) { showToast("❌ อ่านไฟล์ไม่ได้: " + err.message); }
                };
                reader.readAsArrayBuffer(file);
                e.target.value = "";
              }}
            />
          </div>

          <div style={S.card}>
            <div style={S.sectionTitle}><span>📋</span> Add Pick Order</div>
            <div style={S.formGrid4}>
              <input style={S.input} placeholder="Pick No" value={newOrder.id} onChange={(e) => setNewOrder({ ...newOrder, id: e.target.value })} onKeyDown={(e) => { if (e.key === "Enter") handleAddOrder(); }} />
              <input style={S.input} placeholder="Customer" value={newOrder.customer} onChange={(e) => setNewOrder({ ...newOrder, customer: e.target.value })} onKeyDown={(e) => { if (e.key === "Enter") handleAddOrder(); }} />
              <input style={S.input} placeholder="Barcode" value={newOrder.barcode} onChange={(e) => setNewOrder({ ...newOrder, barcode: e.target.value })} onKeyDown={(e) => { if (e.key === "Enter") handleAddOrder(); }} />
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: 10, alignItems: "center", marginBottom: 0 }}>
              <input style={S.input} placeholder="Qty" type="number" value={newOrder.qty} onChange={(e) => setNewOrder({ ...newOrder, qty: e.target.value })} onKeyDown={(e) => { if (e.key === "Enter") handleAddOrder(); }} />
              <input
                style={S.input}
                placeholder="Location (ถ้าไม่กรอกใช้จาก Stock)"
                value={newOrder.location}
                onChange={(e) => setNewOrder({ ...newOrder, location: e.target.value })}
                onKeyDown={(e) => { if (e.key === "Enter") handleAddOrder(); }}
              />
              <button style={{ ...S.btn, ...S.btnBlue }} onClick={handleAddOrder}>+ Add Order</button>
            </div>
            <hr style={{ margin: "16px 0", border: "none", borderTop: "1px solid #f0f2f7" }} />
            <div style={S.sectionTitle}><span>📁</span> Upload Excel Orders</div>
            <div style={{ fontSize: 11, color: "#9aa5bb", marginBottom: 6 }}>
              รองรับชื่อคอลัมน์: <strong>id/pick no/เลขที่, customer/ลูกค้า, barcode/บาร์โค้ด, qty/จำนวน, location/ที่เก็บ/ตำแหน่ง</strong>
            </div>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={(e) => {
                const file = e.target.files[0];
                if (!file) return;
                const reader = new FileReader();
                reader.onload = (ev) => {
                  try {
                    const wb = XLSX.read(ev.target.result, { type: "array" });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
                    if (!rows.length) { showToast("❌ ไม่พบข้อมูลในไฟล์"); return; }
                    const keys = Object.keys(rows[0]);
                    const findCol = (aliases) => keys.find((k) => aliases.some((a) => k.toLowerCase().trim() === a.toLowerCase().trim())) || "";
                    const colId = findCol(["id", "pick no", "pickno", "เลขที่", "รหัสคำสั่ง", "เลขออเดอร์", "order no", "pick_no"]);
                    const colCustomer = findCol(["customer", "ลูกค้า", "ชื่อลูกค้า", "บริษัท", "company"]);
                    const colBarcode = findCol(["barcode", "บาร์โค้ด", "รหัสสินค้า", "รหัส"]);
                    const colQty = findCol(["qty", "จำนวน", "quantity", "required", "จำนวนที่ต้องการ"]);
                    const colLocation = findCol(["location", "ที่เก็บ", "ตำแหน่ง", "loc", "โลเคชั่น", "ชั้น"]);
                    if (!colBarcode) { showToast(`❌ ไม่พบคอลัมน์ barcode — คอลัมน์ในไฟล์: ${keys.join(", ")}`); return; }
                    if (!colId) { showToast(`❌ ไม่พบคอลัมน์ id/pick no — คอลัมน์ในไฟล์: ${keys.join(", ")}`); return; }
                    const mapped = rows.map((r) => ({
                      id: String(r[colId] || "").trim(),
                      customer: String(colCustomer ? r[colCustomer] : "").trim(),
                      barcode: String(r[colBarcode] || "").trim(),
                      qty: Number(colQty ? r[colQty] : 1),
                      location: String(colLocation ? r[colLocation] : "").trim(),
                    })).filter((r) => r.id && r.barcode);
                    if (!mapped.length) { showToast("❌ ไม่พบข้อมูลที่มี id และ barcode"); return; }
                    const newOrders = mapped.map((r) => {
                      const stockItem = stock.find((s) => s.barcode === r.barcode);
                      return {
                        id: r.id, customer: r.customer, barcode: r.barcode,
                        product: stockItem?.product || "-",
                        location: r.location || stockItem?.location || "-",
                        required: r.qty, picked: 0, status: "PENDING",
                        createdAt: new Date().toISOString(), completedAt: null, autoTime: null,
                      };
                    });
                    setOrders((prev) => [...prev, ...newOrders]);
                    showToast(`✅ นำเข้า Order สำเร็จ ${newOrders.length} รายการ`);
                  } catch (err) { showToast("❌ อ่านไฟล์ไม่ได้: " + err.message); }
                };
                reader.readAsArrayBuffer(file);
                e.target.value = "";
              }}
            />
          </div>
        </div>

        {/* STOCK คงเหลือ */}
        <div style={S.card}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <div style={S.sectionTitle}><span>📦</span> Stock คงเหลือ</div>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <input
                style={{ ...S.input, width: 220 }}
                placeholder="ค้นหา Stock..."
                value={stockSearch}
                onChange={(e) => setStockSearch(e.target.value)}
              />
              <button
                style={{ ...S.btn, ...S.btnRed }}
                onClick={() => {
                  if (window.confirm("ยืนยันลบ Stock ทั้งหมด? (เพื่อ Upload ใหม่)")) {
                    setStock([]);
                    showToast("🗑️ ลบ Stock ทั้งหมดแล้ว พร้อม Upload ใหม่");
                  }
                }}
              >
                🗑️ ลบทั้งหมด
              </button>
            </div>
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={S.table}>
              <thead>
                <tr>
                  {["Barcode", "Product", "Stock", "Location", "แก้ไข"].map((h) => (
                    <th key={h} style={S.th}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredStock.map((s, i) => (
                  <tr key={s.barcode}>
                    {editStockIdx === stock.indexOf(s) ? (
                      <>
                        <td style={S.td}>
                          <input style={{ ...S.input, width: 150 }} value={editStockData.barcode}
                            onChange={(e) => setEditStockData({ ...editStockData, barcode: e.target.value })}
                            onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); } }} />
                        </td>
                        <td style={S.td}>
                          <input style={{ ...S.input, width: 130 }} value={editStockData.product}
                            onChange={(e) => setEditStockData({ ...editStockData, product: e.target.value })} />
                        </td>
                        <td style={S.td}>
                          <input style={{ ...S.input, width: 80 }} type="number" value={editStockData.stock}
                            onChange={(e) => setEditStockData({ ...editStockData, stock: e.target.value })} />
                        </td>
                        <td style={S.td}>
                          <input style={{ ...S.input, width: 100 }} value={editStockData.location}
                            onChange={(e) => setEditStockData({ ...editStockData, location: e.target.value })} />
                        </td>
                        <td style={S.td}>
                          <button style={{ ...S.btn, ...S.btnGreen, padding: "4px 10px", fontSize: 12 }} onClick={saveEditStock}>
                            💾 บันทึก
                          </button>
                        </td>
                      </>
                    ) : (
                      <>
                        <td style={S.td}>{s.barcode}</td>
                        <td style={S.td}>{s.product}</td>
                        <td style={S.td}>
                          <strong style={{ color: s.stock <= 10 ? "#ef4444" : s.stock <= 30 ? "#d97706" : "#16a34a" }}>
                            {s.stock}
                          </strong>
                        </td>
                        <td style={S.td}>{s.location}</td>
                        <td style={S.td}>
                          <div style={{ display: "flex", gap: 4 }}>
                            <button style={{ ...S.btn, background: "#2563eb", color: "#fff", padding: "4px 10px", fontSize: 12 }}
                              onClick={() => startEditStock(stock.indexOf(s))}>
                              ✏️ แก้ไข
                            </button>
                            <button style={{ ...S.btn, ...S.btnRed, padding: "4px 10px", fontSize: 12 }}
                              onClick={() => deleteStock(stock.indexOf(s))}>
                              🗑️ ลบ
                            </button>
                          </div>
                        </td>
                      </>
                    )}
                  </tr>
                ))}
                {filteredStock.length === 0 && (
                  <tr>
                    <td colSpan={5} style={{ ...S.td, textAlign: "center", color: "#9aa5bb", padding: 24 }}>
                      ไม่พบข้อมูล Stock
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* FOOTER */}
      <footer style={S.footer}>
        <span>📦 Warehouse Picking System v1.0</span>
        <span>Auto Delete: ระบบจะลบข้อมูลอัตโนมัติเมื่อข้ามวัน</span>
        <span>© 2024 All Rights Reserved</span>
      </footer>

      {/* CAMERA MODAL */}
      {cameraOpen && (
        <div style={{
          position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)",
          zIndex: 9000, display: "flex", flexDirection: "column",
          alignItems: "center", justifyContent: "center", gap: 16,
        }}>
          <div style={{
            background: "#1a2233", borderRadius: 16, padding: 24,
            width: "min(420px,95vw)", boxShadow: "0 8px 40px rgba(0,0,0,.6)",
          }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <span style={{ color: "#fff", fontWeight: 700, fontSize: 16 }}>📷 สแกน Barcode ด้วยกล้อง</span>
              <button style={{ ...S.btn, ...S.btnRed, padding: "4px 12px" }} onClick={closeCamera}>✕ ปิด</button>
            </div>
            {cameraError ? (
              <div style={{ color: "#fca5a5", fontSize: 13, marginBottom: 12, lineHeight: 1.6 }}>{cameraError}</div>
            ) : (
              <div style={{ position: "relative", borderRadius: 10, overflow: "hidden", background: "#000", aspectRatio: "4/3" }}>
                <video ref={videoRef} style={{ width: "100%", height: "100%", objectFit: "cover" }} muted playsInline autoPlay />
                <div style={{
                  position: "absolute", top: "50%", left: "10%", right: "10%",
                  height: 2, background: "#22d3ee", boxShadow: "0 0 8px #22d3ee",
                  animation: "scanline 1.5s ease-in-out infinite",
                }} />
                <div style={{
                  position: "absolute", inset: 0, border: "2px solid rgba(34,211,238,0.4)",
                  borderRadius: 8, pointerEvents: "none",
                }} />
              </div>
            )}
            {scanning && !cameraError && (
              <div style={{ color: "#7dd3fc", fontSize: 12, textAlign: "center", marginTop: 10 }}>
                🔍 กำลังสแกน... เล็งบาร์โค้ดให้อยู่กลางกล้อง
              </div>
            )}
            <div style={{ marginTop: 14 }}>
              <div style={{ color: "#9aa5bb", fontSize: 12, marginBottom: 6 }}>หรือกรอก Barcode ตรงนี้:</div>
              <div style={{ display: "flex", gap: 8 }}>
                <input
                  style={{ ...S.input, flex: 1, background: "#2d3748", border: "1px solid #4a5568", color: "#fff" }}
                  placeholder="กรอก Barcode แล้วกด Enter"
                  autoFocus
                  onKeyDown={(e) => {
                    if (e.key === "Enter" && e.target.value.trim()) {
                      setScanBarcode(e.target.value.trim());
                      setScanStatus(`⌨️ กรอก Barcode: ${e.target.value.trim()}`);
                      showToast(`✅ Barcode: ${e.target.value.trim()}`);
                      closeCamera();
                    }
                  }}
                />
                <button
                  style={{ ...S.btn, ...S.btnGreen }}
                  onClick={(e) => {
                    const inp = e.target.closest("div").querySelector("input");
                    if (inp?.value.trim()) {
                      setScanBarcode(inp.value.trim());
                      setScanStatus(`⌨️ กรอก Barcode: ${inp.value.trim()}`);
                      showToast(`✅ Barcode: ${inp.value.trim()}`);
                      closeCamera();
                    }
                  }}
                >
                  ยืนยัน
                </button>
              </div>
            </div>
          </div>
          <style>{`@keyframes scanline{0%{top:20%}50%{top:75%}100%{top:20%}}`}</style>
        </div>
      )}

      {toast && <Toast msg={toast} onClose={() => setToast(null)} />}
    </div>
  );
}
