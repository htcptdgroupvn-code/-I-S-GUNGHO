import React, { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";
import {
  Users, ClipboardList, BarChart3, Bell, LogOut, CheckCircle2, XCircle,
  Send, UserPlus, ShoppingBag, TrendingUp, Award, Store, Wallet,
  ArrowRightLeft, RefreshCw, Phone, MapPin, Plus, ChevronRight, Inbox,
  ClipboardCheck, Building2, Landmark, AlertCircle, X, Clock, Download, FileText, Search, Megaphone, Pin, Pencil, Trash2
} from "lucide-react";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  PieChart, Pie, Cell, AreaChart, Area, Legend,
} from "recharts";

// ---------------------------------------------------------------------------
// Bảng màu dùng chung cho biểu đồ (đồng bộ với bảng màu giao diện teal/amber/...)
// ---------------------------------------------------------------------------
const CHART_COLORS = ["#0f766e", "#d97706", "#4f46e5", "#e11d48", "#0891b2", "#65a30d", "#9333ea", "#ea580c"];
const CHART_TOOLTIP_STYLE = {
  contentStyle: { borderRadius: 12, border: "1px solid #e2e8f0", boxShadow: "0 4px 16px rgba(15,23,42,.08)", fontSize: 12.5 },
  labelStyle: { color: "#334155", fontWeight: 600, marginBottom: 4 },
};
function shortMoney(n) {
  const v = Number(n) || 0;
  if (v >= 1_000_000_000) return (v / 1_000_000_000).toFixed(1).replace(/\.0$/, "") + " tỷ";
  if (v >= 1_000_000) return (v / 1_000_000).toFixed(1).replace(/\.0$/, "") + " tr";
  if (v >= 1_000) return (v / 1_000).toFixed(0) + " k";
  return String(v);
}
// Nhãn rút gọn dùng cho trục/chú thích của các biểu đồ recharts còn lại (tránh tràn chữ)
function truncateLabel(s, n = 14) {
  if (!s) return "";
  return s.length > n ? s.slice(0, n - 1) + "…" : s;
}

// ---------------------------------------------------------------------------
// Static demo data
// ---------------------------------------------------------------------------

const COMPANIES = [
  {
    name: "I. Công ty Cổ phần Thương mại I - Khối xe máy",
    products: ["Xe máy, xe đạp/máy điện", "Bảo hiểm xe máy", "Phụ tùng bán lẻ (phụ tùng, phụ kiện)", "Dịch vụ sửa chữa"],
    branches: [
      { name: "Phú Tài Đức 1", address: "Số 09, đường Trần Phú, phường Thành Sen, tỉnh Hà Tĩnh" },
      { name: "Phú Tài Đức 2", address: "Số 218, Đường mòn Hồ Chí Minh, xã Hương Khê, tỉnh Hà Tĩnh" },
      { name: "Phú Tài Đức 3", address: "Số 190, đường Lý Tự Trọng, xã Thạch Hà, tỉnh Hà Tĩnh" },
      { name: "Phú Tài Đức 4", address: "Số 19, đường Xô Viết Nghệ Tĩnh, xã Can Lộc, tỉnh Hà Tĩnh" },
      { name: "Phú Tài Đức 5", address: "Số 183, Hà Huy Tập, xã Cẩm Xuyên, tỉnh Hà Tĩnh" },
      { name: "Phú Tài Đức 6", address: "Số 302, đường Lê Đại Hành, phường Sông Trí, tỉnh Hà Tĩnh" },
      { name: "Phú Tài Đức 9", address: "Số 103, đường Lê Đại Hành, phường Sông Trí, tỉnh Hà Tĩnh" },
      { name: "Phú Tài Đức 10", address: "Số 120, đường Trần Phú, phường Thành Sen, tỉnh Hà Tĩnh" },
      { name: "Phú Tài Đức 12", address: "Số 328, đường Lê Lợi, xã Hương Sơn, tỉnh Hà Tĩnh" },
      { name: "Phú Tài Đức 13", address: "TDP8, đường Phan Đình Phùng, xã Đức Thọ, tỉnh Hà Tĩnh" },
    ],
  },
  {
    name: "II. Công ty Cổ phần Thương mại I - Khối ô tô",
    products: ["Xe mới ô tô", "Phụ kiện, chăm sóc xe", "Bảo hiểm ô tô", "Sửa chữa, bảo dưỡng"],
    branches: [
      { name: "Đại lý Toyota Phú Tài Đức", address: "Số 15 đường Trần Phú, phường Thành Sen, tỉnh Hà Tĩnh" },
    ],
  },
  {
    name: "III. HTC",
    products: ["Đặc sản địa phương", "Phòng nghỉ", "Các tiệc lưu động, hội họp, hội nghị, sinh nhật, liên hoan, ăn sáng", "Vé máy bay", "Tour"],
    branches: [
      { name: "Khách sạn Bình Minh", address: "Số 09, đường Trần Phú, phường Thành Sen, tỉnh Hà Tĩnh" },
      { name: "Khách sạn White Place", address: "Số 139, Hà Huy Tập, phường Thành Sen, tỉnh Hà Tĩnh" },
      { name: "Khách sạn Hoành Sơn", address: "Phường Hoành Sơn, tỉnh Hà Tĩnh" },
      { name: "Văn phòng lữ hành", address: "Tầng 4, tòa nhà Toyota, số 15 đường Trần Phú, phường Thành Sen, tỉnh Hà Tĩnh" },
    ],
  },
  {
    name: "IV. VYC",
    products: ["Sản phẩm thời trang"],
    branches: [
      { name: "Văn phòng, phòng trưng bày và xưởng sản xuất VYC", address: "Số 15 đường Trần Phú, phường Thành Sen, tỉnh Hà Tĩnh" },
    ],
  },
  {
    name: "V. Vật tư nông nghiệp",
    products: ["Phân NPK Hàn Việt", "Phân NPK Hà Lan", "Phân đạm cánh diều (Đạm trắng và đạm vàng)", "Phân hữu cơ cánh diều", "Một số sản phẩm phân bón khác"],
    branches: [
      { name: "Kho bãi", address: "Số 359, đường Hà Huy Tập, phường Thành Sen, tỉnh Hà Tĩnh" },
    ],
  },
];

const ALL_BRANCHES = COMPANIES.flatMap((c) => c.branches.map((b) => ({ ...b, company: c.name })));
const branchInfo = (name) => ALL_BRANCHES.find((b) => b.name === name);

// Phân loại sản phẩm theo công thức tính điểm bình xét thi đua Gung Ho
const PRODUCT_CATEGORY = {
  "Bảo hiểm xe máy": "BX",
  "Phụ tùng bán lẻ (phụ tùng, phụ kiện)": "DVX",
  "Dịch vụ sửa chữa": "DVX",
  "Phụ kiện, chăm sóc xe": "DVO",
  "Bảo hiểm ô tô": "BO",
  "Sửa chữa, bảo dưỡng": "DVO",
  "Đặc sản địa phương": "NH",
  "Phòng nghỉ": "KS",
  "Các tiệc lưu động, hội họp, hội nghị, sinh nhật, liên hoan, ăn sáng": "NH",
  "Vé máy bay": "V",
  "Tour": "TO",
  "Sản phẩm thời trang": "VYC",
  "Phân NPK Hàn Việt": "VT",
  "Phân NPK Hà Lan": "VT",
  "Phân đạm cánh diều (Đạm trắng và đạm vàng)": "VT",
  "Phân hữu cơ cánh diều": "VT",
  "Một số sản phẩm phân bón khác": "VT",
};
// Các hạng mục tính theo số lượng (năm bảo hiểm / vé / tour) thay vì doanh thu
const COUNT_CATEGORIES = new Set(["BX", "BO", "TO", "V"]);
const CATEGORY_LABELS = { BX: "Bảo hiểm xe máy (năm)", BO: "Bảo hiểm ô tô (số lượng)", TO: "Tour (số lượng)", V: "Vé máy bay (số lượng)" };

// TD = BX*0.5 + DVX/500.000 + BO*2 + DVO/1.000.000 + NH/1.000.000 + KS/500.000 + TO*20 + V*1 + VYC/300.000 + VT/500.000
function computeTD(paidOrders) {
  const s = { BX: 0, DVX: 0, BO: 0, DVO: 0, NH: 0, KS: 0, TO: 0, V: 0, VYC: 0, VT: 0 };
  paidOrders.forEach((o) => {
    const cat = PRODUCT_CATEGORY[o.product];
    if (!cat) return;
    if (COUNT_CATEGORIES.has(cat)) {
      s[cat] += Number(o.quantity) || 0;
    } else {
      s[cat] += o.finalAmount ?? o.totalAmount ?? 0;
    }
  });
  const td = s.BX * 0.5 + s.DVX / 500000 + s.BO * 2 + s.DVO / 1000000 + s.NH / 1000000 + s.KS / 500000 + s.TO * 20 + s.V * 1 + s.VYC / 300000 + s.VT / 500000;
  return Math.round(td * 10) / 10;
}

function buildRevenueLeaderboard(orders, groupKeyFn) {
  const map = new Map();
  orders.filter((o) => o.status === "da_thanh_toan").forEach((o) => {
    const key = groupKeyFn(o);
    if (!key) return;
    const cur = map.get(key) || { revenue: 0, count: 0 };
    cur.revenue += o.finalAmount ?? o.totalAmount ?? 0;
    cur.count += 1;
    map.set(key, cur);
  });
  return [...map.entries()].map(([name, v]) => ({ name, ...v })).sort((a, b) => b.revenue - a.revenue);
}

function buildTDLeaderboard(orders, groupKeyFn) {
  const paid = orders.filter((o) => o.status === "da_thanh_toan");
  const map = new Map();
  paid.forEach((o) => {
    const key = groupKeyFn(o);
    if (!key) return;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(o);
  });
  return [...map.entries()].map(([name, list]) => ({ name, td: computeTD(list) })).sort((a, b) => b.td - a.td);
}

// Dựng dữ liệu xu hướng doanh thu theo tháng (6 tháng gần nhất) cho biểu đồ đường/vùng
function buildMonthlyTrend(orders, monthsBack = 6) {
  const paid = orders.filter((o) => o.status === "da_thanh_toan");
  const now = new Date();
  const buckets = [];
  for (let i = monthsBack - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    buckets.push({ y: d.getFullYear(), m: d.getMonth(), label: `Th${d.getMonth() + 1}/${String(d.getFullYear()).slice(2)}`, revenue: 0, count: 0 });
  }
  paid.forEach((o) => {
    const d = new Date(o.updatedAt || o.createdAt);
    const b = buckets.find((x) => x.y === d.getFullYear() && x.m === d.getMonth());
    if (b) {
      b.revenue += o.finalAmount ?? o.totalAmount ?? 0;
      b.count += 1;
    }
  });
  return buckets;
}

// Dựng dữ liệu tỉ trọng doanh thu theo nhóm sản phẩm (dùng cho biểu đồ tròn)
function buildProductMix(orders, topN = 6) {
  const rows = buildRevenueLeaderboard(orders, (o) => o.product);
  const top = rows.slice(0, topN);
  const rest = rows.slice(topN);
  const restSum = rest.reduce((s, r) => s + r.revenue, 0);
  const result = top.map((r) => ({ name: r.name, value: r.revenue }));
  if (restSum > 0) result.push({ name: "Khác", value: restSum });
  return result.filter((r) => r.value > 0);
}

function exportToExcel(sheets, filename) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => {
    const ws = XLSX.utils.json_to_sheet(rows && rows.length ? rows : [{ "Không có dữ liệu": "" }]);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}

function orderExportRow(o) {
  return {
    "Mã đơn hàng": o.orderCode || o.id,
    "Mã khách hàng": o.customerCode || "",
    "Khách hàng": o.customerName,
    "SĐT": o.customerPhone,
    "Khối công ty": o.company,
    "Cửa hàng / chi nhánh": o.store,
    "Sản phẩm": o.product,
    "Đại sứ": o.createdByName,
    "Người chăm sóc": o.assignedHandlerName || "",
    "Trạng thái": STATUS_META[o.status]?.label || o.status,
    "Số tiền": o.status === "da_thanh_toan" ? (o.finalAmount ?? o.totalAmount) : "",
    "Hoa hồng": o.status === "da_thanh_toan" ? o.commissionAmount || 0 : "",
    "Ngày tạo": fmtDate(o.createdAt),
  };
}

// Mã khối công ty dùng trong mẫu báo cáo chính thức
const COMPANY_CODE = {
  "I. Công ty Cổ phần Thương mại I - Khối xe máy": "TM1",
  "II. Công ty Cổ phần Thương mại I - Khối ô tô": "TM2",
  "III. HTC": "HTC",
  "IV. VYC": "VYC",
  "V. Vật tư nông nghiệp": "VTNN",
};

function gunghoStatusLabel(status) {
  if (status === "da_thanh_toan") return "Thành công";
  if (status === "khong_thanh_toan") return "Không thành công";
  return "Đang chờ";
}

function daysBetween(isoLater, isoEarlier) {
  const ms = new Date(isoLater) - new Date(isoEarlier);
  return Math.max(Math.round(ms / (1000 * 60 * 60 * 24)), 0);
}

// Xuất đúng theo mẫu "KẾT QUẢ GUNG HO CHI TIẾT NHÂN VIÊN"
function exportGungHoNhanVienTemplate({ ambassador, orders, fromDate, toDate }) {
  const filtered = orders
    .filter((o) => o.createdBy === ambassador.id)
    .filter((o) => {
      if (fromDate && new Date(o.createdAt) < new Date(fromDate)) return false;
      if (toDate && new Date(o.createdAt) > new Date(toDate + "T23:59:59")) return false;
      return true;
    })
    .sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));

  const branch = branchInfo(ambassador.store);
  const companyCode = COMPANY_CODE[branch?.company] || "";

  const header = [
    ["CÔNG TY", companyCode, "", "", "Từ ngày tháng năm", fromDate ? fmtDate(fromDate).split(" ")[0] : "", "Đến ngày tháng năm", toDate ? fmtDate(toDate).split(" ")[0] : ""],
    ["ĐƠN VỊ", ambassador.store || ""],
    ["HỌ VÀ TÊN GUNG HO", ambassador.name],
    ["CHỨC VỤ", ambassador.position || ""],
    [],
    ["KẾT QUẢ GUNG HO CHI TIẾT NHÂN VIÊN"],
    [
      "TT", "Họ tên khách hàng", "Ngày, giờ đăng ký Gung Ho", "Nội dung đăng ký",
      "Họ và tên KH sử dụng SPDV", "Ngày, giờ sử dụng", "Sản phẩm sử dụng",
      "Doanh thu", "Giảm giá", "Số ngày Gung Ho", "Trạng thái", "Chiết khấu", "Ghi chú",
    ],
  ];

  const totalChietKhau = filtered.reduce((s, o) => s + (o.status === "da_thanh_toan" ? (o.commissionAmount || 0) : 0), 0);
  const totalGiamGia = filtered.reduce((s, o) => s + (o.discountAmount || 0), 0);
  const totalRow = ["", "TỔNG", "", "", "", "", "", "", totalGiamGia, "", "", totalChietKhau, ""];

  const dataRows = filtered.map((o, i) => {
    const cat = PRODUCT_CATEGORY[o.product];
    const isPaid = o.status === "da_thanh_toan";
    const doanhThu = !isPaid ? "" : COUNT_CATEGORIES.has(cat) ? (o.quantity || 0) : (o.finalAmount ?? o.totalAmount ?? 0);
    return [
      i + 1,
      o.customerName || "",
      fmtDate(o.createdAt),
      o.product || "",
      o.customerName || "",
      isPaid ? fmtDate(o.updatedAt) : "",
      o.product || "",
      doanhThu,
      o.discountAmount || 0,
      isPaid ? daysBetween(o.updatedAt, o.createdAt) : "",
      gunghoStatusLabel(o.status),
      isPaid ? (o.commissionAmount || 0) : "",
      o.accountantNote || o.handlerNote || "",
    ];
  });

  const aoa = [...header, totalRow, ...dataRows];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!merges"] = [{ s: { r: 5, c: 0 }, e: { r: 5, c: 12 } }];
  ws["!cols"] = [
    { wch: 4 }, { wch: 20 }, { wch: 20 }, { wch: 22 }, { wch: 20 }, { wch: 20 },
    { wch: 20 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 20 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Chi tiết nhân viên");
  XLSX.writeFile(wb, `KetQuaGungHo_${ambassador.name.replace(/\s+/g, "")}_${Date.now()}.xlsx`);
}

// Danh sách sản phẩm cụ thể dùng để tính từng cột trong mẫu "Chi tiết công ty theo đơn vị"
const P = {
  XE_MAY: "Xe máy, xe đạp/máy điện",
  BAO_HIEM_XE_MAY: "Bảo hiểm xe máy",
  PHU_TUNG: "Phụ tùng bán lẻ (phụ tùng, phụ kiện)",
  SUA_CHUA_XE_MAY: "Dịch vụ sửa chữa",
  O_TO: "Xe mới ô tô",
  BAO_HIEM_O_TO: "Bảo hiểm ô tô",
  PHU_KIEN_O_TO: "Phụ kiện, chăm sóc xe",
  SUA_CHUA_O_TO: "Sửa chữa, bảo dưỡng",
  DAC_SAN: "Đặc sản địa phương",
  PHONG_NGHI: "Phòng nghỉ",
  TIEC: "Các tiệc lưu động, hội họp, hội nghị, sinh nhật, liên hoan, ăn sáng",
  VE_MAY_BAY: "Vé máy bay",
  TOUR: "Tour",
  VYC: "Sản phẩm thời trang",
  VTNN: ["Phân NPK Hàn Việt", "Phân NPK Hà Lan", "Phân đạm cánh diều (Đạm trắng và đạm vàng)", "Phân hữu cơ cánh diều", "Một số sản phẩm phân bón khác"],
};

// Quy định thưởng Gung Ho riêng cho 4 sản phẩm thuộc Khối xe máy (Phú Tài Đức 1-10)
// Xe máy: thưởng theo mức giảm giá/xe, phân biệt theo loại xe
function computeXeMayRewardPerUnit(vehicleType, discountPerUnit) {
  const threshold = vehicleType === "ga_con" ? 400000 : 300000; // xe ga/côn: 400k, xe số/điện: 300k
  return Number(discountPerUnit) <= threshold ? 200000 : 100000;
}
// Bảo hiểm xe máy: thưởng theo thời hạn hợp đồng
function baoHiemRewardPerUnit(years) {
  return years === 2 ? 20000 : years === 3 ? 25000 : 15000;
}
// Phụ tùng bán lẻ / Dịch vụ sửa chữa / Phụ kiện ô tô: % doanh thu nếu KHÔNG giảm giá, 0đ nếu có giảm giá (xem ServiceRevenueForm)
const XE_MAY_SPECIAL_PRODUCTS = [P.XE_MAY, P.BAO_HIEM_XE_MAY, P.PHU_TUNG, P.SUA_CHUA_XE_MAY];

// Quy định thưởng Gung Ho riêng cho 4 sản phẩm thuộc Khối ô tô
// Xe mới ô tô: thưởng 900.000đ/xe, tỷ lệ hưởng theo mức giảm giá ngoài chính sách
function otoRewardPerUnit(discountPerUnit) {
  const d = Number(discountPerUnit) || 0;
  const base = 900000;
  if (d <= 0) return base; // không giảm giá ngoài chính sách: 100%
  if (d <= 2000000) return Math.round(base * 0.7); // ≤2.000.000đ: 70%
  return Math.round(base * 0.5); // >2.000.000đ: 50%
}
// Bảo hiểm ô tô: hoa hồng = 82% x mức chiết khấu chính sách (do THT thông báo, nhập tay) x doanh thu
function baoHiemOTOReward(revenue, hasDiscount, policyRatePercent) {
  if (hasDiscount) return 0;
  const gunghoRate = (0.82 * (Number(policyRatePercent) || 0)) / 100;
  return Math.round(Number(revenue || 0) * gunghoRate);
}
const OTO_SPECIAL_PRODUCTS = [P.O_TO, P.BAO_HIEM_O_TO, P.PHU_KIEN_O_TO, P.SUA_CHUA_O_TO];

// Quy định thưởng Gung Ho riêng cho các sản phẩm thuộc Khối HTC
// Đặc sản địa phương / Phòng nghỉ / Tiệc lưu động - nhà hàng - ăn sáng: 3% doanh thu (không phân biệt giảm giá)
// Vé máy bay: theo số lượng vé, phân biệt khách lẻ (15.000đ/vé) và khách đoàn (10.000đ/vé)
function veMayBayRewardPerUnit(customerType) {
  return customerType === "doan" ? 10000 : 15000;
}
// Tour nội địa: 500.000đ/hợp đồng thành công
const TOUR_REWARD_PER_CONTRACT = 500000;
const HTC_SPECIAL_PRODUCTS = [P.DAC_SAN, P.PHONG_NGHI, P.TIEC, P.VE_MAY_BAY, P.TOUR];
// VYC: Sản phẩm thời trang — 5% doanh thu. Vật tư nông nghiệp: Phân bón — 2% doanh thu.
const VYC_SPECIAL_PRODUCTS = [P.VYC];
const VTNN_SPECIAL_PRODUCTS = P.VTNN;

function sumRevenue(orders, productNames) {
  const set = new Set(Array.isArray(productNames) ? productNames : [productNames]);
  return orders.filter((o) => o.status === "da_thanh_toan" && set.has(o.product)).reduce((s, o) => s + (o.finalAmount ?? o.totalAmount ?? 0), 0);
}
function countPaidOrders(orders, productNames) {
  const set = new Set(Array.isArray(productNames) ? productNames : [productNames]);
  return orders.filter((o) => o.status === "da_thanh_toan" && set.has(o.product)).length;
}
function sumQuantity(orders, productName) {
  return orders.filter((o) => o.status === "da_thanh_toan" && o.product === productName).reduce((s, o) => s + (Number(o.quantity) || 0), 0);
}

function buildUnitRow(unitOrders) {
  const scXeMay = sumRevenue(unitOrders, P.SUA_CHUA_XE_MAY);
  const ptBanLe = sumRevenue(unitOrders, P.PHU_TUNG);
  const scOTo = sumRevenue(unitOrders, P.SUA_CHUA_O_TO);
  const pkOTo = sumRevenue(unitOrders, P.PHU_KIEN_O_TO);
  const tiec = sumRevenue(unitOrders, P.TIEC);
  const luuTru = sumRevenue(unitOrders, P.PHONG_NGHI);
  return {
    gunghoCount: unitOrders.length,
    xeMaySL: countPaidOrders(unitOrders, P.XE_MAY),
    baoHiemXeMay: sumQuantity(unitOrders, P.BAO_HIEM_XE_MAY),
    dvXeMay: scXeMay + ptBanLe,
    soLuotDVXeMay: countPaidOrders(unitOrders, [P.SUA_CHUA_XE_MAY, P.PHU_TUNG]),
    scXeMay,
    ptBanLe,
    oToSL: countPaidOrders(unitOrders, P.O_TO),
    baoHiemOTo: sumQuantity(unitOrders, P.BAO_HIEM_O_TO),
    dvOTo: scOTo + pkOTo,
    soLuotDVOTo: countPaidOrders(unitOrders, [P.SUA_CHUA_O_TO, P.PHU_KIEN_O_TO]),
    scOTo,
    pkOTo,
    dacSan: sumRevenue(unitOrders, P.DAC_SAN),
    nhaHangKhachSan: tiec + luuTru,
    tiec,
    luuTru,
    tour: sumQuantity(unitOrders, P.TOUR),
    veMayBay: sumQuantity(unitOrders, P.VE_MAY_BAY),
    vycSL: countPaidOrders(unitOrders, P.VYC),
    vycDoanhThu: sumRevenue(unitOrders, P.VYC),
    vtnn: sumRevenue(unitOrders, P.VTNN),
  };
}

function metricFields(r) {
  return [
    r.gunghoCount,
    r.xeMaySL, r.baoHiemXeMay, r.dvXeMay, r.soLuotDVXeMay, r.scXeMay, r.ptBanLe,
    r.oToSL, r.baoHiemOTo, r.dvOTo, r.soLuotDVOTo, r.scOTo, r.pkOTo,
    r.dacSan, r.nhaHangKhachSan, r.tiec, r.luuTru, r.tour, r.veMayBay,
    r.vycSL, r.vycDoanhThu, r.vtnn,
  ];
}
function rowToArray(tt, companyCode, unitName, r) {
  return [tt, companyCode, unitName, ...metricFields(r)];
}
function periodRowToArray(tt, label, r) {
  return [tt, label, ...metricFields(r)];
}

// Xuất đúng theo mẫu "KẾT QUẢ GUNG HO CHI TIẾT CÔNG TY THEO ĐƠN VỊ"
function exportGungHoCongTyTemplate({ companyName, orders, fromDate, toDate }) {
  const companyCode = COMPANY_CODE[companyName] || "";
  const branches = COMPANIES.find((c) => c.name === companyName)?.branches || [];

  const inRange = (o) => {
    if (fromDate && new Date(o.createdAt) < new Date(fromDate)) return false;
    if (toDate && new Date(o.createdAt) > new Date(toDate + "T23:59:59")) return false;
    return true;
  };

  const header = [
    [companyCode, "Từ ngày tháng năm", fromDate ? fmtDate(fromDate).split(" ")[0] : "", "", "Đến ngày tháng năm", toDate ? fmtDate(toDate).split(" ")[0] : ""],
    ["KẾT QUẢ GUNG HO CHI TIẾT CÔNG TY THEO ĐƠN VỊ"],
    [
      "", "", "", "",
      "THƯƠNG MẠI 1", "", "", "", "", "",
      "TOYOTA", "", "", "", "", "",
      "HTC", "", "", "", "", "",
      "VYC", "",
      "VTNN",
    ],
    [
      "TT", "CÔNG TY", "ĐƠN VỊ", "SỐ LƯỢNG GUNG HO",
      "SL Xe máy, xe điện", "Bảo hiểm xe máy", "Doanh thu DV xe máy", "Số lượt DV xe máy", "Sửa chữa, bảo dưỡng XM", "Phụ tùng bán lẻ",
      "SL Ô tô", "Bảo hiểm ô tô", "Doanh thu DV ô tô", "Số lượt DV ô tô", "Sửa chữa, bảo dưỡng ô tô", "Phụ kiện, chăm sóc xe",
      "Đặc sản địa phương", "Nhà hàng, khách sạn", "Doanh thu nhà hàng, tiệc", "Doanh thu lưu trú khách sạn", "Tour", "Vé máy bay",
      "Số lượng VYC", "Doanh thu VYC",
      "Doanh thu phân bón",
    ],
  ];

  const branchRows = branches.map((b, i) => {
    const allForUnit = orders.filter((o) => o.store === b.name).filter(inRange);
    return rowToArray(i + 1, companyCode, b.name, buildUnitRow(allForUnit));
  });

  const allForCompany = orders.filter((o) => o.company === companyName).filter(inRange);
  const companyTotalRow = rowToArray("", companyCode, "Công ty", buildUnitRow(allForCompany));

  const aoa = [...header, ...branchRows, companyTotalRow];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!merges"] = [
    { s: { r: 1, c: 0 }, e: { r: 1, c: 24 } },
    { s: { r: 2, c: 4 }, e: { r: 2, c: 9 } },
    { s: { r: 2, c: 10 }, e: { r: 2, c: 15 } },
    { s: { r: 2, c: 16 }, e: { r: 2, c: 21 } },
    { s: { r: 2, c: 22 }, e: { r: 2, c: 23 } },
  ];
  ws["!cols"] = Array.from({ length: 25 }, (_, i) => ({ wch: i < 3 ? 16 : 14 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Chi tiết theo đơn vị");
  XLSX.writeFile(wb, `KetQuaGungHo_${companyCode}_TheoDonVi_${Date.now()}.xlsx`);
}

// Xuất đúng theo mẫu "KẾT QUẢ GUNG HO CHI TIẾT CÔNG TY THEO THỜI GIAN"
function exportGungHoThoiGianTemplate({ unitName, orders, year }) {
  const branch = branchInfo(unitName);
  const companyCode = COMPANY_CODE[branch?.company] || "";
  const unitOrders = orders.filter((o) => o.store === unitName);

  const header = [
    ["Đơn vị", unitName],
    [companyCode, "Từ ngày tháng năm", `01/01/${year}`, "", "Đến ngày tháng năm", `31/12/${year}`],
    ["KẾT QUẢ GUNG HO CHI TIẾT CÔNG TY THEO THỜI GIAN"],
    [
      "", "", "",
      "THƯƠNG MẠI 1", "", "", "", "", "",
      "TOYOTA", "", "", "", "", "",
      "HTC", "", "", "", "", "",
      "VYC", "",
      "VTNN",
    ],
    [
      "TT", "CÔNG TY", "SỐ LƯỢNG GUNG HO",
      "SL Xe máy, xe điện", "Bảo hiểm xe máy", "Doanh thu DV xe máy", "Số lượt DV xe máy", "Sửa chữa, bảo dưỡng XM", "Phụ tùng bán lẻ",
      "SL Ô tô", "Bảo hiểm ô tô", "Doanh thu DV ô tô", "Số lượt DV ô tô", "Sửa chữa, bảo dưỡng ô tô", "Phụ kiện, chăm sóc xe",
      "Đặc sản địa phương", "Nhà hàng, khách sạn", "Doanh thu nhà hàng, tiệc", "Doanh thu lưu trú khách sạn", "Tour", "Vé máy bay",
      "Số lượng VYC", "Doanh thu VYC",
      "Doanh thu phân bón",
    ],
  ];

  const monthRows = [];
  for (let m = 1; m <= 12; m++) {
    const monthOrders = unitOrders.filter((o) => {
      const d = new Date(o.createdAt);
      return d.getFullYear() === Number(year) && d.getMonth() + 1 === m;
    });
    monthRows.push(periodRowToArray(m, `Tháng ${m}`, buildUnitRow(monthOrders)));
  }

  const yearOrders = unitOrders.filter((o) => new Date(o.createdAt).getFullYear() === Number(year));
  const totalRow = periodRowToArray(13, "Tổng", buildUnitRow(yearOrders));

  const aoa = [...header, ...monthRows, totalRow];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!merges"] = [
    { s: { r: 2, c: 0 }, e: { r: 2, c: 23 } },
    { s: { r: 3, c: 3 }, e: { r: 3, c: 8 } },
    { s: { r: 3, c: 9 }, e: { r: 3, c: 14 } },
    { s: { r: 3, c: 15 }, e: { r: 3, c: 20 } },
    { s: { r: 3, c: 21 }, e: { r: 3, c: 22 } },
  ];
  ws["!cols"] = Array.from({ length: 24 }, (_, i) => ({ wch: i < 2 ? 14 : 13 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Chi tiết theo thời gian");
  XLSX.writeFile(wb, `KetQuaGungHo_${unitName.replace(/\s+/g, "")}_${year}_${Date.now()}.xlsx`);
}

// Xuất đúng theo mẫu "THỨ HẠNG GUNG HO"
function exportGungHoRankingTemplate({ companyName, orders, fromDate, toDate }) {
  const companyCode = COMPANY_CODE[companyName] || "";
  const employees = USERS.filter((u) => u.role === "dai_su" && branchInfo(u.store)?.company === companyName);

  const inRange = (o) => {
    if (fromDate && new Date(o.createdAt) < new Date(fromDate)) return false;
    if (toDate && new Date(o.createdAt) > new Date(toDate + "T23:59:59")) return false;
    return true;
  };

  const ranked = employees
    .map((emp) => {
      const empOrders = orders.filter((o) => o.createdBy === emp.id).filter(inRange);
      const r = buildUnitRow(empOrders);
      const td = computeTD(empOrders);
      return { emp, r, td };
    })
    .sort((a, b) => b.td - a.td);

  const header = [
    ["CÔNG TY", companyCode],
    ["THỨ HẠNG GUNG HO"],
    [
      "", "", "", "", "",
      "THƯƠNG MẠI 1", "", "", "",
      "TOYOTA", "", "", "",
      "HTC", "", "", "", "", "",
      "VYC",
      "VTNN",
    ],
    [
      "TT", "Nhân viên", "Đơn vị", "Tổng điểm Gung Ho", "Thứ hạng",
      "Bảo hiểm xe máy", "Doanh thu DV xe máy", "Doanh thu DV sửa chữa xe máy", "Phụ tùng bán lẻ xe máy",
      "Bảo hiểm ô tô", "Doanh thu DV ô tô", "Sửa chữa, bảo dưỡng ô tô", "Phụ kiện, chăm sóc xe",
      "Đặc sản địa phương", "Nhà hàng, khách sạn", "Doanh thu nhà hàng, tiệc", "Doanh thu lưu trú khách sạn", "Tour", "Vé máy bay",
      "Doanh thu VYC",
      "Doanh thu phân bón",
    ],
  ];

  const dataRows = ranked.map(({ emp, r, td }, i) => [
    i + 1, emp.name, emp.store || "", td, i + 1,
    r.baoHiemXeMay, r.dvXeMay, r.scXeMay, r.ptBanLe,
    r.baoHiemOTo, r.dvOTo, r.scOTo, r.pkOTo,
    r.dacSan, r.nhaHangKhachSan, r.tiec, r.luuTru, r.tour, r.veMayBay,
    r.vycDoanhThu,
    r.vtnn,
  ]);

  const aoa = [...header, ...dataRows];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!merges"] = [
    { s: { r: 1, c: 0 }, e: { r: 1, c: 20 } },
    { s: { r: 2, c: 5 }, e: { r: 2, c: 8 } },
    { s: { r: 2, c: 9 }, e: { r: 2, c: 12 } },
    { s: { r: 2, c: 13 }, e: { r: 2, c: 18 } },
  ];
  ws["!cols"] = Array.from({ length: 21 }, (_, i) => ({ wch: i < 3 ? 16 : 14 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Thứ hạng Gung Ho");
  XLSX.writeFile(wb, `ThuHangGungHo_${companyCode}_${Date.now()}.xlsx`);
}

// Danh sách nhân viên được nạp thật từ bảng `employees` trên Supabase khi ứng dụng khởi động
// (mảng này được cập nhật (mutate) trong App, không import lại) — giữ nguyên tên USERS
// để toàn bộ phần code còn lại (vốn tham chiếu USERS trực tiếp) không cần sửa thêm.
let USERS = [];

const ROLE_META = {
  dai_su: { label: "Đại sứ Gungho", short: "Đại sứ", color: "bg-teal-50 text-teal-700 border-teal-200" },
  xu_ly: { label: "Nhân viên xử lý - chăm sóc", short: "Xử lý CSKH", color: "bg-indigo-50 text-indigo-700 border-indigo-200" },
  cht: { label: "Trưởng đơn vị", short: "Trưởng đơn vị", color: "bg-amber-50 text-amber-800 border-amber-200" },
  ke_toan: { label: "Kế toán", short: "Kế toán", color: "bg-rose-50 text-rose-700 border-rose-200" },
  admin: { label: "Quản trị hệ thống", short: "Admin", color: "bg-slate-800 text-white border-slate-800" },
};

const STATUS_META = {
  cho_phan_cong: { label: "Chờ xác nhận", color: "bg-slate-100 text-slate-600 border-slate-300" },
  cho_xu_ly: { label: "Chờ xác nhận", color: "bg-slate-100 text-slate-600 border-slate-300" },
  dang_cham_soc: { label: "Chờ thanh toán", color: "bg-amber-50 text-amber-700 border-amber-300" },
  cho_ke_toan: { label: "Chờ thanh toán", color: "bg-amber-50 text-amber-700 border-amber-300" },
  da_thanh_toan: { label: "Đơn hàng đã được ghi nhận", color: "bg-emerald-50 text-emerald-700 border-emerald-300" },
  khong_thanh_toan: { label: "Không thanh toán", color: "bg-rose-50 text-rose-700 border-rose-300" },
};

// Quy định riêng khối HTC: đơn hàng chờ thanh toán quá số ngày này sẽ tự động
// đóng khi tài khoản kế toán/admin mở app — ghi nhận doanh thu & doanh số
// Gungho cho Đại sứ, nhưng KHÔNG ghi nhận hoa hồng Gungho.
const HTC_COMPANY_NAME = "III. HTC";
const HTC_AUTO_CLOSE_DAYS = 30;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const fmtMoney = (n) => (Number(n) || 0).toLocaleString("vi-VN") + " đ";
const fmtDate = (iso) =>
  new Date(iso).toLocaleString("vi-VN", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
const uid = (p) => `${p}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`;

function genOrderCode(existingOrders) {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const prefix = `ORD${yy}${mm}${dd}`;
  const todayCount = existingOrders.filter((o) => o.orderCode?.startsWith(prefix)).length;
  return `${prefix}${String(todayCount + 1).padStart(4, "0")}`;
}

// Mã khách hàng riêng theo từng Đại sứ: <mã nhân viên>-<số thứ tự>, ví dụ ds1-001
function genCustomerCode(existingCustomers, employeeCode) {
  const prefix = `${employeeCode || "KH"}-`;
  const count = existingCustomers.filter((c) => c.customerCode?.startsWith(prefix)).length;
  return `${prefix}${String(count + 1).padStart(3, "0")}`;
}
const userById = (id) => USERS.find((u) => u.id === id);

// ---- Supabase data mapping (DB dùng snake_case, app dùng camelCase) ----

function mapEmployee(e) {
  return { id: e.id, employeeCode: e.employee_code, name: e.name, role: e.role, store: e.store, position: e.position };
}
function mapCustomer(c) {
  return {
    id: c.id, customerCode: c.customer_code, name: c.name, phone: c.phone, address: c.address,
    createdBy: c.created_by, createdByName: c.created_by_name, createdAt: c.created_at,
  };
}
function mapOrder(o) {
  return {
    id: o.id, orderCode: o.order_code, customerId: o.customer_id,
    customerName: o.customer_name, customerPhone: o.customer_phone, customerCode: o.customer_code,
    company: o.company, store: o.store, storeAddress: o.store_address,
    product: o.product, quantity: o.quantity,
    totalAmount: Number(o.total_amount) || 0,
    discountAmount: Number(o.discount_amount) || 0,
    finalAmount: o.final_amount === null || o.final_amount === undefined ? undefined : Number(o.final_amount),
    commissionAmount: Number(o.commission_amount) || 0,
    invoiceName: o.invoice_name || "", invoiceNumber: o.invoice_number || "",
    createdBy: o.created_by, createdByName: o.created_by_name,
    assignedHandler: o.assigned_handler, assignedHandlerName: o.assigned_handler_name,
    status: o.status, handlerNote: o.handler_note, accountantNote: o.accountant_note,
    history: o.history || [], createdAt: o.created_at, updatedAt: o.updated_at,
  };
}
function mapNotification(n) {
  return { id: n.id, toUserId: n.to_employee_id, message: n.message, orderId: n.order_id, read: n.read, createdAt: n.created_at };
}

function orderToRow(o) {
  return {
    order_code: o.orderCode, customer_id: o.customerId, customer_name: o.customerName, customer_phone: o.customerPhone, customer_code: o.customerCode,
    company: o.company, store: o.store, store_address: o.storeAddress, product: o.product,
    quantity: o.quantity ?? 0, total_amount: o.totalAmount ?? 0, discount_amount: o.discountAmount ?? 0,
    final_amount: o.finalAmount ?? null, commission_amount: o.commissionAmount ?? 0,
    created_by: o.createdBy, created_by_name: o.createdByName,
    assigned_handler: o.assignedHandler || null, assigned_handler_name: o.assignedHandlerName || null,
    status: o.status, handler_note: o.handlerNote || "", accountant_note: o.accountantNote || "",
    history: o.history || [], updated_at: new Date().toISOString(),
  };
}

function mapAnnouncement(a) {
  return {
    id: a.id, title: a.title, content: a.content,
    targetRoles: a.target_roles || ["all"], isPinned: !!a.is_pinned,
    createdBy: a.created_by, createdByName: a.created_by_name,
    createdAt: a.created_at, updatedAt: a.updated_at,
  };
}

async function fetchAll() {
  const [emp, cust, ord, notif, announ] = await Promise.all([
    supabase.from("employees").select("*"),
    supabase.from("customers").select("*").order("created_at", { ascending: false }),
    supabase.from("orders").select("*").order("created_at", { ascending: false }),
    supabase.from("notifications").select("*").order("created_at", { ascending: false }),
    supabase.from("announcements").select("*").order("created_at", { ascending: false }),
  ]);
  if (emp.error) console.error("fetch employees error", emp.error);
  if (cust.error) console.error("fetch customers error", cust.error);
  if (ord.error) console.error("fetch orders error", ord.error);
  if (notif.error) console.error("fetch notifications error", notif.error);
  if (announ.error) console.error("fetch announcements error", announ.error);
  return {
    employees: (emp.data || []).map(mapEmployee),
    customers: (cust.data || []).map(mapCustomer),
    orders: (ord.data || []).map(mapOrder),
    notifications: (notif.data || []).map(mapNotification),
    announcements: (announ.data || []).map(mapAnnouncement),
  };
}

const SESSION_KEY = "gungho_session_employee_id";

// ---------------------------------------------------------------------------
// Small UI atoms
// ---------------------------------------------------------------------------

function Badge({ children, className = "" }) {
  return (
    <span className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-medium border whitespace-nowrap ${className}`}>
      {children}
    </span>
  );
}

function StatusBadge({ status }) {
  const m = STATUS_META[status] || STATUS_META.cho_phan_cong;
  return <Badge className={m.color}>{m.label}</Badge>;
}

function Card({ children, className = "" }) {
  return <div className={`bg-white rounded-2xl border border-slate-200/80 shadow-sm hover:shadow-md transition-shadow duration-200 ${className}`}>{children}</div>;
}

function SectionTitle({ icon: Icon, title, subtitle }) {
  return (
    <div className="flex items-start gap-3 mb-4">
      <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-teal-700 to-teal-900 text-white flex items-center justify-center shrink-0 shadow-sm shadow-teal-900/20">
        <Icon size={19} />
      </div>
      <div>
        <h2 className="text-lg font-semibold text-slate-800 tracking-tight">{title}</h2>
        {subtitle && <p className="text-sm text-slate-500">{subtitle}</p>}
      </div>
    </div>
  );
}

function EmptyState({ icon: Icon, text }) {
  return (
    <div className="flex flex-col items-center justify-center py-14 text-slate-400">
      <div className="w-14 h-14 rounded-2xl bg-slate-50 flex items-center justify-center mb-3">
        <Icon size={26} className="opacity-50" />
      </div>
      <p className="text-sm">{text}</p>
    </div>
  );
}

function MetricCard({ label, value, icon: Icon, accent = "teal" }) {
  const accents = {
    teal: { text: "text-teal-700", bg: "bg-teal-50", ring: "ring-teal-600/10", bar: "bg-teal-600" },
    amber: { text: "text-amber-700", bg: "bg-amber-50", ring: "ring-amber-600/10", bar: "bg-amber-600" },
    indigo: { text: "text-indigo-700", bg: "bg-indigo-50", ring: "ring-indigo-600/10", bar: "bg-indigo-600" },
    rose: { text: "text-rose-700", bg: "bg-rose-50", ring: "ring-rose-600/10", bar: "bg-rose-600" },
  };
  const a = accents[accent] || accents.teal;
  return (
    <div className="relative bg-white rounded-2xl border border-slate-200/80 shadow-sm hover:shadow-md transition-shadow duration-200 p-4 flex items-center gap-3 overflow-hidden">
      <span className={`absolute left-0 top-0 bottom-0 w-1 ${a.bar}`} />
      <div className={`w-11 h-11 rounded-xl flex items-center justify-center shrink-0 ring-4 ${a.bg} ${a.text} ${a.ring}`}>
        <Icon size={19} />
      </div>
      <div className="min-w-0">
        <p className="text-xs text-slate-500 truncate">{label}</p>
        <p className="text-lg font-semibold text-slate-800 truncate tracking-tight">{value}</p>
      </div>
    </div>
  );
}

function TextField({ label, ...props }) {
  return (
    <label className="block">
      <span className="block text-xs font-medium text-slate-600 mb-1">{label}</span>
      <input
        {...props}
        className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm transition focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-600"
      />
    </label>
  );
}

// ---------------------------------------------------------------------------
// Chuyển số tiền sang chữ tiếng Việt (hiển thị dưới các ô nhập tiền)
// ---------------------------------------------------------------------------
const CHU_SO_VN = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];
const DON_VI_NHOM_VN = ["", "nghìn", "triệu", "tỷ"];

function docNhom3VN(so) {
  const tram = Math.floor(so / 100);
  const chuc = Math.floor((so % 100) / 10);
  const donvi = so % 10;
  let s = "";
  if (tram > 0) {
    s += CHU_SO_VN[tram] + " trăm";
    if (chuc === 0 && donvi > 0) s += " linh";
  }
  if (chuc > 1) {
    s += (s ? " " : "") + CHU_SO_VN[chuc] + " mươi";
    if (donvi === 1) s += " mốt";
    else if (donvi === 5) s += " lăm";
    else if (donvi > 0) s += " " + CHU_SO_VN[donvi];
  } else if (chuc === 1) {
    s += (s ? " " : "") + "mười";
    if (donvi === 1) s += " một";
    else if (donvi === 5) s += " lăm";
    else if (donvi > 0) s += " " + CHU_SO_VN[donvi];
  } else if (chuc === 0 && donvi > 0) {
    s += (s ? " " : "") + CHU_SO_VN[donvi];
  }
  return s.trim();
}

function soTienThanhChu(n) {
  let so = Math.round(Number(n) || 0);
  if (so === 0) return "";
  const amDau = so < 0;
  so = Math.abs(so);
  const nhom = [];
  while (so > 0) {
    nhom.unshift(so % 1000);
    so = Math.floor(so / 1000);
  }
  const total = nhom.length;
  const parts = [];
  nhom.forEach((g, i) => {
    if (g === 0) return;
    const bac = total - i - 1;
    let chu = docNhom3VN(g);
    if (bac > 0 && bac <= 3) chu += " " + DON_VI_NHOM_VN[bac];
    parts.push(chu);
  });
  let result = parts.join(" ").replace(/\s+/g, " ").trim();
  result = result.charAt(0).toUpperCase() + result.slice(1);
  return (amDau ? "Âm " : "") + result + " đồng";
}

// Ô nhập số tiền: tự hiện dấu phẩy ngăn cách hàng nghìn khi gõ, có dòng chữ
// nhỏ đọc số tiền bằng chữ bên dưới để nhân viên dễ kiểm tra.
function MoneyField({ label, value, onChange, placeholder = "0" }) {
  const numericValue = value === "" || value === null || value === undefined ? "" : Number(value);
  const displayValue = numericValue === "" || Number.isNaN(numericValue) ? "" : numericValue.toLocaleString("vi-VN");
  const words = numericValue !== "" && numericValue > 0 ? soTienThanhChu(numericValue) : "";

  const handleChange = (e) => {
    const digitsOnly = e.target.value.replace(/[^\d]/g, "");
    onChange(digitsOnly);
  };

  return (
    <label className="block">
      <span className="block text-xs font-medium text-slate-600 mb-1">{label}</span>
      <input
        type="text"
        inputMode="numeric"
        value={displayValue}
        onChange={handleChange}
        placeholder={placeholder}
        className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm transition focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-600"
      />
      {words && <span className="block text-[11px] text-slate-400 mt-1 italic">{words}</span>}
    </label>
  );
}


function SelectField({ label, children, ...props }) {
  return (
    <label className="block">
      <span className="block text-xs font-medium text-slate-600 mb-1">{label}</span>
      <select
        {...props}
        className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm bg-white transition focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-600"
      >
        {children}
      </select>
    </label>
  );
}

function TextAreaField({ label, ...props }) {
  return (
    <label className="block">
      <span className="block text-xs font-medium text-slate-600 mb-1">{label}</span>
      <textarea
        {...props}
        rows={2}
        className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm transition focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-600"
      />
    </label>
  );
}

function PrimaryButton({ children, className = "", ...props }) {
  return (
    <button
      {...props}
      className={`inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-gradient-to-b from-teal-700 to-teal-800 text-white text-sm font-medium shadow-sm shadow-teal-900/20 hover:from-teal-800 hover:to-teal-900 active:scale-[0.98] transition disabled:opacity-40 disabled:cursor-not-allowed ${className}`}
    >
      {children}
    </button>
  );
}
function GhostButton({ children, className = "", ...props }) {
  return (
    <button
      {...props}
      className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl border border-slate-300 bg-white text-slate-700 text-sm font-medium hover:bg-slate-50 hover:border-slate-400 active:scale-[0.98] transition disabled:opacity-40 ${className}`}
    >
      {children}
    </button>
  );
}
function DangerButton({ children, className = "", ...props }) {
  return (
    <button
      {...props}
      className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl border border-rose-300 bg-white text-rose-700 text-sm font-medium hover:bg-rose-50 active:scale-[0.98] transition disabled:opacity-40 ${className}`}
    >
      {children}
    </button>
  );
}

function Toast({ toast }) {
  if (!toast) return null;
  return (
    <div className="fixed bottom-5 right-5 z-50 bg-slate-900 text-white text-sm px-4 py-2.5 rounded-xl shadow-lg shadow-slate-900/30 animate-[fadeIn_.15s_ease-out]" style={{ marginBottom: "env(safe-area-inset-bottom)" }}>
      {toast}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Login screen
// ---------------------------------------------------------------------------

function LoginScreen({ onLogin }) {
  const [code, setCode] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async () => {
    if (!code.trim() || !password) {
      setError("Vui lòng nhập mã nhân viên và mật khẩu.");
      return;
    }
    setError("");
    setLoading(true);
    try {
      const { data, error: qErr } = await supabase
        .from("employees")
        .select("*")
        .eq("employee_code", code.trim())
        .eq("password", password)
        .maybeSingle();
      if (qErr) throw qErr;
      if (!data) {
        setError("Mã nhân viên hoặc mật khẩu không đúng.");
        return;
      }
      onLogin(mapEmployee(data));
    } catch (e) {
      console.error(e);
      setError("Không kết nối được máy chủ, vui lòng thử lại.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-[600px] flex items-center justify-center bg-gradient-to-b from-teal-50 via-slate-50 to-slate-50 px-4 py-10">
      <div className="w-full max-w-sm">
        <div className="text-center mb-8">
          <img src="/logo.png" alt="Phú Tài Đức Group" className="h-16 mx-auto mb-3 object-contain" />
          <h1 className="text-2xl font-semibold text-slate-800 tracking-tight">GUNGHO PTD</h1>
          <p className="text-sm text-slate-500 mt-1">Đăng nhập bằng mã nhân viên</p>
        </div>
        <div className="bg-white rounded-2xl border border-slate-200/80 shadow-xl shadow-slate-900/5 p-5 space-y-4">
          <TextField label="Mã nhân viên" value={code} onChange={(e) => setCode(e.target.value)} placeholder="Ví dụ: ds1" onKeyDown={(e) => e.key === "Enter" && handleSubmit()} />
          <TextField label="Mật khẩu" type="password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••" onKeyDown={(e) => e.key === "Enter" && handleSubmit()} />
          {error && (
            <p className="text-sm text-rose-600 flex items-center gap-1.5"><AlertCircle size={14} /> {error}</p>
          )}
          <PrimaryButton onClick={handleSubmit} disabled={loading} className="w-full justify-center">
            {loading ? "Đang đăng nhập..." : "Đăng nhập"}
          </PrimaryButton>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Notification bell
// ---------------------------------------------------------------------------

function NotifBell({ notifications, currentUser, onMarkRead }) {
  const [open, setOpen] = useState(false);
  const mine = notifications.filter((n) => n.toUserId === currentUser.id).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  const unread = mine.filter((n) => !n.read).length;

  return (
    <div className="relative">
      <button
        onClick={() => setOpen((o) => !o)}
        className="relative w-9 h-9 rounded-xl border border-slate-200 flex items-center justify-center hover:bg-slate-50 transition"
      >
        <Bell size={17} className="text-slate-600" />
        {unread > 0 && (
          <span className="absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-rose-600 text-white text-[10px] font-semibold flex items-center justify-center">
            {unread > 9 ? "9+" : unread}
          </span>
        )}
      </button>
      {open && (
        <>
          <div className="fixed inset-0 z-10" onClick={() => setOpen(false)} />
          <div className="absolute right-0 mt-2 w-80 max-h-96 overflow-y-auto bg-white rounded-xl border border-slate-200 shadow-lg z-20">
            <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between">
              <p className="text-sm font-semibold text-slate-800">Thông báo</p>
              {unread > 0 && (
                <button
                  className="text-xs text-teal-700 hover:underline"
                  onClick={() => onMarkRead(mine.map((n) => n.id))}
                >
                  Đánh dấu đã đọc hết
                </button>
              )}
            </div>
            {mine.length === 0 ? (
              <div className="py-10 text-center text-sm text-slate-400">Chưa có thông báo nào</div>
            ) : (
              mine.map((n) => (
                <div
                  key={n.id}
                  onClick={() => !n.read && onMarkRead([n.id])}
                  className={`px-4 py-3 border-b border-slate-50 last:border-0 cursor-pointer ${!n.read ? "bg-teal-50/50" : ""}`}
                >
                  <p className={`text-sm ${!n.read ? "font-medium text-slate-800" : "text-slate-600"}`}>{n.message}</p>
                  <p className="text-[11px] text-slate-400 mt-1">{fmtDate(n.createdAt)}</p>
                </div>
              ))
            )}
          </div>
        </>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Main App
// ---------------------------------------------------------------------------

export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [customers, setCustomers] = useState([]);
  const [orders, setOrders] = useState([]);
  const [notifications, setNotifications] = useState([]);
  const [announcements, setAnnouncements] = useState([]);
  const [tab, setTab] = useState(null);
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState(null);
  const [, forceRerender] = useState(0);

  const showToast = useCallback((msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 2500);
  }, []);

  const refreshAll = useCallback(async () => {
    const data = await fetchAll();
    const closedCount = await autoCloseOverdueHTCOrders(data.orders);
    const finalData = closedCount > 0 ? await fetchAll() : data;
    USERS.length = 0;
    USERS.push(...finalData.employees);
    setCustomers(finalData.customers);
    setOrders(finalData.orders);
    setNotifications(finalData.notifications);
    setAnnouncements(finalData.announcements);
    forceRerender((n) => n + 1);
    if (closedCount > 0) {
      showToast(`Đã tự động đóng ${closedCount} đơn hàng HTC quá 30 ngày chưa thanh toán`);
    }
    return finalData;
  }, []);

  useEffect(() => {
    (async () => {
      setLoading(true);
      const data = await refreshAll();
      const savedId = localStorage.getItem(SESSION_KEY);
      if (savedId) {
        const found = data.employees.find((e) => e.id === savedId);
        if (found) setCurrentUser(found);
        else localStorage.removeItem(SESSION_KEY);
      }
      setLoading(false);
    })();
  }, [refreshAll]);

  useEffect(() => {
    if (!currentUser) return;
    const interval = setInterval(() => {
      refreshAll();
    }, 15000); // tự động làm mới dữ liệu mỗi 15 giây
    return () => clearInterval(interval);
  }, [currentUser, refreshAll]);

  useEffect(() => {
    if (!currentUser) return;
    const defaults = { dai_su: "khach_hang", xu_ly: "duoc_giao", cht: "phan_cong", ke_toan: "cho_xac_nhan", admin: "bao_cao_cht" };
    setTab(defaults[currentUser.role]);
  }, [currentUser]);

  const handleLogin = (employee) => {
    setCurrentUser(employee);
    localStorage.setItem(SESSION_KEY, employee.id);
  };
  const handleLogout = () => {
    setCurrentUser(null);
    localStorage.removeItem(SESSION_KEY);
  };

  const notifRow = (toUserId, message, orderId) => ({
    to_employee_id: toUserId,
    message,
    order_id: orderId || null,
  });

  const insertNotifications = async (rows) => {
    if (rows.length === 0) return;
    const { error } = await supabase.from("notifications").insert(rows);
    if (error) console.error("insert notifications error", error);
  };

  const markRead = async (ids) => {
    setNotifications((prev) => prev.map((n) => (ids.includes(n.id) ? { ...n, read: true } : n)));
    const { error } = await supabase.from("notifications").update({ read: true }).in("id", ids);
    if (error) console.error("mark read error", error);
  };

  // ---- action handlers -----------------------------------------------

  const addCustomer = async ({ name, phone, address }) => {
    const myCustomers = customers.filter((c) => c.createdBy === currentUser.id);
    const customerCode = genCustomerCode(myCustomers, currentUser.employeeCode);
    const { data, error } = await supabase
      .from("customers")
      .insert({ name, phone, address, customer_code: customerCode, created_by: currentUser.id, created_by_name: currentUser.name })
      .select()
      .single();
    if (error) {
      console.error("addCustomer error", error);
      showToast("Không thêm được khách hàng, vui lòng thử lại");
      throw error;
    }
    const c = mapCustomer(data);
    setCustomers((prev) => [c, ...prev]);
    showToast("Đã thêm khách hàng mới");
    return c;
  };

  const addAnnouncement = async ({ title, content, targetRoles, isPinned }) => {
    const { data, error } = await supabase
      .from("announcements")
      .insert({
        title, content, target_roles: targetRoles?.length ? targetRoles : ["all"], is_pinned: !!isPinned,
        created_by: currentUser.id, created_by_name: currentUser.name,
      })
      .select()
      .single();
    if (error) {
      console.error("addAnnouncement error", error);
      showToast("Không đăng được thông báo, vui lòng thử lại");
      throw error;
    }
    setAnnouncements((prev) => [mapAnnouncement(data), ...prev]);
    showToast("Đã đăng thông báo mới");
  };

  const updateAnnouncement = async (id, { title, content, targetRoles, isPinned }) => {
    const { error } = await supabase
      .from("announcements")
      .update({ title, content, target_roles: targetRoles?.length ? targetRoles : ["all"], is_pinned: !!isPinned, updated_at: new Date().toISOString() })
      .eq("id", id);
    if (error) {
      console.error("updateAnnouncement error", error);
      showToast("Không cập nhật được thông báo, vui lòng thử lại");
      throw error;
    }
    await refreshAll();
    showToast("Đã cập nhật thông báo");
  };

  const deleteAnnouncement = async (id) => {
    const { error } = await supabase.from("announcements").delete().eq("id", id);
    if (error) {
      console.error("deleteAnnouncement error", error);
      showToast("Không xoá được thông báo, vui lòng thử lại");
      throw error;
    }
    setAnnouncements((prev) => prev.filter((a) => a.id !== id));
    showToast("Đã xoá thông báo");
  };

  const createOrder = async ({ customerId, company, store, storeAddress, product, handlerId }) => {
    const cust = customers.find((c) => c.id === customerId);
    const handler = handlerId ? userById(handlerId) : null;
    const status = handler ? "cho_xu_ly" : "cho_phan_cong";
    const orderDraft = {
      orderCode: genOrderCode(orders),
      customerId, customerName: cust?.name, customerPhone: cust?.phone, customerCode: cust?.customerCode,
      company, store, storeAddress, product, totalAmount: 0,
      createdBy: currentUser.id, createdByName: currentUser.name,
      assignedHandler: handler?.id || null, assignedHandlerName: handler?.name || null,
      status, handlerNote: "", discountAmount: 0, commissionAmount: 0, accountantNote: "",
      history: [`${fmtDate(new Date().toISOString())} — ${currentUser.name} tạo đơn hàng`],
    };
    const { data, error } = await supabase.from("orders").insert(orderToRow(orderDraft)).select().single();
    if (error) {
      console.error("createOrder error", error);
      showToast("Không tạo được đơn hàng, vui lòng thử lại");
      throw error;
    }
    const order = mapOrder(data);
    setOrders((prev) => [order, ...prev]);

    let notifRows = [];
    if (handler) {
      notifRows.push(notifRow(handler.id, `Bạn được giao chăm sóc đơn hàng của khách "${cust?.name}" (${product}).`, order.id));
    } else {
      const storeManagers = USERS.filter((u) => u.role === "cht" && u.store === store);
      const targets = storeManagers.length > 0 ? storeManagers : USERS.filter((u) => u.role === "cht");
      targets.forEach((cht) => {
        notifRows.push(notifRow(cht.id, `Có đơn hàng mới tại "${store}" (${company}) cần phân công nhân sự xử lý (khách "${cust?.name}").`, order.id));
      });
    }
    await insertNotifications(notifRows);
    await refreshAll();
    showToast("Đã tạo đơn hàng");
  };

  const updateOrder = async (orderId, patch) => {
    const { error } = await supabase.from("orders").update(patch).eq("id", orderId);
    if (error) {
      console.error("updateOrder error", error);
      showToast("Cập nhật đơn hàng thất bại, vui lòng thử lại");
      throw error;
    }
  };

  const assignHandler = async (orderId, handlerId, note) => {
    const handler = userById(handlerId);
    const order = orders.find((o) => o.id === orderId);
    const newHistory = [...order.history, `${fmtDate(new Date().toISOString())} — ${currentUser.name} phân công cho ${handler.name}`];
    await updateOrder(orderId, {
      assigned_handler: handler.id, assigned_handler_name: handler.name, status: "cho_xu_ly",
      history: newHistory, updated_at: new Date().toISOString(),
    });
    await insertNotifications([
      notifRow(handler.id, `Bạn được phân công chăm sóc đơn hàng của khách "${order.customerName}" (${order.product}).`, orderId),
      notifRow(order.createdBy, `Đơn hàng của khách "${order.customerName}" đã được ${ROLE_META.cht.short} phân công cho ${handler.name}.`, orderId),
    ]);
    await refreshAll();
    showToast("Đã phân công nhân sự xử lý");
  };

  const confirmHandling = async (orderId, note) => {
    const order = orders.find((o) => o.id === orderId);
    const newHistory = [...order.history, `${fmtDate(new Date().toISOString())} — ${currentUser.name} xác nhận đang chăm sóc khách hàng`];
    await updateOrder(orderId, { status: "dang_cham_soc", handler_note: note, history: newHistory, updated_at: new Date().toISOString() });
    await insertNotifications([
      notifRow(order.createdBy, `${currentUser.name} đã xác nhận đang chăm sóc khách "${order.customerName}".`, orderId),
    ]);
    await refreshAll();
    showToast("Đã xác nhận chăm sóc");
  };

  const forwardToAccounting = async (orderId, note, htcInfo) => {
    const order = orders.find((o) => o.id === orderId);
    const historyLine = htcInfo
      ? `${fmtDate(new Date().toISOString())} — ${currentUser.name} chuyển đơn cho kế toán (khách đồng ý mua). Số tiền: ${fmtMoney(htcInfo.amount)} · Chiết khấu: ${fmtMoney(htcInfo.discountAmount)} · Tên xuất HĐ: ${htcInfo.invoiceName} · Số HĐ: ${htcInfo.invoiceNumber}`
      : `${fmtDate(new Date().toISOString())} — ${currentUser.name} chuyển đơn cho kế toán (khách đồng ý mua)`;
    const newHistory = [...order.history, historyLine];
    const patch = { status: "cho_ke_toan", handler_note: note, history: newHistory, updated_at: new Date().toISOString() };
    if (htcInfo) {
      patch.total_amount = htcInfo.amount;
      patch.discount_amount = htcInfo.discountAmount;
      patch.invoice_name = htcInfo.invoiceName;
      patch.invoice_number = htcInfo.invoiceNumber;
    }
    await updateOrder(orderId, patch);
    const storeAccountants = USERS.filter((u) => u.role === "ke_toan" && u.store === order.store);
    const targets = storeAccountants.length > 0 ? storeAccountants : USERS.filter((u) => u.role === "ke_toan");
    await insertNotifications([
      notifRow(order.createdBy, `Khách "${order.customerName}" đồng ý mua hàng — đơn đã được chuyển kế toán xử lý.`, orderId),
      ...targets.map((kt) => notifRow(kt.id, `Có đơn hàng mới cần xác nhận thanh toán (khách "${order.customerName}").`, orderId)),
    ]);
    await refreshAll();
    showToast("Đã chuyển đơn cho kế toán");
  };

  const declineOrder = async (orderId, note) => {
    const order = orders.find((o) => o.id === orderId);
    const newHistory = [...order.history, `${fmtDate(new Date().toISOString())} — ${currentUser.name} ghi nhận khách không mua hàng`];
    await updateOrder(orderId, { status: "khong_thanh_toan", handler_note: note, history: newHistory, updated_at: new Date().toISOString() });
    await insertNotifications([
      notifRow(order.createdBy, `Khách "${order.customerName}" không mua hàng. Ghi chú: ${note || "(không có)"}`, orderId),
    ]);
    await refreshAll();
    showToast("Đã cập nhật trạng thái đơn hàng");
  };

  const confirmPayment = async (orderId, { amount, discountAmount, commissionAmount, quantity, note }) => {
    const order = orders.find((o) => o.id === orderId);
    const totalAmount = Number(amount) || 0;
    const finalAmount = Math.max(totalAmount - Number(discountAmount || 0), 0);
    const newHistory = [
      ...order.history,
      `${fmtDate(new Date().toISOString())} — ${currentUser.name} xác nhận thanh toán ${fmtMoney(totalAmount)}, cập nhật hoa hồng ${fmtMoney(commissionAmount)}`,
    ];
    await updateOrder(orderId, {
      status: "da_thanh_toan", total_amount: totalAmount, discount_amount: Number(discountAmount) || 0,
      final_amount: finalAmount, commission_amount: Number(commissionAmount) || 0, quantity: Number(quantity) || 1,
      accountant_note: note, history: newHistory, updated_at: new Date().toISOString(),
    });
    await insertNotifications([
      notifRow(order.createdBy, `Đơn hàng của khách "${order.customerName}" đã được kế toán xác nhận thanh toán. Hoa hồng: ${fmtMoney(commissionAmount)}.`, orderId),
    ]);
    await refreshAll();
    showToast("Đã xác nhận thanh toán");
  };

  const rejectPayment = async (orderId, note) => {
    const order = orders.find((o) => o.id === orderId);
    const newHistory = [...order.history, `${fmtDate(new Date().toISOString())} — ${currentUser.name} (kế toán) ghi nhận không thanh toán`];
    await updateOrder(orderId, { status: "khong_thanh_toan", accountant_note: note, history: newHistory, updated_at: new Date().toISOString() });
    await insertNotifications([
      notifRow(order.createdBy, `Đơn hàng của khách "${order.customerName}" không được thanh toán. Lý do: ${note || "(không có)"}`, orderId),
    ]);
    await refreshAll();
    showToast("Đã cập nhật: không thanh toán");
  };

  // Quét đơn hàng khối HTC đang "Chờ thanh toán" (cho_ke_toan) quá
  // HTC_AUTO_CLOSE_DAYS ngày kể từ lần cập nhật gần nhất → tự động đóng:
  // ghi nhận doanh thu/doanh số Gungho cho Đại sứ nhưng KHÔNG tính hoa hồng.
  const autoCloseOverdueHTCOrders = async (orderList) => {
    const now = Date.now();
    const eligible = orderList.filter((o) => {
      if (o.status !== "cho_ke_toan") return false;
      if (o.company !== HTC_COMPANY_NAME) return false;
      if (!o.updatedAt) return false;
      const daysElapsed = (now - new Date(o.updatedAt).getTime()) / 86400000;
      return daysElapsed > HTC_AUTO_CLOSE_DAYS;
    });
    for (const o of eligible) {
      const newHistory = [
        ...(o.history || []),
        `${fmtDate(new Date().toISOString())} — Hệ thống tự động đóng đơn: quá ${HTC_AUTO_CLOSE_DAYS} ngày chưa thanh toán (khối HTC). Ghi nhận doanh thu & doanh số Gungho, KHÔNG tính hoa hồng.`,
      ];
      try {
        await updateOrder(o.id, {
          status: "da_thanh_toan",
          total_amount: o.totalAmount,
          discount_amount: 0,
          final_amount: o.totalAmount,
          commission_amount: 0,
          accountant_note: [o.accountantNote, `Tự động đóng: quá ${HTC_AUTO_CLOSE_DAYS} ngày chưa thanh toán (HTC) — không tính hoa hồng.`].filter(Boolean).join(" | "),
          history: newHistory,
          updated_at: new Date().toISOString(),
        });
        await insertNotifications([
          notifRow(o.createdBy, `Đơn hàng của khách "${o.customerName}" đã tự động đóng do quá ${HTC_AUTO_CLOSE_DAYS} ngày chưa thanh toán (khối HTC). Đã ghi nhận doanh thu & doanh số, không có hoa hồng.`, o.id),
        ]);
      } catch (e) {
        console.error("auto-close HTC order failed", o.id, e);
      }
    }
    return eligible.length;
  };

  // ---- render -----------------------------------------------------------

  if (loading) {
    return (
      <div className="min-h-[400px] flex items-center justify-center text-slate-400 text-sm gap-2">
        <RefreshCw size={16} className="animate-spin" /> Đang tải dữ liệu...
      </div>
    );
  }

  if (!currentUser) {
    return <LoginScreen onLogin={handleLogin} />;
  }

  const GROUP_BAN_HANG = [
    { key: "khach_hang", label: "Khách hàng", icon: Users },
    { key: "don_hang_ds", label: "Đơn hàng của tôi", icon: ClipboardList },
    { key: "bao_cao_ds", label: "Báo cáo", icon: BarChart3 },
  ];
  const GROUP_CSKH = [
    { key: "duoc_giao", label: "Đơn được giao", icon: Inbox },
    { key: "don_hang_cskh", label: "Đơn hàng CSKH", icon: ClipboardList },
    { key: "bao_cao_cskh", label: "Báo cáo CSKH", icon: BarChart3 },
  ];
  const GROUP_QUAN_LY = [
    { key: "phan_cong", label: "Phân công", icon: ArrowRightLeft },
    { key: "bao_cao_cht", label: "Báo cáo doanh số", icon: BarChart3 },
  ];
  const GROUP_KE_TOAN = [
    { key: "cho_xac_nhan", label: "Chờ xác nhận", icon: ClipboardCheck },
    { key: "lich_su", label: "Lịch sử", icon: Wallet },
  ];
  const GROUP_THONG_BAO = [
    { key: "thong_bao", label: "Thông báo", icon: Megaphone },
  ];
  const NAV_GROUPS = {
    dai_su: [GROUP_BAN_HANG, GROUP_THONG_BAO],
    xu_ly: [GROUP_BAN_HANG, GROUP_CSKH, GROUP_THONG_BAO],
    cht: [GROUP_BAN_HANG, GROUP_QUAN_LY, GROUP_THONG_BAO],
    ke_toan: [GROUP_BAN_HANG, GROUP_KE_TOAN, GROUP_THONG_BAO],
    admin: [GROUP_BAN_HANG, GROUP_CSKH, GROUP_QUAN_LY, GROUP_KE_TOAN, GROUP_THONG_BAO],
  };
  const navGroups = NAV_GROUPS[currentUser.role];

  return (
    <div className="min-h-[600px] bg-slate-50">
      <style>{`@keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}`}</style>
      {/* top bar */}
      <div className="bg-white/90 backdrop-blur border-b border-slate-200 sticky top-0 z-30" style={{ paddingTop: "env(safe-area-inset-top)" }}>
        <div className="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between gap-3">
          <div className="flex items-center gap-2.5 min-w-0">
            <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-teal-700 to-teal-900 text-white flex items-center justify-center shrink-0 shadow-sm shadow-teal-900/25">
              <TrendingUp size={18} />
            </div>
            <div className="min-w-0">
              <p className="text-sm font-semibold text-slate-800 leading-tight tracking-tight">GUNGHO PTD</p>
              <p className="text-[11px] text-slate-400 leading-tight truncate">Theo dõi doanh thu &amp; đơn hàng</p>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <GhostButton onClick={refreshAll} className="!px-2.5" title="Làm mới dữ liệu">
              <RefreshCw size={14} />
            </GhostButton>
            <NotifBell notifications={notifications} currentUser={currentUser} onMarkRead={markRead} />
            <div className="flex items-center gap-2 pl-2 border-l border-slate-200">
              <div className="w-8 h-8 rounded-full bg-gradient-to-br from-teal-700 to-teal-900 text-white flex items-center justify-center text-xs font-semibold shrink-0 shadow-sm">
                {currentUser.name.split(" ").slice(-1)[0][0]}
              </div>
              <div className="hidden md:block text-right leading-tight">
                <p className="text-xs font-medium text-slate-800">{currentUser.name}</p>
                <p className="text-[11px] text-slate-400">{currentUser.store || ROLE_META[currentUser.role].short}</p>
              </div>
              <button onClick={handleLogout} className="w-8 h-8 rounded-xl border border-slate-200 flex items-center justify-center hover:bg-slate-50 transition" title="Đổi tài khoản">
                <LogOut size={14} className="text-slate-500" />
              </button>
            </div>
          </div>
        </div>
        {/* tabs — nhóm theo khu vực chức năng, tự xuống dòng trên máy tính/tablet, cuộn ngang trên điện thoại */}
        <div className="max-w-6xl mx-auto px-4 pb-2 flex items-center gap-1 overflow-x-auto sm:flex-wrap sm:overflow-visible">
          {navGroups.map((group, gi) => (
            <React.Fragment key={gi}>
              {gi > 0 && <span className="w-px h-6 bg-slate-200 mx-1.5 shrink-0" />}
              <div className="flex items-center gap-1 shrink-0">
                {group.map((n) => (
                  <button
                    key={n.key}
                    onClick={() => setTab(n.key)}
                    className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-sm font-medium whitespace-nowrap transition ${
                      tab === n.key
                        ? "bg-teal-800 text-white shadow-sm shadow-teal-900/20"
                        : "text-slate-500 hover:bg-slate-100 hover:text-slate-700"
                    }`}
                  >
                    <n.icon size={15} /> {n.label}
                  </button>
                ))}
              </div>
            </React.Fragment>
          ))}
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-4 py-6">
        {tab === "khach_hang" && (
          <DaiSuKhachHang currentUser={currentUser} customers={customers} orders={orders} onAdd={addCustomer} />
        )}
        {tab === "don_hang_ds" && (
          <DaiSuDonHang currentUser={currentUser} customers={customers} orders={orders} onCreate={createOrder} />
        )}
        {tab === "bao_cao_ds" && <DaiSuBaoCao currentUser={currentUser} orders={orders} />}

        {tab === "duoc_giao" && (
          <XuLyDuocGiao currentUser={currentUser} orders={orders} onConfirm={confirmHandling} onForward={forwardToAccounting} onDecline={declineOrder} />
        )}
        {tab === "don_hang_cskh" && <XuLyDonHang currentUser={currentUser} orders={orders} />}
        {tab === "bao_cao_cskh" && <XuLyBaoCao currentUser={currentUser} orders={orders} />}

        {tab === "phan_cong" && (
          <ChtPhanCong currentUser={currentUser} orders={orders} onAssign={assignHandler} />
        )}
        {tab === "bao_cao_cht" && <ChtBaoCao currentUser={currentUser} orders={orders} />}

        {tab === "cho_xac_nhan" && (
          <KeToanChoXacNhan currentUser={currentUser} orders={orders} onConfirm={confirmPayment} onReject={rejectPayment} />
        )}
        {tab === "lich_su" && <KeToanLichSu currentUser={currentUser} orders={orders} />}

        {tab === "thong_bao" && (
          <AnnouncementsPage
            currentUser={currentUser}
            announcements={announcements}
            onAdd={addAnnouncement}
            onUpdate={updateAnnouncement}
            onDelete={deleteAnnouncement}
          />
        )}
      </div>
      <Toast toast={toast} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// ĐẠI SỨ — Khách hàng
// ---------------------------------------------------------------------------

function CustomerMiniCard({ c, orderCount }) {
  const initial = (c.name || "?").trim()[0]?.toUpperCase() || "?";
  return (
    <Card className="p-4 hover:shadow-md hover:-translate-y-0.5 transition-all duration-150">
      <div className="flex items-start gap-3">
        <div className="w-10 h-10 rounded-full bg-gradient-to-br from-teal-700 to-teal-900 text-white flex items-center justify-center text-sm font-semibold shrink-0">
          {initial}
        </div>
        <div className="min-w-0 flex-1">
          <div className="flex items-center justify-between gap-2">
            <p className="font-medium text-slate-800 truncate">{c.name}</p>
            {c.customerCode && <Badge className="bg-teal-50 text-teal-700 border-teal-200 shrink-0">{c.customerCode}</Badge>}
          </div>
          <p className="text-sm text-slate-500 flex items-center gap-1.5 mt-1"><Phone size={13} className="shrink-0" /> {c.phone}</p>
          {c.address && <p className="text-sm text-slate-500 flex items-center gap-1.5 mt-1"><MapPin size={13} className="shrink-0" /> <span className="truncate">{c.address}</span></p>}
        </div>
      </div>
      <div className="flex items-center justify-between mt-3 pt-3 border-t border-slate-100">
        <span className="text-[11px] text-slate-400">Thêm lúc {fmtDate(c.createdAt)}</span>
        <Badge className={orderCount > 0 ? "bg-amber-50 text-amber-700 border-amber-200" : "bg-slate-50 text-slate-400 border-slate-200"}>
          <ShoppingBag size={11} /> {orderCount} đơn
        </Badge>
      </div>
    </Card>
  );
}

function DaiSuKhachHang({ currentUser, customers, orders, onAdd }) {
  const [form, setForm] = useState({ name: "", phone: "", address: "" });
  const [showForm, setShowForm] = useState(false);
  const [query, setQuery] = useState("");
  const isAdmin = currentUser.role === "admin";
  const mine = isAdmin ? customers : customers.filter((c) => c.createdBy === currentUser.id);

  const orderCountByCustomer = useMemo(() => {
    const map = new Map();
    (orders || []).forEach((o) => map.set(o.customerId, (map.get(o.customerId) || 0) + 1));
    return map;
  }, [orders]);

  const q = query.trim().toLowerCase();
  const filtered = q
    ? mine.filter((c) => c.name?.toLowerCase().includes(q) || c.phone?.includes(q) || c.customerCode?.toLowerCase().includes(q))
    : mine;

  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  const submit = async () => {
    if (!form.name.trim() || !form.phone.trim()) {
      setError("Vui lòng nhập tên và số điện thoại khách hàng.");
      return;
    }
    setError("");
    setSaving(true);
    try {
      await onAdd(form);
      setForm({ name: "", phone: "", address: "" });
      setShowForm(false);
    } catch (e) {
      setError("Không lưu được khách hàng, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
        <SectionTitle icon={Users} title="Danh sách khách hàng" subtitle={isAdmin ? `Toàn bộ ${mine.length} khách hàng, nhóm theo Đại sứ Gungho` : `Bạn đang theo dõi ${mine.length} khách hàng`} />
        <PrimaryButton onClick={() => setShowForm((s) => !s)}>
          <UserPlus size={15} /> Thêm khách hàng
        </PrimaryButton>
      </div>

      {showForm && (
        <Card className="p-4 mb-5">
          <div className="grid sm:grid-cols-3 gap-3 items-end">
            <TextField label="Tên khách hàng" required value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="Nguyễn Thị Hoa" />
            <TextField label="Số điện thoại" required value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} placeholder="09xxxxxxxx" />
            <TextField label="Địa chỉ" value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} placeholder="Số nhà, đường, phường/xã" />
            {error && (
              <p className="sm:col-span-3 text-sm text-rose-600 flex items-center gap-1.5"><AlertCircle size={14} /> {error}</p>
            )}
            <div className="sm:col-span-3 flex gap-2">
              <PrimaryButton type="button" onClick={submit} disabled={saving}>
                {saving ? "Đang lưu..." : "Lưu khách hàng"}
              </PrimaryButton>
              <GhostButton type="button" onClick={() => { setShowForm(false); setError(""); }}>Hủy</GhostButton>
            </div>
          </div>
        </Card>
      )}

      {mine.length > 0 && (
        <Card className="p-3 mb-4 flex items-center gap-2">
          <Search size={15} className="text-slate-400 shrink-0" />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Tìm theo tên, số điện thoại hoặc mã khách hàng..."
            className="flex-1 text-sm outline-none placeholder:text-slate-400"
          />
          {query && <button onClick={() => setQuery("")} className="text-slate-400 hover:text-slate-600"><X size={15} /></button>}
        </Card>
      )}

      {mine.length === 0 ? (
        <EmptyState icon={Users} text="Chưa có khách hàng nào — bấm “Thêm khách hàng” để bắt đầu." />
      ) : filtered.length === 0 ? (
        <EmptyState icon={Search} text="Không tìm thấy khách hàng phù hợp." />
      ) : isAdmin ? (
        <AdminCustomerGroups customers={filtered} orderCountByCustomer={orderCountByCustomer} />
      ) : (
        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">
          {filtered.map((c) => (
            <CustomerMiniCard key={c.id} c={c} orderCount={orderCountByCustomer.get(c.id) || 0} />
          ))}
        </div>
      )}
    </div>
  );
}

function AdminCustomerGroups({ customers, orderCountByCustomer }) {
  const groups = new Map();
  customers.forEach((c) => {
    const key = c.createdByName || "Không rõ người tạo";
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(c);
  });
  const sortedGroups = [...groups.entries()].sort((a, b) => b[1].length - a[1].length);

  return (
    <div className="space-y-6">
      {sortedGroups.map(([name, list]) => (
        <div key={name}>
          <div className="flex items-center gap-2 mb-3">
            <div className="w-7 h-7 rounded-full bg-teal-800 text-white flex items-center justify-center text-xs font-semibold shrink-0">
              {name.split(" ").slice(-1)[0][0]}
            </div>
            <p className="text-sm font-semibold text-slate-700">{name}</p>
            <Badge className="bg-slate-100 text-slate-600 border-slate-200">{list.length} khách hàng</Badge>
          </div>
          <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">
            {list.map((c) => (
              <CustomerMiniCard key={c.id} c={c} orderCount={orderCountByCustomer?.get(c.id) || 0} />
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

// ---------------------------------------------------------------------------
// THÔNG BÁO / HƯỚNG DẪN SỬ DỤNG — chung cho mọi vai trò
// ---------------------------------------------------------------------------

const ANNOUNCEMENT_ROLE_OPTIONS = [
  { key: "dai_su", label: "Đại sứ Gungho" },
  { key: "xu_ly", label: "Xử lý - CSKH" },
  { key: "cht", label: "Trưởng đơn vị" },
  { key: "ke_toan", label: "Kế toán" },
];

function AnnouncementForm({ initial, onSubmit, onCancel, saving }) {
  const [title, setTitle] = useState(initial?.title || "");
  const [content, setContent] = useState(initial?.content || "");
  const [isPinned, setIsPinned] = useState(initial?.isPinned || false);
  const [roles, setRoles] = useState(initial?.targetRoles?.includes("all") ? [] : (initial?.targetRoles || []));
  const [allRoles, setAllRoles] = useState(!initial || initial.targetRoles?.includes("all"));
  const [error, setError] = useState("");

  const toggleRole = (key) => {
    setRoles((prev) => (prev.includes(key) ? prev.filter((r) => r !== key) : [...prev, key]));
  };

  const submit = () => {
    if (!title.trim() || !content.trim()) {
      setError("Vui lòng nhập đủ tiêu đề và nội dung thông báo.");
      return;
    }
    if (!allRoles && roles.length === 0) {
      setError("Vui lòng chọn ít nhất 1 vai trò áp dụng, hoặc chọn “Tất cả vai trò”.");
      return;
    }
    setError("");
    onSubmit({ title: title.trim(), content: content.trim(), targetRoles: allRoles ? ["all"] : roles, isPinned });
  };

  return (
    <Card className="p-4 mb-5">
      <div className="space-y-3">
        <TextField label="Tiêu đề" required value={title} onChange={(e) => setTitle(e.target.value)} placeholder="VD: Cập nhật quy định hoa hồng khối HTC" />
        <label className="block">
          <span className="block text-xs font-medium text-slate-600 mb-1">Nội dung</span>
          <textarea
            value={content}
            onChange={(e) => setContent(e.target.value)}
            rows={5}
            placeholder="Nội dung hướng dẫn / thông báo chi tiết..."
            className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm transition focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-600"
          />
        </label>
        <div>
          <span className="block text-xs font-medium text-slate-600 mb-2">Áp dụng cho</span>
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              onClick={() => setAllRoles(true)}
              className={`px-3 py-1.5 rounded-full text-xs font-medium border transition ${allRoles ? "bg-teal-800 text-white border-teal-800" : "bg-white text-slate-600 border-slate-300 hover:bg-slate-50"}`}
            >
              Tất cả vai trò
            </button>
            {ANNOUNCEMENT_ROLE_OPTIONS.map((r) => (
              <button
                key={r.key}
                type="button"
                onClick={() => { setAllRoles(false); toggleRole(r.key); }}
                className={`px-3 py-1.5 rounded-full text-xs font-medium border transition ${!allRoles && roles.includes(r.key) ? "bg-teal-800 text-white border-teal-800" : "bg-white text-slate-600 border-slate-300 hover:bg-slate-50"}`}
              >
                {r.label}
              </button>
            ))}
          </div>
        </div>
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input type="checkbox" checked={isPinned} onChange={(e) => setIsPinned(e.target.checked)} className="rounded border-slate-300" />
          Ghim lên đầu danh sách
        </label>
        {error && <p className="text-sm text-rose-600 flex items-center gap-1.5"><AlertCircle size={14} /> {error}</p>}
        <div className="flex gap-2">
          <PrimaryButton type="button" onClick={submit} disabled={saving}>
            {saving ? "Đang lưu..." : initial ? "Lưu thay đổi" : "Đăng thông báo"}
          </PrimaryButton>
          <GhostButton type="button" onClick={onCancel}>Hủy</GhostButton>
        </div>
      </div>
    </Card>
  );
}

function AnnouncementCard({ a, isAdmin, onEdit, onDelete }) {
  const [confirmingDelete, setConfirmingDelete] = useState(false);
  const roleLabels = a.targetRoles.includes("all")
    ? "Tất cả vai trò"
    : a.targetRoles.map((r) => ANNOUNCEMENT_ROLE_OPTIONS.find((o) => o.key === r)?.label || r).join(", ");

  return (
    <Card className={`p-4 ${a.isPinned ? "border-amber-300 bg-amber-50/30" : ""}`}>
      <div className="flex items-start justify-between gap-3">
        <div className="flex items-center gap-2 flex-wrap">
          {a.isPinned && <Pin size={14} className="text-amber-600 shrink-0" />}
          <p className="font-semibold text-slate-800">{a.title}</p>
        </div>
        {isAdmin && (
          <div className="flex items-center gap-1 shrink-0">
            <button onClick={() => onEdit(a)} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-400 hover:text-teal-700 transition" title="Sửa">
              <Pencil size={14} />
            </button>
            {confirmingDelete ? (
              <div className="flex items-center gap-1">
                <button onClick={() => onDelete(a.id)} className="text-xs text-rose-600 font-medium px-2 py-1 rounded-lg bg-rose-50 hover:bg-rose-100">Xoá?</button>
                <button onClick={() => setConfirmingDelete(false)} className="text-xs text-slate-500 px-2 py-1">Hủy</button>
              </div>
            ) : (
              <button onClick={() => setConfirmingDelete(true)} className="p-1.5 rounded-lg hover:bg-rose-50 text-slate-400 hover:text-rose-600 transition" title="Xoá">
                <Trash2 size={14} />
              </button>
            )}
          </div>
        )}
      </div>
      <p className="text-sm text-slate-600 mt-2 whitespace-pre-wrap">{a.content}</p>
      <div className="flex flex-wrap items-center gap-x-3 gap-y-1 mt-3 pt-3 border-t border-slate-100 text-[11px] text-slate-400">
        <span>{a.createdByName || "Ban quản lý Gungho"}</span>
        <span>·</span>
        <span>{fmtDate(a.createdAt)}</span>
        <span>·</span>
        <Badge className="bg-slate-100 text-slate-500 border-slate-200 !text-[10px] !py-0.5">{roleLabels}</Badge>
      </div>
    </Card>
  );
}

function AnnouncementsPage({ currentUser, announcements, onAdd, onUpdate, onDelete }) {
  const isAdmin = currentUser.role === "admin";
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [saving, setSaving] = useState(false);

  const visible = announcements
    .filter((a) => a.targetRoles.includes("all") || a.targetRoles.includes(currentUser.role))
    .sort((a, b) => (b.isPinned - a.isPinned) || (new Date(b.createdAt) - new Date(a.createdAt)));

  const handleAdd = async (payload) => {
    setSaving(true);
    try {
      await onAdd(payload);
      setShowForm(false);
    } finally {
      setSaving(false);
    }
  };

  const handleUpdate = async (payload) => {
    setSaving(true);
    try {
      await onUpdate(editing.id, payload);
      setEditing(null);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
        <SectionTitle icon={Megaphone} title="Thông báo & Hướng dẫn sử dụng" subtitle={isAdmin ? "Đăng thông báo, hướng dẫn cho từng vai trò trong hệ thống" : "Các thông báo và hướng dẫn từ Ban quản lý Gungho"} />
        {isAdmin && !showForm && !editing && (
          <PrimaryButton onClick={() => setShowForm(true)}><Plus size={15} /> Đăng thông báo mới</PrimaryButton>
        )}
      </div>

      {isAdmin && showForm && (
        <AnnouncementForm onSubmit={handleAdd} onCancel={() => setShowForm(false)} saving={saving} />
      )}
      {isAdmin && editing && (
        <AnnouncementForm initial={editing} onSubmit={handleUpdate} onCancel={() => setEditing(null)} saving={saving} />
      )}

      {visible.length === 0 ? (
        <EmptyState icon={Megaphone} text="Chưa có thông báo nào." />
      ) : (
        <div className="space-y-3">
          {visible.map((a) => (
            <AnnouncementCard key={a.id} a={a} isAdmin={isAdmin} onEdit={(item) => { setEditing(item); setShowForm(false); }} onDelete={onDelete} />
          ))}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// ĐẠI SỨ — Đơn hàng
// ---------------------------------------------------------------------------

function DaiSuDonHang({ currentUser, customers, orders, onCreate }) {
  const [showForm, setShowForm] = useState(false);
  const isAdmin = currentUser.role === "admin";
  const mineCustomers = isAdmin ? customers : customers.filter((c) => c.createdBy === currentUser.id);
  const mineOrders = (isAdmin ? orders : orders.filter((o) => o.createdBy === currentUser.id)).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  const defaultBranch = branchInfo(currentUser.store) || ALL_BRANCHES[0];
  const defaultProducts = COMPANIES.find((c) => c.name === defaultBranch.company)?.products || [];

  const [form, setForm] = useState({
    customerId: "",
    company: defaultBranch.company,
    store: defaultBranch.name,
    product: defaultProducts[0] || "",
    handlerId: "",
  });
  const [customProduct, setCustomProduct] = useState(false);

  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const branchesForCompany = COMPANIES.find((c) => c.name === form.company)?.branches || [];
  const productsForCompany = COMPANIES.find((c) => c.name === form.company)?.products || [];
  const selectedBranch = branchInfo(form.store);
  const storeHandlers = USERS.filter((u) => u.role === "xu_ly" && u.store === form.store);
  const handlers = storeHandlers.length > 0 ? storeHandlers : USERS.filter((u) => u.role === "xu_ly");

  const submit = async () => {
    if (!form.customerId || !form.product.trim()) {
      setError("Vui lòng chọn khách hàng và sản phẩm.");
      return;
    }
    setError("");
    setSaving(true);
    try {
      await onCreate({ ...form, storeAddress: selectedBranch?.address || "" });
      setForm({ customerId: "", company: defaultBranch.company, store: defaultBranch.name, product: defaultProducts[0] || "", handlerId: "" });
      setCustomProduct(false);
      setShowForm(false);
    } catch (e) {
      setError("Không tạo được đơn hàng, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-4">
        <SectionTitle icon={ClipboardList} title="Đơn hàng của tôi" subtitle={`${mineOrders.length} đơn hàng đã tạo`} />
        <PrimaryButton onClick={() => setShowForm((s) => !s)} disabled={mineCustomers.length === 0}>
          <Plus size={15} /> Tạo đơn hàng
        </PrimaryButton>
      </div>

      {mineCustomers.length === 0 && (
        <div className="mb-4 flex items-center gap-2 text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-xl px-3 py-2">
          <AlertCircle size={15} /> Bạn cần thêm khách hàng trước khi tạo đơn hàng.
        </div>
      )}

      {showForm && (
        <Card className="p-4 mb-5">
          <div className="grid sm:grid-cols-2 gap-3">
            <SelectField label="Khách hàng" required value={form.customerId} onChange={(e) => setForm({ ...form, customerId: e.target.value })}>
              <option value="">— Chọn khách hàng —</option>
              {mineCustomers.map((c) => (
                <option key={c.id} value={c.id}>{c.customerCode ? `[${c.customerCode}] ` : ""}{c.name} — {c.phone}</option>
              ))}
            </SelectField>
            <SelectField label="Khối công ty / đối tác" required value={form.company} onChange={(e) => {
              const newCompany = e.target.value;
              const co = COMPANIES.find((c) => c.name === newCompany);
              const firstBranch = co?.branches[0];
              setCustomProduct(false);
              setForm({ ...form, company: newCompany, store: firstBranch?.name || "", product: co?.products[0] || "", handlerId: "" });
            }}>
              {COMPANIES.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
            </SelectField>
            <SelectField label="Cửa hàng / chi nhánh" required value={form.store} onChange={(e) => setForm({ ...form, store: e.target.value, handlerId: "" })}>
              {branchesForCompany.map((b) => <option key={b.name} value={b.name}>{b.name}</option>)}
            </SelectField>
            {selectedBranch?.address && (
              <p className="sm:col-span-2 -mt-1 text-xs text-slate-500 flex items-center gap-1.5"><MapPin size={13} /> {selectedBranch.address}</p>
            )}
            {customProduct ? (
              <div>
                <TextField
                  label="Sản phẩm (nhập tay)"
                  required
                  value={form.product}
                  onChange={(e) => setForm({ ...form, product: e.target.value })}
                  placeholder="Nhập tên sản phẩm/dịch vụ"
                />
                <button type="button" className="text-xs text-teal-700 hover:underline mt-1" onClick={() => { setCustomProduct(false); setForm({ ...form, product: productsForCompany[0] || "" }); }}>
                  ← Chọn từ danh mục có sẵn
                </button>
              </div>
            ) : (
              <SelectField
                label="Sản phẩm"
                required
                value={form.product}
                onChange={(e) => {
                  if (e.target.value === "__custom__") {
                    setCustomProduct(true);
                    setForm({ ...form, product: "" });
                  } else {
                    setForm({ ...form, product: e.target.value });
                  }
                }}
              >
                {productsForCompany.map((p) => <option key={p} value={p}>{p}</option>)}
                <option value="__custom__">Khác (nhập tay)...</option>
              </SelectField>
            )}
            <SelectField label="Người chăm sóc (tùy chọn)" value={form.handlerId} onChange={(e) => setForm({ ...form, handlerId: e.target.value })}>
              <option value="">— Không chọn, gửi về trưởng đơn vị —</option>
              {handlers.map((h) => <option key={h.id} value={h.id}>{h.name}</option>)}
            </SelectField>
            {error && (
              <p className="sm:col-span-2 text-sm text-rose-600 flex items-center gap-1.5"><AlertCircle size={14} /> {error}</p>
            )}
            <div className="sm:col-span-2 flex gap-2">
              <PrimaryButton type="button" onClick={submit} disabled={saving}>
                {saving ? "Đang tạo..." : "Tạo đơn hàng"}
              </PrimaryButton>
              <GhostButton type="button" onClick={() => { setShowForm(false); setError(""); }}>Hủy</GhostButton>
            </div>
          </div>
        </Card>
      )}

      {mineOrders.length === 0 ? (
        <EmptyState icon={ClipboardList} text="Chưa có đơn hàng nào." />
      ) : (
        <div className="space-y-2">
          {mineOrders.map((o) => <OrderRow key={o.id} order={o} showCommission />)}
        </div>
      )}
    </div>
  );
}

function OrderRow({ order, showCommission, right }) {
  const isPaid = order.status === "da_thanh_toan";
  return (
    <Card className="p-4">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="flex items-center gap-2 flex-wrap">
            <p className="font-medium text-slate-800">{order.customerName}</p>
            <StatusBadge status={order.status} />
          </div>
          <p className="text-sm text-slate-500 mt-1">
            {order.product} · <Store size={12} className="inline -mt-0.5" /> {order.store}
          </p>
          {order.assignedHandlerName && (
            <p className="text-xs text-slate-400 mt-1">Người chăm sóc: {order.assignedHandlerName}</p>
          )}
          {order.handlerNote && <p className="text-xs text-slate-500 mt-1 italic">Ghi chú: {order.handlerNote}</p>}
        </div>
        <div className="text-right shrink-0">
          {isPaid ? (
            <p className="font-semibold text-slate-800">{fmtMoney(order.finalAmount ?? order.totalAmount)}</p>
          ) : (
            <p className="text-xs text-slate-400">Chưa xác định số tiền</p>
          )}
          {showCommission && isPaid && (
            <p className="text-xs text-amber-700 font-medium mt-1">+ {fmtMoney(order.commissionAmount)} hoa hồng</p>
          )}
          {right}
        </div>
      </div>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// ĐẠI SỨ — Báo cáo
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Bảng xếp hạng dạng "leaderboard": huy chương top 3, tên đầy đủ không bị cắt,
// thanh màu gradient theo hạng — dùng chung cho doanh thu & điểm thi đua Gungho
// ---------------------------------------------------------------------------
const RANK_MEDAL = ["🥇", "🥈", "🥉"];
const RANK_BAR_STYLE = [
  "from-amber-400 to-amber-600",
  "from-slate-300 to-slate-500",
  "from-orange-400 to-orange-600",
];
const RANK_ROW_BG = ["bg-amber-50/60", "bg-slate-50", "bg-orange-50/60"];

function RankedLeaderRow({ rank, name, valueLabel, percent }) {
  const isTop3 = rank < 3;
  return (
    <div className={`flex items-center gap-3 rounded-xl px-2.5 py-2 transition hover:bg-slate-50 ${isTop3 ? RANK_ROW_BG[rank] : ""}`}>
      <div className="w-8 h-8 rounded-full flex items-center justify-center shrink-0 text-base font-bold">
        {isTop3 ? (
          <span className="text-xl leading-none">{RANK_MEDAL[rank]}</span>
        ) : (
          <span className="w-7 h-7 rounded-full bg-slate-100 text-slate-500 flex items-center justify-center text-xs">{rank + 1}</span>
        )}
      </div>
      <div className="flex-1 min-w-0">
        <div className="flex items-center justify-between gap-2 mb-1">
          <p className={`text-sm truncate ${isTop3 ? "font-semibold text-slate-800" : "font-medium text-slate-700"}`} title={name}>{name}</p>
          <p className={`text-sm shrink-0 font-semibold ${isTop3 ? "text-slate-800" : "text-slate-600"}`}>{valueLabel}</p>
        </div>
        <div className="h-2 bg-slate-100 rounded-full overflow-hidden">
          <div
            className={`h-full rounded-full bg-gradient-to-r transition-all duration-500 ${isTop3 ? RANK_BAR_STYLE[rank] : "from-teal-500 to-teal-700"}`}
            style={{ width: `${Math.max(percent, 3)}%` }}
          />
        </div>
      </div>
    </div>
  );
}

function LeaderBoard({ orders, groupKeyFn, title, icon }) {
  const rows = buildRevenueLeaderboard(orders, groupKeyFn).slice(0, 8);
  const max = rows[0]?.revenue || 0;
  return (
    <Card className="p-4 sm:p-5">
      <SectionTitle icon={icon} title={title} />
      {rows.length === 0 ? (
        <EmptyState icon={Award} text="Chưa có dữ liệu xếp hạng." />
      ) : (
        <div className="space-y-1">
          {rows.map((r, i) => (
            <RankedLeaderRow
              key={r.name}
              rank={i}
              name={r.name}
              valueLabel={fmtMoney(r.revenue)}
              percent={max > 0 ? (r.revenue / max) * 100 : 0}
            />
          ))}
        </div>
      )}
    </Card>
  );
}

function GunghoLeaderBoard({ orders, groupKeyFn, title, icon = Award }) {
  const rows = buildTDLeaderboard(orders, groupKeyFn).slice(0, 8);
  const max = rows[0]?.td || 0;
  return (
    <Card className="p-4 sm:p-5">
      <SectionTitle icon={icon} title={title} subtitle="TD = BX×0.5 + DVX/500.000 + BO×2 + DVO/1.000.000 + NH/1.000.000 + KS/500.000 + TO×20 + V×1 + VYC/300.000 + VT/500.000" />
      {rows.length === 0 ? (
        <EmptyState icon={Award} text="Chưa có dữ liệu xếp hạng." />
      ) : (
        <div className="space-y-1">
          {rows.map((r, i) => (
            <RankedLeaderRow
              key={r.name}
              rank={i}
              name={r.name}
              valueLabel={`${r.td.toFixed(1)} điểm`}
              percent={max > 0 ? (r.td / max) * 100 : 0}
            />
          ))}
        </div>
      )}
    </Card>
  );
}

// Biểu đồ vùng thể hiện xu hướng doanh thu 6 tháng gần nhất
function RevenueTrendChart({ orders, title = "Xu hướng doanh thu 6 tháng gần đây" }) {
  const data = buildMonthlyTrend(orders, 6);
  const hasData = data.some((d) => d.revenue > 0);
  return (
    <Card className="p-4 sm:p-5">
      <SectionTitle icon={TrendingUp} title={title} />
      {!hasData ? (
        <EmptyState icon={TrendingUp} text="Chưa có đủ dữ liệu để hiển thị xu hướng." />
      ) : (
        <div style={{ width: "100%", height: 220 }}>
          <ResponsiveContainer>
            <AreaChart data={data} margin={{ top: 8, right: 12, left: -12, bottom: 0 }}>
              <defs>
                <linearGradient id="revenueFill" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor="#0f766e" stopOpacity={0.35} />
                  <stop offset="100%" stopColor="#0f766e" stopOpacity={0.02} />
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="label" tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} />
              <YAxis tickFormatter={shortMoney} tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} width={48} />
              <Tooltip formatter={(v) => fmtMoney(v)} {...CHART_TOOLTIP_STYLE} />
              <Area type="monotone" dataKey="revenue" stroke="#0f766e" strokeWidth={2.5} fill="url(#revenueFill)" name="Doanh thu" />
            </AreaChart>
          </ResponsiveContainer>
        </div>
      )}
    </Card>
  );
}

// Biểu đồ tròn tỉ trọng doanh thu theo sản phẩm
function ProductMixPieChart({ orders, title = "Tỉ trọng doanh thu theo sản phẩm" }) {
  const data = buildProductMix(orders);
  return (
    <Card className="p-4 sm:p-5">
      <SectionTitle icon={ShoppingBag} title={title} />
      {data.length === 0 ? (
        <EmptyState icon={ShoppingBag} text="Chưa có dữ liệu doanh thu." />
      ) : (
        <div style={{ width: "100%", height: 260 }}>
          <ResponsiveContainer>
            <PieChart>
              <Pie data={data} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={2}>
                {data.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
              </Pie>
              <Tooltip formatter={(v) => fmtMoney(v)} {...CHART_TOOLTIP_STYLE} />
              <Legend
                layout="vertical"
                verticalAlign="middle"
                align="right"
                wrapperStyle={{ fontSize: 11.5, lineHeight: "20px", maxWidth: "45%" }}
                formatter={(value) => <span className="text-slate-600">{truncateLabel(value, 22)}</span>}
              />
            </PieChart>
          </ResponsiveContainer>
        </div>
      )}
    </Card>
  );
}

function DaiSuBaoCao({ currentUser, orders }) {
  const isAdmin = currentUser.role === "admin";
  // Thanh lọc theo công ty — chỉ hiển thị và có tác dụng khi xem bằng tài khoản Admin
  const [viewCompany, setViewCompany] = useState("");
  const viewOrders = isAdmin && viewCompany ? orders.filter((o) => o.company === viewCompany) : orders;
  const mine = isAdmin ? viewOrders : orders.filter((o) => o.createdBy === currentUser.id);
  const paid = mine.filter((o) => o.status === "da_thanh_toan");
  const revenue = paid.reduce((s, o) => s + (o.finalAmount ?? o.totalAmount), 0);
  const commission = paid.reduce((s, o) => s + (o.commissionAmount || 0), 0);
  const caring = mine.filter((o) => o.status === "dang_cham_soc").length;
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");

  const handleExport = () => {
    const sheets = [
      { name: "Đơn hàng của tôi", rows: mine.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt)).map(orderExportRow) },
      { name: "Xếp hạng Gungho (TD)", rows: buildTDLeaderboard(orders, (o) => o.createdByName).map((r) => ({ "Đại sứ": r.name, "Điểm TD": r.td })) },
      { name: "Doanh thu theo công ty", rows: buildRevenueLeaderboard(orders, (o) => o.company).map((r) => ({ "Khối công ty": r.name, "Doanh thu": r.revenue, "Số đơn": r.count })) },
      { name: "Xếp hạng sản phẩm", rows: buildRevenueLeaderboard(orders, (o) => o.product).map((r) => ({ "Sản phẩm": r.name, "Doanh thu": r.revenue, "Số đơn": r.count })) },
    ];
    exportToExcel(sheets, `BaoCao_${currentUser.name.replace(/\s+/g, "")}_${Date.now()}.xlsx`);
  };

  const handleTemplateExport = () => {
    exportGungHoNhanVienTemplate({ ambassador: currentUser, orders, fromDate, toDate });
  };

  return (
    <div className="space-y-5">
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <SectionTitle
          icon={BarChart3}
          title="Báo cáo tổng hợp"
          subtitle={isAdmin && viewCompany ? `Đang xem: ${viewCompany}` : "Toàn bộ đơn hàng, hoa hồng và xếp hạng"}
        />
        <GhostButton onClick={handleExport}><Download size={15} /> Xuất Excel</GhostButton>
      </div>

      {isAdmin && (
        <Card className="p-3 sm:p-4">
          <div className="flex items-center gap-3 flex-wrap">
            <Building2 size={16} className="text-teal-700 shrink-0" />
            <p className="text-sm font-medium text-slate-700 shrink-0">Xem báo cáo &amp; xếp hạng theo:</p>
            <div className="w-full sm:w-72">
              <select
                value={viewCompany}
                onChange={(e) => setViewCompany(e.target.value)}
                className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm bg-white font-medium text-slate-700 transition focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-600"
              >
                <option value="">— Tất cả (toàn tập đoàn) —</option>
                {COMPANIES.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
              </select>
            </div>
            {viewCompany && (
              <button onClick={() => setViewCompany("")} className="text-xs text-teal-700 hover:underline shrink-0">
                ← Xem lại toàn tập đoàn
              </button>
            )}
          </div>
        </Card>
      )}

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <MetricCard label="Tổng doanh thu" value={fmtMoney(revenue)} icon={TrendingUp} accent="teal" />
        <MetricCard label="Hoa hồng nhận được" value={fmtMoney(commission)} icon={Wallet} accent="amber" />
        <MetricCard label="Đơn đang chăm sóc" value={caring} icon={ClipboardList} accent="indigo" />
        <MetricCard label="Tổng số đơn hàng" value={mine.length} icon={ShoppingBag} accent="rose" />
      </div>

      <Card className="p-4">
        <p className="font-semibold text-slate-800 text-sm mb-1">Xuất mẫu "Kết quả Gung Ho chi tiết nhân viên"</p>
        <p className="text-xs text-slate-500 mb-3">Chọn khoảng thời gian cần tra cứu (bỏ trống nếu muốn lấy toàn bộ dữ liệu).</p>
        <div className="flex flex-wrap items-end gap-3">
          <TextField label="Từ ngày" type="date" value={fromDate} onChange={(e) => setFromDate(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={toDate} onChange={(e) => setToDate(e.target.value)} />
          <PrimaryButton onClick={handleTemplateExport}><Download size={15} /> Xuất theo mẫu</PrimaryButton>
        </div>
      </Card>

      <Card className="p-0 overflow-hidden">
        <div className="p-4 border-b border-slate-100">
          <p className="font-semibold text-slate-800 text-sm">Danh sách đơn hàng &amp; tình trạng</p>
        </div>
        {mine.length === 0 ? (
          <EmptyState icon={Inbox} text="Chưa có đơn hàng." />
        ) : (
          <div className="divide-y divide-slate-100">
            {mine.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt)).map((o) => (
              <div key={o.id} className="p-4 flex flex-wrap items-center justify-between gap-2">
                <div className="min-w-0">
                  <p className="text-sm font-medium text-slate-800">{o.customerName} — {o.product}</p>
                  <p className="text-xs text-slate-400">{o.store} · {fmtDate(o.createdAt)}</p>
                </div>
                <div className="flex items-center gap-3">
                  <StatusBadge status={o.status} />
                  <span className="text-sm font-semibold text-slate-800 w-28 text-right">
                    {o.status === "da_thanh_toan" ? fmtMoney(o.finalAmount ?? o.totalAmount) : "—"}
                  </span>
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>

      <div className="grid lg:grid-cols-2 gap-5">
        <RevenueTrendChart orders={mine} title={isAdmin && viewCompany ? `Xu hướng doanh thu — ${viewCompany}` : "Xu hướng doanh thu 6 tháng gần đây"} />
        <ProductMixPieChart orders={mine} title={isAdmin && viewCompany ? `Tỉ trọng sản phẩm — ${viewCompany}` : "Tỉ trọng doanh thu theo sản phẩm"} />
      </div>

      <GunghoLeaderBoard
        orders={viewOrders}
        groupKeyFn={(o) => o.createdByName}
        title={isAdmin && viewCompany ? `Bảng xếp hạng Gungho — nhân viên ${viewCompany}` : "Bảng xếp hạng Gungho (theo điểm thi đua)"}
        icon={Award}
      />
      <div className="grid lg:grid-cols-2 gap-5">
        {isAdmin && viewCompany ? (
          <LeaderBoard orders={viewOrders} groupKeyFn={(o) => o.store} title={`Doanh thu theo cửa hàng / chi nhánh — ${viewCompany}`} icon={Store} />
        ) : (
          <LeaderBoard orders={viewOrders} groupKeyFn={(o) => o.company} title="Doanh thu theo khối công ty" icon={Building2} />
        )}
        <LeaderBoard orders={viewOrders} groupKeyFn={(o) => o.product} title="Xếp hạng theo sản phẩm" icon={ShoppingBag} />
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// XỬ LÝ - CHĂM SÓC
// ---------------------------------------------------------------------------

function HandlerActionCard({ order, onConfirm, onForward, onDecline }) {
  const [note, setNote] = useState(order.handlerNote || "");
  const isHTCOrder = order.company === HTC_COMPANY_NAME;
  const [amount, setAmount] = useState(order.totalAmount || "");
  const [discount, setDiscount] = useState(order.discountAmount || 0);
  const [invoiceName, setInvoiceName] = useState(order.invoiceName || "");
  const [invoiceNumber, setInvoiceNumber] = useState(order.invoiceNumber || "");
  const [invoiceError, setInvoiceError] = useState("");

  const handleForward = () => {
    if (isHTCOrder) {
      if (!amount || Number(amount) <= 0 || !invoiceName.trim() || !invoiceNumber.trim()) {
        setInvoiceError("Vui lòng điền đủ Số tiền đơn hàng, Tên khách hàng xuất hóa đơn và Số hóa đơn trước khi chuyển kế toán.");
        return;
      }
      setInvoiceError("");
      onForward(order.id, note, {
        amount: Number(amount) || 0, discountAmount: Number(discount) || 0,
        invoiceName: invoiceName.trim(), invoiceNumber: invoiceNumber.trim(),
      });
    } else {
      onForward(order.id, note);
    }
  };

  return (
    <Card className="p-4">
      <div className="flex flex-wrap items-start justify-between gap-3 mb-3">
        <div>
          <div className="flex items-center gap-2 flex-wrap">
            <p className="font-medium text-slate-800">{order.customerName}</p>
            <StatusBadge status={order.status} />
          </div>
          <p className="text-sm text-slate-500 mt-1 flex items-center gap-1.5"><Phone size={13} /> {order.customerPhone}</p>
          <p className="text-sm text-slate-500 mt-1">{order.product} · <Store size={12} className="inline -mt-0.5" /> {order.store}</p>
          <p className="text-xs text-slate-400 mt-1">Đại sứ phụ trách: {order.createdByName}</p>
        </div>
      </div>
      <TextAreaField label="Ghi chú chăm sóc (gửi về Đại sứ Gungho)" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ví dụ: đã liên hệ, khách đang cân nhắc..." />
      {order.status === "dang_cham_soc" && isHTCOrder && (
        <div className="mt-3 pt-3 border-t border-slate-100">
          <p className="text-xs font-medium text-slate-600 mb-2">Thông tin đơn hàng & hóa đơn (riêng khối HTC — điền trước khi chuyển kế toán)</p>
          <div className="grid sm:grid-cols-2 gap-3 mb-3">
            <MoneyField label="Số tiền đơn hàng (đ)" value={amount} onChange={setAmount} />
            <MoneyField label="Chiết khấu (đ)" value={discount} onChange={setDiscount} />
            <TextField label="Tên khách hàng xuất hóa đơn" value={invoiceName} onChange={(e) => setInvoiceName(e.target.value)} placeholder="Tên trên hóa đơn" />
            <TextField label="Số hóa đơn" value={invoiceNumber} onChange={(e) => setInvoiceNumber(e.target.value)} placeholder="VD: HD-000123" />
          </div>
          {invoiceError && <p className="text-xs text-rose-600 mt-1">{invoiceError}</p>}
        </div>
      )}
      <div className="flex flex-wrap gap-2 mt-3">
        {order.status === "cho_xu_ly" && (
          <PrimaryButton onClick={() => onConfirm(order.id, note)}><CheckCircle2 size={15} /> Xác nhận chăm sóc</PrimaryButton>
        )}
        {order.status === "dang_cham_soc" && (
          <>
            <PrimaryButton onClick={handleForward}><Send size={15} /> Khách đồng ý mua — Chuyển kế toán</PrimaryButton>
            <DangerButton onClick={() => onDecline(order.id, note)}><XCircle size={15} /> Khách không mua</DangerButton>
          </>
        )}
      </div>
    </Card>
  );
}


function XuLyDuocGiao({ currentUser, orders, onConfirm, onForward, onDecline }) {
  const isAdmin = currentUser.role === "admin";
  const mine = orders
    .filter((o) => (isAdmin || o.assignedHandler === currentUser.id) && ["cho_xu_ly", "dang_cham_soc"].includes(o.status))
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  return (
    <div>
      <SectionTitle icon={Inbox} title="Đơn hàng được giao" subtitle={`${mine.length} đơn đang chờ bạn xử lý`} />
      {mine.length === 0 ? (
        <EmptyState icon={Inbox} text="Hiện chưa có đơn hàng nào được giao cho bạn." />
      ) : (
        <div className="space-y-3">
          {mine.map((o) => (
            <HandlerActionCard key={o.id} order={o} onConfirm={onConfirm} onForward={onForward} onDecline={onDecline} />
          ))}
        </div>
      )}
    </div>
  );
}

const ORDER_GROUPS = [
  { key: "cho_xac_nhan", label: "Chờ xác nhận", statuses: ["cho_xu_ly"], badge: "bg-amber-50 text-amber-700 border-amber-300" },
  { key: "da_xac_nhan", label: "Đã xác nhận", statuses: ["dang_cham_soc", "cho_ke_toan"], badge: "bg-sky-50 text-sky-700 border-sky-300" },
  { key: "hoan_thanh", label: "Hoàn thành", statuses: ["da_thanh_toan"], badge: "bg-emerald-50 text-emerald-700 border-emerald-300" },
  { key: "da_huy", label: "Đã hủy", statuses: ["khong_thanh_toan"], badge: "bg-rose-50 text-rose-700 border-rose-300" },
];

function OrderListCard({ order, group }) {
  return (
    <Card className="p-4">
      <div className="flex items-start justify-between gap-3 mb-2">
        <p className="font-medium text-slate-800">{order.orderCode || order.id}</p>
        <Badge className={group.badge}>{group.label}</Badge>
      </div>
      <div className="space-y-1.5 text-sm text-slate-600">
        <p className="flex items-center gap-2"><Users size={14} className="text-slate-400" /> {order.customerName}</p>
        <p className="flex items-center gap-2"><Store size={14} className="text-slate-400" /> {order.store}</p>
        <p className="flex items-center gap-2"><UserPlus size={14} className="text-slate-400" /> Tạo bởi: {order.createdByName}</p>
        <p className="flex items-center gap-2 text-slate-400 text-xs"><Clock size={13} /> {fmtDate(order.createdAt)}</p>
      </div>
      <p className="text-sm text-amber-700 font-medium mt-2 flex items-center gap-1.5">
        <ShoppingBag size={14} /> {order.product}
      </p>
    </Card>
  );
}

function XuLyDonHang({ currentUser, orders }) {
  const [groupKey, setGroupKey] = useState("cho_xac_nhan");
  const [phoneQuery, setPhoneQuery] = useState("");
  const [storeFilter, setStoreFilter] = useState("");
  const isAdmin = currentUser.role === "admin";

  const mine = isAdmin ? orders : orders.filter((o) => o.assignedHandler === currentUser.id);
  const activeGroup = ORDER_GROUPS.find((g) => g.key === groupKey);
  const filtered = mine
    .filter((o) => activeGroup.statuses.includes(o.status))
    .filter((o) => !phoneQuery.trim() || (o.customerPhone || "").includes(phoneQuery.trim()))
    .filter((o) => !storeFilter || o.store === storeFilter)
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  return (
    <div>
      <SectionTitle icon={ClipboardList} title="Đơn hàng" subtitle={isAdmin ? `${mine.length} đơn hàng (tất cả nhân viên)` : `${mine.length} đơn hàng được giao cho bạn`} />

      <div className="flex flex-wrap gap-3 mb-3">
        <div className="relative flex-1 min-w-[200px]">
          <Phone size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            value={phoneQuery}
            onChange={(e) => setPhoneQuery(e.target.value)}
            placeholder="Lọc theo SĐT..."
            className="w-full pl-9 pr-3 py-2.5 rounded-xl border border-slate-300 text-sm transition focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-600"
          />
        </div>
        {isAdmin && (
          <div className="w-full sm:w-64">
            <select
              value={storeFilter}
              onChange={(e) => setStoreFilter(e.target.value)}
              className="w-full px-3 py-2.5 rounded-xl border border-slate-300 text-sm bg-white transition focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-600"
            >
              <option value="">— Tất cả chi nhánh —</option>
              {ALL_BRANCHES.map((b) => <option key={b.name} value={b.name}>{b.name}</option>)}
            </select>
          </div>
        )}
      </div>

      <div className="flex gap-1 overflow-x-auto mb-4 border-b border-slate-200">
        {ORDER_GROUPS.map((g) => {
          const count = mine.filter((o) => g.statuses.includes(o.status)).length;
          return (
            <button
              key={g.key}
              onClick={() => setGroupKey(g.key)}
              className={`px-3.5 py-2.5 text-sm font-medium border-b-2 whitespace-nowrap transition ${
                groupKey === g.key ? "border-teal-800 text-teal-800" : "border-transparent text-slate-500 hover:text-slate-700"
              }`}
            >
              {g.label} {count > 0 && <span className="text-xs text-slate-400">({count})</span>}
            </button>
          );
        })}
      </div>

      {filtered.length === 0 ? (
        <EmptyState icon={ClipboardList} text="Không có đơn hàng nào trong mục này." />
      ) : (
        <div className="space-y-3">
          {filtered.map((o) => <OrderListCard key={o.id} order={o} group={activeGroup} />)}
        </div>
      )}
    </div>
  );
}

function XuLyBaoCao({ currentUser, orders }) {
  const isAdmin = currentUser.role === "admin";
  const mine = isAdmin ? orders : orders.filter((o) => o.assignedHandler === currentUser.id);
  const byProduct = new Map();
  mine.forEach((o) => {
    const cur = byProduct.get(o.product) || 0;
    byProduct.set(o.product, cur + 1);
  });
  const rows = [...byProduct.entries()].sort((a, b) => b[1] - a[1]);
  const caring = mine.filter((o) => o.status === "dang_cham_soc").length;
  const done = mine.filter((o) => o.status === "da_thanh_toan").length;

  const handleExport = () => {
    const sheets = [
      { name: "Đơn được giao", rows: mine.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt)).map(orderExportRow) },
      { name: "Theo sản phẩm", rows: rows.map(([product, count]) => ({ "Sản phẩm": product, "Số khách": count })) },
    ];
    exportToExcel(sheets, `BaoCaoCSKH_${currentUser.name.replace(/\s+/g, "")}_${Date.now()}.xlsx`);
  };

  return (
    <div className="space-y-5">
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <SectionTitle icon={BarChart3} title="Báo cáo xử lý & chăm sóc" />
        <GhostButton onClick={handleExport}><Download size={15} /> Xuất Excel</GhostButton>
      </div>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <MetricCard label="Tổng khách được giao" value={mine.length} icon={Users} accent="teal" />
        <MetricCard label="Đang chăm sóc" value={caring} icon={ClipboardList} accent="amber" />
        <MetricCard label="Đã hoàn tất" value={done} icon={CheckCircle2} accent="indigo" />
        <MetricCard label="Số sản phẩm khác nhau" value={byProduct.size} icon={ShoppingBag} accent="rose" />
      </div>
      <Card className="p-4 sm:p-5">
        <p className="font-semibold text-slate-800 text-sm mb-3">Tình trạng xử lý theo sản phẩm</p>
        {rows.length === 0 ? (
          <EmptyState icon={ShoppingBag} text="Chưa có dữ liệu." />
        ) : (
          <>
            <div style={{ width: "100%", height: Math.max(rows.length * 32, 140) }}>
              <ResponsiveContainer>
                <BarChart data={rows.map(([p, c]) => ({ name: truncateLabel(p, 18), count: c }))} layout="vertical" margin={{ top: 0, right: 24, left: 0, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e2e8f0" />
                  <XAxis type="number" allowDecimals={false} tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 12, fill: "#475569" }} axisLine={false} tickLine={false} />
                  <Tooltip formatter={(v) => `${v} khách`} {...CHART_TOOLTIP_STYLE} />
                  <Bar dataKey="count" fill="#4f46e5" radius={[0, 6, 6, 0]} maxBarSize={20} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </>
        )}
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// CỬA HÀNG TRƯỞNG
// ---------------------------------------------------------------------------

function AssignCard({ order, onAssign }) {
  const [handlerId, setHandlerId] = useState("");
  const storeHandlers = USERS.filter((u) => u.role === "xu_ly" && u.store === order.store);
  const handlers = storeHandlers.length > 0 ? storeHandlers : USERS.filter((u) => u.role === "xu_ly");
  return (
    <Card className="p-4">
      <div className="flex flex-wrap items-start justify-between gap-3 mb-3">
        <div>
          <p className="font-medium text-slate-800">{order.customerName}</p>
          <p className="text-sm text-slate-500 mt-1">{order.product} · Đại sứ: {order.createdByName}</p>
          <p className="text-xs text-slate-400 mt-1 flex items-center gap-1.5"><Store size={12} /> {order.store}{order.company ? ` — ${order.company}` : ""}</p>
        </div>
        <StatusBadge status={order.status} />
      </div>
      <div className="flex flex-wrap items-end gap-2">
        <div className="flex-1 min-w-[180px]">
          <SelectField label="Chọn nhân viên xử lý" value={handlerId} onChange={(e) => setHandlerId(e.target.value)}>
            <option value="">— Chọn —</option>
            {handlers.map((h) => <option key={h.id} value={h.id}>{h.name}</option>)}
          </SelectField>
        </div>
        <PrimaryButton disabled={!handlerId} onClick={() => onAssign(order.id, handlerId)}>
          <ArrowRightLeft size={15} /> Phân công
        </PrimaryButton>
      </div>
    </Card>
  );
}

function ChtPhanCong({ currentUser, orders, onAssign }) {
  const pending = orders
    .filter((o) => o.status === "cho_phan_cong")
    .sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
  return (
    <div>
      <SectionTitle icon={ArrowRightLeft} title="Đơn hàng chờ phân công" subtitle="Tất cả khối công ty / chi nhánh" />
      {pending.length === 0 ? (
        <EmptyState icon={ArrowRightLeft} text="Không có đơn hàng nào đang chờ phân công." />
      ) : (
        <div className="space-y-3">
          {pending.map((o) => <AssignCard key={o.id} order={o} onAssign={onAssign} />)}
        </div>
      )}
    </div>
  );
}

function ChtBaoCao({ currentUser, orders }) {
  // Thanh lọc xem báo cáo theo công ty: để trống ("") = xem toàn tập đoàn
  const [viewCompany, setViewCompany] = useState("");
  const viewOrders = viewCompany ? orders.filter((o) => o.company === viewCompany) : orders;

  const paid = viewOrders.filter((o) => o.status === "da_thanh_toan");
  const revenue = paid.reduce((s, o) => s + (o.finalAmount ?? o.totalAmount), 0);
  const defaultCompany = branchInfo(currentUser.store)?.company || COMPANIES[0].name;
  const [tplCompany, setTplCompany] = useState(defaultCompany);
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const [tplUnit, setTplUnit] = useState(currentUser.store || ALL_BRANCHES[0].name);
  const [tplYear, setTplYear] = useState(String(new Date().getFullYear()));

  const handleExport = () => {
    const sheets = [
      { name: "Tất cả đơn hàng", rows: orders.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt)).map(orderExportRow) },
      { name: "Xếp hạng khối công ty", rows: buildRevenueLeaderboard(orders, (o) => o.company).map((r) => ({ "Khối công ty": r.name, "Doanh thu": r.revenue, "Số đơn": r.count })) },
      { name: "Xếp hạng cửa hàng", rows: buildRevenueLeaderboard(orders, (o) => o.store).map((r) => ({ "Cửa hàng / chi nhánh": r.name, "Doanh thu": r.revenue, "Số đơn": r.count })) },
      { name: "Xếp hạng Gungho (TD)", rows: buildTDLeaderboard(orders, (o) => o.createdByName).map((r) => ({ "Đại sứ": r.name, "Điểm TD": r.td })) },
      { name: "Xếp hạng sản phẩm", rows: buildRevenueLeaderboard(orders, (o) => o.product).map((r) => ({ "Sản phẩm": r.name, "Doanh thu": r.revenue, "Số đơn": r.count })) },
    ];
    exportToExcel(sheets, `BaoCao_DoanhSo_${currentUser.store?.replace(/\s+/g, "") || "TatCa"}_${Date.now()}.xlsx`);
  };

  const handleTemplateExport = () => {
    exportGungHoCongTyTemplate({ companyName: tplCompany, orders, fromDate, toDate });
  };

  const handleTimeTemplateExport = () => {
    exportGungHoThoiGianTemplate({ unitName: tplUnit, orders, year: tplYear });
  };

  const handleRankingExport = () => {
    exportGungHoRankingTemplate({ companyName: tplCompany, orders, fromDate, toDate });
  };

  return (
    <div className="space-y-5">
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <SectionTitle
          icon={BarChart3}
          title="Báo cáo doanh số Gungho"
          subtitle={viewCompany ? `Đang xem: ${viewCompany}` : "Toàn bộ khối công ty & chi nhánh (Tập đoàn)"}
        />
        <GhostButton onClick={handleExport}><Download size={15} /> Xuất Excel</GhostButton>
      </div>

      <Card className="p-3 sm:p-4">
        <div className="flex items-center gap-3 flex-wrap">
          <Building2 size={16} className="text-teal-700 shrink-0" />
          <p className="text-sm font-medium text-slate-700 shrink-0">Xem báo cáo &amp; xếp hạng theo:</p>
          <div className="w-full sm:w-72">
            <select
              value={viewCompany}
              onChange={(e) => setViewCompany(e.target.value)}
              className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm bg-white font-medium text-slate-700 transition focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-600"
            >
              <option value="">— Tất cả (toàn tập đoàn) —</option>
              {COMPANIES.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
            </select>
          </div>
          {viewCompany && (
            <button onClick={() => setViewCompany("")} className="text-xs text-teal-700 hover:underline shrink-0">
              ← Xem lại toàn tập đoàn
            </button>
          )}
        </div>
      </Card>

      <div className="grid grid-cols-2 lg:grid-cols-3 gap-3">
        <MetricCard label="Tổng doanh thu" value={fmtMoney(revenue)} icon={Building2} accent="teal" />
        <MetricCard label="Đơn đã hoàn tất" value={paid.length} icon={CheckCircle2} accent="amber" />
        <MetricCard label="Tổng số đơn hàng" value={viewOrders.length} icon={ShoppingBag} accent="indigo" />
      </div>

      <Card className="p-4">
        <p className="font-semibold text-slate-800 text-sm mb-1">Xuất mẫu "Kết quả Gung Ho chi tiết công ty theo đơn vị"</p>
        <p className="text-xs text-slate-500 mb-3">Chọn khối công ty và khoảng thời gian cần tra cứu (bỏ trống ngày nếu muốn lấy toàn bộ dữ liệu).</p>
        <div className="flex flex-wrap items-end gap-3">
          <SelectField label="Khối công ty" value={tplCompany} onChange={(e) => setTplCompany(e.target.value)}>
            {COMPANIES.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
          </SelectField>
          <TextField label="Từ ngày" type="date" value={fromDate} onChange={(e) => setFromDate(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={toDate} onChange={(e) => setToDate(e.target.value)} />
          <PrimaryButton onClick={handleTemplateExport}><Download size={15} /> Xuất theo mẫu</PrimaryButton>
        </div>
      </Card>

      <Card className="p-4">
        <p className="font-semibold text-slate-800 text-sm mb-1">Xuất mẫu "Thứ hạng Gung Ho"</p>
        <p className="text-xs text-slate-500 mb-3">Xếp hạng nhân viên (Đại sứ) trong khối công ty đã chọn ở trên theo điểm thi đua TD, khoảng thời gian dùng chung với mục phía trên.</p>
        <PrimaryButton onClick={handleRankingExport}><Download size={15} /> Xuất theo mẫu</PrimaryButton>
      </Card>

      <Card className="p-4">
        <p className="font-semibold text-slate-800 text-sm mb-1">Xuất mẫu "Kết quả Gung Ho chi tiết công ty theo thời gian"</p>
        <p className="text-xs text-slate-500 mb-3">Chọn 1 đơn vị/chi nhánh và năm cần xem — chia theo 12 tháng.</p>
        <div className="flex flex-wrap items-end gap-3">
          <SelectField label="Đơn vị / chi nhánh" value={tplUnit} onChange={(e) => setTplUnit(e.target.value)}>
            {ALL_BRANCHES.map((b) => <option key={b.name} value={b.name}>{b.name}</option>)}
          </SelectField>
          <TextField label="Năm" type="number" value={tplYear} onChange={(e) => setTplYear(e.target.value)} />
          <PrimaryButton onClick={handleTimeTemplateExport}><Download size={15} /> Xuất theo mẫu</PrimaryButton>
        </div>
      </Card>

      <div className="grid lg:grid-cols-2 gap-5">
        <RevenueTrendChart orders={viewOrders} title={viewCompany ? `Xu hướng doanh thu — ${viewCompany}` : "Xu hướng doanh thu 6 tháng gần đây (toàn tập đoàn)"} />
        <ProductMixPieChart orders={viewOrders} title={viewCompany ? `Tỉ trọng sản phẩm — ${viewCompany}` : "Tỉ trọng doanh thu theo sản phẩm (toàn tập đoàn)"} />
      </div>

      <div className="grid lg:grid-cols-2 gap-5">
        {viewCompany ? (
          <LeaderBoard orders={viewOrders} groupKeyFn={(o) => o.store} title={`Xếp hạng theo cửa hàng / chi nhánh — ${viewCompany}`} icon={Store} />
        ) : (
          <LeaderBoard orders={viewOrders} groupKeyFn={(o) => o.company} title="Xếp hạng theo khối công ty" icon={Building2} />
        )}
        <LeaderBoard orders={viewOrders} groupKeyFn={(o) => o.product} title={viewCompany ? `Xếp hạng sản phẩm bán chạy — ${viewCompany}` : "Xếp hạng sản phẩm bán chạy"} icon={ShoppingBag} />
      </div>
      <GunghoLeaderBoard
        orders={viewOrders}
        groupKeyFn={(o) => o.createdByName}
        title={viewCompany ? `Bảng xếp hạng Gungho — nhân viên ${viewCompany}` : "Bảng xếp hạng Gungho (toàn tập đoàn, theo điểm thi đua)"}
        icon={Award}
      />
    </div>
  );
}

// ---------------------------------------------------------------------------
// KẾ TOÁN
// ---------------------------------------------------------------------------

function AccountingCard({ order, onConfirm, onReject }) {
  if (
    XE_MAY_SPECIAL_PRODUCTS.includes(order.product) ||
    OTO_SPECIAL_PRODUCTS.includes(order.product) ||
    HTC_SPECIAL_PRODUCTS.includes(order.product) ||
    VYC_SPECIAL_PRODUCTS.includes(order.product) ||
    VTNN_SPECIAL_PRODUCTS.includes(order.product)
  ) {
    return <XeMaySpecialAccountingCard order={order} onConfirm={onConfirm} onReject={onReject} />;
  }
  return <GenericAccountingCard order={order} onConfirm={onConfirm} onReject={onReject} />;
}

function InvoiceInfoStrip({ order }) {
  if (order.company !== HTC_COMPANY_NAME || (!order.invoiceName && !order.invoiceNumber)) return null;
  return (
    <div className="flex flex-wrap items-center gap-x-4 gap-y-1 bg-sky-50 border border-sky-200 rounded-xl px-3 py-2 mb-3 text-xs text-sky-800">
      <span className="font-medium flex items-center gap-1"><FileText size={12} /> CSKH đã gửi (HTC)</span>
      <span>Tên xuất HĐ: <span className="font-medium">{order.invoiceName || "—"}</span></span>
      <span>Số HĐ: <span className="font-medium">{order.invoiceNumber || "—"}</span></span>
      <span>Số tiền: <span className="font-medium">{fmtMoney(order.totalAmount)}</span></span>
      <span>Chiết khấu: <span className="font-medium">{fmtMoney(order.discountAmount)}</span></span>
    </div>
  );
}

// ---- Sản phẩm thông thường (giữ nguyên form cũ) ----

function GenericAccountingCard({ order, onConfirm, onReject }) {
  const [amount, setAmount] = useState("");
  const [discount, setDiscount] = useState(0);
  const [commission, setCommission] = useState(0);
  const [quantity, setQuantity] = useState(1);
  const [note, setNote] = useState("");
  const finalAmount = Math.max((Number(amount) || 0) - Number(discount || 0), 0);
  const [amountError, setAmountError] = useState("");

  const category = PRODUCT_CATEGORY[order.product];
  const needsQuantity = COUNT_CATEGORIES.has(category);

  const handleConfirm = () => {
    if (!amount || Number(amount) <= 0) {
      setAmountError("Vui lòng nhập số tiền đơn hàng.");
      return;
    }
    setAmountError("");
    onConfirm(order.id, { amount, discountAmount: discount, commissionAmount: commission, quantity: needsQuantity ? quantity : 1, note });
  };

  return (
    <Card className="p-4">
      <div className="flex flex-wrap items-start justify-between gap-3 mb-3">
        <div>
          <p className="font-medium text-slate-800">{order.customerName}</p>
          <p className="text-sm text-slate-500 mt-1">{order.product} · <Store size={12} className="inline -mt-0.5" /> {order.store}</p>
          <p className="text-xs text-slate-400 mt-1">Đại sứ: {order.createdByName} · Người chăm sóc: {order.assignedHandlerName}</p>
          {order.handlerNote && <p className="text-xs text-slate-500 mt-1 italic">Ghi chú CSKH: {order.handlerNote}</p>}
        </div>
      </div>
      <InvoiceInfoStrip order={order} />
      <div className="grid sm:grid-cols-3 gap-3">
        <MoneyField label="Số tiền đơn hàng (đ)" value={amount} onChange={setAmount} />
        <MoneyField label="Chiết khấu (đ)" value={discount} onChange={setDiscount} />
        <MoneyField label="Hoa hồng đại sứ (đ)" value={commission} onChange={setCommission} />
        {needsQuantity && (
          <TextField label={CATEGORY_LABELS[category]} type="number" min="0" required value={quantity} onChange={(e) => setQuantity(e.target.value)} />
        )}
      </div>
      <div className="flex items-center justify-between bg-slate-50 rounded-xl px-3 py-2 mt-3">
        <span className="text-sm text-slate-500">Thành tiền sau chiết khấu</span>
        <span className="text-sm font-semibold text-slate-800">{fmtMoney(finalAmount)}</span>
      </div>
      {amountError && (
        <p className="text-sm text-rose-600 flex items-center gap-1.5 mt-2"><AlertCircle size={14} /> {amountError}</p>
      )}
      <div className="mt-3">
        <TextAreaField label="Ghi chú kế toán" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ghi chú giao dịch, hình thức thanh toán..." />
      </div>
      <div className="flex flex-wrap gap-2 mt-3">
        <PrimaryButton onClick={handleConfirm}>
          <CheckCircle2 size={15} /> Xác nhận thanh toán
        </PrimaryButton>
        <DangerButton onClick={() => onReject(order.id, note)}>
          <XCircle size={15} /> Không thanh toán
        </DangerButton>
      </div>
    </Card>
  );
}

// ---- Xe máy, xe đạp/máy điện: thưởng theo mức giảm giá/xe ----
function XeMayForm({ order, onConfirm, onReject, note, setNote }) {
  const [vehicleType, setVehicleType] = useState("so_dien");
  const [quantity, setQuantity] = useState(1);
  const [unitPrice, setUnitPrice] = useState("");
  const [discountPerUnit, setDiscountPerUnit] = useState(0);
  const [error, setError] = useState("");

  const qty = Number(quantity) || 0;
  const price = Number(unitPrice) || 0;
  const discPerUnit = Number(discountPerUnit) || 0;
  const totalAmount = qty * price;
  const totalDiscount = qty * discPerUnit;
  const rewardPerUnit = computeXeMayRewardPerUnit(vehicleType, discPerUnit);
  const totalCommission = qty * rewardPerUnit;

  const handleConfirm = () => {
    if (qty <= 0 || price <= 0) { setError("Vui lòng nhập số lượng và giá hợp lệ."); return; }
    setError("");
    onConfirm(order.id, { amount: totalAmount, discountAmount: totalDiscount, commissionAmount: totalCommission, quantity: qty, note });
  };

  return (
    <>
      <div className="grid sm:grid-cols-2 gap-3">
        <SelectField label="Loại xe" value={vehicleType} onChange={(e) => setVehicleType(e.target.value)}>
          <option value="so_dien">Xe số, xe điện (ngưỡng giảm giá 300.000đ/xe)</option>
          <option value="ga_con">Xe ga, xe côn (ngưỡng giảm giá 400.000đ/xe)</option>
        </SelectField>
        <TextField label="Số lượng xe" type="number" min="1" required value={quantity} onChange={(e) => setQuantity(e.target.value)} />
        <TextField label="Giá / xe (đ)" type="number" min="0" required value={unitPrice} onChange={(e) => setUnitPrice(e.target.value)} placeholder="0" />
        <TextField label="Chiết khấu / xe (đ)" type="number" min="0" value={discountPerUnit} onChange={(e) => setDiscountPerUnit(e.target.value)} />
      </div>
      <div className="grid sm:grid-cols-3 gap-2 mt-3">
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">Tổng tiền hàng</p>
          <p className="text-sm font-semibold text-slate-800">{fmtMoney(totalAmount)}</p>
        </div>
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">Thưởng/xe (tự tính)</p>
          <p className="text-sm font-semibold text-slate-800">{fmtMoney(rewardPerUnit)}</p>
        </div>
        <div className="bg-amber-50 rounded-xl px-3 py-2">
          <p className="text-xs text-amber-700">Tổng hoa hồng</p>
          <p className="text-sm font-semibold text-amber-800">{fmtMoney(totalCommission)}</p>
        </div>
      </div>
      {error && <p className="text-sm text-rose-600 flex items-center gap-1.5 mt-2"><AlertCircle size={14} /> {error}</p>}
      <div className="mt-3">
        <TextAreaField label="Ghi chú kế toán" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ghi chú giao dịch..." />
      </div>
      <div className="flex flex-wrap gap-2 mt-3">
        <PrimaryButton onClick={handleConfirm}><CheckCircle2 size={15} /> Xác nhận thanh toán</PrimaryButton>
        <DangerButton onClick={() => onReject(order.id, note)}><XCircle size={15} /> Không thanh toán</DangerButton>
      </div>
    </>
  );
}

// ---- Bảo hiểm xe máy: thưởng theo thời hạn hợp đồng ----
function BaoHiemXeMayForm({ order, onConfirm, onReject, note, setNote }) {
  const [years, setYears] = useState(1);
  const [quantity, setQuantity] = useState(1);
  const [unitPrice, setUnitPrice] = useState("");
  const [error, setError] = useState("");

  const qty = Number(quantity) || 0;
  const price = Number(unitPrice) || 0;
  const totalAmount = qty * price;
  const rewardPerUnit = baoHiemRewardPerUnit(Number(years));
  const totalCommission = qty * rewardPerUnit;

  const handleConfirm = () => {
    if (qty <= 0 || price <= 0) { setError("Vui lòng nhập số lượng và giá hợp lệ."); return; }
    setError("");
    onConfirm(order.id, { amount: totalAmount, discountAmount: 0, commissionAmount: totalCommission, quantity: qty, note });
  };

  return (
    <>
      <div className="grid sm:grid-cols-3 gap-3">
        <SelectField label="Thời hạn bảo hiểm" value={years} onChange={(e) => setYears(Number(e.target.value))}>
          <option value={1}>1 năm — thưởng 15.000đ/bảo hiểm</option>
          <option value={2}>2 năm — thưởng 20.000đ/bảo hiểm</option>
          <option value={3}>3 năm — thưởng 25.000đ/bảo hiểm</option>
        </SelectField>
        <TextField label="Số lượng bảo hiểm" type="number" min="1" required value={quantity} onChange={(e) => setQuantity(e.target.value)} />
        <TextField label="Giá / bảo hiểm (đ)" type="number" min="0" required value={unitPrice} onChange={(e) => setUnitPrice(e.target.value)} placeholder="0" />
      </div>
      <div className="grid sm:grid-cols-2 gap-2 mt-3">
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">Tổng tiền hàng</p>
          <p className="text-sm font-semibold text-slate-800">{fmtMoney(totalAmount)}</p>
        </div>
        <div className="bg-amber-50 rounded-xl px-3 py-2">
          <p className="text-xs text-amber-700">Tổng hoa hồng</p>
          <p className="text-sm font-semibold text-amber-800">{fmtMoney(totalCommission)}</p>
        </div>
      </div>
      {error && <p className="text-sm text-rose-600 flex items-center gap-1.5 mt-2"><AlertCircle size={14} /> {error}</p>}
      <div className="mt-3">
        <TextAreaField label="Ghi chú kế toán" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ghi chú giao dịch..." />
      </div>
      <div className="flex flex-wrap gap-2 mt-3">
        <PrimaryButton onClick={handleConfirm}><CheckCircle2 size={15} /> Xác nhận thanh toán</PrimaryButton>
        <DangerButton onClick={() => onReject(order.id, note)}><XCircle size={15} /> Không thanh toán</DangerButton>
      </div>
    </>
  );
}

// ---- Doanh thu theo tỷ lệ %: dùng chung cho Phụ tùng bán lẻ / Dịch vụ sửa chữa xe máy (5%), Phụ kiện ô tô (9%) ----
function ServiceRevenueForm({ order, onConfirm, onReject, note, setNote, ratePercent = 5 }) {
  const [amount, setAmount] = useState("");
  const [hasDiscount, setHasDiscount] = useState(false);
  const [discount, setDiscount] = useState(0);
  const [error, setError] = useState("");

  const revenue = Number(amount) || 0;
  const totalCommission = hasDiscount ? 0 : Math.round(revenue * (ratePercent / 100));
  const totalDiscount = hasDiscount ? Number(discount) || 0 : 0;

  const handleConfirm = () => {
    if (revenue <= 0) { setError("Vui lòng nhập doanh thu hợp lệ."); return; }
    setError("");
    onConfirm(order.id, { amount: revenue, discountAmount: totalDiscount, commissionAmount: totalCommission, quantity: 1, note });
  };

  return (
    <>
      <div className="grid sm:grid-cols-2 gap-3">
        <TextField label="Doanh thu (đ)" type="number" min="0" required value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" />
        <label className="flex items-center gap-2 mt-6">
          <input type="checkbox" checked={hasDiscount} onChange={(e) => setHasDiscount(e.target.checked)} className="w-4 h-4 accent-teal-800" />
          <span className="text-sm text-slate-700">Đơn hàng có giảm giá</span>
        </label>
        {hasDiscount && (
          <TextField label="Số tiền giảm giá (đ)" type="number" min="0" value={discount} onChange={(e) => setDiscount(e.target.value)} />
        )}
      </div>
      <div className="grid sm:grid-cols-2 gap-2 mt-3">
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">{hasDiscount ? "Có giảm giá — chỉ ghi nhận chỉ tiêu" : `Không giảm giá — thưởng ${ratePercent}% doanh thu`}</p>
        </div>
        <div className={`rounded-xl px-3 py-2 ${hasDiscount ? "bg-slate-50" : "bg-amber-50"}`}>
          <p className={`text-xs ${hasDiscount ? "text-slate-500" : "text-amber-700"}`}>Tổng hoa hồng</p>
          <p className={`text-sm font-semibold ${hasDiscount ? "text-slate-600" : "text-amber-800"}`}>{fmtMoney(totalCommission)}</p>
        </div>
      </div>
      {error && <p className="text-sm text-rose-600 flex items-center gap-1.5 mt-2"><AlertCircle size={14} /> {error}</p>}
      <div className="mt-3">
        <TextAreaField label="Ghi chú kế toán" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ghi chú giao dịch..." />
      </div>
      <div className="flex flex-wrap gap-2 mt-3">
        <PrimaryButton onClick={handleConfirm}><CheckCircle2 size={15} /> Xác nhận thanh toán</PrimaryButton>
        <DangerButton onClick={() => onReject(order.id, note)}><XCircle size={15} /> Không thanh toán</DangerButton>
      </div>
    </>
  );
}

// ---- Xe mới ô tô: thưởng 900.000đ/xe theo mức giảm giá ngoài chính sách ----
function OTOMoiForm({ order, onConfirm, onReject, note, setNote }) {
  const [quantity, setQuantity] = useState(1);
  const [unitPrice, setUnitPrice] = useState("");
  const [discountPerUnit, setDiscountPerUnit] = useState(0);
  const [error, setError] = useState("");

  const qty = Number(quantity) || 0;
  const price = Number(unitPrice) || 0;
  const discPerUnit = Number(discountPerUnit) || 0;
  const totalAmount = qty * price;
  const totalDiscount = qty * discPerUnit;
  const rewardPerUnit = otoRewardPerUnit(discPerUnit);
  const totalCommission = qty * rewardPerUnit;
  const tierLabel = discPerUnit <= 0 ? "100% (không giảm giá ngoài chính sách)" : discPerUnit <= 2000000 ? "70% (giảm giá ≤ 2.000.000đ)" : "50% (giảm giá > 2.000.000đ)";

  const handleConfirm = () => {
    if (qty <= 0 || price <= 0) { setError("Vui lòng nhập số lượng và giá hợp lệ."); return; }
    setError("");
    onConfirm(order.id, { amount: totalAmount, discountAmount: totalDiscount, commissionAmount: totalCommission, quantity: qty, note });
  };

  return (
    <>
      <div className="grid sm:grid-cols-3 gap-3">
        <TextField label="Số lượng xe" type="number" min="1" required value={quantity} onChange={(e) => setQuantity(e.target.value)} />
        <TextField label="Giá / xe (đ)" type="number" min="0" required value={unitPrice} onChange={(e) => setUnitPrice(e.target.value)} placeholder="0" />
        <TextField label="Giảm giá ngoài chính sách / xe (đ)" type="number" min="0" value={discountPerUnit} onChange={(e) => setDiscountPerUnit(e.target.value)} />
      </div>
      <div className="grid sm:grid-cols-3 gap-2 mt-3">
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">Tổng tiền hàng</p>
          <p className="text-sm font-semibold text-slate-800">{fmtMoney(totalAmount)}</p>
        </div>
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">Mức hưởng</p>
          <p className="text-sm font-semibold text-slate-800">{tierLabel}</p>
        </div>
        <div className="bg-amber-50 rounded-xl px-3 py-2">
          <p className="text-xs text-amber-700">Tổng hoa hồng</p>
          <p className="text-sm font-semibold text-amber-800">{fmtMoney(totalCommission)}</p>
        </div>
      </div>
      {error && <p className="text-sm text-rose-600 flex items-center gap-1.5 mt-2"><AlertCircle size={14} /> {error}</p>}
      <div className="mt-3">
        <TextAreaField label="Ghi chú kế toán" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ghi chú giao dịch..." />
      </div>
      <div className="flex flex-wrap gap-2 mt-3">
        <PrimaryButton onClick={handleConfirm}><CheckCircle2 size={15} /> Xác nhận thanh toán</PrimaryButton>
        <DangerButton onClick={() => onReject(order.id, note)}><XCircle size={15} /> Không thanh toán</DangerButton>
      </div>
    </>
  );
}

// ---- Bảo hiểm ô tô: hoa hồng = 82% x mức chiết khấu chính sách (THT thông báo) x doanh thu ----
function BaoHiemOTOForm({ order, onConfirm, onReject, note, setNote }) {
  const [amount, setAmount] = useState("");
  const [policyRate, setPolicyRate] = useState("");
  const [hasDiscount, setHasDiscount] = useState(false);
  const [discount, setDiscount] = useState(0);
  const [error, setError] = useState("");

  const revenue = Number(amount) || 0;
  const totalCommission = baoHiemOTOReward(revenue, hasDiscount, policyRate);
  const totalDiscount = hasDiscount ? Number(discount) || 0 : 0;

  const handleConfirm = () => {
    if (revenue <= 0) { setError("Vui lòng nhập doanh thu bảo hiểm hợp lệ."); return; }
    if (!hasDiscount && !policyRate) { setError("Vui lòng nhập mức chiết khấu theo chính sách THT."); return; }
    setError("");
    onConfirm(order.id, { amount: revenue, discountAmount: totalDiscount, commissionAmount: totalCommission, quantity: 1, note });
  };

  return (
    <>
      <div className="grid sm:grid-cols-2 gap-3">
        <TextField label="Doanh thu bảo hiểm (đ)" type="number" min="0" required value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" />
        <TextField label="Mức chiết khấu theo chính sách THT (%)" type="number" min="0" step="0.1" value={policyRate} onChange={(e) => setPolicyRate(e.target.value)} placeholder="Do THT thông báo theo từng thời điểm" disabled={hasDiscount} />
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={hasDiscount} onChange={(e) => setHasDiscount(e.target.checked)} className="w-4 h-4 accent-teal-800" />
          <span className="text-sm text-slate-700">Đơn hàng có giảm giá</span>
        </label>
        {hasDiscount && (
          <TextField label="Số tiền giảm giá (đ)" type="number" min="0" value={discount} onChange={(e) => setDiscount(e.target.value)} />
        )}
      </div>
      <div className="grid sm:grid-cols-2 gap-2 mt-3">
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">{hasDiscount ? "Có giảm giá — chỉ ghi nhận chỉ tiêu" : "Hoa hồng = 82% × mức chiết khấu chính sách × doanh thu"}</p>
        </div>
        <div className={`rounded-xl px-3 py-2 ${hasDiscount ? "bg-slate-50" : "bg-amber-50"}`}>
          <p className={`text-xs ${hasDiscount ? "text-slate-500" : "text-amber-700"}`}>Tổng hoa hồng</p>
          <p className={`text-sm font-semibold ${hasDiscount ? "text-slate-600" : "text-amber-800"}`}>{fmtMoney(totalCommission)}</p>
        </div>
      </div>
      {error && <p className="text-sm text-rose-600 flex items-center gap-1.5 mt-2"><AlertCircle size={14} /> {error}</p>}
      <div className="mt-3">
        <TextAreaField label="Ghi chú kế toán" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ghi chú giao dịch..." />
      </div>
      <div className="flex flex-wrap gap-2 mt-3">
        <PrimaryButton onClick={handleConfirm}><CheckCircle2 size={15} /> Xác nhận thanh toán</PrimaryButton>
        <DangerButton onClick={() => onReject(order.id, note)}><XCircle size={15} /> Không thanh toán</DangerButton>
      </div>
    </>
  );
}

// ---- Doanh thu × % cố định, không phân biệt giảm giá — dùng cho Đặc sản địa phương / Phòng nghỉ / Tiệc ----
function FlatRevenueForm({ order, onConfirm, onReject, note, setNote, ratePercent }) {
  const isHTCOrder = order.company === HTC_COMPANY_NAME;
  const [amount, setAmount] = useState(isHTCOrder && order.totalAmount ? String(order.totalAmount) : "");
  const [discount, setDiscount] = useState(isHTCOrder ? (order.discountAmount || 0) : 0);
  const [error, setError] = useState("");

  const revenue = Number(amount) || 0;
  const finalAmount = Math.max(revenue - (Number(discount) || 0), 0);
  const totalCommission = Math.round(finalAmount * (ratePercent / 100));

  const handleConfirm = () => {
    if (revenue <= 0) { setError("Vui lòng nhập doanh thu hợp lệ."); return; }
    setError("");
    onConfirm(order.id, { amount: revenue, discountAmount: Number(discount) || 0, commissionAmount: totalCommission, quantity: 1, note });
  };

  return (
    <>
      {isHTCOrder && order.totalAmount > 0 && (
        <p className="text-xs text-sky-700 mb-2">Doanh thu & chiết khấu đã tự điền theo thông tin CSKH gửi lên — kiểm tra lại và chỉnh sửa nếu cần.</p>
      )}
      <div className="grid sm:grid-cols-3 gap-3">
        <MoneyField label="Doanh thu (đ)" value={amount} onChange={setAmount} />
        <MoneyField label="Chiết khấu (đ)" value={discount} onChange={setDiscount} />
        <div className="bg-amber-50 rounded-xl px-3 py-2 flex flex-col justify-center">
          <p className="text-xs text-amber-700">Hoa hồng ({ratePercent}% sau chiết khấu)</p>
          <p className="text-sm font-semibold text-amber-800">{fmtMoney(totalCommission)}</p>
        </div>
      </div>
      {error && <p className="text-sm text-rose-600 flex items-center gap-1.5 mt-2"><AlertCircle size={14} /> {error}</p>}
      <div className="mt-3">
        <TextAreaField label="Ghi chú kế toán" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ghi chú giao dịch..." />
      </div>
      <div className="flex flex-wrap gap-2 mt-3">
        <PrimaryButton onClick={handleConfirm}><CheckCircle2 size={15} /> Xác nhận thanh toán</PrimaryButton>
        <DangerButton onClick={() => onReject(order.id, note)}><XCircle size={15} /> Không thanh toán</DangerButton>
      </div>
    </>
  );
}

// ---- Vé máy bay: thưởng theo số lượng vé, phân biệt khách lẻ / khách đoàn ----
function VeMayBayForm({ order, onConfirm, onReject, note, setNote }) {
  const [customerType, setCustomerType] = useState("le");
  const [quantity, setQuantity] = useState(1);
  const [unitPrice, setUnitPrice] = useState("");
  const [error, setError] = useState("");

  const qty = Number(quantity) || 0;
  const price = Number(unitPrice) || 0;
  const totalAmount = qty * price;
  const rewardPerUnit = veMayBayRewardPerUnit(customerType);
  const totalCommission = qty * rewardPerUnit;

  const handleConfirm = () => {
    if (qty <= 0) { setError("Vui lòng nhập số lượng vé hợp lệ."); return; }
    setError("");
    onConfirm(order.id, { amount: totalAmount, discountAmount: 0, commissionAmount: totalCommission, quantity: qty, note });
  };

  return (
    <>
      <div className="grid sm:grid-cols-3 gap-3">
        <SelectField label="Loại khách" value={customerType} onChange={(e) => setCustomerType(e.target.value)}>
          <option value="le">Khách lẻ — 15.000đ/vé</option>
          <option value="doan">Khách đoàn thông thường — 10.000đ/vé</option>
        </SelectField>
        <TextField label="Số lượng vé" type="number" min="1" required value={quantity} onChange={(e) => setQuantity(e.target.value)} />
        <TextField label="Giá / vé (đ)" type="number" min="0" value={unitPrice} onChange={(e) => setUnitPrice(e.target.value)} placeholder="0" />
      </div>
      <div className="grid sm:grid-cols-2 gap-2 mt-3">
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">Tổng tiền vé</p>
          <p className="text-sm font-semibold text-slate-800">{fmtMoney(totalAmount)}</p>
        </div>
        <div className="bg-amber-50 rounded-xl px-3 py-2">
          <p className="text-xs text-amber-700">Tổng hoa hồng</p>
          <p className="text-sm font-semibold text-amber-800">{fmtMoney(totalCommission)}</p>
        </div>
      </div>
      {error && <p className="text-sm text-rose-600 flex items-center gap-1.5 mt-2"><AlertCircle size={14} /> {error}</p>}
      <div className="mt-3">
        <TextAreaField label="Ghi chú kế toán" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ghi chú giao dịch..." />
      </div>
      <div className="flex flex-wrap gap-2 mt-3">
        <PrimaryButton onClick={handleConfirm}><CheckCircle2 size={15} /> Xác nhận thanh toán</PrimaryButton>
        <DangerButton onClick={() => onReject(order.id, note)}><XCircle size={15} /> Không thanh toán</DangerButton>
      </div>
    </>
  );
}

// ---- Tour nội địa: 500.000đ/hợp đồng thành công ----
function TourForm({ order, onConfirm, onReject, note, setNote }) {
  const [quantity, setQuantity] = useState(1);
  const [unitPrice, setUnitPrice] = useState("");
  const [error, setError] = useState("");

  const qty = Number(quantity) || 0;
  const price = Number(unitPrice) || 0;
  const totalAmount = qty * price;
  const totalCommission = qty * TOUR_REWARD_PER_CONTRACT;

  const handleConfirm = () => {
    if (qty <= 0) { setError("Vui lòng nhập số lượng hợp đồng hợp lệ."); return; }
    setError("");
    onConfirm(order.id, { amount: totalAmount, discountAmount: 0, commissionAmount: totalCommission, quantity: qty, note });
  };

  return (
    <>
      <div className="grid sm:grid-cols-2 gap-3">
        <TextField label="Số lượng hợp đồng thành công" type="number" min="1" required value={quantity} onChange={(e) => setQuantity(e.target.value)} />
        <TextField label="Giá trị / hợp đồng (đ)" type="number" min="0" value={unitPrice} onChange={(e) => setUnitPrice(e.target.value)} placeholder="0" />
      </div>
      <div className="grid sm:grid-cols-2 gap-2 mt-3">
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">Tổng doanh thu</p>
          <p className="text-sm font-semibold text-slate-800">{fmtMoney(totalAmount)}</p>
        </div>
        <div className="bg-amber-50 rounded-xl px-3 py-2">
          <p className="text-xs text-amber-700">Tổng hoa hồng (500.000đ/hợp đồng)</p>
          <p className="text-sm font-semibold text-amber-800">{fmtMoney(totalCommission)}</p>
        </div>
      </div>
      {error && <p className="text-sm text-rose-600 flex items-center gap-1.5 mt-2"><AlertCircle size={14} /> {error}</p>}
      <div className="mt-3">
        <TextAreaField label="Ghi chú kế toán" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ghi chú giao dịch..." />
      </div>
      <div className="flex flex-wrap gap-2 mt-3">
        <PrimaryButton onClick={handleConfirm}><CheckCircle2 size={15} /> Xác nhận thanh toán</PrimaryButton>
        <DangerButton onClick={() => onReject(order.id, note)}><XCircle size={15} /> Không thanh toán</DangerButton>
      </div>
    </>
  );
}

function XeMaySpecialAccountingCard({ order, onConfirm, onReject }) {
  const [note, setNote] = useState("");
  const isOTO = OTO_SPECIAL_PRODUCTS.includes(order.product);
  const isHTC = HTC_SPECIAL_PRODUCTS.includes(order.product);
  const isVYC = VYC_SPECIAL_PRODUCTS.includes(order.product);
  const isVTNN = VTNN_SPECIAL_PRODUCTS.includes(order.product);
  const groupLabel = isOTO ? "ô tô" : isHTC ? "HTC" : isVYC ? "VYC" : isVTNN ? "Vật tư nông nghiệp" : "xe máy";
  return (
    <Card className="p-4">
      <div className="flex flex-wrap items-start justify-between gap-3 mb-3">
        <div>
          <p className="font-medium text-slate-800">{order.customerName}</p>
          <p className="text-sm text-slate-500 mt-1">{order.product} · <Store size={12} className="inline -mt-0.5" /> {order.store}</p>
          <p className="text-xs text-slate-400 mt-1">Đại sứ: {order.createdByName} · Người chăm sóc: {order.assignedHandlerName}</p>
          {order.handlerNote && <p className="text-xs text-slate-500 mt-1 italic">Ghi chú CSKH: {order.handlerNote}</p>}
        </div>
        <Badge className="bg-teal-50 text-teal-700 border-teal-200">Quy định thưởng Khối {groupLabel}</Badge>
      </div>
      <InvoiceInfoStrip order={order} />
      {isHTC && order.company === HTC_COMPANY_NAME && order.updatedAt && (() => {
        const daysElapsed = (Date.now() - new Date(order.updatedAt).getTime()) / 86400000;
        const daysLeft = Math.ceil(HTC_AUTO_CLOSE_DAYS - daysElapsed);
        const overdue = daysLeft <= 0;
        return (
          <p className={`text-xs mb-3 -mt-2 flex items-center gap-1 ${overdue ? "text-rose-600 font-medium" : daysLeft <= 5 ? "text-amber-600" : "text-slate-400"}`}>
            <Clock size={12} />
            {overdue
              ? `Đã quá ${HTC_AUTO_CLOSE_DAYS} ngày chưa thanh toán — sẽ tự động đóng (ghi nhận doanh thu, không hoa hồng) ở lần tải lại tiếp theo.`
              : `Còn ${daysLeft} ngày trước khi tự động đóng nếu chưa xác nhận thanh toán (quy định riêng khối HTC).`}
          </p>
        );
      })()}
      {order.product === P.XE_MAY && <XeMayForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} />}
      {order.product === P.BAO_HIEM_XE_MAY && <BaoHiemXeMayForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} />}
      {order.product === P.PHU_TUNG && <ServiceRevenueForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} ratePercent={5} />}
      {order.product === P.SUA_CHUA_XE_MAY && <ServiceRevenueForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} ratePercent={5} />}
      {order.product === P.O_TO && <OTOMoiForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} />}
      {order.product === P.PHU_KIEN_O_TO && <ServiceRevenueForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} ratePercent={9} />}
      {order.product === P.BAO_HIEM_O_TO && <BaoHiemOTOForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} />}
      {order.product === P.SUA_CHUA_O_TO && <ServiceRevenueForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} ratePercent={5} />}
      {(order.product === P.DAC_SAN || order.product === P.PHONG_NGHI || order.product === P.TIEC) && (
        <FlatRevenueForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} ratePercent={3} />
      )}
      {order.product === P.VE_MAY_BAY && <VeMayBayForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} />}
      {order.product === P.TOUR && <TourForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} />}
      {VYC_SPECIAL_PRODUCTS.includes(order.product) && (
        <FlatRevenueForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} ratePercent={5} />
      )}
      {VTNN_SPECIAL_PRODUCTS.includes(order.product) && (
        <FlatRevenueForm order={order} onConfirm={onConfirm} onReject={onReject} note={note} setNote={setNote} ratePercent={2} />
      )}
    </Card>
  );
}

function KeToanChoXacNhan({ currentUser, orders, onConfirm, onReject }) {
  const isAdmin = currentUser.role === "admin";
  const [storeFilter, setStoreFilter] = useState("");
  const effectiveStore = isAdmin ? storeFilter : currentUser.store;

  const pending = orders
    .filter((o) => o.status === "cho_ke_toan")
    .filter((o) => !effectiveStore || o.store === effectiveStore)
    .sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));

  return (
    <div>
      <div className="flex items-start justify-between gap-3 flex-wrap mb-1">
        <SectionTitle icon={ClipboardCheck} title="Đơn hàng chờ xác nhận thanh toán" subtitle={effectiveStore ? `${pending.length} đơn tại ${effectiveStore}` : `${pending.length} đơn đang chờ (tất cả chi nhánh)`} />
        {isAdmin && (
          <div className="w-full sm:w-64">
            <SelectField label="Lọc theo chi nhánh" value={storeFilter} onChange={(e) => setStoreFilter(e.target.value)}>
              <option value="">— Tất cả chi nhánh —</option>
              {ALL_BRANCHES.map((b) => <option key={b.name} value={b.name}>{b.name}</option>)}
            </SelectField>
          </div>
        )}
      </div>
      {pending.length === 0 ? (
        <EmptyState icon={ClipboardCheck} text="Không có đơn hàng nào chờ xác nhận." />
      ) : (
        <div className="space-y-3">
          {pending.map((o) => <AccountingCard key={o.id} order={o} onConfirm={onConfirm} onReject={onReject} />)}
        </div>
      )}
    </div>
  );
}

function KeToanLichSu({ currentUser, orders }) {
  const isAdmin = currentUser.role === "admin";
  const [storeFilter, setStoreFilter] = useState("");
  const effectiveStore = isAdmin ? storeFilter : currentUser.store;

  const done = orders
    .filter((o) => ["da_thanh_toan", "khong_thanh_toan"].includes(o.status))
    .filter((o) => !effectiveStore || o.store === effectiveStore)
    .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
  const totalRevenue = done.filter((o) => o.status === "da_thanh_toan").reduce((s, o) => s + (o.finalAmount ?? o.totalAmount), 0);
  const totalCommission = done.reduce((s, o) => s + (o.commissionAmount || 0), 0);

  const handleExport = () => {
    const rows = done.map((o) => ({
      "Mã đơn hàng": o.orderCode || o.id,
      "Khách hàng": o.customerName,
      "SĐT": o.customerPhone,
      "Khối công ty": o.company,
      "Cửa hàng / chi nhánh": o.store,
      "Sản phẩm": o.product,
      "Đại sứ": o.createdByName,
      "Trạng thái": STATUS_META[o.status]?.label || o.status,
      "Số tiền đơn hàng": o.totalAmount || 0,
      "Chiết khấu": o.discountAmount || 0,
      "Thành tiền": o.finalAmount ?? o.totalAmount ?? 0,
      "Hoa hồng đại sứ": o.commissionAmount || 0,
      "Ghi chú kế toán": o.accountantNote || "",
      "Ngày xử lý": fmtDate(o.updatedAt),
    }));
    exportToExcel([{ name: "Lịch sử giao dịch", rows }], `LichSuGiaoDich_${Date.now()}.xlsx`);
  };

  return (
    <div className="space-y-5">
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <SectionTitle icon={Landmark} title="Lịch sử giao dịch" subtitle={effectiveStore || (isAdmin ? "Tất cả chi nhánh" : "")} />
        <GhostButton onClick={handleExport}><Download size={15} /> Xuất Excel</GhostButton>
      </div>
      {isAdmin && (
        <div className="sm:w-64">
          <SelectField label="Lọc theo chi nhánh" value={storeFilter} onChange={(e) => setStoreFilter(e.target.value)}>
            <option value="">— Tất cả chi nhánh —</option>
            {ALL_BRANCHES.map((b) => <option key={b.name} value={b.name}>{b.name}</option>)}
          </SelectField>
        </div>
      )}
      <div className="grid grid-cols-2 lg:grid-cols-3 gap-3">
        <MetricCard label="Tổng doanh thu" value={fmtMoney(totalRevenue)} icon={TrendingUp} accent="teal" />
        <MetricCard label="Tổng hoa hồng đã chi" value={fmtMoney(totalCommission)} icon={Wallet} accent="amber" />
        <MetricCard label="Số giao dịch" value={done.length} icon={ClipboardList} accent="indigo" />
      </div>
      {done.length === 0 ? (
        <EmptyState icon={Landmark} text="Chưa có giao dịch nào được xử lý." />
      ) : (
        <div className="space-y-2">
          {done.map((o) => (
            <OrderRow key={o.id} order={o} showCommission />
          ))}
        </div>
      )}
    </div>
  );
}
