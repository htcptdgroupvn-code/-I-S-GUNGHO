import React, { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";
import {
  Users, ClipboardList, BarChart3, Bell, LogOut, CheckCircle2, XCircle,
  Send, UserPlus, ShoppingBag, TrendingUp, Award, Store, Wallet,
  ArrowRightLeft, RefreshCw, Phone, MapPin, Plus, ChevronRight, Inbox,
  ClipboardCheck, Building2, Landmark, AlertCircle, X, Clock, Download
} from "lucide-react";

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

async function fetchAll() {
  const [emp, cust, ord, notif] = await Promise.all([
    supabase.from("employees").select("*"),
    supabase.from("customers").select("*").order("created_at", { ascending: false }),
    supabase.from("orders").select("*").order("created_at", { ascending: false }),
    supabase.from("notifications").select("*").order("created_at", { ascending: false }),
  ]);
  if (emp.error) console.error("fetch employees error", emp.error);
  if (cust.error) console.error("fetch customers error", cust.error);
  if (ord.error) console.error("fetch orders error", ord.error);
  if (notif.error) console.error("fetch notifications error", notif.error);
  return {
    employees: (emp.data || []).map(mapEmployee),
    customers: (cust.data || []).map(mapCustomer),
    orders: (ord.data || []).map(mapOrder),
    notifications: (notif.data || []).map(mapNotification),
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
  return <div className={`bg-white rounded-xl border border-slate-200 shadow-sm ${className}`}>{children}</div>;
}

function SectionTitle({ icon: Icon, title, subtitle }) {
  return (
    <div className="flex items-start gap-3 mb-4">
      <div className="w-10 h-10 rounded-lg bg-teal-800 text-white flex items-center justify-center shrink-0">
        <Icon size={20} />
      </div>
      <div>
        <h2 className="text-lg font-semibold text-slate-800">{title}</h2>
        {subtitle && <p className="text-sm text-slate-500">{subtitle}</p>}
      </div>
    </div>
  );
}

function EmptyState({ icon: Icon, text }) {
  return (
    <div className="flex flex-col items-center justify-center py-14 text-slate-400">
      <Icon size={34} className="mb-2 opacity-60" />
      <p className="text-sm">{text}</p>
    </div>
  );
}

function MetricCard({ label, value, icon: Icon, accent = "teal" }) {
  const accents = {
    teal: "text-teal-700 bg-teal-50",
    amber: "text-amber-700 bg-amber-50",
    indigo: "text-indigo-700 bg-indigo-50",
    rose: "text-rose-700 bg-rose-50",
  };
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 flex items-center gap-3">
      <div className={`w-10 h-10 rounded-lg flex items-center justify-center shrink-0 ${accents[accent]}`}>
        <Icon size={18} />
      </div>
      <div className="min-w-0">
        <p className="text-xs text-slate-500 truncate">{label}</p>
        <p className="text-lg font-semibold text-slate-800 truncate">{value}</p>
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
        className="w-full px-3 py-2 rounded-lg border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-teal-500"
      />
    </label>
  );
}

function SelectField({ label, children, ...props }) {
  return (
    <label className="block">
      <span className="block text-xs font-medium text-slate-600 mb-1">{label}</span>
      <select
        {...props}
        className="w-full px-3 py-2 rounded-lg border border-slate-300 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-teal-500"
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
        className="w-full px-3 py-2 rounded-lg border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-teal-500"
      />
    </label>
  );
}

function PrimaryButton({ children, className = "", ...props }) {
  return (
    <button
      {...props}
      className={`inline-flex items-center gap-1.5 px-4 py-2 rounded-lg bg-teal-800 text-white text-sm font-medium hover:bg-teal-900 active:scale-[0.98] transition disabled:opacity-40 disabled:cursor-not-allowed ${className}`}
    >
      {children}
    </button>
  );
}
function GhostButton({ children, className = "", ...props }) {
  return (
    <button
      {...props}
      className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg border border-slate-300 text-slate-700 text-sm font-medium hover:bg-slate-50 active:scale-[0.98] transition disabled:opacity-40 ${className}`}
    >
      {children}
    </button>
  );
}
function DangerButton({ children, className = "", ...props }) {
  return (
    <button
      {...props}
      className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg border border-rose-300 text-rose-700 text-sm font-medium hover:bg-rose-50 active:scale-[0.98] transition disabled:opacity-40 ${className}`}
    >
      {children}
    </button>
  );
}

function Toast({ toast }) {
  if (!toast) return null;
  return (
    <div className="fixed bottom-5 right-5 z-50 bg-slate-900 text-white text-sm px-4 py-2.5 rounded-lg shadow-lg animate-[fadeIn_.15s_ease-out]" style={{ marginBottom: "env(safe-area-inset-bottom)" }}>
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
    <div className="min-h-[600px] flex items-center justify-center bg-slate-50 px-4 py-10">
      <div className="w-full max-w-sm">
        <div className="text-center mb-8">
          <img src="/logo.png" alt="Phú Tài Đức Group" className="h-16 mx-auto mb-3 object-contain" />
          <h1 className="text-2xl font-semibold text-slate-800">Gungho Sales</h1>
          <p className="text-sm text-slate-500 mt-1">Đăng nhập bằng mã nhân viên</p>
        </div>
        <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-4">
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
        className="relative w-9 h-9 rounded-lg border border-slate-200 flex items-center justify-center hover:bg-slate-50"
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
    USERS.length = 0;
    USERS.push(...data.employees);
    setCustomers(data.customers);
    setOrders(data.orders);
    setNotifications(data.notifications);
    forceRerender((n) => n + 1);
    return data;
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

  const forwardToAccounting = async (orderId, note) => {
    const order = orders.find((o) => o.id === orderId);
    const newHistory = [...order.history, `${fmtDate(new Date().toISOString())} — ${currentUser.name} chuyển đơn cho kế toán (khách đồng ý mua)`];
    await updateOrder(orderId, { status: "cho_ke_toan", handler_note: note, history: newHistory, updated_at: new Date().toISOString() });
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

  const BASE_NAV = [
    { key: "khach_hang", label: "Khách hàng", icon: Users },
    { key: "don_hang_ds", label: "Đơn hàng của tôi", icon: ClipboardList },
    { key: "bao_cao_ds", label: "Báo cáo", icon: BarChart3 },
  ];
  const NAV = {
    dai_su: BASE_NAV,
    xu_ly: [
      ...BASE_NAV,
      { key: "duoc_giao", label: "Đơn được giao", icon: Inbox },
      { key: "don_hang_cskh", label: "Đơn hàng CSKH", icon: ClipboardList },
      { key: "bao_cao_cskh", label: "Báo cáo CSKH", icon: BarChart3 },
    ],
    cht: [
      ...BASE_NAV,
      { key: "phan_cong", label: "Phân công", icon: ArrowRightLeft },
      { key: "bao_cao_cht", label: "Báo cáo doanh số", icon: BarChart3 },
    ],
    ke_toan: [
      ...BASE_NAV,
      { key: "cho_xac_nhan", label: "Chờ xác nhận", icon: ClipboardCheck },
      { key: "lich_su", label: "Lịch sử", icon: Wallet },
    ],
    admin: [
      ...BASE_NAV,
      { key: "duoc_giao", label: "Đơn được giao", icon: Inbox },
      { key: "don_hang_cskh", label: "Đơn hàng CSKH", icon: ClipboardList },
      { key: "bao_cao_cskh", label: "Báo cáo CSKH", icon: BarChart3 },
      { key: "phan_cong", label: "Phân công", icon: ArrowRightLeft },
      { key: "bao_cao_cht", label: "Báo cáo doanh số", icon: BarChart3 },
      { key: "cho_xac_nhan", label: "Chờ xác nhận", icon: ClipboardCheck },
      { key: "lich_su", label: "Lịch sử", icon: Wallet },
    ],
  };
  const navItems = NAV[currentUser.role];

  return (
    <div className="min-h-[600px] bg-slate-50">
      <style>{`@keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}`}</style>
      {/* top bar */}
      <div className="bg-white border-b border-slate-200 sticky top-0 z-30" style={{ paddingTop: "env(safe-area-inset-top)" }}>
        <div className="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between gap-3">
          <div className="flex items-center gap-2.5 min-w-0">
            <div className="w-9 h-9 rounded-lg bg-teal-800 text-white flex items-center justify-center shrink-0">
              <TrendingUp size={18} />
            </div>
            <div className="min-w-0">
              <p className="text-sm font-semibold text-slate-800 leading-tight">Gungho Sales</p>
              <p className="text-[11px] text-slate-400 leading-tight truncate">Theo dõi doanh thu &amp; đơn hàng</p>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <GhostButton onClick={refreshAll} className="!px-2.5" title="Làm mới dữ liệu">
              <RefreshCw size={14} />
            </GhostButton>
            <NotifBell notifications={notifications} currentUser={currentUser} onMarkRead={markRead} />
            <div className="flex items-center gap-2 pl-2 border-l border-slate-200">
              <div className="w-8 h-8 rounded-full bg-teal-800 text-white flex items-center justify-center text-xs font-semibold shrink-0">
                {currentUser.name.split(" ").slice(-1)[0][0]}
              </div>
              <div className="hidden md:block text-right leading-tight">
                <p className="text-xs font-medium text-slate-800">{currentUser.name}</p>
                <p className="text-[11px] text-slate-400">{currentUser.store || ROLE_META[currentUser.role].short}</p>
              </div>
              <button onClick={handleLogout} className="w-8 h-8 rounded-lg border border-slate-200 flex items-center justify-center hover:bg-slate-50" title="Đổi tài khoản">
                <LogOut size={14} className="text-slate-500" />
              </button>
            </div>
          </div>
        </div>
        {/* tabs */}
        <div className="max-w-6xl mx-auto px-4 flex gap-1 overflow-x-auto">
          {navItems.map((n) => (
            <button
              key={n.key}
              onClick={() => setTab(n.key)}
              className={`flex items-center gap-1.5 px-3.5 py-2.5 text-sm font-medium border-b-2 whitespace-nowrap transition ${
                tab === n.key ? "border-teal-800 text-teal-800" : "border-transparent text-slate-500 hover:text-slate-700"
              }`}
            >
              <n.icon size={15} /> {n.label}
            </button>
          ))}
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-4 py-6">
        {tab === "khach_hang" && (
          <DaiSuKhachHang currentUser={currentUser} customers={customers} onAdd={addCustomer} />
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
      </div>
      <Toast toast={toast} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// ĐẠI SỨ — Khách hàng
// ---------------------------------------------------------------------------

function DaiSuKhachHang({ currentUser, customers, onAdd }) {
  const [form, setForm] = useState({ name: "", phone: "", address: "" });
  const [showForm, setShowForm] = useState(false);
  const isAdmin = currentUser.role === "admin";
  const mine = isAdmin ? customers : customers.filter((c) => c.createdBy === currentUser.id);

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
      <div className="flex items-center justify-between mb-4">
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

      {mine.length === 0 ? (
        <EmptyState icon={Users} text="Chưa có khách hàng nào — bấm “Thêm khách hàng” để bắt đầu." />
      ) : isAdmin ? (
        <AdminCustomerGroups customers={mine} />
      ) : (
        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">
          {mine.map((c) => (
            <Card key={c.id} className="p-4">
              <div className="flex items-center justify-between gap-2">
                <p className="font-medium text-slate-800">{c.name}</p>
                {c.customerCode && <Badge className="bg-teal-50 text-teal-700 border-teal-200">{c.customerCode}</Badge>}
              </div>
              <p className="text-sm text-slate-500 flex items-center gap-1.5 mt-1"><Phone size={13} /> {c.phone}</p>
              {c.address && <p className="text-sm text-slate-500 flex items-center gap-1.5 mt-1"><MapPin size={13} /> {c.address}</p>}
              <p className="text-[11px] text-slate-400 mt-2">Thêm lúc {fmtDate(c.createdAt)}</p>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

function AdminCustomerGroups({ customers }) {
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
              <Card key={c.id} className="p-4">
                <div className="flex items-center justify-between gap-2">
                  <p className="font-medium text-slate-800">{c.name}</p>
                  {c.customerCode && <Badge className="bg-teal-50 text-teal-700 border-teal-200">{c.customerCode}</Badge>}
                </div>
                <p className="text-sm text-slate-500 flex items-center gap-1.5 mt-1"><Phone size={13} /> {c.phone}</p>
                {c.address && <p className="text-sm text-slate-500 flex items-center gap-1.5 mt-1"><MapPin size={13} /> {c.address}</p>}
                <p className="text-[11px] text-slate-400 mt-2">Thêm lúc {fmtDate(c.createdAt)}</p>
              </Card>
            ))}
          </div>
        </div>
      ))}
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
        <div className="mb-4 flex items-center gap-2 text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
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

function LeaderBoard({ orders, groupKeyFn, title, icon }) {
  const rows = buildRevenueLeaderboard(orders, groupKeyFn).slice(0, 8);
  return (
    <Card className="p-4">
      <SectionTitle icon={icon} title={title} />
      {rows.length === 0 ? (
        <EmptyState icon={Award} text="Chưa có dữ liệu xếp hạng." />
      ) : (
        <div className="space-y-2">
          {rows.map((r, i) => (
            <div key={r.name} className="flex items-center gap-3">
              <div className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-semibold shrink-0 ${i < 3 ? "bg-amber-100 text-amber-800" : "bg-slate-100 text-slate-500"}`}>
                {i + 1}
              </div>
              <div className="flex-1 min-w-0">
                <p className="text-sm font-medium text-slate-700 truncate">{r.name}</p>
                <div className="h-1.5 bg-slate-100 rounded-full mt-1 overflow-hidden">
                  <div className="h-full bg-teal-700 rounded-full" style={{ width: `${(r.revenue / rows[0].revenue) * 100}%` }} />
                </div>
              </div>
              <p className="text-sm font-semibold text-slate-800 shrink-0">{fmtMoney(r.revenue)}</p>
            </div>
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
    <Card className="p-4">
      <SectionTitle icon={icon} title={title} subtitle="TD = BX×0.5 + DVX/500.000 + BO×2 + DVO/1.000.000 + NH/1.000.000 + KS/500.000 + TO×20 + V×1 + VYC/300.000 + VT/500.000" />
      {rows.length === 0 ? (
        <EmptyState icon={Award} text="Chưa có dữ liệu xếp hạng." />
      ) : (
        <div className="space-y-2">
          {rows.map((r, i) => (
            <div key={r.name} className="flex items-center gap-3">
              <div className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-semibold shrink-0 ${i < 3 ? "bg-amber-100 text-amber-800" : "bg-slate-100 text-slate-500"}`}>
                {i + 1}
              </div>
              <div className="flex-1 min-w-0">
                <p className="text-sm font-medium text-slate-700 truncate">{r.name}</p>
                <div className="h-1.5 bg-slate-100 rounded-full mt-1 overflow-hidden">
                  <div className="h-full bg-amber-600 rounded-full" style={{ width: `${max > 0 ? (r.td / max) * 100 : 0}%` }} />
                </div>
              </div>
              <p className="text-sm font-semibold text-slate-800 shrink-0">{r.td.toFixed(1)} điểm</p>
            </div>
          ))}
        </div>
      )}
    </Card>
  );
}

function DaiSuBaoCao({ currentUser, orders }) {
  const isAdmin = currentUser.role === "admin";
  const mine = isAdmin ? orders : orders.filter((o) => o.createdBy === currentUser.id);
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
        <SectionTitle icon={BarChart3} title="Báo cáo tổng hợp" subtitle="Toàn bộ đơn hàng, hoa hồng và xếp hạng" />
        <GhostButton onClick={handleExport}><Download size={15} /> Xuất Excel</GhostButton>
      </div>
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

      <GunghoLeaderBoard orders={orders} groupKeyFn={(o) => o.createdByName} title="Bảng xếp hạng Gungho (theo điểm thi đua)" icon={Award} />
      <LeaderBoard orders={orders} groupKeyFn={(o) => o.company} title="Doanh thu theo khối công ty" icon={Building2} />
      <LeaderBoard orders={orders} groupKeyFn={(o) => o.product} title="Xếp hạng theo sản phẩm" icon={ShoppingBag} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// XỬ LÝ - CHĂM SÓC
// ---------------------------------------------------------------------------

function HandlerActionCard({ order, onConfirm, onForward, onDecline }) {
  const [note, setNote] = useState(order.handlerNote || "");
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
      <div className="flex flex-wrap gap-2 mt-3">
        {order.status === "cho_xu_ly" && (
          <PrimaryButton onClick={() => onConfirm(order.id, note)}><CheckCircle2 size={15} /> Xác nhận chăm sóc</PrimaryButton>
        )}
        {order.status === "dang_cham_soc" && (
          <>
            <PrimaryButton onClick={() => onForward(order.id, note)}><Send size={15} /> Khách đồng ý mua — Chuyển kế toán</PrimaryButton>
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
            className="w-full pl-9 pr-3 py-2.5 rounded-lg border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-teal-500"
          />
        </div>
        {isAdmin && (
          <div className="w-full sm:w-64">
            <select
              value={storeFilter}
              onChange={(e) => setStoreFilter(e.target.value)}
              className="w-full px-3 py-2.5 rounded-lg border border-slate-300 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-teal-500"
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
      <Card className="p-4">
        <p className="font-semibold text-slate-800 text-sm mb-3">Tình trạng xử lý theo sản phẩm</p>
        {rows.length === 0 ? (
          <EmptyState icon={ShoppingBag} text="Chưa có dữ liệu." />
        ) : (
          <div className="space-y-2">
            {rows.map(([p, c]) => (
              <div key={p} className="flex items-center justify-between text-sm py-1.5 border-b border-slate-50 last:border-0">
                <span className="text-slate-700">{p}</span>
                <Badge className="bg-slate-100 text-slate-600 border-slate-200">{c} khách</Badge>
              </div>
            ))}
          </div>
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
  const paid = orders.filter((o) => o.status === "da_thanh_toan");
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
        <SectionTitle icon={BarChart3} title="Báo cáo doanh số Gungho" subtitle="Toàn bộ khối công ty & chi nhánh" />
        <GhostButton onClick={handleExport}><Download size={15} /> Xuất Excel</GhostButton>
      </div>
      <div className="grid grid-cols-2 lg:grid-cols-3 gap-3">
        <MetricCard label="Tổng doanh thu" value={fmtMoney(revenue)} icon={Building2} accent="teal" />
        <MetricCard label="Đơn đã hoàn tất" value={paid.length} icon={CheckCircle2} accent="amber" />
        <MetricCard label="Tổng số đơn hàng" value={orders.length} icon={ShoppingBag} accent="indigo" />
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

      <LeaderBoard orders={orders} groupKeyFn={(o) => o.company} title="Xếp hạng theo khối công ty" icon={Building2} />
      <LeaderBoard orders={orders} groupKeyFn={(o) => o.store} title="Xếp hạng theo cửa hàng / chi nhánh" icon={Store} />
      <GunghoLeaderBoard orders={orders} groupKeyFn={(o) => o.createdByName} title="Bảng xếp hạng Gungho (theo điểm thi đua)" icon={Award} />
      <LeaderBoard orders={orders} groupKeyFn={(o) => o.product} title="Xếp hạng sản phẩm bán chạy" icon={ShoppingBag} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// KẾ TOÁN
// ---------------------------------------------------------------------------

function AccountingCard({ order, onConfirm, onReject }) {
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
      <div className="grid sm:grid-cols-3 gap-3">
        <TextField label="Số tiền đơn hàng (đ)" type="number" min="0" required value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" />
        <TextField label="Chiết khấu (đ)" type="number" min="0" value={discount} onChange={(e) => setDiscount(e.target.value)} />
        <TextField label="Hoa hồng đại sứ (đ)" type="number" min="0" value={commission} onChange={(e) => setCommission(e.target.value)} />
        {needsQuantity && (
          <TextField label={CATEGORY_LABELS[category]} type="number" min="0" required value={quantity} onChange={(e) => setQuantity(e.target.value)} />
        )}
      </div>
      <div className="flex items-center justify-between bg-slate-50 rounded-lg px-3 py-2 mt-3">
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
