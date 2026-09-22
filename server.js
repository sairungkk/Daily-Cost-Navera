const express = require('express');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { exec } = require('child_process');
const multer = require('multer');

// กำหนดการเก็บไฟล์ชั่วคราวใน Memory เพื่อประมวลผลทันที (จำกัดขนาดไฟล์อัปโหลดไม่เกิน 25 MB)
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_FILE = path.join(__dirname, 'Revenue_Cost_Data.xlsx');
const OUTLETS_FILE = path.join(__dirname, 'outlets.json');
// เป้าหมายงบประมาณ (budget) เก็บแยกไฟล์ JSON ตามรูปแบบเดียวกับ outlets.json/topsale.json
// โครงสร้าง: { "2026-09": { "<Outlet>": { monthly: {...}, daily: [...] } } }
const BUDGET_FILE = path.join(__dirname, 'budget.json');
// หมวดของตัวเลข budget — revenue 4 หมวด + cost แยก 4 ประเภท + covers แยก 4 ประเภท (ตามหมวดที่ระบบกรอกรายวัน)
const BUDGET_METRICS = ['food', 'bev', 'wine', 'other', 'costFood', 'costBev', 'costWine', 'costOther', 'coversFood', 'coversBev', 'coversWine', 'coversOther'];

app.use(express.json());
// serve เฉพาะโฟลเดอร์ public (ไม่ serve ทั้งโฟลเดอร์โปรเจกต์ เพราะจะทำให้ดาวน์โหลด
// ไฟล์ฐานข้อมูล Excel, backups, server.js ได้โดยไม่ต้องล็อกอิน)
app.use(express.static(path.join(__dirname, 'public')));

// ===== Authentication (ตรวจสิทธิ์ฝั่งเซิร์ฟเวอร์) =====
const STAFF_PASSWORD = process.env.STAFF_PASSWORD || 'Aap#navera';
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'admin001405';
const TOKEN_TTL_MS = 12 * 60 * 60 * 1000; // token มีอายุ 12 ชั่วโมง
const activeTokens = new Map(); // token -> { role, expiresAt }

function createToken(role) {
  const token = crypto.randomBytes(32).toString('hex');
  activeTokens.set(token, { role, expiresAt: Date.now() + TOKEN_TTL_MS });
  return token;
}

function getTokenInfo(token) {
  if (!token) return null;
  const info = activeTokens.get(token);
  if (!info) return null;
  if (Date.now() > info.expiresAt) {
    activeTokens.delete(token);
    return null;
  }
  return info;
}

// เก็บกวาด token ที่หมดอายุทุก 1 ชั่วโมง
setInterval(() => {
  const now = Date.now();
  for (const [token, info] of activeTokens) {
    if (now > info.expiresAt) activeTokens.delete(token);
  }
}, 60 * 60 * 1000).unref();

// Middleware ตรวจสิทธิ์: requireAuth() = staff/admin, requireAuth('admin') = เฉพาะผู้ดูแลระบบ
function requireAuth(minRole) {
  return (req, res, next) => {
    const info = getTokenInfo(req.get('x-auth-token'));
    if (!info) {
      return res.status(401).json({ success: false, message: 'ไม่ได้เข้าสู่ระบบ หรือหมดเวลาการใช้งาน กรุณาเข้าสู่ระบบใหม่' });
    }
    if (minRole === 'admin' && info.role !== 'admin') {
      return res.status(403).json({ success: false, message: 'ต้องการสิทธิ์ผู้ดูแลระบบ (Admin) เท่านั้น' });
    }
    req.authRole = info.role;
    next();
  };
}

app.post('/api/login', (req, res) => {
  const password = String((req.body && req.body.password) || '');
  let role = null;
  if (password && password === ADMIN_PASSWORD) role = 'admin';
  else if (password && password === STAFF_PASSWORD) role = 'staff';

  if (!role) {
    return res.status(401).json({ success: false, message: 'รหัสผ่านไม่ถูกต้อง' });
  }
  res.json({ success: true, role, token: createToken(role) });
});

app.post('/api/logout', (req, res) => {
  const token = req.get('x-auth-token');
  if (token) activeTokens.delete(token);
  res.json({ success: true });
});

// ฟังก์ชันสำหรับอ่านรายชื่อห้องอาหารจากไฟล์ outlets.json
function getOutlets() {
  const defaultOutlets = [
    'Reveira House',
    'Junsai',
    'Tapanyaki',
    'Terrace',
    'Juniper',
    'Marcele',
    'Canteen'
  ];
  if (!fs.existsSync(OUTLETS_FILE)) {
    fs.writeFileSync(OUTLETS_FILE, JSON.stringify(defaultOutlets, null, 2));
    return defaultOutlets;
  }
  try {
    const data = fs.readFileSync(OUTLETS_FILE, 'utf8');
    const parsed = JSON.parse(data);
    return Array.isArray(parsed) && parsed.length > 0 ? parsed : defaultOutlets;
  } catch (err) {
    return defaultOutlets;
  }
}

// API ดึงรายชื่อห้องอาหาร
app.get('/api/outlets', requireAuth(), (req, res) => {
  try {
    const outlets = getOutlets();
    res.json({ success: true, outlets });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// API สำหรับเพิ่ม/แก้ไข/ลบ รายชื่อห้องอาหาร
app.post('/api/outlets/update', requireAuth('admin'), (req, res) => {
  try {
    const { outlets } = req.body;
    if (!Array.isArray(outlets)) {
      return res.status(400).json({ success: false, message: 'Invalid data format' });
    }
    fs.writeFileSync(OUTLETS_FILE, JSON.stringify(outlets, null, 2));
    res.json({ success: true, message: 'อัปเดตรายชื่อห้องอาหารสำเร็จ' });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get('/', (req, res) => {
  const possiblePaths = [
    path.join(__dirname, 'public', 'index.html'),
    path.join(__dirname, 'public', 'index.htm'),
    path.join(__dirname, 'public', 'index'),
    path.join(__dirname, 'index.html'),
    path.join(__dirname, 'Revenue_Cost_Data.html')
  ];

  for (const p of possiblePaths) {
    if (fs.existsSync(p) && fs.statSync(p).size > 0) {
      return res.sendFile(p);
    }
  }

  res.send('<h3>⚠️ ไม่พบไฟล์ HTML หน้าแรก</h3>');
});

function getFormattedTimestamp() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
}

// ฟังก์ชันแปลงวันที่รองรับทั้ง Date object, DD-MM-YY, DD-MM-YYYY, DD/MM/YYYY, YYYY-MM-DD
function getCellDateStr(cellValue) {
  if (!cellValue) return '';
  if (cellValue instanceof Date) {
    const year = cellValue.getFullYear();
    const month = String(cellValue.getMonth() + 1).padStart(2, '0');
    const day = String(cellValue.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  if (typeof cellValue === 'object' && cellValue.text) {
    cellValue = cellValue.text;
  } else if (typeof cellValue === 'object' && cellValue.result) {
    cellValue = cellValue.result;
  }
  let s = String(cellValue).trim().split('T')[0];
  if (s.includes('-')) {
    const p = s.split('-');
    if (p.length === 3) {
      if (p[0].length === 4) return `${p[0]}-${p[1].padStart(2, '0')}-${p[2].padStart(2, '0')}`;
      const y = p[2].length === 2 ? '20' + p[2] : p[2];
      return `${y}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
    }
  }
  if (s.includes('/')) {
    const p = s.split('/');
    if (p.length === 3) {
      if (p[0].length === 4) return `${p[0]}-${p[1].padStart(2, '0')}-${p[2].padStart(2, '0')}`;
      const y = p[2].length === 2 ? '20' + p[2] : p[2];
      return `${y}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
    }
  }
  return s.substring(0, 10);
}

// ฟังก์ชันสำรองข้อมูลไฟล์ Excel อัตโนมัติ
function backupExcelFile() {
  if (!fs.existsSync(EXCEL_FILE)) return;
  
  const backupDir = path.join(__dirname, 'backups');
  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir, { recursive: true });
  }

  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const timestampStr = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}-${pad(now.getSeconds())}`;
  
  const backupFileName = `Revenue_Cost_Data_backup_${timestampStr}.xlsx`;
  const backupFilePath = path.join(backupDir, backupFileName);

  try {
    fs.copyFileSync(EXCEL_FILE, backupFilePath);
    console.log(`✅ สำรองข้อมูลสำเร็จ: ${backupFileName}`);

    const MAX_BACKUPS = 30;
    const backups = fs.readdirSync(backupDir)
      .filter(f => f.startsWith('Revenue_Cost_Data_backup_') && f.endsWith('.xlsx'))
      .sort();
    while (backups.length > MAX_BACKUPS) {
      const oldest = backups.shift();
      try {
        fs.unlinkSync(path.join(backupDir, oldest));
        console.log(`🗑️ ลบไฟล์สำรองเก่า: ${oldest}`);
      } catch (cleanupErr) {
        console.error('❌ ลบไฟล์สำรองเก่าไม่สำเร็จ:', cleanupErr.message);
        break;
      }
    }
  } catch (err) {
    console.error('❌ สำรองข้อมูลไม่สำเร็จ:', err.message);
  }
}

async function saveWorkbook(workbook) {
  const tmpFile = EXCEL_FILE + '.tmp';
  await workbook.xlsx.writeFile(tmpFile);
  try {
    fs.renameSync(tmpFile, EXCEL_FILE);
  } catch (err) {
    try { fs.unlinkSync(tmpFile); } catch (cleanupErr) { }
    throw new Error('ไฟล์ Excel ถูกล็อกอยู่ (อาจเปิดค้างไว้ในโปรแกรม Excel) กรุณาปิดไฟล์แล้วบันทึกใหม่อีกครั้ง');
  }
}

const DATA_COLUMNS = [
  { header: 'Date', key: 'date', width: 14 },
  { header: 'Outlet', key: 'outlet', width: 18 },
  { header: 'Categories', key: 'category', width: 16 },
  { header: 'Meal', key: 'meal', width: 14 },
  { header: 'Revenue', key: 'revenue', width: 16 },
  { header: 'Cover', key: 'cover', width: 12 },
  { header: 'Cost', key: 'cost', width: 16 },
  { header: 'Timestamp', key: 'timestamp', width: 22 }
];

const PUR_COLUMNS = [
  { header: 'Date', key: 'date', width: 14 },
  { header: 'Outlet', key: 'outlet', width: 18 },
  { header: 'Direct Purchase Food (THB)', key: 'purchaseFood', width: 22 },
  { header: 'Direct Purchase Bev (THB)', key: 'purchaseBev', width: 22 },
  { header: 'Store Issue Food (THB)', key: 'issueFood', width: 22 },
  { header: 'Store Issue Wine (THB)', key: 'issueWine', width: 22 },
  { header: 'Store Issue Bev (THB)', key: 'issueBev', width: 22 },
  { header: 'Store Issue Other (THB)', key: 'issueOther', width: 22 },
  { header: 'Timestamp', key: 'timestamp', width: 22 }
];

const COST_ADJ_COLUMNS = [
  { header: 'Date', key: 'date', width: 14 },
  { header: 'Outlet', key: 'outlet', width: 18 },
  { header: 'Transfer In Food (THB)', key: 'tInFood', width: 17 },
  { header: 'Transfer In Bev (THB)', key: 'tInBev', width: 17 },
  { header: 'Transfer In Wine (THB)', key: 'tInWine', width: 17 },
  { header: 'Transfer In Other (THB)', key: 'tInOther', width: 17 },
  { header: 'Transfer Out Food (THB)', key: 'tOutFood', width: 18 },
  { header: 'Transfer Out Bev (THB)', key: 'tOutBev', width: 18 },
  { header: 'Transfer Out Wine (THB)', key: 'tOutWine', width: 18 },
  { header: 'Transfer Out Other (THB)', key: 'tOutOther', width: 18 },
  { header: 'Credit Cost Food (THB)', key: 'crFood', width: 17 },
  { header: 'Credit Cost Bev (THB)', key: 'crBev', width: 17 },
  { header: 'Credit Cost Wine (THB)', key: 'crWine', width: 17 },
  { header: 'Credit Cost Other (THB)', key: 'crOther', width: 17 },
  { header: 'Timestamp', key: 'timestamp', width: 22 }
];

const HOTEL_STATS_COLUMNS = [
  { header: 'Date', key: 'date', width: 14 },
  { header: 'Room Available', key: 'roomAvailable', width: 16 },
  { header: 'Room Sold', key: 'roomSold', width: 14 },
  { header: 'Com./HU. Room', key: 'comHuRoom', width: 16 },
  { header: 'Total Guests', key: 'totalGuest', width: 14 },
  { header: 'Weather', key: 'weather', width: 14 },
  { header: 'Timestamp', key: 'timestamp', width: 22 }
];

const SUMMARY_COLUMNS = [
  { header: 'Date', key: 'date', width: 13 },
  { header: 'Outlet', key: 'outlet', width: 18 },
  { header: 'Meal', key: 'meal', width: 12 },
  { header: 'Food Rev', key: 'foodRev', width: 13 },
  { header: 'Bev Rev', key: 'bevRev', width: 13 },
  { header: 'Wine Rev', key: 'wineRev', width: 13 },
  { header: 'Other Rev', key: 'otherRev', width: 13 },
  { header: 'Total Rev', key: 'totalRev', width: 13 },
  { header: 'Cover', key: 'cover', width: 10 },
  { header: 'Food Cost', key: 'foodCost', width: 13 },
  { header: 'Bev Cost', key: 'bevCost', width: 13 },
  { header: 'Wine Cost', key: 'wineCost', width: 13 },
  { header: 'Other Cost', key: 'otherCost', width: 13 },
  { header: 'Total Cost', key: 'totalCost', width: 13 },
  { header: 'Direct Purchase Food', key: 'purchaseFood', width: 17 },
  { header: 'Direct Purchase Bev', key: 'purchaseBev', width: 17 },
  { header: 'Store Issue Food', key: 'issueFood', width: 15 },
  { header: 'Store Issue Wine', key: 'issueWine', width: 15 },
  { header: 'Store Issue Bev', key: 'issueBev', width: 15 },
  { header: 'Store Issue Other', key: 'issueOther', width: 15 },
  { header: 'Transfer In Food', key: 'tInFood', width: 14 },
  { header: 'Transfer In Bev', key: 'tInBev', width: 14 },
  { header: 'Transfer In Wine', key: 'tInWine', width: 14 },
  { header: 'Transfer In Other', key: 'tInOther', width: 14 },
  { header: 'Transfer Out Food', key: 'tOutFood', width: 15 },
  { header: 'Transfer Out Bev', key: 'tOutBev', width: 15 },
  { header: 'Transfer Out Wine', key: 'tOutWine', width: 15 },
  { header: 'Transfer Out Other', key: 'tOutOther', width: 15 },
  { header: 'Credit Cost Food', key: 'crFood', width: 14 },
  { header: 'Credit Cost Bev', key: 'crBev', width: 14 },
  { header: 'Credit Cost Wine', key: 'crWine', width: 14 },
  { header: 'Credit Cost Other', key: 'crOther', width: 14 },
  { header: 'Timestamp', key: 'timestamp', width: 22 }
];

const PIVOT_SUMMARY_COLUMNS = [
  { header: 'Outlet', key: 'outlet', width: 20 },
  { header: 'Total Food Rev', key: 'totalFoodRev', width: 15 },
  { header: 'Total Bev Rev', key: 'totalBevRev', width: 15 },
  { header: 'Total Wine Rev', key: 'totalWineRev', width: 15 },
  { header: 'Total Other Rev', key: 'totalOtherRev', width: 15 },
  { header: 'Grand Total Rev', key: 'totalRev', width: 16 },
  { header: 'Total Cover', key: 'totalCover', width: 12 },
  { header: 'Total Food Cost', key: 'totalFoodCost', width: 15 },
  { header: 'Total Bev Cost', key: 'totalBevCost', width: 15 },
  { header: 'Total Wine Cost', key: 'totalWineCost', width: 15 },
  { header: 'Total Other Cost', key: 'totalOtherCost', width: 15 },
  { header: 'Total Cost', key: 'totalCost', width: 15 },
  { header: 'Direct Purchase Food', key: 'totalPurchaseFood', width: 17 },
  { header: 'Direct Purchase Bev', key: 'totalPurchaseBev', width: 17 },
  { header: 'Store Issue Food', key: 'totalIssueFood', width: 15 },
  { header: 'Store Issue Wine', key: 'totalIssueWine', width: 15 },
  { header: 'Store Issue Bev', key: 'totalIssueBev', width: 15 },
  { header: 'Store Issue Other', key: 'totalIssueOther', width: 15 },
  { header: 'Transfer Net Food', key: 'totalTransferFood', width: 15 },
  { header: 'Transfer Net Bev', key: 'totalTransferBev', width: 15 },
  { header: 'Transfer Net Wine', key: 'totalTransferWine', width: 15 },
  { header: 'Transfer Net Other', key: 'totalTransferOther', width: 15 },
  { header: 'Credit Cost Food', key: 'totalCrFood', width: 14 },
  { header: 'Credit Cost Bev', key: 'totalCrBev', width: 14 },
  { header: 'Credit Cost Wine', key: 'totalCrWine', width: 14 },
  { header: 'Credit Cost Other', key: 'totalCrOther', width: 14 }
];

function setupSheetColumns(sheet, columns) {
  sheet.columns = columns;
  sheet.getRow(1).font = { bold: true };
}

function updateDynamicPivotSheet(workbook) {
  let pivotSheet = workbook.getWorksheet('Pivot_Summary');
  if (!pivotSheet) {
    pivotSheet = workbook.addWorksheet('Pivot_Summary');
  }
  setupSheetColumns(pivotSheet, PIVOT_SUMMARY_COLUMNS);

  while (pivotSheet.rowCount > 1) {
    pivotSheet.spliceRows(2, 1);
  }

  const outlets = getOutlets();

  outlets.forEach((outName, index) => {
    const rowNum = index + 2;
    pivotSheet.addRow({
      outlet: outName,
      totalFoodRev: { formula: `SUMIFS(Outlet_Summary!D:D, Outlet_Summary!B:B, A${rowNum})` },
      totalBevRev: { formula: `SUMIFS(Outlet_Summary!E:E, Outlet_Summary!B:B, A${rowNum})` },
      totalWineRev: { formula: `SUMIFS(Outlet_Summary!F:F, Outlet_Summary!B:B, A${rowNum})` },
      totalOtherRev: { formula: `SUMIFS(Outlet_Summary!G:G, Outlet_Summary!B:B, A${rowNum})` },
      totalRev: { formula: `SUMIFS(Outlet_Summary!H:H, Outlet_Summary!B:B, A${rowNum})` },
      totalCover: { formula: `SUMIFS(Outlet_Summary!I:I, Outlet_Summary!B:B, A${rowNum})` },
      totalFoodCost: { formula: `SUMIFS(Outlet_Summary!J:J, Outlet_Summary!B:B, A${rowNum})` },
      totalBevCost: { formula: `SUMIFS(Outlet_Summary!K:K, Outlet_Summary!B:B, A${rowNum})` },
      totalWineCost: { formula: `SUMIFS(Outlet_Summary!L:L, Outlet_Summary!B:B, A${rowNum})` },
      totalOtherCost: { formula: `SUMIFS(Outlet_Summary!M:M, Outlet_Summary!B:B, A${rowNum})` },
      totalCost: { formula: `SUMIFS(Outlet_Summary!N:N, Outlet_Summary!B:B, A${rowNum})` },
      totalPurchaseFood: { formula: `SUMIFS(Outlet_Summary!O:O, Outlet_Summary!B:B, A${rowNum})` },
      totalPurchaseBev: { formula: `SUMIFS(Outlet_Summary!P:P, Outlet_Summary!B:B, A${rowNum})` },
      totalIssueFood: { formula: `SUMIFS(Outlet_Summary!Q:Q, Outlet_Summary!B:B, A${rowNum})` },
      totalIssueWine: { formula: `SUMIFS(Outlet_Summary!R:R, Outlet_Summary!B:B, A${rowNum})` },
      totalIssueBev: { formula: `SUMIFS(Outlet_Summary!S:S, Outlet_Summary!B:B, A${rowNum})` },
      totalIssueOther: { formula: `SUMIFS(Outlet_Summary!T:T, Outlet_Summary!B:B, A${rowNum})` },
      totalTransferFood: { formula: `SUMIFS(Outlet_Summary!U:U, Outlet_Summary!B:B, A${rowNum})-SUMIFS(Outlet_Summary!Y:Y, Outlet_Summary!B:B, A${rowNum})` },
      totalTransferBev: { formula: `SUMIFS(Outlet_Summary!V:V, Outlet_Summary!B:B, A${rowNum})-SUMIFS(Outlet_Summary!Z:Z, Outlet_Summary!B:B, A${rowNum})` },
      totalTransferWine: { formula: `SUMIFS(Outlet_Summary!W:W, Outlet_Summary!B:B, A${rowNum})-SUMIFS(Outlet_Summary!AA:AA, Outlet_Summary!B:B, A${rowNum})` },
      totalTransferOther: { formula: `SUMIFS(Outlet_Summary!X:X, Outlet_Summary!B:B, A${rowNum})-SUMIFS(Outlet_Summary!AB:AB, Outlet_Summary!B:B, A${rowNum})` },
      totalCrFood: { formula: `SUMIFS(Outlet_Summary!AC:AC, Outlet_Summary!B:B, A${rowNum})` },
      totalCrBev: { formula: `SUMIFS(Outlet_Summary!AD:AD, Outlet_Summary!B:B, A${rowNum})` },
      totalCrWine: { formula: `SUMIFS(Outlet_Summary!AE:AE, Outlet_Summary!B:B, A${rowNum})` },
      totalCrOther: { formula: `SUMIFS(Outlet_Summary!AF:AF, Outlet_Summary!B:B, A${rowNum})` }
    });
  });

  const startRow = 2;
  const endRow = outlets.length + 1;
  const sumCols = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
  const totalRowValues = { outlet: 'Grand Total' };
  const pivotKeys = ['totalFoodRev', 'totalBevRev', 'totalWineRev', 'totalOtherRev', 'totalRev', 'totalCover', 'totalFoodCost', 'totalBevCost', 'totalWineCost', 'totalOtherCost', 'totalCost', 'totalPurchaseFood', 'totalPurchaseBev', 'totalIssueFood', 'totalIssueWine', 'totalIssueBev', 'totalIssueOther', 'totalTransferFood', 'totalTransferBev', 'totalTransferWine', 'totalTransferOther', 'totalCrFood', 'totalCrBev', 'totalCrWine', 'totalCrOther'];
  pivotKeys.forEach((k, i) => {
    const col = sumCols[i];
    totalRowValues[k] = { formula: `SUM(${col}${startRow}:${col}${endRow})` };
  });
  const totalRow = pivotSheet.addRow(totalRowValues);

  totalRow.font = { bold: true };
}

// ===== Migration: ย้ายชีตเก่า (Direct Purchase / Store Issue แบบรวม) ไปชุดฟิลด์ใหม่ =====
function readSheetHeaders(sheet) {
  const headers = [];
  if (!sheet) return headers;
  sheet.getRow(1).eachCell((cell, col) => {
    headers[col] = String(cell.value || '').trim();
  });
  return headers;
}

// เวอร์ชัน layout ของชีต Purchases_Issues: v1 = คอลัมน์รวมเดิม, v2 = แยก 4, v3 = แยก 6 (ปัจจุบัน)
function purSheetLayout(sheet) {
  const headers = readSheetHeaders(sheet);
  if (headers.includes('Store Issue Food (THB)')) return 'v3';
  if (headers.includes('Direct Purchase Food (THB)')) return 'v2';
  if (headers.some(h => h.includes('Direct Purchase') || h.includes('Store Issue'))) return 'v1';
  return 'v3'; // ชีตว่าง/ใหม่ ถือว่าเป็น layout ปัจจุบัน
}

// เวอร์ชัน layout ของชีต Cost_Adjustments: v1 = คอลัมน์รวมเดิม, v2 = แยก 4 หมวด (ปัจจุบัน)
function adjSheetLayout(sheet) {
  const headers = readSheetHeaders(sheet);
  if (headers.includes('Transfer In Food (THB)')) return 'v2';
  if (headers.some(h => h.includes('Transfer In') || h.includes('Credit Cost'))) return 'v1';
  return 'v2'; // ชีตว่าง/ใหม่ ถือว่าเป็น layout ปัจจุบัน
}

function migratePurchasesSheet(purSheet) {
  if (!purSheet || purSheet.rowCount <= 1) return false;
  const layout = purSheetLayout(purSheet);
  if (layout === 'v3') return false;

  // v1: Direct Purchase → Food, Store Issue → Bev | v2: ย้ายตำแหน่ง เติม Issue Food/Other = 0
  const oldRows = [];
  purSheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    let pf = 0, pb = 0, iw = 0, ib = 0;
    if (layout === 'v2') {
      pf = Number(row.getCell(3).value) || 0;
      pb = Number(row.getCell(4).value) || 0;
      iw = Number(row.getCell(5).value) || 0;
      ib = Number(row.getCell(6).value) || 0;
    } else {
      pf = Number(row.getCell(3).value) || 0;
      ib = Number(row.getCell(4).value) || 0;
    }
    oldRows.push({
      date: getCellDateStr(row.getCell(1).value),
      outlet: String(row.getCell(2).value || '').trim(),
      purchaseFood: pf,
      purchaseBev: pb,
      issueFood: 0,
      issueWine: iw,
      issueBev: ib,
      issueOther: 0,
      timestamp: row.getCell(layout === 'v2' ? 7 : 5).value || ''
    });
  });

  while (purSheet.rowCount > 1) purSheet.spliceRows(2, 1);
  setupSheetColumns(purSheet, PUR_COLUMNS);
  oldRows.forEach(r => purSheet.addRow(r));
  return true;
}

function migrateAdjustmentsSheet(adjSheet) {
  if (!adjSheet || adjSheet.rowCount <= 1) return false;
  if (adjSheetLayout(adjSheet) === 'v2') return false;

  // ชีตเก่าเป็นคอลัมน์รวม (ไม่ทราบหมวด) → แบ่งเท่าๆ กัน 4 หมวด ให้ผลรวมเท่าเดิม
  const oldRows = [];
  adjSheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const tIn = (Number(row.getCell(3).value) || 0) / 4;
    const tOut = (Number(row.getCell(4).value) || 0) / 4;
    const cr = (Number(row.getCell(5).value) || 0) / 4;
    oldRows.push({
      date: getCellDateStr(row.getCell(1).value),
      outlet: String(row.getCell(2).value || '').trim(),
      tInFood: tIn, tInBev: tIn, tInWine: tIn, tInOther: tIn,
      tOutFood: tOut, tOutBev: tOut, tOutWine: tOut, tOutOther: tOut,
      crFood: cr, crBev: cr, crWine: cr, crOther: cr,
      timestamp: row.getCell(6).value || ''
    });
  });

  while (adjSheet.rowCount > 1) adjSheet.spliceRows(2, 1);
  setupSheetColumns(adjSheet, COST_ADJ_COLUMNS);
  oldRows.forEach(r => adjSheet.addRow(r));
  return true;
}

// ตรวจว่ายอดปรับปรุง (คอลัมน์ 15-32) ยังซ้ำอยู่หลายแถวต่อวัน/outlet หรืออยู่ผิดแถว (ไม่ใช่ Dinner ทั้งที่มี Dinner)
function summaryNeedsAdjustmentDedup(summarySheet) {
  const groups = new Map();
  summarySheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const d = getCellDateStr(row.getCell(1).value);
    const out = String(row.getCell(2).value || '').trim();
    if (!d || !out) return;
    let adjSum = 0;
    for (let c = 15; c <= 32; c++) adjSum += Math.abs(Number(row.getCell(c).value) || 0);
    const key = `${d}___${out}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ meal: String(row.getCell(3).value || '').trim(), adjSum });
  });
  for (const rows of groups.values()) {
    const withAdj = rows.filter(r => r.adjSum > 0.005);
    if (withAdj.length > 1) return true;
    if (withAdj.length === 1) {
      const hasDinner = rows.some(r => r.meal === 'Dinner');
      if (hasDinner && withAdj[0].meal !== 'Dinner') return true;
    }
  }
  return false;
}

function migrateSummarySheet(summarySheet, dataSheet, purSheet, costAdjSheet, purAlreadyMigrated, adjAlreadyMigrated) {
  if (!summarySheet || summarySheet.rowCount <= 1) return false;
  const headers = readSheetHeaders(summarySheet);
  // header ใน Outlet_Summary คือ 'Transfer In Food' (ไม่มี (THB) ต่างจากชีต Cost_Adjustments)
  const isLatestLayout = headers.includes('Transfer In Food');
  if (!isLatestLayout && !headers.some(h => h.includes('Food Rev'))) return false; // ชีตว่าเปล่า
  // layout ล่าสุดอยู่แล้ว แต่ยอดปรับปรุงยังซ้ำหลายมื้อ (ข้อมูลเก่า) → rebuild ให้เหลือแถวเดียวที่ Dinner
  if (isLatestLayout && !summaryNeedsAdjustmentDedup(summarySheet)) return false;

  // รวมยอดจากชีต Data ใหม่ทั้งชีต เพื่อกู้คืน Wine/Other ที่ชุดเก่ารวมไว้ใน Bev
  const mealAgg = {};
  const pairsWithMeals = new Set();
  if (dataSheet) {
    dataSheet.eachRow((row, rowNum) => {
      if (rowNum === 1) return;
      const d = getCellDateStr(row.getCell(1).value);
      const out = String(row.getCell(2).value || '').trim();
      const cat = String(row.getCell(3).value || '').trim();
      const meal = String(row.getCell(4).value || '').trim() || 'General';
      const rev = Number(row.getCell(5).value) || 0;
      const cov = Number(row.getCell(6).value) || 0;
      const cost = Number(row.getCell(7).value) || 0;
      if (!d || !out) return;

      const pairKey = `${d}___${out}`;
      pairsWithMeals.add(pairKey);
      const key = `${pairKey}___${meal}`;
      if (!mealAgg[key]) {
        mealAgg[key] = { date: d, outlet: out, meal, foodRev: 0, bevRev: 0, wineRev: 0, otherRev: 0, cover: 0, foodCost: 0, bevCost: 0, wineCost: 0, otherCost: 0 };
      }
      const agg = mealAgg[key];
      agg.cover += cov;
      if (cat === 'RV-Food') {
        agg.foodRev += rev;
        agg.foodCost += cost;
      } else if (cat === 'RV-Wine') {
        agg.wineRev += rev;
        agg.wineCost += cost;
      } else if (cat === 'RV-Other') {
        agg.otherRev += rev;
        agg.otherCost += cost;
      } else {
        agg.bevRev += rev;
        agg.bevCost += cost;
      }
    });
  }

  // ยอด Direct Purchase / Store Issue (รองรับทุกเวอร์ชัน layout)
  const purMap = {};
  if (purSheet && purSheet.rowCount > 1) {
    const layout = purAlreadyMigrated ? 'v3' : purSheetLayout(purSheet);
    purSheet.eachRow((row, rowNum) => {
      if (rowNum === 1) return;
      const d = getCellDateStr(row.getCell(1).value);
      const out = String(row.getCell(2).value || '').trim();
      if (!d || !out) return;
      let pf = 0, pb = 0, iF = 0, iw = 0, ib = 0, iO = 0;
      if (layout === 'v3') {
        pf = Number(row.getCell(3).value) || 0;
        pb = Number(row.getCell(4).value) || 0;
        iF = Number(row.getCell(5).value) || 0;
        iw = Number(row.getCell(6).value) || 0;
        ib = Number(row.getCell(7).value) || 0;
        iO = Number(row.getCell(8).value) || 0;
      } else if (layout === 'v2') {
        pf = Number(row.getCell(3).value) || 0;
        pb = Number(row.getCell(4).value) || 0;
        iw = Number(row.getCell(5).value) || 0;
        ib = Number(row.getCell(6).value) || 0;
      } else {
        pf = Number(row.getCell(3).value) || 0;
        ib = Number(row.getCell(4).value) || 0;
      }
      purMap[`${d}___${out}`] = {
        purchaseFood: pf,
        purchaseBev: pb,
        issueFood: iF,
        issueWine: iw,
        issueBev: ib,
        issueOther: iO
      };
    });
  }

  const adjMap = {};
  if (costAdjSheet && costAdjSheet.rowCount > 1) {
    const layout = adjAlreadyMigrated ? 'v2' : adjSheetLayout(costAdjSheet);
    costAdjSheet.eachRow((row, rowNum) => {
      if (rowNum === 1) return;
      const d = getCellDateStr(row.getCell(1).value);
      const out = String(row.getCell(2).value || '').trim();
      if (!d || !out) return;
      if (layout === 'v2') {
        adjMap[`${d}___${out}`] = {
          tInFood: Number(row.getCell(3).value) || 0,
          tInBev: Number(row.getCell(4).value) || 0,
          tInWine: Number(row.getCell(5).value) || 0,
          tInOther: Number(row.getCell(6).value) || 0,
          tOutFood: Number(row.getCell(7).value) || 0,
          tOutBev: Number(row.getCell(8).value) || 0,
          tOutWine: Number(row.getCell(9).value) || 0,
          tOutOther: Number(row.getCell(10).value) || 0,
          crFood: Number(row.getCell(11).value) || 0,
          crBev: Number(row.getCell(12).value) || 0,
          crWine: Number(row.getCell(13).value) || 0,
          crOther: Number(row.getCell(14).value) || 0
        };
      } else {
        // คอลัมน์รวมเดิม → แบ่งเท่าๆ 4 หมวด
        const tIn = (Number(row.getCell(3).value) || 0) / 4;
        const tOut = (Number(row.getCell(4).value) || 0) / 4;
        const cr = (Number(row.getCell(5).value) || 0) / 4;
        adjMap[`${d}___${out}`] = {
          tInFood: tIn, tInBev: tIn, tInWine: tIn, tInOther: tIn,
          tOutFood: tOut, tOutBev: tOut, tOutWine: tOut, tOutOther: tOut,
          crFood: cr, crBev: cr, crWine: cr, crOther: cr
        };
      }
    });
  }

  const emptyPur = { purchaseFood: 0, purchaseBev: 0, issueFood: 0, issueWine: 0, issueBev: 0, issueOther: 0 };
  const emptyAdj = { tInFood: 0, tInBev: 0, tInWine: 0, tInOther: 0, tOutFood: 0, tOutBev: 0, tOutWine: 0, tOutOther: 0, crFood: 0, crBev: 0, crWine: 0, crOther: 0 };

  while (summarySheet.rowCount > 1) summarySheet.spliceRows(2, 1);
  setupSheetColumns(summarySheet, SUMMARY_COLUMNS);

  // จัดกลุ่มตามวัน/outlet — ยอดปรับปรุงใส่แค่แถวเดียวที่มื้อ Dinner (ไม่มี Dinner ไล่ Supper → Lunch → Breakfast → ABF)
  const mealsByPair = new Map();
  Object.values(mealAgg).forEach(agg => {
    const pk = `${agg.date}___${agg.outlet}`;
    if (!mealsByPair.has(pk)) mealsByPair.set(pk, []);
    mealsByPair.get(pk).push(agg);
  });

  mealsByPair.forEach((meals, pk) => {
    const pur = purMap[pk] || emptyPur;
    const adj = adjMap[pk] || emptyAdj;
    const carrierMeal = ['Dinner', 'Supper', 'Lunch', 'Breakfast', 'ABF'].find(m => meals.some(x => x.meal === m)) || meals[0].meal;
    meals.forEach(agg => {
      const isCarrier = agg.meal === carrierMeal;
      summarySheet.addRow({
        date: agg.date,
        outlet: agg.outlet,
        meal: agg.meal,
        foodRev: agg.foodRev,
        bevRev: agg.bevRev,
        wineRev: agg.wineRev,
        otherRev: agg.otherRev,
        totalRev: agg.foodRev + agg.bevRev + agg.wineRev + agg.otherRev,
        cover: agg.cover,
        foodCost: agg.foodCost,
        bevCost: agg.bevCost,
        wineCost: agg.wineCost,
        otherCost: agg.otherCost,
        totalCost: agg.foodCost + agg.bevCost + agg.wineCost + agg.otherCost,
        purchaseFood: isCarrier ? pur.purchaseFood : 0,
        purchaseBev: isCarrier ? pur.purchaseBev : 0,
        issueFood: isCarrier ? pur.issueFood : 0,
        issueWine: isCarrier ? pur.issueWine : 0,
        issueBev: isCarrier ? pur.issueBev : 0,
        issueOther: isCarrier ? pur.issueOther : 0,
        tInFood: isCarrier ? adj.tInFood : 0,
        tInBev: isCarrier ? adj.tInBev : 0,
        tInWine: isCarrier ? adj.tInWine : 0,
        tInOther: isCarrier ? adj.tInOther : 0,
        tOutFood: isCarrier ? adj.tOutFood : 0,
        tOutBev: isCarrier ? adj.tOutBev : 0,
        tOutWine: isCarrier ? adj.tOutWine : 0,
        tOutOther: isCarrier ? adj.tOutOther : 0,
        crFood: isCarrier ? adj.crFood : 0,
        crBev: isCarrier ? adj.crBev : 0,
        crWine: isCarrier ? adj.crWine : 0,
        crOther: isCarrier ? adj.crOther : 0,
        timestamp: getFormattedTimestamp()
      });
    });
  });

  // วัน/outlet ที่มีแต่ยอดปรับปรุงโดยไม่มีรายการขาย → เก็บเป็นแถว meal '-' เหมือนโค้ดเดิม
  const allPairs = new Set([...pairsWithMeals, ...Object.keys(purMap), ...Object.keys(adjMap)]);
  allPairs.forEach(pair => {
    if (pairsWithMeals.has(pair)) return;
    const pur = purMap[pair] || emptyPur;
    const adj = adjMap[pair] || emptyAdj;
    const adjSum = [...Object.values(pur), ...Object.values(adj)].reduce((s, v) => s + v, 0);
    if (adjSum === 0) return;
    const [d, out] = pair.split('___');
    summarySheet.addRow({
      date: d,
      outlet: out,
      meal: '-',
      foodRev: 0,
      bevRev: 0,
      wineRev: 0,
      otherRev: 0,
      totalRev: 0,
      cover: 0,
      foodCost: 0,
      bevCost: 0,
      wineCost: 0,
      otherCost: 0,
      totalCost: 0,
      purchaseFood: pur.purchaseFood,
      purchaseBev: pur.purchaseBev,
      issueFood: pur.issueFood,
      issueWine: pur.issueWine,
      issueBev: pur.issueBev,
      issueOther: pur.issueOther,
      tInFood: adj.tInFood,
      tInBev: adj.tInBev,
      tInWine: adj.tInWine,
      tInOther: adj.tInOther,
      tOutFood: adj.tOutFood,
      tOutBev: adj.tOutBev,
      tOutWine: adj.tOutWine,
      tOutOther: adj.tOutOther,
      crFood: adj.crFood,
      crBev: adj.crBev,
      crWine: adj.crWine,
      crOther: adj.crOther,
      timestamp: getFormattedTimestamp()
    });
  });

  sortSheetByDate(summarySheet);
  return true;
}

let initChain = Promise.resolve();
let sheetLayoutMigrated = false;

async function buildWorkbook() {
  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(EXCEL_FILE)) {
    await workbook.xlsx.readFile(EXCEL_FILE);
  }

  const ALL_OUTLETS = getOutlets();

  let dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');
  if (!dataSheet) {
    dataSheet = workbook.addWorksheet('Data');
  }

  ALL_OUTLETS.forEach(outletName => {
    if (!workbook.getWorksheet(outletName)) {
      workbook.addWorksheet(outletName);
    }
  });

  let purSheet = workbook.getWorksheet('Purchases_Issues');
  if (!purSheet) {
    purSheet = workbook.addWorksheet('Purchases_Issues');
  }

  let costAdjSheet = workbook.getWorksheet('Cost_Adjustments');
  if (!costAdjSheet) {
    costAdjSheet = workbook.addWorksheet('Cost_Adjustments');
  }

  let statsSheet = workbook.getWorksheet('Hotel_Stats');
  if (!statsSheet) {
    statsSheet = workbook.addWorksheet('Hotel_Stats');
  }

  let summarySheet = workbook.getWorksheet('Outlet_Summary');
  if (!summarySheet) {
    summarySheet = workbook.addWorksheet('Outlet_Summary');
  }

  // ตรวจและย้ายโครงสร้างชีตเก่าก่อนเขียน header ชุดใหม่ทับ
  const purchasesMigrated = migratePurchasesSheet(purSheet);
  const adjustmentsMigrated = migrateAdjustmentsSheet(costAdjSheet);
  const summaryMigrated = migrateSummarySheet(summarySheet, dataSheet, purSheet, costAdjSheet, purchasesMigrated, adjustmentsMigrated);
  sheetLayoutMigrated = purchasesMigrated || adjustmentsMigrated || summaryMigrated;
  if (sheetLayoutMigrated) {
    console.log(`🔄 Migrated old sheet layout (Purchases_Issues: ${purchasesMigrated}, Cost_Adjustments: ${adjustmentsMigrated}, Outlet_Summary: ${summaryMigrated})`);
  }

  setupSheetColumns(dataSheet, DATA_COLUMNS);

  ALL_OUTLETS.forEach(outletName => {
    setupSheetColumns(workbook.getWorksheet(outletName), DATA_COLUMNS);
  });

  setupSheetColumns(purSheet, PUR_COLUMNS);

  setupSheetColumns(costAdjSheet, COST_ADJ_COLUMNS);

  setupSheetColumns(statsSheet, HOTEL_STATS_COLUMNS);

  setupSheetColumns(summarySheet, SUMMARY_COLUMNS);

  updateDynamicPivotSheet(workbook);

  return workbook;
}

function initWorkbook() {
  // อ่าน/ย้ายโครงสร้างไฟล์ต้องรันทีละครั้ง กันเซฟพร้อมกันชนกัน
  initChain = initChain.then(buildWorkbook, buildWorkbook);
  return initChain;
}

function sortSheetByDate(sheet) {
  if (!sheet || sheet.rowCount <= 2) return;

  const rowsData = [];
  sheet.eachRow((row, rowNum) => {
    if (rowNum > 1) {
      rowsData.push(row.values);
    }
  });

  rowsData.sort((a, b) => String(a[1] || '').localeCompare(String(b[1] || '')));

  while (sheet.rowCount > 1) {
    sheet.spliceRows(2, 1);
  }

  rowsData.forEach(rowValues => {
    sheet.addRow(rowValues.slice(1));
  });
}

// ===== API สำหรับ Auto Import ทุกชีทในไฟล์ Excel พร้อมกัน =====
app.post('/api/import-excel', requireAuth('admin'), upload.single('file'), async (req, res) => {
  try {
    if (!req.file || !req.file.buffer) {
      return res.status(400).json({ success: false, message: 'กรุณาแนบไฟล์ Excel หรือ CSV' });
    }

    const uploadedWorkbook = new ExcelJS.Workbook();
    await uploadedWorkbook.xlsx.load(req.file.buffer);

    const targetWorkbook = await initWorkbook();
    const dataSheet = targetWorkbook.getWorksheet('Data') || targetWorkbook.getWorksheet('Revenue_Cost');
    const purSheet = targetWorkbook.getWorksheet('Purchases_Issues');
    const costAdjSheet = targetWorkbook.getWorksheet('Cost_Adjustments');
    const statsSheet = targetWorkbook.getWorksheet('Hotel_Stats');
    const summarySheet = targetWorkbook.getWorksheet('Outlet_Summary');
    const currentTimestamp = getFormattedTimestamp();

    let totalImportedRows = 0;
    const importedDateOutlets = new Set();
    const mealMapAggregator = {};

    uploadedWorkbook.worksheets.forEach(sheet => {
      const sheetName = sheet.name.trim().toLowerCase();
      if (sheet.rowCount <= 1) return;

      const headerRow = sheet.getRow(1);
      const colMap = {};
      headerRow.eachCell((cell, colNum) => {
        const val = String(cell.value || '').trim().toLowerCase();
        if (val.includes('date') || val.includes('วัน')) colMap.date = colNum;
        else if (val.includes('outlet') || val.includes('ห้องอาหาร') || val.includes('location')) colMap.outlet = colNum;
        else if (val.includes('cat') || val.includes('ประเภท') || val.includes('หมวด')) colMap.category = colNum;
        else if (val.includes('meal') || val.includes('shift') || val.includes('มื้อ')) colMap.meal = colNum;
        else if (val.includes('rev') || val.includes('ยอดขาย') || val.includes('amount')) colMap.revenue = colNum;
        else if (val.includes('cover') || val.includes('pax') || val.includes('จำนวนคน')) colMap.cover = colNum;
        else if (val.includes('cost') && !val.includes('cr')) colMap.cost = colNum;
        else if (val.includes('purchase food')) colMap.purchaseFood = colNum;
        else if (val.includes('purchase bev')) colMap.purchaseBev = colNum;
        else if (val.includes('issue food')) colMap.issueFood = colNum;
        else if (val.includes('issue wine')) colMap.issueWine = colNum;
        else if (val.includes('issue bev')) colMap.issueBev = colNum;
        else if (val.includes('issue other')) colMap.issueOther = colNum;
        else if (val.includes('purch') || val.includes('ซื้อ')) colMap.purchase = colNum;
        else if (val.includes('issue') || val.includes('เบิก')) colMap.issue = colNum;
        else if (val.includes('transfer in food')) colMap.tInFood = colNum;
        else if (val.includes('transfer in bev')) colMap.tInBev = colNum;
        else if (val.includes('transfer in wine')) colMap.tInWine = colNum;
        else if (val.includes('transfer in other')) colMap.tInOther = colNum;
        else if (val.includes('transfer in') || val.includes('โอนเข้า')) colMap.transferIn = colNum;
        else if (val.includes('transfer out food')) colMap.tOutFood = colNum;
        else if (val.includes('transfer out bev')) colMap.tOutBev = colNum;
        else if (val.includes('transfer out wine')) colMap.tOutWine = colNum;
        else if (val.includes('transfer out other')) colMap.tOutOther = colNum;
        else if (val.includes('transfer out') || val.includes('โอนออก')) colMap.transferOut = colNum;
        else if (val.includes('credit cost food')) colMap.crFood = colNum;
        else if (val.includes('credit cost bev')) colMap.crBev = colNum;
        else if (val.includes('credit cost wine')) colMap.crWine = colNum;
        else if (val.includes('credit cost other')) colMap.crOther = colNum;
        else if (val.includes('cr') || val.includes('oc')) colMap.crCost = colNum;
        else if (val.includes('room avail')) colMap.roomAvailable = colNum;
        else if (val.includes('room sold')) colMap.roomSold = colNum;
        else if (val.includes('com') || val.includes('hu')) colMap.comHuRoom = colNum;
        else if (val.includes('guest') || val.includes('แขก')) colMap.totalGuest = colNum;
        else if (val.includes('weather') || val.includes('อากาศ')) colMap.weather = colNum;
      });

      if (sheetName.includes('purchases') || sheetName.includes('issue')) {
        sheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const d = colMap.date ? getCellDateStr(row.getCell(colMap.date).value) : '';
          const out = colMap.outlet ? String(row.getCell(colMap.outlet).value || '').trim() : '';
          const pf = colMap.purchaseFood ? Number(row.getCell(colMap.purchaseFood).value) || 0 : 0;
          const pb = colMap.purchaseBev ? Number(row.getCell(colMap.purchaseBev).value) || 0 : 0;
          const iF = colMap.issueFood ? Number(row.getCell(colMap.issueFood).value) || 0 : 0;
          const iw = colMap.issueWine ? Number(row.getCell(colMap.issueWine).value) || 0 : 0;
          const ib = colMap.issueBev ? Number(row.getCell(colMap.issueBev).value) || 0 : 0;
          const iO = colMap.issueOther ? Number(row.getCell(colMap.issueOther).value) || 0 : 0;
          // ไฟล์เก่าที่ยังมีคอลัมน์รวมอยู่: Direct Purchase → Food, Store Issue → Bev
          const pfVal = pf + (colMap.purchase ? Number(row.getCell(colMap.purchase).value) || 0 : 0);
          const pbVal = pb;
          const iFVal = iF;
          const iwVal = iw;
          const ibVal = ib + (colMap.issue ? Number(row.getCell(colMap.issue).value) || 0 : 0);
          const iOVal = iO;

          if (d && out && purSheet) {
            for (let idx = purSheet.rowCount; idx >= 2; idx--) {
              const r = purSheet.getRow(idx);
              if (getCellDateStr(r.getCell(1).value) === d && String(r.getCell(2).value || '').trim() === out) {
                purSheet.spliceRows(idx, 1);
              }
            }
            if (pfVal > 0 || pbVal > 0 || iFVal > 0 || iwVal > 0 || ibVal > 0 || iOVal > 0) {
              purSheet.addRow({ date: d, outlet: out, purchaseFood: pfVal, purchaseBev: pbVal, issueFood: iFVal, issueWine: iwVal, issueBev: ibVal, issueOther: iOVal, timestamp: currentTimestamp });
              totalImportedRows++;
            }
          }
        });
      }
      else if (sheetName.includes('adjustment') || sheetName.includes('transfer')) {
        sheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const d = colMap.date ? getCellDateStr(row.getCell(colMap.date).value) : '';
          const out = colMap.outlet ? String(row.getCell(colMap.outlet).value || '').trim() : '';
          const gTIn = colMap.transferIn ? Number(row.getCell(colMap.transferIn).value) || 0 : 0;
          const gTOut = colMap.transferOut ? Number(row.getCell(colMap.transferOut).value) || 0 : 0;
          const gCr = colMap.crCost ? Number(row.getCell(colMap.crCost).value) || 0 : 0;
          // คอลัมน์แยกหมวด (ถ้าไม่มี ใช้ค่ารวมเดิมหาร 4)
          const tInF = colMap.tInFood ? Number(row.getCell(colMap.tInFood).value) || 0 : gTIn / 4;
          const tInB = colMap.tInBev ? Number(row.getCell(colMap.tInBev).value) || 0 : gTIn / 4;
          const tInW = colMap.tInWine ? Number(row.getCell(colMap.tInWine).value) || 0 : gTIn / 4;
          const tInO = colMap.tInOther ? Number(row.getCell(colMap.tInOther).value) || 0 : gTIn / 4;
          const tOutF = colMap.tOutFood ? Number(row.getCell(colMap.tOutFood).value) || 0 : gTOut / 4;
          const tOutB = colMap.tOutBev ? Number(row.getCell(colMap.tOutBev).value) || 0 : gTOut / 4;
          const tOutW = colMap.tOutWine ? Number(row.getCell(colMap.tOutWine).value) || 0 : gTOut / 4;
          const tOutO = colMap.tOutOther ? Number(row.getCell(colMap.tOutOther).value) || 0 : gTOut / 4;
          const crF = colMap.crFood ? Number(row.getCell(colMap.crFood).value) || 0 : gCr / 4;
          const crB = colMap.crBev ? Number(row.getCell(colMap.crBev).value) || 0 : gCr / 4;
          const crW = colMap.crWine ? Number(row.getCell(colMap.crWine).value) || 0 : gCr / 4;
          const crO = colMap.crOther ? Number(row.getCell(colMap.crOther).value) || 0 : gCr / 4;

          if (d && out && costAdjSheet) {
            for (let idx = costAdjSheet.rowCount; idx >= 2; idx--) {
              const r = costAdjSheet.getRow(idx);
              if (getCellDateStr(r.getCell(1).value) === d && String(r.getCell(2).value || '').trim() === out) {
                costAdjSheet.spliceRows(idx, 1);
              }
            }
            if (tInF > 0 || tInB > 0 || tInW > 0 || tInO > 0 || tOutF > 0 || tOutB > 0 || tOutW > 0 || tOutO > 0 ||
                crF > 0 || crB > 0 || crW > 0 || crO > 0) {
              costAdjSheet.addRow({ date: d, outlet: out, tInFood: tInF, tInBev: tInB, tInWine: tInW, tInOther: tInO, tOutFood: tOutF, tOutBev: tOutB, tOutWine: tOutW, tOutOther: tOutO, crFood: crF, crBev: crB, crWine: crW, crOther: crO, timestamp: currentTimestamp });
              totalImportedRows++;
            }
          }
        });
      }
      else if (sheetName.includes('stat') || sheetName.includes('hotel')) {
        sheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const d = colMap.date ? getCellDateStr(row.getCell(colMap.date).value) : '';
          const rAvail = colMap.roomAvailable ? Number(row.getCell(colMap.roomAvailable).value) || 0 : 0;
          const rSold = colMap.roomSold ? Number(row.getCell(colMap.roomSold).value) || 0 : 0;
          const cHu = colMap.comHuRoom ? Number(row.getCell(colMap.comHuRoom).value) || 0 : 0;
          const guests = colMap.totalGuest ? Number(row.getCell(colMap.totalGuest).value) || 0 : 0;
          const weather = colMap.weather ? String(row.getCell(colMap.weather).value || '').trim() : 'Sunny';

          if (d && statsSheet) {
            for (let idx = statsSheet.rowCount; idx >= 2; idx--) {
              const r = statsSheet.getRow(idx);
              if (getCellDateStr(r.getCell(1).value) === d) {
                statsSheet.spliceRows(idx, 1);
              }
            }
            statsSheet.addRow({ date: d, roomAvailable: rAvail, roomSold: rSold, comHuRoom: cHu, totalGuest: guests, weather, timestamp: currentTimestamp });
            totalImportedRows++;
          }
        });
      }
      else if (sheetName === 'data' || sheetName === 'revenue_cost' || sheetName === 'canteen') {
        sheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const d = colMap.date ? getCellDateStr(row.getCell(colMap.date).value) : '';
          const out = colMap.outlet ? String(row.getCell(colMap.outlet).value || '').trim() : (sheetName === 'canteen' ? 'Canteen' : '');
          let cat = colMap.category ? String(row.getCell(colMap.category).value || '').trim() : 'RV-Food';
          let meal = colMap.meal ? String(row.getCell(colMap.meal).value || '').trim() : 'Lunch';
          const rev = colMap.revenue ? Number(row.getCell(colMap.revenue).value) || 0 : 0;
          const cov = colMap.cover ? Number(row.getCell(colMap.cover).value) || 0 : 0;
          const cost = colMap.cost ? Number(row.getCell(colMap.cost).value) || 0 : 0;

          if (!d || !out) return;

          const pairKey = `${d}___${out}`;
          if (!importedDateOutlets.has(pairKey)) {
            importedDateOutlets.add(pairKey);
            if (dataSheet) {
              for (let i = dataSheet.rowCount; i >= 2; i--) {
                const r = dataSheet.getRow(i);
                if (getCellDateStr(r.getCell(1).value) === d && String(r.getCell(2).value || '').trim() === out) {
                  dataSheet.spliceRows(i, 1);
                }
              }
            }
            const outSheet = targetWorkbook.getWorksheet(out);
            if (outSheet) {
              for (let i = outSheet.rowCount; i >= 2; i--) {
                const r = outSheet.getRow(i);
                if (getCellDateStr(r.getCell(1).value) === d) {
                  outSheet.spliceRows(i, 1);
                }
              }
            }
          }

          if (cat.toLowerCase().includes('food')) cat = 'RV-Food';
          else if (cat.toLowerCase().includes('bev')) cat = 'RV-Beverage';
          else if (cat.toLowerCase().includes('wine')) cat = 'RV-Wine';
          else if (cat.toLowerCase().includes('other')) cat = 'RV-Other';

          const validMeals = ['ABF', 'Breakfast', 'Lunch', 'Dinner', 'Supper'];
          const matchedMeal = validMeals.find(m => m.toLowerCase() === meal.toLowerCase());
          if (matchedMeal) meal = matchedMeal;

          const rowObj = { date: d, outlet: out, category: cat, meal: meal, revenue: rev, cover: cov, cost: cost, timestamp: currentTimestamp };
          if (dataSheet) dataSheet.addRow(rowObj);
          
          let outletSheet = targetWorkbook.getWorksheet(out);
          if (!outletSheet) {
            outletSheet = targetWorkbook.addWorksheet(out);
            setupSheetColumns(outletSheet, DATA_COLUMNS);
          }
          outletSheet.addRow(rowObj);
          totalImportedRows++;

          const sumKey = `${d}___${out}___${meal}`;
          if (!mealMapAggregator[sumKey]) {
            mealMapAggregator[sumKey] = { date: d, outlet: out, meal: meal, foodRev: 0, bevRev: 0, wineRev: 0, otherRev: 0, cover: 0, foodCost: 0, bevCost: 0, wineCost: 0, otherCost: 0 };
          }
          if (cat === 'RV-Food') {
            mealMapAggregator[sumKey].foodRev += rev;
            mealMapAggregator[sumKey].foodCost += cost;
          } else if (cat === 'RV-Wine') {
            mealMapAggregator[sumKey].wineRev += rev;
            mealMapAggregator[sumKey].wineCost += cost;
          } else if (cat === 'RV-Other') {
            mealMapAggregator[sumKey].otherRev += rev;
            mealMapAggregator[sumKey].otherCost += cost;
          } else {
            mealMapAggregator[sumKey].bevRev += rev;
            mealMapAggregator[sumKey].bevCost += cost;
          }
          mealMapAggregator[sumKey].cover += cov;
        });
      }
    });

    if (summarySheet) {
      importedDateOutlets.forEach(pair => {
        const [d, out] = pair.split('___');
        for (let idx = summarySheet.rowCount; idx >= 2; idx--) {
          const r = summarySheet.getRow(idx);
          if (getCellDateStr(r.getCell(1).value) === d && String(r.getCell(2).value || '').trim() === out) {
            summarySheet.spliceRows(idx, 1);
          }
        }
      });

      Object.values(mealMapAggregator).forEach(item => {
        summarySheet.addRow({
          date: item.date,
          outlet: item.outlet,
          meal: item.meal,
          foodRev: item.foodRev,
          bevRev: item.bevRev,
          wineRev: item.wineRev,
          otherRev: item.otherRev,
          totalRev: item.foodRev + item.bevRev + item.wineRev + item.otherRev,
          cover: item.cover,
          foodCost: item.foodCost,
          bevCost: item.bevCost,
          wineCost: item.wineCost,
          otherCost: item.otherCost,
          totalCost: item.foodCost + item.bevCost + item.wineCost + item.otherCost,
          purchaseFood: 0,
          purchaseBev: 0,
          issueFood: 0,
          issueWine: 0,
          issueBev: 0,
          issueOther: 0,
          tInFood: 0,
          tInBev: 0,
          tInWine: 0,
          tInOther: 0,
          tOutFood: 0,
          tOutBev: 0,
          tOutWine: 0,
          tOutOther: 0,
          crFood: 0,
          crBev: 0,
          crWine: 0,
          crOther: 0,
          timestamp: currentTimestamp
        });
      });
      sortSheetByDate(summarySheet);
    }

    if (dataSheet) sortSheetByDate(dataSheet);
    if (purSheet) sortSheetByDate(purSheet);
    if (costAdjSheet) sortSheetByDate(costAdjSheet);
    if (statsSheet) sortSheetByDate(statsSheet);

    updateDynamicPivotSheet(targetWorkbook);
    backupExcelFile();

    await saveWorkbook(targetWorkbook);
    res.json({ 
      success: true, 
      message: `✅ Auto Import ทุกชีทสำเร็จ! รวมนำเข้าข้อมูลทั้งสิ้น ${totalImportedRows} รายการ และบันทึกลงระบบเรียบร้อยแล้ว` 
    });

  } catch (err) {
    console.error('Multi-sheet Import Error:', err);
    res.status(500).json({ success: false, message: 'เกิดข้อผิดพลาดในการ Import: ' + err.message });
  }
});

app.post('/api/save-excel', requireAuth('admin'), async (req, res) => {
  try {
    const {
      date, outlet, purchase, issue,
      purchaseFood, purchaseBev, issueFood, issueWine, issueBev, issueOther,
      transferIn, transferOut, crCost,
      tInFood, tInBev, tInWine, tInOther,
      tOutFood, tOutBev, tOutWine, tOutOther,
      crFood, crBev, crWine, crOther,
      roomAvailable, roomSold, comHuRoom, totalGuest, weather, records
    } = req.body;

    const workbook = await initWorkbook();
    const dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');
    const outletSheet = workbook.getWorksheet(outlet);
    const purSheet = workbook.getWorksheet('Purchases_Issues');
    const costAdjSheet = workbook.getWorksheet('Cost_Adjustments');
    const statsSheet = workbook.getWorksheet('Hotel_Stats');
    const summarySheet = workbook.getWorksheet('Outlet_Summary');
    const currentTimestamp = getFormattedTimestamp();

    if (dataSheet) {
      for (let i = dataSheet.rowCount; i >= 2; i--) {
        const row = dataSheet.getRow(i);
        const rDate = getCellDateStr(row.getCell(1).value);
        const rOutlet = String(row.getCell(2).value || '').trim();
        if (rDate === date && rOutlet === outlet) {
          dataSheet.spliceRows(i, 1);
        }
      }
    }

    if (outletSheet) {
      for (let i = outletSheet.rowCount; i >= 2; i--) {
        const row = outletSheet.getRow(i);
        const rDate = getCellDateStr(row.getCell(1).value);
        if (rDate === date) {
          outletSheet.spliceRows(i, 1);
        }
      }
    }

    if (records && records.length > 0) {
      records.forEach(r => {
        const rowObj = {
          date: r.date,
          outlet: r.outlet,
          category: r.category,
          meal: r.meal,
          revenue: Number(r.revenue) || 0,
          cover: Number(r.cover) || 0,
          cost: Number(r.cost) || 0,
          timestamp: currentTimestamp
        };

        if (dataSheet) dataSheet.addRow(rowObj);
        if (outletSheet) outletSheet.addRow(rowObj);
      });

      if (dataSheet) sortSheetByDate(dataSheet);
      if (outletSheet) sortSheetByDate(outletSheet);
    }

    // ฟิลด์แยกหมวด 18 ตัว (รองรับ client เก่า: purchase/issue รวม, transfer/cr รวม → หาร 4 กันยอดเพี้ยน)
    const q = (v) => Number(v) || 0;
    const pFood = q(purchaseFood ?? purchase);
    const pBev = q(purchaseBev);
    const iFood = q(issueFood);
    const iWine = q(issueWine);
    const iBev = q(issueBev ?? issue);
    const iOther = q(issueOther);
    const hasCatTIn = tInFood !== undefined || tInBev !== undefined || tInWine !== undefined || tInOther !== undefined;
    const hasCatTOut = tOutFood !== undefined || tOutBev !== undefined || tOutWine !== undefined || tOutOther !== undefined;
    const hasCatCr = crFood !== undefined || crBev !== undefined || crWine !== undefined || crOther !== undefined;
    const legTIn = q(transferIn), legTOut = q(transferOut), legCr = q(crCost);
    const tInF = hasCatTIn ? q(tInFood) : legTIn / 4;
    const tInB = hasCatTIn ? q(tInBev) : legTIn / 4;
    const tInW = hasCatTIn ? q(tInWine) : legTIn / 4;
    const tInO = hasCatTIn ? q(tInOther) : legTIn / 4;
    const tOutF = hasCatTOut ? q(tOutFood) : legTOut / 4;
    const tOutB = hasCatTOut ? q(tOutBev) : legTOut / 4;
    const tOutW = hasCatTOut ? q(tOutWine) : legTOut / 4;
    const tOutO = hasCatTOut ? q(tOutOther) : legTOut / 4;
    const crF = hasCatCr ? q(crFood) : legCr / 4;
    const crB = hasCatCr ? q(crBev) : legCr / 4;
    const crW = hasCatCr ? q(crWine) : legCr / 4;
    const crO = hasCatCr ? q(crOther) : legCr / 4;

    if (purSheet) {
      for (let idx = purSheet.rowCount; idx >= 2; idx--) {
        const row = purSheet.getRow(idx);
        if (getCellDateStr(row.getCell(1).value) === date && String(row.getCell(2).value || '').trim() === outlet) {
          purSheet.spliceRows(idx, 1);
        }
      }
    }
    if (pFood > 0 || pBev > 0 || iFood > 0 || iWine > 0 || iBev > 0 || iOther > 0) {
      purSheet.addRow({
        date: date,
        outlet: outlet,
        purchaseFood: pFood,
        purchaseBev: pBev,
        issueFood: iFood,
        issueWine: iWine,
        issueBev: iBev,
        issueOther: iOther,
        timestamp: currentTimestamp
      });
      sortSheetByDate(purSheet);
    }

    if (costAdjSheet) {
      for (let idx = costAdjSheet.rowCount; idx >= 2; idx--) {
        const row = costAdjSheet.getRow(idx);
        if (getCellDateStr(row.getCell(1).value) === date && String(row.getCell(2).value || '').trim() === outlet) {
          costAdjSheet.spliceRows(idx, 1);
        }
      }
    }
    if (tInF > 0 || tInB > 0 || tInW > 0 || tInO > 0 || tOutF > 0 || tOutB > 0 || tOutW > 0 || tOutO > 0 ||
        crF > 0 || crB > 0 || crW > 0 || crO > 0) {
      costAdjSheet.addRow({
        date: date,
        outlet: outlet,
        tInFood: tInF,
        tInBev: tInB,
        tInWine: tInW,
        tInOther: tInO,
        tOutFood: tOutF,
        tOutBev: tOutB,
        tOutWine: tOutW,
        tOutOther: tOutO,
        crFood: crF,
        crBev: crB,
        crWine: crW,
        crOther: crO,
        timestamp: currentTimestamp
      });
      sortSheetByDate(costAdjSheet);
    }

    if (summarySheet) {
      for (let idx = summarySheet.rowCount; idx >= 2; idx--) {
        const row = summarySheet.getRow(idx);
        if (getCellDateStr(row.getCell(1).value) === date && String(row.getCell(2).value || '').trim() === outlet) {
          summarySheet.spliceRows(idx, 1);
        }
      }

      const mealMap = {};
      if (records && records.length > 0) {
        records.forEach(r => {
          const m = r.meal || 'General';
          if (!mealMap[m]) {
            mealMap[m] = { foodRev: 0, bevRev: 0, wineRev: 0, otherRev: 0, cover: 0, foodCost: 0, bevCost: 0, wineCost: 0, otherCost: 0 };
          }
          const rev = Number(r.revenue) || 0;
          const cost = Number(r.cost) || 0;
          const cov = Number(r.cover) || 0;

          if (r.category === 'RV-Food') {
            mealMap[m].foodRev += rev;
            mealMap[m].foodCost += cost;
          } else if (r.category === 'RV-Wine') {
            mealMap[m].wineRev += rev;
            mealMap[m].wineCost += cost;
          } else if (r.category === 'RV-Other') {
            mealMap[m].otherRev += rev;
            mealMap[m].otherCost += cost;
          } else {
            mealMap[m].bevRev += rev;
            mealMap[m].bevCost += cost;
          }
          mealMap[m].cover += cov;
        });
      }

      const mealsList = Object.keys(mealMap);
      if (mealsList.length > 0) {
        // ยอดปรับปรุงบันทึกแค่แถวเดียวต่อวัน/outlet ที่มื้อ Dinner (ถ้าไม่มี Dinner ไล่ Supper → Lunch → Breakfast → ABF)
        const carrierMeal = ['Dinner', 'Supper', 'Lunch', 'Breakfast', 'ABF'].find(m => mealsList.includes(m)) || mealsList[0];
        mealsList.forEach(m => {
          const item = mealMap[m];
          const isCarrier = m === carrierMeal;
          summarySheet.addRow({
            date: date,
            outlet: outlet,
            meal: m,
            foodRev: item.foodRev,
            bevRev: item.bevRev,
            wineRev: item.wineRev,
            otherRev: item.otherRev,
            totalRev: item.foodRev + item.bevRev + item.wineRev + item.otherRev,
            cover: item.cover,
            foodCost: item.foodCost,
            bevCost: item.bevCost,
            wineCost: item.wineCost,
            otherCost: item.otherCost,
            totalCost: item.foodCost + item.bevCost + item.wineCost + item.otherCost,
            purchaseFood: isCarrier ? pFood : 0,
            purchaseBev: isCarrier ? pBev : 0,
            issueFood: isCarrier ? iFood : 0,
            issueWine: isCarrier ? iWine : 0,
            issueBev: isCarrier ? iBev : 0,
            issueOther: isCarrier ? iOther : 0,
            tInFood: isCarrier ? tInF : 0,
            tInBev: isCarrier ? tInB : 0,
            tInWine: isCarrier ? tInW : 0,
            tInOther: isCarrier ? tInO : 0,
            tOutFood: isCarrier ? tOutF : 0,
            tOutBev: isCarrier ? tOutB : 0,
            tOutWine: isCarrier ? tOutW : 0,
            tOutOther: isCarrier ? tOutO : 0,
            crFood: isCarrier ? crF : 0,
            crBev: isCarrier ? crB : 0,
            crWine: isCarrier ? crW : 0,
            crOther: isCarrier ? crO : 0,
            timestamp: currentTimestamp
          });
        });
      } else if (pFood > 0 || pBev > 0 || iFood > 0 || iWine > 0 || iBev > 0 || iOther > 0 ||
                 tInF > 0 || tInB > 0 || tInW > 0 || tInO > 0 ||
                 tOutF > 0 || tOutB > 0 || tOutW > 0 || tOutO > 0 ||
                 crF > 0 || crB > 0 || crW > 0 || crO > 0) {
        summarySheet.addRow({
          date: date,
          outlet: outlet,
          meal: '-',
          foodRev: 0,
          bevRev: 0,
          wineRev: 0,
          otherRev: 0,
          totalRev: 0,
          cover: 0,
          foodCost: 0,
          bevCost: 0,
          wineCost: 0,
          otherCost: 0,
          totalCost: 0,
          purchaseFood: pFood,
          purchaseBev: pBev,
          issueFood: iFood,
          issueWine: iWine,
          issueBev: iBev,
          issueOther: iOther,
          tInFood: tInF,
          tInBev: tInB,
          tInWine: tInW,
          tInOther: tInO,
          tOutFood: tOutF,
          tOutBev: tOutB,
          tOutWine: tOutW,
          tOutOther: tOutO,
          crFood: crF,
          crBev: crB,
          crWine: crW,
          crOther: crO,
          timestamp: currentTimestamp
        });
      }
      sortSheetByDate(summarySheet);
    }

    updateDynamicPivotSheet(workbook);

    let existingRow = null;
    if (statsSheet) {
      statsSheet.eachRow((row, rowNum) => {
        if (rowNum > 1 && getCellDateStr(row.getCell(1).value) === date) {
          existingRow = row;
        }
      });
    }

    // บันทึกสถิติห้องพักและสภาพอากาศเสมอเมื่อมีการระบุสภาพอากาศหรือตัวเลขสถิติ
    if (statsSheet && (weather || Number(roomAvailable) >= 0)) {
      if (existingRow) {
        if (Number(roomAvailable) > 0) existingRow.getCell(2).value = Number(roomAvailable);
        if (Number(roomSold) > 0) existingRow.getCell(3).value = Number(roomSold);
        if (Number(comHuRoom) > 0) existingRow.getCell(4).value = Number(comHuRoom);
        if (Number(totalGuest) > 0) existingRow.getCell(5).value = Number(totalGuest);
        if (weather) existingRow.getCell(6).value = weather;
        existingRow.getCell(7).value = currentTimestamp;
      } else {
        statsSheet.addRow({
          date: date,
          roomAvailable: Number(roomAvailable) || 0,
          roomSold: Number(roomSold) || 0,
          comHuRoom: Number(comHuRoom) || 0,
          totalGuest: Number(totalGuest) || 0,
          weather: weather || '',
          timestamp: currentTimestamp
        });
      }
      sortSheetByDate(statsSheet);
    }

    backupExcelFile();

    await saveWorkbook(workbook);
    res.json({ success: true, message: 'บันทึก สำรองข้อมูล และซิงค์ทุกชีทสำเร็จเรียบร้อยแล้ว' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: 'เกิดข้อผิดพลาด: ' + err.message });
  }
});

app.post('/api/check-duplicate', requireAuth(), async (req, res) => {
  try {
    const { date, outlet } = req.body;
    if (!fs.existsSync(EXCEL_FILE)) return res.json({ isDuplicate: false });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    const dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');

    let exists = false;
    if (dataSheet) {
      dataSheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
          const rowDate = getCellDateStr(row.getCell(1).value);
          const rowOutlet = String(row.getCell(2).value || '').trim();
          if (rowDate === date && rowOutlet === outlet) exists = true;
        }
      });
    }

    res.json({
      isDuplicate: exists,
      message: exists ? `ตรวจพบว่าเคยบันทึกข้อมูลวันที่ ${date} แผนก ${outlet} แล้ว` : ''
    });
  } catch (err) {
    res.json({ isDuplicate: false });
  }
});

app.get('/api/dashboard-summary', requireAuth(), async (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.json({
        success: true,
        data: {
          kpi: { 
            totalRevenue: 0, totalCost: 0, totalNetCost: 0, totalPurchase: 0, totalIssue: 0, 
            totalTransferIn: 0, totalTransferOut: 0, totalCrCost: 0, costPercentage: 0, 
            grossProfit: 0, totalCover: 0, totalRoomAvail: 0, totalRoomSold: 0, totalComHu: 0, 
            occupancyRate: 0, roomSoldToday: 0, roomSoldMtd: 0, guestToday: 0, guestMtd: 0, weather: '-',
            foodRevenue: 0, foodCostPct: 0, foodCostRaw: 0, bevRevenue: 0, bevCostPct: 0, bevCostRaw: 0
          },
          byOutlet: { labels: [], revenues: [], costs: [], netCosts: [], covers: [], purchases: [], issues: [], transferIns: [], transferOuts: [], crCosts: [] },
          byOutletFood: { labels: [], revenues: [], costs: [], rawCosts: [], covers: [] },
          byOutletBev: { labels: [], revenues: [], costs: [], rawCosts: [], covers: [] },
          byOutletWine: { labels: [], revenues: [], costs: [], rawCosts: [], covers: [] },
          byOutletOther: { labels: [], revenues: [], costs: [], rawCosts: [], covers: [] },
          canteen: { revenue: 0, netCost: 0, rawCost: 0, cover: 0, purchase: 0, issue: 0, transferIn: 0, transferOut: 0, crCost: 0, costPerDay: 0, costPerMeal: 0, costPct: '0.00%' },
          byCategory: { labels: ['RV-Food', 'RV-Beverage', 'RV-Wine', 'RV-Other'], revenues: [0, 0, 0, 0] },
          dailyTrend: { dates: [], revenues: [], costs: [], netCosts: [] },
          occupancyTrend: { dates: [], occupancy: [], average: 0 },
          dateRange: { min: '', max: '' }
        }
      });
    }

    const { startDate, endDate, outlet } = req.query;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);

    const dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');
    const purSheet = workbook.getWorksheet('Purchases_Issues');
    const costAdjSheet = workbook.getWorksheet('Cost_Adjustments');
    const statsSheet = workbook.getWorksheet('Hotel_Stats');

    let totalRevenue = 0, totalCost = 0, totalCover = 0;
    let foodRevenue = 0, foodCostRaw = 0, bevRevenue = 0, bevCostRaw = 0;
    let totalPurchase = 0, totalIssue = 0, totalTransferIn = 0, totalTransferOut = 0, totalCrCost = 0;
    let totalRoomAvail = 0, totalRoomSold = 0, totalComHu = 0;
    
    let roomSoldToday = 0, roomSoldMtd = 0;
    let guestToday = 0, guestMtd = 0;
    let todayWeather = '-';
    
    let canteenRev = 0, canteenCost = 0, canteenCover = 0, canteenPur = 0, canteenIss = 0, canteenTIn = 0, canteenTOut = 0, canteenCr = 0;

    const outletMap = {}, outletFoodMap = {}, outletBevMap = {}, outletWineMap = {}, outletOtherMap = {}, catMap = { 'RV-Food': 0, 'RV-Beverage': 0, 'RV-Wine': 0, 'RV-Other': 0 }, dailyMap = {};
    const outletCatPool = {};   // Net Cost pool รายหมวดของแต่ละ outlet: { food, bev, wine, other }
    const dailyAdjPool = {};    // Net Cost pool รวมรายวัน (Purchase + Issue + Transfer In − Transfer Out − Cr Cost)
    let dates = [];

    const addToCatPool = (out, cat, amount) => {
      if (!outletCatPool[out]) outletCatPool[out] = { food: 0, bev: 0, wine: 0, other: 0 };
      outletCatPool[out][cat] += amount;
    };

    const ALL_OUTLETS = getOutlets();
    const ALL_ACTIVE_OUTLETS = outlet && outlet !== 'ALL' ? [outlet] : ALL_OUTLETS;

    ALL_ACTIVE_OUTLETS.forEach(outName => {
      outletMap[outName] = { rev: 0, cost: 0, cov: 0, pur: 0, iss: 0, tIn: 0, tOut: 0, cr: 0 };
      outletFoodMap[outName] = { rev: 0, cost: 0, cov: 0 };
      outletBevMap[outName] = { rev: 0, cost: 0, cov: 0 };
      outletWineMap[outName] = { rev: 0, cost: 0, cov: 0 };
      outletOtherMap[outName] = { rev: 0, cost: 0, cov: 0 };
    });

    if (dataSheet) {
      let colDate = 1, colOutlet = 2, colCat = 3, colRev = 5, colCover = 6, colCost = 7;

      const headerRow = dataSheet.getRow(1);
      headerRow.eachCell((cell, colNumber) => {
        const text = String(cell.value || '').trim().toLowerCase();
        if (text.startsWith('date')) colDate = colNumber;
        else if (text.startsWith('outlet')) colOutlet = colNumber;
        else if (text.includes('categor')) colCat = colNumber;
        else if (text.startsWith('revenue')) colRev = colNumber;
        else if (text.startsWith('cover')) colCover = colNumber;
        else if (text.startsWith('cost')) colCost = colNumber;
      });

      dataSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;

        const d = getCellDateStr(row.getCell(colDate).value);
        const out = String(row.getCell(colOutlet).value || '').trim();
        const cat = String(row.getCell(colCat).value || '').trim();
        const rev = Number(row.getCell(colRev).value) || 0;
        const cov = Number(row.getCell(colCover).value) || 0;
        const cost = Number(row.getCell(colCost).value) || 0;

        if (d) dates.push(d);

        if (startDate && d < startDate) return;
        if (endDate && d > endDate) return;
        if (outlet && outlet !== 'ALL' && out !== outlet) return;

        if (out === 'Canteen') {
          canteenRev += rev;
          canteenCost += cost;
          canteenCover += cov;
          return;
        }

        totalRevenue += rev;
        totalCost += cost;
        totalCover += cov;

        if (!outletMap[out]) {
          outletMap[out] = { rev: 0, cost: 0, cov: 0, pur: 0, iss: 0, tIn: 0, tOut: 0, cr: 0 };
        }
        if (!outletFoodMap[out]) outletFoodMap[out] = { rev: 0, cost: 0, cov: 0 };
        if (!outletBevMap[out]) outletBevMap[out] = { rev: 0, cost: 0, cov: 0 };
        if (!outletWineMap[out]) outletWineMap[out] = { rev: 0, cost: 0, cov: 0 };
        if (!outletOtherMap[out]) outletOtherMap[out] = { rev: 0, cost: 0, cov: 0 };

        outletMap[out].rev += rev;
        outletMap[out].cost += cost;
        outletMap[out].cov += cov;

        if (cat === 'RV-Food') {
          foodRevenue += rev;
          foodCostRaw += cost;
          outletFoodMap[out].rev += rev;
          outletFoodMap[out].cost += cost;
          outletFoodMap[out].cov += cov;
        } else if (cat === 'RV-Beverage') {
          bevRevenue += rev;
          bevCostRaw += cost;
          outletBevMap[out].rev += rev;
          outletBevMap[out].cost += cost;
          outletBevMap[out].cov += cov;
        } else if (cat === 'RV-Wine') {
          outletWineMap[out].rev += rev;
          outletWineMap[out].cost += cost;
          outletWineMap[out].cov += cov;
        } else if (cat === 'RV-Other') {
          outletOtherMap[out].rev += rev;
          outletOtherMap[out].cost += cost;
          outletOtherMap[out].cov += cov;
        }

        if (catMap[cat] !== undefined) catMap[cat] += rev;

        if (!dailyMap[d]) dailyMap[d] = { rev: 0, cost: 0 };
        dailyMap[d].rev += rev;
        dailyMap[d].cost += cost;
      });
    }

    const purMap = {};
    if (purSheet) {
      const layout = purSheetLayout(purSheet);
      purSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const d = getCellDateStr(row.getCell(1).value);
        const out = String(row.getCell(2).value || '').trim();
        // แยกตามหมวด (ใช้ mapping เดียวกับ /api/shift-summary)
        let pf = 0, pb = 0, iF = 0, iw = 0, ib = 0, iO = 0;
        if (layout === 'v3') {
          pf = Number(row.getCell(3).value) || 0;
          pb = Number(row.getCell(4).value) || 0;
          iF = Number(row.getCell(5).value) || 0;
          iw = Number(row.getCell(6).value) || 0;
          ib = Number(row.getCell(7).value) || 0;
          iO = Number(row.getCell(8).value) || 0;
        } else if (layout === 'v2') {
          pf = Number(row.getCell(3).value) || 0;
          pb = Number(row.getCell(4).value) || 0;
          iw = Number(row.getCell(5).value) || 0;
          ib = Number(row.getCell(6).value) || 0;
        } else {
          pf = Number(row.getCell(3).value) || 0;
          ib = Number(row.getCell(4).value) || 0;
        }
        const pur = pf + pb;
        const iss = iF + iw + ib + iO;

        if (startDate && d < startDate) return;
        if (endDate && d > endDate) return;
        if (outlet && outlet !== 'ALL' && out !== outlet) return;

        const key = `${d}_${out}`;
        purMap[key] = { out, d, pur, iss, pf, pb, iF, iw, ib, iO };
      });
    }

    Object.values(purMap).forEach(item => {
      if (item.out === 'Canteen') {
        canteenPur += item.pur;
        canteenIss += item.iss;
        return;
      }
      totalPurchase += item.pur;
      totalIssue += item.iss;
      if (!outletMap[item.out]) {
        outletMap[item.out] = { rev: 0, cost: 0, cov: 0, pur: 0, iss: 0, tIn: 0, tOut: 0, cr: 0 };
      }
      outletMap[item.out].pur += item.pur;
      outletMap[item.out].iss += item.iss;

      addToCatPool(item.out, 'food', item.pf + item.iF);
      addToCatPool(item.out, 'bev', item.pb + item.ib);
      addToCatPool(item.out, 'wine', item.iw);
      addToCatPool(item.out, 'other', item.iO);
      dailyAdjPool[item.d] = (dailyAdjPool[item.d] || 0) + item.pur + item.iss;
    });

    const adjMap = {};
    if (costAdjSheet) {
      const layout = adjSheetLayout(costAdjSheet);
      costAdjSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const d = getCellDateStr(row.getCell(1).value);
        const out = String(row.getCell(2).value || '').trim();
        // แยกตามหมวด (ใช้ mapping เดียวกับ /api/shift-summary)
        let tInF = 0, tInB = 0, tInW = 0, tInO = 0, tOutF = 0, tOutB = 0, tOutW = 0, tOutO = 0, crF = 0, crB = 0, crW = 0, crO = 0;
        if (layout === 'v2') {
          tInF = Number(row.getCell(3).value) || 0;
          tInB = Number(row.getCell(4).value) || 0;
          tInW = Number(row.getCell(5).value) || 0;
          tInO = Number(row.getCell(6).value) || 0;
          tOutF = Number(row.getCell(7).value) || 0;
          tOutB = Number(row.getCell(8).value) || 0;
          tOutW = Number(row.getCell(9).value) || 0;
          tOutO = Number(row.getCell(10).value) || 0;
          crF = Number(row.getCell(11).value) || 0;
          crB = Number(row.getCell(12).value) || 0;
          crW = Number(row.getCell(13).value) || 0;
          crO = Number(row.getCell(14).value) || 0;
        } else {
          // คอลัมน์รวมเดิม → แบ่งเท่าๆ 4 หมวด
          const tIn = (Number(row.getCell(3).value) || 0) / 4;
          const tOut = (Number(row.getCell(4).value) || 0) / 4;
          const cr = (Number(row.getCell(5).value) || 0) / 4;
          tInF = tIn; tInB = tIn; tInW = tIn; tInO = tIn;
          tOutF = tOut; tOutB = tOut; tOutW = tOut; tOutO = tOut;
          crF = cr; crB = cr; crW = cr; crO = cr;
        }
        const tIn = tInF + tInB + tInW + tInO;
        const tOut = tOutF + tOutB + tOutW + tOutO;
        const cr = crF + crB + crW + crO;

        if (startDate && d < startDate) return;
        if (endDate && d > endDate) return;
        if (outlet && outlet !== 'ALL' && out !== outlet) return;

        const key = `${d}_${out}`;
        adjMap[key] = { out, d, tIn, tOut, cr, tInF, tInB, tInW, tInO, tOutF, tOutB, tOutW, tOutO, crF, crB, crW, crO };
      });
    }

    Object.values(adjMap).forEach(item => {
      if (item.out === 'Canteen') {
        canteenTIn += item.tIn;
        canteenTOut += item.tOut;
        canteenCr += item.cr;
        return;
      }
      totalTransferIn += item.tIn;
      totalTransferOut += item.tOut;
      totalCrCost += item.cr;
      if (!outletMap[item.out]) {
        outletMap[item.out] = { rev: 0, cost: 0, cov: 0, pur: 0, iss: 0, tIn: 0, tOut: 0, cr: 0 };
      }
      outletMap[item.out].tIn += item.tIn;
      outletMap[item.out].tOut += item.tOut;
      outletMap[item.out].cr += item.cr;

      addToCatPool(item.out, 'food', item.tInF - item.tOutF - item.crF);
      addToCatPool(item.out, 'bev', item.tInB - item.tOutB - item.crB);
      addToCatPool(item.out, 'wine', item.tInW - item.tOutW - item.crW);
      addToCatPool(item.out, 'other', item.tInO - item.tOutO - item.crO);
      dailyAdjPool[item.d] = (dailyAdjPool[item.d] || 0) + item.tIn - item.tOut - item.cr;
    });

    // ===== ประมวลผล Hotel Stats (Today, MTD, Weather, Occupancy) =====
    const hotelStatsRecords = [];
    if (statsSheet) {
      statsSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const d = getCellDateStr(row.getCell(1).value);
        const rAvail = Number(row.getCell(2).value) || 0;
        const rSold = Number(row.getCell(3).value) || 0;
        const cHu = Number(row.getCell(4).value) || 0;
        const guests = Number(row.getCell(5).value) || 0;
        const weatherVal = String(row.getCell(6).value || '').trim();

        if (d) {
          hotelStatsRecords.push({ 
            date: d, 
            roomAvailable: rAvail, 
            roomSold: rSold, 
            comHuRoom: cHu, 
            totalGuest: guests, 
            weather: weatherVal 
          });
        }
      });
    }

    hotelStatsRecords.sort((a, b) => a.date.localeCompare(b.date));

    // กำหนดวันที่เป้าหมาย (Anchor Date)
    const targetAnchorDate = endDate || (dates.length ? dates[dates.length - 1] : new Date().toISOString().split('T')[0]);

    // 1. ค้นหาแถวสถิติของวันเป้าหมาย (ถ้าไม่มี ให้หาวันล่าสุดที่มีข้อมูลจริง)
    let exactTargetStat = hotelStatsRecords.find(r => r.date === targetAnchorDate);
    if (!exactTargetStat && hotelStatsRecords.length > 0) {
      const recordedBeforeAnchor = hotelStatsRecords.filter(r => r.date <= targetAnchorDate);
      exactTargetStat = recordedBeforeAnchor.length > 0 ? recordedBeforeAnchor[recordedBeforeAnchor.length - 1] : hotelStatsRecords[hotelStatsRecords.length - 1];
    }

    if (exactTargetStat) {
      roomSoldToday = exactTargetStat.roomSold;
      guestToday = exactTargetStat.totalGuest;
      todayWeather = exactTargetStat.weather || '-';
    }

    // 2. คำนวณยอด MTD สะสมตั้งแต่วันที่ 1 ของเดือน จนถึง Target Date
    const targetMonthPrefix = targetAnchorDate.substring(0, 7);
    hotelStatsRecords.forEach(r => {
      if (r.date.startsWith(targetMonthPrefix) && r.date <= targetAnchorDate) {
        roomSoldMtd += r.roomSold;
        guestMtd += r.totalGuest;
      }
    });

    // 3. สะสมยอด Occupancy ตามช่วงวันที่ที่เลือกใน Filter
    hotelStatsRecords.forEach(r => {
      if (startDate && r.date < startDate) return;
      if (endDate && r.date > endDate) return;

      totalRoomAvail += r.roomAvailable;
      totalRoomSold += r.roomSold;
      totalComHu += r.comHuRoom;
    });

    // ===== Room Occupancy รายวัน (MTD: วันที่ 1 ของเดือนถึงวันเป้าหมาย) =====
    const occMonthStart = targetAnchorDate.substring(0, 7) + '-01';
    const occDates = [], occValues = [];
    hotelStatsRecords.forEach(r => {
      if (r.date >= occMonthStart && r.date <= targetAnchorDate && r.roomAvailable > 0) {
        occDates.push(r.date);
        occValues.push(Number(((r.roomSold / r.roomAvailable) * 100).toFixed(2)));
      }
    });
    const occAverage = occValues.length ? Number((occValues.reduce((a, b) => a + b, 0) / occValues.length).toFixed(2)) : 0;

    const occupancyRate = totalRoomAvail > 0 ? (totalRoomSold / totalRoomAvail) * 100 : 0;
    // Net Cost = Purchase + Issue + Transfer In − Transfer Out − Cr Cost
    // (นิยามเดียวกับ Allocated Net Cost ในหน้า Shift Analytics — Raw Cost ใช้เป็นสัดส่วนปันเท่านั้น ไม่นำมาบวกซ้ำ)
    const totalNetCost = totalPurchase + totalIssue + totalTransferIn - totalTransferOut - totalCrCost;

    const poolSum = (cat) => Object.values(outletCatPool).reduce((sum, p) => sum + (p[cat] || 0), 0);
    const foodNetCost = poolSum('food');
    const bevNetCost = poolSum('bev');

    const foodCostPct = foodRevenue > 0 ? (foodNetCost / foodRevenue) * 100 : 0;
    const bevCostPct = bevRevenue > 0 ? (bevNetCost / bevRevenue) * 100 : 0;

    const canteenNetCost = canteenCost + canteenPur + canteenIss + canteenTIn - canteenTOut - canteenCr;
    const totalCoverCount = canteenCover > 0 ? canteenCover : 1;
    const canteenCostPerDay = canteenNetCost / totalCoverCount;
    const canteenCostPerMeal = canteenCostPerDay / 2;

    const sortedDaily = Object.keys(dailyMap).sort();
    const outletKeys = Object.keys(outletMap).filter(k => (outletMap[k].rev > 0 || outletMap[k].cost > 0 || outletMap[k].pur > 0 || outletMap[k].iss > 0) && k !== 'Canteen');

    const EMPTY_CAT = { rev: 0, cost: 0, cov: 0 };
    const EMPTY_POOL = { food: 0, bev: 0, wine: 0, other: 0 };

    // Cost รายหมวดของแต่ละ outlet = Net Cost pool ของหมวดนั้น (Purchase + Issue + Transfer In − Transfer Out − Cr Cost)
    const foodNetCostsArr = outletKeys.map(k => (outletCatPool[k] || EMPTY_POOL).food);
    const bevNetCostsArr = outletKeys.map(k => (outletCatPool[k] || EMPTY_POOL).bev);
    const wineNetCostsArr = outletKeys.map(k => (outletCatPool[k] || EMPTY_POOL).wine);
    const otherNetCostsArr = outletKeys.map(k => (outletCatPool[k] || EMPTY_POOL).other);

    dates.sort();

    res.json({
      success: true,
      data: {
        kpi: {
          totalRevenue,
          totalCost,
          totalNetCost,
          totalPurchase,
          totalIssue,
          totalTransferIn,
          totalTransferOut,
          totalCrCost,
          costPercentage: totalRevenue > 0 ? (totalNetCost / totalRevenue) * 100 : 0,
          grossProfit: totalRevenue - totalNetCost,
          totalCover,
          totalRoomAvail,
          totalRoomSold,
          totalComHu,
          occupancyRate,
          roomSoldToday,
          roomSoldMtd,
          guestToday,
          guestMtd,
          weather: todayWeather,
          foodRevenue,
          foodCostPct,
          foodCostRaw,
          bevRevenue,
          bevCostPct,
          bevCostRaw
        },
        byOutlet: {
          labels: outletKeys,
          revenues: outletKeys.map(k => outletMap[k].rev),
          costs: outletKeys.map(k => outletMap[k].cost),
          // Net Cost (pool) = Purchase + Issue + Transfer In − Transfer Out − Cr Cost
          netCosts: outletKeys.map(k => {
            const o = outletMap[k];
            return o.pur + o.iss + o.tIn - o.tOut - o.cr;
          }),
          covers: outletKeys.map(k => outletMap[k].cov),
          purchases: outletKeys.map(k => outletMap[k].pur),
          issues: outletKeys.map(k => outletMap[k].iss),
          transferIns: outletKeys.map(k => outletMap[k].tIn),
          transferOuts: outletKeys.map(k => outletMap[k].tOut),
          crCosts: outletKeys.map(k => outletMap[k].cr)
        },
        byOutletFood: {
          labels: outletKeys,
          revenues: outletKeys.map(k => (outletFoodMap[k] || EMPTY_CAT).rev),
          costs: foodNetCostsArr,
          rawCosts: outletKeys.map(k => (outletFoodMap[k] || EMPTY_CAT).cost),
          covers: outletKeys.map(k => (outletFoodMap[k] || EMPTY_CAT).cov)
        },
        byOutletBev: {
          labels: outletKeys,
          revenues: outletKeys.map(k => (outletBevMap[k] || EMPTY_CAT).rev),
          costs: bevNetCostsArr,
          rawCosts: outletKeys.map(k => (outletBevMap[k] || EMPTY_CAT).cost),
          covers: outletKeys.map(k => (outletBevMap[k] || EMPTY_CAT).cov)
        },
        byOutletWine: {
          labels: outletKeys,
          revenues: outletKeys.map(k => (outletWineMap[k] || EMPTY_CAT).rev),
          costs: wineNetCostsArr,
          rawCosts: outletKeys.map(k => (outletWineMap[k] || EMPTY_CAT).cost),
          covers: outletKeys.map(k => (outletWineMap[k] || EMPTY_CAT).cov)
        },
        byOutletOther: {
          labels: outletKeys,
          revenues: outletKeys.map(k => (outletOtherMap[k] || EMPTY_CAT).rev),
          costs: otherNetCostsArr,
          rawCosts: outletKeys.map(k => (outletOtherMap[k] || EMPTY_CAT).cost),
          covers: outletKeys.map(k => (outletOtherMap[k] || EMPTY_CAT).cov)
        },
        canteen: {
          revenue: canteenRev,
          netCost: canteenNetCost,
          rawCost: canteenCost,
          cover: canteenCover,
          purchase: canteenPur,
          issue: canteenIss,
          transferIn: canteenTIn,
          transferOut: canteenTOut,
          crCost: canteenCr,
          costPerDay: canteenCostPerDay,
          costPerMeal: canteenCostPerMeal,
          costPct: canteenRev > 0 ? ((canteenNetCost / canteenRev) * 100).toFixed(2) + '%' : '0.00%'
        },
        byCategory: {
          labels: Object.keys(catMap),
          revenues: Object.values(catMap)
        },
        dailyTrend: {
          dates: sortedDaily,
          revenues: sortedDaily.map(d => dailyMap[d].rev),
          costs: sortedDaily.map(d => dailyMap[d].cost),
          // Net Cost รายวัน = pool จริงของวันนั้น (Purchase + Issue + Transfer In − Transfer Out − Cr Cost)
          netCosts: sortedDaily.map(d => dailyAdjPool[d] || 0)
        },
        occupancyTrend: {
          dates: occDates,
          occupancy: occValues,
          average: occAverage
        },
        dateRange: {
          min: dates.length ? dates[0] : '',
          max: dates.length ? dates[dates.length - 1] : ''
        }
      }
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// ===== Budget (เป้าหมายรายเดือนต่อ outlet กระจายเป็น daily target) =====
function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function emptyBudgetMetrics() {
  return { food: 0, bev: 0, wine: 0, other: 0, cost: 0, covers: 0 };
}

function readBudgetStore() {
  try {
    if (!fs.existsSync(BUDGET_FILE)) return {};
    const parsed = JSON.parse(fs.readFileSync(BUDGET_FILE, 'utf8'));
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (err) {
    return {};
  }
}

function writeBudgetStore(store) {
  fs.writeFileSync(BUDGET_FILE, JSON.stringify(store, null, 2));
}

// แบ่งยอดเป้ารายเดือนเป็น daily target แบบ Even — ทศนิยม 2 ตำแหน่ง วันสุดท้ายดูดเศษให้พอดี
// covers เป็นจำนวนคน ต้องได้เลขจำนวนเต็ม
function evenSplitMonthly(monthlyTotal, days, isInteger) {
  const result = new Array(days).fill(0);
  if (!Number.isFinite(monthlyTotal) || monthlyTotal <= 0 || days <= 0) return result;
  if (isInteger) {
    const base = Math.floor(monthlyTotal / days);
    let remainder = monthlyTotal - base * days;
    for (let i = 0; i < days; i++) result[i] = base + (remainder-- > 0 ? 1 : 0);
    return result;
  }
  const perDay = Math.round((monthlyTotal / days) * 100) / 100;
  for (let i = 0; i < days; i++) result[i] = perDay;
  result[days - 1] = Math.round((monthlyTotal - perDay * (days - 1)) * 100) / 100;
  return result;
}

// API อ่าน budget ของเดือน — ?year&month → { "<Outlet>": { monthly, daily } }
app.get('/api/budget', requireAuth(), (req, res) => {
  try {
    const year = parseInt(req.query.year, 10);
    const month = parseInt(req.query.month, 10);
    if (!year || !month || month < 1 || month > 12) {
      return res.status(400).json({ success: false, message: 'ระบุ year และ month ให้ถูกต้อง' });
    }
    const monthKey = `${year}-${String(month).padStart(2, '0')}`;
    const store = readBudgetStore();
    res.json({ success: true, month: monthKey, data: store[monthKey] || {} });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// API บันทึก budget เดือน+outlet (admin) — ถ้าไม่ส่ง daily มา ให้กระจาย Even อัตโนมัติ
app.post('/api/budget', requireAuth('admin'), (req, res) => {
  try {
    const { year, month, outlet, monthly, daily } = req.body || {};
    const y = parseInt(year, 10);
    const m = parseInt(month, 10);
    if (!y || !m || m < 1 || m > 12) {
      return res.status(400).json({ success: false, message: 'ระบุ year และ month ให้ถูกต้อง' });
    }
    const outletName = String(outlet || '').trim();
    if (!outletName || !getOutlets().some(o => o.toLowerCase() === outletName.toLowerCase())) {
      return res.status(400).json({ success: false, message: 'Outlet ไม่ถูกต้อง' });
    }
    const monthKey = `${y}-${String(m).padStart(2, '0')}`;
    const days = getDaysInMonth(y, m);

    const monthlyNums = emptyBudgetMetrics();
    BUDGET_METRICS.forEach(k => { monthlyNums[k] = Number(monthly && monthly[k]) || 0; });

    let dailyRows;
    if (Array.isArray(daily) && daily.length) {
      dailyRows = [];
      for (let d = 0; d < days; d++) {
        const row = daily[d] || {};
        const entry = {};
        BUDGET_METRICS.forEach(k => { entry[k] = Number(row[k]) || 0; });
        dailyRows.push(entry);
      }
    } else {
      const metricSplits = {};
      BUDGET_METRICS.forEach(k => { metricSplits[k] = evenSplitMonthly(monthlyNums[k], days, k.startsWith('covers')); });
      dailyRows = [];
      for (let d = 0; d < days; d++) {
        const entry = {};
        BUDGET_METRICS.forEach(k => { entry[k] = metricSplits[k][d]; });
        dailyRows.push(entry);
      }
    }

    const store = readBudgetStore();
    if (!store[monthKey]) store[monthKey] = {};
    store[monthKey][outletName] = { monthly: monthlyNums, daily: dailyRows };
    writeBudgetStore(store);
    res.json({ success: true, message: `บันทึก Budget ของ ${outletName} เดือน ${monthKey} เรียบร้อยแล้ว` });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// API รายงานประจำเดือน — ตัวเลขจริงรายวันรายหมวด + สรุปทั้งเดือน
// ?year&month&outlet (outlet = ALL หรือชื่อ outlet; โหมด ALL ตัด Canteen ออกเหมือนหน้า Dashboard)
app.get('/api/monthly-report', requireAuth(), async (req, res) => {
  try {
    const year = parseInt(req.query.year, 10);
    const month = parseInt(req.query.month, 10);
    if (!year || !month || month < 1 || month > 12) {
      return res.status(400).json({ success: false, message: 'ระบุ year และ month ให้ถูกต้อง' });
    }
    const outletFilter = String(req.query.outlet || 'ALL').trim() || 'ALL';
    const monthKey = `${year}-${String(month).padStart(2, '0')}`;
    const days = getDaysInMonth(year, month);
    const WEEKDAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    const dailyRows = [];
    const dayIndex = {};
    for (let d = 1; d <= days; d++) {
      const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
      const rowDay = {
        date: dateStr,
        day: d,
        weekday: WEEKDAYS[new Date(year, month - 1, d).getDay()],
        food: 0, bev: 0, wine: 0, other: 0,
        costFood: 0, costBev: 0, costWine: 0, costOther: 0,
        coversFood: 0, coversBev: 0, coversWine: 0, coversOther: 0,
        revTotal: 0, covers: 0, rawCost: 0, netCost: 0,
        hasData: false
      };
      dailyRows.push(rowDay);
      dayIndex[dateStr] = rowDay;
    }

    const totals = { food: 0, bev: 0, wine: 0, other: 0, revTotal: 0, covers: 0, coversFood: 0, coversBev: 0, coversWine: 0, coversOther: 0, rawCost: 0, netCost: 0, costFood: 0, costBev: 0, costWine: 0, costOther: 0 };
    const netPoolTotals = { food: 0, bev: 0, wine: 0, other: 0 };
    const dayCatPool = {};
    const bumpPool = (d, cat, amount) => {
      if (!dayCatPool[d]) dayCatPool[d] = { food: 0, bev: 0, wine: 0, other: 0 };
      dayCatPool[d][cat] += amount;
    };
    const includeOutlet = (out) => {
      if (outletFilter !== 'ALL') return out === outletFilter;
      return out !== 'Canteen';
    };

    if (fs.existsSync(EXCEL_FILE)) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(EXCEL_FILE);
      const dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');
      const purSheet = workbook.getWorksheet('Purchases_Issues');
      const costAdjSheet = workbook.getWorksheet('Cost_Adjustments');

      if (dataSheet) {
        let colDate = 1, colOutlet = 2, colCat = 3, colRev = 5, colCover = 6, colCost = 7;
        const headerRow = dataSheet.getRow(1);
        headerRow.eachCell((cell, colNumber) => {
          const text = String(cell.value || '').trim().toLowerCase();
          if (text.startsWith('date')) colDate = colNumber;
          else if (text.startsWith('outlet')) colOutlet = colNumber;
          else if (text.includes('categor')) colCat = colNumber;
          else if (text.startsWith('revenue')) colRev = colNumber;
          else if (text.startsWith('cover')) colCover = colNumber;
          else if (text.startsWith('cost')) colCost = colNumber;
        });

        dataSheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const d = getCellDateStr(row.getCell(colDate).value);
          const rowDay = dayIndex[d];
          if (!rowDay) return;
          const out = String(row.getCell(colOutlet).value || '').trim();
          if (!includeOutlet(out)) return;
          const cat = String(row.getCell(colCat).value || '').trim();
          const rev = Number(row.getCell(colRev).value) || 0;
          const cov = Number(row.getCell(colCover).value) || 0;
          const cost = Number(row.getCell(colCost).value) || 0;

          rowDay.hasData = true;
          rowDay.revTotal += rev;
          rowDay.covers += cov;
          rowDay.rawCost += cost;
          if (cat === 'RV-Food') { rowDay.food += rev; rowDay.coversFood += cov; }
          else if (cat === 'RV-Beverage') { rowDay.bev += rev; rowDay.coversBev += cov; }
          else if (cat === 'RV-Wine') { rowDay.wine += rev; rowDay.coversWine += cov; }
          else if (cat === 'RV-Other') { rowDay.other += rev; rowDay.coversOther += cov; }
        });
      }

      // ซื้อตรง/เบิกของ — dedupe ตาม date+outlet เหมือน dashboard-summary กันยอดซ้ำ
      const purMap = {};
      if (purSheet) {
        const layout = purSheetLayout(purSheet);
        purSheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const d = getCellDateStr(row.getCell(1).value);
          if (!dayIndex[d]) return;
          const out = String(row.getCell(2).value || '').trim();
          if (!includeOutlet(out)) return;
          let pf = 0, pb = 0, iF = 0, iw = 0, ib = 0, iO = 0;
          if (layout === 'v3') {
            pf = Number(row.getCell(3).value) || 0;
            pb = Number(row.getCell(4).value) || 0;
            iF = Number(row.getCell(5).value) || 0;
            iw = Number(row.getCell(6).value) || 0;
            ib = Number(row.getCell(7).value) || 0;
            iO = Number(row.getCell(8).value) || 0;
          } else if (layout === 'v2') {
            pf = Number(row.getCell(3).value) || 0;
            pb = Number(row.getCell(4).value) || 0;
            iw = Number(row.getCell(5).value) || 0;
            ib = Number(row.getCell(6).value) || 0;
          } else {
            pf = Number(row.getCell(3).value) || 0;
            ib = Number(row.getCell(4).value) || 0;
          }
          purMap[`${d}_${out}`] = { d, pf, pb, iF, iw, ib, iO, pur: pf + pb, iss: iF + iw + ib + iO };
        });
      }

      Object.values(purMap).forEach(item => {
        const rowDay = dayIndex[item.d];
        rowDay.hasData = true;
        rowDay.netCost += item.pur + item.iss;
        bumpPool(item.d, 'food', item.pf + item.iF);
        bumpPool(item.d, 'bev', item.pb + item.ib);
        bumpPool(item.d, 'wine', item.iw);
        bumpPool(item.d, 'other', item.iO);
      });

      const adjMap = {};
      if (costAdjSheet) {
        const layout = adjSheetLayout(costAdjSheet);
        costAdjSheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const d = getCellDateStr(row.getCell(1).value);
          if (!dayIndex[d]) return;
          const out = String(row.getCell(2).value || '').trim();
          if (!includeOutlet(out)) return;
          let tInF = 0, tInB = 0, tInW = 0, tInO = 0, tOutF = 0, tOutB = 0, tOutW = 0, tOutO = 0, crF = 0, crB = 0, crW = 0, crO = 0;
          if (layout === 'v2') {
            tInF = Number(row.getCell(3).value) || 0;
            tInB = Number(row.getCell(4).value) || 0;
            tInW = Number(row.getCell(5).value) || 0;
            tInO = Number(row.getCell(6).value) || 0;
            tOutF = Number(row.getCell(7).value) || 0;
            tOutB = Number(row.getCell(8).value) || 0;
            tOutW = Number(row.getCell(9).value) || 0;
            tOutO = Number(row.getCell(10).value) || 0;
            crF = Number(row.getCell(11).value) || 0;
            crB = Number(row.getCell(12).value) || 0;
            crW = Number(row.getCell(13).value) || 0;
            crO = Number(row.getCell(14).value) || 0;
          } else {
            const tIn = (Number(row.getCell(3).value) || 0) / 4;
            const tOut = (Number(row.getCell(4).value) || 0) / 4;
            const cr = (Number(row.getCell(5).value) || 0) / 4;
            tInF = tIn; tInB = tIn; tInW = tIn; tInO = tIn;
            tOutF = tOut; tOutB = tOut; tOutW = tOut; tOutO = tOut;
            crF = cr; crB = cr; crW = cr; crO = cr;
          }
          adjMap[`${d}_${out}`] = {
            d,
            tIn: tInF + tInB + tInW + tInO,
            tOut: tOutF + tOutB + tOutW + tOutO,
            cr: crF + crB + crW + crO,
            tInF, tInB, tInW, tInO, tOutF, tOutB, tOutW, tOutO, crF, crB, crW, crO
          };
        });
      }

      Object.values(adjMap).forEach(item => {
        const rowDay = dayIndex[item.d];
        rowDay.hasData = true;
        rowDay.netCost += item.tIn - item.tOut - item.cr;
        bumpPool(item.d, 'food', item.tInF - item.tOutF - item.crF);
        bumpPool(item.d, 'bev', item.tInB - item.tOutB - item.crB);
        bumpPool(item.d, 'wine', item.tInW - item.tOutW - item.crW);
        bumpPool(item.d, 'other', item.tInO - item.tOutO - item.crO);
      });

      // Net Cost รายหมวดของแต่ละวัน จาก pool (Purchase + Issue + Transfer In − Transfer Out − Cr Cost)
      Object.keys(dayCatPool).forEach(d => {
        const rowDay = dayIndex[d];
        const p = dayCatPool[d];
        rowDay.costFood += p.food;
        rowDay.costBev += p.bev;
        rowDay.costWine += p.wine;
        rowDay.costOther += p.other;
      });
    }

    dailyRows.forEach(r => {
      totals.food += r.food;
      totals.bev += r.bev;
      totals.wine += r.wine;
      totals.other += r.other;
      totals.revTotal += r.revTotal;
      totals.covers += r.covers;
      totals.coversFood += r.coversFood;
      totals.coversBev += r.coversBev;
      totals.coversWine += r.coversWine;
      totals.coversOther += r.coversOther;
      totals.rawCost += r.rawCost;
      totals.netCost += r.netCost;
      totals.costFood += r.costFood;
      totals.costBev += r.costBev;
      totals.costWine += r.costWine;
      totals.costOther += r.costOther;
    });
    Object.values(dayCatPool).forEach(p => {
      netPoolTotals.food += p.food;
      netPoolTotals.bev += p.bev;
      netPoolTotals.wine += p.wine;
      netPoolTotals.other += p.other;
    });

    const kpi = {
      avgCheck: totals.covers > 0 ? totals.revTotal / totals.covers : 0,
      foodNetCost: netPoolTotals.food,
      bevNetCost: netPoolTotals.bev,
      foodCostPct: totals.food > 0 ? (netPoolTotals.food / totals.food) * 100 : 0,
      bevCostPct: totals.bev > 0 ? (netPoolTotals.bev / totals.bev) * 100 : 0,
      costPct: totals.revTotal > 0 ? (totals.netCost / totals.revTotal) * 100 : 0
    };

    res.json({
      success: true,
      data: {
        year, month, monthKey, daysInMonth: days, outlet: outletFilter,
        daily: dailyRows,
        totals,
        kpi
      }
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// ===== Up Sale / Goals (เป้ายอดขายเพิ่มรายเดือน + บันทึกยอดทำได้รายวัน) =====
// เก็บใน upsell.json รูปแบบเดียวกับ budget.json/topsale.json:
//   { goals: { "2026-09": { "<Outlet>": { monthly, daily[] } } },
//     entries: [ { id, date, outlet, amount, note, timestamp } ] }
const UPSELL_FILE = path.join(__dirname, 'upsell.json');

function readUpsellStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(UPSELL_FILE, 'utf8'));
    return parsed && typeof parsed === 'object' ? parsed : { goals: {}, entries: [] };
  } catch (err) {
    return { goals: {}, entries: [] };
  }
}

function writeUpsellStore(store) {
  fs.writeFileSync(UPSELL_FILE, JSON.stringify(store, null, 2));
}

// API อ่านเป้า + รายการของเดือน — ?year&month
app.get('/api/upsell', requireAuth(), (req, res) => {
  try {
    const year = parseInt(req.query.year, 10);
    const month = parseInt(req.query.month, 10);
    if (!year || !month || month < 1 || month > 12) {
      return res.status(400).json({ success: false, message: 'ระบุ year และ month ให้ถูกต้อง' });
    }
    const monthKey = `${year}-${String(month).padStart(2, '0')}`;
    const store = readUpsellStore();
    res.json({
      success: true,
      data: {
        month: monthKey,
        days: getDaysInMonth(year, month),
        goals: (store.goals || {})[monthKey] || {},
        entries: (store.entries || []).filter(e => String(e.date || '').startsWith(monthKey))
      }
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// ตั้งเป้ารายเดือนต่อ outlet (admin) — กระจายเป้ารายวันแบบ Even อัตโนมัติ (เศษไปวันสุดท้าย)
// ส่ง monthly = 0 เพื่อล้างเป้าของ outlet นั้น
app.post('/api/upsell/goal', requireAuth('admin'), (req, res) => {
  try {
    const { year, month, outlet, monthly } = req.body || {};
    const y = parseInt(year, 10);
    const m = parseInt(month, 10);
    if (!y || !m || m < 1 || m > 12) {
      return res.status(400).json({ success: false, message: 'ระบุ year และ month ให้ถูกต้อง' });
    }
    const outletName = String(outlet || '').trim();
    if (!outletName || !getOutlets().some(o => o.toLowerCase() === outletName.toLowerCase())) {
      return res.status(400).json({ success: false, message: 'Outlet ไม่ถูกต้อง' });
    }
    const monthlyNum = Math.max(0, Number(monthly) || 0);
    const monthKey = `${y}-${String(m).padStart(2, '0')}`;
    const store = readUpsellStore();
    if (!store.goals || typeof store.goals !== 'object') store.goals = {};
    if (!store.goals[monthKey]) store.goals[monthKey] = {};
    if (monthlyNum > 0) {
      store.goals[monthKey][outletName] = { monthly: monthlyNum, daily: evenSplitMonthly(monthlyNum, getDaysInMonth(y, m), false) };
    } else {
      delete store.goals[monthKey][outletName];
    }
    writeUpsellStore(store);
    res.json({ success: true, message: monthlyNum > 0 ? `บันทึกเป้า Up Sale ของ ${outletName} เดือน ${monthKey} เรียบร้อยแล้ว` : `ล้างเป้า Up Sale ของ ${outletName} เดือน ${monthKey} แล้ว` });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// บันทึกยอด Up Sale รายวัน (admin) — 1 แถวต่อวัน/outlet บันทึกซ้ำวันเดียวกัน = แทนที่ค่าเดิม
app.post('/api/upsell/entry', requireAuth('admin'), (req, res) => {
  try {
    const { date, outlet, amount, note } = req.body || {};
    const dateStr = getCellDateStr(date);
    const outletName = String(outlet || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      return res.status(400).json({ success: false, message: 'รูปแบบวันที่ไม่ถูกต้อง' });
    }
    if (!outletName || !getOutlets().some(o => o.toLowerCase() === outletName.toLowerCase())) {
      return res.status(400).json({ success: false, message: 'Outlet ไม่ถูกต้อง' });
    }
    const amountNum = Math.max(0, Number(amount) || 0);
    const noteStr = String(note || '').trim();
    const store = readUpsellStore();
    if (!Array.isArray(store.entries)) store.entries = [];
    store.entries = store.entries.filter(e => !(e.date === dateStr && String(e.outlet || '').trim() === outletName));
    if (amountNum > 0 || noteStr) {
      store.entries.push({
        id: Date.now(),
        date: dateStr,
        outlet: outletName,
        amount: amountNum,
        note: noteStr,
        timestamp: getFormattedTimestamp()
      });
      store.entries.sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.outlet).localeCompare(String(b.outlet)));
    }
    writeUpsellStore(store);
    res.json({ success: true, message: `บันทึกยอด Up Sale ของ ${outletName} วันที่ ${dateStr} เรียบร้อยแล้ว` });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// ลบรายการ Up Sale ตาม id (admin)
app.delete('/api/upsell/entry', requireAuth('admin'), (req, res) => {
  try {
    const id = Number(req.body && req.body.id);
    const store = readUpsellStore();
    const before = Array.isArray(store.entries) ? store.entries.length : 0;
    store.entries = (store.entries || []).filter(e => Number(e.id) !== id);
    writeUpsellStore(store);
    const removed = store.entries.length < before;
    res.json({ success: removed, message: removed ? 'ลบรายการเรียบร้อยแล้ว' : 'ไม่พบรายการที่ต้องการลบ' });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/delete-records', requireAuth('admin'), async (req, res) => {
  try {
    const { date, outlet } = req.body;
    if (!fs.existsSync(EXCEL_FILE)) return res.json({ success: false, message: 'ไม่พบไฟล์ฐานข้อมูล Excel' });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);

    const ALL_OUTLETS = getOutlets();
    const sheetsToCheck = ['Data', 'Revenue_Cost', 'Purchases_Issues', 'Cost_Adjustments', 'Hotel_Stats', 'Outlet_Summary', ...ALL_OUTLETS];

    sheetsToCheck.forEach(sheetName => {
      const sheet = workbook.getWorksheet(sheetName);
      if (sheet) {
        for (let i = sheet.rowCount; i >= 2; i--) {
          const row = sheet.getRow(i);
          const rDate = getCellDateStr(row.getCell(1).value);
          const rOutlet = String(row.getCell(2).value || '').trim();

          if (sheetName === 'Data' || sheetName === 'Revenue_Cost' || sheetName === 'Outlet_Summary') {
            if (rDate === date && (!outlet || outlet === 'ALL' || rOutlet === outlet)) {
              sheet.spliceRows(i, 1);
            }
          } 
          else if (ALL_OUTLETS.includes(sheetName)) {
            if (rDate === date && (!outlet || outlet === 'ALL' || outlet === sheetName)) {
              sheet.spliceRows(i, 1);
            }
          } 
          else {
            if (rDate === date && (!outlet || outlet === 'ALL' || !rOutlet || rOutlet === outlet)) {
              sheet.spliceRows(i, 1);
            }
          }
        }
      }
    });

    updateDynamicPivotSheet(workbook);
    backupExcelFile();

    await saveWorkbook(workbook);
    res.json({ success: true, message: `ลบข้อมูลวันที่ ${date} ของ ${outlet || 'All Outlets'} สำเร็จทุกชีท` });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get('/api/open-excel', requireAuth('admin'), (req, res) => {
  if (!fs.existsSync(EXCEL_FILE)) return res.json({ success: false, message: 'ยังไม่มีไฟล์ Excel ถูกสร้างขึ้น' });

  const command = process.platform === 'win32' ? `start "" "${EXCEL_FILE}"` :
                  process.platform === 'darwin' ? `open "${EXCEL_FILE}"` : `xdg-open "${EXCEL_FILE}"`;

  exec(command, (err) => {
    if (err) return res.json({ success: false, message: 'ไม่สามารถเปิดไฟล์ได้: ' + err.message });
    res.json({ success: true, message: 'เปิดไฟล์สำเร็จ' });
  });
});

app.post('/api/refresh-excel', requireAuth(), async (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.json({ success: false, message: 'ไม่พบไฟล์ Excel ในระบบ' });
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    res.json({ success: true, message: 'รีเฟรชและโหลดข้อมูลล่าสุดจาก Excel สำเร็จแล้ว' });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/rebuild-excel', requireAuth('admin'), async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const ALL_OUTLETS = getOutlets();

    const dataSheet = workbook.addWorksheet('Data');
    setupSheetColumns(dataSheet, DATA_COLUMNS);

    ALL_OUTLETS.forEach(outletName => {
      const s = workbook.addWorksheet(outletName);
      setupSheetColumns(s, DATA_COLUMNS);
    });

    const purSheet = workbook.addWorksheet('Purchases_Issues');
    setupSheetColumns(purSheet, PUR_COLUMNS);

    const costAdjSheet = workbook.addWorksheet('Cost_Adjustments');
    setupSheetColumns(costAdjSheet, COST_ADJ_COLUMNS);

    const statsSheet = workbook.addWorksheet('Hotel_Stats');
    setupSheetColumns(statsSheet, HOTEL_STATS_COLUMNS);

    const summarySheet = workbook.addWorksheet('Outlet_Summary');
    setupSheetColumns(summarySheet, SUMMARY_COLUMNS);

    updateDynamicPivotSheet(workbook);
    backupExcelFile();
    await saveWorkbook(workbook);
    res.json({ success: true, message: 'สร้างโครงสร้างไฟล์ Excel ใหม่สำเร็จ' });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/seed-random', requireAuth('admin'), async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const ALL_OUTLETS = getOutlets();

    const dataSheet = workbook.addWorksheet('Data');
    setupSheetColumns(dataSheet, DATA_COLUMNS);

    const outletSheetsMap = {};
    ALL_OUTLETS.forEach(outName => {
      const s = workbook.addWorksheet(outName);
      setupSheetColumns(s, DATA_COLUMNS);
      outletSheetsMap[outName] = s;
    });

    const purSheet = workbook.addWorksheet('Purchases_Issues');
    setupSheetColumns(purSheet, PUR_COLUMNS);

    const costAdjSheet = workbook.addWorksheet('Cost_Adjustments');
    setupSheetColumns(costAdjSheet, COST_ADJ_COLUMNS);

    const statsSheet = workbook.addWorksheet('Hotel_Stats');
    setupSheetColumns(statsSheet, HOTEL_STATS_COLUMNS);

    const summarySheet = workbook.addWorksheet('Outlet_Summary');
    setupSheetColumns(summarySheet, SUMMARY_COLUMNS);

    const meals = ['ABF', 'Breakfast', 'Lunch', 'Dinner', 'Supper'];
    const types = [
      { type: 'Food', cat: 'RV-Food' },
      { type: 'Beverage', cat: 'RV-Beverage' },
      { type: 'Wine', cat: 'RV-Wine' },
      { type: 'Other', cat: 'RV-Other' }
    ];
    const weathers = ['Sunny', 'Cloudy', 'Rainy', 'Stormy'];

    const toLocalDateStr = (dt) => `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
    const today = new Date();
    for (let i = 30; i >= 0; i--) {
      const d = new Date(today);
      d.setDate(d.getDate() - i);
      const dateStr = toLocalDateStr(d);
      const randomTime = `${String(Math.floor(Math.random() * 12) + 10).padStart(2, '0')}:${String(Math.floor(Math.random() * 60)).padStart(2, '0')}:${String(Math.floor(Math.random() * 60)).padStart(2, '0')}`;
      const rowTimestamp = `${dateStr} ${randomTime}`;

      const roomAvail = 120;
      const roomSold = Math.floor(Math.random() * 40) + 70;
      const comHu = Math.floor(Math.random() * 10) + 2;
      const guests = Math.floor(roomSold * (1.6 + Math.random() * 0.4));
      const weather = weathers[Math.floor(Math.random() * weathers.length)];

      statsSheet.addRow({
        date: dateStr,
        roomAvailable: roomAvail,
        roomSold: roomSold,
        comHuRoom: comHu,
        totalGuest: guests,
        weather: weather,
        timestamp: `${dateStr} ${randomTime}`
      });

      ALL_OUTLETS.forEach(out => {
        const purFoodVal = Math.floor(Math.random() * 5000) + 800;
        const purBevVal = Math.floor(Math.random() * 2500) + 200;
        const issFoodVal = Math.floor(Math.random() * 1200) + 100;
        const issWineVal = Math.floor(Math.random() * 1500) + 100;
        const issBevVal = Math.floor(Math.random() * 2500) + 300;
        const issOtherVal = Math.floor(Math.random() * 800) + 50;
        const tInFVal = Math.floor(Math.random() * 300) + 20;
        const tInBVal = Math.floor(Math.random() * 200) + 10;
        const tInWVal = Math.floor(Math.random() * 150) + 10;
        const tInOVal = Math.floor(Math.random() * 100) + 10;
        const tOutFVal = Math.floor(Math.random() * 200) + 10;
        const tOutBVal = Math.floor(Math.random() * 150) + 10;
        const tOutWVal = Math.floor(Math.random() * 100) + 5;
        const tOutOVal = Math.floor(Math.random() * 80) + 5;
        const crFVal = Math.floor(Math.random() * 600) + 50;
        const crBVal = Math.floor(Math.random() * 400) + 30;
        const crWVal = Math.floor(Math.random() * 300) + 20;
        const crOVal = Math.floor(Math.random() * 200) + 10;

        if (out === 'Canteen') {
          const canteenCover = Math.floor(Math.random() * 50) + 80;
          const canteenCost = canteenCover * (55 + Math.random() * 10);

          const rowData = {
            date: dateStr,
            outlet: 'Canteen',
            category: 'RV-Food',
            meal: 'Lunch',
            revenue: 0,
            cover: canteenCover,
            cost: Number(canteenCost.toFixed ? canteenCost.toFixed(2) : canteenCost),
            timestamp: rowTimestamp
          };

          dataSheet.addRow(rowData);
          if (outletSheetsMap['Canteen']) outletSheetsMap['Canteen'].addRow(rowData);

          summarySheet.addRow({
            date: dateStr,
            outlet: 'Canteen',
            meal: 'Lunch',
            foodRev: 0,
            bevRev: 0,
            wineRev: 0,
            otherRev: 0,
            totalRev: 0,
            cover: canteenCover,
            foodCost: rowData.cost,
            bevCost: 0,
            wineCost: 0,
            otherCost: 0,
            totalCost: rowData.cost,
            purchaseFood: purFoodVal,
            purchaseBev: purBevVal,
            issueFood: issFoodVal,
            issueWine: issWineVal,
            issueBev: issBevVal,
            issueOther: issOtherVal,
            tInFood: tInFVal,
            tInBev: tInBVal,
            tInWine: tInWVal,
            tInOther: tInOVal,
            tOutFood: tOutFVal,
            tOutBev: tOutBVal,
            tOutWine: tOutWVal,
            tOutOther: tOutOVal,
            crFood: crFVal,
            crBev: crBVal,
            crWine: crWVal,
            crOther: crOVal,
            timestamp: rowTimestamp
          });

        } else {
          const mealMap = {};
          meals.forEach(m => {
            mealMap[m] = { foodRev: 0, bevRev: 0, wineRev: 0, otherRev: 0, cover: 0, foodCost: 0, bevCost: 0, wineCost: 0, otherCost: 0 };
            types.forEach(t => {
              const rev = Math.floor(Math.random() * 5000) + 500;
              const cov = Math.floor(Math.random() * 20) + 1;
              const cost = Math.floor(rev * (0.28 + Math.random() * 0.12));

              const rowData = {
                date: dateStr,
                outlet: out,
                category: t.cat,
                meal: m,
                revenue: rev,
                cover: cov,
                cost: cost,
                timestamp: rowTimestamp
              };

              dataSheet.addRow(rowData);
              if (outletSheetsMap[out]) outletSheetsMap[out].addRow(rowData);

              if (t.cat === 'RV-Food') {
                mealMap[m].foodRev += rev;
                mealMap[m].foodCost += cost;
              } else if (t.cat === 'RV-Wine') {
                mealMap[m].wineRev += rev;
                mealMap[m].wineCost += cost;
              } else if (t.cat === 'RV-Other') {
                mealMap[m].otherRev += rev;
                mealMap[m].otherCost += cost;
              } else {
                mealMap[m].bevRev += rev;
                mealMap[m].bevCost += cost;
              }
              mealMap[m].cover += cov;
            });
          });

          meals.forEach(m => {
            const item = mealMap[m];
            const isCarrier = m === 'Dinner'; // ยอดปรับปรุงใส่แค่แถว Dinner
            summarySheet.addRow({
              date: dateStr,
              outlet: out,
              meal: m,
              foodRev: item.foodRev,
              bevRev: item.bevRev,
              wineRev: item.wineRev,
              otherRev: item.otherRev,
              totalRev: item.foodRev + item.bevRev + item.wineRev + item.otherRev,
              cover: item.cover,
              foodCost: item.foodCost,
              bevCost: item.bevCost,
              wineCost: item.wineCost,
              otherCost: item.otherCost,
              totalCost: item.foodCost + item.bevCost + item.wineCost + item.otherCost,
              purchaseFood: isCarrier ? purFoodVal : 0,
              purchaseBev: isCarrier ? purBevVal : 0,
              issueFood: isCarrier ? issFoodVal : 0,
              issueWine: isCarrier ? issWineVal : 0,
              issueBev: isCarrier ? issBevVal : 0,
              issueOther: isCarrier ? issOtherVal : 0,
              tInFood: isCarrier ? tInFVal : 0,
              tInBev: isCarrier ? tInBVal : 0,
              tInWine: isCarrier ? tInWVal : 0,
              tInOther: isCarrier ? tInOVal : 0,
              tOutFood: isCarrier ? tOutFVal : 0,
              tOutBev: isCarrier ? tOutBVal : 0,
              tOutWine: isCarrier ? tOutWVal : 0,
              tOutOther: isCarrier ? tOutOVal : 0,
              crFood: isCarrier ? crFVal : 0,
              crBev: isCarrier ? crBVal : 0,
              crWine: isCarrier ? crWVal : 0,
              crOther: isCarrier ? crOVal : 0,
              timestamp: rowTimestamp
            });
          });
        }

        const randomTimePur = `${String(Math.floor(Math.random() * 12) + 10).padStart(2, '0')}:${String(Math.floor(Math.random() * 60)).padStart(2, '0')}:${String(Math.floor(Math.random() * 60)).padStart(2, '0')}`;
        purSheet.addRow({
          date: dateStr,
          outlet: out,
          purchaseFood: purFoodVal,
          purchaseBev: purBevVal,
          issueFood: issFoodVal,
          issueWine: issWineVal,
          issueBev: issBevVal,
          issueOther: issOtherVal,
          timestamp: `${dateStr} ${randomTimePur}`
        });

        costAdjSheet.addRow({
          date: dateStr,
          outlet: out,
          tInFood: tInFVal,
          tInBev: tInBVal,
          tInWine: tInWVal,
          tInOther: tInOVal,
          tOutFood: tOutFVal,
          tOutBev: tOutBVal,
          tOutWine: tOutWVal,
          tOutOther: tOutOVal,
          crFood: crFVal,
          crBev: crBVal,
          crWine: crWVal,
          crOther: crOVal,
          timestamp: `${dateStr} ${randomTimePur}`
        });
      });
    }

    sortSheetByDate(statsSheet);
    sortSheetByDate(summarySheet);
    updateDynamicPivotSheet(workbook);

    backupExcelFile();
    await saveWorkbook(workbook);
    res.json({ success: true, message: 'สุ่มสร้างข้อมูล 30 วันสำเร็จ' });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/get-record-by-date', requireAuth(), async (req, res) => {
  try {
    const { date, outlet } = req.body;
    if (!fs.existsSync(EXCEL_FILE)) return res.json({ success: false, data: null });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);

    const dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');
    const purSheet = workbook.getWorksheet('Purchases_Issues');
    const costAdjSheet = workbook.getWorksheet('Cost_Adjustments');
    const statsSheet = workbook.getWorksheet('Hotel_Stats');

    const records = [];
    let purchaseFood = 0, purchaseBev = 0, issueFood = 0, issueWine = 0, issueBev = 0, issueOther = 0;
    let tInFood = 0, tInBev = 0, tInWine = 0, tInOther = 0;
    let tOutFood = 0, tOutBev = 0, tOutWine = 0, tOutOther = 0;
    let crFood = 0, crBev = 0, crWine = 0, crOther = 0;
    let roomAvailable = 0, roomSold = 0, comHuRoom = 0, totalGuest = 0, weather = 'Sunny';

    if (dataSheet) {
      dataSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const rDate = getCellDateStr(row.getCell(1).value);
        const rOutlet = String(row.getCell(2).value || '').trim();
        if (rDate === date && rOutlet === outlet) {
          records.push({
            category: String(row.getCell(3).value || '').trim(),
            meal: String(row.getCell(4).value || '').trim(),
            revenue: Number(row.getCell(5).value) || 0,
            cover: Number(row.getCell(6).value) || 0,
            cost: Number(row.getCell(7).value) || 0
          });
        }
      });
    }

    if (purSheet) {
      const layout = purSheetLayout(purSheet);
      purSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const rDate = getCellDateStr(row.getCell(1).value);
        const rOutlet = String(row.getCell(2).value || '').trim();
        if (rDate === date && rOutlet === outlet) {
          if (layout === 'v3') {
            purchaseFood = Number(row.getCell(3).value) || 0;
            purchaseBev = Number(row.getCell(4).value) || 0;
            issueFood = Number(row.getCell(5).value) || 0;
            issueWine = Number(row.getCell(6).value) || 0;
            issueBev = Number(row.getCell(7).value) || 0;
            issueOther = Number(row.getCell(8).value) || 0;
          } else if (layout === 'v2') {
            purchaseFood = Number(row.getCell(3).value) || 0;
            purchaseBev = Number(row.getCell(4).value) || 0;
            issueWine = Number(row.getCell(5).value) || 0;
            issueBev = Number(row.getCell(6).value) || 0;
          } else {
            // ชีตเก่า: Direct Purchase → Food, Store Issue → Bev
            purchaseFood = Number(row.getCell(3).value) || 0;
            issueBev = Number(row.getCell(4).value) || 0;
          }
        }
      });
    }

    if (costAdjSheet) {
      const layout = adjSheetLayout(costAdjSheet);
      costAdjSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const rDate = getCellDateStr(row.getCell(1).value);
        const rOutlet = String(row.getCell(2).value || '').trim();
        if (rDate === date && rOutlet === outlet) {
          if (layout === 'v2') {
            tInFood = Number(row.getCell(3).value) || 0;
            tInBev = Number(row.getCell(4).value) || 0;
            tInWine = Number(row.getCell(5).value) || 0;
            tInOther = Number(row.getCell(6).value) || 0;
            tOutFood = Number(row.getCell(7).value) || 0;
            tOutBev = Number(row.getCell(8).value) || 0;
            tOutWine = Number(row.getCell(9).value) || 0;
            tOutOther = Number(row.getCell(10).value) || 0;
            crFood = Number(row.getCell(11).value) || 0;
            crBev = Number(row.getCell(12).value) || 0;
            crWine = Number(row.getCell(13).value) || 0;
            crOther = Number(row.getCell(14).value) || 0;
          } else {
            const tIn = (Number(row.getCell(3).value) || 0) / 4;
            const tOut = (Number(row.getCell(4).value) || 0) / 4;
            const cr = (Number(row.getCell(5).value) || 0) / 4;
            tInFood = tIn; tInBev = tIn; tInWine = tIn; tInOther = tIn;
            tOutFood = tOut; tOutBev = tOut; tOutWine = tOut; tOutOther = tOut;
            crFood = cr; crBev = cr; crWine = cr; crOther = cr;
          }
        }
      });
    }

    if (statsSheet) {
      statsSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const rDate = getCellDateStr(row.getCell(1).value);
        if (rDate === date) {
          roomAvailable = Number(row.getCell(2).value) || 0;
          roomSold = Number(row.getCell(3).value) || 0;
          comHuRoom = Number(row.getCell(4).value) || 0;
          totalGuest = Number(row.getCell(5).value) || 0;
          weather = String(row.getCell(6).value || 'Sunny').trim();
        }
      });
    }

    res.json({
      success: true,
      data: {
        purchase: purchaseFood + purchaseBev,
        issue: issueFood + issueWine + issueBev + issueOther,
        purchaseFood, purchaseBev, issueFood, issueWine, issueBev, issueOther,
        transferIn: tInFood + tInBev + tInWine + tInOther,
        transferOut: tOutFood + tOutBev + tOutWine + tOutOther,
        crCost: crFood + crBev + crWine + crOther,
        tInFood, tInBev, tInWine, tInOther,
        tOutFood, tOutBev, tOutWine, tOutOther,
        crFood, crBev, crWine, crOther,
        roomAvailable, roomSold, comHuRoom, totalGuest, weather, records
      }
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// ===== เส้น API: Shift & Meal Summary (ปรับสูตร Allocated Net Cost) =====
app.get('/api/shift-summary', requireAuth(), async (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.json({ success: true, data: { rows: [], chartData: { labels: [], datasets: [] }, dateRange: { min: '', max: '' } } });
    }

    const { startDate, endDate, mode } = req.query;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);

    const dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');
    const purSheet = workbook.getWorksheet('Purchases_Issues');
    const costAdjSheet = workbook.getWorksheet('Cost_Adjustments');
    let dates = [];
    const aggregator = {};
    const outletCatTotals = {};
    const outletTotals = {}; 
    const mealsList = ['ABF', 'Breakfast', 'Lunch', 'Dinner', 'Supper'];
    const mealColors = {
      'ABF': '#D9A441',
      'Breakfast': '#7B9FC7',
      'Lunch': '#A98CB5',
      'Dinner': '#E8836B',
      'Supper': '#9DBB9A'
    };

    if (dataSheet) {
      dataSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const d = getCellDateStr(row.getCell(1).value);
        const out = String(row.getCell(2).value || '').trim();
        const cat = String(row.getCell(3).value || '').trim();
        const meal = String(row.getCell(4).value || '').trim();
        const rev = Number(row.getCell(5).value) || 0;
        const cov = Number(row.getCell(6).value) || 0;
        const cost = Number(row.getCell(7).value) || 0;

        if (d) dates.push(d);
        if (startDate && d < startDate) return;
        if (endDate && d > endDate) return;
        if (!out || out === 'Canteen') return;

        const key = `${out}_${meal}`;
        if (!aggregator[key]) {
          aggregator[key] = { 
            outlet: out, 
            meal: meal, 
            foodRev: 0, foodCost: 0,
            bevRev: 0, bevCost: 0,
            wineRev: 0, wineCost: 0,
            otherRev: 0, otherCost: 0,
            totalRev: 0, totalCost: 0,
            cover: 0, 
            rawTotalCost: 0 
          };
        }

        aggregator[key].cover += cov;
        aggregator[key].rawTotalCost += cost;
        aggregator[key].totalRev += rev;
        aggregator[key].totalCost += cost;

        if (cat === 'RV-Food') {
          aggregator[key].foodRev += rev;
          aggregator[key].foodCost += cost;
        } else if (cat === 'RV-Beverage') {
          aggregator[key].bevRev += rev;
          aggregator[key].bevCost += cost;
        } else if (cat === 'RV-Wine') {
          aggregator[key].wineRev += rev;
          aggregator[key].wineCost += cost;
        } else if (cat === 'RV-Other') {
          aggregator[key].otherRev += rev;
          aggregator[key].otherCost += cost;
        }

        if (!outletTotals[out]) {
          outletTotals[out] = { totalRev: 0, totalRawCost: 0 };
          outletCatTotals[out] = { food: 0, bev: 0, wine: 0, other: 0, foodRev: 0, bevRev: 0, wineRev: 0, otherRev: 0 };
        }
        outletTotals[out].totalRev += rev;
        outletTotals[out].totalRawCost += cost;

        if (cat === 'RV-Food') {
          outletCatTotals[out].food += cost;
          outletCatTotals[out].foodRev += rev;
        } else if (cat === 'RV-Beverage') {
          outletCatTotals[out].bev += cost;
          outletCatTotals[out].bevRev += rev;
        } else if (cat === 'RV-Wine') {
          outletCatTotals[out].wine += cost;
          outletCatTotals[out].wineRev += rev;
        } else if (cat === 'RV-Other') {
          outletCatTotals[out].other += cost;
          outletCatTotals[out].otherRev += rev;
        }
      });
    }

    // Actual Net Cost ต่อหมวดของแต่ละ Outlet:
    //   Food  = Purchase Food + Issue Food + Transfer In Food − Transfer Out Food − Cr Cost Food
    //   Bev   = Purchase Bev + Issue Bev + Transfer In Bev − Transfer Out Bev − Cr Cost Bev
    //   Wine  = Issue Wine + Transfer In Wine − Transfer Out Wine − Cr Cost Wine
    //   Other = Issue Other + Transfer In Other − Transfer Out Other − Cr Cost Other
    const outletCatPool = {};   // { food, bev, wine, other }

    if (mode === 'allocated') {
      const addToPool = (out, cat, amount) => {
        if (!outletCatPool[out]) outletCatPool[out] = { food: 0, bev: 0, wine: 0, other: 0 };
        outletCatPool[out][cat] += amount;
      };

      if (purSheet) {
        const layout = purSheetLayout(purSheet);
        purSheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const d = getCellDateStr(row.getCell(1).value);
          const out = String(row.getCell(2).value || '').trim();
          if (startDate && d < startDate) return;
          if (endDate && d > endDate) return;
          if (!out || out === 'Canteen') return;

          let pf = 0, pb = 0, iF = 0, iw = 0, ib = 0, iO = 0;
          if (layout === 'v3') {
            pf = Number(row.getCell(3).value) || 0;
            pb = Number(row.getCell(4).value) || 0;
            iF = Number(row.getCell(5).value) || 0;
            iw = Number(row.getCell(6).value) || 0;
            ib = Number(row.getCell(7).value) || 0;
            iO = Number(row.getCell(8).value) || 0;
          } else if (layout === 'v2') {
            pf = Number(row.getCell(3).value) || 0;
            pb = Number(row.getCell(4).value) || 0;
            iw = Number(row.getCell(5).value) || 0;
            ib = Number(row.getCell(6).value) || 0;
          } else {
            // ชีตเก่า: Direct Purchase → Food, Store Issue → Bev
            pf = Number(row.getCell(3).value) || 0;
            ib = Number(row.getCell(4).value) || 0;
          }

          addToPool(out, 'food', pf + iF);
          addToPool(out, 'bev', pb + ib);
          addToPool(out, 'wine', iw);
          addToPool(out, 'other', iO);
        });
      }

      if (costAdjSheet) {
        const layout = adjSheetLayout(costAdjSheet);
        costAdjSheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const d = getCellDateStr(row.getCell(1).value);
          const out = String(row.getCell(2).value || '').trim();
          if (startDate && d < startDate) return;
          if (endDate && d > endDate) return;
          if (!out || out === 'Canteen') return;

          let tInF, tInB, tInW, tInO, tOutF, tOutB, tOutW, tOutO, crF, crB, crW, crO;
          if (layout === 'v2') {
            tInF = Number(row.getCell(3).value) || 0;
            tInB = Number(row.getCell(4).value) || 0;
            tInW = Number(row.getCell(5).value) || 0;
            tInO = Number(row.getCell(6).value) || 0;
            tOutF = Number(row.getCell(7).value) || 0;
            tOutB = Number(row.getCell(8).value) || 0;
            tOutW = Number(row.getCell(9).value) || 0;
            tOutO = Number(row.getCell(10).value) || 0;
            crF = Number(row.getCell(11).value) || 0;
            crB = Number(row.getCell(12).value) || 0;
            crW = Number(row.getCell(13).value) || 0;
            crO = Number(row.getCell(14).value) || 0;
          } else {
            // คอลัมน์รวมเดิม → แบ่งเท่าๆ 4 หมวด
            const tIn = (Number(row.getCell(3).value) || 0) / 4;
            const tOut = (Number(row.getCell(4).value) || 0) / 4;
            const cr = (Number(row.getCell(5).value) || 0) / 4;
            tInF = tIn; tInB = tIn; tInW = tIn; tInO = tIn;
            tOutF = tOut; tOutB = tOut; tOutW = tOut; tOutO = tOut;
            crF = cr; crB = cr; crW = cr; crO = cr;
          }

          addToPool(out, 'food', tInF - tOutF - crF);
          addToPool(out, 'bev', tInB - tOutB - crB);
          addToPool(out, 'wine', tInW - tOutW - crW);
          addToPool(out, 'other', tInO - tOutO - crO);
        });
      }
    }

    const rows = [];
    Object.keys(aggregator).forEach(key => {
      const item = aggregator[key];
      const outSummary = outletTotals[item.outlet] || { totalRev: 0, totalRawCost: 0 };
      const catTotals = outletCatTotals[item.outlet] || { food: 0, bev: 0, wine: 0, other: 0, foodRev: 0, bevRev: 0, wineRev: 0, otherRev: 0 };
      const pool = outletCatPool[item.outlet] || { food: 0, bev: 0, wine: 0, other: 0 };

      let foodCostFinal = item.foodCost;
      let bevCostFinal = item.bevCost;
      let wineCostFinal = item.wineCost;
      let otherCostFinal = item.otherCost;
      let totalCostFinal = item.totalCost;

      if (mode === 'allocated') {
        // สัดส่วนมื้อ (fallback): Raw Cost ตาม Recipe, ถ้าไม่มีให้ใช้ Revenue
        let mealShare = 0;
        if (outSummary.totalRawCost > 0) {
          mealShare = item.rawTotalCost / outSummary.totalRawCost;
        } else if (outSummary.totalRev > 0) {
          mealShare = item.totalRev / outSummary.totalRev;
        }

        // Actual ต่อหมวด = Purchase X + Issue X + Transfer In X − Transfer Out X − Cr Cost X
        // ปันเข้ามื้อตามสัดส่วน raw cost ของหมวดนั้นใน Outlet
        const catMealShare = (catCost, catRev, catTotalCost, catTotalRev) => {
          if (catTotalCost > 0) return catCost / catTotalCost;
          if (catTotalRev > 0) return catRev / catTotalRev;
          return mealShare; // หมวดนี้ไม่มี raw cost/revenue เลย → ปันตามสัดส่วนมื้อทั่วไป
        };

        foodCostFinal = pool.food * catMealShare(item.foodCost, item.foodRev, catTotals.food, catTotals.foodRev);
        bevCostFinal = pool.bev * catMealShare(item.bevCost, item.bevRev, catTotals.bev, catTotals.bevRev);
        wineCostFinal = pool.wine * catMealShare(item.wineCost, item.wineRev, catTotals.wine, catTotals.wineRev);
        otherCostFinal = pool.other * catMealShare(item.otherCost, item.otherRev, catTotals.other, catTotals.otherRev);
        totalCostFinal = foodCostFinal + bevCostFinal + wineCostFinal + otherCostFinal;
      }

      rows.push({
        outlet: item.outlet,
        meal: item.meal,
        foodRev: item.foodRev,
        foodCost: foodCostFinal,
        bevRev: item.bevRev,
        bevCost: bevCostFinal,
        wineRev: item.wineRev,
        wineCost: wineCostFinal,
        otherRev: item.otherRev,
        otherCost: otherCostFinal,
        totalRev: item.totalRev,
        cover: item.cover,
        totalCost: totalCostFinal
      });
    });

    rows.sort((a, b) => a.outlet.localeCompare(b.outlet) || mealsList.indexOf(a.meal) - mealsList.indexOf(b.meal));

    dates.sort();
    const activeOutlets = [...new Set(rows.map(r => r.outlet))];
    const datasets = mealsList.map(mealName => {
      const dataPoints = activeOutlets.map(outletName => {
        const found = rows.find(r => r.outlet === outletName && r.meal === mealName);
        return found ? found.totalRev : 0;
      });
      return {
        label: mealName,
        data: dataPoints,
        backgroundColor: mealColors[mealName] || '#D9A441',
        borderRadius: 3
      };
    });

    res.json({
      success: true,
      data: {
        rows,
        chartData: {
          labels: activeOutlets,
          datasets: datasets
        },
        dateRange: {
          min: dates.length ? dates[0] : '',
          max: dates.length ? dates[dates.length - 1] : ''
        }
      }
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// ====================================================================
// Top Sale Food & Beverage + Top Purchasing (นำเข้าจาก Excel)
// เก็บข้อมูลเป็นไฟล์ JSON (topsale.json / toppurchase.json) ข้าง server.js
// การ import แต่ละครั้งจะ "แทนที่" ข้อมูลของวันที่ที่อยู่ในไฟล์ที่นำเข้า
// ====================================================================
const TOPSALE_FILE = path.join(__dirname, 'topsale.json');
const TOPPURCHASE_FILE = path.join(__dirname, 'toppurchase.json');

function readTopData(file) {
  try {
    const parsed = JSON.parse(fs.readFileSync(file, 'utf8'));
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function writeTopData(file, rows) {
  fs.writeFileSync(file, JSON.stringify(rows, null, 2));
}

// แปลงวันที่จาก Excel (Date / serial / ข้อความ) เป็น YYYY-MM-DD
// รองรับ: 2026-09-05, 05/09/2026, 05-09-26 (ปี 2 หลัก), 05.09.26, ปี พ.ศ. (2568)
function topNormalizeDate(value) {
  if (value instanceof Date && !isNaN(value)) {
    return `${value.getUTCFullYear()}-${String(value.getUTCMonth() + 1).padStart(2, '0')}-${String(value.getUTCDate()).padStart(2, '0')}`;
  }
  if (typeof value === 'number' && isFinite(value)) {
    const d = new Date(Math.round((value - 25569) * 86400 * 1000));
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
  }
  const s = String(value ?? '').trim();
  if (!s) return '';
  const fixYear = y => {
    let year = parseInt(y, 10);
    if (year > 2400) year -= 543; // พ.ศ. -> ค.ศ.
    return year;
  };
  const pad = n => String(n).padStart(2, '0');
  // YYYY-M-D (รองรับปี พ.ศ. เช่น 2568-09-05)
  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) {
    const year = fixYear(m[1]);
    return `${year}-${pad(m[2])}-${pad(m[3])}`;
  }
  // D-M-YY หรือ D-M-YYYY (ปี 2 หลัก: 00-68 -> 25xx/20xx, 69-99 -> 19xx)
  m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
  if (m) {
    let year = fixYear(m[3]);
    if (year < 100) year += (year <= 68 ? 2000 : 1900);
    return `${year}-${pad(m[2])}-${pad(m[1])}`;
  }
  return s;
}

function topToNumber(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const n = parseFloat(String(value ?? '').replace(/[,฿\s]/g, ''));
  return isFinite(n) ? n : 0;
}

function topCellValue(value) {
  if (value instanceof Date) return value; // Excel แปลงข้อความวันที่เป็น Date ให้เองแล้ว ห้ามแปลงเป็นค่าว่าง
  if (value && typeof value === 'object') {
    return value.result ?? value.text ?? value.hyperlink ?? '';
  }
  return value;
}

// สร้างชุด endpoint GET / import / DELETE ให้ทั้งสองหน้า
function registerTopEndpoints({ basePath, file, requiredCols, buildRow }) {
  app.get(basePath, requireAuth(), (req, res) => {
    res.json({ success: true, data: { rows: readTopData(file) } });
  });

  app.delete(basePath, requireAuth('admin'), (req, res) => {
    writeTopData(file, []);
    res.json({ success: true, message: 'ลบข้อมูลทั้งหมดเรียบร้อยแล้ว' });
  });

  app.post(`${basePath}/import`, requireAuth('admin'), upload.single('file'), async (req, res) => {
    try {
      if (!req.file || !req.file.buffer) {
        return res.status(400).json({ success: false, message: 'กรุณาแนบไฟล์ Excel หรือ CSV' });
      }

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(req.file.buffer);
      const sheet = workbook.worksheets[0];
      if (!sheet || sheet.rowCount <= 1) {
        return res.status(400).json({ success: false, message: 'ไฟล์ไม่มีข้อมูล (ต้องมีหัวตารางและข้อมูลอย่างน้อย 1 แถว)' });
      }

      const colMap = {};
      sheet.getRow(1).eachCell((cell, colNum) => {
        const val = String(topCellValue(cell.value) || '').trim().toLowerCase();
        requiredCols.forEach(([key, words]) => {
          if (!colMap[key] && words.some(word => val.includes(word))) colMap[key] = colNum;
        });
      });
      if (!colMap.item || !colMap.amount) {
        return res.status(400).json({ success: false, message: 'หาคอลัมน์ "Item (ชื่อเมนู/รายการ)" หรือ "Amount (ยอดเงิน)" ไม่พบ กรุณาใช้ไฟล์ Template ที่ระบบให้' });
      }

      const rows = [];
      for (let r = 2; r <= sheet.rowCount; r++) {
        const get = key => topCellValue(colMap[key] ? sheet.getRow(r).getCell(colMap[key]).value : null);
        const row = buildRow(get, topNormalizeDate, topToNumber);
        if (row) rows.push(row);
      }
      if (!rows.length) {
        return res.status(400).json({ success: false, message: 'ไม่พบข้อมูลที่นำเข้าได้ในไฟล์' });
      }

      const importedDates = new Set(rows.map(row => row.date).filter(Boolean));
      const existing = readTopData(file).filter(row => !importedDates.has(row.date));
      writeTopData(file, existing.concat(rows));
      res.json({ success: true, imported: rows.length, total: existing.length + rows.length, message: `นำเข้า ${rows.length} รายการสำเร็จ` });
    } catch (err) {
      res.status(500).json({ success: false, message: 'นำเข้าไม่สำเร็จ: ' + err.message });
    }
  });
}

registerTopEndpoints({
  basePath: '/api/topsale',
  file: TOPSALE_FILE,
  requiredCols: [
    ['date', ['date', 'วันที่', 'วัน']],
    ['outlet', ['outlet', 'ร้าน', 'ห้องอาหาร', 'location']],
    ['category', ['cat', 'หมวด', 'ประเภท', 'type']],
    ['item', ['item', 'เมนู', 'รายการ', 'product', 'ชื่อ']],
    ['qty', ['qty', 'จำนวน', 'quantity']],
    ['amount', ['ยอดขาย', 'sales', 'amount', 'ราคารวม', 'total']]
  ],
  buildRow: (get, normDate, toNum) => {
    const item = String(get('item') ?? '').trim();
    if (!item) return null;
    return {
      date: normDate(get('date')),
      outlet: String(get('outlet') ?? '').trim() || '-',
      category: String(get('category') ?? '').trim() || '-',
      item,
      qty: toNum(get('qty')),
      amount: toNum(get('amount'))
    };
  }
});

registerTopEndpoints({
  basePath: '/api/toppurchase',
  file: TOPPURCHASE_FILE,
  requiredCols: [
    ['date', ['date', 'วันที่', 'วัน']],
    ['outlet', ['outlet', 'ร้าน', 'ห้องอาหาร', 'location']],
    ['supplier', ['supplier', 'ผู้ขาย', 'vendor', 'บริษัท']],
    ['category', ['cat', 'หมวด', 'ประเภท', 'type']],
    ['item', ['item', 'รายการ', 'product', 'ชื่อ', 'เมนู']],
    ['qty', ['qty', 'จำนวน', 'quantity']],
    ['unit', ['unit', 'หน่วย']],
    ['amount', ['ยอดซื้อ', 'amount', 'total', 'มูลค่า', 'ราคา']]
  ],
  buildRow: (get, normDate, toNum) => {
    const item = String(get('item') ?? '').trim();
    if (!item) return null;
    return {
      date: normDate(get('date')),
      outlet: String(get('outlet') ?? '').trim() || '-',
      supplier: String(get('supplier') ?? '').trim() || '-',
      category: String(get('category') ?? '').trim() || '-',
      item,
      qty: toNum(get('qty')),
      unit: String(get('unit') ?? '').trim() || '',
      amount: toNum(get('amount'))
    };
  }
});

// Error handler กลาง
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    return res.status(400).json({ success: false, message: 'ไฟล์มีขนาดใหญ่เกินที่ระบบรองรับ (จำกัด 25 MB)' });
  }
  console.error(err);
  res.status(500).json({ success: false, message: err.message });
});

// เตรียมไฟล์ Excel ตั้งแต่สตาร์ทเซิร์ฟเวอร์ (รวมถึงย้ายโครงสร้างชีตเก่าเป็นชุดฟิลด์ใหม่ถ้าพบ)
(async () => {
  try {
    const wb = await initWorkbook();
    if (sheetLayoutMigrated) {
      backupExcelFile();
      await saveWorkbook(wb);
      console.log('✅ ย้ายข้อมูลไปโครงสร้างชีตใหม่เรียบร้อย (Direct Purchase Food/Bev, Store Issue Wine/Bev, คอลัมน์ Wine/Other ใน Outlet_Summary)');
    }
  } catch (err) {
    console.error('❌ เตรียมไฟล์ Excel ตอนสตาร์ทไม่สำเร็จ:', err.message);
  }
})();

app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Server is running at http://0.0.0.0:${PORT}`);
});
