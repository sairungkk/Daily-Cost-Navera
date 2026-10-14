// สคริปต์ตรวจสอบหลังแก้โค้ด: leftover patterns + syntax
const fs = require('fs');
const vm = require('vm');

const root = 'C:/Users/DELL/excel-node-app/excel-node-app';
const html = fs.readFileSync(root + '/public/index.html', 'utf8');

const patterns = [
  'toISOString',
  'PASSWORDS',
  'admin001405',
  'Password:1234',
  'await fetch(',
  'colspan="16"',
  'event.currentTarget',
  "querySelectorAll('.outlet-pill')",
  "querySelectorAll('#tab-shift .outlet-pill')"
];

let problems = 0;
patterns.forEach(p => {
  const count = html.split(p).length - 1;
  const bad = p === 'await fetch(' ? count !== 2   // เหลือเฉพาะใน apiFetch + logout
    : count !== 0;
  if (bad) problems++;
  console.log((bad ? 'LEFTOVER!! ' : 'clean      ') + JSON.stringify(p) + ' -> ' + count);
});

// apiFetch ต้องถูกเรียก 13 จุด (fetch เดิมทั้งหมด) และประาาม 1 จุด
const apiCalls = html.split('await apiFetch(').length - 1;
console.log('apiFetch call sites ->', apiCalls, apiCalls === 13 ? 'OK' : 'CHECK!');

const goodPatterns = [
  ['async function apiFetch', 1],
  ['function localDateStr', 1],
  ['localTodayStr()', 6],          // def + 6 จุดเรียก (revDate init, 3 Today btns, save-reset, +1 ใน def)
  ['escapeHtml(', 9],
  ['requireAuth(', 0],             // server side ไม่อยู่ในไฟล์นี้
  ['colspan="17"', 1],
  ['exportExecutivePDF(this)', 1],
  ['#costModeToggle', 1],          // selector ใน JS (id ใน HTML ไม่มี #)
  ['#outletPillsContainer .outlet-pill', 1]
];
goodPatterns.forEach(([p, expect]) => {
  const count = html.split(p).length - 1;
  console.log('pattern ' + JSON.stringify(p) + ' -> ' + count + ' (expect ' + expect + ')');
});

const m = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)];
m.forEach((b, i) => {
  try {
    new vm.Script(b[1]);
    console.log('inline script ' + (i + 1) + ' syntax OK');
  } catch (e) {
    problems++;
    console.log('inline script ' + (i + 1) + ' SYNTAX ERROR: ' + e.message);
  }
});

console.log(problems === 0 ? '\nALL HTML CHECKS PASSED' : '\nHTML CHECKS: ' + problems + ' PROBLEM(S)');
