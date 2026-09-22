const fs = require('fs');
const path = require('path');
const root = 'C:/Users/DELL/excel-node-app/excel-node-app';

// 1) หาสำเนา database.json ที่เหลืออยู่ในโปรเจกต์ (รวมโฟลเดอร์ v1)
let copies = [];
const walk = (dir) => {
  for (const f of fs.readdirSync(dir)) {
    const p = path.join(dir, f);
    const st = fs.statSync(p);
    if (st.isDirectory()) {
      if (f !== 'node_modules' && f !== 'backups') walk(p);
    } else if (f === 'database.json') {
      copies.push(`${p} (${st.size} bytes)`);
    }
  }
};
walk(root);
console.log('database.json copies found:', copies.length ? copies.join(' | ') : 'NONE');

// 2) มีโค้ดอ้างถึง database.json หรือไม่
for (const f of ['server.js', 'public/index.html', 'package.json']) {
  const t = fs.readFileSync(path.join(root, f), 'utf8');
  console.log(`${f} references database.json:`, t.includes('database.json'));
}

// 3) dependency ที่ถอดออก: ยืนยันโค้ดไม่ได้ require
const sv = fs.readFileSync(path.join(root, 'server.js'), 'utf8');
console.log("server.js requires 'cors':", sv.includes("require('cors')"));
console.log("server.js requires 'xlsx':", sv.includes("require('xlsx')"));
