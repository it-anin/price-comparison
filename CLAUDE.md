# CLAUDE.md — Price Comparison Module

## โปรเจกต์นี้คืออะไร
React + Vite app สำหรับเปรียบเทียบราคาทุนสินค้า ระหว่าง:
- **ไฟล์ 1**: Product Master (ทุนใหม่)
- **ไฟล์ 2**: Promaxx (ทุนเก่า + ราคาทุกระดับ)

ผู้ใช้อัปโหลด 2 ไฟล์ → ระบบคำนวณราคาใหม่ทุกหน่วย → export เป็น TXT หรือ XLSX

---

## Column Mapping (อย่าเปลี่ยนโดยไม่ตรวจสอบข้อมูลจริงก่อน)

### ไฟล์ 1 — Product Master (`PM` constant)
| Field | Column | หมายเหตุ |
|-------|--------|---------|
| SKU | A | รหัสสินค้า |
| NAME | B | ชื่อสินค้า |
| BASE_UNIT | C | หน่วยฐาน — **ตรงกับ ColE ไฟล์ 2 เสมอ** |
| LEGACY_UNIT | D | หน่วยสำรอง (fallback) |
| UNIT | H | หน่วยขาย |
| COST | J | ทุนใหม่ (fallback: ColC) |

### ไฟล์ 2 — Promaxx (`PR` constant)
| Field | Column | หมายเหตุ |
|-------|--------|---------|
| SKU | B | รหัสสินค้า |
| UNIT | E | หน่วย — **ใช้จับคู่กับ BASE_UNIT (ColC ไฟล์ 1)** |
| FILTERORLEVEL | F | ระดับราคา: "0","1","4","5" |
| LV0_PRICE | G | ราคาระดับ 0 (เก่า) |
| COL_H | H | ราคาระดับ 1/4/5 (เก่า) และ ทุนเก่า |

---

## Logic หลัก (runCompare)

### การหา baseOldCost
```js
// ใช้ BASE_UNIT (ColC ไฟล์ 1) จับคู่กับ ColE ไฟล์ 2
const basePromaxx = skuUnitMap?.get(baseUnit) ?? {};
const baseOldCost = basePromaxx.lv4 ?? null;
```
> ⚠️ เดิมใช้ `file1Unit` (ColH/ColD) — ถูกเปลี่ยนเป็น `baseUnit` (ColC) แล้ว เพราะข้อมูลจริง ColC ตรงกับ ColE เสมอ

### การคำนวณราคาใหม่ต่อหน่วย
```
newCostUnit = (file1Cost / baseOldCost) × oldCostUnit
lv0_new = round((newCostUnit / oldCostUnit) × lv0_old)
lv1_new = round((newCostUnit / oldCostUnit) × lv1_old)
lv5_new = round((newCostUnit / oldCostUnit) × lv5_old)
```

### ColF กำหนดระดับราคา
- `"0"` → lv0_old ← ColG
- `"1"` → lv1_old ← ColH
- `"4"` → lv4 (ทุนเก่า) ← ColH
- `"5"` → lv5_old ← ColH

---

## Export Functions

| ฟังก์ชัน | ปุ่ม | เงื่อนไข | output |
|---------|------|---------|--------|
| `exportTextByLevel(level)` | สีน้ำเงิน | ทุนใหม่ > ทุนเก่า | `SKU\tราคา\tLevel.txt` |
| `exportBaseUnit(level)` | สีเขียว | ทุนใหม่ > ทุนเก่า **และ** ColC ไฟล์ 1 = ColE ไฟล์ 2 | `BASE_UNIT_ราคาระดับ X.txt` |
| `exportTextByUnitAndLevel(unit, level)` | ปุ่มแยกหน่วย | filter หน่วย + ทุนใหม่ > ทุนเก่า | `unit_ราคาระดับ X.txt` |
| `exportExcel()` | Export All | ทุกแถวใน displayRows | `Price_Comparison_Report.xlsx` (1 sheet) |

---

## ปัญหาที่เจอ + วิธีแก้

### 1. การ inject โค้ดผ่าน Node.js heredoc บน Windows
**ปัญหา:** เมื่อใช้ `node << 'NODEJS'` เพื่อ inject โค้ด:
- `'\r\n'` กลายเป็น literal CRLF ในไฟล์ → `Unterminated string literal` error
- `'\uFEFF'` กลายเป็น BOM character จริงๆ แทนที่จะเป็น escape sequence
- `'\\'` (backslash) ถูก shell ตัดเหลือ `'\'` → SyntaxError

**วิธีแก้:** ใช้ `String.fromCharCode()` แทน escape sequences ทุกตัว:
```js
const BS = String.fromCharCode(92);    // backslash
const BOM = String.fromCharCode(0xFEFF); // BOM
const CR = String.fromCharCode(13);
const LF = String.fromCharCode(10);
```

### 2. ภาษาไทยเพี้ยนใน Excel headers
**ปัญหา:** ใช้ string literal ไทยใน heredoc แล้ว encoding เสียหาย เช่น `"เธฃเธซเธฑเธชเธชเธดเธเธเนเธฒ"` แทน `"รหัสสินค้า"`

**วิธีแก้:** เขียน Unicode escape หรือใช้ `\uXXXX` form แทน เมื่อต้องการ inject ผ่าน Node.js script

### 3. BASE_UNIT section วางผิด nesting level
**ปัญหา:** เพิ่ม section เข้าไปอยู่ใน `{ternary ? ... : ...}` ทำให้ไม่ render

**วิธีแก้:** ตรวจ indent และ closing tags ให้ section อยู่ **นอก** ternary แต่ **ใน** parent div เสมอ

---

## โครงสร้าง UI Export Tab
```
Export Options Panel
├── ปุ่มสีน้ำเงิน: ราคาระดับ 0/1/4/5 (ทุกแถว)
├── BASE_UNIT Export (สีเขียว): ระดับ 0/1/4/5
├── Export All (.xlsx)
└── Export แยกตามหน่วย + ระดับราคา
    └── [unit] ระดับ 0 / 1 / 4 / 5
```

---

## Git / GitHub

- Remote: `https://github.com/it-anin/price-comparison`
- Branch: `main`
- Push ครั้งแรกต้องใช้ PAT ของ account `it-anin` (account `bigyasrcit-rgb` ไม่มีสิทธิ์)

```bash
git remote set-url origin https://it-anin:<TOKEN>@github.com/it-anin/price-comparison.git
git push -u origin main
```

---

## ข้อควรระวัง

1. **อย่าใช้ `Edit` tool กับ string ที่มี Thai Unicode** ถ้า match ไม่ได้ ให้ใช้ `node` script แทน พร้อม `String.fromCharCode` สำหรับ special chars
2. **`LEGACY_COST` และ `BASE_UNIT` ชี้ไปที่ ColC เหมือนกัน** — ตั้งใจแล้ว ไม่ใช่ bug
3. **`file1Unit` (ColH/ColD) ยังใช้อยู่** สำหรับ fallback กรณีไม่พบ units ใน file 2 เท่านั้น
4. **Excel export ใช้ `bookType: undefined`** (default) — ถ้าต้องการเปลี่ยน encoding ต้องระวัง BOM
5. **ข้อมูลผ่าน client-side เท่านั้น** ไม่มี server ไม่มี API ทุกอย่างอยู่ใน browser
