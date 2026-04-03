# ระบบจัดตารางสอน — โรงเรียนดาราวิทยาลัย v3

## ความต้องการของระบบ

- **Node.js** เวอร์ชัน 18 ขึ้นไป (แนะนำ 20+)
- **npm** (มากับ Node.js)

### ตรวจสอบว่ามี Node.js หรือยัง
```bash
node --version
npm --version
```
ถ้ายังไม่มี ดาวน์โหลดที่ https://nodejs.org/

---

## วิธีติดตั้งและรัน

### 1. แตกไฟล์โปรเจกต์
แตก ZIP แล้ววางไว้ที่ไหนก็ได้ เช่น `C:\dara-timetable` หรือ `Desktop/dara-timetable`

### 2. เปิด Terminal / Command Prompt
```bash
cd dara-timetable
```

### 3. ติดตั้ง Dependencies
```bash
npm install
```
รอประมาณ 30 วินาที - 1 นาที

### 4. รันโปรแกรม
```bash
npm run dev
```

### 5. เปิด Browser
เข้าที่ **http://localhost:3000**

---

## ใช้งานบน Network ของโรงเรียน

เมื่อรัน `npm run dev` จะแสดง URL 2 บรรทัด:

```
  ➜  Local:   http://localhost:3000/
  ➜  Network: http://192.168.x.x:3000/
```

เครื่องอื่นในวง Network เดียวกันสามารถเข้าผ่าน **http://192.168.x.x:3000** ได้เลย
(แทน x.x ด้วย IP ของเครื่องที่รัน)

---

## Build สำหรับ Production

```bash
npm run build
```

จะได้โฟลเดอร์ `dist/` ที่สามารถนำไปวางบน Web Server ใดก็ได้

### รัน Production Preview
```bash
npm run preview
```

---

## ฟีเจอร์หลัก

| ฟีเจอร์ | รายละเอียด |
|---------|-----------|
| ระดับชั้น / ห้องเรียน | สร้าง แก้ไข ลบ ระดับชั้น + ห้องเรียน |
| แผนการเรียน | สร้างแผนแยกต่างหาก ใช้ร่วมข้ามระดับได้ |
| กลุ่มสาระ | สร้าง แก้ไข ลบ + สีแยกอัตโนมัติ |
| จัดการครู | เพิ่ม/แก้ไข/ลบ + Import/Export Excel + คาบที่ได้รับ |
| จัดการวิชา | เพิ่ม/แก้ไข/ลบ + ระบุระดับชั้น + Import/Export Excel |
| มอบหมายงาน | เลือกวิชา→แสดงห้องเฉพาะระดับนั้น + countdown คาบเหลือ |
| คาบล็อค/ประชุม | กำหนดคาบที่กลุ่มสาระประชุม |
| จัดตารางสอน | Drag & Drop + ล็อคคาบ + ป้องกันข้ามระดับ + จำกัดคาบ/ห้อง |
| รายงาน/Export | Export Excel ทุกห้อง/ทุกคน/รายงานสถานะ |

## โครงสร้างไฟล์

```
dara-timetable/
├── index.html          # หน้า HTML หลัก
├── package.json        # Dependencies
├── vite.config.js      # Vite config
├── README.md           # ไฟล์นี้
└── src/
    ├── main.jsx        # Entry point
    └── App.jsx         # แอปทั้งหมด (~730 บรรทัด)
```

## ข้อมูลถูกเก็บที่ไหน?

ข้อมูลทั้งหมดเก็บใน **localStorage** ของ Browser
- ปิด Browser แล้วเปิดใหม่ข้อมูลยังอยู่
- ถ้าต้องการรีเซ็ต ให้เปิด DevTools (F12) → Console → พิมพ์ `localStorage.clear()` แล้ว refresh

---

## ปัญหาที่พบบ่อย

| ปัญหา | วิธีแก้ |
|-------|--------|
| `npm: command not found` | ติดตั้ง Node.js จาก nodejs.org |
| Port 3000 ถูกใช้แล้ว | แก้ port ใน vite.config.js |
| เครื่องอื่นเข้าไม่ได้ | ตรวจ Firewall ให้เปิด port 3000 |
| ตัวอักษรภาษาไทยเพี้ยน | ตรวจว่าเปิดไฟล์เป็น UTF-8 |
