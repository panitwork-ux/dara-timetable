# ติดตั้งและอัปเดต

1. สำรองข้อมูล JSON จากระบบก่อนอัปเดต
2. ตั้งค่า Firebase Web App ใน `.env.local` ตาม `.env.example`
3. รัน `npm ci` แล้ว `npm run build:live`
4. เผยแพร่เฉพาะไฟล์ใน `dist` ไปยังสาขา `gh-pages` ส่วนซอร์สอยู่ใน `main`

เปิดทดสอบในเครื่องด้วย `npm run dev` และใช้ http://localhost:5173/

การอัปเดตเว็บไม่เปลี่ยนข้อมูล Firestore และไม่เผยแพร่กฎ Firebase โดยอัตโนมัติ ห้ามอัปโหลด `.env.local` หรือกุญแจ Admin SDK ลง Git
