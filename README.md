# แบบสำรวจสภาพอาคาร — สถานะปัจจุบัน

**Backend ที่ใช้งานจริงตอนนี้คือ Google Apps Script** (`apps_script/Code.gs`)
ไม่ใช่ Python อีกต่อไป — ดูวิธี Deploy ในคอมเมนต์ด้านบนของไฟล์ `apps_script/Code.gs`

## ไฟล์ Python เดิม (`app.py`, `report_generator.py`, `drive_uploader.py`,
`sheets_uploader.py`, `requirements.txt`, `.env.example`) — **ไม่ได้ใช้งานแล้ว**
เก็บไว้เป็นข้อมูลอ้างอิง/เผื่ออยากย้อนกลับไปใช้ Python ในอนาคต ไม่ต้องรัน `uvicorn`
อีกต่อไปสำหรับการใช้งานจริง

## โครงสร้างที่ใช้งานจริงตอนนี้

```
templates/index.html   ← หน้าเว็บทั้งหมด (เหมือนเดิม) เรียก Apps Script แทน Python
apps_script/Code.gs    ← Backend ตัวจริง (วางใน script.google.com)
```

`index.html` เรียก 2 การทำงานผ่าน `APPS_SCRIPT_URL` (ตัวแปรบนสุดของ `<script>`):
- `action: "checkBranch"` — เช็คสาขาซ้ำ
- `action: "submit"` — ส่งแบบสำรวจ (เขียน Sheet + อัปโหลด Drive + สร้าง PDF)

## ข้อจำกัดเทียบกับฉบับ Python เดิม

- **ไม่มีการย่อ/บีบอัดรูปก่อนอัปโหลด** — Apps Script ไม่มีเครื่องมือแบบ Pillow ในตัว
  รูปขึ้น Drive ตามขนาดจริงที่ถ่ายมา ต้องเผื่อพื้นที่ Drive ให้พอ
- **Layout ของ PDF ต่างจากเดิมเล็กน้อย** — ส่วนรูปภาพเรียงทีละแถว ขึ้นหน้าใหม่ทุก 8 รูป
  แทนตาราง 2 คอลัมน์ × 4 แถว แบบเป๊ะๆ เหมือนฉบับ Python (reportlab ควบคุม layout ได้ละเอียดกว่า)
