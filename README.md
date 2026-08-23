# 🎓 ระบบเช็คชื่อนักศึกษาออนไลน์ - Anti-Cheat & Mobile-First Edition

ระบบเช็คชื่อนักศึกษาออนไลน์ระดับองค์กร พร้อมมาตรการป้องกันการโกง 4 ชั้น (Dynamic TOTP 15s, GPS Geofencing, Device Token Binding, Live Selfie) ใช้งานร่วมกับ Google Sheets หรือเปิดใช้งานแบบ Standalone PWA

---

## 🌟 จุดเด่นและฟังก์ชันการทำงาน

### 🛡️ 1. มาตรการป้องกันการโกง 4 ชั้น (Anti-Cheat Engine)
1. **Dynamic TOTP / Live QR Code**: รหัส OTP 6 หลักและ QR Code รีเฟรชใหม่ทุก 15 วินาที สำหรับฉายโปรเจกเตอร์หน้าห้อง (ป้องกันการแคปหน้าจอส่งต่อข้ามห้อง)
2. **GPS Geofencing**: ตรวจสอบพิกัด GPS ของนักศึกษาว่าอยู่ภายในรัศมีห้องเรียนจริงหรือไม่ (Haversine Formula)
3. **Device Token Binding (1 เครื่อง / 1 บัญชี)**: ผูก Session กับอุปกรณ์ ป้องกันการเช็คชื่อแทนกันในเครื่องเดียว
4. **Live Selfie Camera**: บังคับถ่ายภาพสดผ่านกล้องหน้าเท่านั้น ปิดการเลือกรูปจากคลังภาพ

### 📱 2. ออกแบบสำหรับสมาร์ทโฟน (Mobile-First UX/UI)
- **Step Progress Bar Tracker**: แถบนำทาง 4 ขั้นตอนสวยงาม (1. ตัวตน ➔ 2. GPS ➔ 3. เซลฟี่ ➔ 4. OTP)
- **Responsive Number Pad OTP**: ช่องกรอกรหัส 6 หลักปรับขนาดพอดีกับหน้าจอมือถือ พร้อมเปิดคีย์บอร์ดตัวเลขทันที
- **Mobile Bottom Navigation Dock**: แถบเมนูด้านล่าง 5 แท็บสำหรับอาจารย์ ใช้งานมือเดียวได้สะดวก

### 📥 3. ระบบนำเข้ารายชื่อนักเรียนขั้นสูง (Excel & Copy-Paste)
- **เลือกห้องเรียนก่อนนำเข้า**: กำหนดห้องเรียนปลายทาง หรือแยกห้องตามไฟล์ Excel (พร้อมปุ่ม + เพิ่มห้องใหม่ ในตัว)
- **2 โหมดการนำเข้า**:
  - 📁 ไฟล์ Excel (.xlsx / .csv): ลากวางไฟล์
  - 📋 คัดลอกและวางข้อความ (Copy & Paste): วางตารางจาก Excel, Word, Google Sheets, Line หรือ Notepad
- **Interactive Preview Table (แก้ไขก่อนบันทึก)**:
  - พิมพ์แก้ไข เลขที่, รหัสนักศึกษา, ชื่อ-นามสกุล, ห้องเรียน ได้โดยตรงในตาราง
  - สลับลำดับแถวขึ้น-ลงได้อิสระด้วยปุ่ม ⬆️ ⬇️
  - ลบแถวที่ไม่ต้องการออกได้ด้วยปุ่ม 🗑️
  - เพิ่มแถวใหม่ได้ด้วยปุ่ม + เพิ่มแถว
  - จัดเรียงเลขที่ 1..N ใหม่อัตโนมัติด้วยปุ่ม จัดเลขที่อัตโนมัติ

### 👨‍🏫 4. ระบบอาจารย์ (Admin Portal - รหัสผ่าน: dmin888)
- **จัดการห้องเรียน**: เพิ่ม/ลบห้องเรียนได้อย่างอิสระ
- **ปุ่มลบรายชื่อทั้งหมด**: พร้อมหน้าต่างยืนยันการลบสีแดงนีออน (Luxury Confirmation Popup)
- **จอโปรเจกเตอร์เต็มจอ (HTML5 Fullscreen API)**: ขยายเต็มหน้าจอ 100% พร้อม Live Feed ตรวจสอบแบบ Real-time
- **รายงานและสถิติ**: Dashboard กราฟสรุปสถิติ, จัดอันดับการมาเรียน, ส่งออกรายงาน **Excel (.xlsx)** และ **PDF Report**

---

## 📁 โครงสร้างไฟล์

`	ext
├── Code.gs        # Backend Google Apps Script (ฐานข้อมูล Google Sheets & Security Engine)
├── Index.html     # Frontend Single-File (HTML + Glassmorphism CSS + JavaScript + Libs)
└── README.md      # คู่มือการใช้งานและเอกสารประกอบโปรเจกต์
`

---

## 🚀 วิธีการนำไปติดตั้งใช้งานบน Google Apps Script (ฐานข้อมูล Google Sheets)

1. เปิด **[Google Sheets](https://sheets.new)** สร้างสเปรดชีตใหม่
2. ไปที่เมนู **ส่วนขยาย (Extensions)** ➔ **Apps Script**
3. สร้างไฟล์ 2 ไฟล์ในโปรเจกต์:
   - ไฟล์ Code.gs: คัดลอกโค้ดจาก Code.gs ไปวาง
   - ไฟล์ Index.html (สร้างแบบไฟล์ HTML): คัดลอกโค้ดจาก Index.html ไปวาง
4. กดบันทึก (Save Project)
5. กดปุ่ม **การทำให้ใช้งานได้ (Deploy)** ➔ **การทำให้ใช้งานได้รายการใหม่ (New deployment)**
6. เลือกประเภทเป็น **เว็บแอป (Web app)**:
   - **คำอธิบาย (Description)**: ระบบเช็คชื่อนักศึกษา
   - **ดำเนินการในฐานะ (Execute as)**: ฉัน (Me)
   - **ผู้มีสิทธิ์เข้าถึง (Who has access)**: ทุกคน (Anyone)
7. กด **ทำให้ใช้งานได้ (Deploy)** และให้สิทธิ์การเข้าถึง (Authorize Access)
8. คัดลอก URL ของเว็บแอปไปเปิดใช้งานบนคอมพิวเตอร์และสมาร์ทโฟนได้ทันที

---

## 🌐 วิธีการนำขึ้น GitHub และเปิดใช้งาน GitHub Pages

### วิธีที่ 1: อัปโหลดผ่านหน้าเว็บ GitHub (ง่ายที่สุด ไม่ต้องลงโปรแกรม)
1. เข้าไปที่ **[GitHub.com](https://github.com)** แล้วสร้าง Repository ใหม่ (เช่น ตั้งชื่อ student-attendance-system)
2. ในหน้า Repository กดปุ่ม **Add file** ➔ **Upload files**
3. ลากไฟล์ Index.html, Code.gs, และ README.md ไปวาง แล้วกด **Commit changes**

### วิธีที่ 2: เปิดหน้าเว็บทดลองใช้งาน (Demo) ด้วย GitHub Pages
1. ในหน้า GitHub Repository ไปที่แท็บ **Settings** ➔ เมนูด้านซ้ายเลือก **Pages**
2. ในส่วน **Build and deployment** > **Branch**:
   - เลือก Branch: main (หรือ master)
   - เลือก Folder: / (root)
3. กด **Save**
4. รอประมาณ 1–2 นาที จะได้รับลิงก์ GitHub Pages (เช่น https://username.github.io/student-attendance-system/Index.html)
5. สามารถเปิดทดลองเล่นระบบบนสมาร์ทโฟนได้ทันที (ระบบมีโหมด Offline LocalStorage รองรับการทดสอบ)

---

## 🔑 ข้อมูลเข้าสู่ระบบ Admin

- **รหัสผ่านเข้าสู่ระบบอาจารย์**: dmin888