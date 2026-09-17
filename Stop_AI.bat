@echo off
chcp 65001 >nul
echo กำลังปิดระบบ AI และเซิร์ฟเวอร์เบื้องหลัง...
taskkill /F /IM python.exe
echo ปิดระบบเรียบร้อยแล้ว!
timeout /t 2 >nul