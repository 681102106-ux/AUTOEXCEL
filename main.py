from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
from google import genai
from google.genai import types
import json
import os
import time
import pymupdf
import asyncio
import random # 🌟 1. นำเข้าไลบรารีสุ่มตัวเลขสำหรับ Jitter
from fastapi.staticfiles import StaticFiles
from dotenv import load_dotenv
import csv
from contextlib import asynccontextmanager 

load_dotenv()

SAVE_DIR = os.path.join(os.getcwd(), "saved_bills")
os.makedirs(SAVE_DIR, exist_ok=True)

VENDOR_MASTER_FILE = "vendor_master.csv"
vendor_db = {} 
vendor_list_for_ui = [] 

raw_keys = os.getenv("GEMINI_API_KEYS", "")
# 🌟 2. ระบบ Circuit Breaker: แยกกุญแจทั้งหมด กับกุญแจที่ยัง "รอดชีวิต"
ALL_API_KEYS = [k.strip() for k in raw_keys.split(",") if k.strip()]
ACTIVE_API_KEYS = ALL_API_KEYS.copy()
current_key_idx = 0

if not ALL_API_KEYS: print("⚠️ แจ้งเตือน: ไม่พบ API Key ในไฟล์ .env")

invoice_schema = types.Schema(
    type=types.Type.ARRAY,
    items=types.Schema(
        type=types.Type.OBJECT,
        properties={
            "is_valid_invoice": types.Schema(type=types.Type.BOOLEAN, description="รูปภาพนี้คือใบกำกับภาษีใช่หรือไม่? ตอบ false หากเป็นเอกสารอื่น"),
            "page_number": types.Schema(type=types.Type.INTEGER, description="หน้าที่พบใบกำกับภาษีนี้ (หน้าแรกคือ 1)"),
            "date": types.Schema(type=types.Type.STRING, description="วันที่ DD/MM/YYYY แปลงเป็น พ.ศ."),
            "invoice_no": types.Schema(type=types.Type.STRING),
            "vendor_name": types.Schema(type=types.Type.STRING, description="ตัดคำว่า สาขา หรือ สำนักงานใหญ่ ออก"),
            "tax_id": types.Schema(type=types.Type.STRING, description="เลขผู้เสียภาษีของผู้ขาย"),
            "buyer_tax_id": types.Schema(type=types.Type.STRING, description="เลขประจำตัวผู้เสียภาษีของผู้ซื้อ (ถ้าไม่มีให้เว้นว่าง)"),
            "is_hq": types.Schema(type=types.Type.BOOLEAN),
            "branch_no": types.Schema(type=types.Type.STRING, description="ถ้าเป็นสำนักงานใหญ่ให้เว้นว่าง"),
            "subtotal": types.Schema(type=types.Type.NUMBER),
            "vat": types.Schema(type=types.Type.NUMBER),
            "total": types.Schema(type=types.Type.NUMBER),
        },
        required=["is_valid_invoice", "page_number", "date", "invoice_no", "vendor_name", "tax_id", "buyer_tax_id", "is_hq", "branch_no", "subtotal", "vat", "total"]
    )
)

def load_vendor_master():
    global vendor_db, vendor_list_for_ui
    vendor_db.clear()
    vendor_list_for_ui.clear()
    
    if os.path.exists(VENDOR_MASTER_FILE):
        with open(VENDOR_MASTER_FILE, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                tax_id = ''.join(filter(str.isdigit, str(row.get('tax_id', ''))))
                v_code = str(row.get('v_code', '')).strip()
                name = str(row.get('vendor_name', '')).strip()
                if tax_id:
                    vendor_db[tax_id] = {"v_code": v_code, "name": name}
                    vendor_list_for_ui.append({"v_code": v_code, "tax_id": tax_id, "name": name})
        print(f"📦 โหลดฐานข้อมูลผู้ขายสำเร็จ: {len(vendor_db)} รายการ")
    else:
        print("⚠️ ไม่พบไฟล์ vendor_master.csv (สร้างฐานข้อมูลว่างเปล่า)")

# 🌟 อัปเกรดเป็นระบบซ่อมบำรุงประจำวัน (ลบไฟล์เก่า + ชุบชีวิตกุญแจ API)
async def daily_maintenance_task():
    global ACTIVE_API_KEYS, ALL_API_KEYS
    while True:
        try:
            # 1. ทำความสะอาดไฟล์รูปที่เก่าเกิน 7 วัน
            now = time.time()
            for filename in os.listdir(SAVE_DIR):
                filepath = os.path.join(SAVE_DIR, filename)
                if os.path.isfile(filepath) and now - os.path.getmtime(filepath) > 604800:
                    os.remove(filepath)
            
            # 2. 🌟 ระบบ Self-Healing: ดึงกุญแจที่โดนแบนกลับมาใหม่ทุกๆ 24 ชม.
            if len(ACTIVE_API_KEYS) < len(ALL_API_KEYS):
                ACTIVE_API_KEYS = ALL_API_KEYS.copy()
                print(f"🔄 [Self-Healing] คืนชีปกุญแจ API ทั้งหมด {len(ACTIVE_API_KEYS)} ดอก สำหรับโควต้าวันใหม่แล้ว!")
                
        except Exception as e: pass
        await asyncio.sleep(86400) # รอ 24 ชั่วโมง (86,400 วินาที) แล้ววนลูปทำใหม่

@asynccontextmanager
async def lifespan(app: FastAPI):
    load_vendor_master() 
    # เปลี่ยนชื่อตัวแปรให้ตรงกับฟังก์ชันใหม่
    maintenance_task = asyncio.create_task(daily_maintenance_task())
    yield
    maintenance_task.cancel()
    
app = FastAPI(lifespan=lifespan)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

app.mount("/static", StaticFiles(directory="static"), name="static")
app.mount("/saved_bills", StaticFiles(directory=SAVE_DIR), name="saved_bills")

@app.get("/get-vendors")
async def get_vendors():
    return {"status": "success", "data": vendor_list_for_ui}

def process_batch_with_ai(file_paths):
    global current_key_idx, ACTIVE_API_KEYS
    prompt = "จงสกัดข้อมูลของ 'ทุกบิลในทุกหน้า' ออกมารวมกันเป็น JSON Array โดยตรวจสอบก่อนว่าเอกสารคือใบกำกับภาษีหรือไม่ และอย่าลืมดึงเลขประจำตัวผู้เสียภาษีของผู้ซื้อ (Buyer Tax ID) ออกมาด้วยเพื่อไว้ตรวจสอบ"
    
    max_attempts = 10 # เผื่อรอบไว้เยอะๆ สำหรับการถอยหลัง Exponential Backoff
    base_delay = 2
    uploaded_files = []
    last_used_key = None
    client = None

    for attempt in range(max_attempts):
        # 🌟 Circuit Breaker เช็กว่ายังมีกุญแจเหลือไหม
        if not ACTIVE_API_KEYS:
            raise Exception("❌ กุญแจ API ทุกดอกพัง หรือ โดนแบน (403) หมดแล้ว! โปรดตรวจสอบ API Keys")

        current_key = ACTIVE_API_KEYS[current_key_idx % len(ACTIVE_API_KEYS)]

        try:
            if current_key != last_used_key:
                if client and uploaded_files:
                    for uf in uploaded_files:
                        try: client.files.delete(name=uf.name)
                        except: pass
                
                uploaded_files = []
                client = genai.Client(api_key=current_key)
                print(f"📤 กำลังอัปโหลด {len(file_paths)} หน้า ด้วยกุญแจดอกที่ {(ALL_API_KEYS.index(current_key) + 1)}...")
                for path in file_paths:
                    uf = client.files.upload(file=path)
                    uploaded_files.append(uf)
                last_used_key = current_key
            
            print(f"🚀 ยิงคำสั่งวิเคราะห์ (รอบที่ {attempt+1})...")
            ai_contents = uploaded_files + [prompt]
            
            response = client.models.generate_content(
                model='gemini-3.6-flash',
                contents=ai_contents,
                config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=invoice_schema, temperature=0.1)
            )
            
            for uf in uploaded_files:
                try: client.files.delete(name=uf.name)
                except: pass
            
            return json.loads(response.text)
            
        except Exception as api_error:
            error_msg = str(api_error)
            
            # เคลียร์ไฟล์ขยะของรอบที่พัง
            if "403" in error_msg or "429" in error_msg:
                for uf in uploaded_files:
                    try: client.files.delete(name=uf.name)
                    except: pass
                last_used_key = None # บังคับให้อัปโหลดใหม่ด้วยคีย์ถัดไป
            
            if "403" in error_msg:
                print(f"🚫 [Circuit Breaker] กุญแจดอกที่ {(ALL_API_KEYS.index(current_key) + 1)} โดนบล็อก (403) เตะออกจากระบบ!")
                ACTIVE_API_KEYS.remove(current_key)
                if ACTIVE_API_KEYS:
                    current_key_idx = current_key_idx % len(ACTIVE_API_KEYS)
                continue # ไม่นับรอบนี้ใน Backoff ให้วนลูปใหม่ด้วยคีย์ถัดไปทันที
                
            elif "429" in error_msg or "503" in error_msg:
                if "429" in error_msg:
                    print(f"⚠️ โควต้าเต็ม (429) 🔄 สลับกุญแจ...")
                    current_key_idx = (current_key_idx + 1) % len(ACTIVE_API_KEYS)
                
                # 🌟 Exponential Backoff + Jitter
                # คำนวณเวลา: 2^0 = 1, 2^1 = 2, 2^2 = 4, 2^3 = 8 ... บวกเลขสุ่ม 0-1 วินาที
                delay = (base_delay ** attempt) + random.uniform(0, 1)
                delay = min(delay, 30) # ล็อกไว้สูงสุดไม่เกิน 30 วินาทีเพื่อไม่ให้รอนานเกินไป
                
                if "503" in error_msg:
                    print(f"⏳ เซิร์ฟเวอร์ไม่ว่าง (503) พัก {delay:.2f} วินาที (ไม่สลับกุญแจ)...")
                else:
                    print(f"⏳ พัก {delay:.2f} วินาที ก่อนลองใหม่...")
                    
                time.sleep(delay)
            else:
                raise api_error
                
    raise Exception("❌ ระบบพยายามเชื่อมต่อหลายครั้งแต่เซิร์ฟเวอร์ขัดข้องเกินไป โปรดลองใหม่ภายหลัง")

def validate_and_clean(raw_data_list, saved_image_paths):
    cleaned_list = []
    
    for data in raw_data_list:
        data['tax_id'] = ''.join(filter(str.isdigit, str(data.get('tax_id', ''))))
        
        tax_id = data['tax_id']
        data['v_code'] = "" 
        
        if tax_id in vendor_db:
            data['vendor_name'] = vendor_db[tax_id]['name']
            data['v_code'] = vendor_db[tax_id]['v_code']
            data['is_vendor_matched'] = True
        else:
            data['is_vendor_matched'] = False
        
        page_idx = data.get('page_number', 1)
        try:
            actual_idx = min(max(1, page_idx), len(saved_image_paths)) - 1
            filename = os.path.basename(saved_image_paths[actual_idx])
            data['file_path'] = f"http://localhost:8080/saved_bills/{filename}"
        except: data['file_path'] = ""
            
        data['buyer_tax_id'] = ''.join(filter(str.isdigit, str(data.get('buyer_tax_id', ''))))
        
        try:
            sub, vat, total = float(data.get('subtotal', 0)), float(data.get('vat', 0)), float(data.get('total', 0))
            data['is_math_valid'] = False if abs((sub + vat) - total) > 0.1 else True
        except: data['is_math_valid'] = False
            
        try:
            parts = str(data.get('date', '')).split('/')
            if len(parts) == 3: data['doc_month'], data['doc_year'] = int(parts[1]), int(parts[2])  
        except: data['doc_month'], data['doc_year'] = 0, 0
        cleaned_list.append(data)
    return cleaned_list

@app.post("/extract-invoice")
async def extract_invoice(file: UploadFile = File(...)):
    _, ext = os.path.splitext(file.filename)
    file_bytes = await file.read()
    batch_id = int(time.time())
    saved_image_paths = [] 
    try:
        if ext.lower() == '.pdf':
            doc = pymupdf.open(stream=file_bytes, filetype="pdf")
            for page_num, page in enumerate(doc):
                pix = page.get_pixmap(dpi=100) 
                # 🌟 3. Payload Reduction: เซฟเป็น .jpg แทน .png เพื่อบีบอัดขนาดไฟล์ก่อนส่งให้ AI
                img_path = os.path.join(SAVE_DIR, f"BILL_{batch_id}_P{page_num+1}.jpg")
                pix.save(img_path)
                saved_image_paths.append(img_path)
        else:
            img_path = os.path.join(SAVE_DIR, f"BILL_{batch_id}{ext}")
            with open(img_path, "wb") as buffer: buffer.write(file_bytes)
            saved_image_paths.append(img_path)

        raw_data = process_batch_with_ai(saved_image_paths)
        if not isinstance(raw_data, list): raw_data = [raw_data]

        final_clean_data = validate_and_clean(raw_data, saved_image_paths)
        return {"status": "success", "data": final_clean_data}
    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8080)