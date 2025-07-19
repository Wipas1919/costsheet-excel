# Excel Template Generator

โปรแกรมสร้าง Excel template สำหรับ Cost Sheet พร้อมระบบอัปโหลดไฟล์ไปยัง MinIO

## คุณสมบัติ

- สร้าง Excel template แบบ fixed position
- รองรับการเพิ่มข้อมูลแบบ dynamic
- อัปโหลดไฟล์ไปยัง MinIO
- สร้าง Presigned URL สำหรับดาวน์โหลด

## การติดตั้ง

```bash
pip install -r requirements.txt
```

## การใช้งาน

### 1. สร้าง Template เปล่า

```python
from excel import main

# เรียกใช้ฟังก์ชันหลัก
result = main(
    minio_access_key="your_access_key",
    minio_secret_key="your_secret_key", 
    bucket_name="your_bucket"
)
print(result["file_url"])
```

### 2. สร้าง Template พร้อมข้อมูล

```python
from excel import main

# ข้อมูลแบบ dynamic
dynamic_data = {
    "A3": "2024-01-15",           # วันที่
    "A4": "Project ABC Exhibition", # ชื่อโปรเจค
    "D4": "500,000",              # งบประมาณ
    "A5": "Booth Design",         # รายการที่ 1
    "D5": "100,000",              # ราคาที่ 1
    "A6": "Construction",         # รายการที่ 2
    "D6": "200,000",              # ราคาที่ 2
    "A7": "Lighting",             # รายการที่ 3
    "D7": "50,000",               # ราคาที่ 3
    "A8": "Total",                # รวม
    "D8": "350,000"               # ราคารวม
}

result = main(
    minio_access_key="your_access_key",
    minio_secret_key="your_secret_key",
    bucket_name="your_bucket",
    dynamic_data=dynamic_data
)
print(result["file_url"])
```

## โครงสร้าง Template

Template จะมีโครงสร้างดังนี้:

- **A1:K1**: หัวข้อหลัก "Cost Sheet" (ขนาด 26pt, Bold)
- **A2:C2**: "Exhibition Booth" (ขนาด 12pt)
- **A3**: "date" (มี border)
- **B3:C3**: ช่องว่างสำหรับวันที่ (merged)
- **D3:F3**: "Booth" (merged)
- **G3:K3**: ช่องว่างสำหรับข้อมูล booth (merged)
- **A4**: "Project name" (มี border)
- **B4:C4**: ช่องว่างสำหรับชื่อโปรเจค (merged)
- **D4:F4**: "Budget" (merged)
- **G4:K4**: ช่องว่างสำหรับงบประมาณ (merged)
- **A5-K20**: แถวสำหรับข้อมูลรายการต่างๆ

## การกำหนดค่า MinIO

โปรแกรมใช้ endpoint 2 แบบ:
- **Internal**: `http://minio:9000` (สำหรับการอัปโหลดภายใน Docker network)
- **Public**: `http://100.81.135.35:9009` (สำหรับสร้าง Presigned URL)

## หมายเหตุ

- ไฟล์จะถูกบันทึกในรูปแบบ `CostSheet-YYYYMMDD-HHMMSS.xlsx`
- Presigned URL มีอายุ 1 ชั่วโมง
- Template รองรับข้อมูลสูงสุด 20 แถว (A5-K20) 