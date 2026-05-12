# app/routes/students.py
from fastapi import APIRouter, Depends, UploadFile, File, HTTPException
from sqlalchemy.orm import Session
from app.database import get_db
from app.dependencies import get_current_user
from app.models.user import User
from app.services.filter_ai_service import process_filter_ai   # không dùng, ta cần một service khác
from app.utils.legacy.excel_filter_ai import ExcelFilterApp
import json

router = APIRouter()

# Tạo instance ExcelFilterApp cho mỗi user? Ta có thể dùng class variable hay tạo mới mỗi lần.
# Ở đây tạo một dict toàn cục (tạm thời) hoặc lưu mapping vào DB.
# Giải pháp đơn giản: dùng cùng một instance cho tất cả user, nhưng cần phân biệt user_id.
# Tốt nhất: lưu mapping vào DB thay vì file JSON. Tuy nhiên để tương thích với code legacy, ta vẫn dùng file JSON riêng cho từng user.
# Ta sẽ tạo một class quản lý student list cho từng user, lưu trong thư mục người dùng.

def get_student_manager(user_id: int):
    # Tạo instance ExcelFilterApp riêng cho user, với thư mục app_data riêng
    import os
    from app.utils.file_utils import get_user_dir
    user_dir = get_user_dir(user_id)
    app_data_dir = os.path.join(user_dir, "app_data")
    os.makedirs(app_data_dir, exist_ok=True)
    app = ExcelFilterApp()
    app.app_data_dir = app_data_dir
    app.student_list_path = os.path.join(app_data_dir, "student_list.json")
    app._load_student_list_from_file()
    return app

@router.get("/")
def get_students(current_user: User = Depends(get_current_user)):
    app = get_student_manager(current_user.id)
    return app.student_mapping

@router.post("/upload-excel")
async def upload_student_excel(file: UploadFile = File(...), current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    # Lưu file tạm
    from app.utils.file_utils import save_upload_file, get_user_dir
    import os
    user_dir = get_user_dir(current_user.id)
    file_path = os.path.join(user_dir, file.filename)
    with open(file_path, "wb") as buffer:
        import shutil
        shutil.copyfileobj(file.file, buffer)
    
    app = get_student_manager(current_user.id)
    success, msg = app.load_student_list_from_excel(file_path)
    os.remove(file_path)  # xóa file tạm
    if not success:
        raise HTTPException(400, msg)
    return {"detail": msg}

@router.post("/")
def add_student(email: str, name: str, current_user: User = Depends(get_current_user)):
    app = get_student_manager(current_user.id)
    success, msg = app.add_student_manually(email, name)
    if not success:
        raise HTTPException(400, msg)
    return {"detail": msg}

@router.delete("/{email}")
def delete_student(email: str, current_user: User = Depends(get_current_user)):
    app = get_student_manager(current_user.id)
    success, msg = app.remove_student(email)
    if not success:
        raise HTTPException(404, msg)
    return {"detail": msg}

@router.delete("/")
def clear_all_students(current_user: User = Depends(get_current_user)):
    app = get_student_manager(current_user.id)
    success, msg = app.clear_student_list()
    if not success:
        raise HTTPException(400, msg)
    return {"detail": msg}