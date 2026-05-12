import os
import shutil
from fastapi import UploadFile
from app.config import settings

def get_user_dir(user_id: int) -> str:
    dir = os.path.join(settings.UPLOAD_DIR, str(user_id))
    os.makedirs(dir, exist_ok=True)
    return dir

async def save_upload_file(upload_file: UploadFile, user_id: int) -> str:
    user_dir = get_user_dir(user_id)
    file_path = os.path.join(user_dir, upload_file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(upload_file.file, buffer)
    return file_path

def delete_file(file_path: str):
    if os.path.exists(file_path):
        os.remove(file_path)