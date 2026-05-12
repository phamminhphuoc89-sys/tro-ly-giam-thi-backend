from fastapi import APIRouter, Depends, UploadFile, File
from sqlalchemy.orm import Session
from app.database import get_db
from app.models.file import UserFile
from app.schemas.file import FileOut
from app.dependencies import get_current_user
from app.utils.file_utils import save_upload_file

router = APIRouter()

@router.post("/upload", response_model=FileOut)
async def upload_file(
    file: UploadFile = File(...),
    file_type: str = "input_excel",   # có thể cho phép chọn
    db: Session = Depends(get_db),
    current_user = Depends(get_current_user)
):
    # Lưu file
    file_path = await save_upload_file(file, current_user.id)
    db_file = UserFile(
        owner_id=current_user.id,
        original_name=file.filename,
        stored_path=file_path,
        file_type=file_type
    )
    db.add(db_file)
    db.commit()
    db.refresh(db_file)
    return db_file

@router.get("/{file_id}", response_model=FileOut)
def get_file(file_id: int, db: Session = Depends(get_db), current_user = Depends(get_current_user)):
    file = db.query(UserFile).filter(UserFile.id == file_id, UserFile.owner_id == current_user.id).first()
    if not file:
        raise HTTPException(404, "File không tồn tại")
    return file