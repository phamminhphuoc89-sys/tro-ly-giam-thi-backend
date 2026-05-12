from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from sqlalchemy.orm import Session
from app.database import get_db
from app.models.user import User
from app.utils.auth import decode_access_token

security = HTTPBearer()

def get_current_user(
    credentials: HTTPAuthorizationCredentials = Depends(security),
    db: Session = Depends(get_db)
) -> User:
    token = credentials.credentials
    payload = decode_access_token(token)
    if payload is None:
        raise HTTPException(status_code=401, detail="Token không hợp lệ")
    user_id = payload.get("sub")
    if user_id is None:
        raise HTTPException(401, detail="Token thiếu thông tin")
    user = db.query(User).filter(User.id == int(user_id)).first()
    if not user:
        raise HTTPException(404, detail="Người dùng không tồn tại")
    return user