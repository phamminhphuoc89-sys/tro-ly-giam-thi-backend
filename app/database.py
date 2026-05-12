from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base
from app.config import settings
import os

engine = create_engine(
    settings.DATABASE_URL,
    pool_pre_ping=True,        # ✅ Tự kiểm tra connection trước khi dùng
    pool_recycle=300,          # ✅ Tái tạo connection sau 5 phút
    pool_size=5,
    max_overflow=10,
    connect_args={
        "sslmode": "require",  # ✅ Bắt buộc SSL với Neon
        "connect_timeout": 10,
    }
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()