from sqlalchemy import Column, Integer, String, DateTime, ForeignKey
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func
from app.database import Base

class UserFile(Base):
    __tablename__ = "user_files"
    id = Column(Integer, primary_key=True, index=True)
    owner_id = Column(Integer, ForeignKey("users.id"))
    original_name = Column(String)
    stored_path = Column(String)
    file_type = Column(String)  # 'template_docx', 'diem_sdb', 'loi_ca_nhan', 'loi_tap_the', 'input_excel', 'output'
    uploaded_at = Column(DateTime(timezone=True), server_default=func.now())

    owner = relationship("User", back_populates="files")