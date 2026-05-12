from sqlalchemy import Column, Integer, String, JSON, DateTime, ForeignKey, relationship
from sqlalchemy.sql import func
from app.database import Base

class ProcessingJob(Base):
    __tablename__ = "processing_jobs"
    id = Column(Integer, primary_key=True, index=True)
    owner_id = Column(Integer, ForeignKey("users.id"))
    job_type = Column(String)   # 'bao_cao_thi_dua', 'filter_ai', 'error_extract', 'score'
    status = Column(String, default="pending")  # pending, processing, completed, failed
    input_files = Column(JSON)  # {"template": id, "diem_sdb": id, ...}
    params = Column(JSON, nullable=True)
    output_file_id = Column(Integer, ForeignKey("user_files.id"), nullable=True)
    error_message = Column(String, nullable=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())

    owner = relationship("User", back_populates="jobs")
    output_file = relationship("UserFile")