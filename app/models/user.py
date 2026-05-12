from sqlalchemy import Column, Integer, String, DateTime, relationship
from sqlalchemy.sql import func
from app.database import Base

class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    email = Column(String, unique=True, index=True)
    username = Column(String, unique=True, index=True)
    hashed_password = Column(String)
    openai_api_key = Column(String, nullable=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    files = relationship("UserFile", back_populates="owner")
    jobs = relationship("ProcessingJob", back_populates="owner")