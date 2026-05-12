import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.database import engine, Base
from app.config import settings
from app.routes import auth, files, jobs, students

# Tạo bảng
Base.metadata.create_all(bind=engine)

# Đảm bảo thư mục upload tồn tại
os.makedirs(settings.UPLOAD_DIR, exist_ok=True)

app = FastAPI(title="Trợ lý Giáo vụ API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(auth.router, prefix="/auth", tags=["auth"])
app.include_router(files.router, prefix="/files", tags=["files"])
app.include_router(jobs.router, prefix="/jobs", tags=["jobs"])
app.include_router(students.router, prefix="/students", tags=["students"])

@app.get("/")
def root():
    return {"message": "Trợ lý Giáo vụ API is running"}