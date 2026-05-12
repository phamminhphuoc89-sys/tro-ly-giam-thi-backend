from pydantic import BaseModel
from typing import Dict, Optional, Any
from datetime import datetime

class JobCreate(BaseModel):
    job_type: str
    input_files: Dict[str, int]   # key -> file_id
    params: Optional[Dict[str, Any]] = {}

class JobOut(BaseModel):
    id: int
    job_type: str
    status: str
    input_files: dict
    params: dict
    output_file_id: Optional[int]
    error_message: Optional[str]
    created_at: datetime

    class Config:
        orm_mode = True