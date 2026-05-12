from pydantic import BaseModel
from datetime import datetime

class FileOut(BaseModel):
    id: int
    original_name: str
    file_type: str
    uploaded_at: datetime

    class Config:
        from_attributes = True