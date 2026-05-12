# app/routes/jobs.py
from fastapi import APIRouter, Depends, HTTPException, BackgroundTasks
from sqlalchemy.orm import Session
from app.database import get_db
from app.models.job import ProcessingJob
from app.models.file import UserFile
from app.schemas.job import JobCreate, JobOut
from app.dependencies import get_current_user
from app.services.bao_cao_service import process_bao_cao
from app.services.filter_ai_service import process_filter_ai
from app.services.error_extract_service import process_error_extract
from app.services.score_service import process_score
import logging

router = APIRouter()
logger = logging.getLogger(__name__)

def run_job(job_id: int):
    from app.database import SessionLocal
    db = SessionLocal()
    try:
        job = db.query(ProcessingJob).filter(ProcessingJob.id == job_id).first()
        if not job:
            return
        job.status = "processing"
        db.commit()

        # Lấy file map
        file_ids = job.input_files.values()
        user_files = db.query(UserFile).filter(UserFile.id.in_(file_ids)).all()
        file_map = {f.id: f for f in user_files}

        output_path = None
        if job.job_type == "bao_cao_thi_dua":
            output_path = process_bao_cao(job.owner_id, job.input_files, job.params, file_map)
        elif job.job_type == "filter_ai":
            output_path = process_filter_ai(job.owner_id, job.input_files, job.params, file_map)
        elif job.job_type == "error_extract":
            output_path = process_error_extract(job.owner_id, job.input_files, job.params, file_map, db)
        elif job.job_type == "score":
            output_path = process_score(job.owner_id, job.input_files, job.params, file_map)
        else:
            raise ValueError("Unknown job type")

        if output_path:
            import os
            filename = os.path.basename(output_path)
            out_file = UserFile(
                owner_id=job.owner_id,
                original_name=filename,
                stored_path=output_path,
                file_type="output"
            )
            db.add(out_file)
            db.commit()
            db.refresh(out_file)
            job.output_file_id = out_file.id
        job.status = "completed"
        db.commit()
    except Exception as e:
        logger.error(f"Job {job_id} failed: {str(e)}")
        job.status = "failed"
        job.error_message = str(e)
        db.commit()
    finally:
        db.close()

@router.post("/", response_model=JobOut)
def create_job(
    job_in: JobCreate,
    background_tasks: BackgroundTasks,
    db: Session = Depends(get_db),
    current_user = Depends(get_current_user)
):
    for file_id in job_in.input_files.values():
        file = db.query(UserFile).filter(UserFile.id == file_id, UserFile.owner_id == current_user.id).first()
        if not file:
            raise HTTPException(400, f"File ID {file_id} không hợp lệ hoặc không thuộc quyền sở hữu")

    job = ProcessingJob(
        owner_id=current_user.id,
        job_type=job_in.job_type,
        input_files=job_in.input_files,
        params=job_in.params,
        status="pending"
    )
    db.add(job)
    db.commit()
    db.refresh(job)
    
    background_tasks.add_task(run_job, job.id)
    return job

@router.get("/{job_id}", response_model=JobOut)
def get_job(job_id: int, db: Session = Depends(get_db), current_user = Depends(get_current_user)):
    job = db.query(ProcessingJob).filter(ProcessingJob.id == job_id, ProcessingJob.owner_id == current_user.id).first()
    if not job:
        raise HTTPException(404, "Job không tồn tại")
    return job