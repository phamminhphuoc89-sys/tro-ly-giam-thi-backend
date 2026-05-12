import os
from app.utils.legacy.filter_v2 import ErrorExtractor, Config as ErrConfig
from app.models.file import UserFile
from app.models.user import User
from typing import Dict
from app.utils.file_utils import get_user_dir
from sqlalchemy.orm import Session

def process_error_extract(user_id: int, input_files: Dict[str, int], params: dict, file_map: Dict[int, UserFile], db: Session = None) -> str:
    input_id = input_files["input_excel"]
    template_id = input_files["template"]

    input_path = file_map[input_id].stored_path
    template_path = file_map[template_id].stored_path
    use_gpt = params.get("use_gpt", False)

    extractor = ErrorExtractor()

    # Lấy API key của user nếu dùng GPT
    if use_gpt and db:
        user = db.query(User).filter(User.id == user_id).first()
        if user and user.openai_api_key:
            extractor.set_api_key(user.openai_api_key)
        else:
            use_gpt = False  # fallback

    success, msg = extractor.load_input_file(input_path)
    if not success:
        raise Exception(f"Lỗi load file: {msg}")

    success, msg = extractor.extract_errors(use_gpt=use_gpt)
    if not success:
        raise Exception(f"Lỗi extract: {msg}")

    output_dir = os.path.join(get_user_dir(user_id), "output")
    os.makedirs(output_dir, exist_ok=True)
    import datetime
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(output_dir, f"LOI_CA_NHAN_{ts}.xlsx")

    extractor.generate_individual_report(template_path, output_path)
    return output_path