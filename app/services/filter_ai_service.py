import os
from app.utils.legacy.excel_filter_ai import ExcelFilterApp
from app.models.file import UserFile
from typing import Dict
from app.utils.file_utils import get_user_dir

def process_filter_ai(user_id: int, input_files: Dict[str, int], params: dict, file_map: Dict[int, UserFile]) -> str:
    input_id = input_files["input_excel"]
    input_path = file_map[input_id].stored_path

    split_all = params.get("split_all_students", True)
    auto_ai = params.get("auto_ai", True)
    auto_map = params.get("auto_map", True)

    app = ExcelFilterApp()
    app.split_all_students = split_all

    success, msg = app.load_file(input_path)
    if not success:
        raise Exception(f"Lỗi load file: {msg}")

    if auto_map and app.student_mapping:
        app._add_student_names()

    if auto_ai and 'đánh giá' in app.original_df.columns:
        app.analyze_students_ai()

    # Áp dụng bộ lọc (đã có sẵn trong app)
    success, msg = app.apply_filters()
    if not success:
        raise Exception(f"Lỗi filter: {msg}")

    output_dir = os.path.join(get_user_dir(user_id), "output")
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"filtered_ai_{os.path.basename(input_path)}")

    success, msg = app.save_filtered_file(output_path)
    if not success:
        raise Exception(f"Lỗi lưu file: {msg}")
    return output_path