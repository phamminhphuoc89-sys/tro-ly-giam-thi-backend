import os
from datetime import date
from app.utils.legacy.bao_cao_thi_dua import (
    read_diem_sdb,
    read_loi_tap_the,
    read_loi_ca_nhan,
    generate_comments,
    update_template
)
from app.models.file import UserFile
from typing import Dict
from app.utils.file_utils import get_user_dir

def process_bao_cao(user_id: int, input_files: Dict[str, int], params: dict, file_map: Dict[int, UserFile]) -> str:
    template_id = input_files["template"]
    diem_id = input_files["diem_sdb"]
    loica_id = input_files["loi_ca_nhan"]
    loitt_id = input_files["loi_tap_the"]

    template_path = file_map[template_id].stored_path
    diem_path = file_map[diem_id].stored_path
    loica_path = file_map[loica_id].stored_path
    loitt_path = file_map[loitt_id].stored_path

    start_date = date.fromisoformat(params["start_date"])
    end_date = date.fromisoformat(params["end_date"])

    # Đọc dữ liệu
    diem_data = read_diem_sdb(diem_path)
    tt_data = read_loi_tap_the(loitt_path)
    cn_data = read_loi_ca_nhan(loica_path)

    comments = generate_comments(diem_data, tt_data, cn_data)

    # Tạo file output
    output_dir = os.path.join(get_user_dir(user_id), "output")
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"BaoCaoThiDua_{date.today().strftime('%Y%m%d')}.docx")

    update_template(template_path, diem_data, tt_data, cn_data, comments, start_date, end_date, output_path)
    return output_path