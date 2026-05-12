import os
from app.utils.legacy.xeploai import (
    xu_ly_che_do_1, xu_ly_che_do_2, xu_ly_che_do_3, xu_ly_che_do_4,
    save_results
)
from app.models.file import UserFile
from typing import Dict, List
from app.utils.file_utils import get_user_dir

def process_score(user_id: int, input_files: Dict[str, int], params: dict, file_map: Dict[int, UserFile]) -> str:
    mode = params.get("mode", 1)
    # Với chế độ 1-3, input_files là dict các tuần: {"week1": id, "week2": id, ...}
    # Với chế độ 4, input_files có thể là list file id
    if mode in (1,2,3):
        file_paths = [None] * 5
        for key, fid in input_files.items():
            if key.startswith("week"):
                idx = int(key.replace("week","")) - 1
                file_paths[idx] = file_map[fid].stored_path
        # Gọi hàm tương ứng
        if mode == 1:
            result = xu_ly_che_do_1(file_paths)
        elif mode == 2:
            result = xu_ly_che_do_2(file_paths)
        else:
            result = xu_ly_che_do_3(file_paths)
    else:
        # mode 4
        file_paths = [file_map[fid].stored_path for fid in input_files.values()]
        result, _ = xu_ly_che_do_4(file_paths)

    if result is None:
        raise Exception("Không có dữ liệu để xử lý")

    output_dir = os.path.join(get_user_dir(user_id), "output")
    os.makedirs(output_dir, exist_ok=True)
    import datetime
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(output_dir, f"score_result_{ts}.xlsx")
    save_results(result, output_path, mode)
    return output_path