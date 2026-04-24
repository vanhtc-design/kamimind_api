import os
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from typing import List
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
import tempfile
import json

# Import các hàm từ 2 script đã viết
from flatten_ctdt import extract_ctdt_data
from rule_based_checker import RuleBasedChecker

app = FastAPI(
    title="KamiMind Tools API",
    description="API cung cấp các công cụ Rule-based và Flattening cho KamiMind AI Agent.",
    version="1.0.0"
)

# Thêm cấu hình CORS để cho phép Web Demo gọi API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Cho phép tất cả các nguồn gọi API
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Thư mục lưu DB tạm thời
DB_PATH = "db_mon_hoc.json"

@app.post("/api/v1/flatten-ctdt", summary="Trải phẳng CTĐT thành JSON (Giai đoạn 1)")
async def flatten_ctdt(files: List[UploadFile] = File(...)):
    """
    Nhận một hoặc nhiều file CTĐT (.docx) và trả về dữ liệu cấu trúc JSON (Môn học, Ma trận 15.3, PLO) gộp lại.
    """
    if not files:
        raise HTTPException(status_code=400, detail="Không có file nào được tải lên.")
        
    try:
        all_courses = []
        all_matrix = []
        all_plos = []
        processed_files = []

        for file in files:
            if not file.filename.endswith('.docx'):
                continue
                
            # Lưu file tạm để xử lý
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
                content = await file.read()
                temp_file.write(content)
                temp_path = temp_file.name

            try:
                # Gọi hàm trích xuất
                courses, matrix, plos = extract_ctdt_data(temp_path)
                all_courses.extend(courses)
                all_matrix.extend(matrix)
                all_plos.extend(plos)
                processed_files.append(file.filename)
            finally:
                # Xóa file tạm
                if os.path.exists(temp_path):
                    os.remove(temp_path)
        
        if not processed_files:
            raise HTTPException(status_code=400, detail="Không có file .docx nào hợp lệ để xử lý.")

        # Cập nhật lại db_mon_hoc.json cục bộ (Gộp tất cả môn học của các ngành)
        with open(DB_PATH, "w", encoding="utf-8") as f:
            json.dump(all_courses, f, ensure_ascii=False, indent=2)

        return JSONResponse(content={
            "status": "success",
            "message": f"Trích xuất dữ liệu thành công từ {len(processed_files)} CTĐT.",
            "data": {
                "processed_files": processed_files,
                "courses_count": len(all_courses),
                "matrix_count": len(all_matrix),
                "plos_count": len(all_plos),
                "courses": all_courses,
                "matrix": all_matrix,
                "plos": all_plos
            }
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xử lý file: {str(e)}")


@app.post("/api/v1/check-syllabus", summary="Kiểm tra Lỗi Cơ học Đề cương (Giai đoạn 2)")
async def check_syllabus(file: UploadFile = File(...)):
    """
    Nhận file Đề cương (.docx), quét các lỗi Rule-based (Form, Tín chỉ, Môn tiên quyết, Tài liệu <2020).
    """
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Chỉ hỗ trợ định dạng .docx")
        
    try:
        # Khởi tạo Checker
        if not os.path.exists(DB_PATH):
            # Nếu chưa có DB, tạo một mảng rỗng để không bị lỗi crash
            checker = RuleBasedChecker("dummy_path_that_fails_safely")
            checker.db_mon_hoc = []
        else:
            checker = RuleBasedChecker(DB_PATH)

        # Lưu file tạm để xử lý
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
            content = await file.read()
            temp_file.write(content)
            temp_path = temp_file.name

        # Gọi hàm check
        errors = checker.check_syllabus(temp_path)
        
        # Xóa file tạm
        os.remove(temp_path)

        return JSONResponse(content={
            "status": "success",
            "file_checked": file.filename,
            "total_errors": len(errors),
            "errors": errors
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xử lý file: {str(e)}")

if __name__ == "__main__":
    print("Khởi động KamiMind API Server tại http://0.0.0.0:8000")
    uvicorn.run("api_server:app", host="0.0.0.0", port=8000, reload=True)
