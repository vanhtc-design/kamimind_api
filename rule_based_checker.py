import docx
import json
import re

class RuleBasedChecker:
    def __init__(self, db_mon_hoc_path):
        # Tải database môn học từ Giai đoạn 1 để kiểm tra môn tiên quyết
        try:
            with open(db_mon_hoc_path, 'r', encoding='utf-8') as f:
                self.db_mon_hoc = json.load(f)
        except Exception as e:
            self.db_mon_hoc = []
            print("Cảnh báo: Không tìm thấy file db_mon_hoc.json. Sẽ bỏ qua check môn tiên quyết với DB.")

    def extract_text(self, doc):
        return "\n".join([p.text for p in doc.paragraphs])

    def check_syllabus(self, docx_path):
        print(f"\n[RULE-BASED CHECKER] Đang rà soát file: {docx_path}")
        try:
            doc = docx.Document(docx_path)
            full_text = self.extract_text(doc)
            errors = []

            # 1. Kiểm tra Lỗi Form mẫu
            if "đề cương học phần" not in full_text.lower():
                errors.append("LỖI FORM MẪU: Không tìm thấy tiêu đề 'ĐỀ CƯƠNG HỌC PHẦN'.")
            if "quuyết định số" not in full_text.lower() and "quyết định số" not in full_text.lower():
                errors.append("LỖI FORM MẪU: Thiếu thông tin 'Quyết định số...' ban hành đề cương.")
            if "thông tin chung về học phần" not in full_text.lower():
                errors.append("LỖI FORM MẪU: Thiếu Mục 'A. THÔNG TIN CHUNG VỀ HỌC PHẦN'.")

            # 2. Kiểm tra Môn học trước (Môn tiên quyết)
            ma_mon_hp_match = re.search(r'Mã số học phần:\s*([A-Za-z0-9]+)', full_text, re.IGNORECASE)
            ma_hp_hien_tai = ma_mon_hp_match.group(1).strip() if ma_mon_hp_match else None
            
            tien_quyet_match = re.search(r'Học phần trước:\s*([^\n]+)', full_text, re.IGNORECASE)
            if tien_quyet_match and ma_hp_hien_tai:
                tq_text = tien_quyet_match.group(1).strip()
                # Đối chiếu với DB
                db_course = next((c for c in self.db_mon_hoc if c["Ma_HP"].lower() == ma_hp_hien_tai.lower()), None)
                if db_course and db_course.get("Mon_Tien_Quyet"):
                    # Nếu DB yêu cầu môn tiên quyết mà giảng viên ghi khác hoặc ghi thiếu mã
                    if db_course["Mon_Tien_Quyet"].lower() not in tq_text.lower():
                        errors.append(f"LỖI QUY CHẾ (Môn tiên quyết): CTĐT yêu cầu môn tiên quyết là '{db_course['Mon_Tien_Quyet']}', nhưng Đề cương ghi '{tq_text}'. Vui lòng ghi rõ và đúng mã môn.")

            # 3. Kiểm tra Tín chỉ & Giờ học (Công thức tính giờ theo ghi chú 4, 5, 6)
            # - Tổng thời gian = Số tín chỉ * 50
            # - Giờ lý thuyết (Giảng dạy trên lớp) = TC Lý thuyết * 15
            # - Giờ thực hành = TC Thực hành * 30
            
            # Trích xuất số tín chỉ trong đề cương
            tc_match = re.search(r'Số tín chỉ:\s*(\d+)', full_text, re.IGNORECASE)
            if tc_match:
                so_tc = int(tc_match.group(1))
                tong_gio_ky_vong = so_tc * 50
                
                # Trích xuất tổng phân bổ thời gian
                tong_gio_match = re.search(r'Phân bổ thời gian:\s*(\d+)\s*giờ', full_text, re.IGNORECASE)
                if tong_gio_match:
                    tong_gio_thuc_te = int(tong_gio_match.group(1))
                    if tong_gio_thuc_te != tong_gio_ky_vong:
                        errors.append(f"LỖI QUY CHẾ (Tính toán giờ): Môn học có {so_tc} tín chỉ, Tổng phân bổ thời gian phải là {tong_gio_ky_vong} giờ (TC * 50). Giảng viên đang ghi là {tong_gio_thuc_te} giờ.")
                
                # Trích xuất giờ giảng dạy trên lớp (Lý thuyết)
                lt_match = re.search(r'Giảng dạy trên lớp[^:]*:\s*(\d+)\s*giờ', full_text, re.IGNORECASE)
                if lt_match:
                    gio_lt_thuc_te = int(lt_match.group(1))
                    # Nếu giờ lý thuyết không chia hết cho 15 thì chắc chắn sai công thức (vì TC lý thuyết là số nguyên)
                    if gio_lt_thuc_te % 15 != 0:
                        errors.append(f"LỖI QUY CHẾ (Giờ lý thuyết): Giờ giảng dạy trên lớp ({gio_lt_thuc_te} giờ) không hợp lệ. Phải tính theo công thức: Số TC lý thuyết * 15 giờ.")
                        
                # Trích xuất giờ thực hành
                th_match = re.search(r'Hoạt động thực hành[^:]*:\s*(\d+)\s*giờ', full_text, re.IGNORECASE)
                if th_match:
                    gio_th_thuc_te = int(th_match.group(1))
                    if gio_th_thuc_te > 0 and gio_th_thuc_te % 30 != 0:
                        errors.append(f"LỖI QUY CHẾ (Giờ thực hành): Giờ thực hành ({gio_th_thuc_te} giờ) không hợp lệ. Phải tính theo công thức: Số TC thực hành * 30 giờ.")

            # Kiểm tra dòng quy định về 30% trực tuyến
            if "không vượt quá 30%" not in full_text.lower():
                if "tiết" in full_text.lower() and "trực tuyến" in full_text.lower():
                    errors.append("LỖI QUY CHẾ (Ghi chú trực tuyến): Ghi sai quy định giờ học trực tuyến. Chữ đúng phải là 'không vượt quá 30% tổng thời gian giảng dạy trực tiếp' thay vì tính bằng tiết.")

            # 4. Kiểm tra Tài liệu tham khảo (Năm xuất bản >= 2020)
            # Tìm đoạn văn bản dưới chữ "Tài liệu bắt buộc" hoặc "Tài liệu chính"
            tl_bat_buoc_idx = full_text.lower().find("tài liệu bắt buộc")
            tl_tham_khao_idx = full_text.lower().find("tài liệu tham khảo", tl_bat_buoc_idx + 1)
            
            if tl_bat_buoc_idx != -1:
                end_idx = tl_tham_khao_idx if tl_tham_khao_idx != -1 else tl_bat_buoc_idx + 1000
                tl_chinh_text = full_text[tl_bat_buoc_idx:end_idx]
                
                # Quét các năm xuất bản (4 chữ số từ 1900 đến 2099)
                years = re.findall(r'\b(19\d{2}|20\d{2})\b', tl_chinh_text)
                for year_str in years:
                    year = int(year_str)
                    if year < 2020:
                        errors.append(f"LỖI QUY CHẾ (Tài liệu): Tài liệu chính phát hiện xuất bản năm {year}. Yêu cầu: Tài liệu tham khảo chính phải xuất bản từ năm 2020 đến nay.")
                        break # Chỉ cần báo 1 lỗi là đủ

            # 5. Kiểm tra sự tồn tại của Ma trận CĐR (Bảng 15.3 tương đương trong đề cương)
            if "ma trận tích hợp" not in full_text.lower() and "ma trận chuẩn đầu ra" not in full_text.lower():
                errors.append("LỖI CẤU TRÚC (Thiếu Bảng): Không tìm thấy 'Ma trận tích hợp giữa CĐR MH và CTĐT'. Bảng này bắt buộc phải có.")

            return errors

        except Exception as e:
            return [f"LỖI HỆ THỐNG: Không thể đọc file {docx_path}. Chi tiết: {e}"]

if __name__ == "__main__":
    # Khởi tạo checker với database đã lấy từ Bước 1
    checker = RuleBasedChecker("db_mon_hoc.json")
    
    # Test thử với file Đề cương ECS701
    test_file = r"D:\Kiemdo\ECS701 Marketing cho Thuong mai dien tu.docx"
    
    errors = checker.check_syllabus(test_file)
    
    print("\n--- BÁO CÁO RÀ SOÁT LỚP CƠ HỌC (RULE-BASED) ---")
    if len(errors) == 0:
        print("✅ Đề cương đạt chuẩn Hình thức & Quy chế cơ bản!")
    else:
        for err in errors:
            print(f"❌ {err}")
