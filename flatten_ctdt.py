import docx
import json
import pandas as pd
import re

def extract_ctdt_data(docx_path):
    print(f"Đang đọc file: {docx_path}")
    doc = docx.Document(docx_path)
    
    course_list = []
    matrix_mapping = {}
    plo_list = []
    
    for table in doc.tables:
        if len(table.rows) == 0:
            continue
            
        # Lấy văn bản của dòng đầu tiên để nhận diện bảng
        header_text = " ".join([cell.text for cell in table.rows[0].cells]).lower()
        
        # =====================================================================
        # 1. Trích xuất Danh mục Môn học & Môn tiên quyết (Bảng 12.2)
        # =====================================================================
        if "mã học phần" in header_text and "tên học phần" in header_text and ("tín chỉ" in header_text or "tc" in header_text):
            print("-> Đã tìm thấy Bảng Khung chương trình (Danh mục môn học).")
            data = [[cell.text.strip() for cell in row.cells] for row in table.rows]
            df = pd.DataFrame(data)
            df.columns = df.iloc[0] # Dùng dòng đầu làm Header
            df = df[1:]
            
            # Tìm các cột tương ứng bằng Regex
            cols = df.columns.astype(str).str.lower()
            ma_hp_col = df.columns[cols.str.contains('mã')]
            ten_hp_col = df.columns[cols.str.contains('tên')]
            tc_col = df.columns[cols.str.contains('tín chỉ|tc')]
            tq_col = df.columns[cols.str.contains('tiên quyết|trước')]
            
            if len(ma_hp_col) > 0 and len(ten_hp_col) > 0:
                for _, row in df.iterrows():
                    ma_hp = row[ma_hp_col[0]]
                    ten_hp = row[ten_hp_col[0]]
                    if not ma_hp or ma_hp.lower() == 'mã học phần':
                        continue
                    
                    tc = row[tc_col[0]] if len(tc_col) > 0 else ""
                    tq = row[tq_col[0]] if len(tq_col) > 0 else ""
                    
                    # Xử lý làm sạch tên môn tiên quyết (nếu có xuống dòng)
                    tq = tq.replace('\n', ', ') if tq else ""
                    
                    course_list.append({
                        "Ma_HP": ma_hp,
                        "Ten_HP": ten_hp,
                        "So_TC": tc,
                        "Mon_Tien_Quyet": tq
                    })
                    
        # =====================================================================
        # 2. Trích xuất Ma trận Môn học - CLO - PI (Bảng 15.3)
        # =====================================================================
        if "clo" in header_text and "pi" in header_text and "học phần" in header_text:
            print("-> Đã tìm thấy Bảng Ma trận đóng góp (Bảng 15.3).")
            data = [[cell.text.strip() for cell in row.cells] for row in table.rows]
            df = pd.DataFrame(data)
            
            # Word thường có header gồm nhiều dòng do merge cells. Gộp 2 dòng đầu.
            header = []
            for col_idx in range(len(df.columns)):
                col_header = " ".join([str(df.iloc[r, col_idx]) for r in range(2)]).strip()
                header.append(col_header)
            
            # Xác định vị trí các cột PI (Ví dụ: PI 1.1, PI 2.1)
            pi_cols = {}
            for idx, h in enumerate(header):
                match = re.search(r'PI\s*(\d+\.\d+)', h, re.IGNORECASE)
                if match:
                    pi_cols[idx] = "PI_" + match.group(1)
            
            # Xác định các cột cơ bản
            ma_idx, ten_idx, clo_idx = -1, -1, -1
            for idx, h in enumerate(header):
                h_low = h.lower()
                if "mã" in h_low: ma_idx = idx
                elif "tên" in h_low: ten_idx = idx
                elif "clo" in h_low: clo_idx = idx
                
            if ma_idx != -1 and clo_idx != -1 and len(pi_cols) > 0:
                current_ma = ""
                current_ten = ""
                for row_idx in range(2, len(df)):
                    row = df.iloc[row_idx]
                    ma = row[ma_idx]
                    
                    # Xử lý ô gộp (merged cells) của Mã học phần
                    if ma: 
                        current_ma = ma
                        current_ten = row[ten_idx] if ten_idx != -1 else ""
                        
                    clo = row[clo_idx]
                    if not current_ma or not clo:
                        continue
                        
                    if current_ma not in matrix_mapping:
                        matrix_mapping[current_ma] = {
                            "Ma_HP": current_ma,
                            "Ten_HP": current_ten,
                            "Mapping": {}
                        }
                        
                    if clo not in matrix_mapping[current_ma]["Mapping"]:
                        matrix_mapping[current_ma]["Mapping"][clo] = {}
                        
                    # Lấy mức độ (thường là 3, 4) từ các cột PI
                    for c_idx, pi_name in pi_cols.items():
                        val = row[c_idx]
                        if val and val.isdigit():
                            matrix_mapping[current_ma]["Mapping"][clo][pi_name] = int(val)

        # =====================================================================
        # 3. Trích xuất Nội dung PLO để đối chiếu Text-Matching (Giai đoạn 3)
        # =====================================================================
        if "plo" in header_text and "mức độ" in header_text and ("cđr" in header_text or "đầu ra" in header_text):
             print("-> Đã tìm thấy Bảng Chuẩn đầu ra (PLOs).")
             data = [[cell.text.strip() for cell in row.cells] for row in table.rows]
             
             for row in data:
                 if len(row) >= 3:
                     # Tìm ô có chứa chữ PLO
                     plo_cell = next((c for c in row if "PLO" in c.upper() and len(c) < 10), None)
                     if plo_cell:
                         # Giả sử mô tả nằm ở ô tiếp theo
                         idx = row.index(plo_cell)
                         if idx + 1 < len(row):
                             plo_desc = row[idx+1]
                             plo_list.append({
                                 "Ma_PLO": plo_cell,
                                 "Mo_Ta_Chinh_Xac": plo_desc
                             })

    return course_list, list(matrix_mapping.values()), plo_list

if __name__ == "__main__":
    # Thay đường dẫn này bằng đường dẫn tới file CTĐT thực tế
    input_file = r"D:\Kiemdo\10.CTDT_He thong thong tin kinh doanh_HTTTQL_CQC.docx"
    
    try:
        courses, matrix, plos = extract_ctdt_data(input_file)
        
        # Lưu kết quả ra file JSON
        with open("db_mon_hoc.json", "w", encoding="utf-8") as f:
            json.dump(courses, f, ensure_ascii=False, indent=2)
            
        with open("db_ma_tran_15_3.json", "w", encoding="utf-8") as f:
            json.dump(matrix, f, ensure_ascii=False, indent=2)
            
        with open("db_plo_text.json", "w", encoding="utf-8") as f:
            json.dump(plos, f, ensure_ascii=False, indent=2)
            
        print("\n=> XUẤT DỮ LIỆU THÀNH CÔNG!")
        print(f"Đã trích xuất {len(courses)} môn học.")
        print(f"Đã trích xuất ma trận cho {len(matrix)} môn học.")
        print(f"Đã trích xuất {len(plos)} PLO để so khớp văn bản.")
        
    except Exception as e:
        print(f"Lỗi: {e}")
