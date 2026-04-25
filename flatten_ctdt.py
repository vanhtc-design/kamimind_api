import docx
import json
import re
import os

def find_table_by_keywords(doc, keywords, exclude_keywords=None):
    """Find a table containing all keywords and none of the exclude_keywords."""
    for i, table in enumerate(doc.tables):
        header_text = ""
        try:
            # Check first few rows for keywords to identify the table
            for r in range(min(5, len(table.rows))):
                header_text += " " + " ".join([cell.text for cell in table.rows[r].cells])
        except:
            continue
            
        header_text = header_text.lower()
        if all(k.lower() in header_text for k in keywords):
            if exclude_keywords and any(ek.lower() in header_text for ek in exclude_keywords):
                continue
            return table, i
    return None, -1

def clean_text(text):
    """Normalize text by removing extra spaces and newlines."""
    if not text: return ""
    return re.sub(r'\s+', ' ', text).strip()

def extract_course_list_12_2(table):
    """Extract course list from Table 12.2 (Khung chương trình)."""
    if not table: return []
    
    # Identify column roles
    header_rows_count = 1
    # Search for the row containing column names
    for r_idx in range(min(5, len(table.rows))):
        row_text = " ".join([cell.text.lower() for cell in table.rows[r_idx].cells])
        if "mã học phần" in row_text and "tên học phần" in row_text:
            header_rows_count = r_idx + 1
            break
            
    # Map column indices
    # We'll use the last header row to find column roles
    header_cells = [clean_text(c.text).lower() for c in table.rows[header_rows_count-1].cells]
    
    col_map = {
        "stt": -1, "ma_hp": -1, "ten_hp": -1, "so_tc": -1, 
        "lt": -1, "th": -1, "khac": -1, "tq": -1, "hk": -1
    }
    
    for idx, text in enumerate(header_cells):
        if "stt" in text or "tt" == text: col_map["stt"] = idx
        elif "mã học phần" in text or "mã hp" in text: col_map["ma_hp"] = idx
        elif "tên học phần" in text or "tên hp" in text: col_map["ten_hp"] = idx
        elif "số tín chỉ" in text or "số tc" in text: col_map["so_tc"] = idx
        elif "lý thuyết" in text or "lt" == text: col_map["lt"] = idx
        elif "thực hành" in text or "th" == text: col_map["th"] = idx
        elif "khác" in text: col_map["khac"] = idx
        elif "tiên quyết" in text or "tq" in text: col_map["tq"] = idx
        elif "học kỳ" in text or "hk" in text: col_map["hk"] = idx

    # If column roles weren't found in one row, search the whole header block
    for r in range(header_rows_count):
        row_cells = [clean_text(c.text).lower() for c in table.rows[r].cells]
        for idx, text in enumerate(row_cells):
            if col_map["stt"] == -1 and ("stt" in text or "tt" == text): col_map["stt"] = idx
            if col_map["ma_hp"] == -1 and ("mã học phần" in text or "mã hp" in text): col_map["ma_hp"] = idx
            if col_map["ten_hp"] == -1 and ("tên học phần" in text or "tên hp" in text): col_map["ten_hp"] = idx
            if col_map["so_tc"] == -1 and ("số tín chỉ" in text or "số tc" in text): col_map["so_tc"] = idx
            if col_map["lt"] == -1 and ("lý thuyết" in text or "lt" == text): col_map["lt"] = idx
            if col_map["th"] == -1 and ("thực hành" in text or "th" == text): col_map["th"] = idx
            if col_map["khac"] == -1 and "khác" in text: col_map["khac"] = idx
            if col_map["tq"] == -1 and ("tiên quyết" in text or "tq" in text): col_map["tq"] = idx
            if col_map["hk"] == -1 and ("học kỳ" in text or "hk" in text): col_map["hk"] = idx

    results = []
    for r_idx in range(header_rows_count, len(table.rows)):
        try:
            row = table.rows[r_idx].cells
            ma_hp = clean_text(row[col_map["ma_hp"]].text) if col_map["ma_hp"] != -1 else ""
            if not ma_hp or len(ma_hp) < 3: continue # Skip category headers
            
            course = {
                "STT": clean_text(row[col_map["stt"]].text) if col_map["stt"] != -1 else "",
                "Ma_HP": ma_hp,
                "Ten_HP": clean_text(row[col_map["ten_hp"]].text) if col_map["ten_hp"] != -1 else "",
                "So_TC": clean_text(row[col_map["so_tc"]].text) if col_map["so_tc"] != -1 else "",
                "Ly_Thuyet": clean_text(row[col_map["lt"]].text) if col_map["lt"] != -1 else "",
                "Thuc_Hanh": clean_text(row[col_map["th"]].text) if col_map["th"] != -1 else "",
                "Khac": clean_text(row[col_map["khac"]].text) if col_map["khac"] != -1 else "",
                "Mon_Tien_Quyet": clean_text(row[col_map["tq"]].text) if col_map["tq"] != -1 else "",
                "Hoc_Ky": clean_text(row[col_map["hk"]].text) if col_map["hk"] != -1 else ""
            }
            results.append(course)
        except: continue
        
    return results

def extract_mapping_15_3(table):
    """Extract mapping from Table 15.3 (CLO to PI/PLO)."""
    if not table: return []

    pi_row_idx, plo_row_idx, header_rows_count = -1, -1, 0
    
    # Scan for PI row (numbers like 1.1)
    for r_idx in range(min(10, len(table.rows))):
        cells = [cell.text.strip() for cell in table.rows[r_idx].cells]
        if any(re.search(r'\d+\.\d+', c) for c in cells):
            pi_row_idx = r_idx
            # Find PLO row above
            for prev_r in range(r_idx - 1, -1, -1):
                prev_cells = [cell.text.strip().upper() for cell in table.rows[prev_r].cells]
                if any("PLO" in c for c in prev_cells):
                    plo_row_idx = prev_r
                    break
            if plo_row_idx == -1: plo_row_idx = max(0, r_idx - 1)
            header_rows_count = r_idx + 1
            break
            
    if pi_row_idx == -1: return []

    # Map column roles
    stt_col, hk_col, ten_hp_col, clo_col = 0, 1, 2, 3
    for r in range(header_rows_count):
        cells = [c.text.strip().lower() for c in table.rows[r].cells]
        for idx, text in enumerate(cells):
            if "stt" in text or "tt" == text: stt_col = idx
            elif "học kỳ" in text or "hk" == text: hk_col = idx
            elif "tên học phần" in text or "tên hp" in text: ten_hp_col = idx
            elif "clo" in text: clo_col = idx

    col_mapping = {}
    pi_cells = table.rows[pi_row_idx].cells
    plo_cells = table.rows[plo_row_idx].cells
    
    for idx in range(clo_col + 1, len(table.columns)):
        try:
            pi_val = clean_text(pi_cells[idx].text)
            plo_val = clean_text(plo_cells[idx].text)
            if not plo_val: # Handle merged
                for back_idx in range(idx - 1, clo_col, -1):
                    prev_plo = clean_text(plo_cells[back_idx].text)
                    if prev_plo:
                        plo_val = prev_plo
                        break
            if pi_val and re.search(r'\d+\.\d+', pi_val):
                col_mapping[idx] = {"PLO": plo_val or "PLO?", "PI": pi_val}
        except: continue

    results = []
    current_course = None
    c_stt, c_hk, c_ten = "", "", ""
    
    for r_idx in range(header_rows_count, len(table.rows)):
        try:
            cells = table.rows[r_idx].cells
            stt = clean_text(cells[stt_col].text)
            ten = clean_text(cells[ten_hp_col].text)
            clo = clean_text(cells[clo_col].text)
            
            if not clo or "tổng" in ten.lower(): continue
            
            if stt: c_stt = stt
            if ten: c_ten = ten
            hk = clean_text(cells[hk_col].text)
            if hk: c_hk = hk
            
            if not current_course or (stt and stt != current_course.get("STT")):
                if current_course: results.append(current_course)
                current_course = {"STT": c_stt, "Ten_HP": c_ten, "Hoc_Ky": c_hk, "Mappings": []}
            
            for col_idx, meta in col_mapping.items():
                val = clean_text(cells[col_idx].text)
                if val and re.search(r'[234xX]', val):
                    level = int(val) if val.isdigit() else 3
                    current_course["Mappings"].append({
                        "CLO": clo, "PLO": meta["PLO"], "PI": meta["PI"], "Level": level
                    })
        except: continue

    if current_course: results.append(current_course)
    return results

def extract_ctdt_data(docx_path):
    """Main entry point for extracting all data from a CTĐT document."""
    print(f"Processing: {docx_path}")
    doc = docx.Document(docx_path)
    
    # 1. Extract Course List (12.2)
    t12_2, _ = find_table_by_keywords(doc, ["mã học phần", "tên học phần"])
    courses = extract_course_list_12_2(t12_2)
    
    # 2. Extract Mapping Matrix (15.3)
    t15_3, _ = find_table_by_keywords(doc, ["clo", "pi", "học phần"])
    if not t15_3:
        t15_3, _ = find_table_by_keywords(doc, ["clo", "plo", "học phần"])
    matrix = extract_mapping_15_3(t15_3)
    
    # 3. Extract PLO Text (for phase 3)
    plo_list = []
    t_plo, _ = find_table_by_keywords(doc, ["plo", "mức độ", "chuẩn đầu ra"], exclude_keywords=["học phần"])
    if t_plo:
        for row in t_plo.rows:
            cells = [clean_text(c.text) for c in row.cells]
            plo_match = next((c for c in cells if re.match(r'^PLO\d+$', c)), None)
            if plo_match:
                idx = cells.index(plo_match)
                if idx + 1 < len(cells):
                    plo_list.append({"Ma_PLO": plo_match, "Mo_Ta": cells[idx+1]})

    return courses, matrix, plo_list

if __name__ == "__main__":
    import sys
    path = r"D:\Kiemdo\10.CTDT_He thong thong tin kinh doanh_HTTTQL_CQC.docx"
    if os.path.exists(path):
        c, m, p = extract_ctdt_data(path)
        print(f"Extracted {len(c)} courses, {len(m)} matrices, {len(p)} PLOs.")
        with open("final_extraction.json", "w", encoding="utf-8") as f:
            json.dump({"courses": c, "matrix": m, "plos": p}, f, ensure_ascii=False, indent=2)
