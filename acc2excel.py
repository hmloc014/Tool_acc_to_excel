# encoding: utf-8
import re
import pandas as pd

# ====== Cấu hình ======
acc_file = "2030-K-MAX.acc"   # Đường dẫn file .acc
excel_file = "output.xlsx"  # File Excel xuất ra

# ====== Hàm đọc file ======
def read_acc_file(file_path):
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
    return lines

# ====== Hàm phân tích nội dung ======
def parse_acc_sections(lines):
    """
    Tách các phần (section) trong file .acc theo tiêu đề lớn (chữ in hoa)
    """
    data_sections = {}
    current_section = None
    current_data = []

    for line in lines:
        # Dòng tiêu đề lớn (VD: "BUS VOLTAGE REPORT")
        if re.match(r'^[A-Z0-9 \-\/()]+$', line.strip()) and len(line.strip()) > 5:
            # Nếu đang có section cũ thì lưu lại
            if current_section and current_data:
                data_sections[current_section] = current_data
            current_section = line.strip()
            current_data = []
        else:
            if current_section:
                current_data.append(line.strip())

    # Lưu section cuối cùng
    if current_section and current_data:
        data_sections[current_section] = current_data

    return data_sections

# ====== Hàm xử lý từng bảng ======
def parse_table(section_lines):
    """
    Tự động nhận diện cột bằng khoảng trắng
    """
    clean_lines = [l for l in section_lines if l and not l.startswith('-')]

    # Tìm header
    header_line = None
    for l in clean_lines:
        if re.search(r'[A-Za-z]', l):
            header_line = l
            break

    if not header_line:
        return None

    headers = re.split(r'\s{2,}', header_line.strip())
    data = []

    # Lấy phần dữ liệu sau header
    header_index = clean_lines.index(header_line)
    for l in clean_lines[header_index + 1:]:
        if l.strip() == "":
            continue
        row = re.split(r'\s{2,}', l.strip())
        if len(row) == len(headers):
            data.append(row)

    if not data:
        return None
    df = pd.DataFrame(data, columns=headers)
    return df

# ====== Thực thi ======
lines = read_acc_file(acc_file)
sections = parse_acc_sections(lines)

with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    for section_name, section_lines in sections.items():
        df = parse_table(section_lines)
        if df is not None:
            # Tên sheet giới hạn 31 ký tự
            sheet_name = section_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"✅ Đã xuất dữ liệu từ {acc_file} ra {excel_file}")
