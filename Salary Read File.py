
def Salary_read_file:

import os
import pandas as pd
from fpdf import FPDF

# Ensure the directory exists
os.makedirs("salary_slips", exist_ok=True)

# Load the workbook
file_path = "Salary_data.xls"
xls = pd.ExcelFile(file_path)

# Process the relevant sheet
try:
    # Parse the sheet (adjust the sheet name as needed)
    data = xls.parse(sheet_name='Salary Sheet')  # Update 'Salary Sheet' if necessary
    #data.to_csv('data.csv')
except ValueError as e:
    print(f"Error loading sheet: {e}")
    exit()

data = data.iloc[10:] 
print(data)
# Skip rows before row 11 (index 10)
 # Keep rows starting from index 10 (row 11 in Excel)

# Create a PDF class that supports Unicode
class PDF(FPDF):
    def header(self):
        self.set_font('DejaVu', '', 14)  # Use the Unicode font
        self.cell(0, 10, 'CÔNG TY TNHH MARUICHI SUN STEEL (HÀ NỘI)', align='L', ln=True)

# Ensure the DejaVuSans.ttf font file is available
font_path = "DejaVuSans.ttf"  # Update the path if necessary
if not os.path.exists(font_path):
    raise FileNotFoundError("The required font file 'DejaVuSans.ttf' is missing. Please add it to the directory.")

# Generate PDFs
for index, row in data.iterrows():
    try:
        # Create a PDF instance and register the Unicode font
        pdf = PDF()
        pdf.add_font("DejaVu", "", font_path, uni=True)
        pdf.add_page()
        pdf.set_font("DejaVu", size=12)  # Use the Unicode font
        # Access columns (update with actual column names or indices)
        employee_name = row['Unnamed: 2'] if 'Unnamed: 2' in row else None # Tong luong
        staff_id = row['Unnamed: 1'] if 'Unnamed: 1' in row else None # Tong luong

        total06 = row['Unnamed: 6'] if 'Unnamed: 6' in row else None # Tong luong
        total07 = row['Unnamed: 7'] if 'Unnamed: 7' in row else None # Tong luong
        total08 = row['Unnamed: 8'] if 'Unnamed: 8' in row else None # Tong luong
        total09 = row['Unnamed: 9'] if 'Unnamed: 9' in row else None # Tong luong
        total10 = row['Unnamed: 10'] if 'Unnamed: 10' in row else None # Tong luong
        total11 = row['Unnamed: 11'] if 'Unnamed: 11' in row else None # Tong luong
        total12 = row['Unnamed: 12'] if 'Unnamed: 12' in row else None # Tong luong
        total13 = row['Unnamed: 13'] if 'Unnamed: 13' in row else None # Tong luong
        total14 = row['Unnamed: 14'] if 'Unnamed: 14' in row else None # Tong luong
        total15 = row['Unnamed: 15'] if 'Unnamed: 15' in row else None # Tong luong
        total16 = row['Unnamed: 16'] if 'Unnamed: 16' in row else None # Tong luong
        total17 = row['Unnamed: 17'] if 'Unnamed: 17' in row else None # Tong luong
        total18 = row['Unnamed: 18'] if 'Unnamed: 18' in row else None # Tong luong
        total19 = row['Unnamed: 19'] if 'Unnamed: 19' in row else None # Tong luong
        total20 = row['Unnamed: 20'] if 'Unnamed: 20' in row else None # Tong luong
        total21 = row['Unnamed: 21'] if 'Unnamed: 21' in row else None # Tong luong
        total22 = row['Unnamed: 22'] if 'Unnamed: 22' in row else None # Tong luong
        #total23 = row['Unnamed: 23'] if 'Unnamed: 23' in row else None # Tong luong
        total24 = row['Unnamed: 24'] if 'Unnamed: 24' in row else None # Tong luong
        #total25 = row['Unnamed: 25'] if 'Unnamed: 25' in row else None # Tong luong
        total26 = row['Unnamed: 26'] if 'Unnamed: 26' in row else None # Tong luong
        total27 = row['Unnamed: 27'] if 'Unnamed: 27' in row else None # Tong luong
        #total28 = row['Unnamed: 28'] if 'Unnamed: 28' in row else None # Tong luong
        total29 = row['Unnamed: 29'] if 'Unnamed: 29' in row else None # Tong luong
        total30 = row['Unnamed: 30'] if 'Unnamed: 30' in row else None # Tong luong
        total31 = row['Unnamed: 31'] if 'Unnamed: 31' in row else None # Tong luong
        total32 = row['Unnamed: 32'] if 'Unnamed: 32' in row else None # Tong luong
        #total33 = row['Unnamed: 33'] if 'Unnamed: 33' in row else None # Tong luong
        #total34 = row['Unnamed: 34'] if 'Unnamed: 34' in row else None # Tong luong
        #total35 = row['Unnamed: 35'] if 'Unnamed: 35' in row else None # Tong luong
        #total36 = row['Unnamed: 36'] if 'Unnamed: 36' in row else None # Tong luong
        total37 = row['Unnamed: 37'] if 'Unnamed: 37' in row else None # Tong luong
        total38 = row['Unnamed: 38'] if 'Unnamed: 38' in row else None # Tong luong
        total39 = row['Unnamed: 39'] if 'Unnamed: 39' in row else None # Tong luong
        total40 = row['Unnamed: 40'] if 'Unnamed: 40' in row else None # Tong luong
        total41 = row['Unnamed: 41'] if 'Unnamed: 41' in row else None # Tong luong
        #total42 = row['Unnamed: 42'] if 'Unnamed: 42' in row else None # Tong luong
        #total43 = row['Unnamed: 43'] if 'Unnamed: 43' in row else None # Tong luong
        total44 = row['Unnamed: 44'] if 'Unnamed: 44' in row else None # Tong luong
        #total45 = row['Unnamed: 45'] if 'Unnamed: 45' in row else None # Tong luong
        total46 = row['Unnamed: 46'] if 'Unnamed: 46' in row else None # Tong luong
        total47 = row['Unnamed: 47'] if 'Unnamed: 47' in row else None # Tong luong

        # Skip if Salary is null
        if pd.isna(total47) or total47 == 0:
            print(f"Skipping row {index} due to null or zero Salary.")
            continue
        
        # Generate content

        pdf.cell(200, 10, txt=f"PHIẾU LƯƠNG", ln=True, align='C')
        pdf.cell(200, 10, txt=f"Họ và Tên / Mã Nhân Viên : {employee_name} / {staff_id} ", ln=True, align='C')
        pdf.cell(200, 10, txt=f"Đơn vị:VNĐ", ln=True, align='R')


        pdf.set_font("DejaVu", size=8)
        pdf.cell(100, 5, txt=f"1 Lương cơ bản : {int(total06):,}", align='L')
        pdf.cell(100, 5, txt=f"19 Làm HC ngày lễ (300%): {int(total24):,}", align='R',ln=True)

        pdf.cell(100, 5, txt=f"2 Ngoại Ngữ : {int(total07):,}",  align='L')
        pdf.cell(100, 5, txt=f"20 Tổng thêm giờ  : {int(total26):,}", ln=True, align='R')
        
        pdf.cell(100, 5, txt=f"3 P.C Chủ quản : {int(total08):,}", align='L')
        pdf.cell(100, 5, txt=f"21 Thu nhập ngoài giờ : {int(total27):,}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"4 P.C Kỹ thuật : {int(total09):,}", align='L')
        pdf.cell(100, 5, txt=f"22 Phụ cấp khác : {int(total29):,}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"5 P.C Trình độ : {int(total10):,}", align='L')
        pdf.cell(100, 5, txt=f"23 Chuyên cần :{int(total17):,} ", ln=True, align='R')

        pdf.cell(100, 5, txt=f"6 Tổng (=1+2+3+4+5) : {int(total11):,}", align='L')
        pdf.cell(100, 5, txt=f"24 Phụ cấp hàn : {int(total30):,}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"7 Mức lương tạm ngừng việc = Lương vùng) :", align='L')
        pdf.cell(100, 5, txt=f"25 Phụ cấp giao thông :{int(total31):,}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"8 Ngày công (HC) (100%)) {int(total12):,}:", align='L')
        pdf.cell(100, 5, txt=f"26 Phụ cấp con nhỏ :{int(total32):,}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"9 Ngày công ca 3 (190%)) :{int(total13):,}",align='L')
        pdf.cell(100, 5, txt=f"27 Thu nhập khác trước thuế  :{int(total37):,}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"10 Tổng ngày công (10=8+9)) :{int(total14):,}", align='L')
        pdf.cell(100, 5, txt=f"28 Thưởng lễ ( nếu có ) :{total38}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"11 Ngày nghỉ tạm ngừng việc :",  align='L')
        pdf.cell(100, 5, txt=f"29 Tổng thu nhập (=14+21+22+23+24+25+26+27+28) :{int(total39):,}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"12 Lương đi làm (=(6/26)*10) :{int(total18):,}", align='L')
        pdf.cell(100, 5, txt=f"30 BHYT, BHXH, BHTN (6*10,5%) :{int(total40):,}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"13 Lương tạm ngừng việc (13=(7/26)*11) :", align='L')
        pdf.cell(100, 5, txt=f"31 Phí công đoàn  :{int(total41):,}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"14 Tổng lương (=12+13) : {int(total18):,}", align='L')
        pdf.cell(100, 5, txt=f"32 Thuế TNCN : {int(total44):,}", ln=True, align='R')    

        pdf.cell(100, 5, txt=f"15 Lương 1 giờ công(=6/26 ngày*8h):  {int(total19):,}", align='L')
        pdf.cell(100, 5, txt=f"33 Tổng các khoản giảm trừ (=30+31+32): {int((total40)+(total41)+(total44)):,} ", ln=True, align='R')

        pdf.cell(100, 5, txt=f"16 Làm thêm ngày thường (150%): {int(total20):,}", align='L')
        pdf.cell(100, 5, txt=f"34 Tạm ứng (nếu có): {int(total46):,}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"17 Làm thêm ngày CN (200%): {int(total22):,}" ,align='L')
        pdf.cell(100, 5, txt=f"35 Thực lĩnh (=29-33-34): {int(total47):,}", ln=True, align='R')

        pdf.cell(100, 5, txt=f"18 Làm thêm giờ ca 3 (260%): {int(total21):,}", align='L')
        pdf.cell(100, 5, txt=f"36 Số ngày phép còn lại:", ln=True, align='R')

        pdf.set_font("DejaVu", size=12)
        pdf.cell(200, 20, txt=f"Cảm ơn bạn đã luôn số gắng,nỗ lực suốt thời gian qua!", ln=True, align='C')

        # Sanitize file name
        sanitized_name = "".join(char for char in str(staff_id)if char.isalnum() or char in " _-")
        file_path = f"salary_slips/{sanitized_name}_Salary_Slip.pdf"

        # Save the PDF
        pdf.output(file_path)
        print(f"Saved: {file_path}")
    except Exception as e:
        print(f"Error processing row {index}: {e}")
