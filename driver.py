
import openpyxl

### 수험자 리스트 엑셀 열기
wb = openpyxl.load_workbook("./data/자격시험수험자리스트_1102.xlsx")
originsheet = wb.active

MAX_ROW = 0

# 이메일 셀 범위 복사 붙여넣기 (C열-> L열로 복사)
columns = originsheet.iter_cols(min_col=3, max_col=3)
for col in columns:
    for cell in col:
        cell_new = originsheet.cell(row=cell.row, column=12, value=cell.value)
        MAX_ROW += 1
MAX_ROW -= 1  # 총 row 개수는 (MAX_ROW-1) 개

# 날짜 삽입 (A열-> M열)
dateValue = (originsheet.cell(row=2, column=1).value) % 10000  ## 1102
month = (int)(dateValue / 100)
day = (dateValue % 100)
dateString = str(month)+"월 "+str(day)+"일"  # 월-일 형태의 스트링으로 변환

for i in range(0,MAX_ROW):
    originsheet.cell(row=i+2, column=13).value = dateString  # M열에 삽입

# step 3. Selenium으로 브라우저 원격 제어

# step 4. 엑셀에 정보 붙여넣기
#sheet.append([title, genre, audience])

wb.save("./result/temp_1102.xlsx")

"""
# 나중에) 메일 전송용 엑셀 만들기
wb_mail = openpyxl.Workbook()
mailsheet = wb_mail.active
mailsheet.append(["수신자 Email 주소", "date", "link"])  # <--맨 첫 행에 추가
wb_mail.save("./result/MASS_INPUT_FORM_today.xlsx")
"""
