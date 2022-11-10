
import openpyxl

# step 1. 메일 전송용 엑셀 만들기
wb_mail = openpyxl.Workbook()
mailsheet = wb_mail.active
mailsheet.append(["수신자 Email 주소", "date", "link"])  # <--맨 첫 행에 추가

# 수험자 리스트 엑셀 열기 (C열-> L열로 복사 / A열-> M열로 복사 / N열에 크롤링 정보 넣을 예정)
wb = openpyxl.load_workbook("./data/자격시험수험자리스트_1101.xlsx")
originsheet = wb.active
old_cell = originsheet.cell(1,1)
new_cell = mailsheet.cell(2,1, value=old_cell.value)

# mailsheet[2:10,1] = originsheet[2:10,3] # 안 됨

""" 
#이렇게 해야 함
allList = []
for row in sheet.iter_rows(min_row=1, max_row=10, min_col=2, max_col=5):
    a = []
    for cell in row:
        a.append(cell.value)
    allList.append(a)
"""

# step 2. 엑셀 중복값 표시
# 나중에

# step 3. Selenium으로 브라우저 원격 제어

# step 4. 엑셀에 정보 붙여넣기
#sheet.append([title, genre, audience])

wb_mail.save("./result/MASS_INPUT_FORM_today.xlsx")
