
import openpyexcel

# step 1. 메일 전송용 엑셀 만들기
wb_mail = openpyexcel.Workbook()
mailsheet = wb_mail.active
mailsheet.append(["수신자 Email 주소", "date", "link"])  # <--맨 첫 행에 추가

# 수험자 리스트 엑셀 열기 (C열-> L열로 복사 / A열-> M열로 복사 / N열에 크롤링 정보 넣을 예정)

wb = openpyexcel.load_workbook("Book.xlsx")

# step 2. 엑셀 중복값 표시
# 나중에

# step 3. Selenium으로 브라우저 원격 제어

# step 4. 엑셀에 정보 붙여넣기
#sheet.append([title, genre, audience])

wb.save("MASS_INPUT_FORM_today.xlsx")
