# 다음영화 - 예매순위 페이지(http://ticket2.movie.daum.net/Movie/MovieRankList.aspx)에서 영화별 상세페이지에 접속하여
# 영화의 제목 / 장르 / 누적관객수 데이터를 수집합니다.
# 예매순위 페이지에서 각 영화 상세페이지로 들어갈 수 있는 링크를 찾습니다.
# 상세페이지에서 원하는 데이터(제목, 장르, 누적관객수)를 찾을 수 있는 선택자를 찾습니다.
# 누적관객수가 5만 이상인 영화의 데이터만 엑셀에 저장합니다.

# 컨테이너 div.movie_join ul li
# 상세페이지 link = div.movie_join ul li a / link.attr["href"]

# 제목 div.subject_movie strong
# 장르 dl dd.txt_main:nth-of-type(1)

# 누적관객수 dl.list_placing dd#totalAudience

import requests
#from bs4 import BeautifulSoup #설치가 안 되지만,, 어챠ㅏ피 엑셀 부분만 참고할 거니까 필요없음
import pyexcel

raw = requests.get("http://ticket2.movie.daum.net/Movie/MovieRankList.aspx"
                   , headers={"User-Agent":"Mozilla/5.0"})
html = BeautifulSoup(raw.text, "html.parser")
wb = pyexcel.Workbook()
sheet = wb.active
sheet.append(["제목", "장르", "관객 수"])  # <--맨 첫 행에 추가

movies = html.select("div.movie_join ul li") #컨테이너

for m in movies:
    link = m.select_one("div.movie_join ul li a")
    url = link.attrs["href"]

    each_raw = requests.get(url, headers={"User-Agent":"Mozilla/5.0"})
    each_html = BeautifulSoup(each_raw.text, "html.parser")

    title = each_html.select_one("div.subject_movie strong").text
    genre = each_html.select_one("dl dd:nth-of-type(1)").text
    audience = each_html.select_one("dl.list_placing dd#totalAudience").text
    audience.replace("명", "") #"명"이라는 문자열 삭제
    audience.replace(",", "")  # ","이라는 문자열 삭제

    if ( int(audience)<50000 ):
        continue
    else:
        sheet.append([title, genre, audience])

wb.save("daumnews_week5.xlsx")
