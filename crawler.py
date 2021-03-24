import requests
from bs4 import BeautifulSoup
from selenium import webdriver
import time
from openpyxl import load_workbook

# 엑셀 불러오기
# data_only=True로 해줘야 수식이 아닌 값으로 받아온다.
load_wb = load_workbook("C:\\Users\\dawin07\\Documents\\Cretop_Data_Crawler\\data.xlsx", data_only=True)
# 시트 이름으로 불러오기
load_ws = load_wb['최종본']
# 지정한 셀의 값 출력
table = load_ws['E123':'J222']

# for row in get_cells:
#     for cell in row:
#         print(row)
driver = webdriver.Chrome("C:\\chromedriver.exe")
# 접속할 url
url = "http://www.cretop.com/"
# 접속 시도
driver.get(url)
# 창 최대화
# driver.maximize_window()

# 1. 로그인
login = {
    "id": "compa999",
    "pw": "compa123$"
}
print("ID: ", login.get("id"))
print("PW: ", login.get("pw"))
# 아이디와 비밀번호를 입력합니다.
# driver.find_element_by_name('id').send_keys('아이디') # "아이디라는 값을 보내준다"
driver.find_element_by_id("in_id").send_keys(login.get("id"))
driver.find_element_by_id("in_pw").send_keys(login.get("pw"))
driver.find_element_by_id("loginBtn1").click()
driver.find_element_by_id("CMCOM07S2Login").click()


for row in table:
    company_registration_number = row[0].value
    print("사업자등록번호: ", company_registration_number)
    # 2. 사업자등록번호 검색
    # 기업 검색란에 사업자등록번호 입력 후 엔터
    driver.find_element_by_xpath("//*[@id=\"_srchNm\"]").send_keys(company_registration_number)
    driver.find_element_by_xpath("//*[@id=\"uniSrch\"]").click()
    driver.implicitly_wait(2)
    # 검색 결과 창에서 기업 선택
    driver.find_element_by_xpath("//*[@id=\"srchListDiv\"]/div/div/div[1]/table/tbody/tr/td[1]/a").click()

    # 3. 기업재무재표에서 자산, 국내매출액, 수출액, R&D비용 구하기
    # 기업재무 클릭
    driver.find_element_by_xpath("//*[@id=\"side\"]/div[1]/ul/li[3]/ul/li[3]/a").click()
    # 개별재무제표 클릭
    driver.find_element_by_xpath("//*[@id=\"side\"]/div[1]/ul/li[3]/ul/li[3]/ul/li[1]/a").click()
    # 2017, 2018, 2019 자산 구하기
    # 2017 항목이 있을 경우만
    if driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[2]").is_enabled():
        # 자산 내역 날짜가 2017-12-31인 경우 엑셀에 해당 자산 값 기입
        if driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[2]").text == "2017-12-31":
            row[3].value = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[1]").text
        # 자산 내역 날짜가 2018-12-31인 경우 엑셀에 해당 자산 값 기입
        if driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[3]").text == "2018-12-31":
            row[4].value = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[2]").text
        # 자산 내역 날짜가 2019-12-31인 경우 엑셀에 해당 자산 값 기입
        if driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[4]").text == "2019-12-31":
            row[5].value = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[3]").text
    load_wb.save("C:\\Users\\dawin07\\Documents\\Cretop_Data_Crawler\\data.xlsx")
