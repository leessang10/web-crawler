import requests
from bs4 import BeautifulSoup
from selenium import webdriver
import time
from openpyxl import load_workbook

# 엑셀 불러오기
# data_only=True로 해줘야 수식이 아닌 값으로 받아온다.
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.select import Select
load_wb = load_workbook("C:\\Users\\dawin07\\Documents\\Cretop_Data_Crawler\\data.xlsx", data_only=True)
# 시트 이름으로 불러오기
load_ws = load_wb['최종본']
# 지정한 셀의 값 출력
table = load_ws['E178':'J178']

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
try:
    # # 동시접속자 알림창이 뜨면 "계속 진행하기" 버튼을 누른다
    driver.find_element_by_id("CMCOM07S2Login").click()
except:
    pass

n = 1
for row in table:
    company_registration_number = row[0].value
    print(n, ". 사업자등록번호: ", company_registration_number, sep="")
    n += 1
    # 2. 사업자등록번호 검색
    # 기업 검색란에 사업자등록번호 입력 후 엔터
    driver.find_element_by_xpath("//*[@id=\"_srchNm\"]").send_keys(company_registration_number)
    driver.find_element_by_xpath("//*[@id=\"uniSrch\"]").click()
    driver.implicitly_wait(3)

    try:
        # 검색 결과 창에서 기업 선택
        driver.find_element_by_xpath("//*[@id=\"srchListDiv\"]/div/div/div[1]/table/tbody/tr/td[1]/a").click()

        # 3. 기업재무재표에서 자산, 국내매출액, 수출액, R&D비용 구하기
        # 기업재무 클릭
        driver.find_element_by_xpath("//*[@id=\"side\"]/div[1]/ul/li[3]/ul/li[3]/a").click()
        # 개별재무제표 클릭
        driver.find_element_by_xpath("//*[@id=\"side\"]/div[1]/ul/li[3]/ul/li[3]/ul/li[1]/a").click()

        try:
            driver.find_element_by_xpath("//*[@id=\"acctDt\"]").click()
            driver.find_element_by_xpath("//*[@id=\"acctDt\"]/option[3]").click()
            driver.find_element_by_xpath("//*[@id=\"tab_5dep\"]/li[1]/a").click()
        except:
            print("//*[@id=\"acctDt\"] Not Found")
        # 2015~2017 자산 정보 구하기
        try:

            # 2017~2019 자산 정보 구하기
            # 자산 내역 날짜가 2017-12-31인 경우 엑셀에 해당 자산 값 기입
            if driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[2]").text == "2015-12-31":
                row[1].value = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[1]").text
            # 자산 내역 날짜가 2018-12-31인 경우 엑셀에 해당 자산 값 기입
            if driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[3]").text == "2016-12-31":
                row[2].value = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[2]").text
        except:
            print("Error 발생함. 데이터 없음")
            pass
        finally:
            load_wb.save("C:\\Users\\dawin07\\Documents\\Cretop_Data_Crawler\\data.xlsx")
    except:
        load_wb.save("C:\\Users\\dawin07\\Documents\\Cretop_Data_Crawler\\data.xlsx")
