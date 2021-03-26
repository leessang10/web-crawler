import time

from selenium import webdriver
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
# 엑셀 불러오기
# data_only=True로 해줘야 수식이 아닌 값으로 받아온다.
from selenium.webdriver.support.wait import WebDriverWait

load_wb = load_workbook("C:\\Users\\dawin07\\Documents\\Cretop_Data_Crawler\\data.xlsx", data_only=True)
# 시트 이름으로 불러오기
load_ws = load_wb['최종본']
# 지정한 셀의 값 출력
table = load_ws['E450':'N450']


# for row in get_cells:
#     for cell in row:
#         print(row)
driver = webdriver.Chrome("C:\\chromedriver.exe")
# 접속할 url
URL = "http://www.cretop.com/"
# 접속 시도
driver.get(URL)
# 창 최대화
# driver.maximize_window()
wait = WebDriverWait(driver, 20)

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
    try:
        driver.implicitly_wait(3)
        # 검색 결과 창에서 기업 선택
        driver.find_element_by_xpath("//*[@id=\"srchListDiv\"]/div/div/div[1]/table/tbody/tr/td[1]/a").click()

        # 3. 기업재무재표에서 자산, 국내매출액, 수출액, R&D비용 구하기
        # 기업재무 클릭
        driver.find_element_by_xpath("//*[@id=\"side\"]/div[1]/ul/li[3]/ul/li[3]/a").click()
        # 개별재무제표 클릭
        driver.find_element_by_xpath("//*[@id=\"side\"]/div[1]/ul/li[3]/ul/li[3]/ul/li[1]/a").click()
        driver.implicitly_wait(3)

        try:
            driver.find_element_by_xpath("//*[@id=\"acctDt\"]").click()
            driver.find_element_by_xpath("//*[@id=\"acctDt\"]/option[3]").click()
            driver.find_element_by_xpath("//*[@id=\"frmENFNS01R0\"]/div[3]/div/table/tbody/tr[2]/td[3]/input").click()
        except:
            print("2017년도 자산 데이터 없음")


        try:
            a = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[3]/td[1]").text
            b = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[3]/td[2]").text
            c = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[3]/td[3]").text
            print(a, b, c)
            # 자산 내역 날짜가 2015-12-31인 경우 엑셀에 해당 자산 값 기입
            if driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[2]").text == "2017-12-31":
                row[8].value = a

            # 자산 내역 날짜가 2016-12-31인 경우 엑셀에 해당 자산 값 기입
            if driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[3]").text == "2018-12-31":
                row[9].value = b

            if driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[4]").text == "2019-12-31":
                row[10].value = c

        except:
            print("Error 발생함. 데이터 없음")
            pass
        finally:
            print("변경 사항 저장")
            load_wb.save("C:\\Users\\dawin07\\Documents\\Cretop_Data_Crawler\\data.xlsx")
    except:
        print("//*[@id=\"srchListDiv\"]/div/div/div[1]/table/tbody/tr/td[1]/a Not Found")
        pass