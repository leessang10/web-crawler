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
table = load_ws['E742':'T2215']

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
    time.sleep(3)
    driver.find_element_by_xpath("//*[@id=\"srchListDiv\"]/div/div/div[1]/table/tbody/tr/td[1]/a").click()

    # 3. 기업재무재표에서 자산, 국내매출액, 수출액, R&D비용 구하기
    # 기업재무 클릭
    driver.find_element_by_xpath("//*[@id=\"side\"]/div[1]/ul/li[3]/ul/li[3]/a").click()
    # 개별재무제표 클릭
    driver.find_element_by_xpath("//*[@id=\"side\"]/div[1]/ul/li[3]/ul/li[3]/ul/li[1]/a").click()
    try:
        # 2015~2017 자산 정보 구하기
        try:
            a = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[1]").text
            b = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[2]").text
            c = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[3]").text
            print("\t자산 2017: ", a, "2018: ", b, "2019: ", c)

            # 자산 내역 날짜가 2017-12-31인 경우 엑셀에 해당 자산 값 기입
            if driver.find_element_by_xpath(
                    "//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[2]").text == "2017-12-31":
                row[3].value = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[1]").text
            # 자산 내역 날짜가 2018-12-31인 경우 엑셀에 해당 자산 값 기입
            if driver.find_element_by_xpath(
                    "//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[3]").text == "2018-12-31":
                row[4].value = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[2]").text
            # 자산 내역 날짜가 2019-12-31인 경우 엑셀에 해당 자산 값 기입
            if driver.find_element_by_xpath(
                    "//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[4]").text == "2019-12-31":
                row[5].value = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[3]").text

        except:
            print("X: 2017, 2018, 2019년도 자산 데이터 없음")
            pass

        # 손익계산서 클릭
        try:
            driver.find_element_by_xpath("//*[@id=\"tab_5dep\"]/li[2]/a").click()
            time.sleep(3)
        except:
            print("//*[@id=\"tab_5dep\"]/li[2]/a Not Found")

        # 2017, 2018, 2019년 국내매출 구하기
        try:
            a = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[3]/td[1]").text
            b = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[3]/td[2]").text
            c = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[3]/td[3]").text
            print("\t국내매출 2017: ", a, "2018: ", b, "2019: ", c)
            row[8].value = a
            row[9].value = b
            row[10].value = c
        except:
            print("X: 2017, 2018, 2019년도 국내매출 데이터 없음")
            pass

        # 2017, 2018, 2019년 수출매출 구하기
        try:
            if driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[4]/th").text == "수출매출":
                e = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[7]/td[1]").text
                f = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[7]/td[2]").text
                g = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[7]/td[3]").text
                print("\t수출매출 2017: ", e, "2018: ", f, "2019: ", g)

                # 자산 내역 날짜가 2015-12-31인 경우 엑셀에 해당 자산 값 기입
                row[13].value = e
                row[14].value = f
                row[15].value = g
        except:
            print("X: 2017, 2018, 2019년도 수출매출 데이터 없음")
            pass

        # 결산일자-2017-12-31 클릭
        try:
            # 재무상태표 클릭
            driver.find_element_by_xpath("//*[@id=\"tab_5dep\"]/li[1]/a").click()
            time.sleep(3)
            # 결산일자 옵션 클릭
            driver.find_element_by_xpath("//*[@id=\"acctDt\"]").click()
            driver.find_element_by_xpath("//*[@id=\"acctDt\"]/option[3]").click()
            # 조회 버튼 클릭
            driver.find_element_by_xpath("//*[@id=\"frmENFNS01R0\"]/div[3]/div/table/tbody/tr[2]/td[3]/input").click()
            time.sleep(3)
        except:
            print("X: 결산일자-2017-12-31 클릭")

        # 2015, 2016년 자산 데이터 구하기
        try:
            a = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[1]").text
            b = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[1]/td[2]").text
            print("\t자산 2015: ", a, "2016: ", b)
            # 자산 내역 날짜가 2015-12-31인 경우 엑셀에 해당 자산 값 기입
            if driver.find_element_by_xpath(
                    "//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[2]").text == "2015-12-31":
                row[1].value = a
            # 자산 내역 날짜가 2016-12-31인 경우 엑셀에 해당 자산 값 기입
            if driver.find_element_by_xpath(
                    "//*[@id=\"ENFNS01S0_TABLE\"]/table/thead/tr[1]/th[3]").text == "2016-12-31":
                row[2].value = b
        except:
            print("X: 2015, 2016년도 자산 데이터 없음")
            pass
        # 손익계산서 클릭
        try:
            driver.find_element_by_xpath("//*[@id=\"tab_5dep\"]/li[2]/a").click()
            time.sleep(3)
        except:
            print("//*[@id=\"tab_5dep\"]/li[2]/a Not Found")

        # 2015, 2016년 국내매출 구하기
        try:

            a = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[3]/td[1]").text
            b = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[3]/td[2]").text
            print("\t국내매출 2015: ", a, "2016: ", b)
            # 자산 내역 날짜가 2015-12-31인 경우 엑셀에 해당 자산 값 기입
            row[6].value = a
            row[7].value = b
        except:
            print("X: 2015, 2016년도 국내매출 데이터 없음")
            pass

        # 2015, 2016년 수출매출 구하기
        try:
            if driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[4]/th").text == "수출매출":
                e = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[7]/td[1]").text
                f = driver.find_element_by_xpath("//*[@id=\"ENFNS01S0_TABLE\"]/table/tbody/tr[7]/td[2]").text
                print("\t수출매출_2017: ", e, "매출_2018: ", f)
                # 자산 내역 날짜가 2015-12-31인 경우 엑셀에 해당 자산 값 기입
                row[11].value = e
                row[12].value = f
        except:
            print("X: 2015, 2016년도  수출매출 데이터 없음")
            pass
    except:
        print("!!!!")
    finally:
        print("\t변경 사항 저장...")
        load_wb.save("C:\\Users\\dawin07\\Documents\\Cretop_Data_Crawler\\data.xlsx")
        print("\t변경사항 저장 완료.")