"# web-crawler"

Q. 왜 만들었나?

A. 사무보조 알바를 할 때 반복적인 데이터 크롤링 업무를 자동화해서 노가다를 줄이고 싶었다.

Q. 어떤 작업이었나?

A. `크레탑`([http://www.cretop.com/](http://www.cretop.com/)) 이라는 기업정보 제공 플랫폼에서 `사업자등록번호`로 기업정보를 조회한다. 
해당 기업의 재무제표에서 2015~2017년의 `자산`, `국내매출`, `수출매출`, `연구비용`, `종업원수`에 대한 데이터를 엑셀로 가져오는 작업이다.

Q. 사용한 언어와 API는?

A. `Python`의 `Selenium`, `openpyxl`을 사용했다. 

1. 작업 순서 정리
    1. `Selenium`의 기본 셋팅
    2. `openpyxl`의 기본 셋팅
    3. `크레탑` 로그인
    4. 엑셀 상 `사업자등록번호` 리스트를 가져옴
        1. `사업자등록번호`로 기업 정보 검색
        2. 사이드바의 `기업정보` - `기업재무` 클릭
        3. 재무제표 상 `자산`에 대한 도표가 출력된다
        4. 재무제표 상 **2017년**, **2018년**, **2019년** `자산` 정보가 있는지 확인
            1. 있으면 `자산` 데이터를 엑셀에 기입
            2. 없으면 아무 작업 없이 건너뜀
        5. 엑셀에 변경사항 내역 저장

2. 수정해야할 사항들
    1. 2015년, 2016년 데이터 정보 가져오는 기능 추가
    2. `재산` 정보 이외에 `국내매출`, `수출매출`, `연구비용`, `종업원수`에 대한 정보를 가져오는 기능 추가
    3. 기업을 검색하고 데이터를 기입정보를 가져오는 작업이 끝날 때마다 매번 저장을 하는 것이 수행시간을 너무 잡아먹는다.
    4. 재무제표 상 **데이터가 없을 때 예외처리**를 추가.
    5. 예외처리를 구현한다면 `try catch`문을 사용해서 에러 발생시 **변경사항을 저장하고 종료**시키면 수행시간을 줄일 수 있을 것 같다.
