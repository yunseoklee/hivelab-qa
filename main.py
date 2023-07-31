################################
# 터미널 열고 python -m unittest -v main.py 입력하여 실행
################################

# unittest 모듈 불러오기
import unittest
# webdriver 모듈을 selenium 패키지에서 불러오기
from selenium import webdriver
# Options 클래스를 selenium.webdriver.chrome.options 모듈에서 불러오기
from selenium.webdriver.chrome.options import Options
# 엑셀 사용을 위한 라이브러리
from openpyxl import load_workbook, Workbook
# By 사용을 위한 라이브러리
from selenium.webdriver.common.by import By
# 키입력을 위한 라이브러리
from selenium.webdriver.common.keys import Keys

# 커스텀 모듈인 functions 불러오기
import functions.function1_copyApplied, functions.function2_highlightsPageReport
import functions.function3_accessoriesPageReport, functions.function4_disclaimerReport
import functions.function5_taggingReport, functions.function6_anchorListReport
import functions.function7_imageSrcCheck, functions.function8_sectionReport

class MainTest(unittest.TestCase):
    
    # 테스트를 위한 설정
    def setUp(self) -> None:
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        
        # 창 열지 않고 실행 원할 시 활성화
        # options.add_argument("headless")

        self.driver = webdriver.Chrome(options=options)
        url = ("https://semdev.hivelab.co.kr:3443/global/galaxy/galaxy-s23/gallery/")
        self.driver.get(url)

        #로그인 설정
        # Dev 이용 시 활성화
        self.driver.find_element(By.NAME, "user_id").send_keys("test")
        self.driver.find_element(By.NAME, "user_pw").send_keys("test")
        self.driver.find_element(By.CLASS_NAME, "btn").send_keys(Keys.ENTER)
        # P6 이용 시 활성화
        # self.driver.find_element(By.NAME, "j_username").send_keys("qauser02")
        # self.driver.find_element(By.NAME, "j_password").send_keys("samsungqa")
        # self.driver.find_element(By.CLASS_NAME, "coral3-Button").send_keys(Keys.ENTER)

    # 테스트 종료 후
    def tearDown(self) -> None:
        self.driver.quit()

    # 1번 산출물: 카피덱 반영 확인
    def _test_copydeck_applied(self):
        mainFunction = functions.function1_copyApplied.MainFunction(self.driver)
        excelfile = 'Copydeck_Result.xlsx'
        sheetname = 'Sheet1'
        mainFunction.copydeck_applied(excelfile, sheetname)
    
    # 2번 산출물: 하이라이트 페이지 리포트
    def _test_page_report(self):
        mainFunction = functions.function2_highlightsPageReport.MainFunction(self.driver)
        excelfile = 'Page_Report.xlsx'
        sheetname = 'Sheet1'
        mainFunction.page_report(excelfile, sheetname, 2)
    
    # 3번 산출물: 악세서리 페이지 리포트 
    # ※※※악세서리 페이지 리포트 추출 하기 전 url 변경 필요※※※
    def _test_acc_report(self):
        mainFunction = functions.function3_accessoriesPageReport.MainFunction(self.driver)
        mainFunction.acc_page_report()
    
    # 4번 산출물: 각주 자동화 리포트
    # 4-1번: 최하단 디스클레이머 카피덱 반영 확인
    def _test_bottomDisclaimer_copy_applied(self):
        disclaimer = functions.function4_disclaimerReport.MainFunction(self.driver)
        excelfile = 'Disclaimer_Result.xlsx'
        sheetname = 'Sheet1'
        disclaimer.copydeck_disclaimerText_applied(excelfile, sheetname, 2)

    # 4-2번: 본문 각주 번호 클릭하려 최하단 디스클레이머 영역으로 이동 확인
    def _test_disclaimerNumber_click(self):
        disclaimer = functions.function4_disclaimerReport.MainFunction(self.driver)
        excelfile = 'Disclaimer_Result.xlsx'
        sheetname = 'Sheet1'
        option = '2'
        disclaimer.body_disclaimerNumber_click(excelfile, sheetname, option, 2)

    # 4-3번: 55개국 최하단 각주번호 순서 확인
    def _test_bottomDisclaimerNumber_order_check(self):
        disclaimer = functions.function4_disclaimerReport.MainFunction(self.driver)
        fronturl = 'https://www.samsung.com/'
        backurl = '/smartphones/galaxy-s23-ultra/'
        excelfile = 'Disclaimer_Result.xlsx'
        sheetname = 'Sheet2'
        disclaimer.bottomDisclaimerNumber_order_check(fronturl, backurl, excelfile, sheetname, 1.5)
    
    # 5번 산출물: 태깅 자동화 리포트
    def _test_tagging_automation(self):
        tagging = functions.function5_taggingReport.MainFunction(self.driver)
        excelfile = 'Tagging_Result.xlsx'
        sheetname = 'Sheet1'
        tagging.tagging_automation_report(excelfile, sheetname, 1)
    
    # 6번 산출물: 앵커 스크린샷 리포트
    def _test_anchor_screenshot_automation(self):
        anchor = functions.function6_anchorListReport.MainFunction(self.driver)
        excelfile = 'anchorList.xlsx'
        sheetname = 'test'
        anchor.anchorList(excelfile, sheetname, 1)

    # 7번 산출물: 이미지 경로 유효성 체크
    def _test_image_src_check(self):
        image_src = functions.function7_imageSrcCheck.MainFunction(self.driver)
        image_src.imageSrcCheck()

    # 8번 산출물: 피쳐, 배너 , CTA 유무 및 url 체크
    def test_section_check(self):
        section = functions.function8_sectionReport.MainFunction(self.driver)
        excelfile = 'url_LIVE.xlsx'
        sheetname = 'Sheet1'
        section.sectioncheck(excelfile, sheetname)


if __name__ == "__main__":
    unittest.main()