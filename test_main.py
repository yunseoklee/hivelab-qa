# unittest 모듈 불러오기
import unittest
# webdriver 모듈을 selenium 패키지에서 불러오기
from selenium import webdriver
# Options 클래스를 selenium.webdriver.chrome.options 모듈에서 불러오기
from selenium.webdriver.chrome.options import Options
# 엑셀 사용을 위한 라이브러리
from openpyxl import load_workbook, Workbook

# 커스텀 모듈인 functions 불러오기
import functions

################################
# 터미널 열고 python -m unittest -v main.py 입력하여 실행
################################

class MainTest(unittest.TestCase):
    
    # 테스트를 위한 설정
    def setUp(self) -> None:
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        
        # 창 열지 않고 실행 원할 시 활성화
        # options.add_argument("headless")

        self.driver = webdriver.Chrome(options=options)
        # url = ("https://semdev.hivelab.co.kr:3443/global/galaxy/galaxy-s23-ultra/")
        url = ("https://semdev.hivelab.co.kr:3443/global/galaxy/galaxy-s23-ultra/accessories/")
        self.driver.get(url)

        #로그인 설정
        loginPage = functions.LoginPage(self.driver)
        # Dev 이용 시 활성화
        loginPage.dev_login("test", "test")
        # P6 이용 시 활성화
        #loginPage.p6_login("qauser02", "samsungqa")

    # 테스트 종료 후
    def tearDown(self) -> None:
        self.driver.quit()

    # 1번 산출물: 카피덱 반영 확인
    def _test_copydeck_applied(self):
        mainFunction = functions.MainFunction(self.driver)
        excelfile = 'Copydeck_Result(Format).xlsx'
        sheetname = 'Sheet1'
        mainFunction.copydeck_applied(excelfile, sheetname)
    
    # 2번 산출물: 하이라이트 페이지 리포트
    def _test_page_report(self):
        mainFunction = functions.MainFunction(self.driver)
        excelfile = 'Page_Report(Format).xlsx'
        sheetname = 'Sheet1'
        mainFunction.page_report(excelfile, sheetname, 2)

    # 3번 산출물: 악세서리 페이지 리포트
    def test_acc_report(self):
        from selenium.webdriver.common.by import By
        mainFunction = functions.MainFunction(self.driver)
        acc_image = self.driver.find_elements(By.CSS_SELECTOR,"picture")
        product_color = self.driver.find_elements(By.CSS_SELECTOR, "div.accessories__product-contents")
        acc_product_count = self.driver.find_elements(By.CSS_SELECTOR,"div.accessories__product")


        product_counts = []# 제품 구간
        for count in  acc_product_count:
            count = count.find_elements(By.CSS_SELECTOR,"li.accessories__product-item")
            product_counts.append(len(count))
        #product_counts = [9,3,2,3]

        product_option = []# 제품 구간 별 제품 옵션 구간
        for count in  acc_product_count:
            count = count.find_elements(By.CSS_SELECTOR,"ul.accessories__product-option-list")
            product_option.append(len(count))
        # product_option = [9, 0, 0, 3]

        color = 0
        counts = 0
        click_count = 0

        img = 0 #이미지 수 
        bann = 0

        all_counts = [] # 이미지가 들어가야 될 행 
        list_number = [] # 숫자
        list_target = [] # 대상
        list_route = [] # 값
        all_images = [] # 이미지 경로 리스트 
        
        for idx, acc in enumerate(acc_image):
            mainFunction.kv_visual_src_img(idx, acc, list_target, list_route, all_counts, self.driver, all_images)
            if idx < 5 and idx > 0 :
                (color, counts, click_count) = mainFunction.product_acc_item_scr_img_alt(product_counts,idx,acc_product_count,product_option,product_color,counts,click_count,color,list_target,list_route,all_counts,self.driver,img,all_images)
            if idx == 4:
                bann = mainFunction.bottom_banner(self.driver, list_target,list_route,all_counts,bann,all_images)
        mainFunction.make_file(list_number,list_target,list_route,all_images, all_counts)
    
    # 4번 산출물: 각주 자동화 리포트
    # 4-1번: 최하단 디스클레이머 카피덱 반영 확인
    def _test_bottomDisclaimer_copy_applied(self):
        disclaimer = functions.MainFunction(self.driver)
        excelfile = 'Disclaimer_Result.xlsx'
        sheetname = 'Sheet1'
        disclaimer.copydeck_disclaimerText_applied(excelfile, sheetname)
        
    # 4-2번: 본문 각주 번호 클릭하려 최하단 디스클레이머 영역으로 이동 확인
    def _test_disclaimerNumber_click(self):
        disclaimer = functions.MainFunction(self.driver)
        excelfile = 'Disclaimer_Result.xlsx'
        sheetname = 'Sheet1'
        option = '1'
        disclaimer.body_disclaimerNumber_click(excelfile, sheetname, option)
    
    # 4-3번: 55개국 최하단 각주번호 순서 확인
    def _test_bottomDisclaimerNumber_order_check(self):
        disclaimer = functions.MainFunction(self.driver)
        fronturl = 'https://www.samsung.com/'
        backurl = '/smartphones/galaxy-s23-ultra/'
        excelfile = 'Disclaimer_Result.xlsx'
        sheetname = 'Sheet2'
        disclaimer.bottomDisclaimerNumber_order_check(fronturl, backurl, excelfile, sheetname)

    # 5번 산출물: 태깅 자동화 리포트
    def _test_tagging_automation(self):
        tagging = functions.MainFunction(self.driver)
        excelfile = 'Tagging_Result.xlsx'
        sheetname = 'Sheet1'
        tagging.tagging_automation_report(excelfile, sheetname, 1)

if __name__ == "__main__":
    unittest.main()