# unittest 모듈 불러오기
import unittest
# webdriver 모듈을 selenium 패키지에서 불러오기
from selenium import webdriver
# Options 클래스를 selenium.webdriver.chrome.options 모듈에서 불러오기
from selenium.webdriver.chrome.options import Options
# ChromeDriverManager 클래스를 webdriver_manager.chrome 모듈에서 불러오기
from webdriver_manager.chrome import ChromeDriverManager
# 엑셀 사용을 위한 라이브러리
from openpyxl import load_workbook, Workbook

class MainTest(unittest.TestCase):
    # 테스트를 위한 설정
    def setUp(self) -> None:
        #드라이버 설정
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        self.driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
        self.driver.get("https://www.samsung.com/uk/smartphones/galaxy-s23-ultra/")
    
    def example1(self):
        print("hello world")

    # 테스트 종료 후 드라이버 닫기
    def tearDown(self):
        self.driver.close()

        
if __name__ == "__main__":
    unittest.main()