from selenium import webdriver
from PIL import Image as PILImage
from io import BytesIO
import openpyxl, os, time

driver = webdriver.Chrome()
driver.maximize_window()
workbook = openpyxl.load_workbook('anchorList.xlsx')
# 엑셀파일 실행
worksheet = workbook['test']
# worksheet = workbook[시트 이름] : 실행 시킨 엑셀 파일에서 지정한 시트 불러오기 
current_directory = os.getcwd()
# 현재 작업 디렉토리(current working directory)를 반환
if not os.path.exists("anchorList_Screenshot"):
    os.makedirs("anchorList_Screenshot")
# 스크린샷을 담을 폴더 생성
folder_name = "anchorList_Screenshot"
screenshot_path = os.path.join(current_directory,folder_name)
# 현제 작업 디렉토리와 생성한 폴더 디렉토리 경로 연결 : 이미지 경로 생성


# 6번 산출물: 엥커 리스트 스크린샷 리포트
#메인 function 클래스
class MainFunction():
    def anchorList(self):
        for i in range(2,5):
            link = worksheet['C'+str(i)].value
            driver.get(link)
            screenshot = driver.get_screenshot_as_png()
            # 드라이버 이미지 스샷 찍기
            image = PILImage.open(BytesIO(screenshot))
            # 바이트로 변경 후 이미지 형식으로 오픈
            image.save(os.path.join(screenshot_path , worksheet['B'+str(i)].value + ".png"))
            # 이미지 파일을 이미지 저장 경로 폴더에 저장, 스샷 이름은 워크시트 B열을 가져옴 : 국가.png"
            time.sleep(0.2)