from PIL import Image as PILImage
from io import BytesIO
import openpyxl, os, time

class BasePage(object):
    def __init__(self, driver):
        self.driver = driver

#메인 function 클래스
class MainFunction(BasePage):
    # 6번 산출물: 앵커 스크린샷 리포트
    def anchorList(self, excelName, sheetName, waitTime):

        # 윈도우 사이즈 최대
        self.driver.maximize_window()

        # 엑셀 설정
        self.wb = openpyxl.load_workbook(filename=excelName)
        self.ws = self.wb[sheetName]

        # 스크린샷 저장을 위한 폴더 설정
        folder_name = "anchorList_Screenshot"
        current_directory = os.getcwd()
        new_folder_path = os.path.join(current_directory, folder_name)
        if not os.path.exists(new_folder_path):
            os.makedirs(new_folder_path)
        else:
            pass

        for i in range(2,200):
            link = self.ws['C'+str(i)].value
            if link is not None:
                self.driver.get(link)
                screenshot = self.driver.get_screenshot_as_png()
                time.sleep(waitTime)
                # 드라이버 이미지 스샷 찍기
                image = PILImage.open(BytesIO(screenshot))
                # 바이트로 변경 후 이미지 형식으로 오픈
                image.save(os.path.join(new_folder_path, self.ws['B'+str(i)].value + ".png"))
                # 이미지 파일을 이미지 저장 경로 폴더에 저장, 스샷 이름은 워크시트 B열을 가져옴 : 국가.png"
                time.sleep(waitTime)
