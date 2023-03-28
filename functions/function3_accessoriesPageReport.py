from selenium.webdriver.common.by import By
import functions.function3_1

class BasePage(object):
    def __init__(self, driver):
        self.driver = driver

# 3번 산출물: 악세서리 페이지 리포트
#메인 function 클래스
class MainFunction(BasePage):
    def acc_page_report(self):
        self.driver.maximize_window()
        mainFunction = functions.function3_1.SectionFunction(self.driver)
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
            if idx < len(acc_image) and idx > 0 :
                (color, counts, click_count) = mainFunction.product_acc_item_scr_img_alt(product_counts,idx,acc_product_count,product_option,product_color,counts,click_count,color,list_target,list_route,all_counts,self.driver,img,all_images)
            if idx == len(acc_image)-1:
                bann = mainFunction.bottom_banner(self.driver, list_target,list_route,all_counts,bann,all_images)
        mainFunction.make_file(list_number,list_target,list_route,all_images, all_counts)