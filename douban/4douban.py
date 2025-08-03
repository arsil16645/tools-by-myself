from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from random import randint
from random import uniform
import time
import cv2
import numpy
import pyautogui
import os

class CalculateDistance:
    def __init__(self,backgroud_path,slide_path,offset_px,offset_py,display):
        self.backgroud_img=cv2.imread(backgroud_path)
        self.offset_px=offset_px
        self.offset_py=offset_py
        self.slide_img=cv2.imread(slide_path,cv2.IMREAD_UNCHANGED)
        scale_x=50/self.slide_img.shape[1]
        self.slide_scale_img=cv2.resize(self.slide_img,(0,0),fx=scale_x,fy=scale_x)
        self.backgroud_cut_img=None
        self.display=display

    def get_distance(self):
        slide_grey_img=cv2.cvtColor(self.slide_scale_img,cv2.COLOR_BGR2GRAY)
        slide_edge_img=cv2.Canny(slide_grey_img,100,250)
        background_grey_img=cv2.cvtColor(self.backgroud_cut_img,cv2.COLOR_BGR2GRAY)
        background_edge_img=cv2.Canny(background_grey_img,200,300)
        h,w=slide_edge_img.shape

        result=cv2.matchTemplate(background_edge_img,slide_edge_img,cv2.TM_CCOEFF_NORMED)
        min_val,max_val,min_loc,max_loc=cv2.minMaxLoc(result)
        top_left=(max_loc[0],max_loc[0])
        bottom_right=(top_left[0]+w,top_left[1]+h)

        if self.display:
            print(top_left)
            print(bottom_right)
            after_img=cv2.rectangle(self.backgroud_cut_img,top_left,bottom_right,(0,0,255),2)
            self.cv_show('after',after_img)

        slide_distance=top_left[0]+w+10
        return slide_distance

    def cut_background(self):
        height=self.slide_scale_img.shape[0]
        self.backgroud_cut_img=self.backgroud_img[self.offset_py-10:self.offset_py+height+10,
                               self.offset_px+height+10:]

    def cv_show(self,name,img):
        cv2.imshow(name,img)
        cv2.waitKey(0)
        cv2.destroyAllWindows()

    def run(self):
        self.cut_background()
        return self.get_distance()

#确认登录按钮位置并点击
def press_button(pic_addr):
    loc=pyautogui.locateOnScreen(pic_addr,confidence=0.9)
    p=pyautogui.center(loc)
    pyautogui.moveTo(p)
    pyautogui.leftClick()

def handle_distance(distance):
    #将移动的距离区别在(-2,15)，将每次预计移动的距离放入列表
    import random
    slow_distance=[]
    while sum(slow_distance)<=distance:
        slow_distance.append(random.randint(-2,15))
    if sum(slow_distance)!=distance:
        slow_distance.append(distance-sum(slow_distance))
    return slow_distance

#移动滑块按钮
def drag_slide(tracks,slide_addr):
    time.sleep(5)
    loc=pyautogui.locateOnScreen(slide_addr,confidence=0.7)
    p1=pyautogui.center(loc)
    pyautogui.moveTo(p1)
    pyautogui.mouseDown()
    for track in tracks:
        pyautogui.move(track,uniform(-2,2),duration=0.15)
    pyautogui.mouseUp()

def login(url,username,password,path):
    #以debugging模式控制浏览器，需要手动启动谷歌浏览器，避免被检测
    servic=Service(r'C:\Users\Administrator\Downloads\chromedriver-win64 (1)\chromedriver-win64\chromedriver.exe')
    opt=Options()
    opt.debugger_address="127.0.0.1:9999"
    browser=webdriver.Chrome(service=servic,options=opt)
    browser.get(url)
    browser.implicitly_wait(4)

    #登录密码
    iframe=browser.find_element(By.XPATH,'//*[@id="anony-reg-new"]/div/div[1]/iframe')
    browser.switch_to.frame(iframe)
    browser.find_element(By.XPATH,'/html/body/div[1]/div[1]/ul[1]/li[2]').click()
    time.sleep(uniform(1,4))
    #输入用户名和密码
    browser.find_element(By.XPATH,'//*[@id="username"]').send_keys(username)
    time.sleep(uniform(1,4))
    browser.find_element(By.XPATH,'//*[@id="password"]').send_keys(password)
    time.sleep(uniform(1, 4))
    #点击登录
    pic_addr = r"C:\Users\Administrator\Desktop\pachong\douban\login.png"
    press_button(pic_addr)
    #进入验证码
    WebDriverWait(browser,4).until(EC.visibility_of_element_located((By.ID,'tcaptcha_iframe_dy')))
    slide_frame=browser.find_element(By.ID,"tcaptcha_iframe_dy")
    browser.switch_to.frame(slide_frame)
    time.sleep(2)

    #获取背景图片
    backgroud_element=browser.find_element(By.ID,'slideBg')
    backgroud_location=backgroud_element.location
    print(backgroud_location)
    backgroud_img=backgroud_element.screenshot_as_png
    filename=int(time.time())
    with open(f'{path}/bg{filename}.png','wb')as f:
        f.write(backgroud_img)
        print("已下载背景图片")
    #获取滑块图片
    slide_element1=browser.find_element(By.XPATH,'//*[@id="tcOperation"]/div[7]')
    slide_element2 = browser.find_element(By.XPATH, '//*[@id="tcOperation"]/div[8]')
    s1=slide_element1.size
    s2=slide_element2.size
    print(s1,s2)
    #根据像素，判断哪个才是真的滑块图片
    if s1['width']>100 and s1['height']<20:
        slide_element=slide_element2
    else:
        slide_element=slide_element1
    slide_location = slide_element.location
    print(slide_location)
    slide_img = slide_element.screenshot_as_png
    with open(f'{path}/sd{filename}.png','wb')as f:
        f.write(slide_img)
        print("已下载滑块图片")
    
    bg_addr=f'{path}/bg{filename}.png'
    sd_addr=f'{path}/sd{filename}.png'
    
    #计算滑块所在背景图片的相对位置
    offset_x=slide_location['x']-backgroud_location['x']
    offset_y = slide_location['y'] - backgroud_location['y']
    
    #计算出需要移动的距离
    slide_offset=CalculateDistance(bg_addr,sd_addr,offset_x,offset_y,0)
    slide_distance=slide_offset.run()
    print(slide_distance)
    
    #设计滑动轨迹的快慢和距离
    tracks=handle_distance(slide_distance)
    slide_img_addr=f'{path}/douban_slide.png'
    #开始移动
    drag_slide(tracks,slide_img_addr)

if __name__=="__main__":
#路径要英文，在浏览器属性的目标后面加 --remote-debugging-port=9999 --user-data-dir="C:\Temp\ChromeDebug"；
    url='https://www.douban.com/'
    username = '1415627886@qq.com'
    password = 'aodfha'
    path='C:/Users/Administrator/Desktop/pachong/douban'
    login(url,username,password,path)

