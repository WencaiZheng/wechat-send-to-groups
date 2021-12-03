# byNoven
import pyautogui
import time
import xlrd
import pyperclip
import win32clipboard as clip
import win32con
from io import BytesIO
from PIL import Image

### 读取待发布文本
with open('target_text.txt',encoding="utf8") as f:
    TEXT = f.read()
    # print(TEXT)


def pic2clip(filepath):
    """
    local picture to clipboard
    """
    image = Image.open(filepath)

    output = BytesIO()
    image.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()
    clip.OpenClipboard()
    clip.EmptyClipboard()
    clip.SetClipboardData(win32con.CF_DIB, data)
    clip.CloseClipboard()
    time.sleep(0.5)


def send2group(clickTimes,lOrR,img,content):
    obj_name = img.split('.')[0].split('/')[-1]
    while True:
        location=pyautogui.locateCenterOnScreen('_util/search_bar.png',confidence=0.9)
        if location is not None:
            if content==3:
                """
                转发链接
                """
                location=pyautogui.locateCenterOnScreen('_util/helper.png',confidence=0.9)
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                location=pyautogui.locateCenterOnScreen('target_link.png',confidence=0.9)
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button='right')
                location=pyautogui.locateCenterOnScreen('_util/retweet.png',confidence=0.9)
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                # paste to receiver list
                pyperclip.copy(obj_name)
                pyautogui.hotkey('ctrl','v') 
                try: 
                    # click the first one searched
                    location=pyautogui.locateCenterOnScreen('_util/down_list.png',confidence=0.9)
                    pyautogui.click(location.x,location.y+100,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                    # find the send button
                    location=pyautogui.locateCenterOnScreen('_util/link_send.png',confidence=0.9)
                    pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                except:
                    print(f'cannot find this group {obj_name}')
                
            
            elif content == 2:
                """
                发送文本和图片
                """
                # TEXT = ""
                # if is pic and TEXT
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                pyperclip.copy(obj_name)
                pyautogui.hotkey('ctrl','v')              
                time.sleep(1)
                pyautogui.click(location.x,location.y+150,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                # copy picture
                pic2clip('target_plot.png')
                pyautogui.hotkey('ctrl','v')
                # copy text
                pyperclip.copy(TEXT)
                pyautogui.hotkey('ctrl','v')
                # send the msg
                location=pyautogui.locateCenterOnScreen('_util/send.png',confidence=0.9)
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)

            elif content == 1:
                """
                仅发送图片
                """
                # TEXT = ""
                # if is pic and TEXT
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                pyperclip.copy(obj_name)
                pyautogui.hotkey('ctrl','v')              
                time.sleep(1)
                pyautogui.click(location.x,location.y+150,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                # copy picture
                pic2clip('target_plot.png')
                pyautogui.hotkey('ctrl','v')
                # send the msg
                location=pyautogui.locateCenterOnScreen('_util/send.png',confidence=0.9)
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            
            elif content == 0:
                """
                发送文本
                """
                # TEXT = ""
                # if is pic and TEXT
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                pyperclip.copy(obj_name)
                pyautogui.hotkey('ctrl','v')              
                time.sleep(1)
                pyautogui.click(location.x,location.y+150,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                # copy text
                pyperclip.copy(TEXT)
                pyautogui.hotkey('ctrl','v')
                # send the msg
                location=pyautogui.locateCenterOnScreen('_util/send.png',confidence=0.9)
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            # 
            break
        else:
            # 搜索
            print("未找到微信搜索栏,请打开微信界面,等待5秒")
            # pyautogui.scroll(-200)
            time.sleep(5)

#任务
def mainWork(img,content):
    i = 1
    while i < sheet1.nrows:
        #取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            send2group(1,"left",img,content=content)
            print("传送中",img)
        #
        i += 1
        time.sleep(1)

if __name__ == '__main__':
    file = 'clients.xls'
    #打开文件
    wb = xlrd.open_workbook(filename=file)
    # 通过索引获取表格sheet页,0为正式名单,1为测试名单
    sheet1 = wb.sheet_by_index(1)
    print('请打开微信界面等待')
    # 0 文字 1 图片 2 图+文 3 链接
    mainWork(sheet1,content = 0)
    print('正常退出')
