# coding=UTF-8
import pyautogui
import time
import xlrd
import pyperclip
import time
import os

#mouseClick

def mouseClick(clickTimes,lOrR,img,reTry):
    if reTry == 1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                break
            print("--img not found,0.1s retry")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                print("--Repeat")
                i += 1
            time.sleep(0.1)


#hotkey

#hotkey_get split by ',' ，max 10 values。
def hotkey_get(hk_g_inputValue):
            newinput = hk_g_inputValue.split(',')
            if len(newinput)==1: 
           			 pyautogui.hotkey(hk_g_inputValue)
            elif len(newinput)==2:
           			 pyautogui.hotkey(newinput[0],newinput[1])
            elif len(newinput)==3:
           			 pyautogui.hotkey(newinput[0],newinput[1],newinput[2])
            elif len(newinput)==4:
           			 pyautogui.hotkey(newinput[0],newinput[1],newinput[2],newinput[3])
            elif len(newinput)==4:
               			 pyautogui.hotkey(newinput[0],newinput[1],newinput[2],newinput[3])
            elif len(newinput)==5:
               			 pyautogui.hotkey(newinput[0],newinput[1],newinput[2],newinput[3],newinput[4])
            elif len(newinput)==6:
               			 pyautogui.hotkey(newinput[0],newinput[1],newinput[2],newinput[3],newinput[4],newinput[5])
            elif len(newinput)==7:
               			 pyautogui.hotkey(newinput[0],newinput[1],newinput[2],newinput[3],newinput[4],newinput[5],newinput[6])       
            elif len(newinput)==8:
               			 pyautogui.hotkey(newinput[0],newinput[1],newinput[2],newinput[3],newinput[4],newinput[5],newinput[6],newinput[7])     
            elif len(newinput)==9:
               			 pyautogui.hotkey(newinput[0],newinput[1],newinput[2],newinput[3],newinput[4],newinput[5],newinput[6],newinput[7],newinput[8])
            elif len(newinput)==10:
               			 pyautogui.hotkey(newinput[0],newinput[1],newinput[2],newinput[3],newinput[4],newinput[5],newinput[6],newinput[7],newinput[8],newinput[9])   
                                                                                                                                                         
#hotkey_Group call hotkey_get，Logic check。
def hotkeyGroup(hotkey_reTry,hkg_inputValue):
    if hotkey_reTry == 1:
            hotkey_get(hkg_inputValue)                  
            print("Exc:",hkg_inputValue)
            time.sleep(0.1)
    elif hotkey_reTry == -1:
        while True:
            hotkey_get(hkg_inputValue)
            print("Exc:",hkg_inputValue)
            time.sleep(0.1)
    elif hotkey_reTry > 1:
        i = 1
        while i < hotkey_reTry + 1:
                hotkey_get(hkg_inputValue)
                print("Exc:",hkg_inputValue)
                i += 1
                time.sleep(0.1)



# dataCheck
# cmdType.value  1.0 Left click    2.0 Left double click  3.0 right click  4.0 input  5.0 wait  6.0 scoll  
# 7.0 hot key
# 8.0 paste time
# 9.0 system cmd
# ctype     blank：0
#           String：1
#           number：2
#           time：3
#           bool：4
#           error：5
def dataCheck(sheet1):
    checkCmd = True
    #行数检查
    if sheet1.nrows<2:
        print("--No Data--")
        checkCmd = False
    #每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0 
        and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0 
        and cmdType.value != 7.0 and cmdType.value != 8.0 and cmdType.value != 9.0):
            print('Row ',i+1,",Column 1 Data Not Correct!")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value ==1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('Row ',i+1,",Column 2 Data Not Correct!")
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('Row ',i+1,",Column 2 Data Not Correct!")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('Row ',i+1,",Column 2 Data Not Correct!")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('Row ',i+1,",Column 2 Data Not Correct!")
                checkCmd = False
        # 7.0 热键组合，内容不能为空
        if cmdType.value == 7.0:
            if cmdValue.ctype == 0:
                print('Row ',i+1,",Column 2 Data Not Correct!")
                checkCmd = False
        # 8.0 时间，内容不能为空
        if cmdType.value == 8.0:
            if cmdValue.ctype == 0:
                print('Row ',i+1,",Column 2 Data Not Correct!")
                checkCmd = False
        # 9.0 系统命令集模式，内容不能为空
        if cmdType.value == 9.0:
            if cmdValue.ctype == 0:
                print('Row ',i+1,",Column 2 Data Not Correct!")
                checkCmd = False
        i += 1
    return checkCmd

#任务
def mainWork(img):
    i = 1
    while i < sheet1.nrows:
        #取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            #Get Img Name
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"left",img,reTry)
            print("Left Click:",img)
        #2代表双击左键
        elif cmdType.value == 2.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            #取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2,"left",img,reTry)
            print("Double Click:",img)
        #3代表右键
        elif cmdType.value == 3.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            #取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"right",img,reTry)
            print("Right Click",img) 
        #4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
            print("Input:",inputValue)                                        
        #5代表等待
        elif cmdType.value == 5.0:
            #取图片名称
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("Wait",waitTime,"s")
        #6代表滚轮
        elif cmdType.value == 6.0:
            #取图片名称
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("Scoll ",int(scroll)," line")     
       #7代表_热键组合
        elif cmdType.value == 7.0:
            #取重试次数,并循环。
            hotkey_reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                hotkey_reTry = sheet1.row(i)[2].value
            inputValue = sheet1.row(i)[1].value
            hotkeyGroup(hotkey_reTry,inputValue)
            time.sleep(0.5)
       #8代表_粘贴当前时间
        elif cmdType.value == 8.0:      
            #设置本机当前时间。
            localtime = time.strftime("%Y-%m-%d %H：%M：%S", time.localtime()) 
            pyperclip.copy(localtime)
            pyautogui.hotkey('ctrl','v')
            print("Paste Time :",localtime)
            time.sleep(0.5)
       #9代表_系统命令集模式
        elif cmdType.value == 9.0:      
            wincmd = sheet1.row(i)[1].value
            os.system(wincmd)
            print("Run Cmd:",wincmd)
            time.sleep(0.5) 
        i += 1

if __name__ == '__main__':
    file = 'steps.xls'
    #打开文件
    wb = xlrd.open_workbook(filename=file)
    #通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print('AutoClick_v20211221')
    #数据检查
    checkCmd = dataCheck(sheet1)
    if checkCmd:
        key=input('Times: 1.One-Time 2.Repeat \n')
        if key=='1':
            #循环拿出每一行指令
            mainWork(sheet1)
        elif key=='2':
            while True:
                mainWork(sheet1)
                time.sleep(0.1)
                print("---LoopEnd wait 0.1s ----")    
    else:
        print('Input Error or Exit!')
