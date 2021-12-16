import pyautogui as pg
import time
import pyperclip as pp
import pandas as pd

def mouseclick(clicktimes, lOrR, img):
    while True:
        location= pg.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pg.click(location.x,location.y,clicks=clicktimes,interval=0.2,duration=0.2,button=lOrR)
            # flag = True;
            break;
        print("cannot find image. retry esle 0.5 seconds")
        time.sleep(0.5)

def run_macro(sheet):
    i = 0
    while i<sheet.shape[0]:
        value = sheet.loc[i][0]
        if value == 1:
            # get image name
            img = sheet.loc[i][1]
            mouseclick(1, "left", img)
            print("one left click ", img)
        elif value == 2:
            img = sheet.loc[i][1]
            mouseclick(2, "left", img)
            print("two left click ", img)
        elif value == 3:
            img = sheet.loc[i][1]
            mouseclick(1, "right", img)
            print("one right click ", img)
        elif value == 4:
            inputcontent = sheet.loc[i][1]
            pp.copy(inputcontent)
            pg.hotkey('command', 'v')
            time.sleep(0.5)
            print("input: ", inputcontent)
        elif value == 5:
            waittime = sheet.loc[i][1]
            time.sleep(waittime)
            print("wait ", waittime, "seconds")
        elif value == 6:
            scroll = sheet.loc[i][1]
            pg.scroll(int(scroll))
            print("scroll ", int(scroll), "pixel")
        elif value == 0:
            hotkeys = df.loc[i][1]
            hotkeys = hotkeys.split("+")
            length = len(hotkeys)
            if(length==1):
                pg.press(hotkeys[0])
            elif(length==2):
                pg.hotkey(hotkeys[0], hotkeys[1], interval=0.25)
            elif(length==3):
                pg.hotkey(hotkeys[0], hotkeys[1], hotkeys[2])
        i+=1




# data check
# value 1 for left one click;  2 for left two click; 3 for rigth one click; 4 for input; 5 for wait
#               6 for scroll; 0 for hotkeys

def dataCheck(sheet):
    checkCmd = True
    if sheet.shape[0]<2:
        print("There is no data in your cmd.xlsx.")
        checkCmd = False
    i = 0;
    while i < sheet.shape[0]-1:
        # check the first col
        cmdvalue = sheet.loc[i][0]
        if (cmdvalue != 1 and cmdvalue != 2 and cmdvalue != 3 
        and cmdvalue != 4 and cmdvalue != 5 and cmdvalue != 6 and cmdvalue != 0):
            print('Data in line',i+1,"is wrong!")
            checkCmd = False
        # check the second col
        convalue = sheet.loc[i][1]
        # when the comtype is 1, 2, 3, the content type should be string
        if cmdvalue ==1 or cmdvalue == 2 or cmdvalue == 3:
            if (str(convalue)).isdigit() == True:
                print("The data in line",i+1, "col 2 is wrong!string")
                checkCmd = False
        # when the comptype is 5 wait. the content type should be number
        if cmdvalue == 5:
            if (str(convalue)).isdigit() == False:
                print("The data in line",i+1, "col 2 is wrong!wait")
                checkCmd = False
        # when the comtype is 6 scroll. the content type should be number.
        if cmdvalue == 6:
            if (str(convalue)).isdigit() == False:
                print("The data in line",i+1, "col 2 is wrong!number")
                checkCmd = False
        i += 1
    return checkCmd

if __name__ == '__main__':
    choice = input("Please input a number, 1: Open bilibili, 2: connect VPN: ")
    df = pd.read_excel('cmd.xlsx', sheet_name='Sheet1'+choice)
    checkCmd = dataCheck(df)
    if checkCmd:
        run_macro(df);

