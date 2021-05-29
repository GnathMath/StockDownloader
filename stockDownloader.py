import pyautogui as pag
import time, os, datetime, shutil
import keyboard
import random
import win32api, win32con
img='NSE Website images\\'
tab=img+'newTab.PNG'
archives=img+'archivesTab.PNG'
dateBox=img+'dateBox.PNG'
leftButton='left.png'
bhavcopy=img+'Bhavcopy.PNG'
download=img+'download.png'
top=img+'topImage.png'
#pag.displayMousePosition()
target=datetime.date(1999,1,1)
def lClick(x,y):
    pag.moveTo(x,y,0.5)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(1) #This pauses the script for 0.05 seconds
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)
def rClick(x,y):
    pag.moveTo(x,y,0.5)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,0,0)
    time.sleep(0.5) #This pauses the script for 0.05 seconds
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,0,0)
def finder(img,side):
    temp=4000
    attempt = 0
    while temp>2999:
        if pag.locateOnScreen(img, grayscale=True) != None:  #region=(150,175,259,194),
            print("I found "+img+"\n and clicking now")
            pic=pag.locateCenterOnScreen(img, grayscale=True)
            cent=pic
            temp=1000
        elif attempt>20:
            pag.press('home')
            time.sleep(1)
            attempt=0
        else:
            print("I could not find "+img)
            attempt += 1
            temp=3000
            cent=[0,0]
            pag.scroll(-200)
            print("Scrolled and checking")
        time.sleep(1)
    if side == 'l':
        lClick(cent[0],cent[1])
    elif side == 'r':
        rClick(cent[0],cent[1])
    else:
        pag.moveTo(cent[0],cent[1],1)

def d(dd):
    if dd<10:
        return '0'+str(dd)
    else:
        return str(dd)

def m(mm):
    months = {1 : 'JAN',
              2 : 'FEB',
              3 : 'MAR',
              4 : 'APR',
              5 : 'MAY',
              6 : 'JUN',
              7 : 'JUL',
              8 : 'AUG',
              9 : 'SEP',
              10: 'OCT',
              11: 'NOV',
              12: 'DEC'}
    return months[mm]

def previousDay(date):
    day31 = [1,3,5,7,8,10,12]
    day30 = [4,6,9,11]
    day = date.day
    month = date.month
    year = date.year
    day = day - 1
    if day == 0:
        month = month - 1
        if month == 0:
            year = year - 1
            month = 12
        if month in day31:
            day = 31
        elif month in day30:
            day = 30
        elif year%4 == 0:
            day = 29
        else:
            day = 28
    return datetime.date(year,month,day)

def previousWeekDay(date):
    if 1<date.isoweekday()<7:
        return previousDay(date)
    elif date.isoweekday() == 7:
        return previousDay(previousDay(date))
    else:
        return previousDay(previousDay(previousDay(date)))
time.sleep(5)
##finder(tab,'l')
##pag.write('https://www.nseindia.com/all-reports', interval=0.25)  # prints out "Hello world!" with a quarter second delay after each character
##pag.press('enter')
##time.sleep(5)
##pag.moveRel(0,100,1)
start = datetime.date.today()
current = previousWeekDay(start)
while target!=current:
    dd=current.day
    mm=current.month
    yy=current.year
    
    if os.path.exists('C:\\Users\\Sujathaselvi\\Downloads\\cm'+d(dd)+m(mm)+str(yy)+'bhav.csv.zip'):
        current = previousWeekDay(current)
        print('File on '+m(mm)+' '+d(dd)+', '+str(yy)+' has already been downloaded')
    else:
        print('Now, I will try to download the file on '+m(mm)+' '+d(dd)+', '+str(yy))
        finder(archives,'l')
        time.sleep(1)
        finder(dateBox,'l')
        time.sleep(1)
        reg=pag.locateCenterOnScreen(img+'week.png',grayscale=True, confidence=0.9)
        if mm == start.month and yy == start.year:
            finder(img+str(dd)+'.png','s')
            pag.click()
            time.sleep(1)
        elif yy == start.year:
            lClick(reg[0],reg[1]-31)
            time.sleep(1)
            finder(img+m(mm)+'.png','l')
            time.sleep(1)
            pag.moveRel(200,0,1)
            finder(img+str(dd)+'.png','s')
            pag.click()
            time.sleep(1)
        else:
            lClick(reg[0],reg[1]-31)
            time.sleep(1)
            clickTimes = start.year - yy
            for i in range(clickTimes):
                finder(img+leftButton,'l')
                time.sleep(1)
            finder(img+m(mm)+'.png','l')
            time.sleep(1)
            pag.moveRel(200,0,1)
            finder(img+str(dd)+'.png','s')
            pag.click()
            time.sleep(1)
        finder(bhavcopy,'s')
##        time.sleep(1)
        reg2 = pag.locateOnScreen(bhavcopy,grayscale=True, confidence=0.9)
        position = pag.locateCenterOnScreen(download,region=reg2,grayscale=True, confidence=0.9)
        lClick(position[0],position[1])
        print('Clicking Download Button')
        waitTime = 0
        while os.path.exists('C:\\Users\\Sujathaselvi\\Downloads\\cm'+d(dd)+m(mm)+str(yy)+'bhav.csv.zip') != True and waitTime<9.9:
            time.sleep(0.5)
            waitTime += 0.6
        if waitTime < 9.9:
            print('cm'+d(dd)+m(mm)+str(yy)+'bhav.csv.zip is downloaded')
        else:
            print('Download is unsuccessful')
        pag.press('F5')
        time.sleep(1)
        print('Refreshing the page')
        pag.moveRel(-150,0)
        print('Scrollig to the top')
        while pag.locateOnScreen(top, grayscale=True, confidence=0.9)== None:
            pag.press('home')
            time.sleep(1)
            pag.moveRel(150,0)
##        shutil.move('C:\\Users\\Sujathaselvi\\Downloads\\cm'+d(dd)+m(mm)+str(yy)+'bhav.csv.zip','BhavCopy Zip files')
        current = previousWeekDay(current)
