import pyautogui
import time
import datetime
import ezgmail
import ezsheets

ss = ezsheets.Spreadsheet('1G3AfREoWoOFsnwuQZ2VhTEOWXYS8NJAU4liSLgCJo-U')
sheet=ss[0]
print(sheet['B'+str(2)])

month=str(datetime.datetime.now().month)
day=str(datetime.datetime.now().day)
if len(day)==1:
    day='0'+day

def tab():
    pyautogui.write('\t')
def enter():
    pyautogui.write('\n')
def openForm():
    pyautogui.moveTo(1350, 220, duration=0.25)
    pyautogui.click(button='right')
    pyautogui.moveTo(1250,650,duration=0.5)
    time.sleep(0.1)
    pyautogui.click()
    time.sleep(2)
    pyautogui.write('https://docs.google.com/forms/d/e/1FAIpQLSd7CfxIWRAzK8GFRWjRi50wHFwHOmjUxco7Z3JpY89DmDqRMw/viewform?gxids=7628')
    enter()

def fillForm():
    time.sleep(2)
    tab()
    tab()
    pyautogui.write(sheet['B'+str(row)])
    tab()
    pyautogui.write(sheet['C'+str(row)])
    tab()
    pyautogui.write(month+day+'2021')
    tab()
    enter()
    for i in range(int(sheet['D'+str(row)])):
        pyautogui.write(['down'])
    enter()
    time.sleep(0.5)
    for i in range(15):
        tab()
    pyautogui.write([' '])
    tab()
    enter()
    time.sleep(2)
    tab()
    tab()
    pyautogui.write([' '])
    tab()
    for i in range(3):
        pyautogui.write(['down'])
        tab()
        tab()
    tab()
    ezgmail.send(sheet['E'+str(row)], sheet['B'+str(row)], 'Your COVID form has been filled, have a nice day :)')
    print('Text send to '+sheet['B'+str(row)])
    enter()
    time.sleep(1)
    tab()
    enter()

print('Starting row?')
start=int(input())
print('Ending row?')
end=int((input()))
print(pyautogui.size())
openForm()


#ending num is the first blank row in sheets
for row in range(start,end):
    time.sleep(1)
    fillForm()

