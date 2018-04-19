#! python3
# Setup Chrome and Excel before run this program
# Excel Edit Cells would be 2nd line position on screen (Firstline Freeze)
import time,pyautogui
def checkKBFeatureV1():
              #import time,pyautogui
              # Set these to the correct coordinates for your computer
              xKnowledgeID=(384,240)
              xSearch=(730,340)
              xEdit=(314,432)
              xFeature=(945,582)
              xCloseN=(399,200)
              xChrome=(86,749)

              xExcel=(337,746)
              xExcelDown=(1356,670)
              xExcelUp=(1354,201)
              xCellEditA=(55,239)
              xCellEditB=(200,242)
              xCellEditC=(267,238)

              # Operation in Excel, copy A
              pyautogui.moveTo(xCellEditA)
              time.sleep(0.25)
              pyautogui.click()
              time.sleep(0.25)
              pyautogui.hotkey('ctrl','c')

              # Open Chrome to Search page
              # Setup Chrome before run this program
              pyautogui.click(xChrome)
              time.sleep(0.75)
              pyautogui.moveTo(xKnowledgeID)
              pyautogui.click()
              pyautogui.hotkey('ctrl','a')
              pyautogui.hotkey('ctrl','v')

              pyautogui.moveTo(xSearch)
              pyautogui.click()
              time.sleep(2.5)

              pyautogui.moveTo(xEdit)
              time.sleep(0.75)
              pyautogui.click()
              time.sleep(2.5)

              for i in range(6):
                            pyautogui.press('pagedown')
                            time.sleep(0.1)

              pyautogui.moveTo(xFeature)
              time.sleep(0.25)
              pyautogui.click()
              pyautogui.hotkey('ctrl','a')
              pyautogui.hotkey('ctrl','c')

              for i in range(6):
                            pyautogui.press('pageup')
                            time.sleep(0.1)

              pyautogui.click(xCloseN)
              time.sleep(0.25)

              # Back to Excel
              pyautogui.moveTo(xExcel)
              time.sleep(0.25)
              pyautogui.click()
              time.sleep(0.5)

              # Move and Paste in B
              pyautogui.moveTo(xCellEditB)
              time.sleep(0.25)
              pyautogui.click()
              time.sleep(0.1)

              pyautogui.hotkey('ctrl','v')
              time.sleep(0.1)
              pyautogui.hotkey('ctrl','s') #Save Excel
              time.sleep(1.5)

              pyautogui.moveTo(xExcelDown)
              time.sleep(0.25)
              pyautogui.click()
              time.sleep(0.5)
# Main program begins
print('Set up Chrome in index page, Status to be all ')

# Let user define row numbers to take action
print('Please input rows to take action: ')
n=input()
# Notice user to open excel window manually
print('Please open Excel in 5s...')
time.sleep(5)
for i in range(int(n)):
              checkKBFeatureV1()
print('Action has taken for '+n+' rows.' )

#MessageBox to notify user finish running
from tkinter import messagebox
messagebox.showinfo('Python-checkKBFeatureV2', 'Action has taken for '+n+' rows.')
