#! python3
# Setup Chrome and Excel before run this program
# Excel Edit Cells would be 2nd line position on screen (Firstline Freeze)
import tkinter as tk
import time,pyautogui,pyperclip

#global e1

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
              time.sleep(0.5)
              pyautogui.click()
              time.sleep(0.5)
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
              
'''
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
'''
'''
              pyautogui.moveTo(xExcelDown)
              time.sleep(0.25)
              pyautogui.click()
              time.sleep(0.5)
'''
def inputBoxTk():
              global e1
              def continueRun(event):        #Send out event
                            var=e1.get()
                            pyperclip.copy(var)
                            print(var)
                            e1.delete(0,20)
                            window.destroy()
                            
              window=tk.Tk()
              window.title('Please input...')
              window.geometry('150x100+18+457')
              window.wm_attributes('-topmost',1)

              e1 =tk.Entry(window,show=None)
              e1.insert(10,'404')
              e1.pack()
 

              b1=tk.Button(window,text='Quit',command=window.quit)
              b1.pack()

              window.bind('<Return>', continueRun)
              #while e1.get()==str(404):
              #             time.sleep(3)
              #Button(root, text='Quit', command=root.quit).grid(row=3, column=0, sticky=W, pady=6)

             # time.sleep(3)

              #Button(master, text='Show', command=show_entry_fields).grid(row=3, column=1, sticky=W, pady=6)

              window.mainloop( )
              
def FillinExcelxCV1():
              # Close current edit area
              xCloseN=(399,200)
              pyautogui.moveTo(xCloseN)
              time.sleep(0.25)
              pyautogui.click()
              time.sleep(0.1)
              
              # Back to Excel
              xExcel=(337,746)
              xExcelDown=(1356,670)
              xCellEditC=(267,238)
              pyautogui.moveTo(xExcel)
              time.sleep(0.25)
              pyautogui.click()
              time.sleep(0.5)
              
              # Move and Paste in C
              pyautogui.moveTo(xCellEditC)
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
print('''Set up Chrome in index page, Status to be all done
              1. Chrome in Index Management
              2. Status contains PUBLISH, ARCHIVED, DISABLED''')

# Let user define row numbers to take action
print('Please input rows to take action: ')
n=input()
# Notice user to open excel window manually
print('Please open Excel in 5s...')
time.sleep(5)
for i in range(int(n)):
              checkKBFeatureV1()
              #import tkinter as TK
              inputBoxTk()
              # while e1.get()==str(404):
              #              time.sleep(3) # TODO: Stop in this step, inputbox not appear
              print('Sleep end +'+str(i))
                            
              FillinExcelxCV1()

print('Action has taken for '+n+' rows.' )

#MessageBox to notify user finish running
from tkinter import messagebox
#import tkMessageBox
messagebox.showinfo('Python-checkKBFeatureV2', 'Action has taken for '+n+' rows.')
