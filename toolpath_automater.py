import os
import subprocess
import pyautogui
import time
import re
import logging
from pathlib import Path
from dxfFileOrganizer import dxfExtractor
from vcrvFileMaker import vcrvMaker


pyautogui.FAILSAFE = True
logging.basicConfig(filename='TP_automater_logging.txt', level=logging.DEBUG,
                    format=' %(asctime)s - %(levelname)s - %(message)s')
logging.disable()


def fileOpener(filename, completePath):
    # regex = re.compile('(birch)|(formica)|(fenix)')
    rx = '(^6mm.*\.dxf$)|(^8mm.*\.dxf$)|(.*birch.*\.dxf$)|(.*oak.*\.dxf$)|(.*walnut.*\.dxf$)|(.*formica.*\.dxf$)|(.*fenix.*\.dxf$)'

    global x_dim
    global y_dim
    global material_thickness
    global material_type

    subprocess.Popen(['start', completePath], shell=True)
    time.sleep(1)

    res = re.match(rx, filename)
    if res:
        # 6mm or 8mm match
        if res.group(1) or res.group(2):
            x_dim = '2440'
            y_dim = '1220'
            material_thickness = '6'
            material_type = 'Generic'
        # birch match
        if res.group(3):
            x_dim = '2440'
            y_dim = '1220'
            material_thickness = '18'
            material_type = 'Birch'
        # oak/walnut match
        if res.group(4) or res.group(5):
            x_dim = '2440'
            y_dim = '1220'
            material_thickness = '18.6'
            material_type = 'Oak/Walnut'
        # formica match
        if res.group(6):
            x_dim = pyautogui.prompt(
                'Please type in length of sheet, carefully (i.e. 2440 for normal or 3050 for jumbo): ')
            y_dim = pyautogui.prompt(
                'Please type in width of sheet, carefully (i.e. 1220 for normal or 1300 for jumbo): ')
            material_thickness = '19.6'
            material_type = 'Formica'
        # fenix match
        if res.group(7):
            x_dim = pyautogui.prompt(
                'Please type in length of sheet, carefully (i.e. 2440 for normal or 3050 for jumbo): ')
            y_dim = pyautogui.prompt(
                'Please type in width of sheet, carefully (i.e. 1220 for normal or 1300 for jumbo): ')
            material_thickness = '20'
            material_type = 'Fenix'


def jobSetup():
    time.sleep(1)
    vcrvApp = pyautogui.getActiveWindow()
    vcrvApp.maximize()
    time.sleep(1)
    pyautogui.doubleClick(368, 595)
    pyautogui.write(x_dim)
    pyautogui.press('\t')
    pyautogui.write(y_dim)
    pyautogui.press('\t')
    pyautogui.write(material_thickness)
    pyautogui.press('\t', presses=4, interval=0.5)
    pyautogui.press('space')  # unclick "use offset"
    # pyautogui.press('\t', presses=6, interval=1)
    # pyautogui.press('space')  # jobSetup ok
    # pyautogui.click('ok.png')
    time.sleep(1)
    pyautogui.click(132, 1848)  # jobSetup ok


def joiningSimple():
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('j')
    pyautogui.press('enter')
    time.sleep(1)
    # pyautogui.click('close.png')
    pyautogui.click(468, 890)  # join_close


def joining():
    time.sleep(2)
    # pyautogui.click('layerTab.png')
    pyautogui.click(480, 1978)  # layerTab
    time.sleep(1)
    # pyautogui.rightClick('10mm_cutout.png')
    pyautogui.rightClick(271, 370)  # 10mm_cutout
    pyautogui.press('down', presses=10, interval=0.3)
    pyautogui.press('enter')
    pyautogui.press('j')
    pyautogui.press('\t')
    pyautogui.press('space')  # join
    time.sleep(1)
    # pyautogui.click('close.png')
    pyautogui.click(468, 890)  # join_close
    pyautogui.move(250, 0, duration=0.1)  # move mouse right

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('j')
    pyautogui.press('\t')
    pyautogui.press('space')  # join
    time.sleep(1)
    # pyautogui.click('close.png')
    pyautogui.click(468, 890)  # join_close


def simpleNest():
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    # pyautogui.click('nestTool.png')
    pyautogui.click(441, 1543)  # nestTool
    pyautogui.press('\t', presses=14, interval=0.3)
    pyautogui.press('space')  # preview
    time.sleep(1)
    # pyautogui.click('ok.png')
    pyautogui.click(140, 1678)  # nestTool ok


def toolSet():
    time.sleep(1)
    pyautogui.click(3820, 196)  # toolpathTab
    time.sleep(1)
    pyautogui.click(3756, 151)  # tabLock
    time.sleep(1)
    pyautogui.click(3263, 312)  # toolSet
    pyautogui.press('\t', presses=10, interval=0.3)
    pyautogui.write('45')
    pyautogui.press('\t', presses=4, interval=0.3)
    pyautogui.write('45')
    pyautogui.press('\t')
    pyautogui.press('space')  # toolSet ok


def loadTP():
    pyautogui.alert(
        'Please look at main handle type or for Urtil features. NOTE IT DOWN, you will be asked for it shortly.')
    time.sleep(1)
    pyautogui.click(3219, 858)  # loadTP
    pyautogui.press('\t', presses=5, interval=0.3)
    pyautogui.press('space')  # addressBar
    pyautogui.write('C:\\Users\\sho25\\Desktop\\Toolpath Templates')
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('\t', presses=4, interval=1)
    material_type_select()


def material_type_select():
    # material_type = pyautogui.confirm(text='What is the sheet material type?: ', buttons=[
    #                                   'Birch', 'Oak/Walnut', 'Formica', 'Fenix'])
    if material_type == 'Birch':
        pyautogui.press('1')  # Birch folder
        pyautogui.press('enter')
    elif material_type == 'Oak/Walnut':
        pyautogui.press('2')  # Oak/Walnut folder
        pyautogui.press('enter')
    elif material_type == 'Formica':
        pyautogui.press('3')  # Formica folder
        pyautogui.press('enter')
    elif material_type == 'Fenix':
        pyautogui.press('4')  # Fenix folder
        pyautogui.press('enter')
    elif material_type == 'Generic':
        pyautogui.press('5')  # Generic folder
        pyautogui.press('enter')
        pyautogui.press('1')  # Urtil-Door-And-Back TP
        pyautogui.press('enter')  # create missing layer
        time.sleep(1)
        pyautogui.press('enter')  # load selected TP


def handle_prompt():
    handle_type = pyautogui.confirm(
        text='What is the feature/handle type?: ', buttons=['None/Grab/Circle', 'D-Pull', 'Edge-Pull', 'SRG/SRC', 'Urtil', 'Omlopp', 'CANCEL'])
    if handle_type == 'None/Grab/Circle':
        time.sleep(1)
        pyautogui.press('1')  # None/Grab/Circle TP
        pyautogui.press('enter')  # create missing layer
        time.sleep(1)
        pyautogui.press('enter')  # load selected TP
    elif handle_type == 'D-Pull':
        time.sleep(1)
        pyautogui.press('2')  # D-Pull TP
        pyautogui.press('enter')  # create missing layer
        time.sleep(1)
        pyautogui.press('enter')  # load selected TP
    elif handle_type == 'Edge-Pull':
        time.sleep(1)
        pyautogui.press('3')  # Edge-Pull TP
        pyautogui.press('enter')  # create missing layer
        time.sleep(1)
        pyautogui.press('enter')  # load selected TP
    elif handle_type == 'SRG/SRC':
        time.sleep(1)
        pyautogui.press('4')  # SRG/SRC TP
        pyautogui.press('enter')  # create missing layer
        time.sleep(1)
        pyautogui.press('enter')  # load selected TP
    elif handle_type == 'Urtil':
        time.sleep(1)
        pyautogui.press('5')  # Urtil TP
        pyautogui.press('enter')  # create missing layer
        time.sleep(1)
        pyautogui.press('enter')  # load selected TP
    elif handle_type == 'Omlopp':
        time.sleep(1)
        pyautogui.press('6')  # Omlopp TP
        pyautogui.press('enter')  # create missing layer
        time.sleep(1)
        pyautogui.press('enter')  # load selected TP
    elif handle_type == 'CANCEL':
        time.sleep(1)
        pyautogui.press('esc')


def save_vcrv():
    time.sleep(2)
    pyautogui.press('esc')
    currentWindow = pyautogui.getActiveWindow()
    pyautogui.hotkey('ctrl', 's')
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.confirm('Toolpath is set and saved. Happy to move to next file?')
    currentWindow = pyautogui.getActiveWindow()
    currentWindow.close()


logging.debug('Start of Program')
pyautogui.alert(
    'Make sure required downloaded CAD zip file is located on desktop')
dxfExtractor()
baseDir = 'C:\\Users\\sho25\\Desktop\\extractedFile'
baseDirList = os.listdir(baseDir)
logging.debug('baseDirList before for loop = (%s%%)' % (baseDirList))
for file_count in baseDirList:
    logging.debug('file_count in first for loop = (%s%%)' % (file_count))
    targetDir = os.path.join(baseDir, file_count)
    logging.debug('targetDir in first for loop = (%s%%)' % (targetDir))
    for client in targetDir:
        clientName = os.listdir(targetDir)
        logging.debug('clientName inside for-for-if loop = (%s%%)' %
                      (clientName))
        clientPath = os.path.join(targetDir, clientName[0])
        logging.debug('clientPath inside for-for-if loop = (%s%%)' %
                      (clientPath))
        clientFileList = os.listdir(clientPath)
        logging.debug('clientFileList inside for-for-if loop = (%s%%)' %
                      (clientFileList))
        for filename in clientFileList:
            logging.debug(
                'filename inside for-for-if-for loop = (%s%%)' % (filename))
            completePath = os.path.join(clientPath, filename)
            logging.debug(
                'completePath inside for-for-if-for loop = (%s%%)' % (completePath))
            ext = os.path.splitext(filename)[-1].lower()
            if ext == '.dxf':
                # also for some reason doesnt work when theres already a .crv file - WHYYYY?
                # NEEDS TO BE SHOWN THROUGH TKINTER - show clientName and file type so that prompt sheet size makes sense.
                pyautogui.alert(text=clientName[0],
                                button='Confirm client\'s name.')
                pyautogui.alert(text=filename, button='Enter into this file.')
                fileOpener(filename, completePath)
                jobSetup()
                if material_thickness == '6':
                    logging.debug('entering for-if-for-if loop == generic TP')
                    joiningSimple()
                    toolSet()
                    loadTP()
                    save_vcrv()

                if material_thickness == '18' or material_thickness == '18.6':
                    logging.debug('entering for-if-for-if loop == wood TP')
                    joining()
                    toolSet()
                    loadTP()
                    handle_prompt()
                    while True:
                        repeat_loadTP = pyautogui.confirm(text='Add more toolpath template for additional handle type?', buttons=[
                                                          'Add another handle type', 'DONE'])
                        if repeat_loadTP == 'DONE':
                            break
                        else:
                            loadTP()
                            handle_prompt()
                    save_vcrv()

                elif material_thickness == '19.6' or material_thickness == '20':
                    logging.debug('entering for-if-for-if loop == laminate TP')
                    joining()
                    simpleNest()
                    toolSet()
                    loadTP()
                    handle_prompt()
                    while True:
                        repeat_loadTP = pyautogui.confirm(text='Add more toolpath template for additional handle type?', buttons=[
                                                          'Add another handle type', 'DONE'])
                        if repeat_loadTP == 'DONE':
                            break
                        else:
                            loadTP()
                            handle_prompt()
                    save_vcrv()
            else:
                logging.debug(
                    'filename is not .dxf - entered else clause. filename = (%s%%)' % (filename))
                break

pyautogui.alert(
    'All Toolpath Automation completed. Completed VCarve Files will be collected onto the Desktop.')
vcrvMaker()
