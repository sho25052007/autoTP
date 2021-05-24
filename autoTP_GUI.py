from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox, Menu, ttk, Checkbutton

# import os
# import subprocess
# import time
# import re
# from pathlib import Path

import pyautogui
from dxfFileOrganizer import dxfExtractor
from vcrvFileMaker import vcrvMaker
from toolpath_automater import completeFileMove, fullAutoTP, jobSetup, simpleNest, singleFolderAccess, joining, toolSetup, loadTP, save_vcrv
import toolpath_automater


class autoTP(Tk):
    def __init__(self):
        super(autoTP, self).__init__()
        self.title("Toolpath Automater")
        self.geometry('3000x1600')

        style = ttk.Style()
        style.theme_create('sea', settings={
            ".": {
                "configure": {
                    "background": '#79ADDA'  # All except tabs
                }
            },
            # "TLabel": {"configure": {"background": '#3D91DA'}},
            # "TFrame": {"configure": {"background": '#3D91DA'}},
            "TNotebook": {
                "configure": {
                    "background": '#CCCCCC',  # Your margin color
                    # margins: left, top, right, separator
                    "tabmargins": [2, 10, 0, 0],
                }
            },
            "TNotebook.Tab": {
                "configure": {
                    "background": '#CCCCCC',  # tab color when not selected
                    # [space between text and horizontal tab-button border, space between text and vertical tab_button border]
                    "padding": [20, 20],
                    "font": ('Helvetica', '10')
                },
                "map": {
                    # Tab color when selected
                    "background": [("selected", '#3D91DA')],
                    "expand": [("selected", [1, 8, 1, 0])]  # text margins
                }
            },
            "TButton": {
                "configure": {
                    "background": '#CCCCCC',
                    "padding": [20, 20],
                    "font": ('Helvetica', '10', 'underline'),
                    "relief": RAISED
                }
            }
        })

        style.theme_use('sea')
        style.configure("Tab", focuscolor=style.map(
            "TNotebook.Tab")["background"])  # makes selection dotted line to be same colour as selected tab background = #3D91DA
        style.configure('TButton', focuscolor=style.map(
            "TNotebook.Tab")["background"])

        self.createMenu()

        tabControl = ttk.Notebook(self)
        self.tab1 = ttk.Frame(tabControl)
        self.tab1.columnconfigure(0, weight=1)
        # self.tab1.rowconfigure(0, weight=1)
        tabControl.add(self.tab1, text="Welcome - AutoTP Guide")

        self.tab2 = ttk.Frame(tabControl)
        self.tab2.columnconfigure(0, weight=1)
        tabControl.add(self.tab2, text="Step 1: Job Setup")

        self.tab3 = ttk.Frame(tabControl)
        self.tab3.columnconfigure(0, weight=1)
        tabControl.add(self.tab3, text="Step 2: Joining Lines")

        self.tab4 = ttk.Frame(tabControl)
        self.tab4.columnconfigure(0, weight=1)
        tabControl.add(
            self.tab4, text="Step 3: Nest (laminates only)")

        self.tab5 = ttk.Frame(tabControl)
        self.tab5.columnconfigure(0, weight=1)
        tabControl.add(self.tab5, text="Step 4: Tool Setup")

        self.tab6 = ttk.Frame(tabControl)
        self.tab6.columnconfigure(0, weight=1)
        tabControl.add(self.tab6, text="Step 5: Load toolpath")

        self.tab7 = ttk.Frame(tabControl)
        self.tab7.columnconfigure(0, weight=1)
        tabControl.add(self.tab7, text="Step 6: Save toolpath")

        self.tab8 = ttk.Frame(tabControl)
        self.tab8.columnconfigure(0, weight=1)
        tabControl.add(self.tab8, text="Complete AutoTP")

        tabControl.pack(expand=1, fill='both')
        self.tab_control = tabControl

        self.create_tab1()
        self.create_tab2()
        self.create_tab3()
        self.create_tab4()
        self.create_tab5()
        self.create_tab6()
        self.create_tab7()
        self.create_tab8()

    def createMenu(self):
        menubar = Menu(self)
        self.config(menu=menubar)

        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label='File', menu=file_menu)
        file_menu.add_command(label='Run Step-by-Step',
                              command=self.select_tab2)
        file_menu.add_command(label='Run Complete AutoTP',
                              command=self.select_tab_autoTP)
        file_menu.add_command(label='Settings')
        file_menu.add_command(label='Help')

        organize_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label='Organizer', menu=organize_menu)
        organize_menu.add_command(label='Extract ZIP', command=self.extractZIP)
        organize_menu.add_command(
            label='Compile VCarve Files', command=self.compileVCRV)

    def extractZIP(self):
        self.confirm_dxfExtractor = messagebox.askokcancel(
            'Confirm Desktop Zipfile Extraction', 'Please make sure downloaded ZIP file of CAD from MegaAdmin is located on Desktop')
        if self.confirm_dxfExtractor:
            dxfExtractor()
        else:
            pass

    def compileVCRV(self):
        self.confirm_compileVCRV = messagebox.askokcancel(
            'Confirm Desktop VCarve Compiling', "This step only works when you've completed the AutoTP. Please confirm to compile all saved VCarve Files onto the Desktop")
        if self.confirm_compileVCRV:
            vcrvMaker()
        else:
            pass

    def select_tab2(self):
        self.tab_control.select(self.tab2)

    def select_tab3(self):
        self.tab_control.select(self.tab3)

    def select_tab4(self):
        self.tab_control.select(self.tab4)

    def select_tab5(self):
        self.tab_control.select(self.tab5)

    def select_tab6(self):
        self.tab_control.select(self.tab6)

    def select_tab7(self):
        self.tab_control.select(self.tab7)

    def select_tab_autoTP(self):
        self.tab_control.select(self.tab8)

    def create_tab1(self):
        welcome_message = "Hello, welcome to AutoTP.\nWith this program you can automate making toolpaths on VCarve from Plykea Ltd. MegaAdmin."
        self.welcome = Label(
            self.tab1, text=welcome_message, font=('Helvetica', 14), justify=CENTER, anchor=N)
        self.welcome.grid(column=0, row=0, padx=10, pady=200)

    def create_tab2(self):
        step_1_intro = "This is a step-by-step method to guide you through AutoTP.\nIt is meant as a learning tool to get acquainted to this program.\nThis first step will extract a single file and open the .dxf in VCarve set up the job.\nFor the complete AutoTP go to final tab."
        self.intro = Label(self.tab2, text=step_1_intro,
                           font=('Helvetica', 14), justify=CENTER)
        self.intro.grid(column=0, row=1, padx=50, pady=20)

        extract_single_file_button = Button(
            self.tab2, text="Open and Extract Single File for Job Setup", command=self.extractSingleFile)
        extract_single_file_button.grid(column=0, row=3, padx=50, pady=10)

    def extractSingleFile(self):
        baseDir = 'C:\\Users\\sho25\\Desktop\\extractedFile'
        global autoTPapp
        autoTPapp = pyautogui.getActiveWindow()
        confirm_desktop_file = messagebox.askokcancel(
            'Download to Desktop', 'Please download a single client CAD zip from MA to the Desktop for extraction.\n\nREMEMBER: Allow the program to run without touching the mouse or keyboard until popup appears!')
        if confirm_desktop_file:
            # global clientFileEmpty
            # clientFileEmpty = False
            dxfExtractor()
            singleFolderAccess(baseDir)
            jobSetup()
        global vcrvApp
        vcrvApp = pyautogui.getActiveWindow()
        autoTPapp.maximize()
        confirm_jobSetup_complete = messagebox.askokcancel(
            'Job Setup Complete', 'Job Setup is done. Next is step 2: Joining Lines')
        if confirm_jobSetup_complete:
            self.select_tab3()

    def create_tab3(self):
        step_2_intro = "Now that the job is setup, we can go ahead and join all the vector lines."
        self.intro = Label(self.tab3, text=step_2_intro,
                           font=('Helvetica', 14), justify=CENTER)
        self.intro.grid(column=0, row=1, padx=50, pady=20)

        join_lines_button = Button(
            self.tab3, text="Join Vector Lines", command=self.joinLines)
        join_lines_button.grid(column=0, row=3, padx=50, pady=10)

    def joinLines(self):
        confirm_join_lines = messagebox.askokcancel(
            'Joining Lines', 'Confirm to join vector lines.\n\nREMEMBER: Allow the program to run without touching the mouse or keyboard until popup appears!')
        if confirm_join_lines:
            autoTPapp.minimize()
            vcrvApp.maximize()
            joining()
        vcrvApp.minimize()
        autoTPapp.maximize()
        confirm_joinLines_complete = messagebox.askokcancel(
            'Joining Lines Complete', 'Joining Lines is done. Next is step 3: Nest (laminates only)')
        if confirm_joinLines_complete:
            self.select_tab4()

    def create_tab4(self):
        step_3_intro = "Now we will Nest the jobs, but only if its laminate."
        self.intro = Label(self.tab4, text=step_3_intro,
                           font=('Helvetica', 14), justify=CENTER)
        self.intro.grid(column=0, row=1, padx=50, pady=20)

        nesting_button = Button(
            self.tab4, text="Nest for Laminates", command=self.nesting)
        nesting_button.grid(column=0, row=3, padx=50, pady=10)

    def nesting(self):
        confirm_nesting = messagebox.askokcancel(
            'Nesting Laminates', 'Confirm to check if its laminate and start nesting.\n\nREMEMBER: Allow the program to run without touching the mouse or keyboard until popup appears!')
        if confirm_nesting:
            autoTPapp.minimize()
            vcrvApp.maximize()
            simpleNest()
        vcrvApp.minimize()
        autoTPapp.maximize()
        confirm_nesting_complete = messagebox.askokcancel(
            'Nesting Complete', 'Nesting is done. Next is step 4: Tool Setup')
        if confirm_nesting_complete:
            self.select_tab5()

    def create_tab5(self):
        step_4_intro = "The vector lines are grouped into different 'layers' depending on what kind of CNC cut they should be applied to.\nIn this step we will setup the tool settings, mainly that it will default to a height 45mm from the top surface of the material sheet."
        self.intro = Label(self.tab5, text=step_4_intro,
                           font=('Helvetica', 14), justify=CENTER)
        self.intro.grid(column=0, row=1, padx=50, pady=20)

        toolSetup_button = Button(
            self.tab5, text="Tool Setup", command=self.setTool)
        toolSetup_button.grid(column=0, row=3, padx=50, pady=10)

    def setTool(self):
        confirm_set_tool = messagebox.askokcancel(
            'Tool Setup', 'Confirm to setup the tool.\n\nREMEMBER: Allow the program to run without touching the mouse or keyboard until popup appears!')
        if confirm_set_tool:
            autoTPapp.minimize()
            vcrvApp.maximize()
            toolSetup()
        vcrvApp.minimize()
        autoTPapp.maximize()
        confirm_set_tool_complete = messagebox.askokcancel(
            'Tool Setup Complete', 'Tool Setup is done. Next is step 5: Load Toolpath')
        if confirm_set_tool_complete:
            self.select_tab6()

    def create_tab6(self):
        step_5_intro = "Load the necessary toolpath depending on the 'material type' and 'handle type'.\nPlease check the VCarve file and note down the type of material and handle\nand select the right options before commencing to 'Load Toolpath'."
        self.intro = Label(self.tab6, text=step_5_intro,
                           font=('Helvetica', 14), justify=CENTER)
        self.intro.grid(column=0, row=1, padx=50, pady=20)

        material_type_instruction = "Please select the material type."
        self.instruction_material_selector = Label(self.tab6, text=material_type_instruction,
                                                   font=('Helvetica', 14), justify=CENTER)
        self.instruction_material_selector.grid(
            column=0, row=3, padx=50, pady=20)

        self.material_type_selector = ttk.Combobox(self.tab6, width=50)
        self.material_type_selector['values'] = (
            'Birch', 'Oak/Walnut', 'Formica', 'Fenix', 'Urtil Back and Door')
        # self.material_type_selector.current(0)  # set the default selected item
        self.material_type_selector.grid(
            column=0, row=4, padx=10, pady=10)

        handle_type_instruction = "Please select the handle type. You may select more than one if necessary."
        self.instruction_handle_selector = Label(self.tab6, text=handle_type_instruction,
                                                 font=('Helvetica', 14), justify=CENTER)
        self.instruction_handle_selector.grid(
            column=0, row=6, padx=50, pady=50)

        check_state1 = BooleanVar()
        check_state2 = BooleanVar()
        check_state3 = BooleanVar()
        check_state4 = BooleanVar()
        check_state5 = BooleanVar()
        check_state6 = BooleanVar()
        check_state1.set(False)
        check_state2.set(False)
        check_state3.set(False)
        check_state4.set(False)
        check_state5.set(False)
        check_state6.set(False)

        choice_none_handle = Checkbutton(
            self.tab6, text='None/Grab/Circle Handle', var=check_state1, font=('Helvetica', '10'), background='#79ADDA', activebackground='#79ADDA')
        choice_none_handle.grid(
            column=0, row=7, padx=10, pady=10)
        choice_D_handle_choice = Checkbutton(
            self.tab6, text='D-Pull Handle', var=check_state2, font=('Helvetica', '10'), background='#79ADDA', activebackground='#79ADDA')
        choice_D_handle_choice.grid(column=0, row=8, padx=10, pady=10)
        choice_edge_handle = Checkbutton(
            self.tab6, text='Edge-Pull Handle', var=check_state3, font=('Helvetica', '10'), background='#79ADDA', activebackground='#79ADDA')
        choice_edge_handle.grid(column=0, row=9, padx=10, pady=10)
        choice_SRG_handle = Checkbutton(
            self.tab6, text='SRG/SRC Handle', var=check_state4, font=('Helvetica', '10'), background='#79ADDA', activebackground='#79ADDA')
        choice_SRG_handle.grid(column=0, row=10, padx=10, pady=10)
        choice_urtil_body = Checkbutton(
            self.tab6, text='Urtil Body', var=check_state5, font=('Helvetica', '10'), background='#79ADDA', activebackground='#79ADDA')
        choice_urtil_body.grid(column=0, row=11, padx=10, pady=10)
        choice_omlopp = Checkbutton(
            self.tab6, text='Omlopp', var=check_state6, font=('Helvetica', '10'), background='#79ADDA', activebackground='#79ADDA')
        choice_omlopp.grid(column=0, row=12, padx=10, pady=10)

        global choice_dict
        global state_list
        choice_dict = {1: 'None/Grab/Circle Handle', 2: 'D-Pull Handle',
                       3: 'Edge-Pull Handle', 4: 'SRG/SRC Handle', 5: 'Urtil Body', 6: 'Omlopp'}
        state_list = [check_state1, check_state2, check_state3,
                      check_state4, check_state5, check_state6]

        loadTP_button = Button(
            self.tab6, text="Load Toolpath", command=self.loadToolpath)
        loadTP_button.grid(column=0, row=14, padx=50, pady=50)

    def loadToolpath(self):
        handle_selection = []
        for count, check_state in enumerate(state_list, start=1):
            if check_state.get() == True:
                handle_selection.append(choice_dict.get(count))
        material_selection = self.material_type_selector.get()
        confirm_loadTP = messagebox.askokcancel(
            'Load Toolpath', f'{handle_selection} was selected for the {material_selection} material. Confirm to load the toolpath for these selections.\n\nREMEMBER: Allow the program to run without touching the mouse or keyboard until popup appears!')
        if confirm_loadTP:
            autoTPapp.minimize()
            vcrvApp.maximize()
            for handle_choice in handle_selection:
                loadTP(handle_choice, material_selection)
        vcrvApp.minimize()
        autoTPapp.maximize()
        confirm_loadTP_complete = messagebox.askokcancel(
            'Loading Toolpath Complete', 'Loading Toolpath is done. Next is step 6: Save Toolpath')
        if confirm_loadTP_complete:
            self.select_tab7()

    def create_tab7(self):
        step_6_intro = "We will save the finished file in VCarve format and save it on the desktop.\nThen delete that file's DXF format.\nAnd finally move onto the next file for the client if they have multiple files.\n(returning to step 1 for the next file)."
        self.intro = Label(self.tab7, text=step_6_intro,
                           font=('Helvetica', 14), justify=CENTER)
        self.intro.grid(column=0, row=1, padx=50, pady=20)

        saveTP_button = Button(
            self.tab7, text="Save Toolpath", command=self.saveTP)
        saveTP_button.grid(column=0, row=3, padx=50, pady=20)

        self.nextFile_instruction = Label(self.tab7, text="Before going to 'Next File', remember to 'Save Toolpath' (button above)",
                                          font=('Helvetica', 14, 'underline'), justify=CENTER)
        self.nextFile_instruction.grid(
            column=0, row=4, padx=50, pady=100)

        nextFile_button = Button(
            self.tab7, text="Next File", command=self.nextFile)
        nextFile_button.grid(column=0, row=5, padx=50, pady=0)

    def saveTP(self):
        baseDir = 'C:\\Users\\sho25\\Desktop\\extractedFile'
        confirm_saveTP = messagebox.askokcancel(
            'Save Toolpath', 'Confirm to save the current toolpath.\n\nREMEMBER: Allow the program to run without touching the mouse or keyboard until popup appears!')
        if confirm_saveTP:
            autoTPapp.minimize()
            vcrvApp.maximize()
            save_vcrv()
            vcrvMaker()
        vcrvApp.minimize()
        autoTPapp.maximize()
        confirm_set_tool_complete = messagebox.askokcancel(
            'Saving Toolpath Complete', 'Saving Toolpath is done. This was the final step for this file, shall we delete the DXF\nand then move onto the next client file if there is one?\nWe will return to step 1 if that is the case.')
        if confirm_set_tool_complete:
            completeFileMove(baseDir)
            self.select_tab7()

    def nextFile(self):
        baseDir = 'C:\\Users\\sho25\\Desktop\\extractedFile'
        confirm_next_file = messagebox.askokcancel(
            'Open Next File', 'We will check for the next file for the client.\nThen set the job up as step 1 did, and then move to step 2.\n\nREMEMBER: Allow the program to run without touching the mouse or keyboard until popup appears!')
        if confirm_next_file:
            singleFolderAccess(baseDir)
            if toolpath_automater.clientFileEmpty != True:
                jobSetup()
                global vcrvApp
                vcrvApp = pyautogui.getActiveWindow()
                vcrvApp.minimize()
            else:
                confirm_all_complete = messagebox.askokcancel(
                    'Client File Empty', 'Could not find anymore DXF to toolpath in client file. All toolpath complete!')
                if confirm_all_complete:
                    self.select_tab1()
        autoTPapp.maximize()
        confirm_jobSetup_complete = messagebox.askokcancel(
            'Job Setup Complete', 'Job Setup is done. Next is step 2: Joining Lines')
        if confirm_jobSetup_complete:
            self.select_tab3()

    def create_tab8(self):
        full_autoTP_intro = "This tab is to initiate the 'Fully Automatic Toolpath Mode'\nIt is recommended that you understand the step-by-step process before commencing this mode."
        self.intro = Label(self.tab8, text=full_autoTP_intro,
                           font=('Helvetica', 14), justify=CENTER)
        self.intro.grid(column=0, row=1, padx=50, pady=20)

        fullAutoTP_button = Button(
            self.tab8, text="Fully Automatic Toolpathing", command=self.completeAutoTP)
        fullAutoTP_button.grid(column=0, row=3, padx=50, pady=10)

    def completeAutoTP(self):
        confirm_fullAutoTP = messagebox.askokcancel(
            'Starting Fully Automatic Toolpathing', 'Confirm to start the fully automatic toolpathing mode!\n\nREMEMBER: you need to have downloaded multiple client ZIP file from MegaAdmin and onto Desktop.')
        if confirm_fullAutoTP:
            autoTPapp = pyautogui.getActiveWindow()
            autoTPapp.minimize()
            fullAutoTP()


root = autoTP()
root.mainloop()
