##########################
# Contains main program and GUI logic. 
# Contains UI for game settings.
##########################
import sys
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
# import functions to generate bingo board, create simulations and run simulation from already written code.
from bingo_general_game import run_game
#import function to  create histogram and create linegraph
from bingo_output_functions import  create_histogram, create_linegraph
from main_bingo_window import MainGame
import threading
from PIL import ImageTk, Image
import os

class GameStart:
    def __init__(self,main_root=None, top=None):
        self.top = top
        self.main_root=main_root
        self.combobox_value = tk.StringVar()  # combobox value
        self.current_setting_window = 1  # control screens for games

        # the following are variable from already provided code
        self.sUSERDECISION = ""
        self.iCARDS = 0
        self.iSIMULATION = 0
        self.iCOLUMNS = 0
        self.iROWS = 0
        self.iZEROS = 0
        self.iBEGINNINGRANGE = 0
        self.iENDINGRANGE = 0
        self.dBINGO_SHEETS = {}
        self.dSIMULATION_SEQUENCES = {}
        self.df_BINGO_RESULTS = {}

        self.Frame1 = tk.Frame(self.top)  # frame to hold children widget
        self.Frame1.place(relx=0.05, rely=0.15, relheight=0.761, relwidth=0.9)  # place frame
        # frame configurations
        self.Frame1.configure(relief='raised',borderwidth="6",background="#000000",cursor="hand2")

        # -----------------the following are widgets to help input game configurations for game type
        # ................ They  include labels, entry and button widgets ------

        self.Label1_1 = tk.Label(self.Frame1, anchor='w', text='''Settings:''', background="#000000",
                                 font="-family {Sitka Heading} -size 36 -weight bold -slant italic",
                                 foreground="#ffffff")
        self.Label1_1.place(relx=0.028, rely=0.015, height=71, width=235)

        self.Label1_1_1 = tk.Label(self.Frame1, text='''Select Game Type''', anchor='w', foreground="skyblue",
                                   background="#000000",
                                   font="-family {Sitka Heading} -size 36 -weight bold -slant italic")
        self.Label1_1_1.place(relx=0.17, rely=0.268, height=71, width=395)

        self.TCombobox = ttk.Combobox(self.Frame1, state="readonly", textvariable=self.combobox_value)
        self.TCombobox.place(relx=0.341, rely=0.474, relheight=0.05, relwidth=0.246)
        self.TCombobox_values = ['Select Game', 'Standard Game', 'Dynamic Game']
        self.TCombobox['values'] = self.TCombobox_values
        self.TCombobox.current(0)

        self.Button_next = tk.Button(self.Frame1, text='''NEXT''', background="skyblue",
                                     font="-family {Segoe UI} -size 22 -weight bold")
        self.Button_next.place(relx=0.554, rely=0.639, height=44, width=97)
        self.Button_next.configure(foreground="#ffffff", pady="0", command=self.next_setting)

        self.Frame2 = tk.Frame(self.top, relief='groove', borderwidth="2", background="#000000")
        self.Frame2.place(relx=0.05, rely=0.0, relheight=0.149, relwidth=0.9)

        self.Label1 = tk.Label(self.Frame2, text='''Welcome to Bingo Board Game''',anchor="center")
        self.Label1.place(relx=0.01, rely=0.3, relheight=0.5, relwidth=0.98)
        self.Label1.configure(anchor='w', background="#000000", foreground="skyblue",
                              font="-family {Sitka Heading} -size 36 -weight bold -slant italic")

    # ----------------------------------------------------end of input widgets for game type..........................
    # --------------------------------------------------- end of input widgets .......................................

    def get_TCombobox_value(self):  # function to get combobox value and return
        selected_value = self.TCombobox.current()
        return selected_value

    def show_selected(self):
        if self.current_setting_window == 1:
            print(str(self.sUSERDECISION) + " game mode selected.")
            messagebox.showinfo("Message", str(self.sUSERDECISION) + " game mode selected.")

    def validate_inputs(self,temp_dict):
        ireturn=0
        for i in temp_dict.keys():
            if temp_dict[i].isnumeric() == False:
                print(
                    '--------------------------------------------------------------------------------------------')
                print(
                    'You have entered in a none integer for number of ' + i + '. Re-enter an integer in this field.')
                print(
                    '-------------------------------------------------------------------------------------------------')
                messagebox.showerror("Error Message",
                                     '''------------------------------------------------------------------------
                                     \nYou have entered in a none integer for number of ''' + i + '''. Re-enter an integer in this field.
                                     \n------------------------------------------------------------------------''')
                ireturn =1
                break
            elif int(temp_dict[i]) <= 0:
                print(
                    '-----------------------------------------------------------------------------------------------------')
                print('You have inputted a number less than 0 for ' + i + '. Insert a number greater than 0.')
                print(
                    '-----------------------------------------------------------------------------------------------------')
                messagebox.showerror("Error Message",
                                     '''------------------------------------------------------------------------
                                     \nYou have inputted a number less than 0 for ''' + i + '''. Insert a number greater than 0.
                                 \n------------------------------------------------------------------------''')
                ireturn = 1
                break
            elif int(temp_dict[i]) > 1000:
                print(
                    '-----------------------------------------------------------------------------------------------------')
                print('You have inputted a number greater than 1000 for ' + i + '. Insert a number less than 1000.')
                print(
                    '-----------------------------------------------------------------------------------------------------')
                messagebox.showerror("Error Message",
                                     '''------------------------------------------------------------------------
                                     \nYou have inputted a number greater than 1000 for ''' + i + '''. Insert a number less than 1000..
                                 \n------------------------------------------------------------------------''')
                ireturn = 1
                break
        return ireturn

    def next_setting(self):  # function called when next button at start menu is called
        if self.current_setting_window == 1:  # 1 represents window with game type selection
            if self.get_TCombobox_value() == 1:  # 1 represents Standard game selection
                self.sUSERDECISION = 'Standard'
                self.show_selected()
                # ----- the following are widgets to help input game configurations for Standard Game type............
                # ----- They include labels, entry and button widgets ...............................................
                self.Label1_1_1.configure(text='Cards Number')
                self.TCombobox.place_forget()
                self.card = tk.Entry(self.Frame1, background="white",
                                     font="-family {Courier New} -size 13 -weight bold", borderwidth="4",
                                     foreground="#000000", insertbackground="black", relief="groove")
                self.card.place(relx=0.65, rely=0.33, relheight=0.05, relwidth=0.246)
                self.Label_simulation = tk.Label(self.Frame1, text='''Simulations''', anchor='w', foreground="skyblue",
                                                 background="#000000",
                                                 font="-family {Sitka Heading} -size 36 -weight bold -slant italic")
                self.Label_simulation.place(relx=0.17, rely=0.4, height=50, width=395)
                self.simulation = tk.Entry(self.Frame1, background="white",
                                           font="-family {Courier New} -size 13 -weight bold", borderwidth="4",
                                           foreground="#000000", insertbackground="black", relief="groove")
                self.simulation.place(relx=0.65, rely=0.46, relheight=0.05, relwidth=0.246)
                self.current_setting_window = 2
                self.Button_next.configure(text="Run")
                # -----------------end of standard game configuration widgets........................................
                # ----------------- end of standard game configuration widgets ......................................
            elif self.get_TCombobox_value() == 2:  # if Dynamic game is selected
                self.sUSERDECISION = 'Dynamic'
                self.show_selected()
                # --------------the following are widgets to help input game configurations for Dynamic Game type......
                # --------------They include labels, entry and button widgets ........................................
                self.Label1_1_1.place_forget()
                self.Label1_1_1.place(relx=0, rely=0.15, height=50, width=250)
                self.Label1_1_1.configure(text='Cards')
                self.TCombobox.place_forget()
                self.card = tk.Entry(self.Frame1, background="white",
                                     font="-family {Courier New} -size 13 -weight bold", borderwidth="4",
                                     foreground="#000000", insertbackground="black", relief="groove")
                self.card.place(relx=0, rely=0.3, relheight=0.05, relwidth=0.246)

                self.Label_simulation = tk.Label(self.Frame1, text='''Simulations''', anchor='w', foreground="skyblue",
                                                 background="#000000",
                                                 font="-family {Sitka Heading} -size 36 -weight bold -slant italic")
                self.Label_simulation.place(relx=0.3, rely=0.15, height=50, width=270)
                self.simulation = tk.Entry(self.Frame1, background="white",
                                           font="-family {Courier New} -size 13 -weight bold", borderwidth="4",
                                           foreground="#000000", insertbackground="black", relief="groove")
                self.simulation.place(relx=0.3, rely=0.3, relheight=0.05, relwidth=0.246)

                self.Label_column = tk.Label(self.Frame1, text='''Columns''', anchor='w', foreground="skyblue",
                                             background="#000000",
                                             font="-family {Sitka Heading} -size 36 -weight bold -slant italic")
                self.Label_column.place(relx=0.65, rely=0.15, height=50, width=200)
                self.column = tk.Entry(self.Frame1, background="white",
                                       font="-family {Courier New} -size 13 -weight bold", borderwidth="4",
                                       foreground="#000000", insertbackground="black", relief="groove")
                self.column.place(relx=0.65, rely=0.3, relheight=0.05, relwidth=0.246)

                self.Label_row = tk.Label(self.Frame1, text='''Rows''', anchor='w', foreground="skyblue",
                                          background="#000000",
                                          font="-family {Sitka Heading} -size 36 -weight bold -slant italic")
                self.Label_row.place(relx=0, rely=0.35, height=50, width=200)
                self.row = tk.Entry(self.Frame1, background="white", font="-family {Courier New} -size 13 -weight bold",
                                    borderwidth="4", foreground="#000000", insertbackground="black", relief="groove")
                self.row.place(relx=0, rely=0.5, relheight=0.05, relwidth=0.246)

                self.Label_zero = tk.Label(self.Frame1, text='''Free spaces''', anchor='w', foreground="skyblue",
                                           background="#000000",
                                           font="-family {Sitka Heading} -size 36 -weight bold -slant italic")
                self.Label_zero.place(relx=0.3, rely=0.35, height=50, width=270)
                self.zero = tk.Entry(self.Frame1, background="white",
                                     font="-family {Courier New} -size 13 -weight bold", borderwidth="4",
                                     foreground="#000000", insertbackground="black", relief="groove")
                self.zero.place(relx=0.3, rely=0.5, relheight=0.05, relwidth=0.246)

                self.Label_brange = tk.Label(self.Frame1, text='''Start Range''', anchor='w', foreground="skyblue",
                                             background="#000000",
                                             font="-family {Sitka Heading} -size 36 -weight bold -slant italic")
                self.Label_brange.place(relx=0.65, rely=0.35, height=70, width=260)
                self.brange = tk.Entry(self.Frame1, background="white",
                                       font="-family {Courier New} -size 13 -weight bold", borderwidth="4",
                                       foreground="#000000", insertbackground="black", relief="groove")
                self.brange.place(relx=0.65, rely=0.5, relheight=0.05, relwidth=0.246)

                self.Label_erange = tk.Label(self.Frame1, text='''End Range''', anchor='w', foreground="skyblue",
                                             background="#000000",
                                             font="-family {Sitka Heading} -size 36 -weight bold -slant italic")
                self.Label_erange.place(relx=0, rely=0.6, height=70, width=230)
                self.erange = tk.Entry(self.Frame1, background="white",
                                       font="-family {Courier New} -size 13 -weight bold", borderwidth="4",
                                       foreground="#000000", insertbackground="black", relief="groove")
                self.erange.place(relx=0, rely=0.75, relheight=0.05, relwidth=0.246)
                self.current_setting_window = 2
                self.Button_next.configure(text="Run")
                # -----------------end of Dynamic game configuration widgets........................................
                # ----------------- end of Dynamic game configuration widgets ........................................

        elif self.current_setting_window == 2:  # 1 represents window with game values inputs like no. of simulations
            if self.sUSERDECISION == 'Standard':
                # -----------------get Standard game inputs and check for errors........................................
                cards, simulations = self.card.get(), self.simulation.get()
                temp_dict = {'cards': cards, 'simulations': simulations}
                ivalidate = self.validate_inputs(temp_dict=temp_dict)
                if ivalidate == 0:
                    self.iCARDS = int(cards)
                    self.iSIMULATION = int(simulations)
                    self.playGame()  # call this functions to direct to run game if all inputs are ok


            elif self.sUSERDECISION == 'Dynamic':
                # -----------------get Dynamic game inputs and check for errors........................................
                cards, simulations = self.card.get(), self.simulation.get()
                zeros, column, rows, startRange, endRange = self.zero.get(), self.column.get(), self.row.get(), self.brange.get(), self.erange.get()

                temp_dict = {'cards':cards, 'simulations':simulations,'Free spaces': zeros, 'columns': column, 'rows': rows,'Start Range':startRange, 'End Range':endRange}
                ivalidate = self.validate_inputs(temp_dict=temp_dict)
                if ivalidate == 0:
                    # self.main_window()  # call this function to direct to run game if all inputs are ok

                    irange = (int(endRange) - int(startRange))+1

                    if int(startRange) >= int(endRange):
                        print(
                            '-----------------------------------------------------------------------------------------------------')
                        print(
                            'Starting range number can not be greater than or equal to the end, select an appropriate range ')
                        print(
                            '-----------------------------------------------------------------------------------------------------')

                        messagebox.showerror("Error Message",
                                             '''------------------------------------------------------------------------
                                             \nStarting range number can not be greater than or equal to the end, select an appropriate range 
                                             \n------------------------------------------------------------------------''')
                    if (int(rows) * int(column)) <= int(zeros):
                        print(
                            '-----------------------------------------------------------------------------------------------------')
                        print(
                            'Number of spaces cannot be equal or greater than board size')
                        print(
                            '-----------------------------------------------------------------------------------------------------')

                        messagebox.showerror("Error Message",
                                             '''------------------------------------------------------------------------
                                             \nNumber of spaces cannot be equal or greater than board size
                                             \n------------------------------------------------------------------------''')
                    elif (int(rows) * int(column)) > irange:
                        print(
                            '-----------------------------------------------------------------------------------------------------')
                        print(
                            'Inputted starting to ending range is smaller than Free spaces on bingo board. Enter in a larger range.  ')
                        print(
                            '-----------------------------------------------------------------------------------------------------')

                        messagebox.showerror("Error Message",
                                             '''------------------------------------------------------------------------
                                             \nInputted starting to ending range is smaller than Free spaces on bingo board. Enter in a larger range.  
                                             \n------------------------------------------------------------------------''')

                    else:

                        self.iCARDS = int(cards)
                        self.iSIMULATION = int(simulations)
                        self.iCOLUMNS = int(column)
                        self.iROWS = int(rows)
                        self.iZEROS = int(zeros)
                        self.iBEGINNINGRANGE = int(startRange)
                        self.iENDINGRANGE = int(endRange)
                        if self.iROWS == 1:  # prevent single rows from generating error
                            if int(irange)/int(column)<int(column):
                                self.iENDINGRANGE = int(column)*int(column)


                        self.playGame()  # call this function to direct to run game if all inputs are ok

    def playGame(self):
        def show_frontend():
            for i in self.Frame1.winfo_children():
                i.destroy()
            self.Label_run = tk.Label(self.Frame1, text='''Game is Generating Boards,Cards, \nSimulations and Analysis for next Window\nPlease wait ...
            ''', anchor='w', foreground="white",background="#000000")
            self.Label_run.place(relx=0.3, rely=0.2,relheight=0.6, relwidth=0.6)

            self.pb = ttk.Progressbar(self.Frame1,orient='horizontal',mode='indeterminate',length=280)
            self.pb.place(relx=0.1, rely=0.9, relheight=0.05, relwidth=0.8)
            self.pb.start()


        def start_running_game():
            self.df_BINGO_RESULTS,self.dBINGO_SHEETS,self.dSIMULATION_SEQUENCES = run_game(sUSERDECISION=self.sUSERDECISION, iCARDS=self.iCARDS,
                                                                 iSIMULATION=self.iSIMULATION, iROWS=self.iROWS, iCOLUMNS=self.iCOLUMNS,
                                                                 iZEROS=self.iZEROS, iBEGINNINGRANGE=self.iBEGINNINGRANGE,
                                                                 iENDINGRANGE=self.iENDINGRANGE)
        def test_thread_running():
            while True:
                if self.thread_backgroud.is_alive():
                    time.sleep(5)
                else:
                    eval("self.main_root.after(0,self.new_window)")  # call next window
                    break
        self.thread_backgroud = threading.Thread(target=start_running_game,daemon=True)
        self.thread_frontend = threading.Thread(target= show_frontend, daemon=True)
        self.thread_check = threading.Thread(target=test_thread_running, daemon=True)
        self.thread_backgroud.start()
        self.thread_frontend.start()
        self.thread_check.start()




    def new_window(self):
        for i in self.top.winfo_children():
            i.destroy()

        left_frame = tk.Frame(self.top)  # create new frame and place it on left
        left_frame.place(relx=0, rely=0, relheight=0.9, relwidth=0.7)
        right_frame = tk.Frame(self.top)  # create new frame and place it on right
        right_frame.place(relx=0.5, rely=0, relheight=0.9, relwidth=0.7)
        self.left_side = MainGame(left_frame)  # create new instance of scrolled window and place it on left frame
        self.right_side = MainGame(right_frame)  # create new instance of scrolled window and place it on right frame

        # create line graph and histogram by calling create_linegraph and create_histogram functions and place them on right instance of scrolled window
        create_histogram(root=self.right_side.Scrolledwindow1_f, df_data=self.df_BINGO_RESULTS,
                         dSIMULATION_SEQUENCES=self.dSIMULATION_SEQUENCES)

        create_linegraph(root=self.right_side.Scrolledwindow1_f, df_data=self.df_BINGO_RESULTS,
                         dSIMULATION_SEQUENCES=self.dSIMULATION_SEQUENCES)

        tk.Label(self.left_side.Scrolledwindow1_f, text='testing ').place(relx=0.1, rely=0, height=200,
                                                                                     width=200)

        for i in range(0,self.iCARDS):
            cwd = os.getcwd()
            image_name ='df_image{0}_new.png'.format(str(i))
            image_path = os.path.join(cwd,  image_name)
            # Create a Label Widget to display the text or Image

            img = Image.open(image_path)
            # img = img.resize((400, 450), Image.ANTIALIAS) Allows for image resizing
            img = ImageTk.PhotoImage(img)
            panel = tk.Label(self.left_side.Scrolledwindow1_f, image=img,padx=20,pady=20)
            panel.image = img
            panel.pack()
            os.remove(image_path)
