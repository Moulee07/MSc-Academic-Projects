##########################
# Run program from here
# Contains code to reset and exit program
##########################
import sys,os
from tkinter import *  # python gui module
from bingo_ui import GameStart  # import class for game setup gui
from tkinter import messagebox

root = Tk()  # create root window
root.geometry("939x637+215+17")
root.minsize(120, 1)  # minimum root window size
root.maxsize(1370, 749)  # maximum root window size
root.resizable(1, 1)  # resizable
root.title("Bingo Board Game")  # window title
root.configure(background="#000000")  # background colour, black

root_frame = Frame(bg='red')
root_frame.place(relx=0.01, rely=0.01, relheight=0.9, relwidth=0.99)


def restart_game():
    for i in root_frame.winfo_children():
        i.destroy()
    GameStart(main_root=root, top=root_frame)  # start class




def exit_game():
    ask = messagebox.askyesno(title='QUIT', message='ARE YOU SURE YOU WANT TO QUIT?')
    if ask == True:
        try:
            root.destroy()
            os._exit(0)  # exit and stop threads
        except RuntimeError:  # prevent matplotlib thread error when closing
            os._exit(0)  # exit and stop threads


button_exit = Button(root, text='''EXIT GAME''', foreground="#ffffff", pady="0", background="skyblue",
                     font="-family {Segoe UI} -size 22 -weight bold", command=exit_game)
button_exit.place(relx=0.2, rely=0.92, height=44, width=250)

button_restart = Button(root, text='''RESTART GAME''', foreground="#ffffff", pady="0", background="skyblue",
                        font="-family {Segoe UI} -size 22 -weight bold", command=restart_game)
button_restart.place(relx=0.5, rely=0.92, height=44, width=250)

restart_game()
import warnings  # ignore deprecation warnings
warnings.filterwarnings("ignore")
root.protocol("WM_DELETE_WINDOW", exit_game)  # do not close window directly or by mistake
root.mainloop()  # event loop
