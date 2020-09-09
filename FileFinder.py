import os                                           #NOTE: I am not using any OOP because this is a very small project and let's be honest, it's not needed.
import win32api
from tkinter import *
import tkinter.messagebox

import threading
drives = win32api.GetLogicalDriveStrings()
drives = drives.split('\000')[:-1]

#compile error fix


def import_pywin32():
    """
    For the module ``pywin32``,
    this function tries to add the path to the DLL to ``PATH``
    before throwing the exception:
    ``DLL load failed: The specified module could not be found``.
    """
    try:
        import win32com
    except ImportError as e :
        if "DLL load failed:" in str(e):
            import os,sys
            path = os.path.join(os.path.split(sys.executable)[0], "Lib","site-packages","pywin32_system32")
            os.environ["PATH"] = os.environ["PATH"] + ";" + path
            try:
                import win32com
            except ImportError as ee :
                dll = os.listdir(path)
                dll = [os.path.join(path,_) for _ in dll if "dll" in _]
                raise ImportError("some DLL must be copied:\n" + "\n".join(dll)) from e
        else :
            raise e

import_pywin32()

def WorkDrive(DriveLetter,filetype):
    for root, dirs, files in os.walk(DriveLetter):
        for name in files:
            f = name.split(".")
            try:
                if f[1] == filetype:
                    print(f"{root}" + rf"\{name}")
                    yield f"{root}" + rf"\{name}"
                else:
                    pass
            except IndexError as e:
                #print(f"File {name} has no extenstion")
                pass

#Start sequence


#CODE ABOVE NEEDS TO DONE BY GUI



def WriteFile(filename, x):
    file = open(f"{filename}.txt","w")

    if len(x) <=0:
        file.write(f"Sorry!\n\nNothing was found on this drive that follows the parameters: {filetype}")
        file.close()

    else:
        file.write("---------------------------------Here's what we found---------------------------------\n\n")
        for i in range(len(x)):
            file.write(x[i]+"\n")
        file.close()

    os.system(f"start {filename}.txt")
    New.destroy()

# uInput = StartSequence()
# filetype = input("What filetype do you wish to find?\n\nNOTE: don't include the . at the start of the filetype!\nExample: docx, exe, pdf\n")
# filename = input("Finally, please write the name of the file you wish to write to\n\n")

#-------------------------------------------------------GUI CODE-------------------------------------------------------#

def CREATE_GUI(bg="#000000", fg="#FFFFFF", abg="#AF2FFF", afg="#222222"):

    def ClosingProto():

        ##CleanUpMethods

        exit()
        sys.exit()

    #I am regretting more and more not using classes for my GUI now..... OOPS
    def LoadingWindow():
        global New
        New = Tk()
        New.config(bg=bg)
        New.resizable(0,0)
        New.geometry("250x250")
        New.title("Processing...")
        New.protocol("WM_DELETE_WINDOW",ClosingProto)
        Processing_Label = Label(text="Scanning for files...\nPlease wait!",bg=bg,fg=fg)
        Processing_Label.place(relx=0.27,rely=0.35)


        New.mainloop()
    def LoadingWindowThread():
        threading.Thread(target=LoadingWindow).start()
    def ButtonGrabber():
        global filetype
        chosen_drive = strVar.get()
        filetype = FileType.get()
        filename = FileName.get()

        if len(filetype) <=0 or len(filename) <= 0:
            print(len(filetype) ,len(filename))
            tkinter.messagebox.showerror("Warning!","Nothing was entered into one of the text fields!\nPlease try again.")

        else:
            Window.destroy()

            LoadingWindowThread()


            x = [f for f in WorkDrive(chosen_drive,filetype)]
            WriteFile(filename, x)

    Window = Tk()

    Window.title("FileFinder")
    Window.resizable(0,0)
    Window.geometry("450x450+400+200")
    Window.config(bg=bg)

    EnterButton = Button(Window, width=10,height =1 ,text="Find files",command=ButtonGrabber,bg=bg,fg=fg,activeforeground=afg,activebackground=abg)
    EnterButton.pack()
    EnterButton.place(relx=0.4,rely=0.8)



    FileType = Entry(width=10)
    FileType.pack()
    FileType.place(relx=0.425,rely=0.38)


    FileName = Entry(width=20)
    FileName.pack()
    FileName.place(relx=0.368,rely=0.6)

    #Dropdown for drives
    strVar = StringVar()
    strVar.set(drives[0])

    #Labels
    ChooseDrive = Label(text="Storage drive to search through",bg=bg, fg=fg)
    ChooseDrive.pack()
    ChooseDrive.place(relx=0.29,rely=0.0)
    ChooseExtension = Label(text="What file type do you wish to find?",bg=bg,fg=fg)
    ChooseExtension.pack()
    ChooseExtension.place(relx=0.29,rely=0.31)
    FileNameLabel = Label(text="What do you wish the output file name to be?",bg=bg,fg=fg)
    FileNameLabel.pack()
    FileNameLabel.place(relx=0.25,rely=0.51)

                                    #Underline Labels
    UL_Text = "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Underline = Label(bg=bg,fg=fg,text=UL_Text)
    Underline.place(relx=0.0,rely=0.05)
    Underline2 = Label(bg=bg,fg=fg,text=UL_Text)
    Underline2.place(relx=0.0,rely=0.25)
    Underline3 = Label(bg=bg,fg=fg,text=UL_Text)
    Underline3.place(relx=0.0,rely=0.45)
    Underline4 = Label(bg=bg,fg=fg,text=UL_Text)
    Underline4.place(relx=0.0,rely=0.7)

    #------

    Options = OptionMenu(Window, strVar, *drives)
    Options.pack()
    Options.config(bg=bg,fg=fg,activeforeground=afg,activebackground=abg)
    Options["highlightthickness"] = 0
    Options.place(relx=0.43,rely=0.1)


    Window.mainloop()




#----------------------------------------------------------------------------------------------------------------------#

CREATE_GUI()
#Waste code:

# def StartSequence():
#     inputString = "----Please pick a drive to search---\n\n----The drives to choose from are----\n\n"
#     for i in range(len(drives)):
#         inputString += f"{i} "+ drives[i] + "\n"
#     inputString += "\n"
#     userInput = input(inputString)
#     try:
#         x = int(userInput)
#         return int(x)
#     except Exception as e:
#         print("An error has occurred! You may have not entered a number!")
#         return
#
# other_found_files = []
# def CurrentFiles():
#     dirs = [f for f in os.listdir(os.curdir) if
#             os.path.isfile(f)]  # Returns current files from current directory into an array
#     return dirs  # returns said array
#
# def CurrentDirs():
#     dirs = [f for f in os.listdir(os.curdir) if os.path.isdir(f)]
#     return dirs
#
# def GrabFiles(*args):
#     FileArray = []
#     for x in args:
#         for y in x:
#             FileArray.append(y)
#     return FileArray
#
# def filterForEXE(array):
#     EXEARRAY = []
#
#     for f in array:
#         y = f.split(".")
#         if y[1].lower() == "exe":
#             EXEARRAY.append(f)
#     if len(EXEARRAY) <= 0:
#         return f"No EXE files found\n\nHowever, the following was found {array}"
#     else:
#         print(f"Length of exe array is {len(EXEARRAY)}")
#         return EXEARRAY
#
