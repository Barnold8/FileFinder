import os
import win32api

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

def StartSequence():
    inputString = "----Please pick a drive to search---\n\n----The drives to choose from are----\n\n"
    for i in range(len(drives)):
        inputString += f"{i} "+ drives[i] + "\n"
    inputString += "\n"
    userInput = input(inputString)
    try:
        x = int(userInput)
        return int(x)
    except Exception as e:
        print("An error has occurred! You may have not entered a number!")
        return


uInput = StartSequence()
filetype = input("What filetype do you wish to find?\n\nNOTE: don't include the . at the start of the filetype!\nExample: docx, exe, pdf\n")
filename = input("Finally, please write the name of the file you wish to write to\n\n")
x = [f for f in WorkDrive(drives[uInput],filetype)]


file = open(f"{filename}.txt","w")
file.write("---------------------------------Here's what we found---------------------------------\n\n")
for i in range(len(x)):
    file.write(x[i]+"\n")
file.close()

os.system(f"start {filename}.txt")



#Waste code:


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
