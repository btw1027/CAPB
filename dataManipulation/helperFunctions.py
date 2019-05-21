from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory, asksaveasfilename

#From http://stackoverflow.com/a/14119223/6816646
# with askdirectory() to request folder path
def setPath(filename='.dataFolderPath'):
    win = Tk()
    win.withdraw()
    win.wm_attributes('-topmost', 1)
    data_folder_path= askdirectory(parent=win)
    if len(data_folder_path)>0:
        with open(filename,'w') as f:
            f.write(data_folder_path)
    win.quit()
    return data_folder_path

def get_file():
    win = Tk()
    win.withdraw()
    win.wm_attributes('-topmost', 1)
    filename = askopenfilename(parent=win)
    return filename

def get_folder():
    win = Tk()
    win.withdraw()
    win.wm_attributes('-topmost', 1)
    folder = askdirectory(parent=win)
    return folder

def saveas_filename():
    win = Tk()
    win.withdraw()
    win.wm_attributes('-topmost', 1)
    save_filename = asksaveasfilename(title='What would you like to name your excel file?',defaultextension = 'xlsx')
    return save_filename