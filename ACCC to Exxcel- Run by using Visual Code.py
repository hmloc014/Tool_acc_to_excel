import os,sys
import Tkinter as tk
import tkMessageBox
import tkFileDialog
import glob
from Tkconstants import INSERT

PSSE_PATH = r'C:\Program Files (x86)\PTI\PSSE33\PSSBIN'
sys.path.append(PSSE_PATH)
os.environ['PATH'] = os.environ['PATH'] + ';' + PSSE_PATH
import pssexcel
#filename = "25-K-MAX",

def selectfolder():
    folder_selected = tkFileDialog.askdirectory()
    outputbar.insert(INSERT,' Begin Converted - Please Wait---\n')
    root.update_idletasks()
    print(folder_selected)
    os.chdir(folder_selected)
    accfiles = glob.glob('*.acc')
    for i in range(len(accfiles)):
        acc_file_name = accfiles[i]
        xlsx_file_name = accfiles[i][:-4].strip() +".xlsx"
        pssexcel.accc(  accfile = acc_file_name,
                        string  = ['b','e'],
                        colabel = '', 
                        stype   = 'contingency',
                        busmsm  = 0.5,
                        sysmsm  = 5.0,
                        rating  = 'a',
                        namesplit = True,
                        ratecon = 'a',
                        baseflowvio = False,
                        basevoltvio = False,
                        flowlimit   = 0,
                        flowchange  = 0.0,
                        voltchange   = 0.0,
                        xlsfile = xlsx_file_name,
                        sheet   = '',
                        overwritesheet = True,
                        show    = False,    
                    )
        outputbar.insert(INSERT,' {} Converted Success\n'.format(accfiles[i]))
        root.update_idletasks()
    outputbar.insert(INSERT,' Finish Converted\n')
    root.update_idletasks()
    tkMessageBox.showinfo(title="Finish", message="Finish Converted")
def quit():
    global root
    root.destroy()  
root = tk.Tk()
root.title = ("Acc to excel ")
sub_btn4=tk.Button(root, text="Close Program", command= quit,background='red',fg='white')
sub_btn4.grid(row=2,column=0)

sub_btn2=tk.Button(text = 'Selected the .acc folder', command = selectfolder,background='blue',fg='white')
sub_btn2.grid(row=1,column=0)
help_label2 = tk.Label(root, text = 'Multiple ACC to XLS Running - Please Close Excel ', font=('calibre',10, 'bold'))

help_label2.place(x=0,y=100)
help_label3 = tk.Label(root, text = 'Power System Dept-TRD-Tran Huu Phuc - EVNPECC2', font=('calibre',10, 'italic'))
help_label3.place(x=0,y=125)
outputbar = tk.Text(root,height=15, width=50)
outputbar.place(x=0,y=150)
root.geometry("420x475")
root.mainloop()
