import os
import pandas as pd
from tkinter import Tk, filedialog
from mylib import getUniqueDataFrame

def getDir():
    print("Work Dir:", os.getcwd())

def setDir():
    root = Tk()
    root.withdraw()
    indir = filedialog.askdirectory(title="请选择工作目录：", initialdir=os.getcwd())
    if len(indir):
        os.chdir(indir)
    root.destroy()
    getDir()

def getUniqueJoin():
    root = Tk()
    root.withdraw()
    infile = filedialog.askopenfilename(title="请选择文件：", initialdir=os.getcwd(), filetypes=(("CSV File", "*.csv;*.txt;"),))
    if len(infile):
        df = pd.read_csv(infile)
        df = getUniqueDataFrame(df, df.columns[0], df.columns[1])
        fpath, fext = os.path.splitext(infile)
        fname = fpath + "_unique" + fext
        if df[df.columns[0]].is_unique and df[df.columns[1]].is_unique:
            df.to_csv(fname, index=False)
            print("Saved as "+ fname)
        else:
            print("Failed to deal with " + infile)
    else:
        print("No input file!")
    root.destroy()

if __name__ != "main":
    getDir()