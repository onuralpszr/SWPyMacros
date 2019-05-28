import sys
import time
import pythoncom
import win32com.client
import tkinter as tk
from tkinter import *

# Solidworks Version (2019)
swYearLastDigit = 9


# Consts
sw = win32com.client.Dispatch(
    "SldWorks.Application.{}".format((20 + (swYearLastDigit - 2)))
)  # e.g. 20 is SW2012,  27 is SW2019
sw.SetUserPreferenceToggle(1, False)
SWAV = 20 + swYearLastDigit - 2
SWV = 2010 + swYearLastDigit
print(f" Solidworks API Version : {SWAV}", "\n", f"Solidworks Version : {SWV}")
Model = sw.ActiveDoc


class Application(tk.Frame):
    def cube(self):
        ARG_NULL = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        Model.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, False, 0, ARG_NULL, 0)
        Model.SketchManager.InsertSketch(True)
        Model.ClearSelection2(True)
        Model.SetPickMode()
        Model.SketchManager.CreateCenterRectangle(0, 0, 0, 0.05, 0.05, 0)
        Model.ClearSelection2(True)
        Model.ShowNamedView2("*Trimetric", 8)
        Model.ViewZoomtofit2()
        Model.FeatureManager.FeatureExtrusion2(
            True,
            False,
            False,
            0,
            0,
            0.1,
            0.01,
            False,
            False,
            False,
            False,
            0,
            0,
            False,
            False,
            False,
            False,
            True,
            True,
            True,
            0,
            0,
            False,
        )
        Model.SelectionManager.EnableContourSelection = False
        Model.ClearSelection2(True)
        swDim = Model.Parameter("D1@Boss-Extrude1")
        swDim.SetSystemValue3(0.4, 1)
        Model.EditRebuild3

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.btnone = tk.Button(self)
        self.btnone["text"] = "Create a Cube !"
        self.btnone["command"] = self.cube
        self.btnone.pack(anchor="center")

        self.scale = tk.Scale(
            self, orient="horizontal", from_=0, to=60000, variable=D1_value
        )
        self.scale.pack(anchor="center")

        self.button = tk.Button(
            self,
            text="Set Cube Value",
            command=lambda: Application.change_value(int(D1_value.get())),
        )
        self.button.pack(anchor="center")

        self.quit = tk.Button(self, text="QUIT", fg="red", command=self.master.destroy)
        self.quit.pack(anchor="center")

    def change_value(val):
        time.sleep(1)
        swDim = Model.Parameter("D1@Boss-Extrude1")
        swDim.SetSystemValue3((val / 1000), 1)
        Model.EditRebuild3
        time.sleep(0.5)


try:
    root = tk.Tk()
    root.title("Solidworks Python Cube GUI")
    root.geometry("500x200")
    root.resizable(0, 0)  # No Resize X,Y directions
    D1_value = DoubleVar()
    app = Application(master=root)
    app.mainloop()
except KeyboardInterrupt:
    sys.exit()
