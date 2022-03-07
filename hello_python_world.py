import sys
import time
import asyncio

import pythoncom

import win32com.client

# Solidworks Version (write the year) :
SWV = 2022
# API Version
SWAV = SWV-1992
# Elapsed time
elapsed = 0

async def main():
    sw = win32com.client.Dispatch("SldWorks.Application.{}".format(SWAV)) 
    print(f" Solidworks API Version : {SWAV}","\n",f"Solidworks Version : {SWV}")
    Model = sw.ActiveDoc
    ARG_NULL = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    ck = Model.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, False, 0, ARG_NULL, 0)
    Model.SketchManager.InsertSketch(True)
    Model.ClearSelection2(True)

    mySketchText = Model.InsertSketchText(
        0, 0, 0, "Hello Python World!", 0, 0, 0, 100, 100
    )
    myFeature = Model.FeatureManager.FeatureExtrusion2(
        True, 
        False,
        False,
        0,
        0,
        0.001,
        0.001,
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
    s = time.perf_counter()
    elapsed = time.perf_counter() - s
    print(f"{__file__} executed in {elapsed:0.2f} seconds until while-loop")
    time.sleep(2)
    while True:
        Model.ViewRotateplusy()
        time.sleep(0.1)


try:
    if __name__ == "__main__":
        asyncio.run(main())
except KeyboardInterrupt:
    sys.exit()
