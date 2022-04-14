import shutil
import os
from datetime import datetime as dt
def SpeedupCode():
    x="BACKUPS/"+dt.now().strftime("%m-%d-%Y-%H:%M:%S")
    os.mkdir(x)
    try:
        shutil.move('../FRT.CAL.xlsx',x)
    except Exception as e:
            print("Error"+str(e))
    try:
        shutil.move('../FirmsSorted.xlsx',x)
    except Exception as e:
        print("Error"+str(e))
    try:
            shutil.move('../BANKS.xlsx',x)
    except Exception as e:
        print("Error"+str(e))
    try:
            shutil.move('vchrs.csv',x)
    except Exception as e:
        print("Error"+str(e))

