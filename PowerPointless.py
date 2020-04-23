# Created by Philip Goosen github.com/goosenphil

import win32com.client
import time
import os
import zipfile
from ppsx_patcher import patch_ppsx
import os.path
# Only for Windows computers running PowerPoint 2007 or newer

use_voice_times = 15
slide_duration = 1 # only matters if use_voice_times = 0
resolution = 720 # Vertical resolution, default at 720p
frame_rate = 24
quality = 60

def pptx_to_mp4(pptx_input,mp4_output):
    print("Converting", pptx_input, "...")
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    presentation = ppt.Presentations.Open2007(pptx_input,WithWindow=False,OpenAndRepair=True)
    presentation.CreateVideo(mp4_output,use_voice_times,slide_duration,resolution,frame_rate,quality)
    start_time_stamp = time.time()
    
    bar = [
    " [=     ]",
    " [ =    ]",
    " [  =   ]",
    " [   =  ]",
    " [    = ]",
    " [     =]",
    " [    = ]",
    " [   =  ]",
    " [  =   ]",
    " [ =    ]",
    ]
    i = 0

    while True:
        time.sleep(0.5)
        try:
            os.rename(mp4_output,mp4_output) # If we can rename the file, the conversion has completed
            print(' '*24 , end="\r") # Clear loading bar
            break
        except Exception:
            print(bar[i % len(bar)],"converting...", end="\r")
            i += 1
            pass
    end_time_stamp=time.time()
    print("Conversion time:", round(end_time_stamp-start_time_stamp, 3), "seconds")
    ppt.Quit()
    pass
  
if __name__ == '__main__':
    print("\nPowerPointless V1.0 (PowerPoint to video converter) by Philip Goosen [19509766@sun.ac.za]\n")
    cwd = os.getcwd() + '\\'
    
    files = os.listdir(os.curdir)
    for file in files:
        if ".ppsx" in file:
            print("Detected ppsx file! -", file, " creating pptx patched version...")
            patch_ppsx(file)
    
    files = os.listdir(os.curdir)
    ppt_counter = 0
    for file in files:
        if ".pptx" in file:
            ppt_counter += 1
            if os.path.isfile(cwd+file[:-5]+'.mp4')==False or os.path.getsize(cwd+file[:-5]+'.mp4')==0: # Check if conversion already happened
                pptx_to_mp4(cwd+file,cwd+file[:-5]+'.mp4')
                # import ipdb; ipdb.set_trace()
                print("[+] Converting", file) 
            else:
                print("[-] Skipping", file)
            
    
    if ppt_counter == 0:
        input("\nPlease put your powerpoint files in the same folder as this program - Press ENTER to exit...")
    else:
        input("\nDone, study well! :) - Press ENTER to exit")