# Created by Philip Goosen github.com/goosenphil

import win32com.client
import time
import os
import zipfile
from ppsx_patcher import patch_ppsx
import os.path
import subprocess
from colorama import Fore, Back, Style, init as colorama_init
# Only for Windows computers running PowerPoint 2007 or newer

use_voice_times = 15
slide_duration = 1 # only matters if use_voice_times = 0
resolution = 720 # Vertical resolution, default at 720p
frame_rate = 24
quality = 60

def is_powerpoint_running():
    tasks = subprocess.check_output('tasklist', shell=True)
    if b"POWERPNT.EXE" in tasks.upper():
        return True
    else:
        return False

def pptx_to_mp4(pptx_input,mp4_output, ppt):
    print("[+] Converting", pptx_input, "...")
    # ppt = win32com.client.Dispatch('PowerPoint.Application')
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
            print(Fore.GREEN,bar[i % len(bar)],Style.RESET_ALL,"converting...", end="\r")
            i += 1
            pass
    end_time_stamp=time.time()
    print(Fore.GREEN + "Converted in:", round(end_time_stamp-start_time_stamp, 3), "seconds" + Style.RESET_ALL)
    # ppt.Quit()
    # pass
  
if __name__ == '__main__':
    colorama_init()
    print("\nPowerPointless V1.1 (PowerPoint to video converter) by Philip Goosen [19509766@sun.ac.za]")
    print("https://github.com/goosenphil/PowerPointless\n")

    print(Fore.YELLOW + "Please don't open PowerPoint while this program is running" + Style.RESET_ALL)

    if is_powerpoint_running():
        print(Fore.RED + "Please close PowerPoint so I can run :/" + Style.RESET_ALL)
        print("If I do nothing for a while after you closed it, check in task manager, PowerPoint might still be running in the background.")
        while is_powerpoint_running():
            time.sleep(0.1)

    ppt = win32com.client.Dispatch('PowerPoint.Application')
    cwd = os.getcwd() + '\\'
    
    files = os.listdir(os.curdir)
    for file in files:
        if ".ppsx" in file:
            print(Fore.CYAN + "Detected ppsx file! -", file, " creating pptx patched version..." + Style.RESET_ALL)
            patch_ppsx(file)
    
    files = os.listdir(os.curdir)
    ppt_counter = 0
    for file in files:
        if ".pptx" in file:
            ppt_counter += 1
            if os.path.isfile(cwd+file[:-5]+'.mp4')==False or os.path.getsize(cwd+file[:-5]+'.mp4')==0: # Check if conversion already happened
                pptx_to_mp4(cwd+file,cwd+file[:-5]+'.mp4', ppt)
            else:
                print("[-] Skipping", file)
    
    ppt.Quit()
    # print("Done!")
    if ppt_counter == 0:
        print(Fore.YELLOW+ "\nPlease put your PowerPoint files in the same folder as this program "+ Style.RESET_ALL + "- Press ENTER to exit...")
        input()
    else:
        print("Waiting for PowerPoint to quit...")
        print(Fore.GREEN + "\nDone, study well! :)"+ Style.RESET_ALL + " - Press ENTER to exit..." )
        input()