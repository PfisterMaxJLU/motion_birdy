import sys
import pkg_resources
import subprocess
import glob
import time
import os
import time
from datetime import datetime

def sum_writer_1hour(worksheet, row):
    #"functional"
    worksheet.write_formula(row, 1, f'=SUMIF(Observations!$A$6:Observations!$A$10006,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1}),),Observations!D$6:Observations!D$10006)-SUMIF(Observations!$A$6:Observations!$A$10006,">"&TIME(HOUR($A{row+1})+1,MINUTE($A{row+1}),),Observations!D$6:Observations!D$10006)')
    worksheet.write_formula(row, 2, f'=SUMIF(Observations!$A$6:Observations!$A$10006,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1}),),Observations!E$6:Observations!E$10006)-SUMIF(Observations!$A$6:Observations!$A$10006,">"&TIME(HOUR($A{row+1})+1,MINUTE($A{row+1}),),Observations!E$6:Observations!E$10006)')
    worksheet.write_formula(row, 3, f'=SUMIF(Observations!$A$6:Observations!$A$10006,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1}),),Observations!F$6:Observations!F$10006)-SUMIF(Observations!$A$6:Observations!$A$10006,">"&TIME(HOUR($A{row+1})+1,MINUTE($A{row+1}),),Observations!F$6:Observations!F$10006)')
    worksheet.write_formula(row, 4, f'=SUMIF(Observations!$A$6:Observations!$A$10006,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1}),),Observations!G$6:Observations!G$10006)-SUMIF(Observations!$A$6:Observations!$A$10006,">"&TIME(HOUR($A{row+1})+1,MINUTE($A{row+1}),),Observations!G$6:Observations!G$10006)')
    worksheet.write_formula(row, 5, f'=SUMIF(Observations!$A$6:Observations!$A$10006,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1}),),Observations!H$6:Observations!H$10006)-SUMIF(Observations!$A$6:Observations!$A$10006,">"&TIME(HOUR($A{row+1})+1,MINUTE($A{row+1}),),Observations!H$6:Observations!H$10006)')
    return worksheet

def sum_writer_5min(worksheet, row):
    worksheet.write_formula(row, 1, f'=SUMIF(Observations!$A$6:Observations!$A$10000,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1}),),Observations!D$6:Observations!D$10000)-SUMIF(Observations!$A$6:Observations!$A$10000,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1})+5,),Observations!D$6:Observations!D$10000)')
    worksheet.write_formula(row, 2, f'=SUMIF(Observations!$A$6:Observations!$A$10000,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1}),),Observations!E$6:Observations!E$10000)-SUMIF(Observations!$A$6:Observations!$A$10000,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1})+5,),Observations!E$6:Observations!E$10000)')
    worksheet.write_formula(row, 3, f'=SUMIF(Observations!$A$6:Observations!$A$10000,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1}),),Observations!F$6:Observations!F$10000)-SUMIF(Observations!$A$6:Observations!$A$10000,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1})+5,),Observations!F$6:Observations!F$10000)')
    worksheet.write_formula(row, 4, f'=SUMIF(Observations!$A$6:Observations!$A$10000,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1}),),Observations!G$6:Observations!G$10000)-SUMIF(Observations!$A$6:Observations!$A$10000,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1})+5,),Observations!G$6:Observations!G$10000)')
    worksheet.write_formula(row, 5, f'=SUMIF(Observations!$A$6:Observations!$A$10000,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1}),),Observations!H$6:Observations!H$10000)-SUMIF(Observations!$A$6:Observations!$A$10000,">"&TIME(HOUR($A{row+1}),MINUTE($A{row+1})+5,),Observations!H$6:Observations!H$10000)')
    return worksheet

def header_fill_standard(worksheet):
    #writes header for summarized sheet
    worksheet.write(0, 0, "User:")
    worksheet.write(1, 0, "ID:")
    worksheet.write(2, 0, "Date:")
    worksheet.write(4, 0, "Interval")
    worksheet.write(4, 1, "Event 1")
    worksheet.write(4, 2, "Event 2")
    worksheet.write(4, 3, "Event 3")
    worksheet.write(4, 4, "Event 4")
    worksheet.write(4, 5, "Event 5")
    return worksheet

def header_fill_1hour(workbook, worksheet):
    #fills the 1 hour summarized sheet
    worksheet = header_fill_standard(worksheet)
    row = 5
    offset = start_hour*60*60
    date_format = workbook.add_format({'num_format': 'hh:mm:ss', 'align': 'right'})
    for one_time in range(0,23-start_hour):
        dt = datetime.strptime(time.strftime('%H:%M:%S', time.gmtime(offset)),'%H:%M:%S')
        worksheet.write_datetime(row, 0, dt, date_format)
        worksheet = sum_writer_1hour(worksheet, row)
        row += 1
        offset += 3600


def header_fill_5min(workbook, worksheet):
    #fills the 5 min summarized sheet
    worksheet = header_fill_standard(worksheet)
    row = 5
    offset = start_hour*60*60
    date_format = workbook.add_format({'num_format': 'hh:mm:ss', 'align': 'right'})
    for one_time in range(0,287-start_hour*12):
        dt = datetime.strptime(time.strftime('%H:%M:%S', time.gmtime(offset)),'%H:%M:%S')
        worksheet.write_datetime(row, 0, dt, date_format)
        worksheet = sum_writer_5min(worksheet, row)
        row += 1
        offset += 300
        
def header_fill(worksheet):
    #writes header for main page
    worksheet.write(0, 0, "User:")
    worksheet.write(1, 0, "ID:")
    worksheet.write(2, 0, "Date:")
    worksheet.write(4, 0, "Start Time")
    worksheet.write(4, 1, "End Time")
    worksheet.write(4, 2, "Filename")
    worksheet.write(4, 3, "Event 1")
    worksheet.write(4, 4, "Event 2")
    worksheet.write(4, 5, "Event 3")
    worksheet.write(4, 6, "Event 4")
    worksheet.write(4, 7, "Event 5")
    return worksheet

def xlsx_writer(time_list):
    #constructs most of the xlsx
    workbook = xlsxwriter.Workbook('Results.xlsx')
    worksheet_1 = workbook.add_worksheet(name="Observations")
    date_format = workbook.add_format({'num_format': 'hh:mm:ss', 'align': 'right'})
    worksheet_1.set_column(2, 2, 35)
    row = 0
    col = 0
    offset = start_hour*60*60 + start_min*60 - vid_length*60
    star_times = []
    end_times = []

    for idx, sublist in enumerate(time_list):
        offset += vid_length*60
        for num, singel_time in enumerate(sublist):
            if singel_time != "-\n":
                just_seconds = get_sec(singel_time)
                just_seconds += offset
                singel_time = time.strftime('%H:%M:%S', time.gmtime(just_seconds))

                if num%2 == 0:
                    star_times.append(singel_time)
                if num%2 != 0:
                    end_times.append(singel_time)
            else:
                num += 1
                print(f"No motion detected in {files_grabbed[idx]}\n")

    worksheet_1 = header_fill(worksheet_1)
    row += 5

    length_list = []
    for part_list in time_list:
        length_list.append(int(len(part_list)/2))

    for num, singel_time in enumerate(star_times):
        dt = datetime.strptime(time.strftime(star_times[num]),'%H:%M:%S')
        worksheet_1.write_datetime(row, col, dt, date_format)
        col += 1
        dt = datetime.strptime(time.strftime(end_times[num]),'%H:%M:%S')
        worksheet_1.write_datetime(row, col, dt, date_format)
        row += 1
        col -= 1

    row = 5
    col = 2

    for num, lenght in enumerate(length_list):
        for _ in range(0, lenght):
            worksheet_1.write(row, col, files_grabbed[num])
            row += 1

    header_fill_5min(workbook, workbook.add_worksheet(name="Sum 5-Min"))
    header_fill_1hour(workbook, workbook.add_worksheet(name="Sum 1-Hour"))
    workbook.close()

def get_sec(time_str):
    # hh:mm:ss --> seconds
    h, m, s = time_str.split(':')
    return int(h) * 3600 + int(m) * 60 + int(s)

def user_input_collector():
    #gets user input at start
    length_list = []
    for file in files_grabbed:
        length_list.append(get_vid_length(file))

    if all(x==length_list[0] for x in length_list):
        vid_length = int(length_list[0]/60)
        print(f"A video duration of {int(vid_length)} minutes was determined.\n")
    else:
        print("No definitive video duration could be determined. How long are the videos?")
        try:
            vid_length = int(input("Video duration (min): "))
        except:
            input("Input not valid restart program and try again.")
            exit()
        print("")

    print("At which hour do your videos start?")
    try:
        start_time = input("Day start (hours minutes e.g. \"13 37\"): ")
        start_time = start_time.split(" ")
        start_hour, start_min = int(start_time[0]), int(start_time[1])
    except:
        input("Input not valid restart program and try again.")
        exit()
    print("")

    print("Do you want to use a custom threshold for the motion detection step?\n(Examples: 4.50, 3.00, 0.75) Lower value > higher sensibility.\nFor the default (3) simply press enter.")
    threshold = input("Threshold: ")

    if threshold == "":
        threshold = 3.0
    try:
        threshold = float(threshold)
    except:
        input("Input not valid restart program and try again.")
        exit()
    print("")

    return vid_length, start_hour, start_min, threshold

def dependency_installer():
    #installs dependency and tells user to do it manually if it fails itself
    required = {'psutil', 'xlsxwriter','dvr-scan', 'opencv-python', 'progress-bar'}
    installed = {pkg.key for pkg in pkg_resources.working_set}
    missing = required - installed

    try:
        if missing:
            python = sys.executable
            subprocess.check_call([python, '-m', 'pip', 'install', *missing], stdout=subprocess.DEVNULL)

        if len(missing) > 0:
            print("\nSetting up some dependency, this will only be required on first startup. The process takes about 30 seconds.")
            time.sleep(30)
    except:
        print("\nAutomatic installation of dependencies failed. Please install the following manually via (python) \"pip\".\n")
        print("pip install psutil")
        print("pip install xlsxwriter")
        print("pip install dvr-scan")
        print("pip install opencv-python")
        print("pip install progress-bar")
        input("\nThis should resolve the issue on the next start. If not, contact Max.Pfister@bio.uni-giessen.de.")
        exit()

print("motion_birdy launched) In case of questions or problems write Max.Pfister@bio.uni-giessen.de")
dependency_installer()

#main processing part with the help of dvr-scan
import psutil 
import xlsxwriter
import cv2

def childCount():
    children = psutil.Process().children()
    return(len(children))

def progress_bar(cmd_len):
    chunk_size = 100/len(files_grabbed)
    jobs_done = len(files_grabbed) - cmd_len - childCount()
    bar_ammount = int(jobs_done*chunk_size)
    print('\r', end="") #hmm, does not work in py-console, still looks ok
    print(f"Progress:[{'>'*bar_ammount}{'-'*(100-bar_ammount)}] ({bar_ammount}%)", end="")

def get_vid_length(filename):
    video = cv2.VideoCapture(filename)
    fps = video.get(cv2.CAP_PROP_FPS)
    frame_count = video.get(cv2.CAP_PROP_FRAME_COUNT)
    duration = frame_count / fps
    return duration

print("\nHello, before I can start, I need some Information for the final Excel file (Enter to confirm):\n")

types = ('*.avi', '*.mp4')
files_grabbed = []
for files in types:
    files_grabbed.extend(glob.glob(files))

vid_length, start_hour, start_min, threshold = user_input_collector()

print("Performing motion (bird) detection on files:\n")

for file in files_grabbed:
    print(file)
print("")
print("Progress bar only updates if a full video is processed. Depending on system speed, this can take a while.\n")

if not os.path.exists("motion_only_videos"):
    os.makedirs("motion_only_videos")
else:
    print("An analysis directory seems to exist already (motion_only_videos),\nif it contains files from prior aborted runs,\nthis will cause problems further down the analysis pipeline.\nIt is highly recommended to delete the folder and their contents.\n")

if not os.path.exists("logs"):
    os.makedirs("logs")
else:
    print("An analysis directory seems to exist already (logs),\nif it contains files from prior aborted runs,\nthis will cause problems further down the analysis pipeline.\nIt is highly recommended to delete the folder and their contents.\n")

cmd = []
for f in files_grabbed:
    just_name =  os.path.splitext(f)[0]
    cmd.append(f"dvr-scan -i {f} -o motion_only_videos/MB_{f} -t {threshold} -q -tc > logs/tmp_logfile_{just_name}.log")

while len(cmd) > 0:
    for i in cmd:
        if childCount() < os.cpu_count():
            cmd.remove(i)
            procs = subprocess.Popen(i, shell=True)
    time.sleep(5)
    progress_bar(len(cmd))

while childCount() > 0:
    time.sleep(5)
    progress_bar(len(cmd))

log_files = glob.glob("logs/*.log")

full_log = open("logs/MB_full_log.log", "w")
for file in log_files:
    with open(file, "r") as singular_log:
        if sum(1 for line in singular_log)>0:
            singular_log.seek(0)
            for line in singular_log: 
                if line != "[DVR-Scan] Comma-separated timecode values:\n":
                    full_log.write(line)     
        else:
            full_log.write("-\n")

for file in log_files:
    if file.startswith("logs\\tmp_logfile"):
        os.remove(file) 
full_log.close()

time_list = []
with open("logs/MB_full_log.log", "r") as full_log:
    for linedata in full_log:
        time_list.append(linedata.split(","))

for i_1, sublist in enumerate(time_list):
        for i_2, singel_time in enumerate(sublist):
            time_list[i_1][i_2] = singel_time.split(".")[0]

xlsx_writer(time_list)

print("Job Done\n")
input("Please confirm the end of analysis by pressing enter") 