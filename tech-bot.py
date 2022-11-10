# Required Packages:
# slackclient
# gspread
# slackeventsapi
# python-dotenv
# oauth2client
# numpy
#
# by: Frederico Wieser

import slack # Slack API Manager
import sys # For exiting the program for reboots
import os # For accesing files and directories in the system
import gspread # Google Sheets Package
import time # For delay function
import numpy as np
from pathlib import Path
from dotenv import load_dotenv
from slackeventsapi import SlackEventAdapter # Message sending package for Slack
from csv import writer # Import writer class from csv module
from oauth2client.service_account import ServiceAccountCredentials
import socket
from datetime import datetime
from datetime import timedelta

###############################################################################
# DEFINING SYSTEM PARAMETERS WHICH MAY CHANGE BY STUDIO

#naama_slack_channel = '#laser-daily-update'
naama_slack_channel = '#bot-test'

# Tech-Bot project folder directory decleration on local system.
env_path = Path('.') / '.env'
load_dotenv(dotenv_path=env_path)

# Checking Slack Token of Tech-Bot App for security in ".env".
client = slack.WebClient(token=os.environ['SLACK_TOKEN'])

# G-Spread Keys
laser-availability-gspread = '1cUgnU57bXlvZl5fFaH3e0w753mAnryLhDda27xQiIuI'
laser-availability-log-gspread = '15019z9lHnEOC97fKy8Cx-EeGG-5qe8XCy3lgq48u2XA'
analysis-downtime-gspread = '1WARgdqpOCO-UUk-X1iIlOWrTQdNqf4P3tOUrOF63zxk'
sytem-parameters-gspread = '1b-trYip_8-vGAHoU395Z1lp-qzCv7rr7mljLy3prTlc'
laser-parameter-log-gspread = '17WVIWLYWEKHy-oMNDOyYspfdiNJGgeOtQUsJecuYVcE'

# Google Drive API Keys


###############################################################################
# Connecting to G-Sheets and defining how long the program will run for.

# Set socket time out to 10 minutes, to avoid server timeouts.
socket.setdefaulttimeout(600)

# Google Sheets API Authentication.
from google.oauth2.service_account import Credentials
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
credentials = Credentials.from_service_account_file(google_json, scopes=scopes)
gc = gspread.authorize(credentials)

# Checking the time when the program is launched using system time.
starttime = time.time()

# Amount of time the whole "Tech-Bot" program runs for.
runtime = 60.0*60*2 # seconds

# Current time
current = time.time()

# To stop Excess Call Errors
time.sleep(5)

###############################################################################
# Laser-Parameter G-Sheets Data Extraction for Checking

# Opening Laser parameter spreadsheet
all_sheets = gc.open_by_key('1b-trYip_8-vGAHoU395Z1lp-qzCv7rr7mljLy3prTlc')

# Treatment room sheets are extracted and defined
Room_1 = all_sheets.get_worksheet(0)
time.sleep(1)
Room_2 = all_sheets.get_worksheet(1)
time.sleep(1)
Room_3 = all_sheets.get_worksheet(2)
time.sleep(1)
Room_4 = all_sheets.get_worksheet(3)
time.sleep(1)
Room_5 = all_sheets.get_worksheet(4)
time.sleep(1)

# Extract room parameter values
Room_1_vals = Room_1.get_all_values()
time.sleep(1)
Room_2_vals = Room_2.get_all_values()
time.sleep(1)
Room_3_vals = Room_3.get_all_values()
time.sleep(1)
Room_4_vals = Room_4.get_all_values()
time.sleep(1)
Room_5_vals = Room_5.get_all_values()
time.sleep(1)

# To stop Excess Call Errors
time.sleep(5)

###############################################################################
# Opening and storing the cell values for the "Laser Availability" Sheet in
# list_of_lists variable.

sh = gc.open_by_key(laser-availability-gsheet)
ws = sh.get_worksheet(0)
list_of_lists = ws.get_all_values() # List of Cell values

###############################################################################
# Morning 9:00 am and night 8:00 pm Slack updates and G-Drive logging.

# Counter definition so we don't repeat logging and sending messages.
counter = 0

# Upper Limit for stoping the bot from making repeat meassages
end_time_morning = datetime.time(8, 50, 0, 0)

# Lower Limit for starting the bot for making meassages
start_time_morning = datetime.time(8, 45, 0, 0)

# Upper Limit for stoping the bot from making repeat meassages
end_time_evening = datetime.time(20, 10, 0, 0)

# Lower Limit for starting the bot for making meassages
start_time_evening = datetime.time(20, 15, 0, 0)

###############################################################################
# Defining Methods which will be later used in the the bot.

def split(word):
    return [char for char in word]

def Save_To_Log():
    """
    Method for saving any changes to the "System-Parameter" Spreadsheet on your corresponding google drive.
    """
    # To stop Excess Call Errors
    time.sleep(5)

    # Open our existing CSV file in append mode
    # Create a file object for this file
    with open('log-parameters.csv', 'a') as f_object:

        # Pass this file object to csv.writer()
        # and get a writer object
        writer_object = writer(f_object)

        # Pass the list as an argument into
        # the writerow()
        writer_object.writerow(List)

        # Close the file object
        f_object.close()

        # Checking Credentials and defining scopes
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

        credentials = ServiceAccountCredentials.from_json_keyfile_name(google_json, scope)
        client = gspread.authorize(credentials)

        spreadsheet = client.open('Laser-Parameter-Log')

        # Opening our log.csv and saving it to the shared drive file
        with open('log-parameters.csv', 'r') as file_obj:
            content = file_obj.read()
            client.import_csv(spreadsheet.id, data=content.encode('utf-8'))

def Slack_LA_Text(room_name, room_id):
    """
    Method for sending updates to NAAMA's Slack channel where the stuido team
    can stay up to date with the studio's laser availability (ideally real-time).

    room_name : Name you will be allocating to room [string]
    """
    update = "Room " + str(room_name) + " : " + ("532 " if list_of_lists_new[int(room_id)][1] == "Y" else "") + ("800 " if list_of_lists_new[int(room_id)][2] == "Y" else "") + ("1064 " if list_of_lists_new[int(room_id)][3] == "Y" else "") + (list_of_lists_new[int(room_id)][4])
    return update

###############################################################################
# MAIN-LOOP # MAIN-LOOP # MAIN-LOOP # MAIN-LOOP # MAIN-LOOP # MAIN-LOOP # MAIN#
###############################################################################

while current < starttime + runtime :
    print(str(time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime(time.time()))))

    ##################################
    # LASER-Availability Bot Section #
    ##################################

    # To stop Excess Call Errors
    time.sleep(5)

    # Opening spreadsheet
    sh_new = gc.open_by_key(laser-availability-gsheet)
    ws_new = sh_new.get_worksheet(0)
    list_of_lists_new = ws_new.get_all_values()

    ###########################################################################
    # Morning 9:00 am
    ###########################################################################

    # Check if time is more than 9 am but less than 9:05 and that the counter
    # is less than 1 so that we don't repeatedly send messages and record logs.
    if current > start_time_morning and current < end_time_morning and counter < 1 :
        counter = counter + 1

        # Sending message to Slack Channel with Laser-Availability at current time
        print("MORNING AUTOMATIC CHECK @ " + str(time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime(time.time()))))
        client.chat_postMessage(channel=naama_slack_channel, text=
        Slack_LA_Text(1, 1) + "\n" +
        Slack_LA_Text(2, 2) + "\n" +
        Slack_LA_Text(3, 3) + "\n" +
        Slack_LA_Text(4, 4) + "\n" +
        Slack_LA_Text(5, 5)
        )
    elif current > start_time_evening and current < end_time_evening and counter < 1 :
        counter = counter + 1

        # Sending message to Slack Channel with Laser-Availability at current time
        print("EVENING AUTOMATIC CHECK @ " + str(time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime(time.time()))))
        client.chat_postMessage(channel=naama_slack_channel, text=
        Slack_LA_Text(1, 1) + "\n" +
        Slack_LA_Text(2, 2) + "\n" +
        Slack_LA_Text(3, 3) + "\n" +
        Slack_LA_Text(4, 4) + "\n" +
        Slack_LA_Text(5, 5)
        )



    # Checking if spreadsheets match
    if list_of_lists == list_of_lists_new:
        # No changes to laser availability so no updates needed
        pass
    else:
        ###############
        # New Changes #
        ###############

        # 1. Send message to Slack channel with current status of lasers
        print("NEW CHANGES TO Laser AVAILABILITY @ " + str(time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime(time.time()))))
        client.chat_postMessage(channel=naama_slack_channel, text=
        Slack_LA_Text(1, 1) + "\n" +
        Slack_LA_Text(2, 2) + "\n" +
        Slack_LA_Text(3, 3) + "\n" +
        Slack_LA_Text(4, 4) + "\n" +
        Slack_LA_Text(5, 5)
        )

        #######################################################################

        # List
        List = [str(time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime(time.time()))),
                  str(list_of_lists_new[1][1]),# Room 1
                  str(list_of_lists_new[1][2]),
                  str(list_of_lists_new[1][3]),
                  str(list_of_lists_new[1][4]),
                  str(list_of_lists_new[2][1]),# Room 2
                  str(list_of_lists_new[2][2]),
                  str(list_of_lists_new[2][3]),
                  str(list_of_lists_new[2][4]),
                  str(list_of_lists_new[3][1]),# Room 3
                  str(list_of_lists_new[3][2]),
                  str(list_of_lists_new[3][3]),
                  str(list_of_lists_new[3][4]),
                  str(list_of_lists_new[4][1]),# Room 4
                  str(list_of_lists_new[4][2]),
                  str(list_of_lists_new[4][3]),
                  str(list_of_lists_new[4][4]),
                  str(list_of_lists_new[5][1]),# Room 5
                  str(list_of_lists_new[5][2]),
                  str(list_of_lists_new[5][3]),
                  str(list_of_lists_new[5][4])]

        # Open our existing CSV file in append mode
        # Create a file object for this file
        with open('log-availability.csv', 'a') as f_object:

            # Pass this file object to csv.writer()
            # and get a writer object
            writer_object = writer(f_object)

            # Pass the list as an argument into
            # the writerow()
            writer_object.writerow(List)

            #Close the file object
            f_object.close()

            # CSV to Google Sheets
            import gspread
            from oauth2client.service_account import ServiceAccountCredentials

            # Checking Credentials and defining scopes
            scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                     "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

            credentials = ServiceAccountCredentials.from_json_keyfile_name(google_json, scope)
            client = gspread.authorize(credentials)

            spreadsheet = client.open('Laser-Availability-Log')

            # Opening our log.csv and saving it to the shared drive file
            with open('log-availability.csv', 'r') as file_obj:
                content = file_obj.read()
                client.import_csv(spreadsheet.id, data=content)

    list_of_lists = list_of_lists_new

    # checking current time so that app can be reset if needed
    current = time.time()

    # Timer
    delay1 = 20.0 #seconds
    time.sleep(delay1 - ((time.time() - starttime) % delay1))

    # checking current time so that app can be reset if needed
    current = time.time()

    ###############################
    # LASER-Parameter Bot Section #
    ###############################

    all_sheets_new = gc.open_by_key('1b-trYip_8-vGAHoU395Z1lp-qzCv7rr7mljLy3prTlc')

    Room_1_new = all_sheets_new.get_worksheet(0)
    time.sleep(0.25)
    Room_2_new = all_sheets_new.get_worksheet(1)
    time.sleep(0.25)
    Room_3_new = all_sheets_new.get_worksheet(2)
    time.sleep(0.25)
    Room_4_new = all_sheets_new.get_worksheet(3)
    time.sleep(0.25)
    Room_5_new = all_sheets_new.get_worksheet(4)
    time.sleep(0.25)

    Room_1_new_vals = Room_1_new.get_all_values()
    time.sleep(0.25)
    Room_2_new_vals = Room_2_new.get_all_values()
    time.sleep(0.25)
    Room_3_new_vals = Room_3_new.get_all_values()
    time.sleep(0.25)
    Room_4_new_vals = Room_4_new.get_all_values()
    time.sleep(0.25)
    Room_5_new_vals = Room_5_new.get_all_values()
    time.sleep(0.25)

    for i in range(len(Room_1_vals[0])):
        for j in range(len(Room_1_vals[:][0])):
            if Room_1_new_vals[i][j] == Room_1_vals[i][j] :
                pass
            else:
                List = [str(time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime(time.time()))),
                        "1",
                        str(Room_1_new_vals[1][j]),
                        str(Room_1_new_vals[i][0]),
                        str(Room_1_new_vals[i][1]),
                        str(Room_1_vals[i][j]),
                        str(Room_1_new_vals[i][j])]

                Save_To_Log()

    for i in range(len(Room_2_vals[0])):
        for j in range(len(Room_2_vals[:][0])):
            if Room_2_new_vals[i][j] == Room_2_vals[i][j] :
                pass
            else:
                List = [str(time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime(time.time()))),
                        "2",
                        str(Room_2_new_vals[1][j]),
                        str(Room_2_new_vals[i][0]),
                        str(Room_2_new_vals[i][1]),
                        str(Room_2_vals[i][j]),
                        str(Room_2_new_vals[i][j])]

                Save_To_Log()

    for i in range(len(Room_3_vals[0])):
        for j in range(len(Room_3_vals[:][0])):
            if Room_3_new_vals[i][j] == Room_3_vals[i][j] :
                pass
            else:
                List = [str(time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime(time.time()))),
                        "3",
                        str(Room_3_new_vals[1][j]),
                        str(Room_3_new_vals[i][0]),
                        str(Room_3_new_vals[i][1]),
                        str(Room_3_vals[i][j]),
                        str(Room_3_new_vals[i][j])]

                Save_To_Log()

    for i in range(len(Room_4_vals[0])):
        for j in range(len(Room_4_vals[:][0])):
            if Room_4_new_vals[i][j] == Room_4_vals[i][j] :
                pass
            else:
                List = [str(time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime(time.time()))),
                        "4",
                        str(Room_4_new_vals[1][j]),
                        str(Room_4_new_vals[i][0]),
                        str(Room_4_new_vals[i][1]),
                        str(Room_4_vals[i][j]),
                        str(Room_4_new_vals[i][j])]

                Save_To_Log()

    for i in range(len(Room_5_vals[0])):
        for j in range(len(Room_5_vals[:][0])):
            if Room_5_new_vals[i][j] == Room_5_vals[i][j] :
                pass
            else:
                List = [str(time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime(time.time()))),
                        "5",
                        str(Room_5_new_vals[1][j]),
                        str(Room_5_new_vals[i][0]),
                        str(Room_5_new_vals[i][1]),
                        str(Room_5_vals[i][j]),
                        str(Room_5_new_vals[i][j])]

                Save_To_Log()

    Room_1_vals = Room_1_new_vals
    Room_2_vals = Room_2_new_vals
    Room_3_vals = Room_3_new_vals
    Room_4_vals = Room_4_new_vals
    Room_5_vals = Room_5_new_vals

    # checking current time so that app can be reset if needed
    current = time.time()

    # Timer
    delay2 = 60.0 #seconds
    time.sleep(delay2 - ((time.time() - starttime) % delay2))

    # checking current time so that app can be reset if needed
    current = time.time()

###############################################################################
# MAIN-LOOP # MAIN-LOOP # MAIN-LOOP # MAIN-LOOP # MAIN-LOOP # MAIN-LOOP # MAIN#
###############################################################################

###########################################
# Analysis of Laser Availability Log Data #
###########################################

log = np.genfromtxt('log-availability.csv', delimiter=",", skip_header=1, dtype= "U20")

def DownTimeCalc(j):
    '''
    Calculates the downtime for a column in the laser-availability log or i.e. for one system.

    j : the index of the column
    '''

    downtime = timedelta(seconds=0)

    for i in range(len(log[:, 0])) :
        if log[-1, j] != "Y" :
            downtime = "Still Down"
        elif log[i, j] == "Y" :
            pass
        elif log[i, j] == "N/A" :
            pass
        else :
            last_y = datetime.strptime(log[i-1, 0],"%d-%m-%Y %H:%M:%S")
            first_n = datetime.strptime(log[i, 0],"%d-%m-%Y %H:%M:%S")

            extra_time = first_n - last_y

            downtime = downtime + extra_time

            if log[i+1, j] == "Y" :
                last_n = datetime.strptime(log[i, 0],"%d-%m-%Y %H:%M:%S")
                first_y = datetime.strptime(log[i+1, 0],"%d-%m-%Y %H:%M:%S")

                extra_time = first_y - last_n

                downtime = downtime + extra_time

                downtime = downtime.total_seconds()
            else:
                pass

    return str(downtime)

def split(word):
    return [char for char in word]

downtime = timedelta(seconds=0)

def DownTimeCalcAllRoomsOneDay(LOG):
    '''
    Calculates the downtime for a column in the laser-availability log or i.e. for one system.

    LOG : Log from the day you are trying to get the downtime
    '''

    availability_columns = [1,2,3,5,6,7,9,10,11,13,14,15,17,18,19]

    downtimes = np.empty(len(availability_columns), dtype='U10')

    # Counter for parsing downtimes array
    k = 0

    for j in availability_columns:

        downtime = 0
        print(downtime)

        if len(LOG[:, 0]) < 2:
            pass
        elif len(LOG[:, 0]) == 2
            # Check if System is still down at the end of the day
            if LOG[-1, j] == "N" :
                final_log_of_day = datetime.strptime(LOG[-1, 0],"%d-%m-%Y %H:%M:%S").timestamp()
                print(final_log_of_day)
                day = split(LOG[-1,0])
                midnight_of_day = datetime(int(day[6]+day[7]+day[8]+day[9]), int(day[3]+day[4]), int(day[0]+day[1]), 23, 59, 59).timestamp()
                print(midnight_of_day)
                downtime = midnight_of_day - final_log_of_day
                print(downtime)
            else:
                pass

            # Check if system was down this morning
            if LOG[0, j] == "N":
                first_log_of_day = datetime.strptime(LOG[1, 0],"%d-%m-%Y %H:%M:%S").timestamp()
                print(first_log_of_day)
                day = split(LOG[0,0])
                midnight_of_yday = datetime(int(day[6]+day[7]+day[8]+day[9]), int(day[3]+day[4]), int(day[0]+day[1]), 23, 59, 59).timestamp()
                print(midnight_of_yday)
                downtime = downtime + midnight_of_yday - first_log_of_day
                print(downtime)
            else:
                pass

        elif len(LOG[:, 0]) > 2:
            for i in range(2, len(LOG[:, 0])-1) :

                # Checking if the system is ok
                if LOG[i, j] == "Y" :
                    pass
                # Checking if the system is installed
                elif LOG[i, j] == np.nan :
                    pass
                elif LOG[i, j] == "N" :
                    last_value = datetime.strptime(LOG[i-1, 0],"%d-%m-%Y %H:%M:%S").timestamp()
                    print(last_value)
                    first_n = datetime.strptime(LOG[i, 0],"%d-%m-%Y %H:%M:%S").timestamp()
                    print(first_n)
                    extra_time = first_n - last_value
                    print(extra_time)
                    downtime = downtime + extra_time
                    print(downtime)
                    if log[i+1, j] == "Y" :
                        last_n = datetime.strptime(LOG[i, 0],"%d-%m-%Y %H:%M:%S").timestamp()
                        print(last_n)
                        first_y = datetime.strptime(LOG[i+1, 0],"%d-%m-%Y %H:%M:%S").timestamp()
                        print(first_y)
                        extra_time = first_y - last_n
                        print(extra_time)
                        downtime = downtime + extra_time
                        print(downtime)
                    else:
                        downtime = downtime
                else:
                    pass

        downtimes[k] = str(downtime)

        k = k + 1

    return downtimes

log = df2.to_numpy()

def DownTimeCalcPerDay(LOG2):
    '''
    Takes in the log at the end of the main loop of Tech-Bot and calculates the
    downtimes for each laser in the system and outputs a 2-D numpy array with
    dowtimes sorted by date and system.
    '''
    from datetime import datetime

    # Create empty array for getting log timestamps but only dates
    dates_all = np.empty(len(LOG2[:, 0]), dtype='U10')

    # For loop for finding all the dates in the log
    for i in range(len(LOG2[:, 0])):
        date = datetime.strptime(LOG2[i, 0],"%d-%m-%Y %H:%M:%S").date()
        dates_all[i] = date

    # Find all unique dates inside of log
    dates = np.unique(dates_all, axis=0)

    # Create Mask 2D array with masks indexed from first date in log
    # to last day in log
    mask = np.empty((len(dates), len(LOG2[:, 0])), dtype="bool")

    # Creating Array for storing the downtimes by date for each system
    dates_downtimes = np.empty((len(dates), 16), dtype='U20')

    # loop to check each unique date in the log and calculate downtime
    # returning a 1-D array which is appended to the "dates_downtime" array.
    for i in range(len(dates)):
        if i == 0:
            dates_downtimes[i, 0] = dates[i]

            mask[i] = (dates_all == dates[i])

            current_mask = LOG2[mask[i], :]

            dates_downtimes[i, 1:16] = DownTimeCalcAllRoomsOneDay(current_mask)
        else:

            dates_downtimes[i, 0] = dates[i]

            mask[i] = (dates_all == dates[i])

            if LOG2[mask[i-1]][-1].ndim == 1:

                current_mask = np.concatenate(([LOG2[mask[i-1]][-1]], LOG2[mask[i]]), axis=0)

                dates_downtimes[i, 1:16] = DownTimeCalcAllRoomsOneDay(current_mask)

            else:

                current_mask = np.concatenate((LOG2[mask[i-1]][-1], LOG2[mask[i]]), axis=0)

                dates_downtimes[i, 1:16] = DownTimeCalcAllRoomsOneDay(current_mask)

    return dates_downtimes

# Opening spreadsheet
downtime_sh = gc.open_by_key('1WARgdqpOCO-UUk-X1iIlOWrTQdNqf4P3tOUrOF63zxk')
downtime_ws = downtime_sh.get_worksheet(0)

# Updating Downtime values in spreadsheet
# Room 1
downtime_ws.update_cell(2, 2, DownTimeCalc(1))
time.sleep(0.25)
downtime_ws.update_cell(2, 3, DownTimeCalc(2))
time.sleep(0.25)
downtime_ws.update_cell(2, 4, DownTimeCalc(3))
time.sleep(0.25)

# Room 2
downtime_ws.update_cell(3, 2, DownTimeCalc(5))
time.sleep(0.25)
downtime_ws.update_cell(3, 3, DownTimeCalc(6))
time.sleep(0.25)
downtime_ws.update_cell(3, 4, DownTimeCalc(7))
time.sleep(0.25)

# Room 3
downtime_ws.update_cell(4, 2, DownTimeCalc(9))
time.sleep(0.25)
downtime_ws.update_cell(4, 3, DownTimeCalc(10))
time.sleep(0.25)
downtime_ws.update_cell(4, 4, DownTimeCalc(11))
time.sleep(0.25)

# Room 4
downtime_ws.update_cell(5, 2, DownTimeCalc(13))
time.sleep(0.25)
downtime_ws.update_cell(5, 3, DownTimeCalc(14))
time.sleep(0.25)
downtime_ws.update_cell(5, 4, DownTimeCalc(15))
time.sleep(0.25)

# Room 5
downtime_ws.update_cell(6, 2, DownTimeCalc(17))
time.sleep(0.25)
downtime_ws.update_cell(6, 3, DownTimeCalc(18))
time.sleep(0.25)
downtime_ws.update_cell(6, 4, DownTimeCalc(19))
time.sleep(0.25)

# Stop The Program
sys.exit()
