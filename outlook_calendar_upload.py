import win32com.client, schedule, time, datetime, os
from datetime import date
from dateutil.parser import *

OUT_DIRECTORY = "C://Users//jacopor//Downloads"
FILE_NAME = date.today().strftime("%m-%d-%Y") + '_10day_meeting_list.csv'
ABSOLUTE_PATH = os.path.join(OUT_DIRECTORY, FILE_NAME)

def get_twelve_hour_time(twenty_four_hr_time):
    date.today()
    hour_num = int(twenty_four_hr_time[:2])
    daytime = 'pm' if hour_num / 12 >= 1 else 'am'
    hour_num %= 12
    hour_num = 12 if hour_num == 0 else hour_num
    min = twenty_four_hr_time.split(':')[1]
    return str(hour_num) + ':' + min + daytime


def write_calendar_to_csv():
    from datetime import date   # no clue why this is needed again
    print('Running calendar update for ' + str(date.today()))
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # fetch calendar of logged-on user
    appointments = namespace.GetDefaultFolder(9).Items

    # sort events by occurrence, include recurring events
    appointments.sort("[Start]")
    appointments.IncludeRecurrences = "True"

    # end date is today in form dd/mm/YYYY
    end = date.today().strftime("%m/%d/%Y")
    begin = (date.today() - datetime.timedelta(days=10)).strftime("%m/%d/%Y")
    appointments = appointments.Restrict("[Start] >= '" + begin + "' AND [END] <= '" + end + "'")

    # populate dictionary of meetings
    apptDict = {}
    for i, appt in enumerate(appointments):
        subject = str(appt.Subject)
        organizer = str(appt.Organizer)
        meetingDate = str(appt.Start)
        date = parse(meetingDate).date().strftime("%m/%d/%Y")
        time = get_twelve_hour_time(str(parse(meetingDate).time()))
        duration = str(appt.Duration)
        apptDict[i] = {"Duration": duration, "Organizer": organizer, "Subject": subject, "Date": date, "Time": time}

    # write to CSV
    csv = open(ABSOLUTE_PATH, 'w')
    csv.write('Date,Time,Duration,Subject,Organizer\n')
    for i in range(len(apptDict)):
        a = apptDict[i]
        csv.write(a['Date'] + ',' + a['Time'] + ',' + a['Duration'] + '\n')

def main():
    write_calendar_to_csv()

main()