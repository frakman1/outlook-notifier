# coding=utf-8

#-------------------------------------------------------------------------------------------------------------------------
#This is a python2 script.
#Install Notes:
#pip2.7 install pypiwin32
#C:\Python27\python.exe outlook.py
#Use git-bash to see colors: $ c:/python27/python -i outlook.py
#For emojis: 
# Use Locale:C Character-Set:UTF-8 in git-bash->Options->Text
# Install https://github.com/eosrei/twemoji-color-font/releases/download/v12.0.1/TwitterColorEmoji-SVGinOT-12.0.1.zip
#-------------------------------------------------------------------------------------------------------------------------
import sys
import win32com.client
import sched, time
from datetime import datetime, timedelta
from colorama import init, Fore, Style

# global font variables
bright_green = Fore.GREEN + Style.BRIGHT
bright_yellow = Fore.YELLOW + Style.BRIGHT
bright_magenta = Fore.MAGENTA + Style.BRIGHT
bright_cyan = Fore.CYAN + Style.BRIGHT
bright_red = Fore.RED + Style.BRIGHT
bright_white = Fore.WHITE + Style.BRIGHT

INTERVAL = 10

s = sched.scheduler(time.time, time.sleep)


def get_timestamp():
    dateTimeObj = datetime.now()
    timestampStr = dateTimeObj.strftime("%d-%b-%Y (%I:%M %p)")
    #print('Current Timestamp : ', timestampStr)
    return timestampStr
     
def getCalendarEntries(days=1):
    """
    Returns calender entries for days default is 1
    """
    Outlook = win32com.client.Dispatch("Outlook.Application")
    ns = Outlook.GetNamespace("MAPI")
    appointments = ns.GetDefaultFolder(9).Items
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = "True"
    today = datetime.today()
    begin = today.date().strftime("%m/%d/%Y")
    tomorrow= timedelta(days=days)+today
    end = tomorrow.date().strftime("%m/%d/%Y")
    appointments = appointments.Restrict("[Start] >= '" +begin+ "' AND [END] <= '" +end+ "'")
    #events={'Start':[],'Subject':[],'Duration':[]}
    events=[]
    for a in appointments:
        adate=datetime.fromtimestamp(int(a.Start))
        #print(type(adate))
        #print (a.Start, a.Subject,a.Duration)
        events.append({'Start':adate,'Subject':a.Subject,'Duration':a.Duration})
        #events['Subject'].append(a.Subject)
        #events['Duration'].append(a.Duration)
    return events


def main_loop(sc):
    print(bright_magenta + "-----------------------------------------------------------------------------------------" + bright_white)
    print(bright_cyan + "📅️ {}".format(get_timestamp()) + bright_white)

    events = getCalendarEntries()    
    for event in events:
        event_end = event['Start'] + timedelta(minutes=event['Duration'])
        if ( event_end) < datetime.now():
            print(bright_green + "    ✅ SAFE ✅      Event \"{}\" happened in the past (Started:{}. Ended:{})".format(event['Subject'], event['Start'].strftime("%I:%M %p"), event_end.strftime("%I:%M %p")) + bright_white)
        if ( event['Start'] <= datetime.now() <= event_end):
            print(bright_red + "    🚨 ALERT 🚨     Event \"{}\" is happening right now (Started {}. Ends {})".format(event['Subject'], event['Start'].strftime("%I:%M %p"), event_end.strftime("%I:%M %p")) + bright_white)
        else:
            alert_time = event['Start']-timedelta(minutes=5)
            if alert_time <= datetime.now() <= event['Start']:
                print(bright_yellow + "    ⏰ WARNING ⏰   Event \"{}\" starts within 5 minutes at {}".format(event['Subject'],event['Start'].strftime("%I:%M %p")) + bright_white)
    
    s.enter(INTERVAL, 1, main_loop, (sc,))
    
def main():
    s.enter(1, 1, main_loop, (s,))
    s.run()
    
if __name__ == "__main__":
    # execute only if run as a script
    main()
    
