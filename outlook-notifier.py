# coding=utf-8

#-------------------------------------------------------------------------------------------------------------------------
#This is a python2 script.
#Install Notes:
#  pip2.7 install pypiwin32
#  pip2.7 install git+https://github.com/frakman1/lazylights@2.0 
#  pip2.7 install playsound
#  pip2.7 install colorama

#Run Notes:
#  C:\Python27\python.exe outlook-notifier.py
#  Use git-bash to see colors: $ c:/python27/python -i outlook-notifier.py
#For emojis: 
#  Use Locale:C Character-Set:UTF-8 in git-bash->Options->Text
#  Install https://github.com/eosrei/twemoji-color-font/releases/download/v12.0.1/TwitterColorEmoji-SVGinOT-12.0.1.zip
#-------------------------------------------------------------------------------------------------------------------------
import sys
import win32com.client
import sched, time
from datetime import datetime, timedelta
from colorama import init, Fore, Style
import lazylights
import binascii
import playsound

# global font variables
bright_green = Fore.GREEN + Style.BRIGHT
bright_yellow = Fore.YELLOW + Style.BRIGHT
bright_magenta = Fore.MAGENTA + Style.BRIGHT
bright_cyan = Fore.CYAN + Style.BRIGHT
bright_red = Fore.RED + Style.BRIGHT
bright_white = Fore.WHITE + Style.BRIGHT

INTERVAL = 10          # Polling Interval
ALERT_TIME = 5         # Minutes before an event to alert you
OFFICE_HOUR_START = 8  # 8AM
OFFICE_HOUR_END = 23   # 6PM

#------------------------------------------------------------------------------------------------------------
# I use this to manually create a bulb using IP and MAC address. 
def createBulb(ip, macString, port = 56700):        
    return lazylights.Bulb(b'LIFXV2', binascii.unhexlify(macString.replace(':', '')), (ip,port))
#------------------------------------------------------------------------------------------------------------	

myBulb1 = createBulb('192.168.86.48','d0:73:d5:02:a9:1e')
myBulb2 = createBulb('192.168.86.49','d0:73:d5:00:41:6d')
myBulb3 = createBulb('192.168.86.62','d0:73:d5:20:ae:2f')
myBulb4 = createBulb('192.168.86.79','d0:73:d5:2a:ce:e6')
myBulb5 = createBulb('192.168.86.63','d0:73:d5:20:a0:00')
myBulb6 = createBulb('192.168.86.47','d0:73:d5:02:6b:04')

bulbs=[myBulb1, myBulb2, myBulb3, myBulb4, myBulb5, myBulb6]


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
        
        print("adate type: "+str(type(adate)))
        print (adate)
        events.append({'Start':adate,'Subject':a.Subject,'Duration':a.Duration})
        #events['Subject'].append(a.Subject)
        #events['Duration'].append(a.Duration)
    return events


def main_loop(sc):
    red_alert = False
    print(bright_magenta + "-----------------------------------------------------------------------------------------" + bright_white)
    print(bright_cyan + "📅️ {}".format(get_timestamp()) + bright_white)

    if  datetime.today().replace(hour=OFFICE_HOUR_START,minute=0) <= datetime.now() <= datetime.today().replace(hour=OFFICE_HOUR_END,minute=0):
        print("Within office hours")
        
        events = getCalendarEntries()    
        if not events:
            print(bright_white + "     No events to report")
            #lazylights.set_state(bulbs, 0, 0, 1, 8000, 500, raw=False)  # white
        else:
            for event in events:
                event_end  = event['Start'] + timedelta(minutes=event['Duration'])
                alert_time = event['Start'] - timedelta(minutes=ALERT_TIME)
                if ( event_end) < datetime.now():
                    print(bright_green + "    ✅ SAFE ✅      Event \"{}\" happened in the past (Started:{}. Ended:{})".format(event['Subject'], event['Start'].strftime("%I:%M %p"), event_end.strftime("%I:%M %p")) + bright_white)
                    #lazylights.set_state(bulbs, 0, 0, 1, 8000, 500, raw=False)  # white

                if alert_time <= datetime.now() <= event['Start']:
                    print(bright_yellow + "    ⏰ WARNING ⏰   Event \"{}\" starts within 5 minutes at {}".format(event['Subject'],event['Start'].strftime("%I:%M %p")) + bright_white)
                    playsound.playsound('UpcomingMeeting.mp3', True)
                    #lazylights.set_state(bulbs, 60, 1, 1, 8000, 500, raw=False)  # yellow

                if ( event['Start'] <= datetime.now() <= event_end):
                    print(bright_red + "    🚨 ALERT 🚨     Event \"{}\" is happening right now (Started {}. Ends {})".format(event['Subject'], event['Start'].strftime("%I:%M %p"), event_end.strftime("%I:%M %p")) + bright_white)
                    red_alert = True
                    #lazylights.set_state(bulbs, 0, 1, 1, 8000, 500, raw=False)  # red
                    
                else:
                    #lazylights.set_state(bulbs, 0, 0, 1, 8000, 500, raw=False)  # white
                    pass

    if red_alert:
        lazylights.set_state(bulbs, 0, 1, 1, 8000, 500, raw=False)  # red
    else:
        lazylights.set_state(bulbs, 0, 0, 1, 8000, 500, raw=False)  # white
        
    s.enter(INTERVAL, 1, main_loop, (sc,))
    
def main():

    s.enter(1, 1, main_loop, (s,))
    s.run()
    
if __name__ == "__main__":
    # execute only if run as a script
    main()
    
