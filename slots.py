""" A script that finds meeting slots in outlook
"""
import sys
import getpass
import datetime
import zoneinfo
from dateutil import tz
import csv

import click
import arrow
from InquirerPy import prompt
from InquirerPy.separator import Separator
from InquirerPy import inquirer
from InquirerPy.prompts.checkbox import CheckboxPrompt

LOCAL_TIMEZONE = datetime.datetime.now().astimezone().tzinfo

def air_port_code_to_tz(air_port_code,filename = 'tzmap.txt'):
    with open(filename, 'r') as file:
        datareader = csv.reader(file, delimiter='\t')
        for code, timezone in datareader:
                if code == air_port_code: return timezone

class OutlookMac():
    def __init__(self):
        from appscript import app, k
        self.outlook = app('Microsoft Outlook')
        self.k = k

    def get_freebusy(self, attendee, start_time, end_time, interval=15):
        """ check freebusy status in all accounts
        """

        start_time = arrow.get(start_time)
        end_time = arrow.get(end_time)
        for account in self.outlook.exchange_account(): 
            try:
                res = self.outlook.query_freebusy ( 
                    account,
                    for_attendees=[attendee],
                    range_start_time=start_time.naive,
                    range_end_time=end_time.naive,
                    interval=interval,
                )
            except:
                continue
            visibilitiy = {}
            attendee_email = res.pop(0)
            current_time = arrow.get(res.pop(0), 'YYYY-MM-DD HH:mm:ss Z')
            while  current_time < end_time:
                if len(res) == 0: break
                name, location, status =  (res.pop(0), res.pop(0), res.pop(0))
                visibilitiy[current_time.to('UTC')] = name, location, status #Store the time visibility in UTC for comparison purposes
                current_time = current_time.shift(minutes=interval)

            variations = set(list(visibilitiy.values()))
            if variations == {('', '', 'no info')}:
                continue 
            break
        else:
            raise Exception("Not found %r in %s accounts. Check email, Try restart Outlook or use VPN" % (attendee, len(outlook.exchange_account()) ))

        return visibilitiy

    def create_event(self, subject, content, start_time, end_time, attendees):
        event = self.outlook.make(
            new=self.k.calendar_event,
            with_properties={
                self.k.subject: subject,
                self.k.content: content,
                self.k.free_busy_status: self.k.busy,
                self.k.start_time: datetime.datetime(start_time.year, start_time.month, start_time.day, start_time.hour, start_time.minute), 
                self.k.end_time: datetime.datetime(end_time.year, end_time.month, end_time.day, end_time.hour, end_time.minute), 
            },
        )
        for email in attendees or []:
            event.make(
                new=self.k.required_attendee,
                with_properties={self.k.email_address: {self.k.address: email}}
            )
        event.open()
        event.activate()

        return event

#https://docs.microsoft.com/en-us/office/vba/api/outlook.exchangeuser.getfreebusy
class OutlookWin():
    def __init__(self):
        import win32com.client # pip install pywin32
        self.outlook = win32com.client.Dispatch('Outlook.Application')
        self.namespace = self.outlook.GetNamespace("MAPI")

    def get_freebusy(self, attendee, start_time, end_time, interval=15):
        """ check freebusy status in all accounts
            This is broken because the function outputs times starting at midnight. The arrow default starts at 11pm.
            https://docs.microsoft.com/en-us/office/vba/api/outlook.exchangeuser.getfreebusy

            Might need to correct this for timezone too.
            Looks like the "Office" field is populated. I could look up the airport code to find timezone.
            https://raw.githubusercontent.com/hroptatyr/dateutils/tzmaps/iata.tzmap
        """
        # saved_args = locals()
        # print("get_freebusy saved_args is", saved_args)

        recipient = self.namespace.CreateRecipient(attendee)

        #this code is so ugly and won't work outside of AWS.
        air_port_code = recipient.AddressEntry.GetExchangeUser().OfficeLocation[:3]
        timezone = air_port_code_to_tz(air_port_code)

        start_time = arrow.get(start_time).replace(hour=0,minute=0, second=0).to('UTC')
        # start_time = arrow.get(datetime.datetime(2013, 5, 5), 'US/Pacific')
        end_time = arrow.get(end_time)
        res = recipient.FreeBusy(start_time.datetime, interval, True)

        # print(f'freeBusy start_time: {start_time}')
        # print(f'start_time as datetime: {start_time.datetime}')
        # print(f'####### res: {res}  ########')

        visibilitiy = {}
        current_time = start_time
        for item in res:
            status = {
                '4': 'oof', #olWorkingElsewhere
                '3': 'oof',
                '2': 'busy',
                '1': 'tentative',
                '0': 'free',
            }.get(item, 'no info')

            if current_time > end_time: break #Don't consider times past the end time

            #Mark Busy if outside working hours
            attendeeTime = current_time.to(timezone)
            if attendeeTime.format('HHmm') >= '1700' or attendeeTime.format('HHmm') < '0800': status = 'busy'
            if attendeeTime.format('ddd') in ("Sat", "Sun"): status = 'busy'
            
            visibilitiy[current_time] = attendee, '', status
            current_time = current_time.shift(minutes=interval)
        return visibilitiy    

    def create_event(self, subject, content, start_time, end_time, attendees):
        mail = self.outlook.CreateItem(1)   #Open new mail

        mail.Start = start_time.to(LOCAL_TIMEZONE).format('YYYY-MM-DD HH:mm:ss') 
        mail.Subject = subject
        mail.Duration = 15 #Duration in minutes. Do the math with the start and end time.
        mail.MeetingStatus = 1 
        mail.Body = content

        for alias in attendees: mail.Recipients.Add(alias)
        mail.Display()

        return None

@click.command()
@click.option('-a', '--attendees', default=None, help='A semicolon separated list of contacts')
@click.option('--start', default='today', help='Begining of the date range to find available slots. Iso date or Arrow Humanised time (ex: "in 30 days"). See https://arrow.readthedocs.io/en/latest/#dehumanize. Default is today. ')
@click.option('--end', default='in 30 days',
              help='The end of the date range to find available slots. Iso date or Arrow Humanised time (ex: "in 30 days"). See https://arrow.readthedocs.io/en/latest/#dehumanize. Default is "in 30 days".')
@click.option('--full/--only-slots', default=False, help='Show slots or a full agenda (default=only slots)')
@click.option('-r', '--rate', default=100, help='Acceptable share of available attendees in percent(%)')
@click.option('--tentative/--no-tentative', default=True, help="Treat tentative meetings as free")
@click.option('-l', '--length', default=60, help='Length of the meeting, in minutes. default=60')
@click.option('-tz', '--alternative_tz', default=None, help='A comma separated list of alternative Timezones.')
@click.option('-hr', '--hours', default='0700-1800', help='Find availability between the hours of... Default = "0800-1900". ')
@click.option('-f', '--fmt', default='HH:mm',
              help='Time format in list of available slots. "HH:mm" or "h:mma". Refer to https://arrow.readthedocs.io/en/latest/#supported-tokens. Default "HH:mm".')
@click.option('-t', '--title', default='', help='Title of the event')
def main(attendees, start, end, full, rate, length, tentative, alternative_tz, hours, fmt, title):
    # saved_args = locals()
    # print("saved_args is", saved_args)

    if start == 'today':
        start = arrow.now(LOCAL_TIMEZONE).replace(minute=0, second=0)
    else:
        try:
            start = arrow.get(LOCAL_TIMEZONE).dehumanize(start).replace(minute=0, second=0)
        except:
            pass 

    between_lower, between_upper = hours.split('-')

    new_end = None
    try: 
        new_end = arrow.get(end)
    except:
        try:
            new_end = arrow.get().dehumanize(end).replace(minute=0, second=0)
        except:
            pass 
    assert new_end, f"I do not understand --end='{end}'. Please use ISO date or ARROW Humanise syntax. ex: 'in 30 days'"
    end = new_end

    if not attendees:
        answer = inquirer.text(
            message="enter attendees",
            long_instruction='Semicolon `;` or new line `\\n` separated list. ',
            instruction=' (Use `esc` followd by `enter` to complete the question.)',
            multiline=True,
            default=''
        ).execute()
        attendees = answer.replace('\n', ";")

    new_attendees = []
    for addr in attendees.split(";"):
        addr = addr.strip()
        if not addr: continue
        if '<' in addr and '>' in addr:
            addr = addr.split('<')[1].split('>')[0]
            new_attendees.append(addr)
        else:
           new_attendees.append(addr)
    attendees = list(set(new_attendees))

    
    alternative_tz = alternative_tz.split(',') if alternative_tz else []

    # print(f'########## alternative_tz: {alternative_tz} ##########')
    
    ### This part breaks it for me a looks like a safety check. Delete for now.
    # for itz, tz in enumerate(alternative_tz): 
    #     print(f'########## tz: {tz}, Available Time Zones {zoneinfo.available_timezones()}')
    #     if tz not in zoneinfo.available_timezones():
    #         alternative_tz[itz] = inquirer.fuzzy(
    #             message="Cannot find the TZ. Please select timezone:",
    #             choices=sorted(zoneinfo.available_timezones()),
    #             default=tz,
    #         ).execute()

    print ('Looking up agendas for', '; '.join(attendees))

    if sys.platform == 'darwin':
        outlook = OutlookMac()
    elif sys.platform.startswith('win'):
        outlook = OutlookWin()
    else:
        raise NotImplementedError(f'{sys.platform} not supported')

    interval = 15 if length <=30 else 30 
    freebusy = {}
    for index, _ in enumerate(attendees):
        while True: # for retry
            attendee = attendees[index]
            try:
                freebusy[attendee] = outlook.get_freebusy(attendee, start, end, interval=interval)
                break
            except Exception as exc:
                print(exc)
                answer = inquirer.text(
                    message=f"Cannot find {attendee} in outlook. Please correct email or delete to skip",
                    long_instruction='',
                    default=attendee
                ).execute()
                if not answer: 
                    freebusy[attendee] = None
                    break
                attendees[index] = answer

    freebusy = dict([item for item in freebusy.items() if item[1] ]) #filter empty 
    
    # print(f'######## freebusy: {freebusy.replace()}')
    # print("{" + "\n".join("{!r}: {!r},".format(k, v) for k, v in freebusy.items()) + "}")

    # print('####### Free Busy Times ########')
    # for att,_ in freebusy.items():
    #     print(f'################ Attendee: {att} ################################') 
    #     for k, v in freebusy[att].items(): 
    #         # print('Hello World')
    #         print(f'{k.to(LOCAL_TIMEZONE)} {v}')


    current_time_start = None
    current_slot_status = 'not started'
    current_busy_attendees = []

    busy_statuses = ['oof', 'busy'] if tentative else ['oof', 'busy', 'tentative']

    # candidateTimes = list(freebusy.values())[0]
    
    # dFreePeople={}
    # for time in candidateTimes:
    #     dFreePeople[time]=sum([(freebusy[att][time][2] not in busy_statuses) for att in freebusy])

    # print("###### dFreePeople #######")
    # for k, v in dFreePeople.items(): 
    #     # print('Hello World')
    #     print(f'{k} {v}')
    
    # dFreePeople = {time:sum([freebusy[att][time][2] not in busy_statuses])for time in candidateTimes}

    # if freebusy[att][time][2] in busy_statuses: free = False

    if full:
        # show all slots
        choices = []
        last_date = ''
        for time in list(freebusy.values())[0]:
            time = time.to(LOCAL_TIMEZONE)
            line  = time.format('    HH:mm ')
            # if time.format('HHmm') > between_upper or time.format('HHmm') < between_lower: continue
            # if time.format('ddd') in ("Sat", "Sun"): continue
            if time.format("dddd DD MMMM") != last_date:
                last_date = time.format("dddd DD MMMM")
                choices.append(Separator(last_date))
            num_free = 0
            for att in freebusy:
                status = freebusy[att][time][2]
                char = {
                    'oof': '▓',
                    'busy':'█',
                    'tentative': '░',
                    'free': ' '
                }.get(status, '?')
                if status not in busy_statuses:
                    num_free +=1
                line +=  char * 5 + '│'
            line +=   f'{num_free:2d}/{len(freebusy)} ' + num_free * '█'
            if (num_free + 0.5) / len(freebusy) * 100 > rate:
                choices.append({"name":line, "value": (time, time.shift(minutes=length))})
            else:
                choices.append(Separator(line))

        res = prompt(questions=[
            {
                "type": "list",
                "message": "Select a Timeslot to create a meeting (busy='█', free=' '):\n            " + ' '.join([key[:5] for key in freebusy]),
                "choices": choices,
                "long_instruction": f"You can choose slots with {rate}% rate. Use '--rate 50' parameter to select slots with partial availability."
            },
        ])

    else:
        # show only slots where everybody is available
        slots = {}

        candidate_times = list(freebusy.values())[0]
        for i, time in enumerate(candidate_times):
            # time = time.to(LOCAL_TIMEZONE)
            num_intervals = int(length/interval)
            num_free = 0
            busy_attendees = []
            for att in freebusy:
                free = True
                #Loop through all the times what would have to be free for the time to work
                # print(f'num intervals: {num_intervals}')
                # print(f'candidate times: {list(candidate_times)}')

                for att_time in list(candidate_times)[i:(i+num_intervals)]:
                    if freebusy[att][att_time][2] in busy_statuses: free = False
                    # if time.format('HHmm') > between_upper or time.format('HHmm') < between_lower: free = False
                    # if time.format('ddd') in ("Sat", "Sun"): free = False
                if free:
                    num_free += 1
                else:
                    busy_attendees.append(att)
            busy_attendees = sorted(busy_attendees)

            # do not show domains if they are all the same
            domains = [email.split('@')[1] for email in busy_attendees]
            if len(set(domains)) == 1:
                busy_attendees = [email.split('@')[0] for email in busy_attendees]

            attendee_free_rate = (num_free + 0.5) / len(freebusy) * 100

            # print(f'Time: {time.to(LOCAL_TIMEZONE)}, Free rate: {attendee_free_rate}, Trigger Value: {attendee_free_rate >=  rate}')

            if attendee_free_rate >=  rate:
                # print('updateing slots')
                slot_name = f'{time.to(LOCAL_TIMEZONE).format("dddd DD MMMM " + fmt)} - {time.to(LOCAL_TIMEZONE).shift(minutes=length).format(fmt + " ZZZ")}'
                slots[slot_name] = (time, time.shift(minutes=length))
                # print(f'slots: {slots}')

        # print(f'slots: {slots}')

        res = prompt(questions=[
            {
                "type": "checkbox",
                "message": "Select a Timeslot to create a meeting (space to multi-select, enter to finish):",
                "choices": [{"name": "Cancel (CTRL+C)", "value": None}] + [Separator()]
                        + [{"name":slots_name, "value": slots_data} for slots_name, slots_data in slots.items()],
                "default": "default",
            },
        ])


    if res[0] is None: 
        return

    # print(f'res: {res}')

    pretty_selected_times = ''
    for start_time, end_time in res[0]:
        # pretty_selected_times += f'\u2022 {start_time.to(LOCAL_TIMEZONE).format(fmt)} - {end_time.to(LOCAL_TIMEZONE).format(fmt)} {str(LOCAL_TIMEZONE)} \n'
        pretty_selected_times += f'\u2022 {start_time.to(LOCAL_TIMEZONE).format("ddd, MMM-DD HH:mm")} - {end_time.to(LOCAL_TIMEZONE).format("HH:mm ZZZ")} \n'


    print(pretty_selected_times)

    for i, (start_time, end_time) in enumerate(res[0]):
        # print(f'######### Creating Meeting {i}  ##########')
        # start_time= start_time.format()[:-9]
        # print(f'start time: {str(start_time)}')
        outlook.create_event(
            subject='[Placeholder] '+title,
            content='This placeholder meeting is part of the following set:\n'+pretty_selected_times,
            attendees=list(freebusy.keys()),
            # start_time='2022-02-03 12:00',#start_time.format(),
            start_time=start_time,#[:-5],
            end_time=end_time,
        )

    # if inquirer.select(
    #         message="event created. are you happy with this event?",
    #         choices=['yes, thank you', 'no, delete event'],
    #     ).execute() == 'no, delete event':
    #     event.delete()



if __name__ == '__main__':
    main()
