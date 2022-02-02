Iakov has already made a wonderful slots tool: https://github.com/iakov-aws/slots



# Slots

## About

This tool shows all slots where attendees are free. It uses Microsoft Outlook to read agenda. Windows and Mac are supported

## Prerequisites
* mac or windows
* python3.9+
* Microsoft Outlook


## Example
```
python3 slots.py  --attendees 'attendee1@domain.com; Atten Dee <attendee2@domain>' -tz 'America/New_York'
? Select a Timeslot to create a meeting:
  Cancel (CTRL+C)
  ---------------
  Thursday 20 January 14:00 - 15:00 CET / 08:00 - 09:00 America/New_York
  Thursday 20 January 16:00 - 18:00 CET / 10:00 - 12:00 America/New_York
  Friday 21 January 14:00 - 15:00 CET / 08:00 - 09:00 America/New_York
 >Friday 21 January 15:30 - 18:00 CET / 09:30 - 12:00 America/New_York
  Monday 24 January 15:00 - 18:00 CET / 09:00 - 12:00 America/New_York
  Tuesday 25 January 15:00 - 18:00 CET / 09:00 - 12:00 America/New_York
  Wednesday 26 January 16:00 - 18:00 CET / 10:00 - 12:00 America/New_York
  Thursday 27 January 14:00 - 15:00 CET / 08:00 - 09:00 America/New_York

```
python3 slots.py  --attendees 'waltmayf@amazon.com' --end 'in 3 days' --length 60
python3 slots.py  --attendees 'waltmayf@amazon.com;chewmaur@amazon.com' --end 'in 3 days' --length 60

## Install
### mac
```
pip install -r requirements/macos.txt
```
### win
```
pip install -r requirements/windows.txt
```


## Advanced usage
Show all slots where 5O% of attendees are free and show alternative timezone:
```
python3 slots.py  --attendees 'attendee1@domain.com; Atten Dee <attendee2@domain>' --rate 50 -tz America/New_York
```

Show full agenda:
```
python3 slots.py  --attendees 'attendee1@domain.com; Atten Dee <attendee2@domain>' --rate 50  --full
```

More:
```
❯ python3 slots.py --help
Usage: slots.py [OPTIONS]

Options:
  -a, --attendees TEXT          A semicolon separated list of contacts
  --start TEXT                  Begining of the date range to find available
                                slots. Iso date or Arrow Humanised time (ex:
                                "in 30 days"). See https://arrow.readthedocs.io/en/latest/#dehumanize. 
                                Default is today.
  --end TEXT                    The end of the date range to find available
                                slots. Iso date or Arrow Humanised time (ex:
                                "in 30 days"). See https://arrow.readthedocs.io/en/latest/#dehumanize. 
                                Default is "in 7
                                days".
  --full / --only-slots         Show slots or a full agenda (default=only
                                slots)
  -r, --rate INTEGER            Acceptable share of available attendees in
                                percent(%)
  --tentative / --no-tentative  Treat tentative meetings as free
  -l, --lenght INTEGER          Length of the meeting, in minutes 
  -tz, --alternative_tz TEXT    A comma separated list of alternative
                                Timezones. if "auto" script will fetch
                                timezones from slack (mwinit --aea requied).
  -hr, --hours TEXT             Find availability between the hours of...
                                Default = "0800-1900".
  -f, --fmt TEXT                Time format in list of available slots.
                                "HH:mm" or "h:mma". Refer to https://arrow.rea
                                dthedocs.io/en/latest/#supported-tokens.
                                Default "HH:mm".
  --help                        Show this message and exit.
 ```
