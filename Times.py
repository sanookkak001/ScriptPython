import datetime as dt

Beforehour = int(input("Enter Beforehour : "))
Beforeminute = int(input("Enter Beforeminute: "))

Afterhour = int(input("Enter Afterhour : "))
Afterminute = int(input("Enter Afterminute : "))

timestart = dt.datetime(year=2024,month=6,day=20,hour=Beforehour,minute=Beforeminute)
timemend = dt.datetime(year=2024,month=6,day=20,hour=Afterhour,minute=Afterminute)
timeanswer = timemend - timestart
total_minutes = timeanswer.total_seconds() / 60

print(f"ผลต่างระหว่างเวลาคือ {timeanswer}")
print(f"ผลต่างในนาทีคือ {total_minutes} นาที")
