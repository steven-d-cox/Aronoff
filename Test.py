import datetime as dt 



def changeTime(tm, hours):
	date = dt.datetime(100, 1, 1, tm.hour, tm.minute, tm.second)
	date = date + dt.timedelta(hours=hours)
	return date.time()

def differentTime(now):
	time = (dt.datetime.combine(date.today(), now) + timedelta(hours=2)).time()

now = dt.datetime.now().time()
before = changeTime(now, -2)
before2 = changeTime(now, -2.5)
test = differentTime(now)

print(now)
print(before)
print(before2)
print(test)


def columnWidth():
	ws.column_dimensions['B'].width = 20
	ws.column_dimensions['G'].width = 40
	ws.column_dimensions['I'].width = 40