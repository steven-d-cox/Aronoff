from openpyxl import load_workbook
from openpyxl import Workbook
import datetime as dt

# Writing to the file
write = Workbook()
ws = write.active

# Function for loading all the Floor supervisor availability
def loadWorkbook(form):
	wb = load_workbook(form, read_only=True)
	sheet = wb['Sheet']

	select = []

	for row in sheet.row: 
		if row[7].value != null: 
			select.append((row[0].value, row[1].value, row[2].value, row[3].value, row[4].value, row[5].value, row[6].value))


#def type():
#	if time is after 5pm = 'N'
#	if time is before 5pm = 'M'
#	if there are two shows in the 'P&G' in one day = 'D'



def main():
	# loop through and grab each Floor supervisor availability form 

	# write out to new file with supervisors initials added to appropriate row

	# include logic for the columns type, number of shifts, availability for each day 

	""" Write Stage """
	wrtite = Workbook()
	ws = write.active
	# Write the header
	ws['A1'] = "Day"
	ws['B1'] = "Date"
	ws['C1'] = "Venue"
	ws['D1'] = "E"
	ws['E1'] = "Status"
	ws['F1'] = "Time"
	ws['G1'] = "Event"
	ws['H1'] = "Type"
	ws['I1'] = "MOD"
	ws['J1'] = "JK"
	ws['K1'] = "5/3"
	ws['L1'] = "Shift"
	ws['M1'] = "Per Day"
	ws['N1'] = "Assigned"
	ws['O1'] = "Remaining"



	# Sum of shifts

	# Sum of per day

	# Sum of remaining 
















