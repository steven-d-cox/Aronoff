from openpyxl import load_workbook
from openpyxl import Workbook
# Targets
# A B C D H J K + 1 new columns

wb = load_workbook('Sept 15 Cal.xlsx', read_only=True)
sheet = wb['Sept 2015']

select = []

for row in sheet.rows:
	if row[3].value:
		if 'P' in row[3].value and row[7].value == 'FIRM':
			select.append((row[0].value, row[1].value, row[2].value, row[3].value, row[7].value, row[9].value, row[10].value))

# Writing to the file
write = Workbook()
ws = write.active
# Write the header
ws['A1'] = "Day"
ws['B1'] = "Date"
ws['C1'] = "Venue"
ws['D1'] = "E"
ws['E1'] = "Status"
ws['F1'] = "Time"
ws['G1'] = "Event"
ws['H1'] = "FS Initials"
ws['I1'] = "Comments"
for i in range(len(select)):
	ws['A'+str(i+2)] = select[i][0]
	ws['B'+str(i+2)] = select[i][1]
	ws['C'+str(i+2)] = select[i][2]
	ws['D'+str(i+2)] = select[i][3]
	ws['E'+str(i+2)] = select[i][4]
	ws['F'+str(i+2)] = select[i][5]
	ws['G'+str(i+2)] = select[i][6]
write.save("sampleout.xlsx")
