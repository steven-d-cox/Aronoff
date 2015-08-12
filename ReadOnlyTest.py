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
#print(len(select))

# Writing to the file
write = Workbook()
ws = write.active
for i in range(len(select)):
	ws['A'+str(i+1)] = select[i][0]
	ws['B'+str(i+1)] = select[i][1]
	ws['C'+str(i+1)] = select[i][2]
	ws['D'+str(i+1)] = select[i][3]
	ws['E'+str(i+1)] = select[i][4]
	ws['F'+str(i+1)] = select[i][5]
	ws['G'+str(i+1)] = select[i][6]
write.save("sampleout.xlsx")
