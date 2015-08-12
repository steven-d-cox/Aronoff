from openpyxl import load_workbook
# Targets
# A B C D H J K + 1 new columns

wb = load_workbook('Sept 15 Cal.xlsx')
sheet = wb['Sept 2015']

for i in range(1,200):
	if sheet['D' + str(i)].value != None:
		if ('P' in sheet['D' + str(i)].value) and (sheet['H' + str(i)].value == 'FIRM'):
			A = sheet[str(i)].value
	
