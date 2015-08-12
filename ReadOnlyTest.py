from openpyxl import load_workbook

# Targets
# A B C D H J K + 1 new columns

wb = load_workbook('Sept 15 Cal.xlsx', read_only=True)
sheet = wb['Sept 2015']

select = []

for row in sheet.rows:
	if row[3].value:
		if 'P' in row[3].value and row[7].value == 'FIRM':
			select.append((row[0].value, row[1].value, row[2].value, row[3].value, row[7].value, row[9].value, row[10].value))
print(select)
