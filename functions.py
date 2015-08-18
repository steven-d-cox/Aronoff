from openpyxl import load_workbook
from openpyxl import Workbook

write = Workbook()
ws = write.active

class functions:
	
	def columnWidth(self, dimensions):

		ws.column_dimensions['B'].width = 20
		ws.column_dimensions['G'].width = 40
		ws.column_dimensions['I'].width = 40

	def header(self):
		
		ws['A1'] = "Day"
		ws['B1'] = "Date"
		ws['C1'] = "Venue"
		ws['D1'] = "E"
		ws['E1'] = "Status"
		ws['F1'] = "Time"
		ws['G1'] = "Event"
		ws['H1'] = "FS Initials"
		ws['I1'] = "Comments"

