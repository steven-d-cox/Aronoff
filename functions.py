from openpyxl import load_workbook
from openpyxl import Workbook

def columnWidth(writer, dimensions):
	for d in dimensions:
		writer.column_dimensions[d[0]].width = d[1]
	return writer
