from openpyxl import load_workbook

file = '/home/emil/Загрузки/out/pdfFile.xlsx'

wb = load_workbook(file)
sheets = wb.get_sheet_names()[1:]  #list of sheets without title-page


def get_column(sheet, cells):
	page = wb.get_sheet_by_name(sheet)
	return [i.value for i in page[cells]]


