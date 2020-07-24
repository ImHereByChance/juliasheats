from openpyxl import load_workbook

# file = '/home/emil/Загрузки/out/pdfFile.xlsx'

# wb = load_workbook(file)
# sheets = wb.get_sheet_names()[1:]  #list of sheets without title-page


def pull_up_files(file_list:list):
	return [load_workbook(f) for f in file_list]


def get_column(book, sheet, cells):
	page = book.get_sheet_by_name(sheet)
	return [i.value for i in page[cells]]


def get_flowerVariety(raw_data:list):
	return [i for i in raw_data if isinstance(i, str) and i[0].isdigit()]
