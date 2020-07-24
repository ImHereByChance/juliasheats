from openpyxl import load_workbook

# file = '/home/emil/Загрузки/out/pdfFile.xlsx'

# wb = load_workbook(file)
# sheets = wb.get_sheet_names()[1:]  #list of sheets without title-page


def pull_up_files(file_list:list):
	return [load_workbook(f) for f in file_list]


def get_column(book, sheet, cells):
	page = book.get_sheet_by_name(sheet)
	return [i.value for i in page[cells]]


def get_variety_and_customer(book):
	sheets = book.get_sheet_names()[1:]  #list of sheets without title-page
	variety = []
	costumer = []
	for sh in sheets:
		raw_variety = get_column(book, sh, 'C')
		variety += [i for i in raw_variety 
							if isinstance(i, str) and i[0].isdigit()]
		raw_costumer = get_column(book, sh, 'F')
		costumer += [i for i in raw_costumer if isinstance(i, str) 
								and i[0].isdigit() and ' ' in i]
	if len(variety) != len(costumer):
		raise Error('Error: len(variety) != len(costumer)')
	return variety, costumer

