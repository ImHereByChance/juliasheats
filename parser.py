from openpyxl import load_workbook


def pull_up_files(file_list:list):
	return [load_workbook(f) for f in file_list]


def find_additional_data(book, sheet:str):
	pass


def find_main_data(book, sheet:str):
	page = book.get_sheet_by_name(sheet)
	
	data_beginning_row = None
	data_range = 0
	for cell in page['A']:
		if cell.value == 'date':
			data_beginning_row = cell.row + 1 
			if isinstance(page[f'A{data_beginning_row}'].value, float):  # check of row follows next to data_beginng row 
				# print('begin:', data_beginning_row)
				continue
			else:
				raise Error("Error: 'date'-row is found but no float-type date-cells follows further")
		if data_beginning_row is not None and type(cell.value) == float:
			data_range += 1
			# print('value:', cell.value, cell.row, 'count:', data_range)
		elif data_beginning_row is not None and data_range > 1:
			# print('end of range sirching wiht result:', data_range)
			break
	if data_beginning_row is None:
		print(f'no main data finded on `{sheet}`')
		return

	# print(book, 'len =', data_range, f'range = A{data_beginning_row}:A{data_beginning_row+data_range-1}', sep=',')
	return data_beginning_row, data_beginning_row + data_range - 1


def get_column(book, sheet:str, cells):
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


if __name__ == '__main__':
	file = '/home/emil/Загрузки/out/pdfFile4.xlsx'
	wb = load_workbook(file)
	sheet = 'Page 2'

	a = find_main_data(wb, sheet)

	print(a)

