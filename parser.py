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
				raise ValueError("Error: 'date'-row is found but float-type date-cell don't follow further")
		if data_beginning_row is not None and type(cell.value) == float:
			data_range += 1
			# print('value:', cell.value, cell.row, 'count:', data_range)
		elif data_beginning_row is not None and data_range > 1:
			# print('end of range sirching wiht result:', data_range)
			break
	if data_beginning_row is None:
		print(f'no main data finded on "{sheet}"')
		return
	# print(book, 'len =', data_range, f'range = A{data_beginning_row}:A{data_beginning_row+data_range-1}', sep=',')
	return data_beginning_row, data_beginning_row + data_range - 1


def find_quantity_columns(book, sheet:str):
		page = book.get_sheet_by_name(sheet) #tmp
		mainData_range = find_main_data(book, sheet)  
		data_beginning_row = str(mainData_range[0])
		beginning = None
		end = None
		for cell in page[data_beginning_row]:
			# print('---ITER---:', cell.value)
			# print('BEGIN:', beginning)
			# print('END:', end)
			if not isinstance(cell.value, int):
				# print(cell.value, 'is not int')
				if beginning is not None:
					# print(f'for {cell.value} beginning {beginning} is not None')
					if not isinstance(cell.value, str):
						# print(cell.value, 'is not float, beginning is set to None')
						beginning = None
						continue	
					elif ',' in cell.value:  #TODO: REG digit-digit-comma-digit-digit
						# print(f'for {cell.value} else, {cell.row} is end')
						end = cell.column
						break
				continue
			elif beginning is None:
				# print(f'elif: {cell.value} is beginning with N {cell.column}')
				beginning = cell.column
		if not end - beginning == 4:
			raise ValueError("Error during defining range of columns containing quantity values")
		return (beginning, end)
			

def get_column(book, sheet:str, cells):
	page = book.get_sheet_by_name(sheet)
	return [i.value for i in page[cells]]


def get_range_from_column(book, sheet:str, col_range:tuple, column_name:str):
		page = book.get_sheet_by_name(sheet)
		return [i[0].value for i in 
					page[f'{column_name}{col_range[0]}':f'{column_name}{col_range[1]}']]


def is_varieties_or_costumers(data_list:list):
	for i in data_list:
		if isinstance(i, str) and i[0].isdigit() and ' ' in i:
			continue
		else:
			return False
	return True


def pars(book):
	sheets = book.get_sheet_names()[1:]  #list of sheets without title-page
	varieties = []
	costumers = []
	
	for sh in sheets:
		mainData_range = find_main_data(book, sh)
		if mainData_range is not None:

				variety = get_range_from_column(book, sh, mainData_range, 'C')
				if not is_varieties_or_costumers(variety):
					raise ValueError(f"Error in column 'variety' in {book} on page '{sheet}'")
				varieties += variety
				
				costumer = get_range_from_column(book, sh, mainData_range, 'F')
				if not is_varieties_or_costumers(costumer):
					raise ValueError(f"Error in column 'costumer' in {book} on page '{sheet}'")
				costumers += costumer

	return varieties, costumers


if __name__ == '__main__':
	file = '/home/emil/Загрузки/out/pdfFile4.xlsx'
	wb = load_workbook(file)
	sheet = 'Page 2'

	print(find_quantity_columns(wb, sheet))


