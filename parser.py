from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import re


def pull_up_files(file_list:list):
	return [load_workbook(f) for f in file_list]


def find_main_data(book, sheet:str):
	page = book.get_sheet_by_name(sheet)
	data_beginning_row = None
	data_range = 0
	for cell in page['A']:
		if cell.value == 'date':
			data_beginning_row = cell.row + 1 
			if isinstance(page[f'A{data_beginning_row}'].value, float):  # check of row follows next to data_beginng row 
				continue
			else:
				raise ValueError("Error: 'date'-row is found but float-type date-cell don't follow further")
		if data_beginning_row is not None and type(cell.value) == float:
			data_range += 1
		elif data_beginning_row is not None and data_range > 1:
			break
	if data_beginning_row is None:
		print(f'no main data finded on "{sheet}"')
		return
	return data_beginning_row, data_beginning_row + data_range - 1


def find_quantity_columns(book, sheet:str, mainData_range):
		page = book.get_sheet_by_name(sheet)  
		data_beginning_row = str(mainData_range[0])
		beginning = None
		end = None
		for cell in page[data_beginning_row]:
			## print('---ITER---:', cell.value)
			## print('BEGIN:', beginning)
			## print('END:', end)
			if not isinstance(cell.value, int):
				## print(cell.value, 'is not int')
				if beginning is not None:
					## print(f'for {cell.value} beginning {beginning} is not None')
					if not isinstance(cell.value, str):
						## print(cell.value, 'is not float, beginning is set to None')
						beginning = None
						continue	
					elif ',' in cell.value:  #TODO: REG digit-digit-comma-digit-digit
						## print(f'for {cell.value} else, {cell.row} is end')
						end = cell.column + 1  # +1 is for include 'code' column, that follows quite after amount ( ',' cell.value)
						break
				continue
			elif beginning is None:
				## print(f'elif: {cell.value} is beginning with N {cell.column}')
				beginning = cell.column
		if not end - beginning == 5:
			raise ValueError("Error during defining range of columns containing quantity values")
		return tuple(get_column_letter(i) for i in range(beginning, end+1))


def is_longFormat_date(date:str):
	result = re.match(r'\d{2}.\d{2}.\d{4}', date)
	return not isinstance(result, type(None))


def find_singleUse_data(book, sheet:str):
	page = book.get_sheet_by_name(sheet)
	data_beginning_row = None
	data_range = 0
	c = 0
	for cell in page['A']:
		if cell.value == 'Single use packaging': 
			if not page[f'A{cell.row+1}'].value == 'Date':  
				raise ValueError('error during searching singleUse_range: \
								 "Date" row don`t follow after "Single use packaging" row')
			data_beginning_row = cell.row + 2
			continue
		if data_beginning_row is not None:
			if is_longFormat_date(cell.value):
				data_range += 1
				continue
			elif cell.value == 'Total':
				break
	if data_beginning_row is None:
		print(f'no "Single use packaging" data finded on "{sheet}"')
		return
	return data_beginning_row, data_beginning_row + data_range - 1


def find_quantities_singleUse(book, sheet:str, singleUse_range):
	page = book.get_sheet_by_name(sheet)  
	data_beginning_row = singleUse_range[0]
	headline_row = str(data_beginning_row - 1)
	quantity_columns = [cell.column for cell in page[headline_row] 
								            if cell.value in ('Number', 'Rate')]
	return tuple(get_column_letter(i) for i in quantity_columns)


def is_rows_merged(rowData:list):
	cell = rowData[0]
	if '\n' in cell:
		return True
	return False 


def check_codes(codes:list, book, sheet):
	for string in codes:
		string = str(string)
		if not len(string) == 3:
			raise ValueError(f'code value {sheet} in {book} on {sheet} does not look like code')
		for i in string:
			if not i.isdigit():
				raise ValueError(f'code value {sheet} in {book} on {sheet} does not look like code')
	return True


	
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


def correct_priece_format(price:str, book, sheet:str):
	if not isinstance(price, int):
		raise ValueError(f'ValueError: while convert to right format value {price}\
							      from column Prise in book {book}, page {sheet}')
	price = str(price)
	if len(price) == 3:
		return f'0,{price}'
	elif len(price) > 3:
		return f'{price[:-3]},{price[-3:]}'
	else:
		raise ValueError(f'ValueError: while convert to right format value {price}\
							      from column Prise in book {book}, page {sheet}')


def parse(book):
	sheets = book.get_sheet_names()[1:]  #list of sheets without title-page
	varieties = []; costumers = []; numbers = []; pieces = []
	totals = []   ; prices = []   ;amounts = [] ; codes = []

	codes_singleUse = []
	
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

			quantity_colums = find_quantity_columns(book, sh, mainData_range)

			number = get_range_from_column(book, sh, mainData_range, quantity_colums[0])
			numbers += number

			piece = get_range_from_column(book, sh, mainData_range, quantity_colums[1])
			pieces += piece

			total = get_range_from_column(book, sh, mainData_range, quantity_colums[2])
			totals += total

			price = get_range_from_column(book, sh, mainData_range, quantity_colums[3])
			price = [correct_priece_format(i, book, sh) for i in price]
			prices += price

			amount = get_range_from_column(book, sh, mainData_range, quantity_colums[4])
			amounts += amount

			code = get_range_from_column(book, sh, mainData_range, quantity_colums[5])
			codes += code

		singleUse_range = find_singleUse_data(book, sh)
		if singleUse_range is not None:

			code_singleUse = get_range_from_column(book, sh, singleUse_range, 'B')
			if len(code_singleUse) == 1 and is_rows_merged(code_singleUse):
				code_singleUse = [i for i in code_singleUse[0].split('\n')]
			check_codes(code_singleUse, book, sh)
			codes_singleUse += code_singleUse





	return varieties, costumers, numbers, pieces, totals, prices, amounts, codes, codes_singleUse


if __name__ == '__main__':
	file = '/home/emil/Загрузки/out/pdfFile.xlsx'
	wb = load_workbook(file)
	sheet = wb.get_sheet_by_name('Page 4')

	sr = find_singleUse_data(wb, 'Page 4')

	qc = find_quantities_singleUse(wb, 'Page 4', sr)
	print(qc)

	# b = get_range_from_column(wb, 'Page 3', col_range=a, column_name='A')
	# print(b)

	

	





	


