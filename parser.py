import openpyxl 
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from collections import namedtuple
import re


Fieldstuple = namedtuple('Fields', 'varieties costumers numbers pieces \
								   	totals    prices    amounts codes  \
								    								   \
								    codes_singleUse numbers_singleUse rates_singleUse')


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
	values_list = page[data_beginning_row]

	mask = {1: (int,), 2: (int,), 3: (int, float),
			4: (int, float), 5: (str,), 6: (int,)}
	pos = 1
	beginning = None
	ending = None
	for i in values_list:
		if type(i.value) in mask[pos]:
			if beginning is None:
				beginning = i.column
			pos +=1
			if pos == 6:
				ending = i.column+1
				break
			continue
		else:
			beginning = None
			pos = 1
	if beginning is None:
		raise ValueError(f"can't find 'Number' 'Pieces' 'Total' 'Price' 'Amount' columns layout on the {sheet}" )
	return tuple(get_column_letter(i) for i in range(beginning, ending+1))


def find_singleUse_section(book, sheet:str):
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
		print(f'no "Single use packaging" section finded on "{sheet}" in {book}')
		return
	return data_beginning_row, data_beginning_row + data_range - 1


def find_quantities_singleUse(book, sheet:str, singleUse_range):
	page = book.get_sheet_by_name(sheet)  
	data_beginning_row = singleUse_range[0]
	headline_row = str(data_beginning_row - 1)
	quantity_columns = [cell.column for cell in page[headline_row] 
								            if cell.value in ('Number', 'Rate')]
	return tuple(get_column_letter(i) for i in quantity_columns)


def get_range_from_column(book, sheet:str, col_range:tuple, column_name:str):
		"""takes tuple containing (bigining-, end-) column numbers 
		from find-functions and returns values of cells in this range
		"""
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


def is_longFormat_date(date:str):
	result = re.match(r'\d{2}.\d{2}.\d{4}', date)
	return not isinstance(result, type(None))


def is_rows_merged(rowData:list):
	cell = str(rowData[0])
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


def check_numbers(numbers:list, book, sheet, column_name):
	row_numb=0
	for numb in numbers:
		row_numb+=1
		numb = str(numb)
		for i in numb:
			if not i.isdigit() and i != '.':
				raise ValueError(f'Error during checking Number "{numb}" value (row {row_numb}) in {book} on {sheet}')
	

def check_rate(rates:list, book, sheet):
	"""cheks is format of string is like 'dd.mm.yyyy' or not"""
	for rate in rates:
		rate = str(rate)
		res = re.match(r'\d{1},\d{2}', rate)
		if isinstance(res, type(None)):
			raise ValueError(f'Error during checking Rate value "{rate}" in {book} on {sheet}')


def correct_priece_format(price:str, book, sheet:str):
	if not isinstance(price, int):
		raise ValueError(f'ValueError while convert to right format value {price}\
							      from column Prise in book {book}, page {sheet}')
	price = str(price)
	if len(price) == 3:
		return float(f'0.{price}')
	elif len(price) > 3:
		return float(f'{price[:-3]}.{price[-3:]}')
	else:
		raise ValueError(f'ValueError: while convert to right format value {price}\
							      from column Prise in book {book}, page {sheet}')

def correct_totals_format(totals:list):
	"""corrects 'total' values like 1,600 that retrieved as floats -> 1.6 
	to correct format -> 1600, int
	"""
	result = []
	for i in totals:
		if isinstance(i, float):
			i = str(i).replace('.', '')
			i = '{0:0<4}'.format(i)
			result.append(int(i))
		else:
			result.append(i)
	return result


def adopt_float_format(rate_value):
	"""make rate record suitable to convert to float 
	and convert them via float(rate)"""
	result = re.sub(',', '.', rate_value)
	return float(result)


def parse(file):
	book = load_workbook(file)
	sheets = book.get_sheet_names()[1:]  #list of sheets without title-page
	
	varieties = []; costumers = []; numbers = []; pieces = []
	totals = []   ; prices = []   ; amounts = []; codes = []
	codes_singleUse = []; numbers_singleUse = []; rates_singleUse = []
	
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
			check_numbers(number, book, sh, 'number')
			numbers += number

			piece = get_range_from_column(book, sh, mainData_range, quantity_colums[1])
			check_numbers(piece, book, sh, 'piece')
			pieces += piece

			total = get_range_from_column(book, sh, mainData_range, quantity_colums[2])
			check_numbers(total, book, sh, 'total')
			total = correct_totals_format(total)
			totals += total

			price = get_range_from_column(book, sh, mainData_range, quantity_colums[3])
			price = [correct_priece_format(i, book, sh) for i in price]
			prices += price

			amount = get_range_from_column(book, sh, mainData_range, quantity_colums[4])
			amount = [adopt_float_format(i) for i in amount]
			amounts += amount

			code = get_range_from_column(book, sh, mainData_range, quantity_colums[5])
			codes += code

		singleUse_range = find_singleUse_section(book, sh)
		if singleUse_range is not None:

			code_singleUse = get_range_from_column(book, sh, singleUse_range, 'B')
			if len(code_singleUse) == 1 and is_rows_merged(code_singleUse):
				code_singleUse = [int(i) for i in code_singleUse[0].split('\n')]
			check_codes(code_singleUse, book, sh)
			codes_singleUse += code_singleUse

			quantities_singleUse = find_quantities_singleUse(book, sh, singleUse_range)

			number_singleUse = get_range_from_column(book, sh, singleUse_range, quantities_singleUse[0])
			if len(number_singleUse) == 1 and is_rows_merged(number_singleUse):
				number_singleUse = [int(i) for i in number_singleUse[0].split('\n')]
			check_numbers(number_singleUse, book, sh, 'number_singleUse')
			numbers_singleUse += number_singleUse

			rate_singleUse = get_range_from_column(book, sh, singleUse_range, quantities_singleUse[1])
			if len(rate_singleUse) == 1 and is_rows_merged(rate_singleUse):
				rate_singleUse = [i for i in rate_singleUse[0].split('\n')]
			check_rate(rate_singleUse, book, sh)
			rate_singleUse = [adopt_float_format(i) for i in rate_singleUse]
			rates_singleUse += rate_singleUse

	retrieved_data = Fieldstuple(varieties, costumers, numbers, pieces, 
								 totals,    prices,    amounts, codes, 
								 codes_singleUse, numbers_singleUse, rates_singleUse)

	return retrieved_data


if __name__ == '__main__':
	file = 'rawxl/pdfFile33.xlsx'
	# dt = parse(file)
	wb = load_workbook(file)

	
	d = parse(file)
	print(d.totals)
		




	

	





	


