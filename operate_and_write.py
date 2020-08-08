from parser import *
import os ###TMP


Calculated = namedtuple('fields', 'costs_singleUse costs_multiUse_deposts costs_multiUse_rents')


def make_codesRates_dict(codes_singleUse:list, rates_singleUse:list):
	"""{codes_codes_singleUse: rates_singleUse, ...etc.}""" 
	return dict(zip(codes_singleUse, rates_singleUse))


def filter_repeated_codes(codes_multiUse):
	"""[1, 1, 2, 2] -> [1, 2]"""
	filtred_codes = []
	for code in codes_multiUse:
		if not code in filtred_codes:
			filtred_codes.append(code) 
	return filtred_codes


def make_codesPayments_dict(codes_multiUse, deposits_multiUse, rents_multiUse):
	"""codes_multiUse; deposits_multiUse; rents_multiUse ->
	-> {codes_multiUse: (deposits_multiUse, rents_multiUse)}"""
	codes_multiUse = filter_repeated_codes(codes_multiUse)
	depRen_tuples = zip(deposits_multiUse, rents_multiUse)
	codes_and_DepRenTuples = zip(codes_multiUse, depRen_tuples)  #(code_multiUse, (deposits_multiUse, rents_multiUse), ...etc.)
	return dict(codes_and_DepRenTuples)


def calc_packings(data):
	codesRates_dict = make_codesRates_dict(data.codes_singleUse, data.rates_singleUse)
	codesPayments_dict = make_codesPayments_dict(data.codes_multiUse, data.deposits_multiUse, 
																	  data.rents_multiUse)
	costs_singleUse = []
	costs_multiUse_deposts = []
	costs_multiUse_rents =[]

	index = 0
	for code in data.codes:
		if code in codesRates_dict:
			rate = codesRates_dict[code]
			number = data.numbers[index]
			cost = round(rate * number, 2)
			
			costs_singleUse.append(cost)
			costs_multiUse_deposts.append(None)
			costs_multiUse_rents.append(None)
			index +=1
		elif code in codesPayments_dict:	
			payment = codesPayments_dict[code]
			number = data.numbers[index]
			deposit = payment[0]
			dep_cost = round(deposit * number, 2) 
			
			rent = payment[1]
			rent_cost = round(rent * number, 2)
			
			costs_multiUse_deposts.append(dep_cost)
			costs_multiUse_rents.append(rent_cost)
			costs_singleUse.append(None)
			index +=1
		else:
			costs_singleUse.append(None)
			costs_multiUse_deposts.append(None)
			costs_multiUse_rents.append(None)
			index +=1

	return Calculated(costs_singleUse, costs_multiUse_deposts, costs_multiUse_rents)


def write_column(sheet, data:list, starting_row, column):
		for i in data:
			sheet.cell(starting_row, column).value = i
			starting_row += 1


def write_results(sample_file, outp_filename, parsed, calcd): 
	wb = load_workbook(sample_file)
	sheet = wb.active

	write_column(sheet, parsed.costumers, 5, 4)   			  # Поставщик
	write_column(sheet, parsed.varieties, 5, 5)   			  # Сорт
	write_column(sheet, parsed.prices, 5, 6)      			  # Цена за стебель
	write_column(sheet, parsed.pieces, 5, 7)   				  # Кол-во в коробке
	write_column(sheet, parsed.numbers, 5, 8)   			  # Коробок шт.
	write_column(sheet, parsed.totals, 5, 9)  				  # Всего стеблей
	write_column(sheet, parsed.amounts, 5, 10)  			  # Purchases of products (NE)
	write_column(sheet, calcd.costs_singleUse, 5, 14)  		  # Single use packaging Connect (NE)
	write_column(sheet, calcd.costs_multiUse_deposts, 5, 15)  # Packaging deposit by Connect (AG)
	write_column(sheet, calcd.costs_multiUse_rents, 5, 16)    # Packaging rent by Connect (NE)
	
	wb.save(outp_filename)

	return outp_filename


if __name__ == '__main__':
	file = '/home/emil/Загрузки/multy/converted/multi16.xlsx'
	
	data = parse(file)
	calcl = calc_packings(data)

	write_results('sample.xlsx', './out/qwakozyabra.xlsx' ,data, calcl)

	os.system("xdg-open ./out/qwakozyabra.xlsx")


	
