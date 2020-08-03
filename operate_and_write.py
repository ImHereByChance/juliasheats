from parser import *
import os ###TMP


def make_rates_dict(codes_singleUse:list, rates_singleUse:list):
	return dict(zip(codes_singleUse, rates_singleUse))


def calc_singleUse_cost(data):
	rates_codes = make_rates_dict(data.codes_singleUse, data.rates_singleUse)
	costs_singleUse = []
	index = 0
	for i in data.codes:
		try:
			rate = rates_codes[i]
		except KeyError:
			costs_singleUse.append('MULTI')
			continue
		number = data.numbers[index]
		cost = round(rate * number, 2)
		costs_singleUse.append(cost)
		index +=1
	return costs_singleUse


def write_column(sheet, data:list, starting_row, column):
		for i in data:
			sheet.cell(starting_row, column).value = i
			starting_row += 1


def write_results(sample_file, outp_filename, parsed, calculated): 
	wb = load_workbook(sample_file)
	sheet = wb.active

	write_column(sheet, parsed.costumers, 5, 4)   # Поставщик
	write_column(sheet, parsed.varieties, 5, 5)   # Сорт
	write_column(sheet, parsed.prices,    5, 6)   # Цена за стебель
	write_column(sheet, parsed.pieces,    5, 7)   # Кол-во в коробке
	write_column(sheet, parsed.numbers,   5, 8)   # Коробок шт.
	write_column(sheet, parsed.totals,    5, 9)   # Всего стеблей
	write_column(sheet, parsed.amounts,   5, 10)  # Purchases of products (NE)
	write_column(sheet, calculated,       5, 14)  # Packaging deposit by Connect (AG)
 
	wb.save(outp_filename)

	return outp_filename


if __name__ == '__main__':
	file = 'pdfFile49.xlsx'
	
	dt = parse(file)

	calcl = calc_singleUse_cost(dt)

	write_results('sample.xlsx', './out/qwakozyabra.xlsx' ,dt, calcl)

	os.system("xdg-open ./out/qwakozyabra.xlsx")



	
