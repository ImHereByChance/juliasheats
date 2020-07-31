from parser import *

def make_rate_dict(codes:list, rates:list):
	return dict(zip(codes, rates))

def calc_singleUse_cost(rates:dict, codes:list, numbers:list):
	costs_singleUse = []
	index = 0
	for i in codes:
		rate = rates[i]
		number = numbers[index]
		cost = rate * number
		cost = round(cost, 2)
		print(i, index, rate, number, cost)
		costs_singleUse.append(cost)
		index +=1
	return costs_singleUse

if __name__ == '__main__':
	file = '/home/emil/Загрузки/out/pdfFile4.xlsx'
	wb = load_workbook(file)
	
	dt = parse(wb)

	dictt = make_rate_dict(dt.codes_singleUse, dt.rates_singleUse)

	print(calc_singleUse_cost(dictt, dt.codes, dt.numbers))



	

