import pdftables_api, requests
import os


# https://pdftables.com/pdf-to-excel-api
API_KEY = os.getenv('API_KEY')

convertor = pdftables_api.Client(API_KEY)


def remaining_conversions():
	req = requests.get(f'https://pdftables.com/api/remaining?key={API_KEY}')
	return f'{int(req.text)} pages available to convert'


def make_same_fileName(file_name):
	splited_path = file_name.split('/')
	primar_name = splited_path[-1]
	splited_primar = primar_name.split('.')
	if 'pdf' not in splited_primar:
		raise ValueError("file extension must be pdf")
	return splited_primar[0] + '.xlsx'


def convert_pdf_to_xlsl(pdf_path, xlsx_path):
	name = make_same_fileName(pdf_path)
	final_path = xlsx_path + name
	convertor.xlsx(pdf_path, final_path)
	print(f'Done. For this api_key remains {remaining_conversions()}')
	return name

if __name__ == '__main__':
	print(remaining_conversions())
