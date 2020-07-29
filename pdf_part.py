import pdftables_api, requests
from closed import api_key

#https://pdftables.com/pdf-to-excel-api
convertor = pdftables_api.Client(api_key)

def remaining_conversions():
	req = requests.get(f'https://pdftables.com/api/remaining?key={api_key}')
	return f'{int(req.text)} pages available to convert'

def make_same_fileName(file_name):
	splited_path = file_name.split('/')
	primar_name = splited_path[-1]
	splited_primar = primar_name.split('.')
	if 'pdf' not in splited_primar:
		raise ValueError("ValueError: file extension must be pdf")
	return splited_primar[0] + '.xlsx'

def convert_pdf_to_xlsl(pdf_path, xlsx_path):
	convertor.xlsx(pdf_path, xlsx_path + make_same_fileName(pdf_path))
	print(f'Done. For this api_key remains {remaining_conversions()}')

