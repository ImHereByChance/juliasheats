import pdftables_api

#https://pdftables.com/pdf-to-excel-api
apiKey_pdftables = 'ouyf4csccg1l'
convertor = pdftables_api.Client(apiKey_pdftables)

def make_same_fileName(file_name):
	splited_path = file_name.split('/')
	primar_name = splited_path[-1]
	splited_primar = primar_name.split('.')
	if 'pdf' not in splited_primar:
		raise ValueError("ValueError: file extension must be pdf")
	return splited_primar[0] + '.xlsx'


def convert_pdf_to_xlsl(pdf_path, xlsx_path):
	convertor.xlsx(pdf_path, xlsx_path + make_same_fileName(pdf_path))
