from pdfpart import *
from parser import *
from operate_and_write import *
import platform
import sys


def open_file(file):
	if platform.system() == 'Linux':
		os.system(f'xdg-open {file}')
	elif platform.system() == 'Windows':
		os.startfile(file)


def create_dirs():
	"""if not exists, creates folder for raw, intermediate xlsx-files converted 
	via API("rawxl") and folder for final xlsx files ("out")"""
	try:
		os.makedirs('out/')
		os.makedirs('rawxl/')
	except:
		pass


def convert_file(source_pdf):
	source_xl = convert_pdf_to_xlsl(source_pdf, './rawxl/') 
	retrieved_data = parse(f'./rawxl/{source_xl}')
	calculated_data = calc_packings(retrieved_data) 
	output_xl = write_results(sample_file='sample.xlsx',
				  outp_filename=f'./out/{source_xl}',
				  parsed=retrieved_data,
				  calculated=calculated_data)
	return output_xl


if __name__ == '__main__':
	create_dirs()
	converted_file = convert_file(str(sys.argv[1]))
	open_file(f'{converted_file}')