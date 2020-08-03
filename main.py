from pdfpart import *
from parser import *
from operate_and_write import *
import platform


def open_file(file):
	if platform.system() == 'Linux':
		os.system(f'xdg-open {file}')
	elif platform.system() == 'Windows':
		os.startfile(file)


def main():
	source_pdf = input('filename or path: ')
	source_xl = convert_pdf_to_xlsl(source_pdf, './rawxl/') 
	retrieved_data = parse(f'./rawxl/{source_xl}')
	calculated_data = calc_singleUse_cost(retrieved_data) 
	output_xl = write_results(sample_file='sample.xlsx',
							  outp_filename=f'./out/{source_xl}',
							  parsed=retrieved_data,
							  calculated=calculated_data)

	open_file(f'{output_xl}')


if __name__ == '__main__':
	main()
	
