# import required libraries
from pandas import ExcelFile, ExcelWriter

# custom class for reading and writing data in excel
class ExcelFormatting:
	"""
	Required Params:
		- excel_input_file (e.g., "file.xlsx")
	Default Params:
		- datetime_format = 'm/dd/yyyy' (sets the default date format when writing back to excel)
		- date_format = 'm/dd/yyyy' (sets the default date format when writing back to excel)

	Methods:
		- excel_file_read: returns the excel file pandas object. You can access the ExcelFile pandas class attribute through inheritiance.
		- excel_file_write: returns the excel file name and datetime properties to be write back to excel in pandas ExcelWriter class
			Required Params:
				excel_file_out: default = "excelfile.xlsx"

	"""
	# initialize default variables/properties
	def __init__(self, datetime_format='m/dd/yyyy', date_format='m/dd/yyyy'):
		self.datetime_format = datetime_format
		self.date_format = date_format

	# method to return excel file, inherits from pandas ExcelFileClass
	def excel_file_read(self, excel_file_input):
		file = ExcelFile(excel_file_input)
		return file

	# method to write excel data back to excel
	def excel_file_write(self,excel_file_out='excelfile.xlsx'):

		# default excel outfile
		excel_file_out = excel_file_out

		# define the writer properties
		writer = ExcelWriter(
			excel_file_out,
			date_format=self.date_format,
			datetime_format=self.datetime_format
		)
		return writer
