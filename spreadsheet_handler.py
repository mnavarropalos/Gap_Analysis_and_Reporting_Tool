import openpyxl


class SpreadsheetHandler:

	@staticmethod
	def extract(file_path, sheet_name, from_row, to_row, from_column, to_column):
		"""

		:param file_path:
		:param sheet_name:
		:param from_row:
		:param to_row:
		:param from_column:
		:param to_column:
		:return:
		"""
		try:

			extracted_data = []
			workbook = openpyxl.load_workbook(file_path)
			sheet = workbook[sheet_name]

			for row in range(from_row, to_row + 1):
				for column in range(from_column, to_column + 1):
					extracted_data.append(str(sheet.cell(row=row, column=column).value))

		except FileNotFoundError:
			print("[ERROR]Source file not found: %s" % file_path)
			return None

		except Exception as ex:
			print("[ERROR]File: %s, sheet: %s. %s" % (file_path, sheet_name, ex))
			return None

		if not extracted_data:
			print("[ERROR]Could not find any valid data in source file: %s" % file_path)
			return None

		return extracted_data

		return extracted_data
