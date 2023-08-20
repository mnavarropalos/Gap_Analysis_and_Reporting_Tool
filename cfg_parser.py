import json


class DataToExtractEntry:

	def __init__(self, from_row, to_row, from_column, to_column):
		"""

		:param from_row:
		:param to_row:
		:param from_column:
		:param to_column:
		"""
		self.from_row = from_row
		self.to_row = to_row
		self.from_column = from_column
		self.to_column = to_column


class CfgEntry:

	def __init__(self, sheet_name, token, ):
		"""

		:param sheet_name:
		:param token:
		"""
		self.sheet_name = sheet_name
		self.token = token
		self.data_to_extract = list()

	def add_data_to_extract(self, from_row, to_row, from_column, to_column):
		"""

		:param from_row:
		:param to_row:
		:param from_column:
		:param to_column:
		:return:
		"""
		if from_row > to_row:
			raise Exception("'from_row' is bigger than 'to_row' for token '%s' in JSON file" % self.token)

		if from_column > to_column:
			raise Exception("'from_column' is bigger than 'to_column' for token '%s' in JSON file" % self.token)

		self.data_to_extract.append(DataToExtractEntry(from_row, to_row, from_column, to_column))


class CfgParser:

	@staticmethod
	def parse(file_path):
		"""

		:param file_path:
		:return:
		"""
		cfg_entries_list = list()

		try:

			with open(file_path, 'r') as input_file:

				json_entries = json.load(input_file)

				for json_entry in json_entries:

					if "sheet_name" not in json_entry:
						raise Exception("'sheet_name' entry is not defined in JSON file")

					if "token" not in json_entry:
						raise Exception("'token' entry is not defined in JSON file")

					if "data_to_extract" not in json_entry:
						raise Exception("'data_to_extract' entry is not defined in JSON file")

					sheet_name = json_entry['sheet_name']
					token = json_entry['token']
					data_to_extract_list = json_entry['data_to_extract']

					if not data_to_extract_list:
						raise Exception("'data_to_extract' entry is empty in the JSON file")

					cfg_entry = CfgEntry(sheet_name, token)

					for data_to_extract_entry in data_to_extract_list:

						if "from_row" not in data_to_extract_entry:
							raise Exception("'from_row' entry is not defined in JSON file")

						if "to_row" not in data_to_extract_entry:
							raise Exception("'to_row' entry is not defined in JSON file")

						if "from_column" not in data_to_extract_entry:
							raise Exception("'from_column' entry is not defined in JSON file")

						if "to_column" not in data_to_extract_entry:
							raise Exception("'data_to_extract' entry is not defined in JSON file")

						from_row = data_to_extract_entry['from_row']
						to_row = data_to_extract_entry['to_row']
						from_column = data_to_extract_entry['from_column']
						to_column = data_to_extract_entry['to_column']

						cfg_entry.add_data_to_extract(from_row, to_row, from_column, to_column)

					cfg_entries_list.append(cfg_entry)

		except FileNotFoundError:
			print("[ERROR]Configuration file not found: %s" % file_path)
			return None

		except json.JSONDecodeError as ex:
			print("[ERROR]Invalid format for configuration file: %s: %s" % (file_path, ex))
			return None

		except Exception as ex:
			print("[ERROR]%s" % ex)
			return None

		if not cfg_entries_list:
			print("[ERROR]Could not find any valid entry in configuration file: %s" % file_path)
			return None

		return cfg_entries_list
