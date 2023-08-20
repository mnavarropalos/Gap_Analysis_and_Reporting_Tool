import argparse
import os

from cfg_parser import CfgParser
from spreadsheet_handler import SpreadsheetHandler
from doc_handler import DocHandler


def parse(input_file, source_file, config_file, output_file):
	"""

	:param input_file:
	:param source_file:
	:param config_file:
	:param output_file:
	:return:
	"""
	cfg_data_list = CfgParser.parse(config_file)
	if cfg_data_list is None:
		return 1

	for cfg_data_entry in cfg_data_list:

		extracted_data = []
		for data_to_extract_cfg_entry in cfg_data_entry.data_to_extract:

			tmp_extracted_data = SpreadsheetHandler.extract(source_file, cfg_data_entry.sheet_name, data_to_extract_cfg_entry.from_row, data_to_extract_cfg_entry.to_row, data_to_extract_cfg_entry.from_column, data_to_extract_cfg_entry.to_column)
			if extracted_data is None:
				return 1
			extracted_data.extend(tmp_extracted_data)

		output_file = DocHandler.insert(input_file, output_file, cfg_data_entry.token, extracted_data)
		if output_file is None:
			return 1

		input_file = output_file

	return 0


if __name__ == "__main__":
	parser = argparse.ArgumentParser(description="Pollito")
	parser.add_argument("-i", "--input_file", type=str, required=True, help="Input word document where the <source> data will be added")
	parser.add_argument("-s", "--source_file", type=str, required=True, help="Input excel document containg the <source> data")
	parser.add_argument("-c", "--config_file", type=str, required=True, help="Input JSON file containg the description of which data to extract from <source>")
	parser.add_argument("-o", "--output_file", type=str, default=None, help="OPTIONAL: Generated word document. DEFAULT_VALUE: Same path as input file")

	args = parser.parse_args()

	if args.output_file is None:
		output_file_path = os.path.dirname(args.input_file)
		output_file_name = os.path.basename(args.input_file)
		output_file_name, output_ext = os.path.splitext(output_file_name)
		output_file_name += "_generated"
		output_file_name += output_ext
		args.output_file = os.path.join(output_file_path, output_file_name)

	if os.path.exists(args.output_file):
		os.remove(args.output_file)

	return_code = parse(args.input_file, args.source_file, args.config_file, args.output_file)
	exit(return_code)
