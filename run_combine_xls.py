#!env/Scripts/python
import xlrd
import xlwt
import argparse
import os
import time
import sys

EXTENSION = ".xls"

def generate_output_filename():
    timestr = time.strftime("%Y%m%d_%H%M%S")
    return timestr + EXTENSION

def main():
    parser = argparse.ArgumentParser(description='Combine .xls files with same header')
    parser.add_argument(dest="search_path", help="input search path where all files are located", metavar="PATH", type=lambda x: is_valid_path(parser, x))
    parser.add_argument("-o", "--output", dest="output_filename", help="out filename where all data will be created", default=generate_output_filename())
    args = parser.parse_args()
    PATH = args.search_path
    OUTPUT_FILENAME = args.output_filename
    print("Searching for all {} files in {}".format(EXTENSION, PATH))
    list_filepath = find_all_files(PATH, EXTENSION)
    # Verify files
    list_filepath_verified = []
    list_expected_header = None
    expected_sheet = None
    count_skipped = 0
    for filepath in list_filepath:
        bool_check, error_message, list_expected_header, expected_sheet = check_alyson_xls(filepath, list_expected_header, expected_sheet)
        if not bool_check:
            print("Skipped {} ({})".format(filepath, error_message))
            count_skipped += 1
        else:
            print("Found {}".format(filepath))
            list_filepath_verified.append(filepath)
    # Read files
    list_all_file_row = []
    for filepath in list_filepath_verified:
        list_per_file_row = read_alyson_xls(filepath)
        print("Read {} - {} data".format(filepath, len(list_per_file_row)))
        list_all_file_row.extend(list_per_file_row)
    # Write data to Main
    if os.path.exists(OUTPUT_FILENAME):
        print("Filename already exists, overriding {}".format(OUTPUT_FILENAME))
    write_alyson_xls(OUTPUT_FILENAME, list_expected_header, expected_sheet, list_all_file_row)
    count_file = len(list_filepath_verified)
    count_data = len(list_all_file_row)
    count_total_file = len(list_filepath)
    print("Expected header: {}".format(list_expected_header))
    print("Expected sheet name: {}".format(expected_sheet))
    print("Data: {}".format(count_data))
    print("Files: {}/{} with {} skipped".format(count_file, count_total_file, count_skipped))
    print("Generated file: {}".format(OUTPUT_FILENAME))
    print("Complete")
    if count_skipped > 0:
        sys.exit(1)


def is_valid_path(parser, arg):
    """
    Check if user input path exists
    """
    path = os.path.realpath(arg)
    if not os.path.exists(path):
        parser.error("The path {} does not exist!".format(path))
    else:
        return path

def find_all_files(path, extension):
    """
    Example of extension is .xls, .txt, .csv
    """
    list_filepath = []
    for root, __, files in os.walk(path):
        for file in files:
            if file.endswith(extension):
                list_filepath.append(os.path.join(root, file))
    return list_filepath

def check_alyson_xls(filepath, list_expected_header, expected_sheet):
    """
    Checks if 1st row is the same
    """
    error_message = ""
    try:
        loc = (filepath) 
        wb = xlrd.open_workbook(loc)
        if expected_sheet is None:
            list_sheet_names = wb.sheet_names()
            expected_sheet = list_sheet_names[0]
    except:
        error_message = "Cannot open file"
        return False, error_message, list_expected_header, expected_sheet

    try:
        sheet = wb.sheet_by_name(expected_sheet)
        sheet.cell_value(0, 0)
    except:
        error_message = "Cannot find sheet named: {}".format(expected_sheet)
        return False, error_message, list_expected_header, expected_sheet

    list_header = sheet.row_values(0)
    if list_expected_header is not None:
        if len(list_header) == len(list_expected_header):
            count = 0
            for header, EXPECTED_HEADER in zip(list_header, list_expected_header):
                count += 1
                # Check if each header value matches
                if header != EXPECTED_HEADER:
                    error_message = "1st Row, {} Column is not expected, expectation:{} vs reality:{}".format(count, EXPECTED_HEADER, header)
                    return False, error_message, list_expected_header, expected_sheet
            return True, None, list_expected_header, expected_sheet
        else:
            error_message = "1st Row is not expected, expectation:{} vs reality:{}".format(str(list_expected_header), str(list_header))
            return False, error_message, list_expected_header, expected_sheet
    else:
        list_expected_header = list_header
        # Always Return True for first file to obtain the list_expected_header
        return True, None, list_expected_header, expected_sheet

def read_alyson_xls(filepath):
    """
    Returns data from one xls
    """
    loc = (filepath) 
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    list_per_file_row = []
    for row_idx in range(1, sheet.nrows):
        list_per_row_value = sheet.row_values(row_idx)
        list_per_file_row.append(list_per_row_value)
    return list_per_file_row

def write_alyson_xls(output_filename, list_expected_header, expected_sheet, list_all_file_row):
    wb = xlwt.Workbook()
    ws = wb.add_sheet(expected_sheet)
    index_column = 0
    index_row = 0
    for header in list_expected_header:
        ws.write(index_row, index_column, header)
        index_column += 1

    index_row = 1
    for list_per_row_value in list_all_file_row:
        for index_column in range(0, len(list_per_row_value)):
            ws.write(index_row, index_column, list_per_row_value[index_column])
        index_row += 1
    wb.save(output_filename)

if __name__ == "__main__":
    main()
