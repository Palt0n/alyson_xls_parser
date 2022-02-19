#!env/Scripts/python
import xlrd, xlwt
import argparse
import os

EXTENSION = ".xls"

LIST_EXPECTED_HEADER = ['Type', 'Name', 'ID', 'Point X', 'Point Y', 'Length (µm)', 'Angle', 'Area', 'Perimeter']

def main():
    parser = argparse.ArgumentParser(description='Combine .xls files with same header')
    parser.add_argument(dest="path", help="input path where all files are located", metavar="PATH", type=lambda x: is_valid_path(parser, x))
    args = parser.parse_args()
    PATH = args.path
    print("Searching for all {} files in {}".format(EXTENSION, PATH))
    list_filepath = find_all_files(PATH, EXTENSION)
    # Verify files
    list_filepath_verified = []
    list_expected_header = None
    count_skipped = 0
    for filepath in list_filepath:
        bool_check, error_message, list_expected_header = check_alyson_xls(filepath, list_expected_header)
        if not bool_check:
            print("Skipped {} ({})".format(filepath, error_message))
            count_skipped += 1
        else:
            print("Found {}".format(filepath))
            list_filepath_verified.append(filepath)
    # Read files and Dump into Main
    for filepath in list_filepath_verified:
        pass
        
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

def check_alyson_xls(filepath, list_expected_header):
    """
    Checks if 1st row is the same
    """
    error_message = ""
    try:
        loc = (filepath) 
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)
        sheet.cell_value(0, 0)
    except:
        error_message = "Cannot open file"
        return False, error_message, None
    list_header = sheet.row_values(0)
    if list_expected_header is not None:
        if len(list_header) == len(list_expected_header):
            count = 0
            for header, EXPECTED_HEADER in zip(list_header, list_expected_header):
                count += 1
                # Check if each header value matches
                if header != EXPECTED_HEADER:
                    error_message = "1st Row, {} Column is not expected, expectation:{} vs reality:{}".format(count, EXPECTED_HEADER, header)
                    return False, error_message, None
            return True, None, list_expected_header
        else:
            error_message = "1st Row is not expected, expectation:{} vs reality:{}".format(str(list_expected_header), str(list_header))
            return False, error_message, None
    else:
        # Always Return True for first file to obtain the list_expected_header
        return True, None, list_expected_header



def read_alyson_xls(filepath):
    """
    Returns data from one xls
    """
    loc = (filepath) 
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    print(sheet.row_values(1))



if __name__ == "__main__":
    main()
