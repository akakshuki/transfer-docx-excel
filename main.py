import argparse
from docx_functions import convert_file_word_to_excel, conver_txt_file

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Converts word files to excel or removes empty lines from a text file')
    parser.add_argument('function_name', choices=[
                        'cv', 'rm'], help='function name (cv or rm)')
    parser.add_argument('input_address', help='input file address')
    parser.add_argument('output_address', nargs='?',
                        default='', help='output file address (optional)')
    args = parser.parse_args()
    if args.function_name == 'cv':
        convert_file_word_to_excel(args.input_address, args.output_address)
    elif args.function_name == 'rm':
        conver_txt_file(args.input_address, args.output_address)
