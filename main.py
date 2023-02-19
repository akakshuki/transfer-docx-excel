import argparse
import re
import sys

import pandas as pd
import docx
from openpyxl import Workbook


# Class object question
class Question:
    def __init__(self, statement, level, option_a, option_b, option_c, option_d):
        self.statement = statement
        self.level = level
        self.option_a = option_a
        self.option_b = option_b
        self.option_c = option_c
        self.option_d = option_d

LEVEL_MAP = {
    'TH': 'Thông hiểu',
    'VD': 'Vận Dụng',
    'VDT': 'Vận Dụng',
    'NB': 'Nhận biết',
}

def convert_file_word_to_excel(input_file_path, output_file_path):
    if input_file_path == "":
        sys.exit(0)
    # Read the contents of the Microsoft Word file
    doc = docx.Document(input_file_path)
    # Store the contents of the file in a list
    data = []
    questions = []
    questions_for_qc = []

    for para in doc.paragraphs:
        data.append(para.text)

    for i, element in enumerate(data):
        if "Câu " in element:
            statement = element
            keyword = re.findall(r'\((.*?)\)', element)
            option_a = data[i+1]
            option_b = data[i+2]
            option_c = data[i+3]
            option_d = data[i+4]
            keyword = re.findall(r'\((.*?)\)', element)[0]
            level = LEVEL_MAP.get(keyword, 'Vận dụng cao')
            questions_for_qc.append(
                Question(statement, level, option_a, option_b, option_c, option_d)
            )
                
    for i, element in enumerate(data):
        if "Câu " in element:
            statement = re.sub(r"Câu\s*\d*\s*\([^)]*\)", "", element)
            keyword = re.findall(r'\((.*?)\)', element)
            option_a = data[i+1].replace("A.", "")
            option_b = data[i+2].replace("B.", "")
            option_c = data[i+3].replace("C.", "")
            option_d = data[i+4].replace("D.", "")
            keyword = re.findall(r'\((.*?)\)', element)[0]
            level = LEVEL_MAP.get(keyword, 'Vận dụng cao')
            questions.append(Question(statement, level, option_a, option_b, option_c, option_d))

    # Convert data list to pandas data frame
    columns=['statement', 'level', 'option_a', 'option_b', 'option_c', 'option_d']
    rows = []
    for question in questions:
        row = [question.statement, question.level, question.option_a,
               question.option_b, question.option_c, question.option_d]
        rows.append(row)
    df_final = pd.DataFrame(rows, columns=columns)

    rows_for_qc = []
    for question in questions_for_qc:
        row = [question.statement, question.level, question.option_a,
               question.option_b, question.option_c, question.option_d]
        rows_for_qc.append(row)
    df_qc = pd.DataFrame(rows_for_qc, columns=columns)

    # Create an Excel Workbook and add the DataFrame as a worksheet
    book = Workbook()
    writer = pd.ExcelWriter(
        './output/output_data.xlsx' if output_file_path == "" else output_file_path, engine='openpyxl')
    writer.book = book
    df_final.to_excel(writer, index=False, header=False, sheet_name='questions')
    df_qc.to_excel(writer, index=False, header=False, sheet_name='question_for_qc')
    # Save the Excel Workbook
    writer.save()

# Remove empty file input
# ==========================

def conver_txt_file(input_file_path, output_file_path):
    # Open the input file
    if input_file_path == "":
        sys.exit(0)
    with open(input, "r", encoding="utf-8") as input_file:
        # Read the contents of the file
        lines = input_file.readlines()
    # Remove empty lines from the list of lines
    lines = list(filter(lambda x: x.strip() != "", lines))

    # Open the output file and write the filtered lines to it
    with open('./output/output_data.txt' if output_file_path == "" else output_file_path, "w", encoding='utf-8') as output_file:
        output_file.writelines(lines)

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
