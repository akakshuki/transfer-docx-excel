import argparse
import re
import sys

import pandas as pd
from openpyxl import Workbook

from google.colab import drive

# Class object question
class Question:
    def __init__(self, statement, level, option_a, option_b, option_c, option_d, answer):
        self.statement = statement
        self.level = level
        self.option_a = option_a
        self.option_b = option_b
        self.option_c = option_c
        self.option_d = option_d
        self.answer = answer


LEVEL_MAP = {
    'TH': '2_Thông hiểu',
    'VD': '3_Vận Dụng',
    'VDT': '3_Vận Dụng',
    'NB': '1_Nhận biết',
}

drive.mount('/content/drive/')

def convert_file_word_to_excel(input_file_path, output_file_path):
    if input_file_path == "":
        sys.exit(0)
    # Store the contents of the file in a list
    data = []
    questions = []
    questions_for_qc = []
    with open(input_file_path, "r", encoding='utf-8') as file:
        data = file.readlines()
    for i, element in enumerate(data):
        if "Câu " in element:
            statement = element[:-1]
            keyword = re.findall(r'\((.*?)\)', element)
            option_a = data[i+1][:-1]
            option_b = data[i+2][:-1]
            option_c = data[i+3][:-1]
            option_d = data[i+4][:-1]
            answer = data[i+5]
            keyword = re.findall(r'\((.*?)\)', element)[0]
            level = LEVEL_MAP.get(keyword, '4_Vận dụng cao')
            questions_for_qc.append(
                Question(statement, level, option_a,
                         option_b, option_c, option_d, answer)
            )
    for i, element in enumerate(data):
        if "Câu " in element:
            statement = re.sub(r"Câu\s*\d*\s*\([^)]*\)", "", element)[:-1]
            keyword = re.findall(r'\((.*?)\)', element)
            option_a = data[i+1].replace("A.", "")[:-1]
            option_b = data[i+2].replace("B.", "")[:-1]
            option_c = data[i+3].replace("C.", "")[:-1]
            option_d = data[i+4].replace("D.", "")[:-1]
            answer = data[i+5].split(':')[1][1:-1]
            keyword = re.findall(r'\((.*?)\)', element)[0]
            level = LEVEL_MAP.get(keyword, '4_Vận dụng cao')
            questions.append(
                Question(statement, level, option_a,
                         option_b, option_c, option_d, answer)
            )
    # Convert data list to pandas data frame
    columns = ['statement', 'level', 'option_a',
               'option_b', 'option_c', 'option_d', 'answer']
    rows = []
    for question in questions:
        row = [question.statement, question.level, question.option_a,
               question.option_b, question.option_c, question.option_d, question.answer]
        rows.append(row)
    df_final = pd.DataFrame(rows, columns=columns)
    rows_for_qc = []
    for question in questions_for_qc:
        row = [question.statement, question.level, question.option_a,
               question.option_b, question.option_c, question.option_d, question.answer]
        rows_for_qc.append(row)
    df_qc = pd.DataFrame(rows_for_qc, columns=columns)
    # Create an Excel Workbook and add the DataFrame as a worksheet
    book = Workbook()
    writer = pd.ExcelWriter(
        './output/output_data.xlsx' if output_file_path == "" else output_file_path, engine='openpyxl')
    # Remove default sheet
    del book['Sheet']
    writer.book = book
    df_final.to_excel(writer, index=False, header=False,
                      sheet_name='questions')
    df_qc.to_excel(writer, index=False, header=False,
                   sheet_name='question_for_qc')
    # Save the Excel Workbook
    writer.save()

# Remove empty file input
# ==========================


def conver_txt_file(input_file_path, output_file_path):
    # Open the input file
    if input_file_path == "":
        sys.exit(0)
    with open(file=input_file_path, mode="r", encoding="utf-8") as input_file:
        # Read the contents of the file
        lines = input_file.readlines()
    # Remove empty lines from the list of lines
    lines = list(filter(lambda x: x.strip() != "", lines))

    # Open the output file and write the filtered lines to it
    with open(file='./output/output_data.txt' if output_file_path == "" else output_file_path, mode="w", encoding='utf-8') as output_file:
        output_file.writelines(lines)
