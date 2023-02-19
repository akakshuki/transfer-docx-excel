
# Word-to-Excel and Empty-Line-Remover
## This script contains two functions:

1. `convert_file_word_to_excel(input, output)`: 
This function converts a Microsoft Word file to an Excel file. The Word file should contain questions with options, and each question should start with the text "Câu" (meaning "Question" in Vietnamese) followed by the question number. The questions should be followed by four options starting with A., B., C., and D. The options should be enclosed in parentheses with the first letter of the option as the identifier.

2. `conver_txt_file(input, output)`: This function removes empty lines from a text file. The input and output file names should be provided as arguments. If no output file name is given, the function creates a file named "output_data.txt" in the "./output" directory.

## Requires

Install libs with command:

```bash
    pip install -r ./requirements.txt
```

To run the script, run the main.py file with the following command:

```bash
python main.py [function_name] [input_address] [output_address]
```
function_name can be either "cv" or "rm" to call either the convert_file_word_to_excel or conver_txt_file function, respectively. input_address should be the address of the input file, and output_address should be the optional address of the output file.

You can use the -h option to see the help message and the available arguments:

```bash
python main.py -h
```