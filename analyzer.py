import os
import shutil
import tempfile
import zipfile

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.xml.constants import ARC_SHARED_STRINGS, SHEET_MAIN_NS
from openpyxl.xml.functions import Element, SubElement, tostring
from lxml.etree import ElementTree
from openpyxl import Workbook
from openpyxl.xml.constants import SHARED_STRINGS
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from lxml.etree import Element, SubElement, tostring
from openpyxl.xml.functions import tostring
from openpyxl.workbook import Workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
from openpyxl.xml.functions import Element, SubElement
from openpyxl.xml.functions import tostring
from openpyxl.xml.functions import tostring
from lxml.etree import Element, SubElement, tostring



import pandas as pd

import xlrd

from extractor import stockInput

def analyzeNetIncome(ticker_list):
    input_filepath = r'C:\Users\cummi\Desktop\webscrap\bin\net_income.xlsx'
    output_filepath = r'C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses\net_income_analysis.xlsx'

    # Create a temporary directory
    temp_dir = tempfile.mkdtemp()

    try:
        # Create a temporary file path with .xlsx extension
        temp_filepath = os.path.join(temp_dir, 'temp_file.xlsx')

        # Copy the input file to the temporary file
        shutil.copy(input_filepath, temp_filepath)

        # Load the temporary file
        wb = load_workbook(filename=temp_filepath)

        # Perform the analysis and modifications on the workbook
        ws = wb.active
        ws['F1'] = 'G1'
        ws['G1'] = 'G2'
        ws['H1'] = 'G3'
        ws['I1'] = 'AAGR'
        for i, ticker in enumerate(ticker_list, start=2):
            row_num = str(i)
            ws['F' + row_num] = '=((D' + row_num + '- E' + row_num + ')/E' + row_num + ')*100'
            ws['G' + row_num] = '=((C' + row_num + '- D' + row_num + ')/D' + row_num + ')*100'
            ws['H' + row_num] = '=((B' + row_num + '- C' + row_num + ')/C' + row_num + ')*100'
            ws['I' + row_num] = '=AVERAGE(F' + row_num + ':H' + row_num + ')'

        # Save the modified workbook without shared strings
        wb.save(output_filepath)

        # Create the sharedStrings.xml file and update the workbook archive
        create_shared_strings_xml(output_filepath, wb.shared_strings)

        print("Analysis completed successfully.")

        # Check if xl/sharedStrings.xml exists in the output Excel archive
        with zipfile.ZipFile(output_filepath, 'r') as archive:
            file_list = archive.namelist()
            if 'xl/sharedStrings.xml' in file_list:
                print("xl/sharedStrings.xml is in the output Excel archive.")
            else:
                print("xl/sharedStrings.xml is NOT in the output Excel archive.")

    except Exception as e:
        print(f"Error analyzing net income: {str(e)}")

    finally:
        # Remove the temporary directory and its contents
        shutil.rmtree(temp_dir)





def generate_shared_strings_xml(shared_strings):
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' \
          '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{0}" uniqueCount="{0}">'.format(
        len(shared_strings))
    for shared_string in shared_strings:
        xml += '<si><t>{}</t></si>'.format(shared_string)
    xml += '</sst>'
    return xml








def analyzeHistoricalPrices(ticker_list):
    xlsx_path = 'C:\\Users\\cummi\\Desktop\\webscrap\\bin\\xlsx\\'
    output_directory = 'C:\\Users\\cummi\\Desktop\\webscrap\\bin\\historical_price_analyses\\'

    for ticker in ticker_list:
        # Load the existing workbook
        wb = load_workbook(xlsx_path + ticker + '.xlsx')
        ws = wb.active

        length = ws.max_row

        # Add the analysis headers
        ws['G1'] = 'Pco'
        ws['H1'] = 'Phl'
        ws['I1'] = 'AvgPco'
        ws['J1'] = 'AvgPhl'

        for i in range(2, length + 1):
            i = str(i)
            ws['G' + i] = '=((B' + i + '-D' + i + ')/(D' + i + '))*100'  # B2-D2
            ws['H' + i] = '=((E' + i + '-F' + i + ')/(F' + i + '))*100'  # E2-F2

        ws['I2'] = '=AVERAGE(G2:G' + str(length) + ')'
        ws['J2'] = '=AVERAGE(H2:H' + str(length) + ')'

        # Save the modified workbook to the output directory
        wb.save(output_directory + ticker + '.xlsx')


def nothing(ticker_list):
    xlsx_path = 'C:\\Users\\cummi\\Desktop\\webscrap\\bin\\xlsx\\'

    for ticker in ticker_list:
        wb = None

        try:
            # Load the existing workbook using openpyxl
            wb = load_workbook(xlsx_path + ticker + '.xlsx', read_only=True, data_only=True)
        except KeyError:
            # If sharedStrings.xml is missing, try opening the archive and modifying the case of the filename
            archive = zipfile.ZipFile(xlsx_path + ticker + '.xlsx')
            shared_strings_path = None

            for file in archive.namelist():
                if file.lower() == 'xl/sharedstrings.xml':
                    shared_strings_path = file
                    break

            if shared_strings_path:
                archive.extract(shared_strings_path, xlsx_path)
                shared_strings_path = xlsx_path + shared_strings_path

                # Modify the case of the extracted sharedStrings.xml filename
                renamed_path = shared_strings_path.replace('xl/sharedstrings.xml', 'xl/SharedStrings.xml')
                os.rename(shared_strings_path, renamed_path)

                # Load the workbook with the modified sharedStrings.xml filename
                wb = load_workbook(xlsx_path + ticker + '.xlsx', read_only=True, data_only=True)

        if wb:
            ws = wb.active

            length = ws.max_row

            # Add the analysis headers
            ws['G1'] = 'Pco'
            ws['H1'] = 'Phl'
            ws['I1'] = 'AvgPco'
            ws['J1'] = 'AvgPhl'

            for i in range(2, length + 1):
                i = str(i)
                ws['G' + i] = '=B' + i + '-D' + i  # B2-D2
                ws['H' + i] = '=E' + i + '-F' + i  # E2-F2

            ws['I2'] = '=AVERAGE(G2:G' + str(length) + ')'
            ws['J2'] = '=AVERAGE(H2:H' + str(length) + ')'

            wb.save(r'C:\Users\cummi\Desktop\webscrap\bin\historical_price_analyses')






def second_old_analyzeNetIncome(ticker_list):
    input_filepath = r'C:\Users\cummi\Desktop\webscrap\bin\net_income.xlsx'
    output_filepath = r'C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses\net_income_analysis.xlsx'

    # Create a temporary directory
    temp_dir = tempfile.mkdtemp()

    try:
        # Create a temporary file path with .xlsx extension
        temp_filepath = os.path.join(temp_dir, 'temp_file.xlsx')

        # Copy the input file to the temporary file
        shutil.copy(input_filepath, temp_filepath)

        # Load the temporary file
        wb = load_workbook(filename=temp_filepath)

        # Perform the analysis and modifications on the workbook
        ws = wb.active
        ws['F1'] = 'G1'
        ws['G1'] = 'G2'
        ws['H1'] = 'G3'
        ws['I1'] = 'AAGR'
        for i in range(2, len(ticker_list) + 2):
            i = str(i)
            ws['F' + i] = '=((D' + i + '- E' + i + ')/E' + i + ')*100'
            ws['G' + i] = '=((C' + i + '- D' + i + ')/D' + i + ')*100'
            ws['H' + i] = '=((B' + i + '- C' + i + ')/C' + i + ')*100'
            ws['I' + i] = '=AVERAGE(F' + i + ':H' + i + ')'

        # Save the modified workbook
        wb.save(output_filepath)
        print("Analysis completed successfully.")

    except Exception as e:
        print(f"Error analyzing net income: {str(e)}")

    finally:
        # Remove the temporary directory and its contents
        shutil.rmtree(temp_dir)



def old_analyzeNetIncome(ticker_list):
    # Loads the net income workbook
    wb = load_workbook(filename=r'C:\Users\cummi\Desktop\webscrap\bin\net_income.xlsx')

    # Loads the first sheet of the net income workbook
    ws = wb.active

    # Assigns cells F1:I1 the values of G1, G2, G3, and AGR, which means growth in period 1, 2, 3, and average annual
    # growth rate.
    ws['F1'] = 'G1'
    ws['G1'] = 'G2'
    ws['H1'] = 'G3'
    ws['I1'] = 'AAGR'

    # Iterates over each row and writes the growth rate formulas
    for i in range(2, len(ticker_list) + 2):
        i = str(i)
        ws['F' + i] = '=((D' + i + '- E' + i + ')/E' + i + ')*100'
        ws['G' + i] = '=((C' + i + '- D' + i + ')/D' + i + ')*100'
        ws['H' + i] = '=((B' + i + '- C' + i + ')/C' + i + ')*100'
        ws['I' + i] = '=AVERAGE(F' + i + ':H' + i + ')'

    # Saves the analysis to the net_income_analyses folder
    wb.save(r'C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses\net_income_analysis.xlsx')
''' 
def create_shared_strings_xml(output_filepath, shared_strings):
    # Create a temporary directory
    temp_dir = tempfile.mkdtemp()

    try:
        # Extract the workbook archive
        with zipfile.ZipFile(output_filepath, 'r') as archive:
            archive.extractall(temp_dir)

        # Create the sharedStrings.xml file
        shared_strings_path = os.path.join(temp_dir, "xl", "sharedStrings.xml")
        shared_strings_xml = generate_shared_strings_xml(shared_strings)
        with open(shared_strings_path, "w", encoding="utf-8") as f:
            f.write(shared_strings_xml)

        # Update the workbook archive
        with zipfile.ZipFile(output_filepath, 'a') as archive:
            shared_strings_rel_path = os.path.join("xl", "sharedStrings.xml")
            archive.write(shared_strings_path, shared_strings_rel_path)

    finally:
        # Remove the temporary directory and its contents
        shutil.rmtree(temp_dir)
'''