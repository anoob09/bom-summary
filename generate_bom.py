import sys
import pandas as pd
from collections import deque
import re
import xlsxwriter
import json


# Required functions


def get_level(row):
    return int(re.search(r'\d+', row).group())


# dictionary to maintain column names
columns = {
    'item_name': 'Item Name',
    'level': 'Level',
    'raw_material': 'Raw material',
    'quantity': 'Quantity',
    'unit': 'Unit'}

raw_material_columns = {
    'item': 'Item Description',
    'quantity': 'Quantity',
    'unit': 'Unit '
}


def get_bom_dictionary(df=None):

    # create stack to maintain parent items
    stack = deque()

    # output dictionary to maintain BOMs
    output = {}

    # maitain two rows to compare levels and parent
    prev_row = None
    curr_row = None

    for index, curr_row in df.iterrows():
        if pd.notnull(curr_row[columns['item_name']]):
            current_item = curr_row[columns['item_name']].strip()
            raw_material_name = curr_row[columns['raw_material']].strip()
            quantity = curr_row[columns['quantity']]
            unit = curr_row[columns['unit']].strip()
            raw_material = {
                raw_material_columns['item']: raw_material_name,
                raw_material_columns['quantity']: quantity,
                raw_material_columns['unit']: unit,
            }

            if prev_row is None:
                parent = current_item
                stack.append(parent)
                output.setdefault(parent, []).append(raw_material)
            else:
                if current_item == prev_row[columns['item_name']].strip():

                    curr_level = get_level(curr_row[columns['level']])
                    prev_level = get_level(prev_row[columns['level']])

                    if curr_level == prev_level:
                        parent = stack[-1]
                        output.setdefault(parent, []).append(raw_material)
                    elif curr_level > prev_level:
                        parent = prev_row[columns['raw_material']].strip()
                        stack.append(parent)
                        output[parent] = []
                        output.setdefault(parent, []).append(raw_material)
                    else:
                        while (prev_level > curr_level):
                            stack.pop()
                            prev_level -= 1
                        parent = stack[-1]
                        stack.append(parent)
                        output.setdefault(parent, []).append(raw_material)
                else:
                    stack.clear()
                    parent = current_item
                    stack.append(parent)
                    output.setdefault(parent, []).append(raw_material)

            prev_row = curr_row

    return output


def generate_bom(output, workbook):

    normal_cell_format = workbook.add_format(
        {'align': 'right'})
    header_cell_format = workbook.add_format(
        {'bold': True, 'align': 'right', 'bg_color': '#7E9ED7'})
    item_cell_format = workbook.add_format(
        {'align': 'right', 'bg_color': '#E8F922'})

    for item_name in output:
        worksheet = workbook.add_worksheet(item_name)
        row = 0

        worksheet.write(row, 0, "Finished Good List", normal_cell_format)
        row += 1

        headers = ["#", "Item Description", "Quantity", "Unit"]

        # writing headers
        for col, header in enumerate(headers):
            worksheet.write(row, col, header, header_cell_format)
        row += 1

        # Item
        item = [1, item_name, 1, "Pc"]
        for col, val in enumerate(item):
            worksheet.write(row, col, val, normal_cell_format)
        row += 1

        worksheet.write(row, 0, "End of FG", normal_cell_format)
        row += 1

        worksheet.write(row, 0, "Raw Material List", normal_cell_format)
        row += 1

        # writing headers
        for col, header in enumerate(headers):
            worksheet.write(row, col, header, header_cell_format)
        row += 1

        for raw_item in output[item_name]:

            col = 0
            item_no = 1
            worksheet.write(row, col, item_no, normal_cell_format)
            col += 1

            for value in raw_item.values():
                worksheet.write(row, col, value, item_cell_format)
                col += 1

            row += 1

        worksheet.write(row, 0, "End of RM", normal_cell_format)
        row += 1


if __name__ == "__main__":

    file_location = '/home/a_noob__/Downloads/BOM file for Data processing.xlsx'

    if len(sys.argv) > 1:
        path = sys.argv[1]
        file_location = path

    df = pd.read_excel(file_location, engine='openpyxl')

    output = get_bom_dictionary(df)

    workbook = xlsxwriter.Workbook('output.xlsx')

    generate_bom(output, workbook)
    workbook.close()
