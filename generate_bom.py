import sys
import pandas as pd
from collections import deque
import re
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


if __name__ == "__main__":

    file_location = '/home/a_noob__/Downloads/BOM file for Data processing.xlsx'

    if len(sys.argv) > 1:
        path = sys.argv[1]
        file_location = path

    df = pd.read_excel(file_location)

    output = get_bom_dictionary(df)

    for key in output:
        print(key)
