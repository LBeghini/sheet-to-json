import ast

import pandas
from openpyxl import load_workbook


def is_formula(value):
    value = str(value)
    if value[0] == '=':
        return True
    return False


def df_to_formatted_json(data_frame, ob, sheet):
    result = []
    for idx, row in data_frame.iterrows():
        parsed_row = {}
        for col_label, v in row.items():
            keys = col_label.split(".")

            current = parsed_row
            for i, k in enumerate(keys):
                if i == len(keys) - 1:
                    if is_formula(v):
                        for key in ob.keys():
                            if key in v:
                                position = v.partition('!')[2]
                                row = int(''.join(filter(str.isdigit, position)))
                                index = row - 2
                                if sheet not in ob[key][index]:
                                    ob[key][index][sheet] = []
                                ob[key][index][sheet].append(parsed_row)
                    else:
                        current[k] = format_value(v)
                else:
                    if k not in current.keys():
                        current[k] = {}
                    current = current[k]
        result.append(parsed_row)
    return result


def format_value(value):
    value = str(value)
    if value[0] in '[{' and value[-1] in ']}':
        return ast.literal_eval(value)
    return value


def excel_column_number(name):
    n = 0
    for c in name:
        n = n * 26 + ord(c) - ord('A')
    return n


def main():
    wb = load_workbook('x.xlsx')
    data = {}
    for sheet in wb.worksheets:
        df = pandas.DataFrame(sheet.values)
        new_header = df.iloc[0]
        df = df[1:]
        df.columns = new_header
        df = df.replace(to_replace='None').dropna()

        data[sheet.title] = df_to_formatted_json(df, data, sheet.title)

    print(data)


if __name__ == "__main__":
    main()
