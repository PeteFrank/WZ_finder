import csv
import os
import re

import click
import xlrd


WZ_XLS_FOLDER = ""


def get_xls_book(xls_file_name: str) -> xlrd.book.Book:
    return xlrd.open_workbook(xls_file_name)


def get_wz_sheet_list(xls_book: xlrd.book.Book) -> list:
    wz_list = xls_book.sheet_names()
    return [item for item in wz_list if re.fullmatch(r"[0-9]{1,3}_[0-9]{1,2}_[0-9]{1,2}_6", item) is not None]


def get_wz_sheet(xls_book: xlrd.book.Book, sheet_name: str) -> xlrd.sheet.Sheet:
    return xls_book.sheet_by_name(sheet_name)


def get_wz_number(row_items: list) -> str:
    if re.fullmatch(r"\d\d\d/[0-1][0-9]/[0-9]{1,2}/6", row_items[-2]) is not None:
        return row_items[-2]
    else:
        return ""


def get_first_row(column: list) -> int:
    i=0
    while column[i]!="Lp":
        i += 1
    return i+1


def check_wz_type(wz_type: str, tested_string: str) -> bool:
    return re.search(wz_type, tested_string, re.IGNORECASE + re.MULTILINE) is None


def get_module_by_wz(wz_sheet: xlrd.sheet.Sheet, module_name: str, wz_type: str="SPRZEDAŻ", start_row=6, start_col=4, end_col=16) -> list[dict]:
    result = []
    module_item = {}
    if check_wz_type(wz_type, wz_sheet.cell_value(15,14)):
        return result
    sr = get_first_row(wz_sheet.col_values(start_col))
    row = wz_sheet.row_values(sr,start_col, end_col)
    while type(row[0]) == float:
        if re.search(module_name, row[1], re.IGNORECASE + re.MULTILINE) is not None:
            module_item["module_name"] = row[1]
            try:
                module_item["quantity"] = int(row[-1])
            except ValueError:
                module_item["quantity"] = int(0)
            number_list = [item.split('.')[0].strip() for item in str(row[7]).split(',')]
            module_item["module_numbers"] = number_list
            row = wz_sheet.row_values(start_row, start_col, end_col)
            module_item["wz_number"] = get_wz_number(row)
            row = wz_sheet.row_values(start_row+3, start_col, end_col)
            module_item["disposition_number"] = str(row[-3] + " " + row[-2])
            module_item["type"] = wz_sheet.cell_value(15,14)
            result.append(module_item.copy())
        sr += 1
        row = wz_sheet.row_values(sr,start_col, end_col)
    return result


@click.command()
@click.argument("wz-file")
@click.argument("module-name", nargs=1)
@click.argument("wz-type", nargs=1)
def main(module_name: str, wz_type: str, wz_file: str):
    """
    Program generuje listę modułów z pliku [wz_file] (plik wz'tów *.xls) o określonej nazwie [module_name] z określoengo typu wz'tów [wz_type].
    Lista elementów zapisywana jest do pliku csv w bieżącym katalogu.
    v1.0.0
    """
    wz_book = get_xls_book(os.path.join(WZ_XLS_FOLDER, wz_file))
    wz_list = get_wz_sheet_list(wz_book)
    mod_list = []
    for wz_name in wz_list:
        wz_sheet = get_wz_sheet(wz_book, wz_name)
        mod_items = get_module_by_wz(wz_sheet,module_name,wz_type)
        if len(mod_items) > 0:
            mod_list.extend(mod_items)

    modules_sum = sum([int(q["quantity"]) for q in mod_list])
    items_sum = len(mod_list)
    print("NAZWA MODUŁU                           NR WZ     NR DYSPOZYCJI     TYP     ILOŚĆ    NUMERY")
    print("=================================== =========== =============== ========= ======= =============")
    for m in mod_list:
        print(f"{m['module_name'][:35]:35} {m['wz_number']:11} {m['disposition_number']:15} {m['type'][:9]:9} {m['quantity']:7} \
            {m['module_numbers']}")
    print(f"RAZEM POZYCJI: {items_sum}, RAZEM MODUŁÓW: {modules_sum}")
    try:
        fieldnames = mod_list[0].keys()
    except IndexError:
        print("Nie znaleziono modułów pasujących do wzorca.")
    else:
        csv_filename = os.path.join(".", module_name+"_"+wz_type+".csv")
        with open(csv_filename, "w", encoding="UTF8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(mod_list)


if __name__ == "__main__":
    main()
    