import csv
import os
import re

import click
from tqdm import tqdm
import xlrd


WZ_XLS_FOLDER = ""


def get_xls_book(xls_file_name: str) -> xlrd.book.Book:
    return xlrd.open_workbook(xls_file_name)


def get_wz_sheet_list(xls_book: xlrd.book.Book) -> list:
    wz_list = xls_book.sheet_names()
    return [item for item in wz_list if re.fullmatch(r"\d\d\d_\d\d_22_6", item) is not None]


def get_wz_sheet(xls_book: xlrd.book.Book, sheet_name: str) -> xlrd.sheet.Sheet:
    return xls_book.sheet_by_name(sheet_name)


def get_wz_number(row_items: list) -> str:
    if re.fullmatch(r"\d\d\d/[0-1][0-9]/2[0-9]/6", row_items[-2]) is not None:
        return row_items[-2]
    else:
        return ""


def compose_wz_item(assortment_dict: dict, item_name: str, item_count: int) -> dict:
    wz_item = {}
    wz_item["index"] = int(assortment_dict.get(item_name, 0))
    wz_item["name"] = item_name
    wz_item["quantity"] = item_count
    return wz_item


def get_module_by_wz(wz_sheet: xlrd.sheet.Sheet, module_name: str, wz_type: str="SPRZEDAŻ", start_row=6, start_col=4, end_col=16) -> dict:
    result = {}
    if re.search(wz_type, wz_sheet.cell_value(15,14)) is None:
        return result
    if wz_sheet.cell_value(0,0) == "WZv1":
        sr = start_row + 13
    else:
        sr = start_row + 15    
    row = wz_sheet.row_values(sr,start_col, end_col)
    while type(row[0]) == float:
        if re.search(module_name, row[1]) is not None:
            result["module_name"] = row[1]
            result["quantity"] = int(row[-1])
            number_list = [item.split('.')[0].strip() for item in str(row[7]).split(',')]
            result["module_numbers"] = number_list
            # for n_str in number_list:
            #     try:
            #          n = int(float(n_str))       
            #     except ValueError:
            #         n = '-'
            #     result["module_numbers"].append(n)
            row = wz_sheet.row_values(start_row, start_col, end_col)
            result["wz_number"] = get_wz_number(row)
            row = wz_sheet.row_values(start_row+3, start_col, end_col)
            result["disposition_number"] = str(row[-3] + " " + row[-2])
            result["type"] = wz_sheet.cell_value(15,14)
            break
        sr +=1
        row = wz_sheet.row_values(sr,start_col, end_col)
    return result


@click.command()
@click.argument("wz-file")
@click.argument("module-name", nargs=1)
@click.argument("wz-type", nargs=1)
def main(module_name: str, wz_type: str, wz_file: str):
    wz_book = get_xls_book(os.path.join(WZ_XLS_FOLDER, wz_file))
    wz_list = get_wz_sheet_list(wz_book)
    mod_list = []
    for wz_name in wz_list:
        wz_sheet = get_wz_sheet(wz_book, wz_name)
        mod = get_module_by_wz(wz_sheet,module_name,wz_type)
        if len(mod) > 0:
            mod_list.append(mod)
    print("NAZWA MODUŁU                 NR WZ     NR DYSPOZYCJI     TYP     ILOŚĆ    NUMERY")
    print("========================= =========== =============== ========= ======= =============")
    for m in mod_list:
        print(f"{m['module_name'][:25]:25} {m['wz_number']:11} {m['disposition_number']:15} {m['type'][:9]:9} {m['quantity']:7} \
            {m['module_numbers']}")
    fieldnames = mod_list[0].keys()
    csv_filename = os.path.join(".", module_name+"_"+wz_type+".csv")
    with open(csv_filename, "w", encoding="UTF8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(mod_list)
    
    
    
    # assortment_dict = get_assortment(assortment_book)
    # print("done.")
    # print("Reading DWS list...", end="")
    # dws_book = get_xls_book(os.path.join(DWS_XLS_FOLDER, "DWSygn v1 2022.xls"))
    # dws_list = get_dws_sheet_list(dws_book)
    # print("done.")
    # t = tqdm(total=len(dws_list), unit=" DWS", desc="Extracting WZ")
    # for dws_name in dws_list:
    #     dws_sheet = get_dws_sheet(dws_book, dws_name)
    #     wz_list = [get_wz_content(assortment_dict, dws_sheet, start_row=sr) for sr in get_wz_start(dws_sheet)]
    #     for wz in wz_list:
    #         save_json(wz, get_json_file_name(WZ_JSON_FOLDER,"WZ_"+wz["WZ_number"]))
    #     t.update(n=1)
    # t.close()


if __name__ == "__main__":
    main()
    