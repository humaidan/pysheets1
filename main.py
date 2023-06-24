import openpyxl

# import warnings
from config import CONFIG


# warnings.simplefilter("ignore", category=UserWarning)
debug = True

print("Opening sheet ...")
wb = openpyxl.load_workbook(CONFIG["excelfile"])
sheet1 = wb[CONFIG["sheetname_source"]]
sheet2 = wb[CONFIG["sheetname_dest"]]

print("Parsing and writing data ...")
start_row = 2
# for row_no, row_val in enumerate(sheet1.iter_rows(min_row=start_row, max_row=2, values_only=True)):
for row_no, row_val in enumerate(sheet1.iter_rows(min_row=start_row, values_only=True)):
    if debug:
        print(f"- looping through record #{row_no+start_row}:")
    newsheet = wb.copy_worksheet(sheet2)

    if debug:
        print(f"\t - set title to: {row_val[CONFIG['sysname_col']] + CONFIG['sysname_col_append']}:")

    newtitle = row_val[CONFIG["sysname_col"]].rstrip()
    if newtitle in CONFIG["short_titles"]:
        newtitle = CONFIG["short_titles"][newtitle]
    newsheet.title = newtitle + CONFIG["sysname_col_append"]

    for target_cell, source_col_str in CONFIG["field_mapping"].items():
        # Read the value in "All Services" sheet in current row + config column
        source_value = ""
        if isinstance(source_col_str, str):
            source_col_val = openpyxl.utils.column_index_from_string(source_col_str)
            if debug:
                print(f"\t - sheet1.cell(row={row_no+start_row}, column={source_col_val}).value")
            source_value = sheet1.cell(row=row_no + start_row, column=source_col_val).value
            if source_value is None:
                source_value = ""

            if target_cell in CONFIG["pre_fills"] and source_value:
                source_value = CONFIG["pre_fills"][target_cell] + str(source_value)

            if debug:
                print(f"\t - reading AllSystems[row={row_no + start_row}, column={source_col_val}]  =  {source_value}.")

        else:  # list of str
            if debug:
                print(f"\t - detected a source list {source_col_str}:")

            source_value = "'"
            for col in source_col_str:
                source_col_val = openpyxl.utils.column_index_from_string(col)
                cell = sheet1.cell(row=row_no + start_row, column=source_col_val).value
                if cell is not None:
                    source_value += "- " + cell + "\n"

        # Write value in new sheet
        if debug:
            print(f"\t=Setting {newsheet.title} sheet cell [{target_cell}] to value: {source_value}.")
        if debug:
            print()
        newsheet[target_cell] = str(source_value).rstrip()


print("Cleaning ...")
# Close the workbook
# wb.close()
wb.save(CONFIG["modified_excelfile"])
wb.close()

print("Done.")
