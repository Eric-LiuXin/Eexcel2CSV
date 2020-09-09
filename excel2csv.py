import xlrd, csv, sys, os

class ExcelHelper:
    @staticmethod
    def get_column_title_list(sheet):
        col_list = list()
        for i in range(sheet.ncols):
            col_list.append(sheet.cell_value(0, i).strip())
        return col_list

    @staticmethod
    def get_row_values(sheet, row):
        value_list = list()
        row_values = sheet.row(row)
        for cell in row_values:
            if cell.ctype == xlrd.XL_CELL_NUMBER and cell.value % 1 == 0:
                value_list.append(str(int(cell.value)))
            elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                if cell.value == 0:
                    value_list.append(False)
                else:
                    value_list.append(True)
            elif cell.ctype == xlrd.XL_CELL_DATE:
                value_list.append(xlrd.xldate_as_datetime(cell.value, 0).strftime("%Y-%m-%d") )
            else:
                value_list.append(str(cell.value).strip())
        return value_list

def excel2csv(file_path, csv_path):
    book = xlrd.open_workbook(file_path)
    sheet = book.sheet_by_index(0)

    title_list = ExcelHelper.get_column_title_list(sheet)
    with open(csv_path, "w", newline="", encoding="utf-8") as csv_file:
        spam_writer = csv.writer(csv_file, dialect='excel')
        spam_writer.writerow(title_list)
        for excel_row in range(1, sheet.nrows):
            spam_writer.writerow(ExcelHelper.get_row_values(sheet, excel_row))

def main():
    if len(sys.argv) < 2:
        print("请指定要转换的Excel文件。")
        return

    file_path = sys.argv[1]
    if not os.path.exists(sys.argv[1]):
        print("指定文件不存在，或不是有效路径。")
        return

    shotname, extension = os.path.splitext(file_path)
    if extension.lower() != '.xlsx':
        print("请指定Excel类型的文件。")
        return

    csv_path = r"%s.csv" % shotname
    excel2csv(file_path, csv_path)
    print("CSV文件输出路径：%s" % csv_path)

if __name__ == '__main__':
    main()
