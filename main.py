import csv
import xlsxwriter
from os import listdir
from os.path import splitext

class csv2xlsx:
    def __init__(self, input_path, filename):
        self.input_path = input_path
        self.filename = filename

    def is_num(self, input_col):
        try:
            return float(input_col)
        except:
            return input_col

    def create_xlsx(self):
        output_path = './output/'
        workbook = xlsxwriter.Workbook(output_path + self.filename + '.xlsx')
        worksheet = workbook.add_worksheet()

        with open(self.input_path, encoding='shift-jis') as csvfile:
            reader = csv.reader(csvfile, delimiter=',')

            row_index = 0
            col_index = 0

            for row in reader:
                worksheet.write(row_index, col_index, self.is_num(row[0]))
                worksheet.write(row_index, col_index + 1, self.is_num(row[1]))
                worksheet.write(row_index, col_index + 2, self.is_num(row[2]))
                row_index += 1

            workbook.close()

if __name__ == '__main__':
    try:
        doc_path = './doc'
        filelist = listdir(doc_path)
        for file in filelist:
            filename, extension = splitext(file)
            filepath = doc_path + '/' + file
            if extension == '.csv':
                f = csv2xlsx(filepath, filename)
                f.create_xlsx()

    except FileNotFoundError:
        filelist = listdir('./')
        print(filelist)
        print("There is no doc directory under your root")
