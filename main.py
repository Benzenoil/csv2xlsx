import csv
import xlsxwriter
from os import listdir, mkdir
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
        workbook = xlsxwriter.Workbook('./output/' + self.filename + '.xlsx')
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
        if 'output' not in listdir('./'):
            mkdir('./output/')
        input_path = './doc/'
        for file in listdir(input_path):
            filename, extension = splitext(file)
            filepath = input_path + file
            if extension == '.csv':
                f = csv2xlsx(filepath, filename)
                f.create_xlsx()
    except:
        print("Unexpected error is occurred")
