from openpyxl import load_workbook
import pandas


def main():
    r = Roster("../Jones_2019.xlsx")
    r.get_student_names()


class Roster(object):
    """A roster editor/ viewer for poor Mrs.Jones"""

    def get_student_names(self):
        """Gets a complete list of all student's names"""
        student_names = []
        student_workbook = load_workbook(self.filename)
        sheet = student_workbook.active
        class_dataframe = pandas.DataFrame(sheet.values)

        for row in sheet.iter_rows(max_col=1, values_only=True):
            for cell in row:
                if isinstance(cell, str) is False:  # skip the 'ID' entry
                    student_names.append(class_dataframe.loc[cell, 1] + ' ' + class_dataframe.loc[cell, 2])

        return student_names

    def get_student(self, student_identifier):
        pandas.read_excel(self.filename)

    def delete_student(self, student_identifier):
        pandas.read_excel(self.filename)

    def __init__(self, filename):
        self.filename = filename
        pass


if __name__ == "__main__":
    main()
