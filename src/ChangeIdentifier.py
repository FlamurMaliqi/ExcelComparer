from ExcelReader import ExcelReader
from ExcelWriter import ExcelWriter
from ExcelFile import ExcelFile
import tkinter
import os
from tkinter import filedialog


class ChangeIdentifier:
    def __init__(self, old: ExcelFile, check: ExcelFile):
        self.old_file = old
        self.check_file = check
        self.max_row = 0
        self.max_column = 0

    def find_heading_difference(self):
        count_head = 1
        for column in range(1, max(self.old_file.max_row, self.check_file.max_row) + 1):
            for row in range(3, max(self.old_file.max_column, self.check_file.max_column) + 1):
                old_cell_value = str(self.old_file.worksheet.cell(column, row).value).lower().strip()
                check_cell_value = str(self.check_file.worksheet.cell(column, row).value).lower().strip()
                if old_cell_value == "heading":
                    self.old_file.headings += 1
                    self.old_file.heading_line_rows.append((f"Heading{count_head}", column))
                if check_cell_value == "heading":
                    self.check_file.headings += 1
                    self.check_file.heading_line_rows.append((f"Heading{count_head}", column))
                    count_head += 1
        return self.old_file.headings, self.check_file.headings

    def get_lines_of_headings(self):
        self._equal_len_lists()
        for i in range(max(self.check_file.headings, self.old_file.headings)):
            diff_old = abs(self.old_file.heading_line_rows[i][1] - self.old_file.heading_line_rows[i + 1][1])
            self.old_file.lines_below_headings.append((self.old_file.heading_line_rows[i][0], diff_old))
            diff_check = abs(self.check_file.heading_line_rows[i][1] - self.check_file.heading_line_rows[i + 1][1])
            self.check_file.lines_below_headings.append((self.check_file.heading_line_rows[i][0], diff_check))

    def _equal_len_lists(self):
        difference = abs(self.check_file.headings - self.old_file.headings)
        for i in range(difference):
            self.old_file.heading_line_rows.append((f"Heading{self.check_file.headings}", self.old_file.max_row))
        self.old_file.heading_line_rows.append((f"Heading{self.check_file.headings}", self.old_file.max_row))
        self.check_file.heading_line_rows.append((f"Heading{self.check_file.headings}", self.check_file.max_row))

    # OID steht für ObjectIdentifier
    def get_object_ids(self, file: ExcelFile):
        object_ids = []
        for col in range(1, file.max_column + 1):
            cell = file.worksheet.cell(column=col, row=1)
            if cell.value == "Object Identifier":
                row = cell.row
                while True:
                    row += 1
                    next_cell = file.worksheet.cell(column=col, row=row)
                    if next_cell.value is None:
                        break
                    object_ids.append(next_cell.value)
                break
        return object_ids or None

    # Sucht Spalte der ObjectIDs raus.
    def get_id_column(self, file, text):
        cell_column = 0
        for col in range(1, file.max_column + 1):
            cell = file.worksheet.cell(column=col, row=1)
            if cell.value == text:
                cell_column = cell.column
        return cell_column

    def retrieve_object_ids_of_header(self, file):
        entries = {}
        cell_column = self.get_id_column(file, "Object Identifier")
        for i in range(file.headings):
            lst = []
            for j in range(file.lines_below_headings[i][1] + 1):
                lst.append([j+file.heading_line_rows[i][1], file.worksheet.cell(column=cell_column, row=j+file.heading_line_rows[i][1]).value])
            header = lst[0][1]
            entries.update({header: lst})
        return entries

    def compare_dicts(self, check_dict, old_dict):
        col_hid = self.get_id_column(self.check_file, "Type of Object")

        # Finden neuer Headings
        value_check = [check_dict[key] for key in check_dict]
        value_old = [old_dict[key] for key in old_dict]
        first_check = [value_check[i][0][1] for i in range(len(check_dict))]
        first_old = [value_old[i][0][1] for i in range(len(old_dict))]
        first = [first for first in first_check if first not in first_old]
        new_headings = [y for x in value_check for y in x if y[1] in first]
        for nh in new_headings:
            if nh[1] is None and self.check_file.worksheet.cell(column=col_hid, row=nh[0]).value != "Heading":
                new_headings.remove(nh)

        # Finden der neu eingefügten TestSpecs (neue Zeilen)
        check = [y[1] for i in value_check for y in i]
        old = [y[1] for i in value_old for y in i]
        lst = [c for c in check if c not in old]
        new_lines = [y for x in value_check for y in x if y[1] in lst]
        for nl in new_lines:
            if nl[1] is None and self.check_file.worksheet.cell(column=col_hid, row=nl[0]).value == "Heading":
                new_lines.remove(nl)

        # Finden der verschobenen TestSpecs
        check = [[key, v[1]] for key in check_dict for v in check_dict[key]]
        old = [[key, v[1]] for key in old_dict for v in old_dict[key]]
        check = [i for i in check if i not in old]
        check = [i for i in check if i[1] not in lst]
        check = [i[1] for i in check if i[1] not in first]
        moved_lines = [y for x in value_check for y in x if y[1] in check]

        # Finden der gelöschten TestSpecs
        check = [y[1] for i in value_check for y in i]
        old = [y[1] for i in value_old for y in i]
        old = [o for o in old if o not in check]
        old = [y for x in value_old for y in x if y[1] in old]
        deleted_lines = []
        for d in old:
            for key in check_dict.keys():  # Alle Heading O_Ids von check
                value = old_dict.get(key)  # Liste aller O_Ids in old (Liste aller Testcases zu einem bestimmten Heading)
                if value is not None:      # None Werte von old nicht zulassen (Neue Headings von check)
                    for entry in value:    # Alle durchgehen und mit der gelöschten Zeile abgleichen (anhand der O_Id) was anderes gits hier nicht.
                        if d[1] == entry[1]:  # Wenn Id gleich ist, dann füge erste Zeilen Id (Heading) und O_Id-Testcase in Liste hinzu.
                            old_to_check = check_dict.get(key)
                            deleted_lines.append([old_to_check[0][0], d[1]])

        return new_headings, new_lines, moved_lines, deleted_lines

    def get_testcases(self, file: ExcelFile):
        id = self.get_id_column(file, "Type of Object")
        entries = []
        for i in range(1, file.max_row + 1):
            entry = file.worksheet.cell(column=id, row=i).value
            if entry == 'Testcase':
                entries.append(i)
        return entries

    def generate_tc_id(self, oid: str):
        c = oid.split('_')
        result = "TC_ID_" + c[0] + "_" + c[1]
        return result

def main():
    try:
        root = tkinter.Tk()
        root.withdraw()
        old_path = filedialog.askopenfilename(title="Select the DOORS Excel file", filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if old_path:
            print("Ausgewählte Datei:", old_path)
        else:
            print("Keine Datei ausgewählt")

        check_path = filedialog.askopenfilename(title="Select the ready to Check Excel file", filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if check_path:
            print("Ausgewählte Datei:", check_path)
        else:
            print("Keine Datei ausgewählt")

        # old_path = "ExcelData/old/old.xlsx"
        # check_path = "ExcelData/check/check.xlsx"

        o = ExcelReader(old_path)
        c = ExcelReader(check_path)
        old = o.init_excel_file()
        check = c.init_excel_file()

        change = ChangeIdentifier(old, check)
        change.find_heading_difference()
        change.get_lines_of_headings()
        change.get_object_ids(old)

        print(f"Headings in Old: {old.headings}")
        print(f"Old Heading Zeilen: {old.heading_line_rows}")
        print(f"Zeilen zu old Headings: {old.lines_below_headings}")

        print(f"Headings in Check: {check.headings}")
        print(f"Check Heading Zeilen: {check.heading_line_rows}")
        print(f"Zeilen zu check Headings: {check.lines_below_headings}")
        check_dict = change.retrieve_object_ids_of_header(check)
        old_dict = change.retrieve_object_ids_of_header(old)
        new_headings, new_lines, moved_lines, deleted_lines = change.compare_dicts(check_dict=check_dict, old_dict=old_dict)
        print(new_headings)
        print(new_lines)
        print(moved_lines)
        print(deleted_lines)

        writer = ExcelWriter(old, check)
        darkblue = 'ff718cff'
        lightblue = 'ff72b5fe'
        yellow = 'ffffff57'
        red = 'ffffafaf'
        writer.add_test_case_id(change.get_testcases(check), change.get_id_column(check, "TC ID"),change.generate_tc_id(list(check_dict.values())[0][0][1]))
        writer.list_colouring(new_headings, darkblue)
        writer.list_colouring(new_lines, lightblue)
        writer.list_colouring(moved_lines, yellow)
        writer.add_new_row(deleted_lines, change.get_id_column(check, "Object Identifier"), red)
        writer.delete_column(change.get_id_column(check, "Result"))
        writer.delete_column(change.get_id_column(check, "Comment"))
        writer.delete_column(change.get_id_column(check, "Software"))
        writer.delete_column(change.get_id_column(check, "Änderungsinfo"))
        writer.create_result(os.path.basename(check_path))
        input("Press enter to exit")
        o.close_excel_file()
        c.close_excel_file()
    except Exception as e:
        print("An Exception occurred:")
        print(e)
    finally:
        o.close_excel_file()
        c.close_excel_file()


if __name__ == '__main__':
    main()
