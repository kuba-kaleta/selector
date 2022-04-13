from openpyxl import load_workbook
import easygui
# print(easygui.fileopenbox())

workbook = load_workbook(filename="D:\\jakub.kaleta\\automat\\iveco\\excel\\tabela_iveco_matryca.xlsx", data_only=True)
sheet = workbook.active

table_dir = "D:\\jakub.kaleta\\automat\\iveco\\excel\\tabele\\"
workbook_baza = load_workbook(filename="D:\\jakub.kaleta\\automat\\iveco\\excel\\tabele\\tabela_iveco_7774.xlsx", data_only=True)
sheet_baza = workbook_baza.active

co1 = 0.7
co2 = 0.2
co3 = 0.1


def compare(cur_sheet):
    scr = 0.0
    scr_x = 0.0

    for i in range(2,82):
        cur1 = float(cur_sheet["A" + str(i)].value)
        baza1 = float(sheet["A" + str(i)].value)
        if abs(cur1 - baza1) < 1:
            scr_x += 2.0
        else:
            scr_x += 1.0

    scr = co1*scr_x + scr
    return scr


score = 0.0

score += compare(sheet)

print(score)


