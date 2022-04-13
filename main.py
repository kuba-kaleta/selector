from openpyxl import load_workbook
import easygui
# print(easygui.fileopenbox())

tabele_dir = "D:\\jakub.kaleta\\automat\\iveco\\excel\\tabele\\"
tabela = "D:\\jakub.kaleta\\automat\\iveco\\excel"

workbook = load_workbook(filename=tabele_dir+"tabela_iveco_7774.xlsx")
sheet = workbook.active
print(sheet["A1"].value)

score = 0.0


def compare_col():
    return 0
