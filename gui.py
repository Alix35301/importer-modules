# import xlrd
#
# print("FSI importer")
# print("OPTIONS")
# files = ["556FSI2","FSI_OFC","556FSID"]
# [print(i+1,j) for i,j in enumerate(files)]
#
# file_option=input("Select the file")
# if file_option ==str(1):
#     print("loading file")
#     xls = xlrd.open_workbook(r'\\admma2\data\SSD\SIS\FSI\data\556FSI2.xlsx', on_demand=True)
#     print("Select the sheet")
#     for i, j in enumerate(xls.sheet_names()):
#         print(j)
#     sheet_name=input("Enter sheet name")
#
# elif file_option==str(2):
#     print("loading file")
# elif file_option==str(3):
#     print("loading file")
# else:
#     print("Invalid Input")
#

import tkinter as tk

cls_sheet= ""
data_sheet = ""

def updatesource(cls_sheet, data_sheet):
    cls_sheet = cls_sheet
    data_sheet = data_sheet


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.one = tk.Button(self)
        self.one["text"] = "FSIU"
        self.one["command"] = self.one_action
        self.one.pack(side="top")

        self.two = tk.Button(self)
        self.two["text"] = "D_INCOME"
        self.two["command"] = self.two_action
        self.two.pack(side="top")

        self.three = tk.Button(self)
        self.three["text"] = "D_BS"
        self.three["command"] = self.three_action
        self.three.pack(side="top")

        self.three = tk.Button(self)
        self.three["text"] = "D_MI"
        self.three["command"] = self.four_action
        self.three.pack(side="top")

        self.three = tk.Button(self)
        self.three["text"] = "OFC_BS"
        self.three["command"] = self.five_action
        self.three.pack(side="top")

        self.three = tk.Button(self)
        self.three["text"] = "OFC_IND"
        self.three["command"] = self.six_action
        self.three.pack(side="top")



        self.quit = tk.Button(self, text="QUIT", fg="red",
                              command=self.master.destroy)
        self.quit.pack(side="bottom")

    def one_action(self):
        updatesource("FSIU","Table 1")


    def two_action(self):
        updatesource("D_INCOME","Annex 2")
    def three_action(self):
        updatesource("DEPOSIT","Annex 3")
    def four_action(self):
        updatesource("MEMO","Annex 4")
    def five_action(self):
        updatesource("OFC_BS","BS")
    def six_action(self):
        updatesource("OFC_IND","Indicators")


root = tk.Tk()
app = Application(master=root)
app.mainloop()



