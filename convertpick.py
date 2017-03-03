import xlrd
import tkinter as tk
from tkinter import messagebox


class PickUpList:
    def __init__(self, route, sheet, row):
        self.route = route
        self.list = []
        self.row = row + 1
        # Do not use '== route' or similar, because the route column is
        # inconsistently also used for dates.
        while (sheet.cell(self.row, 0).value != '' and
               sheet.cell(self.row, 1).value != ''):
            # Values should be in these rows.
            self.add_pickup(sheet.cell(self.row, 1).value,
                            sheet.cell(self.row, 2).value,
                            sheet.cell(self.row, 4).value,
                            sheet.cell(self.row, 5).value,
                            sheet.cell(self.row, 7).value)
            self.row += 1

    def add_pickup(self, customer, action, contact, inv, description):
        new_description = description.split(',')
        new_description[:] = ["    " + x.lstrip() for x in new_description]
        new_description = "\r\n".join(new_description)
        self.list.append([customer, action, contact, inv, new_description])


def getlists(filename):
    try:
        book = xlrd.open_workbook(filename)
    except FileNotFoundError:
        error = "File was not found. Have you selected an Excel file?"
        tk.messagebox.showerror("Pick Ups To Notes", error)
    except xlrd.biffh.XLRDError:
        error = "This is not an Excel file. Please choose a different file."
        tk.messagebox.showerror("Pick Ups To Notes", error)
    else:
        sheet = book.sheet_by_index(0)
        rows, cols = sheet.nrows, sheet.ncols
        pickups = []
        for row in range(rows):
            if (sheet.cell(row, 0).value != '' and sheet.cell(row, 1).value == ''):
                pickups.append(PickUpList(sheet.cell(row, 0).value, sheet, row))
        return pickups
