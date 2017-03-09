import pyodbc
from convertpick import *
from route_translation import route_translation
import tkinter as tk
from tkinter import messagebox
from os import path, getcwd
import sys
import re


def createpick(filename):
    config = get_config()
    connect_string = 'DRIVER={SQL Server};SERVER='+config[0]+';DATABASE='+config[1]+';UID='+config[2]+';PWD='+config[3]
    try:
        loading = tk.Tk()
        loading.title("Pick Ups To Notes")
        loading.iconbitmap(getcwd() + '\putn.ico')
        label = tk.Label(loading, text="Connecting. . .")
        label.pack(side="top", fill="both", expand=True, padx=20, pady=20)
        loading.update()
        stalmic = pyodbc.connect(connect_string)
        loading.destroy()
    except pyodbc.Error:
        loading.destroy()
        error = "Could not connect to database."
        tk.messagebox.showerror("Pick Ups To Notes", error)
    else:
        cursor = stalmic.cursor()
        stalmic.setdecoding(pyodbc.SQL_CHAR, encoding='utf-8')
        stalmic.setdecoding(pyodbc.SQL_WCHAR, encoding='utf-8')
        stalmic.setencoding('utf-8')
        picklist = getlists(filename)
        cursor.execute("SELECT RouteNum, NoteID, NoteText FROM dbo.Note INNER JOIN dbo.Route ON RecordID = RouteID WHERE ModuleCode = 'Route' AND NoteText LIKE 'PICKUPS%'")
        routes = cursor.fetchall()
        route_notes = {}
        for route in routes:
            route_notes[route[0]] = []
        for route in routes:
            route_notes[route[0]].append([route[1], route[2]])
        write_pickups(picklist, route_notes, cursor)
        stalmic.close()


def get_routeIDs(cursor):
    cursor.execute("SELECT RouteNum, RouteID FROM dbo.Route WHERE Active = 1")
    route_id_list = {}
    routes = cursor.fetchall()
    for route in routes:
        route_id_list[route[0]] = route[1]
    return route_id_list


def check_existing(picklist, route_notes):
    existing = []
    for pick in picklist:
        try:
            route_translation[pick.route]
        except KeyError:
            error = "The route "
            error += pick.route
            error += " in the pickup list is not mapped to a known bMobile route."
            error += " Pickup notes for this route will not be sent."
            error += "\n\nClick OK to continue."
            tk.messagebox.showwarning("Pick Ups To Notes", error)
        else:
            # Get pre=existing notes
            for bm_route in route_translation[pick.route]:
                if bm_route in route_notes and bm_route not in existing:
                    existing.append(bm_route)
    return existing


def write_pickups(picklist, route_notes, cursor):
    notes = {}
    for pick in picklist:
        try:
            route_translation[pick.route]
        except KeyError:
            error = "The route "
            error += pick.route
            error += " in the pickup list is not mapped to a known bMobile route."
            error += " Pickup notes for this route will not be sent."
            error += "\n\nClick OK to continue."
            tk.messagebox.showwarning("Pick Ups To Notes", error)
            picklist.remove(pick)
        else:
            for bm_route in route_translation[pick.route]:
                if bm_route not in notes and pick.list != []:
                    notes[bm_route] = "PICKUPS"
                for item in pick.list:
                    notes[bm_route] += "\r\n\r\n"
                    notes[bm_route] += item[0] + "\r\n"
                    notes[bm_route] += item[1] + " - " + item[3] + "\r\n"
                    notes[bm_route] += item[4]
    confirm = "The following routes will be modified:\n"
    for route in sorted(notes.keys()):
        confirm += "\n %s: %i pick ups" % (route, len(notes[route].split('Pick Up - ')) - 1)
    confirm += "\n\nClick OK to continue or Cancel to abort."
    if not tk.messagebox.askokcancel("Pick Ups To Notes", confirm):
        return
    else:
        existing = check_existing(picklist, route_notes)
        route_id_list = get_routeIDs(cursor)
        for route in notes:
            if route in existing:
                for item in route_notes[route]:
                    cursor.execute("UPDATE dbo.Note SET NoteText = ? WHERE NoteID = ?", (notes[route], str(item[0])))
            else:
                cursor.execute("INSERT INTO dbo.Note (ModuleCode, RecordID, NoteDate, NoteText, SendToDevice, AlwaysSendToDevice) VALUES ('Route', ?, '2000-01-01 00:00:00', ?, 1, 1)", (str(route_id_list[route]), notes[route]))
        cursor.commit()
        tk.messagebox.showinfo("Pick Ups To Notes", "Route notes updated.")


def get_config():
    if getattr(sys, 'frozen', False):
        application_path = path.dirname(sys.executable)
    elif __file__:
        application_path = path.dirname(__file__)

    config_path = path.join(application_path, 'config.cfg')

    try:
        with open(config_path, 'r') as file:
            contents = file.read()
            server = re.search('server = (.*)', contents).group(1)
            database = re.search('database = (.*)', contents).group(1)
            username = re.search('username = (.*)', contents).group(1)
            password = re.search('password = (.*)', contents).group(1)
    except FileNotFoundError:
        error = "There was no configuration file found in the root folder."
        error += " Please make sure a config file is included that includes"
        error += " the server, database, username, and password."
        tk.messagebox.showerror("Pick Ups To Notes", error)
    except AttributeError:
        error = "There was a problem with the configuration file."
        error += " Please make sure it includes a server, database, username, and password."
        tk.messagebox.showerror("Pick Ups To Notes", error)
    return [server, database, username, password]
