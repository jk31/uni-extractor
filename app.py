import os
import shutil
import glob

import re

import PySimpleGUI as sg
import pandas as pd
from zipfile import ZipFile
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger

layout = [[sg.Text('Select .zip file:')],
          [sg.Input(key="file_input"),
           sg.FileBrowse()], [sg.Text('Select name:')],
          [sg.Input(default_text="NAME", key="name_input", size=(20, ))],
          [sg.Text('If merging, select cut point:')],
          [sg.Input(default_text="2", key="cut_input", size=(2, ))],
          [sg.Text('Select process:')],
          [
              sg.Button("Merge", key="merge"),
              sg.Button("Extract", key="extract"),
              sg.Button("Extract and Merge", key="extract_and_merge")
          ], [sg.Text('Waiting...', key="process")]]

#
window = sg.Window('MERGER').Layout(layout)

# Variables
file_name = ""
file_folder = ""
temp_folder = ""


# Functions
def extracting(file_path):

    global file_name
    global file_folder
    global temp_folder

    file_name = file_path.split("/")[-1]
    file_folder = "/".join(file_path.split("/")[:-1])

    # make a temporary folder
    temp_folder = file_folder + "/extracted_" + values["name_input"]
    if not os.path.exists(temp_folder):
        os.mkdir(temp_folder)

    # extract pdf from zip in temporary folder
    with ZipFile(file_path, "r") as zip:

        alle = zip.namelist()

        df = pd.DataFrame(alle, columns=["names"])
        df["pair"] = df["names"].apply(
            lambda x: re.findall(r"[A-Za-z0-9 ]*-", x)[0])
        df = df.groupby("pair").agg({"names": "first"})

        for i in df.names:
            zip.extract(i, temp_folder)

        zip.close()


def merging(temp_folder, file_folder):

    merger = PdfFileWriter()
    # check if we want to cut the .pdfs
    if values["cut_input"] != "":
        cut = int(values["cut_input"])

        # select all .pdfs and loop over the papers until the cut or max number of pages
        for file in glob.glob(temp_folder + "/*.pdf"):
            temp_file = PdfFileReader(file, "rb")
            for p in range(0, min(cut, temp_file.getNumPages())):
                page = temp_file.getPage(p)
                merger.addPage(page)

        if not os.path.exists(file_folder + "\\" + values["name_input"] +
                              ".pdf"):
            with open(file_folder + "\\" + values["name_input"] + ".pdf",
                      'wb') as f:
                merger.write(f)
    else:

        # append all files form temporary folder
        for file in glob.glob(temp_folder + "/*.pdf"):
            merger.append(file)

        # if MERGE file does not exist create one with the appended files

        if not os.path.exists(file_folder + "\\" + values["name_input"] +
                              ".pdf"):
            merger.write(file_folder + "\\" + values["name_input"] + ".pdf")
        merger.close()


def extract_button(file_path):

    extracting(file_path)


def merge_button(file_path):

    extracting(file_path)
    merging(temp_folder, file_folder)

    shutil.rmtree(temp_folder)


def extract_and_merge_button(file_path):

    extracting(file_path)
    merging(temp_folder, file_folder)


# While running window
while True:

    event, values = window.Read()

    if event == "extract":
        if os.path.exists(values["Browse"]):
            window.FindElement("process").Update(value="Working...",
                                                 text_color="#e0741f")
            window.Refresh()

            extract_button(values["Browse"])

            window.FindElement("process").Update(value="Done!",
                                                 text_color="#08c65b")
        else:
            print("not file selected")

    if event == "merge":
        if os.path.exists(values["Browse"]):
            window.FindElement("process").Update(value="Working...",
                                                 text_color="#e0741f")
            window.Refresh()

            merge_button(values["Browse"])

            window.FindElement("process").Update(value="Done!",
                                                 text_color="#08c65b")
        else:
            print("not file selected")

    if event == "extract_and_merge":
        if os.path.exists(values["Browse"]):
            window.FindElement("process").Update(value="Working...",
                                                 text_color="#e0741f")
            window.Refresh()

            extract_and_merge_button(values["Browse"])

            window.FindElement("process").Update(value="Done!",
                                                 text_color="#08c65b")
        else:
            print("not file selected")

    # closing program
    if event is None or event == 'Exit':
        break

window.Close()
