import os
import shutil
import glob

import time
import re

import PySimpleGUI as sg
import pandas as pd
from zipfile import ZipFile
from PyPDF2 import PdfFileWriter, PdfFileReader

# pyinstaller -F --noconsole -n extractor_comparer app.py

layout = [
    [sg.Text('_' * 50)],
    [sg.Text('Merging and Extracting', font=('Helvetica', 15))],
    [sg.Text('Select .zip file:')],
    [sg.Input(key='zip_input'), sg.FileBrowse(key="zip_input_browse")],
    [sg.Text('Select name:')],
    [sg.Input(default_text='NAME', key='name_input', size=(20, ))],
    [sg.Text('If you want to remove unneccesary pages for merging,\nset the cutpoint here:')],
    [sg.Input(default_text='2', key='cut_input', size=(2, ))],
    [sg.Text('Select process:')],
    [sg.Button('Merge', key='merge'), sg.Button('Extract', key='extract'
     ), sg.Button('Extract and Merge', key='extract_and_merge')],
    [sg.Text('_' * 50)],
    [sg.Text('Compare Documents' , font=('Helvetica', 15))],
    [sg.Text("The results are in the folder of the first document.")],
    [sg.Input(key='document_1_input'), sg.FileBrowse(key="document_1_browse")],
    [sg.Input(key='document_2_input'), sg.FileBrowse(key="document_2_browse")],
    [sg.Button('Compare', key='compare')],
    [sg.Text('_' * 50)],
    [sg.Text('Waiting...', key='process')],
    ]

# sg.theme("LightGreen")
window = sg.Window('App').Layout(layout)


# Variables
file_name = ""
file_folder = ""
temp_folder = ""


# Functions
def extracting(file_path):

    global file_folder
    global temp_folder

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

        # select all .pdfs and loop over the papers until
        # the cut or max number of pages
        for file in glob.glob(temp_folder + "/*.pdf"):
            temp_file = PdfFileReader(file, "rb")
            for p in range(0, min(cut, temp_file.getNumPages())):
                page = temp_file.getPage(p)
                merger.addPage(page)

        if not os.path.exists(file_folder + "\\" + values["name_input"]
                              + ".pdf"):
            with open(file_folder + "\\" + values["name_input"] + ".pdf",
                      'wb') as f:
                merger.write(f)
    else:
        # append all files form temporary folder
        for file in glob.glob(temp_folder + "/*.pdf"):
            merger.append(file)

        # if MERGE file does not exist create one with the appended file
        if not os.path.exists(file_folder + "\\" + values["name_input"]
                              + ".pdf"):
            merger.write(file_folder + "\\" + values["name_input"] + ".pdf")
        merger.close()

def comparing(document_1, document_2):

    file_path = "/".join(document_1.split("/")[:-1])

    def extract_text(file, extract_list):

        with open(file, 'rb') as pdf:
            pdf_read = PdfFileReader(pdf)

            for page in pdf_read.pages:
                text = page.extractText()
                text = re.sub('\n+', '', text)
                text = re.split("\. |\? ", text)
                extract_list.extend(text)

    document_1_sentences = []
    document_2_sentences = []

    extract_text(document_1, document_1_sentences)
    extract_text(document_2, document_2_sentences)

    in_both = list(set(document_1_sentences).intersection(document_2_sentences))

    in_both_str = " \n\n".join(in_both)
    with open(file_path + "\\" + f"compare-{time.time()}.txt", "w") as results:
        results.write(in_both_str)

# While running window
while True:

    event, values = window.Read()

    if event == "extract":
        if os.path.exists(values["zip_input_browse"]):
            window.FindElement("process").Update(value="Working...",
                                                 text_color="#e0741f")
            window.Refresh()

            extracting(values["zip_input_browse"])

            window.FindElement("process").Update(value="Done!",
                                                 text_color="#08c65b")
        else:
            print("not file selected")

    if event == "merge":
        if os.path.exists(values["zip_input_browse"]):
            window.FindElement("process").Update(value="Working...",
                                                 text_color="#e0741f")
            window.Refresh()

            extracting(values["zip_input_browse"])
            merging(temp_folder, file_folder)
            shutil.rmtree(temp_folder)

            window.FindElement("process").Update(value="Done!",
                                                 text_color="#08c65b")
        else:
            print("not file selected")

    if event == "extract_and_merge":
        if os.path.exists(values["zip_input_browse"]):
            window.FindElement("process").Update(value="Working...",
                                                 text_color="#e0741f")
            window.Refresh()

            extracting(values["zip_input_browse"])
            merging(temp_folder, file_folder)

            window.FindElement("process").Update(value="Done!",
                                                 text_color="#08c65b")
        else:
            print("not file selected")

    if event == "compare":
        if os.path.exists(values["document_1_browse"]) and os.path.exists(values["document_2_browse"]):
            window.FindElement("process").Update(value="Working...",
                                                 text_color="#e0741f")
            window.Refresh()

            comparing(values["document_1_browse"], values["document_2_browse"])

            window.FindElement("process").Update(value="Done!",
                                                 text_color="#08c65b")
        else:
            print("no files selected")

    # closing program
    if event is None or event == 'Exit':
        break

window.Close()
