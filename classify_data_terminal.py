import os
import tkinter as tk
from tkinter import filedialog

import pandas as pd
from colorama import Fore
from datetime import datetime
from pathlib import Path

import spacy

global cols

# getting file path from user ----------
root = tk.Tk()
root.withdraw()
# get the path for real
file_path = filedialog.askopenfilename()
path_str = Path(file_path).stem

# do some stuff with date time
now = datetime.now()
dtstring = now.strftime("%Y%m%d")
new_filename = dtstring + "_" + path_str + "_categorized.xlsx"

# import excel data for use
df = pd.read_excel(file_path)
# --------------------------------------


# get nlp ready to go. ------------------
nlp = spacy.load("en_core_web_lg")
# ---------------------------------------


# function to clear console--------------
def clear_console():
    os.system('cls')
# --------------------------------------


# ----------------------------------------
def format_print(dct):
    out = ""
    for fb, ids in dct.items():
        out += (("{}: {}".format(fb, ids)) + "\n")

    return out.rstrip()
# ----------------------------------------


# --------------------------------------
# function to go add keys to dictionary if they don't exist
def check_keys(res, out, id_index):
    # this should create a decent dictionary
    # ---------------------------------------
    # iterate through the feedback, could turn into a function.
    global cols
    new_key = ''

    fb = strip(fb)

    for fb in out:
        # save trouble if empty string
        if fb == '':
            break

        # initiate boolean
        is_present = False

        # get the first phrase going for similarity
        phrase_one = nlp(fb)

        # check if the current key exists in the dictionary
        for key in res:

            # second phrase
            phrase_two = nlp(key)
            # check similarity
            if phrase_one.similarity(phrase_two) >= .99:
                new_key = key
                is_present = True

        # either create the key or just append based on its presence.
        if is_present is False:
            res[fb] = []
            res[fb].append(df.at[id_index, cols[0]])
        elif is_present:
            res[new_key].append(df.at[id_index, cols[0]])

    return res
# ---------------------------------------------


# big function-------------------------------------------------------------------------------------
def classify_data():
    # get list of columns
    global cols
    cols = df.columns.values.tolist()

    # master list of feedback sure for now, with an ID col only if it exists.
    if cols[0].strip() == 'ID' or cols[0].strip() == 'Participant ID' or cols[0].strip() == 'Participant':
        master = ["Feedback"]
    else:
        master = []

    # can use df.at[x, "ID"] to get the ID label. 0 index, so start a tmp var at zero and increment
    # iterate through cols
    for i in cols:

        # feedback array
        res = {}

        # temp var for confirming it's not an id col
        is_id = False

        # temporary var for IDs
        id_index = 0

        # iterate through col contents
        for item in df[i]:
            # checks to make sure that we aren't looking at the ID column.
            if i.strip() == 'ID' or i.strip() == 'Participant ID':
                is_id = True
                break
            else:
                is_id = False

            # checking to make sure there's a value there.
            if item is None or item == '' or pd.isna(item):
                tmp = 'N/A'
            else:
                # makes sure console is clear
                clear_console()

                # prints out the question and answer that needs to be coded/classified.
                print(Fore.GREEN + "Question: \n-------------------------------\n", Fore.WHITE, i, "\n")
                print(Fore.RED + "Data: \n-------------------------------\n", Fore.WHITE, item, "\n")

                # getting user input
                print(Fore.CYAN + "Categorize Feedback (separate each entry with a comma):\n-----------------------------------------------------------\n")
                print("Recent Responses (reusing these will make life easier):\n")

                # get keys and print them in a makeshift list
                ls = res.keys()
                for key in ls:
                    print("- " + key + "\n")

                # print a divider to make people happy
                print("-----------------------------------------------------------\n")
                # takes input as comma separated feedback.
                tmp = input(Fore.WHITE)

            # make a new list and add the current ID to 0 zero index. splits cs feedback into a list based on the commas
            out = tmp.split(",")

            # use nlp to merge any similar keys
            res = check_keys(res, out, id_index)

            # increment ID var.
            id_index = id_index + 1

        # make a nice string from the dictionary.
        app = format_print(res)

        # append to master list, assuming its not an id column
        if not is_id:
            master.append(app)

    # return the data we need.
    return master
# --------------------------------------------------------------------------------------------------

# this adds feedback to the bottom cells.
addition = classify_data()
print(addition)
df.loc[df.shape[0]] = addition
# print to excel or to csv? I fucking hate excel. but like...
df.to_excel(new_filename)
