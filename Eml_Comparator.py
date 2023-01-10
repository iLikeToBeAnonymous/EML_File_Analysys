# Largely based on my previous file found here:
# https://github.com/iLikeToBeAnonymous/Category_Extraction_From_Product_Names/blob/separate_files/RakeTest_fileLoad.py

import string
import json # json lib for functions and prettifying things.
import re # so you co do regex matching and stuff.
import os # ChatGPT
import email # for parsing .eml files
from os import path
from collections import Counter # ChatGPT

# # Compile the regex used to match the field name
# fieldNmRegex = re.compile("^(?<fieldName>[A-Za-z0-9\-]+):.*$")

# myDir = path.abspath(path.dirname(__file__))
# targetFolder = 'ThunderbirdExports-TESTING'
# targetFilename = 'Mail delivery failed  returning message to sender - Mail Delivery System (Mailer-Daemon@se1-lax1.servconfig.com) - 2022-11-23 1649.eml'

# # JOIN THE THREE SEPARATE COMPONENTS TOGETHER TO MAKE A COMPLETE LITERAL PATH.
# literalPath = path.join(myDir, targetFolder, targetFilename)
# # print('My dir:  ' + literalPath)


# rawData = open(literalPath, 'r') # Opens the source data which contains the text from which to extract keywords

# # #REMEMBER! .read() reads the file as a blob, but readlines() reads the file into an array of strings, each row delimiting a line.
# # myText = rawData.read() # reads the complete source text into a variable
# myText = rawData.readlines()

# myctr = 0 # Set the line counter to zero
# for eachEle in myText:
#     myctr += 1
#     # print("Row " + str(myctr) + ":  " + eachEle)

###############################
### BEGIN CODE FROM ChatGPT ###
def extract_fields(eml_file):
    with open(eml_file, 'rb') as fileContents:
        msg = email.message_from_binary_file(fileContents)
        fields = []
        for field in msg.keys():
            fields.append(field)
        return fields

# V1 from ChatGPT---------------------------------#
# def analyze_folder(folder):
#     fields_list = []
#     for eml_file in os.listdir(folder):
#         if eml_file.endswith('.eml'):
#             fields = extract_fields(os.path.join(folder, eml_file))
#             fields_list.extend(fields)
#     return fields_list
#-------- End V1 ---------------------------------#

# V2 from ChatGPT---------------------------------#
# def analyze_folder(folder):
#     fields_list = []
#     for eml_file in os.listdir(folder):
#         if eml_file.endswith('.eml'):
#             fields = extract_fields(os.path.join(folder, eml_file))
#             fields_list.extend(fields)
#     return Counter(fields_list)
#-------- End V2 ---------------------------------#

# V3 from ChatGPT---------------------------------#
def analyze_folder(folder):
    fields_list = []
    for eml_file in os.listdir(folder):
        if eml_file.endswith('.eml'):
            fields = extract_fields(os.path.join(folder, eml_file))
            fields_list.extend(fields)
    fields_counter = Counter(fields_list)
    fields_dict = {}
    for field, count in fields_counter.items():
        fields_dict[field] = count
    return fields_dict
#-------- End V3 ---------------------------------#

# folder = '/path/to/folder' # Original from ChatGPT. Below is my version.
#------------------BEGIN MINE ---------------------#
myDir = path.abspath(path.dirname(__file__)) # mine
targetFolder = 'ThunderbirdExports-TESTING' # mine
folder_path = path.join(myDir, targetFolder) # mine
#------------------ END MINE ----------------------#

# # Below two lines work with analyze_folder V1 and V2
# fields_list = analyze_folder(folder_path) 
# print(fields_list)

# Below two lines work with analyze_folder V3
fields_dict = analyze_folder(folder_path)
print(json.dumps(fields_dict, indent=4))
#### END CODE FROM ChatGPT ####
###############################
