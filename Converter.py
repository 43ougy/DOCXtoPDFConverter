#!/bin/python3

import convertapi
import os
import sys
from dotenv import load_dotenv

load_dotenv()

convertapi.api_secret = os.getenv("API_SECRET")
convertapi.api_key = os.getenv("API_KEY")

dir_list=os.listdir(os.getcwd())
convertDir="ConvertFiles"
argumentIndex=1

def ManageConvertDir():
    global argumentIndex
    global convertDir
    if (sys.argv[argumentIndex] == "-d"):
        argumentIndex += 2
        convertDir = sys.argv[2]
    if (os.path.isdir(convertDir) == False):
        os.mkdir(convertDir)

def CheckEmptyDir(directory: str) -> None:
    filesList=os.listdir(directory)
    for index in filesList:
        if (os.path.isdir(directory + '/' + index)):
            CheckEmptyDir(directory + '/' + index)
    if (len(os.listdir(directory)) == 0):
        print("Directory is empty --> ", directory)
        os.rmdir(directory)
    else:
        print("Directory is NOT empty --> ", directory)

#function to convert a docx file to a pdf with the api
def ConvertToPdf(file: str, directory: str) -> None:
    base_name, _ = os.path.splitext(file)
    new_file_path = base_name + ".pdf"
    if (os.path.isdir(directory) == True and os.path.exists(convertDir + '/' + new_file_path)):
        print("File already exist --> ", directory + '/' + new_file_path)
    else:
        convertapi.convert('pdf', { 'File': file }, from_format='docx').save_files(directory)
        print(file, " --> convert to pdf")

#function to recursivly go through directory and convert docx files
def DirectoryConvert(directory: str) -> None:
    newDir=convertDir+'/'+directory
    if (os.path.isdir(newDir) == False):
        os.mkdir(newDir)
    filesList=os.listdir(directory)
    for index in filesList:
        if (os.path.isdir(directory + '/' + index) and len(os.listdir(directory + '/' + index)) > 0):
            print(index, " --> is a directory in ", directory)
            DirectoryConvert(directory + '/' + index)
    for index in filesList:
        print("in directory ", directory, " --> ", index)
        if (index.lower().endswith(".docx")):
            print(directory + '/' + index)
            ConvertToPdf(directory + '/' + index, newDir)

if (len(sys.argv) == 1):
    print("./converter [flag -d (path to convert directory)] ; [path to file] - [path to directory]")

elif (len(sys.argv) >= 2):
    ManageConvertDir()
    print(convertDir)
    if (sys.argv[argumentIndex] == 'all' and len(sys.argv) == argumentIndex + 1):
        for index in dir_list:
            if (os.path.isdir(index) and index != '.venv'):
                DirectoryConvert(index)
            elif (index.lower().endswith(".docx")):
                print(index, " --> is a file in ", os.getcwd())
                ConvertToPdf(index, convertDir)
    
    else:
        for index in sys.argv:
            if (os.path.isdir(index) and index != '.venv'):
                DirectoryConvert(index)
            elif (index.lower().endswith(".docx")):
                print(index, " --> is a file in ", os.getcwd())
                ConvertToPdf(index, convertDir)
    CheckEmptyDir(convertDir)



# API command to convert from docx to pdf
#
#                       can use a string instead of a dictionary here
#                                            |
#                           _________________|_________________
# convertapi.convert('pdf', { 'File': '/path/to/my_file.docx' }, from_format = 'docx').save_files('/path/to/dir')