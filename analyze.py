#!python3
import os
import re
import sys
import graphs
import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import NamedStyle
from operator import itemgetter

#(Date)(, )(Hours:Minutes)(Seconds)(AM/PM)(:/ -)( )(Sender)(: )(Message)
message_format = re.compile(r'^([0-9]{1,2}/[0-9]{2}/[0-9]{2})(, )([0-9]{1,2}:[0-9]{2})(:[0-9]{2})*( [A-Z]{2})( -|:)( )(.*?)(: )(.*)')
#[Hours]:[Minutes]
time_split = re.compile(r'([0-9]{1,2}):([0-9]{2})')

#To get rid of file extension when making graphs
file_split = re.compile(r'(.*)(.[a-zA-Z0-9]{3,4})')

#Initializing empty lists of dictionaries for date, time, person and word, no of messages to 0 and a message list
number_of_messages = 0
date_dictionary = {}
time_dictionary = {}
person_dictionary = {}
word_dictionary = {}
message_list = []

#If value exists in the dictionary, increment count. Else add a new entry.
def add_to_dictionary(dictionary, value):
    if value in dictionary:
        dictionary[value] += 1
    else:
        dictionary[value] = 1
    return dictionary

#Sort a dictionary in descending order of its count
def sort(dictionary):
    return sorted(dictionary, key=itemgetter('Count'), reverse=True)

#Get a dictionary of the most used words in the chat and the number of times they're used while ignoring commonly used words.
def get_word_frequency(message_list):
    with open('commonWords.txt', 'r') as word_file:
        common_words_list = word_file.read()
    clean_word = re.compile(r'[a-zA-z0-9]+')
    frequency_list = {}
    for messages in message_list:
        words_list = messages.split(' ')
        for words in words_list:
            word = word.lower()
            if word in common_words_list:
                continue
            if clean_word.search(word):
                word = clean_word.search(word).group()
                if word in frequency_list:
                    frequency_list[word] += 1
                else:
                    frequency_list[word] = 1
    return sort(frequency_list)

"""
Function to create a new sheet or find an existing sheet.
Returns a Worksheet object as well as the Workbook object.
First, it checks if the file: data.xlsx already exists. 
If the file exists, just create a new sheet with the given sheet_name argument. 
Otherwise, create a new Workbook and set the sheet name. 
Return the Workbook, Worksheet
"""
def get_workbook_and_sheet(sheet_name):
    if os.path.isfile('output/data.xlsx'):
        xl = openpyxl.load_workbook('output/data.xlsx')
        xl.create_sheet(title=sheet_name)
        sheet = xl[sheet_name]
    else:
        if not os.path.exists('output'):
            os.mkdir('output')
        xl = openpyxl.Workbook()
        sheet = xl.active
        sheet.title = sheet_name
    return xl, sheet

"""
Function to save a dictionary to an Excel file.
Returns nothing.
Calls the get_workbook_and_sheet() function to get the workbook and sheet.
Sets the column width & headers before inserting each item of the dictionary to the sheet.
Then saves the workbook.
"""
def to_xl(dictionary, sheet_name, col1, col2):
    xl, sheet = get_workbook_and_sheet(sheet_name)
    # Column Width
    sheet.column_dimensions['A'].width = 30
    sheet.column_dimensions['B'].width = 15
    # Column Headers
    sheet.cell(row=1, column=1).value = col1
    sheet.cell(row=2, column=2).value = col2
    # Insert Data
    for row, item in enumerate(dictionary):
        sheet.cell(row=row + 2, column=1).value = item[col1]
        sheet.cell(row=row + 2, column=2).value = item['Count']
    xl.save('output/data.xlsx')

"""
Function to accept command-line arguments.
If no arguments are found, then it prints an error and stops the script.
Else it returns the arguments.
"""
def get_file_name():
    #If no filename is given
    if len(sys.argv) < 2:
        print('Usage: analyze.py file_name.extension')
        sys.exit()
    #Get file name from command line
    return ' '.join(sys.argv[1:])

"""
Function to read the input file of chats. 
Copies it to a variable and returns it.
"""
def read_file(file_name):
    with open(file_name, 'r', encoding='utf-8') as fi:
        text_to_analyze = fi.readlines()
    return text_to_analyze
    
file_name = get_file_name()

#Analyze file
text_to_analyze = read_file(file_name)

#for each line, check if it is a new message. If it is, split it up into components.
#The date and sender components are added to their dictionaries.
#Split up the time into hours, minutes, seconds and AM/PM.
#If it is PM, add 12 hours to it. If hours is a single digit, convert it to double digits.
#Each message is appended to a list for later analysis.
for lines in text_to_analyze:
    if message_format.search(lines):
        found = message_format.search(lines)
        date_dictionary = add_to_dictionary(date_dictionary, found[1])
        person_dictionary = add_to_dictionary(person_dictionary, found[8])
        #Time
        time = time_split.search(found[3])
        hours = time[1]
        if len(hours) == 1:
            hours = '0' + str(hours)
        if found[5] == ' PM':
            if (int(hours) + 12) >= 24:
                hours = 0
            else:
                hours = int(hours) + 12
        time_dictionary = add_to_dictionary(time_dictionary, str(hours))
        #Message. Ignore media.
        if '<Media omitted>' not in found[10] and '<‎attached>' not in found[10]:
            messgae_list.append(found[10])
        number_of_messages += 1

#Sort dictionaries and get word dictionary from analysing the message list
date_dictionary = sort(date_dictionary)
person_dictionary = sort(person_dictionary)
time_dictionary = sorted(time_dictionary, key=itemgetter('Time'))
word_dictionary = get_word_frequency(messgae_list)

#remove old data sheets
if os.path.isfile('output/data.xlsx'):
    os.unlink('output/data.xlsx')

#Add to excel sheet
to_xl(date_dictionary, 'Dates', 'Date', 'No. of Messages')
to_xl(person_dictionary, 'People', 'Sender', 'No. of Messages')
to_xl(time_dictionary, 'Times', 'Time', 'No. of Messages')
to_xl(word_dictionary, 'Words', 'Word', 'No. of Uses')

#Generate graphs
new_file_name = file_split.search(file_name)[1]
graphs.histogram(time_dictionary, 'Message Time Chart in ' + new_file_name, 'output/timeActivity.png')
graphs.barGraph(word_dictionary[:15], 'Word', 'Uses', 'Most used words in ' + str(number_of_messages) + ' messages in ' + new_file_name, 'output/wordFrequency.png')
graphs.barGraph(date_dictionary[:15], 'Date', 'Messages', 'Most Messages in ' + new_file_name, 'output/dateActivity.png')
graphs.barGraph(person_dictionary[:15], 'Sender', 'Messages', 'Most active person in ' + new_file_name, 'output/personActivity.png')
