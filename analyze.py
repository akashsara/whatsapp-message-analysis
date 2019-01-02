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
messageFormat = re.compile(r'^([0-9]{1,2}/[0-9]{2}/[0-9]{2})(, )([0-9]{1,2}:[0-9]{2})(:[0-9]{2})*( [A-Z]{2})( -|:)( )(.*?)(: )(.*)')
#[Hours]:[Minutes]
timeSplit = re.compile(r'([0-9]{1,2}):([0-9]{2})')

#To get rid of file extension when making graphs
fileSplit = re.compile(r'(.*)(.[a-zA-Z0-9]{3,4})')

#Initializing empty lists of dictionaries for date, time, person and word, no of messages to 0 and a message list
noMessages = 0
dateDictionary = []
timeDictionary = []
personDictionary = []
wordDictionary = []
messageList = []

#If value exists in the dictionary, increment count. Else add a new entry.
def analyze(dictionary, value, keyName):
    for d in dictionary:
        if d[keyName] == value:
            d['Count'] += 1
            break
    else:
        dictionary.append({keyName: value, 'Count': 1})
    return dictionary

#Sort a dictionary in descending order of its count
def sort(dictionary):
    return sorted(dictionary, key=itemgetter('Count'), reverse=True)

#Get a dictionary of the most used words in the chat and the number of times they're used while ignoring commonly used words.
def getWordFrequency(messageList):
    wordFile = open('commonWords.txt', 'r')
    commonWordsList = wordFile.read()
    wordFile.close()
    stripChars = re.compile(r'[a-zA-z0-9]+')
    frequencyList = []
    for messages in messageList:
        wordsList = messages.split(' ')
        for words in wordsList:
            words = words.lower()
            if words in commonWordsList:
                continue
            if stripChars.search(words):
                words = stripChars.search(words).group()
                for existingWords in frequencyList:
                    if existingWords['Word'] == words:
                        existingWords['Count'] += 1
                        break
                else:
                    frequencyList.append({'Word': words, 'Count': 1})
    return sort(frequencyList)

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
text_to_analyze = readFile(file_name)

#for each line, check if it is a new message. If it is, split it up into components.
#The date and sender components are added to their dictionaries.
#Split up the time into hours, minutes, seconds and AM/PM.
#If it is PM, add 12 hours to it. If hours is a single digit, convert it to double digits.
#Each message is appended to a list for later analysis.
for lines in text_to_analyze:
    if messageFormat.search(lines):
        found = messageFormat.search(lines)
        dateDictionary = analyze(dateDictionary, found[1], 'Date')
        personDictionary = analyze(personDictionary, found[8], 'Sender')
        #Time
        time = timeSplit.search(found[3])
        hours = time[1]
        if len(hours) == 1:
            hours = '0' + str(hours)
        if found[5] == ' PM':
            if (int(hours) + 12) >= 24:
                hours = 0
            else:
                hours = int(hours) + 12
        timeDictionary = analyze(timeDictionary, str(hours), 'Time')
        #Message. Ignore media.
        if '<Media omitted>' not in found[10] and '<‎attached>' not in found[10]:
            messageList.append(found[10])
        noMessages += 1

#Sort dictionaries and get word dictionary from analysing the message list
dateDictionary = sort(dateDictionary)
personDictionary = sort(personDictionary)
timeDictionary = sorted(timeDictionary, key=itemgetter('Time'))
wordDictionary = getWordFrequency(messageList)

#remove old data sheets
if os.path.isfile('output/data.xlsx'):
    os.unlink('output/data.xlsx')

#Add to excel sheet
to_xl(dateDictionary, 'Dates', 'Date', 'No. of Messages')
to_xl(personDictionary, 'People', 'Sender', 'No. of Messages')
to_xl(timeDictionary, 'Times', 'Time', 'No. of Messages')
to_xl(wordDictionary, 'Words', 'Word', 'No. of Uses')

#Generate graphs
newFileName = fileSplit.search(fileName)[1]
graphs.histogram(timeDictionary, 'Message Time Chart in ' + newFileName, 'output/timeActivity.png')
graphs.barGraph(wordDictionary[:15], 'Word', 'Uses', 'Most used words in ' + str(noMessages) + ' messages in ' + newFileName, 'output/wordFrequency.png')
graphs.barGraph(dateDictionary[:15], 'Date', 'Messages', 'Most Messages in ' + newFileName, 'output/dateActivity.png')
graphs.barGraph(personDictionary[:15], 'Sender', 'Messages', 'Most active person in ' + newFileName, 'output/personActivity.png')
