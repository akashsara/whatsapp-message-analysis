#!python3
import os, re, sys, graphs
from operator import itemgetter

#(Date)(: )(Time)(: )(Sender)(: )(Message)
messageFormat = re.compile(r'^([0-9]{2}/[0-9]{2}/[0-9]{2})(, )([0-9]{1,2}:[0-9]{2}:[0-9]{2} [A-Z]{2})(: )(.*?)(: )(.*)')
#[Hours]:[Minutes]:[Seconds] [AM/PM]
timeSplit = re.compile(r'([0-9]{1,2}):([0-9]{2}):([0-9]{2}) ([A-Z]{2})')

#To get rid of file extension when making graphs
fileSplit = re.compile(r'(.*)(.[a-zA-Z0-9]{3,4})')

#Initializing empty lists of dictionaries for date, time, person and word, no of messages to 0 and a message list
noMessages = 0
dateDictionary = []
timeDictionary = []
personDictionary = []
wordDictionary = []
messageList = []

#Copy contents of file to a variable.
def readFile(fileName):
    fi = open(fileName, 'r', encoding='utf-8')
    textToAnalyze = fi.readlines()
    fi.close()
    return textToAnalyze

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

#Copy contents of a dictionary to a file
def toFile(dictionary, keyOne, fileName):
    fi = open(fileName, 'w', encoding='utf-8')
    for items in dictionary:
        fi.write(items[keyOne] + '::' + str(items['Count']) + '\n')
    fi.close()

#Get a list of the most used words of length 5 or more while ignoring links and case. Words are separated by ' '.
def getWordFrequency(messageList):
    stripChars = re.compile(r'[a-zA-z0-9]+')
    frequencyList = []
    for messages in messageList:
        wordsList = messages.split(' ')
        for words in wordsList:
            if stripChars.search(words):
                words = (stripChars.search(words).group()).lower()
                if len(words) < 5 or words == 'https':
                    continue
                for existingWords in frequencyList:
                    if existingWords['Word'] == words:
                        existingWords['Count'] += 1
                        break
                else:
                    frequencyList.append({'Word': words.lower(), 'Count': 1})
    return sort(frequencyList)

#Gets file name from command line arguments
def getFileName():
    #If no filename is given
    if len(sys.argv) < 2:
        print('Usage: analyze.py fileName')
        sys.exit()
    #Get file name from command line
    return ' '.join(sys.argv[1:])


fileName = getFileName()

#Analyze file
textToAnalyze = readFile(fileName)

#for each line, check if it is a new message. If it is, split it up into components.
#The date and sender components are added to their dictionaries.
#Split up the time into hours, minutes, seconds and AM/PM.
#If it is PM, add 12 hours to it. If hours is a single digit, convert it to double digits.
#Each message is appended to a list for later analysis.
for lines in textToAnalyze:
    if messageFormat.search(lines):
        found = messageFormat.search(lines)
        dateDictionary = analyze(dateDictionary, found[1], 'Date')
        personDictionary = analyze(personDictionary, found[5], 'Sender')
        #Time
        time = timeSplit.search(found[3])
        hours = time[1]
        if len(hours) == 1:
            hours = '0' + str(hours)
        if time[4] == 'PM':
            if (int(hours) + 12) >= 24:
                hours = 0
            else:
                hours = int(hours) + 12
        timeDictionary = analyze(timeDictionary, str(hours), 'Time')
        #Message
        messageList.append(found[7])
        noMessages += 1

#Sort dictionaries and get word dictionary from analysing the message list
dateDictionary = sort(dateDictionary)
personDictionary = sort(personDictionary)
timeDictionary = sorted(timeDictionary, key=itemgetter('Time'))
wordDictionary = getWordFrequency(messageList)

#Print number of messages
print('\nTotal number of messages = %s' %noMessages)

#Generate graphs
newFileName = fileSplit.search(fileName)[1]
graphs.histogram(timeDictionary, 'Message Time Chart in ' + newFileName, 'timeActivity.png')
graphs.barGraph(wordDictionary[:15], 'Word', 'Uses', 'Most used words in ' + newFileName, 'wordFrequency.png')
graphs.barGraph(dateDictionary[:15], 'Date', 'Messages', 'Most Messages in ' + newFileName, 'dateActivity.png')
graphs.barGraph(personDictionary[:15], 'Sender', 'Messages', 'Most active person in ' + newFileName, 'personActivity.png')


#Write to files
toFile(dateDictionary, 'Date', 'dates.txt')
toFile(personDictionary, 'Sender', 'people.txt')
toFile(timeDictionary, 'Time', 'time.txt')
toFile(wordDictionary, 'Word', 'word.txt')
