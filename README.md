# WhatsApp Message Analyzer

## What is it?

It's a script that analyzes all the messages in a given WhatsApp group chat and visualizes the most active users, the most used words and the dates and times with the most activity.

## Setup & Usage:

_Requires Python3 and the following libraries: matplotlib, numpy_

1. Get a chat by going into a group -> Settings -> Email Chat -> No Media.

2. Place the downloaded .txt file containing the chats in the same folder as this script.

3. Run the script in the commandline with the following format:
`analyze.py chat.txt` where chat.txt is the file containing the chats

4. The below images will be generated in the folder along with 4 text files containing the formatted data:

```
dates.txt
people.txt
time.txt
word.txt
dateActivity.png
personActivity.png
timeActivity.png
wordFrequency.png
```

## Screenshots:

![Most Used Words](https://i.imgur.com/cc2LlIo.png)

![Most Active Users](https://i.imgur.com/jfB7Og5.png)

![Message Time Chart](https://i.imgur.com/dl6zYEs.png)

![Most Active Dates](https://i.imgur.com/DsGjMRP.png)
