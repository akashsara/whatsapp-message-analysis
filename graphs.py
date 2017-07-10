#!python3
import matplotlib.pyplot as plt
import numpy as np
from operator import itemgetter

#Set up the words as x and its count as y
def barGraph(dictionary, keyName, label, title, fileName):
    x = []
    y = []
    for items in dictionary:
        x.append(items[keyName])
        y.append(items['Count'])
    makeBarGraph(x, y, label, title, fileName)

def makeBarGraph(x, y, label, title, saveAs):
        fig, ax = plt.subplots()
        y_pos = np.arange(len(y))
        ax.barh(y_pos, y)
        ax.set_yticks(y_pos)
        ax.set_yticklabels(x)
        ax.invert_yaxis()
        ax.set_xlabel(label)
        ax.set_title(title)
        plt.tight_layout()
        plt.savefig(saveAs)

def histogram(dictionary, title, saveAs):
    dataSet = []
    bins = []
    for items in dictionary:
        for i in range(items['Count']):
            dataSet.append(int(items['Time']))
        bins.append(int(items['Time']))
    plt.xlim(0,23)
    plt.xticks(bins)
    plt.hist(dataSet, bins, align='mid')
    plt.xlabel('Time')
    plt.ylabel('Messages')
    plt.title(title)
    plt.tight_layout()
    plt.savefig(saveAs)
