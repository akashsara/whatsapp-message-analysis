#!python3
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
from operator import itemgetter

def bar_graph(dictionary, limit, label, title, fileName):
    x = []
    y = []
    for count, (key, value) in enumerate(dictionary.items()):
        if count == limit:
            break
        x.append(key)
        y.append(value)
    make_bar_graph(x, y, label, title, fileName)

def make_bar_graph(x, y, label, title, save_as):
    plt.figure(figsize=(10,6), dpi=160)
    ax = sns.barplot(y, x)
    ax.set_xlabel(label)
    ax.set_title(title)
    plt.tight_layout()
    plt.savefig(save_as)

def histogram(dictionary, title, save_as):
    plt.figure(figsize=(10,6), dpi=160)
    sns.barplot(
        x=[int(x)for x in dictionary.keys()],
        y=list(dictionary.values())
    )
    plt.xlabel('Time')
    plt.ylabel('Messages')
    plt.title(title)
    plt.tight_layout()
    plt.savefig(save_as)