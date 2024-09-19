##########################
# Card Creation
# Creates Linegraph and Histogram with mean, min, max and standard deviation displayed
# SciPy is a scientific computation library that uses NumPy underneath. 
# SciPy stands for Scientific Python. It provides more utility functions for optimization, stats and signal processing. 
# in this program  scipy. kurtosis a submodule of scipy is used. 
# Kurtosis is a measure of whether the data are heavy-tailed or light-tailed relative to a normal distribution.
##########################

from fpdf import FPDF
from datetime import datetime
import copy
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure as Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import statistics
import scipy.stats as stats

def create_linegraph(root,df_data,dSIMULATION_SEQUENCES):
    all_data_set = copy.deepcopy(df_data)  # Deep copy - Copies the whole array without referencing the original copy.

    figure = Figure(figsize=(5, 4), dpi=100)
    ax = figure.add_subplot(111)
    linegraph = FigureCanvasTkAgg(figure, root)

    all_data_set['max_win'] = all_data_set.iloc[:, 2:-1].max(axis=1)
    all_data_set['mean_win'] = all_data_set.iloc[:, 2:-2].mean(axis=1)
    all_data_set['stdv_win'] = all_data_set.iloc[:, 2:-3].std(axis=1)
    # mean
    data = all_data_set['mean_win']
    res = stats.cumfreq(data, numbins=25)
    x = res.lowerlimit + np.linspace(0, res.binsize * res.cumcount.size, res.cumcount.size)
    plt.plot(x, res.cumcount,label="Mean")

    # SDEV
    sdev = all_data_set['stdv_win'].mean()
    res2 = np.histogram(data, 25)[0].tolist()  # get stdev bins from hist function
    # plt.plot(x, res.cumcount)
    plt.fill_between(x, res.cumcount - res2, res.cumcount + res2, alpha=0.3,label="Stdev")

    # max
    data = all_data_set['max_win']
    res = stats.cumfreq(data, numbins=25)
    plt.plot(x, res.cumcount,'--',label="Min")
    # min
    data = all_data_set['first win']
    res = stats.cumfreq(data, numbins=25)
    plt.plot(x, res.cumcount,'--',label="Max",)
    ax.legend(loc='upper left')



    # plotting
    plt.xlabel('Total Number Called')
    plt.ylabel('Winners')

    plt.title('Number of winners per numbers called')


    linegraph.get_tk_widget().pack()

def create_histogram(root,df_data,dSIMULATION_SEQUENCES):
    all_data_set = copy.deepcopy(df_data)
    # sort data by simulation and first win
    value_counts = all_data_set['first win'].value_counts().sort_index()

    # converting to df and assigning new names to the columns
    df_value_counts = pd.DataFrame(value_counts)
    df_value_counts = df_value_counts.reset_index()
    df_value_counts.columns = ['numbers', 'wins']  # change column names
    df1 = df_value_counts
    res = stats.cumfreq(df1['numbers'], numbins=df1.shape[0],
                        defaultreallimits=(1.5, 5))

    # generating random values
    data = all_data_set['first win']

    res = stats.cumfreq(data,
                        numbins=25)

    x = res.lowerlimit + np.linspace(0, res.binsize * res.cumcount.size,
                                     res.cumcount.size)

    figure = Figure(figsize=(5, 4), dpi=100)
    ax = figure.add_subplot(111)
    linegraph = FigureCanvasTkAgg(figure, root)

    # cumulative graph
    ax.bar(x, res.cumcount, width=4, color="blue")

    # setting up the title
    ax.set_title('Number of times BINGO was called occurred on a number')
    plt.xlabel("Number of numbers")  # add X-axis label
    plt.ylabel("Number of Bingos")  # add Y-axis label
    ax.set_xlim([x.min(), x.max()])

    # display the figure(histogram)
    linegraph.get_tk_widget().pack()



















