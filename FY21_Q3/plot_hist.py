#!/usr/bin/env pyhton
# -*- coding: UTF-8 -*-


__author__ = 'Chao Wu'
__date__ = '06/16/2021'
__version__ = '1.0'


r'''
This script plots histogram 

python C:\Users\cwu\Desktop\Software\Aspen_automation\Scripts\case\FY21_Q3\plot_hist.py
'''


OUT_DIR = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\plot_hist'
DATA_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\plot_hist\MFSP-Rin.xlsx'
XLABEL = 'MFSP - Rin ($/GGE)'


import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


def read_data(data_file):
	'''
	Parameters
	data_file: str, data file
	'''
	
	data = pd.read_excel(data_file, header = None, index_col = None, squeeze = True)
	
	return data	
	
	
def plot_hist_and_save(out_dir, data, xlabel, percentile = 10):
	'''
	Parameters
	out_dir: str, output directory
	data: ser, data to plot
	xlabel: str, label of xaxis
	percentile: float of 0 - 100, lines indicate percentile% and 1 - percentile% will be plotted
	'''
	
	fig, ax1 = plt.subplots()
	sns.distplot(data, rug = True, kde = False, hist = True, ax = ax1)
	ax1.set_xlabel(xlabel, fontsize = 15)
	ax1.set_ylabel('Count', color = 'steelblue', fontsize = 15)
	
	ax2 = ax1.twinx()
	sns.distplot(data, rug = True, kde = True, hist = False, ax = ax2)   # if plot kde, y axis can not be Count
	ax2.set_ylabel('')
	ax2.set_yticks([])
	ax2.spines['left'].set_visible(False)
	ax2.spines['right'].set_visible(False)
	ax2.spines['top'].set_visible(False)
	ax2.spines['bottom'].set_visible(False)
	
	counts, edges = np.histogram(data, bins = int(data.size/10) if data.size > 20 else data.size)
	x = (edges[:-1]+edges[1:])/2
	y = np.cumsum(counts)/np.sum(counts)
	p1, p2 = np.percentile(data, [percentile, 100-percentile])
	
	ax3 = ax1.twinx()
	ax3.plot(x, y, color = 'seagreen')
	ax3.set_ylabel('Cumulative probabilty', color = 'seagreen', fontsize = 15)
	ax3.vlines(x = p1, ymin = 0, ymax = 1, linestyles = 'dashed', color = 'gray')
	ax3.vlines(x = p2, ymin = 0, ymax = 1, linestyles = 'dashed', color = 'gray')
	
	fileName = re.sub(r'\s*\(.*?\)\s*', '', xlabel)
	fig.savefig('%s/%s.jpg' % (out_dir, fileName), dpi = 300, bbox_inches = 'tight')
	plt.close()




if __name__ == '__main__':
	
	data = read_data(DATA_FILE)
	
	plot_hist_and_save(OUT_DIR, data, XLABEL)
	
	
	


