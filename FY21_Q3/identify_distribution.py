#!/usr/bin/env pyhton
# -*- coding: UTF-8 -*-


__author__ = 'Chao Wu'
__date__ = '05/15/2021'
__version__ = '1.1'


r'''
This script identifies the distribution of a continuous variable by fitting to the following unimodal distributions: "alpha", "beta", "triangular", "normal", "gamma" and "pareto"

python C:\Users\cwu\Desktop\Software\Aspen_automation\Scripts\case\FY21_Q3\identify_distribution.py
'''


OUT_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\Identify_distribution\fitted_distributions.jpg'
DATA_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\Identify_distribution\data.xlsx'
DATA_LABEL = 'Weekly values'


import os
import numpy as np
import pandas as pd
from scipy import stats
from scipy.stats import kstest
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings("ignore")


def read_data(data_file):
	'''
	Parameters
	data_file: str, data file
	
	Returns
	data: ser
	'''
	
	data = pd.read_excel(data_file, header = 0, index_col = 0, squeeze = True)
	
	return data
	
	
def identify_distribution(data):
	'''
	Parameters
	data: ser of data
	
	Returns
	fitPDFs: dict of ser
	'''
	
	distNames = ['alpha', 'gamma', 'beta', 'triang', 'norm', 'pareto', 'uniform']
	
	fitPDFs = {}
	for distName in distNames:
		
		# fit to known distribution
		dist = getattr(stats, distName)
		params = dist.fit(data)
		
		pvalue = kstest(data, distName, args = params)[1]
		
		*shapeParams, loc, scale = params
		
		print('%s pvalue: %.4f\nparams: %s, loc: %.4f, scale: %.4f' % (distName, pvalue, shapeParams, loc, scale))
		
		# generate PDF of fitted distribution
		xstart = dist.ppf(0.01, *shapeParams, loc = loc, scale = scale)
		xend = dist.ppf(0.99, *shapeParams, loc = loc, scale = scale)
		xend = min(xend, data.max()*1.2)
		
		xs = np.linspace(xstart, xend, 1000)
		PDF = dist.pdf(xs, *params[:-2], loc = params[-2], scale = params[-1])
		fitPDFs[distName] = pd.Series(PDF, index = xs)
	
	return fitPDFs
	

def plot_results(out_file, data, data_label, fitted):
	'''
	Parameters
	out_file: str, output file
	data: ser of data
	data_label: str, data label for xaxes
	fitted: dict of ser, keys are distribution names, values are values
	'''
	
	plt.hist(data, bins = 50 if data.size > 100 else 10)
	plt.xlabel(data_label)
	plt.ylabel('Count')

	ax = plt.twinx()
	for distName, PDF in fitted.items():
		ax.plot(PDF.index, PDF.values, label = distName)

	ax.set_ylabel('Probability density function')
	ax.legend()

	plt.savefig(out_file, dpi = 300, bbox_inches = 'tight')



	
if __name__ == '__main__':
	
	data = read_data(DATA_FILE)
	
	fittedPDFs = identify_distribution(data)
	
	plot_results(OUT_FILE, data, DATA_LABEL, fittedPDFs)
	
	
	
	
	
	
	
	
	
	
	
	
		
	
	
	
	
