#!/usr/bin/env pyhton
# -*- coding: UTF-8 -*-


__author__ = 'Chao Wu'
__date__ = '08/22/2021'
__version__ = '1.0'


r'''
This script fitts the distribution of a continuous variable using kernel density estimation, and optionally generates random samples from the model

python C:\Users\cwu\Desktop\Software\Aspen_automation\Scripts\case\FY21_Q3\identify_distribution_kde.py
'''


OUT_DIR = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\Identify_distribution2\Weekly_LCFS'
DATA_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\Identify_distribution2\Weekly_LCFS\Weekly_LCFS.xlsx'
DATA_LABEL = 'Weekly values'
BANDWIDTH = 2

NEED_RANDOM_VALUES = True
SIZE = 5000   # works if NEED_RANDOM_VALUES == True


import os
import numpy as np
import pandas as pd
from sklearn.neighbors import KernelDensity
import matplotlib.pyplot as plt


def read_data(data_file):
	'''
	Parameters
	data_file: str, data file
	
	Returns
	data: ser
	'''
	
	data = pd.read_excel(data_file, header = 0, index_col = 0, squeeze = True)
	
	return data
	
	
def fit_distribution(data, bandwidth = 1):
	'''
	Parameters
	data: ser of data
	bandwidth: bandwidth of the kernel, large bandwidth leads to smooth density distribution
	'''
	
	kde = KernelDensity(kernel = 'gaussian', bandwidth = bandwidth).fit(data[:, np.newaxis])
	
	return kde
	

def plot_results(out_dir, data, data_label, model):
	'''
	Parameters
	out_dir: str, output directory
	data: ser of data
	data_label: str, data label for x axix
	fit_infos: list of namedtuples, fields are ['dist_name', 'shape_params', 'loc', 'scale', 'pvalue', 'pdf']
	'''
	
	os.makedirs(out_dir, exist_ok = True)
	
	plt.hist(data, bins = 50 if data.size > 100 else 10, density = True)
	plt.xlabel(data_label)
	plt.ylabel('Frequency (probability)')

	x = np.sort(data)
	prob = np.exp(model.score_samples(x[:, np.newaxis]))
	plt.plot(x, prob)
	
	plt.savefig('%s/fitted_distribution_kde.jpg' % out_dir, dpi = 300, bbox_inches = 'tight')
	
	
def generate_random_values(out_dir, model, size):
	'''
	Parameters
	out_dir: str, output directory
	model: KDE model
	size: int, # of random values to generate
	'''
	
	os.makedirs(out_dir, exist_ok = True)
	
	values = model.sample(size).reshape(-1)
	
	pd.Series(values).to_excel('%s/random_values.xlsx' % out_dir, header = False, index = False)
	



if __name__ == '__main__':
	
	data = read_data(DATA_FILE)
	
	kde = fit_distribution(data, bandwidth = BANDWIDTH)
	
	plot_results(OUT_DIR, data, DATA_LABEL, kde)
	
	if NEED_RANDOM_VALUES:
		generate_random_values(OUT_DIR, kde, SIZE)
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	



	
	
	

