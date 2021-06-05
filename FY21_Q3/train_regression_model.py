#!/usr/bin/env pyhton
# -*- coding: UTF-8 -*-


__author__ = 'Chao Wu'
__date__ = '05/25/2021'
__version__ = '1.0'


r'''
This script trains and tunes a ridge regression model using polynomial kernel.

python C:\Users\cwu\Desktop\Software\Aspen_automation\Scripts\case\FY21_Q3\train_regression_model.py
'''


OUT_DIR = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\new_model'
DATA_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\new_model\training_data.xlsx'


import sys
import warnings
import os
if not sys.warnoptions:   
	warnings.simplefilter('ignore')
	os.environ['PYTHONWARNINGS'] = 'ignore'
import numpy as np
import pandas as pd
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import PolynomialFeatures
from sklearn.linear_model import Ridge
from sklearn.model_selection import GridSearchCV
from scipy.stats import pearsonr
import matplotlib.pyplot as plt
from joblib import dump


def read_data(data_file):
	'''
	Parameters
	data_file: str, data file
	
	Returns
	features: df
	targets: ser
	'''
	
	dataInfo = pd.read_excel(data_file, sheet_name = ['Inputs', 'Output'])
	inputInfo = dataInfo['Inputs']
	outputInfo = dataInfo['Output'].squeeze()
	
	inputValues = inputInfo['Values'].str.split(',')
	features = pd.DataFrame(dict(zip(inputInfo['Input variable'], inputValues)), dtype = float)
	
	outputValues = outputInfo['Values'].split(',')
	targets = pd.Series(outputValues, name = outputInfo['Output variable'], dtype = float)
	
	return features, targets
	
	
def train_and_turn(features, targets, nfolds = 5, njobs = 3):
	'''
	Parameters
	features: df, training features
	targets: ser, training targets
	nfolds: int, # of cross validation folds
	njobs: int, # of jobs in parallel
	
	Returns
	bestModel: model
	bestParams: dict
	true_vs_pred: df
	R2: float
	'''
	
	pipe = Pipeline(steps = [('poly', PolynomialFeatures()), ('ridge', Ridge())])
	paramGrid = {'poly__degree': [1, 2, 3, 4, 5],
				 'ridge__alpha': [0.1, 1, 5, 10],
				 'ridge__fit_intercept': [True, False],
				 'ridge__normalize': [True, False]}
	
	regModels = GridSearchCV(pipe, paramGrid, cv = nfolds, n_jobs = njobs)
	regModels.fit(features, targets)
	
	bestModel = regModels.best_estimator_
	bestParams = regModels.best_params_
	
	predicted = bestModel.predict(features)
	R2 = pearsonr(predicted, targets)[0]**2
	
	true_vs_pred = pd.DataFrame({'True': targets.values, 'Predicted': predicted})
	
	return bestModel, bestParams, true_vs_pred, R2
	
	
def display_results(best_params, true_vs_pred, r2):
	'''
	Parameters
	best_params: dict, keys are params, values are values
	true_vs_pred: df, columns are ['True', 'Predicted']
	r2: float, square of correlation coefficient
	'''
	
	print('best parameters:')
	for param, value in best_params.items():
		print('%s = %s' % (param, value))
	print('')
		
	print('predicted vs true')
	for _, (vpred, vtrue) in true_vs_pred.iterrows():
		print(round(vpred, 4), round(vtrue, 4))
	print('')
	
	print('R2: %.4f' % r2)
	
	
def save_results(out_dir, model, true_vs_pred):
	'''
	Parameters
	out_dir: str, output directory
	model: trained model
	true_vs_pred: df, columns are ['True', 'Predicted']
	'''
	
	dump(model, '%s/regression.mod' % out_dir)
	
	true_vs_pred.to_excel('%s/true_vs_predicted.xlsx' % out_dir, header = True, index = False)


def plot_true_vs_predicted(out_dir, true_vs_pred, r2):
	'''
	Parameters
	out_dir: str, output directory
	true_vs_pred: df, columns are ['True', 'Predicted']
	r2: float, square of correlation coefficient
	'''
	
	fig, ax = plt.subplots()
	ax.scatter(true_vs_pred['True'], true_vs_pred['Predicted'])
	ax.set_xlabel('True values', fontsize = 15)
	ax.set_ylabel('Predicted values', fontsize = 15)
	ax.text(0.1, 0.8, '$R^2$ = %.3f' % r2, fontsize = 15, transform = ax.transAxes)
	
	lineLB, lineUB = true_vs_pred.min().min(), true_vs_pred.max().max()
	ax.plot([lineLB, lineUB], [lineLB, lineUB], linestyle = '--', color = '#d62728')

	fig.savefig('%s/true_vs_predicted.jpg' % out_dir, dpi = 300, bbox_inches = 'tight')




if __name__ == '__main__':
	
	features, targets = read_data(DATA_FILE)
	
	bestModel, bestParams, trueVSpred, R2 = train_and_turn(features, targets)
	
	display_results(bestParams, trueVSpred, R2)
	save_results(OUT_DIR, bestModel, trueVSpred)
	plot_true_vs_predicted(OUT_DIR, trueVSpred, R2)
	
	
	
	
	
	
	
	
	

