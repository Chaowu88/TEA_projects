#!/usr/bin/env pyhton
# -*- coding: UTF-8 -*-


__author__ = 'Chao Wu'
__date__ = '06/02/2021'
__version__ = '1.0'


r'''
This script predicts output MSP using a trained regression model and performs:
1. Sensitivity analysis with one query input;
2. Response analysis with two query inputs;
3. Monte Carlo simulation with three or more query inputs 

python C:\Users\cwu\Desktop\Software\Aspen_automation\Scripts\case\FY21_Q3\predict_and_simulate.py
'''


OUT_DIR = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\new_model'
CONFIG_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\new_model\config.xlsx'
MODEL_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\new_model\regression.mod'
XLABEL = 'MFSP ($/GGE)'


import os
import re
from itertools import product
from collections import namedtuple
import numpy as np
import pandas as pd
from scipy import stats
from joblib import load
import matplotlib.pyplot as plt
import seaborn as sns


class BaseHandler:
	
	def __init__(self, config, baseline):
		'''
		Parameters
		config: df, columns are ['Input variable', 'Bounds', 'Distribution', 'Parameters', 'Size']
		baseline: df, columns are ['Input variable', 'Baseline value']
		'''
		
		self.config = config
		self.baseline = baseline
		
		
	@staticmethod
	def generate_random_values(dist_name, size, bounds, *params):
		'''
		Parameters
		dist_name: str, distribution name
		size: int, # of random values to generate
		bounds: tuple, (lower bound, upper bound)
		params: tuple, parameters of dist_name
		
		Returns
		values: array
		'''
		
		dist = getattr(stats, dist_name)
		lb, ub = bounds
		
		if dist_name == 'uniform':
			values = dist.rvs(loc = lb, scale = ub-lb, size = size)

		elif dist_name == 'bernoulli':
			pl, ph = params
			labels = dist.rvs(pl, size = size)
			values = [lb if label else ub for label in labels]
		
		else:
			*shapeParams, loc, scale = params
			
			values = []
			count = 0
			while count < size:
				value = dist.rvs(*shapeParams, loc = loc, scale = scale)
				if lb <= value <= ub:
					count += 1
					values.append(value)
		values = np.array(values)			
					
		return values

	
	def load_model(self, model_file):
		'''
		Parameters
		model_file: str, model file
		'''
		
		self.model = load(model_file)
		
		
	def simulate(self):
		
		self.outputs = []
		Output = namedtuple('Output', ['name', 'values'])
		for singleInput in self.inputs:
			
			predicted = self.model.predict(singleInput.data)
			
			singleOutput = Output(singleInput.name, predicted)
			self.outputs.append(singleOutput)
			
	
	def plot_hist_and_save(self, out_dir, folder_name, xlabel, percentile = 10):
		'''
		Parameters
		out_dir: str, output directory
		folder_name: str, folder name
		xlabel: str, label of xaxis
		percentile: float of 0 - 100, lines indicate percentile% and 1 - percentile% will be plotted
		'''
		
		for singleOutput in self.outputs:
			
			varName = get_var_name(singleOutput.name)
			values = singleOutput.values
			fileName = get_var_name(xlabel)
			
			saveDir = '%s/%s/%s' % (out_dir, folder_name, varName)
			os.makedirs(saveDir, exist_ok = True)
			
			fig, ax1 = plt.subplots()
			sns.distplot(values, rug = True, kde = False, hist = True, ax = ax1)
			ax1.set_xlabel(xlabel, fontsize = 15)
			ax1.set_ylabel('Count', color = 'steelblue', fontsize = 15)
			
			ax2 = ax1.twinx()
			sns.distplot(values, rug = True, kde = True, hist = False, ax = ax2)   # if plot kde, y axis can not be Count
			ax2.set_yticks([])
			ax2.spines['left'].set_visible(False)
			ax2.spines['right'].set_visible(False)
			ax2.spines['top'].set_visible(False)
			ax2.spines['bottom'].set_visible(False)
			
			counts, edges = np.histogram(values, bins = int(values.size/10) if values.size > 20 else values.size)
			x = (edges[:-1]+edges[1:])/2
			y = np.cumsum(counts)/np.sum(counts)
			p1, p2 = np.percentile(values, [percentile, 100-percentile])
			
			ax3 = ax1.twinx()
			ax3.plot(x, y, color = 'seagreen')
			ax3.set_ylabel('Cumulative probabilty', color = 'seagreen', fontsize = 15)
			ax3.vlines(x = p1, ymin = 0, ymax = 1, linestyles = 'dashed', color = 'gray')
			ax3.vlines(x = p2, ymin = 0, ymax = 1, linestyles = 'dashed', color = 'gray')
			
			fig.savefig('%s/%s.jpg' % (saveDir, fileName), dpi = 300, bbox_inches = 'tight')
			
			pd.Series(values).to_excel('%s/%s.xlsx' % (saveDir, fileName), header = False, index = False)
	
	
	def plot_contour_and_save(self, out_dir, folder_name, xlabel):
		'''
		Parameters
		out_dir: str, output directory
		folder_name: str, folder name
		xlabel: str, label of xaxis
		'''
		
		for singleInput, singleOutput in zip(self.inputs, self.outputs):
			
			varName = get_var_name(singleOutput.name)
			fileName = get_var_name(xlabel)
			
			saveDir = '%s/%s/%s' % (out_dir, folder_name, varName)
			os.makedirs(saveDir, exist_ok = True)
			
			xvar, yvar = singleInput.name.split('_')
			
			X, Y = np.meshgrid(singleInput.xvalues, singleInput.yvalues)
			Z = singleOutput.values.reshape(singleInput.xvalues.size, singleInput.yvalues.size).T
			
			fig, ax = plt.subplots()
	
			ctf = ax.contourf(X, Y, Z, 50, cmap = plt.cm.get_cmap('RdBu').reversed())
			
			ax.set_xlabel(xvar, fontsize = 15)
			ax.set_ylabel(yvar, fontsize = 15)
			
			cbar = fig.colorbar(ctf)
			cbar.set_label(xlabel, labelpad = 15, rotation = 270, fontsize = 15)
			
			if Z.max() - Z.min() > 0.001:	
				ct = ax.contour(X, Y, Z, ctf.levels[1::6], colors = 'dimgray', linewidths = 1, linestyles ='dashed')
				ax.clabel(ct, ctf.levels[1::6], inline = True, fontsize = 7, colors = 'k')
			
			fig.savefig('%s/%s.jpg' % (saveDir, fileName), dpi = 300, bbox_inches = 'tight')
			
			pd.DataFrame(Z).to_excel('%s/%s.xlsx' % (saveDir, fileName), header = xvar, index = yvar)
		

class OneInputHandler(BaseHandler):
	
	def generate_input_matrix(self):
		
		self.inputs = []
		Input = namedtuple('Input', ['name', 'data'])
		for _, [inputVar, bnds, distName, params, size] in self.config.iterrows():
			
			bnds = map(float, bnds.split(','))
			if isinstance(params, str):
				params = map(float, params.split(','))
			elif np.isnan(params):
				params = ()
			size = int(size)
			values = self.generate_random_values(distName, size, bnds, *params)
			
			baseInput = self.baseline.set_index('Input variable').T
			baseInputMat = pd.concat([baseInput]*size, ignore_index = True)
			
			baseInputMat[inputVar] = values
			
			singleInput = Input(inputVar, baseInputMat.values)
			self.inputs.append(singleInput)
			
	
	def plot_and_save(self, out_dir, xlabel):
		'''
		Parameters
		out_dir: str, output directory
		xlabel: str, label of xaxis
		'''
		
		self.plot_hist_and_save(out_dir, 'one_input', xlabel)
		

class TwoInputsHandler(BaseHandler):
	
	def generate_input_matrix(self):
		
		self.inputs = []
		Input = namedtuple('Input', ['name', 'xvalues', 'yvalues', 'data'])
		for _, [inputVars, bndss, sizes] in self.config.iterrows():
			
			inputVarX, inputVarY = inputVars.split('|')
			bndsX, bndsY = (tuple(map(float, bnds.split(','))) for bnds in bndss.split('|'))
			sizeX, sizeY = map(int, sizes.split('|'))
			
			valuesX = np.linspace(*bndsX, sizeX)
			valuesY = np.linspace(*bndsY, sizeY)
			
			baseInput = self.baseline.set_index('Input variable').T
			baseInputMat = pd.concat([baseInput]*sizeX*sizeY, ignore_index = True)
			
			baseInputMat[[inputVarX, inputVarY]] = list(product(valuesX, valuesY))
			
			singleInput = Input(inputVarX+'_'+inputVarY, valuesX, valuesY, baseInputMat.values)
			self.inputs.append(singleInput)
		
	
	def plot_and_save(self, out_dir, xlabel):
		'''
		Parameters
		out_dir: str, output directory
		xlabel: str, label of xaxis
		'''
		
		self.plot_contour_and_save(out_dir, 'two_input', xlabel)
	
	
class MoreInputsHandler(BaseHandler):
	
	def generate_input_matrix(self):
			
		self.inputs = []
		Input = namedtuple('Input', ['name', 'data'])
		for _, [inputVars, bndss, distNames, paramss, size] in self.config.iterrows():
		
			baseInput = self.baseline.set_index('Input variable').T
			baseInputMat = pd.concat([baseInput]*size, ignore_index = True)
			
			inputVars = inputVars.split('|')
			bndss = bndss.split('|')
			distNames = distNames.split('|')
			paramss = paramss.split('|')
			size = int(size)
			for inputVar, bnds, distName, params in zip(inputVars, bndss, distNames, paramss):
			
				bnds = map(float, bnds.split(','))
				if params == '':
					params = ()
				else:
					params = map(float, params.split(','))
				values = self.generate_random_values(distName, size, bnds, *params)
				
				baseInputMat[inputVar] = values
			
			singleInput = Input('_'.join(inputVars), baseInputMat.values)
			self.inputs.append(singleInput)
				
	
	def plot_and_save(self, out_dir, xlabel):
		'''
		Parameters
		out_dir: str, output directory
		xlabel: str, label of xaxis
		'''
		
		self.plot_hist_and_save(out_dir, 'more_input', xlabel)
		

def parse_config_file(config_file):
	'''
	Parameters
	config_file: str, path of config file. Note the order of variables in sheet Baseline should identical 
				 with the order of model features
	
	Returns
	oneInput, twoInputs, moreInputs, baseline: df
	'''
	
	configInfo = pd.read_excel(config_file, sheet_name = ['One-input', 'Two-inputs', 'More-inputs', 'Baseline'])
	oneInput = configInfo['One-input']
	twoInputs = configInfo['Two-inputs']
	moreInputs = configInfo['More-inputs']
	baseline = configInfo['Baseline']
	
	return oneInput, twoInputs, moreInputs, baseline
	
	
def get_var_name(name_with_unit):		
	'''
	Parameters
	name_with_unit: str, variable name with unit
	
	Returns
	name: str
	'''
	
	return re.sub(r'\s*\(.*?\)\s*', '', name_with_unit)
	
	
	

if __name__ == '__main__':
	
	*configs, baseline = parse_config_file(CONFIG_FILE)
	Handlers = [OneInputHandler, TwoInputsHandler, MoreInputsHandler]
	
	for config, Handler in zip(configs, Handlers):
		if not config.empty:
			handler = Handler(config, baseline)
			handler.generate_input_matrix()
			handler.load_model(MODEL_FILE)
			handler.simulate()
			handler.plot_and_save(OUT_DIR, XLABEL)
			
	
	
	
	
	
	
	
	
	
	
	
	
	


