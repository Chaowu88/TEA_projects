#!/usr/bin/env pyhton
# -*- coding: UTF-8 -*-


__author__ = 'Chao Wu'
__date__ = '05/22/2021'
__version__ = '1.0'


r'''
This script generates training dataset using Aspen model and .xslm calculator.

Python C:\Users\cwu\Desktop\Software\Aspen_automation\Scripts\case\FY21_Q3\generate_dataset.py
'''


DATASET_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\new_model\training_data.xlsx'
#r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\new_model\test\training_data.xlsx'
#r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\new_model\training_data.xlsx'
ASPEN_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Data\FY2021_Q3\Feedstock_Cost\BC_FY30MYPP_BDO_Combined_modified_lignin_0.51_scale_final.bkp'
#r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\new_model\test\PETase_Depoly_Base_v17.bkp'
#r'C:\Users\cwu\Desktop\Software\Aspen_automation\Data\FY2021_Q3\Feedstock_Cost\BC_FY30MYPP_BDO_Combined_modified_lignin_0.51_scale_final.bkp'
CALCULATOR_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Data\FY2021_Q3\Feedstock_Cost\BC_BDO_FY30Projection.xlsm'
#r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\FY2021_Q3\new_model\test\PETase_Depoly_Base_v17_corr.xlsm'
#r'C:\Users\cwu\Desktop\Software\Aspen_automation\Data\FY2021_Q3\Feedstock_Cost\BC_BDO_FY30Projection.xlsm'
NRUNS = 3
SCRIPTS_DIR = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Scripts'


import sys
sys.path.append(SCRIPTS_DIR)
import os
from collections import namedtuple
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from classes import Aspen, Excel


def parse_data_file(data_file):
	'''
	Parameters
	data_file: str, path of data file
	
	Returns
	inputInfos, outputInfo: df
	'''
	
	dataInfo = pd.read_excel(data_file, sheet_name = ['Inputs', 'Output'])
	inputInfo = dataInfo['Inputs']
	outputInfo = dataInfo['Output']
	
	return inputInfo, outputInfo
	
	
def run_and_update(data_file, input_infos, output_info, aspen_file, calculator_file, nruns):	
	'''
	Parameters
	data_file: str, dataset file
	input_infos: df, columns are ['Input variable', 'Type', 'Location', 'Values']
	output_info: df, columns are ['Output variable', 'Location', 'Values']
	aspen_file: str, Aspen model file
	calculator_file: .xslm calculator file
	nruns: int, total # of runs
	'''
	
	*others, values = output_info.squeeze()
	if isinstance(values, str):
		values = list(map(float, values.split(',')))
	elif np.isnan(values):
		values = []
	else:
		raise TypeError("what's in the Values column of Output sheet?")
		
	OutputInfo = namedtuple('OutputInfo', ['name', 'loc', 'values'])
	outputInfo = OutputInfo(*others, values)
	
	
	nrunsCompl = len(outputInfo.values)
	nrunsLeft = nruns - nrunsCompl
	if nrunsLeft >= 0:
		print('%s runs left.' % nrunsLeft)
	else:
		print('detected runs exceeds the number of required.')
	
	if nrunsLeft != 0:
		
		# remake input_infos
		InputInfo = namedtuple('InputInfo', ['name', 'type', 'loc', 'values'])
		inputInfos = []
		for _, [*others, values] in input_infos.iterrows():
			
			values = list(map(float, values.split(',')))
			inputInfos.append(InputInfo(*others, values))
		
		# run
		outDir = os.path.dirname(data_file)
		tmpDir = outDir + '/tmp'
		os.makedirs(tmpDir, exist_ok = True)
	
		aspenModel = Aspen(aspen_file)
		calculator = Excel(calculator_file)

		for i in range(nrunsCompl, nruns):
			print('run %s:' % (i+1))
			
			# set Aspen variables
			for inputInfo in inputInfos:
				if inputInfo.type == 'bkp':
					aspenModel.set_value(inputInfo.loc, inputInfo.values[i], False)
					
				elif inputInfo.type == 'bkp_fortran':
					aspenModel.set_value(inputInfo.loc, inputInfo.values[i], True)
				
				else:
					continue

			# run Aspen model
			aspenModel.run_model()
			
			tmpFile = '%s/%s.bkp' % (tmpDir, i)
			aspenModel.save_model(tmpFile)
			
			# set calculator variables
			for inputInfo in inputInfos:
				if inputInfo.type == 'xlsm':
					inputSheet, inputCell = inputInfo.loc.split('!')
					calculator.set_cell(inputInfo.values[i], inputSheet, loc = inputCell)
				
				else:
					continue
			
			# run calculator
			calculator.load_aspenModel(tmpFile)
			calculator.run_macro('solvedcfror')
			
			outputSheet, outputCell = outputInfo.loc.split('!')
			output = calculator.get_cell(outputSheet, loc = outputCell)
			outputInfo.values.append(output)
			
			# update dataset
			outputValues = ','.join(map(str, outputInfo.values))
			output_info = pd.DataFrame([[outputInfo.name, outputInfo.loc, outputValues]],
									  columns = ['Output variable', 'Location', 'Values'])
			
			with pd.ExcelWriter(data_file) as writer:
				input_infos.to_excel(writer, sheet_name = 'Inputs', index = False)
				output_info.to_excel(writer, sheet_name = 'Output', index = False)
				writer.save()
			
			print('done.')
			
		aspenModel.close()
		calculator.close()
		
	print('all done.')
	
	
	
	
if __name__ == '__main__':
	
	inputsInfo, outputInfo = parse_data_file(DATASET_FILE)
	
	run_and_update(DATASET_FILE, inputsInfo, outputInfo, ASPEN_FILE, CALCULATOR_FILE, NRUNS)
	
	
	
	
	
	

