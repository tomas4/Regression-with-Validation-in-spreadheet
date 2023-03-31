# RegressionValidation.py
# Script to do (multiple) linear regression with crossvalidation within Calc spreadsheet "Regression with validation.ods"

import uno

from scriptforge import CreateScriptService
from sklearn.model_selection import KFold
from sklearn.model_selection import cross_val_score
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, r2_score
import numpy as np
import pandas as pd
# For future use
#import seaborn as sns
#import matplotlib.pyplot as plt 

# Get the current document
doc = CreateScriptService("Calc")
# Load the Basic service
bas = CreateScriptService("Basic")
# Load the UI service
ui = CreateScriptService("UI")


def RegressionValidation(df):
	""" Does (multiple) linear regression with cross-validation, returns results as formatted text. """
	# Data preparation
	columns = df.columns
	y = df.pop(columns[0])
	X = df
	# Linear regression
	lm = LinearRegression()
	model = lm.fit(X, y)
	y_predict = model.predict(X)
	# Compupte adjusted coef of determination
	COD = model.score(X,y)
	dfshape = df.shape
	# n - number of samples
	n = dfshape[0]
	# p - number of dependent variables
	p = dfshape[1]
	adjusted_COD = 1 - (1 - COD) * (n-1) / (n-p-1)
	# Regression type (SLR/MLR)
	if p > 1 :
		rtype = "Multiple linear regression"
	else:
		rtype = "Simple linear regression"
	# Cross-validation
	k = ReadSettings("B4")
	kf = KFold(n_splits=k, random_state=None)
	# For now CV results in msgbox
	# list of crossval MSE and RMSE scores in folds
	CV_MSE = np.abs(cross_val_score(model , X, y, scoring = 'neg_mean_squared_error', cv = kf))
	# Average CV MSE and CV RMSE
	AVCV_MSE = np.mean(CV_MSE)
	AVCV_RMSE = np.sqrt(AVCV_MSE)
	
	# Results text
	global reg_results
	global cv_results
	reg_results = rtype + ":" +  \
		"\nnumber of dependent variables p=" + str(p) + \
		"\nnumber of observations n=" + str(n) + \
		"\nitercept a0=" + str(model.intercept_) + \
		"\nslopes [a1 a2 ..an]=" + str(model.coef_) + \
		"\ncoef of determination R2=" + str(COD) + \
		"\nadjusted coef of determination R2adj=" + str(adjusted_COD) 
	cv_results = "Validated MSE scores in folds: " + str(CV_MSE) + \
		"\nAverage validated MSE=" + str(AVCV_MSE) + \
		"\nOverall validated RMSE=" + str(AVCV_RMSE)
	# Return tuple with reg_results (string), cv_results (string), origData (dataframe). TODO: Saves scatterplot as image scatterplot.png (above). The latter only in case of SLR (TODO: or MLR with two variables).
	return None

def ButtonRun(*args ):
	"""Function to start when a button is pressed."""
	#Note: the *args in function declaration is needed, so that the function acceps unneeded arguments supplied by the button.
	
	if doc.IsCalc:
		# read data from selection into pandas dataframe
		df = ReadValues()
		# check for empty cells in the dataset, warn the user  and drop the affected rows 
		# replace empty strings with np.NaN
		df = df.replace('', np.nan)
		if df.isnull().values.any():
			# warning about the empty cells
			bas.MsgBox("Information: The selected  range contains empty cells. Please note that the rows with empty cells will be dropped from the dataset prior regression and cross-validation.")
			# drop the rows with empty cells
			df.dropna(inplace=True)
		# Save copy of the dataframe in global variable, making sure the data are as float numbers
		global cleanData
		cleanData = df.copy()
		cleanData = cleanData.astype(float)
		# proceed with regression and validation
		RegressionValidation(df)
		# create document with results report
		ResultsReport()
		# msgbox for now
		#bas.MsgBox(results[0] + "\n\nCross-validation:\n" + results[1])
	else:
		bas.MsgBox("Please use this script in Libreoffice Calc 7.2 and higher version only!")
	return None

def ReadValues():
	"""Function to read the input values and column headers from the sheet selection and return pandas dataframe"""
	# Should do:  Read column headinggs in A2, B2, read data from cells below.
	# Read selection with input data including column headers (in spreadsheet row 2 in the RegressionWithValidation.ods) containing y-data in first column, x1-data in second, and so on, if there is more independent x variables. NOTE: result is a tuple
	Selection = doc.GetValue("~.~")
	#Create Pandas dataframe with named columns from the selection - forst row of selection makes column names, the rest the y, x1, x2, ... etc. data columns.
	# first split the header row and the data to arrays
	ColHeaders = np.array(Selection[0])
	TheData = np.array(Selection[1:-1])
	df = pd.DataFrame(TheData,  columns=ColHeaders)
	return df
	
def ButtonClearTable(*args):
	"""Function to clear the whole data area of the table and reset value names"""
	#warning
	confirmation = bas.MsgBox("This will ERASE all the data in the table and reset the edited variable names. Are you sure?", 4)
	if confirmation == 6 :
		doc.ClearValues("~.A4:F1000000")
		theCaptions = [ "symbol_y [unit]", "symbol_x1 [unit]", "symbol_x2 [unit]", "symbol_x3 [unit]", "symbol_x4 [unit]", "symbol_x5 [unit]"]
		doc.SetValue("A3:F3", theCaptions)
	return None
	
def ReadSettings(cell):
	""" Function to read settings from the Settings sheet at the location (like B4, i.e. without sheet name) specified as input parameter """
	cell = "$Settings." + cell
	value = doc.GetValue(cell)
	return value
	
def ButtonChangeSettingsSMLR(*args):
	""" Function to change the SMLR settings """
	# TO BE IMPLEMENTED
	# For now:
	bas.MsgBox("Not yet implemented. For now you can change number of folds in the sheet Settings directly (cell B4)")
	return None

def ResultsReport():
	""" Function to open results report as new Calc document"""
	# Create and open new Writer tocument
	# msgbox for now
	# Open report Writer document
	reportDoc = ui.CreateDocument("Calc")
	# Copy regression results title template to A1
	source_range = doc.Range("Templates.A1")
	reportDoc.copyToCell(source_range, "A1")
	# Paste the results of regression to A2
	RegrResultsArray = reg_results.split("\n")
	reportDoc.setArray("A2", RegrResultsArray)
	# Copy CV results title template to A10
	source_range = doc.Range("Templates.A2")
	reportDoc.copyToCell(source_range, "A10")
	# Paste the results of CV to A11
	CVResultsArray = cv_results.split("\n")
	reportDoc.setArray("A11", CVResultsArray)
	# Copy input data title template to F1
	source_range = doc.Range("Templates.A3")
	reportDoc.copyToCell(source_range, "F1")
	# Paste input data from F2 by columns
	# Here we increment the column address letter by 1 in a loop over column names to get right spreadsheet column letter for every name
	colLetter = "F"
	for (ColName, ColData) in cleanData.iteritems():
		# Write column name
		#bas.MsgBox("ColName type: " + str(type(ColName)) + "\nColData type: " + str(type(ColData)))
		reportDoc.setValue(colLetter + "2", ColName)
		reportDoc.setArray(colLetter + "3", ColData.tolist())
		n = ord(colLetter) + 1
		colLetter = chr(n)
	return None