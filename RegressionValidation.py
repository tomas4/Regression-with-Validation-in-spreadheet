# RegressionValidation.py
# See https://events.documentfoundation.org/media/libocon2021/submissions/FCU9NB/resources/LibOCon_-_ScriptForge_2021-09_-_Rafael_Lima_DWkWAXc.pdf

import uno

from scriptforge import CreateScriptService
from sklearn.model_selection import KFold
from sklearn.model_selection import cross_val_score
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, r2_score
import numpy as np
import pandas as pd
# Get the current document
doc = CreateScriptService("Calc")
# Load the Basic service
bas = CreateScriptService("Basic")

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
	k = 10
	kf = KFold(n_splits=k, random_state=None)
	# For now CV results in msgbox
	# list of crossval MSE scores in folds
	CV_MSE = np.abs(cross_val_score(model , X, y, scoring = 'neg_mean_squared_error', cv = kf))
	# Average CV MSE
	AVCV_MSE = np.mean(CV_MSE)
	# Root of average CV MSE
	AVCV_RMSE = np.sqrt(AVCV_MSE)
	
	# Results text
	results = " REGRESSION RESULTS " + \
		"\n regression type: " + rtype +  \
		"\n number of dependent variables p=" + str(p) + \
		"\n number of observations n=" + str(n) + \
		"\n itercept a0=" + str(model.intercept_) + \
		"\n slopes [a1 a2 ..an]=" + str(model.coef_) + \
		"\n model data MSE=" + str(mean_squared_error(y, y_predict)) + \
		"\n model data RMSE=" + str(np.sqrt(mean_squared_error(y, y_predict))) + \
		"\n coef of determination R2=" + str(COD) + \
		"\n adjusted coef of determination R2adj=" + str(adjusted_COD) + \
		"\n\n CROSS-VALIDATION RESULTS" + \
		"\n Validated MSE scores in folds: " + str(CV_MSE) + \
		"\n Average validated MSE=" + str(AVCV_MSE) + \
		"\n Overall validated RMSE=" + str(AVCV_RMSE) + \
		"\n\n Note: Overall validated RMSE is root square of average validated MSE, not average of RMSE in folds."
	
	return results

def ButtonRun(*args ):
	"""Function to call ReadValues() when a button is pressed."""
	#Note: the *args in function declaration is needed, so that the function acceps arguments supplied by the button.
	
	if doc.IsCalc:
		# read data from selection into pandas dataframe
		df = ReadValues()
		# check for empty cells in the dataset, warn the user  and drop the affected rows 
		# replace empty strings with np.NaN
		df = df.replace('', np.nan)
		if df.isnull().values.any():
			# warning about the empty cells
			bas.MsgBox("Warning: The selected  range contains empty cells. Please note that the rows with empty cells will be dropped from the dataset prior regression and cross-validation.")
			# drop the rows with empty cells
			df.dropna(inplace=True)
		# proceed with regression and validation
		results = RegressionValidation(df)
		# create document with results report
		# msgbox for now
		bas.MsgBox(results)
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
	# Get the current document
	doc = CreateScriptService("Calc")
	# Load the Basic service
	bas = CreateScriptService("Basic")
	#warning
	confirmation = bas.MsgBox("This will ERASE all the data in the table and reset the edited variable names. Are you sure?", 4)
	if confirmation == 6 :
		doc.ClearValues("~.A4:F1000000")
		theCaptions = [ "symbol_y [unit]", "symbol_x1 [unit]", "symbol_x2 [unit]", "symbol_x3 [unit]", "symbol_x4 [unit]", "symbol_x5 [unit]"]
		doc.SetValue("A3:F3", theCaptions)
	return None
		