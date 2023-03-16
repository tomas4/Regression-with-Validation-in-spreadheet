# Regression-with-Validation-in-spreadheet
Spreadsheet file with embedded Python script for easy regression with cross-validation.

![screenshot v. 0.1a](https://github.com/tomas4/Regression-with-Validation-in-spreadheet/blob/main/screenshot-v0.1a.png)

This project is intended to bring simple way to do cross-validation of regression models in a freely downloadable spreadsheet document. The user should only paste his/her data into sheet of this spreadseet document, set parameters and run the computation. The computaion is caried on in Python, the initial spreadheet application employed is LibreOffice, but Excel support is planned as well.

This initial version allows for simple or multiple linear regression (linear fit) with K-fold cross-validation with fixed k=10. User selected k for k-fold cross-validation and LOOCV cross-validation are the immediate development plans, as well as producing more detailed output in the form of a report document.

The underlying Python script is uploaded in the repository also separately, but it is not needed to download it for using the program. All what is needed is the spreadheet file _Regression with validation.ods_. User also has to install dependencies of the script, that is to have installed Python with numpy, scikit-learn and pandas libraries. 
