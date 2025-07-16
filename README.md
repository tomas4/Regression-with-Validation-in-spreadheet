# Regression-with-Validation-in-spreadheet
Spreadsheet file with embedded Python script for easy regression with cross-validation.

![screenshot v. 0.2a](https://github.com/tomas4/Regression-with-Validation-in-spreadheet/blob/main/screenshot-v0.2a.png)

This project aims to provide a simple way to perform cross-validation of regression models in a freely downloadable spreadsheet document. All the user has to do is insert their data into the spreadsheet, set the parameters and perform the calculation. The calculations are performed in Python. LibreOffice is initially used as the spreadsheet programme, but an Excel version is also planned.

This development version allows for simple or multiple linear regression (linear fitting) with K-fold cross-validation with selectable k. LOOCV cross-validation, polynomial and other non-linear fits are planned for the future, as well as producing more detailed and streamlined outputs in the report document.

The underlying Python script is also uploaded separately to the repository, but does not need to be downloaded to use the programme. All you need is the worksheet file _Regression with validation.ods_ in which the script is embedded. The user also needs to install the script's dependencies, that is Python with numpy, scikit-learn and pandas libraries. For users currently not using LibreOffice or OpenOffice and not wiling or able to install these products (for example in corporate environment with limited rights) I can recommend [PortableApps project](https://portableapps.com) and [portable LibreOffice](https://portableapps.com/apps/office/libreoffice_portable).

## Future development plans
* Dependencies instalation automation (currently working on it)
* Non-linear fitting
* Excel version
