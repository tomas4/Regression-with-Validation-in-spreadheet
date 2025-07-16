# Regression-with-Validation-in-spreadheet
Spreadsheet file with embedded Python script for easy regression with cross-validation.

![screenshot v. 0.2a](https://github.com/tomas4/Regression-with-Validation-in-spreadheet/blob/main/screenshot-v0.2a.png)

This project aims to provide a simple way to perform cross-validation of regression models in a freely downloadable spreadsheet document. All the user has to do is insert their data into the spreadsheet, set the parameters and perform the calculation. The calculations are performed in Python. LibreOffice is initially used as the spreadsheet, but an Excel version is also planned.

This development version allows for simple or multiple linear regression (linear fitting) with K-fold cross-validation with selectable k. LOOCV cross-validation, polynomial and other non-linear fits are planned for the future, as well as producing more detailed and streamlined outputs in the report document.

## Installation of dependencies
You need a LibreOffice (LO) installation. For Windows users currently not using LibreOffice (or OpenOffice, but it was not tested by the author) and not willing or able to do standard install of one of these products (for example in a corporate environment with limited rights) I can recommend [PortableApps project](https://portableapps.com), and [portable LibreOffice](https://portableapps.com/apps/office/libreoffice_portable). For standard installation in all opreating systems I would recommend to use either your operating system's standard software repository (Microsoft Store on Windows, Apt system on Debian Linux, etc...), or package downloded from the official [LibreOffice site](https://www.libreoffice.org/). On Linux, there are some [problems with Python scripting in Appimage installs of LibreOoffice](https://duckduckgo.com/?t=midori&q=problems+with+Python+scripting+in+Appimage+installs+of+LibreOoffice&ia=web), at least with some versions.

The underlying Python script is uploaded separately to the repository, but does not need to be downloaded to use the app. All you need is the worksheet file _Regression with validation.ods_ in which the script is embedded already. The user needs to install the script's dependencies, however. That is Python modules [numpy](https://extensions.libreoffice.org/en/extensions/show/41995), scikit-learn and [pandas](https://extensions.libreoffice.org/en/extensions/show/41998) libraries. While it is possible to install numpy and pandas by the linked LO extensions, it is currently not possible for scikit-learn. The standard way to install Python modules into LO Python is using pip. The procedure is described in files [Installation-of-dependencies.md](https://github.com/tomas4/Regression-with-Validation-in-spreadheet/edit/main/Installation-of-dependencies.md) and [Installation-of-dependencies.odt](https://github.com/tomas4/Regression-with-Validation-in-spreadheet/edit/main/Installation-of-dependencies.odt). I work on automation of this process within the app.

## Future development plans
* Dependencies installation automation (currently working on it)
* Non-linear fitting
* Variable transformation
* Excel version
