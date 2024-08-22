# Excel Print to PDF
This started as a Visual Basic applications script and later turned into a Windows Forms application using Visual Basic .NET.

It prints all Excel files in a specific folder to PDF. Two option exist:
- print all sheets in the file to PDF, or
- only print the first sheet to PDF.

## Prerequisites
To use this app the following is required, which is not supplied in this repo:
- Microsoft Excel for Windows
- Microsoft Windows Desktop Runtime (x86), the current version used is `8.0.8`.

## Installing
Other than the prerequisites above, nothing else needs to be installed.

# Usage

When the app is executed for the first time a `filesToPrint` folder will be created in the `root` folder where the executable is. All Excel files which need to be printed to PDF are to be placed in this folder.

The UI window will present the user with two radio buttons:
- Print all sheets
- Print first sheet

Select one of these options and click on the `Print` button. All Excel files in the `filesToPrint` folder will be printed to PDF and saved into the same folder.

# License
This project is licensed under the MIT license - see the `LICENSE` file for details.
