# EBI (Equella Bulk Importer)
The EBI is a popular tool for importing content into EQUELLA. It can also be used for updating, deleting and exporting content. 

The EBI is written in Python and compiled to a standalone version (i.e. will run on a computer without Python) for Windows and Macintosh. The source files are included in the Windows package so that the EBI can be used on Linux computers with Python (and wxPython, see Dependencies) installed.

## Dependencies
The EBI requires both Python 2.7+ and the GUI framework wxPython. However, to use a standalone EBI package neither Python nor wxPython need be installed as they are both included in the standalone package. To run the EBI from source files (as required on Linux) both Python 2.7+ and wxPython must be installed.
To make modifications to and test EBI Python 2.7.x and wxPython must be installed on the developer’s workstation. To compile the EBI as a standalone package then, as well as Python 2.7.x and wxPython, one of the following is required on the workstation:
•	py2exe (for Windows), or
•	py2app (for Macintosh)

## Compiling/Packaging Standalone
EBI should be compiled as a standalone package for Windows and Macintosh to remove the need for end users to install Python. Compiling a Windows version must be done from a Windows computer and compiling a Macintosh version must be done from a Macintosh computer.

### Compiling a Windows Standalone Package
Make certain all source code files, package.bat, setup.py (for Windows) and package.py is in the same folder on a Windows computer with Python 2.7+, wxPython and py2exe installed.

Run package.bat to create ebi.zip (it will appear in the same folder as the source files).
package.bat does the following automatically:
1.	Removes any previous packages from the working folder
2.	 Invokes setup.py and py2exe to generate a standalone package in a folder called “dist” Dist, amongst other files, contains ebi.exe which is the resulting standalone Windows EBI executable.
3.	Renames the “dist” folder to “ebi”
4.	Copies the source files into the “ebi” folder
5.	Invokes package.py which zips the “ebi” folder to ebi.zip

### Compiling a Macintosh Standalone Package
Use a Macintosh computer with Python 2.7+, wxPython and py2app installed. Place package.command, setup.py (for Macintosh) and ebi.command in the same folder. In that folder create a sub folder called source and in source create a sub folder called ebi. Place all the EBI source files in /source/ebi.
Run package.command to create ebi.dmg (it will appear in a sub folder called dist).
package.command does the following:
1.	Removes any previous packages from the working folder
2.	Invokes setup.py and py2app to generate a standalone app package called “ebi.app” in a folder called “dist”
3.	Copies the image files and ebi.command into the “ebi.app” package
4.	Creates a DMG called ebi.dmg in the “dist” folder from ebi.app (by using hdituil) 



