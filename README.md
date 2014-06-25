# WHA Revit Launcher #

This utility automates the creation and launching of Revit "Local" files to get you working faster

## Features: ##
- Lists all active Revit Projects for the office
- Automatically creates a "local file" on your local computer in the "MyDocuments/Revit" folder
- Launches your project with the correct version of Revit
- Saves a back-up of the last two "local files" created for a particular project
- Launches the Worksharing Monitor to show who else is in the project
- Allows for easy "detaching" of a model

## How It Works: ##

The utility points to a file located on the network.  This file has a list of all active Revit
"Central" files for the company.  The file also includes additional information about the
project including:

* The project name and number
* the version of Revit the project is supporting 
* The working directory of the project 

When a project is chosen to be launched, the Revit Laucher checks this file for the correct settings.
 It will then create a backup of any old local files and a new local file is created from the latest
central file.  The new local is then opened with the correct version of the Revit software.