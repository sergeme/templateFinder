# templateFinder
a small, old c# tool to fix old word files that take forever to open in word.

## what's this?
- very quick and even dirtier
- sometimes, the "attached template" property of a word file can cause a problem, if it's set as a UNC path to a server no longer available.
- this tool crawls through a specified folder structure and looks for all word files present.
- lists all the found documents and their attached template property.
- by defining a new value for the "attached template" property, and a substring to be searched for (for instance \\oldserver) you can replace the property for all affected files.
- allows to export a report for all crawled files
- generates some log files, since some documents cause problems / are password protected etc.
- tested on win7 with office 2010 & win10 with office 2016
- probably crashes all the time
