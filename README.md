# Regulatory-Request-Helper


Overview: 

The Regulatory-Request-Helper is a python application which utilizes the PyQt4 GUI to aid in the automation of a tedious excel-based process for the Regulatory Department of Griffith Foods Inc.


Problem: 

Upwards of fifty sales requests (excel spreadsheets) are sent into the Regulatory Department of Griffith Foods every day.  The relevant information from each of these spreadsheets must be copied into a Master Database on a Network Drive.  Copying and Pasting the information from all fifty of these spreadsheets was getting tedious and incredibly time consuming.


Solution: 

The Regulatory-Request-Helper solves this problem by creating a user-friendly Windows Application that takes in all those requests at once and compiles their relevant information into one excel spreadsheet that can then be copied to the Master Database.  Now instead of fifty excel spreadsheets, one only has to run the program and copy information from one excel spreadsheet.

*NOTE* The Regulatory-Request-Helper does have the capability to completely bypass the middle, "copying" excel spreadsheet, by directly uploading data to the Master Database.  However, this had the possibility to cause syncing issues in the distant future.  Although, the code that executes this operation is still within the regualtory-helper.py 


How to Use and How it Works:  

At the end of the day, the person tasked with the copying and pasting of the sales requests only has to copy the paths of the sales request files and paste them into the Regulatory-Request-Helper’s editable textbox.  After clicking the “Convert to One” button, the program executes multiple strains of commands that open the individual files and look for the relevant information.  The program stores this information and then copies it to the new to_copy.xls.  Once the program is complete, to_copy.xls will appear in the same file location as the program itself.  The Excel file can be opened and all the information within it can then be copied to the Master Database. 

This application uses many encompassing Python modules, but namely PyQt4, xlrd, and xlwt to execute most of the aforementioned operations.  The program has both 32-bit and 64-bit versions which were created using separate versions of Python.  All the raw files are available within their folders, as well as their respective .exe files.  One may notice that the folders contain various photos.  These are to improve the overall graphics and presentation of the GUI. 
