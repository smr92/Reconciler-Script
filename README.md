# Reconciler-Script

A python script that reconciles a balance sheet account run from an accounting software using the openpyxl library

Python BTC Reconciler: Readme

###Introduction

This script attempts to reduce the amount manual work required to identify and remove the offsetting transactions in a report for a balance sheet account that one has run from an accounting software. The program is limited to seeking and identifying exact matches. To illustrate:
	
  John Doe . . . 80
	
  John Doe . . .(80)
	
  John Doe . . . 50
	
  John Doe . . .(25)
	
  John Doe . . .(25)

In this example, the program will be able to remove the 80 and (80), as they match each other exactly. However, even though the 50 and (25) (25) are transactions that offset each other, the program will overlook these.

That being said, in the Excel file with which testing was performed, a real world example, the computer is capable of reducing the number of rows from 1270 to 580. It has been estimated, then, that the program is capable of performing roughly 80% of the work, which is no insignificant time-saver.

###Instructions

  1. Prepare the excel file (N.B. See an Excel file titled "Example" in this repository for reference:
    - The script will read the first worksheet in the workbook. Make sure, therefore, that the data to be reconciled is so.
    - The transaction description must be in column A. The amounts must be in column B. Omit column headings.
    - Close out of the file.
  2. Run script:
    - The program will immediately prompt the user to enter file path. The best way to retrieve this is to shift + right click the file and select "copy as path."
    - If the Excel file is in the same directory as the program file, the name alone should suffice.
    - Paste file path into prompt. Remove quotation marks if present. Make sure the file extension (e.g. .xlsm) is also included.
    - Press enter. The program will now run.
  3. Open file:
    - You should now see a new worksheet titled "Reconciled" in the workbook, which the script has created. The script will have identified and eliminated offsetting entries
    - The original will also be present. To check that the scipt performed well, you may sum the amounts on each worksheet and see that they are equal.
    - 
