# BlocksFromExcel - sample code - use at own risk

Reads blockname and attribute property from an Excel sheet and places blocks accordingly by name and displays the attribute using the property from Excel.

The blocks need to exist in the drawing already, even if not visible (e.g. using a template drawing that contains all these blocks).

Sample DWG can be found in this repository as well as the Excel sheet that's needed for the test.

The test block (name: "asdf") needs to contain an attribute called "FDSA", which is visible (contained already in the "p3d.DWG"). 

Compile and load the dll with "netload" command in AutoCAD, then you can use the "BlocksFromExcel" command to run the sample script, it will prompt you to select the Excel sheet.
