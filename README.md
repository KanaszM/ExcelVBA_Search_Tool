# Excel_Search_Tool

About
•	This tool allows you to search any information (up to 20 entries) from all Excel files in the folder chosen by the user (non-recursive).
•	After all the files in the selected folder have been processed, a new sheet will be created under the name “Result” in which the search results are found.
•	The columns on the "Result" sheet represent:
	File = the name of the file in which the information was found.
	Sheet = the name of the sheet in the file where the information was found.
	Cell Address = the location of the cell where the information in that file is located.
	Link = Clicking on this link will open the file in which the information was found and place the cursor on the cell where the information is located.
	Value = All the contents of the cell in which the respective information was found.
•	The tool opens automatically in Read-Only and no changes can be saved to it.
•	The “Initialization” sheet as well as the VBA source code are properly protected.
•	After selecting the folder, the tool will search all the Excel files present in it. Excel files found in another folder under the selected folder will be ignored. The search is non-recursive for performance and optimization reasons.
•	The duration of searches in a single file can be up to 30 seconds, depending on the performance of the operating system of the user or the volume of the file.
