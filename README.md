# Excel_Search_Tool

## About

* This tool allows you to search any information (up to 20 entries) from all Excel files in the folder chosen by the user (non-recursive).
* After all the files in the selected folder have been processed, a new sheet will be created under the name “Result” in which the search results are found.
* The columns on the "Result" sheet represent:
  * File = the name of the file in which the information was found.
  * Sheet = the name of the sheet in the file where the information was found.
  * Cell Address = the location of the cell where the information in that file is located.
  * Link = Clicking on this link will open the file in which the information was found and place the cursor on the cell where the information is located.
  * Value = All the contents of the cell in which the respective information was found.
* The tool opens automatically in Read-Only and no changes can be saved to it.
* The “Initialization” sheet as well as the VBA source code are properly protected.
* After selecting the folder, the tool will search all the Excel files present in it. Excel files found in another folder under the selected folder will be ignored. The search is non-recursive for performance and optimization reasons.
* The duration of searches in a single file can be up to 30 seconds, depending on the performance of the operating system of the user or the volume of the file.

## How to use

1. If the macro activation notification appears (like the one in the picture below), click Enable Content. This step is usually required only once per user / operating system, Excel should remember this choice. If you do not enable macros, the tool will not work.

![alt text](https://github.com/KanaszM/Excel_Search_Tool/blob/main/ReadMe_Resources/Picture1.png)

2. We enter between 1 and 20 pieces of information (configurable, more of than on the "For developers" section) that we want to search for. Do not proceed further if you have not entered anything.

![alt text](https://github.com/KanaszM/Excel_Search_Tool/blob/main/ReadMe_Resources/Picture2.png)

3.	Click the "Select Directory > Search" button and select the folder that contains the files from which you want to search for the information entered in step 2.

![alt text](https://github.com/KanaszM/Excel_Search_Tool/blob/main/ReadMe_Resources/Picture3.png)

4.	Wait until the search process is completed. We can track the progress in the loading window (like the one in the picture below). Upon completion, the window closes automatically, and the current sheet will change to a sheet called "Result".

![alt text](https://github.com/KanaszM/Excel_Search_Tool/blob/main/ReadMe_Resources/Picture4.png)

### Careful!
When exiting the tool, saving is not necessary unless you want to save a copy of the tool to another location.
(This behaviour can be changed, details on the "For developers" section)

![alt text](https://github.com/KanaszM/Excel_Search_Tool/blob/main/ReadMe_Resources/Picture5.png)
