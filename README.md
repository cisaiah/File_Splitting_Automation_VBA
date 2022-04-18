# File_Splitting_Automation_VBA

This is a Visual Basic Macro that automates the process of splitting data in an excel spreadsheet into different spreadsheet
The macro is created to be dynamic so that it can be used by different users on spreadsheets with different data content and split columns


#### Example of Use Case Scenario

Assuming you have prepared an excel spreadsheet that contains formulas, validations, conditional formatting, etc. The spreadsheet contains data for different vendors of your business or for different departments within your company.
You want to distribute the files to the different vendors/departments, but you only need data relating to each vendor/department to be included on each file. Also, you want to maintain all conditional formats, validations, and formulas.
In this case, you would need to copy the data for each vendor/department into a seperate spreadsheet. This may take a very long time especially if you have a lot of vendors/departments - for example, 100+ vendors

This VBA Macro can automate the process of file splitting for you


#### How It Works

On running the macro, you would be asked to select the header of the column you want to split the file on. In the case of the example above, this will be the vendor/department column
You can also specify if you want to hide the split-column (as all data on this column will be the same)
You can specify if you want to show a filter button on the split spreadsheets
You can specify if you want to freeze panes and on what cell you want the freeze panes to apply

The split files (sub files) will be automatically saved in the same file path as that of the main file. The name of the sub files will be the name of the main files plus the data on the split cells of each file. No changes will be made on the main file.


### Working Sample

Click  on this link to view a working sample - https://youtu.be/4xIrSbZvtCw
