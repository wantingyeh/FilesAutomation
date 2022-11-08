# FilesAutomation

The script in this file provides quick automation for splitting files or converting files.

## Split excel files based on a certain column name
- files: SplitExcelColumn.ipnb
- language: python
- usage: Choose the excel files that you want to split to mulitple files based on a certain column.
- example:

![image](https://user-images.githubusercontent.com/83806848/200459577-f643fa28-add5-42c8-9e3d-bbf9c527e6b9.png)
  - Split the column based on *fruit*
  - You will have individual fruit files (ie., "apple.xlsx", "orange.xlsx", "kiwi.xlsx")
  
## Convert excel to txt files
- files: GetTextfromExcel.xlsm
- language: vba
- usage: batch converting multiple excel files into txt files within a folder
- how to:
  - open *GetTextfromExcel.xlsm*
  - go to the ribbon [Developer > Macro > Batch_TextSave] > Edit
  - Edit the directory
    - MyDir = "C:\Users\USER\PycharmProjects\UniqueWordCalculator\text_agent_condition\parent\condition" 
  - Press [run]
