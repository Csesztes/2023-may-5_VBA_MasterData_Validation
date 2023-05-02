# VBA_MasterData_KPI-s-master-data-template-generator- 
Committed on 2023-May-2

Scripts runs on a weekly basis.

Aims:
  1. Count and categorize items
  2. Create summary and template files which will be uploaded to SAP APO
  
  Operation:
  1.
  The script loops trough 176 files based on a mapping file. It checks how many rows there are in extract and adds it to a category based on the mapping.
  
  2.  
  Summary contains all the relevant data from all the workbooks opened based on the mapping, plus the the category and a timestamp. Then it will be feeded into a tablue   report.
 
  For the templates scripts also run troughs the files, based on mapping checks which columns should be used for copying into the templates. 
    Once it is figured out, it will save those into a seperate excel with the category name and current date.

Some of the techniques included in the scripts are arrays,dictionaries,enums, collection and little bit of classes as well.
