# find_and_replace_txt_autocad
This script uses as a parameter an excel file to find and replace text in Autocad Files

To execute this script successfuly, you should follow to steps bellow:
1. Install requirements with : pip install -r requirements.txt
2. copy and paste the script file (edit_text.pv) to the folder containing drawing files
3. create an excel file "excel_file_name.xlsx" with 3 columns named :
  a. files: contains files' name
  b. find: text to find in the corresponding file
  c. replace: new text to add
4. run the script with the command: python edit_text.py "excel_file_name.xlsx"

Enjoy the magic!!

