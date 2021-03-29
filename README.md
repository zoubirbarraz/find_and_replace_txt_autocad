# Find and replace txt in autocad files
This script uses as a parameter an excel file to find and replace text in Autocad Files

To execute this script successfuly, you should follow to steps bellow:
1. **Install** requirements with : _pip install -r requirements.txt_
2. **Copy and paste** the script file (edit_text.py) to the folder containing drawing files
3. **Create** an excel file "excel_file_name.xlsx" with 3 columns named :
  3a. files: contains files' name
  3b. find: text to find in the corresponding file
  3c. replace: new text to add
4. **Run** the script with the command: python edit_text.py "excel_file_name.xlsx"

Enjoy the magic!!




![Preview_GIF](/images/find_and_replace_text.gif)
