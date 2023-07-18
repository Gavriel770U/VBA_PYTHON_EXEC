# VBA_PYTHON_EXEC
VBA code that runs on .xlsm or .docm files, installs python, and executes given python code.
______________________________________________________________________________________________
## How To Use?
Well basically the code can be used in .xlsm or .docm files as macros for workbooks, sheets, modules and etc.
To execute your custom code you have to use it in the next way:
  1. The code have to work in python.
  2. The first line of the code is just like regular python.
  3. After the first line, every line have to be separated by '|' character.
  4. If the line starts with TABs or SPACEs you have to add them after the '|' character.
  5. All the code have to be in one line.
  6. The line to change is:
     
      
      ##### python_script = "print('hello world')|print('banana')|print('apple')|if(1==1):|    print('yes')"
      
     `python_script` is a variable that stores the code, so it shouldn't be changed, but the content of the string can be changed for your needs.
