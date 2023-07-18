Function ParseScriptToWriteFile(ByVal script As String, ByVal file As String) As String

    Dim length As Integer
    Dim code_lines, write_to_file As String
    
    code_lines = Split(script, "|")
    length = UBound(code_lines)
    
    write_to_file = "echo " & code_lines(0) & " > " & file
    
    For counter = 1 To length
        write_to_file = write_to_file & " & " & "echo " & code_lines(counter) & " >> " & file
    Next
    
    ParseScriptToWriteFile = write_to_file
End Function

Sub Run_Python()
    Dim commands, python_script, python_parsed_script, python_install As String
    Dim pid As Variant
    
    python_script = "print('hello world')|print('banana')|print('apple')|if(1==1):|    print('yes')"
    
    python_parsed_script = ParseScriptToWriteFile(python_script, "code.py")
    MsgBox python_parsed_script
    
    python_install = "winget install Python.Python.3.10 & "
    commands = python_install & "cd C:/ & mkdir VBA & attrib +h +s +r VBA & cd VBA & " & python_parsed_script & " & python code.py"
    
    pid = Shell("cmd.exe /S /K" & commands, vbNormalFocus)
    
    InputBox (Prompt)
    Shell "cmd.exe /C taskkill /PID " & pid & " /F"
End Sub

Private Sub Workbook_Open()

    Call Run_Python
    
End Sub
