Attribute VB_Name = "執行腳本"

Sub RunScript()

   
    Dim DeviceName As String
    Dim ClassName As String
    Dim CaseName As String
    Dim PackageName, APP_PackageName As String
    Dim CaseNumber As Integer
    
    ActiveWorkbook.Save
    Application.Wait Now() + TimeValue("00:00:02") '暫緩2秒
    
    
    If Dir(CStr("C:\TUTK_QA_TestTool\TestTool\")) = "" Then y = MsgBox("找不到C:\TUTK_QA_TestTool\TestTool路徑", 0 + 16, "Error"): Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile("C:\TUTK_QA_TestTool\TestTool\Uiautomator.bat", True)
    
    CheckRestDataInformation = ResetAPPData()

    If CheckAPPPackageName = False Then Exit Sub
    If CheckRestDataInformation = False Then Exit Sub
    
    
    PackageName = Sheets("infor").Cells(2, "B").Text
    ClassName = Sheets("infor").Cells(2, "C").Text
    APP_PackageName = Sheets("infor").Cells(2, "F").Text
    
    k = 2: CaseNumber = 0
    Do While Sheets("infor").Cells(k, "D") <> ""
        CaseNumber = CaseNumber + 1
        k = k + 1
    Loop
    
    i = 2
    Do
        DeviceName = Sheets("infor").Cells(i, "A").Text
        temp = ""
        
        If Sheets("infor").Cells(2, "E").Text = "TRUE" Then
            
            temp = temp & "echo Reset APP: && adb -s " & DeviceName & " shell pm clear " & APP_PackageName & " && " & "echo Device Name:" & DeviceName
        Else
            temp = temp & "echo Device Name:" & DeviceName
        
        End If
        
        If CaseNumber > 1 Then
            x = 2
            Do
                CaseName = "#" & Sheets("infor").Cells(x, "D").Text
                temp = temp & " && " & "adb -s " & DeviceName & " shell am instrument -w -r   -e debug false -e class " & PackageName & "." & ClassName & CaseName & " " & PackageName & ".test/android.support.test.runner.AndroidJUnitRunner"
                x = x + 1
            Loop Until Sheets("infor").Cells(x, "D") = ""
            
            oFile.WriteLine "start cmd /k " & Chr(34) & temp & Chr(34)
        ElseIf CaseNumber = 1 Then
            CaseName = "#" & Sheets("infor").Cells(2, "D").Text
            oFile.WriteLine "start cmd /k " & Chr(34) & temp & " && " & "adb -s " & DeviceName & " shell am instrument -w -r   -e debug false -e class " & PackageName & "." & ClassName & CaseName & " " & PackageName & ".test/android.support.test.runner.AndroidJUnitRunner" & Chr(34)
    
        Else
            
            oFile.WriteLine "start cmd /k " & Chr(34) & temp & " && " & "adb -s " & DeviceName & " shell am instrument -w -r   -e debug false -e class " & PackageName & "." & ClassName & " " & PackageName & ".test/android.support.test.runner.AndroidJUnitRunner" & Chr(34)
        
        End If
        i = i + 1
    Loop Until Sheets("infor").Cells(i, "A") = ""
    
    oFile.WriteLine "exit"
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
    Application.Wait Now() + TimeValue("00:00:02") '暫緩2秒
    r = Shell(Environ("windir") & "\system32\cmd.exe cmd/k" & "C:\TUTK_QA_TestTool\TestTool\Uiautomator.bat", 1)    '啟動cmd

End Sub
Function ResetAPPData()

    Sheets("infor").Cells(2, "E").NumberFormatLocal = "G/通用格式"
    If Sheets("infor").Cells(2, "E") = "False" Or Sheets("infor").Cells(2, "E") = "FALSE" Or Sheets("infor").Cells(2, "E") = "false" Then
        
        Sheets("infor").Cells(2, "E") = "False"
        Sheets("infor").Cells(2, "E").NumberFormatLocal = "G/通用格式"
        Sheets("infor").Cells(2, "E").Font.Color = RGB(0, 0, 0)
        ResetAPPData = True
        
    ElseIf Sheets("infor").Cells(2, "E") = "True" Or Sheets("infor").Cells(2, "E") = "TRUE" Or Sheets("infor").Cells(2, "E") = "true" Then
    
        Sheets("infor").Cells(2, "E") = "True"
        Sheets("infor").Cells(2, "E").NumberFormatLocal = "G/通用格式"
        Sheets("infor").Cells(2, "E").Font.Color = RGB(0, 0, 0)
        ResetAPPData = True
    Else
        y = MsgBox("Reset APP Data欄位請輸入大寫TRUE或FALSE", 0 + 16, "Error")
        Sheets("infor").Cells(2, "E").Font.Color = RGB(255, 0, 0)
        ResetAPPData = False
        Exit Function
        
    End If

End Function

Function CheckAPPPackageName()
    If Sheets("infor").Cells(2, "F").Text = "" Then
        
        y = MsgBox("請輸入測試的APP PackageName", 0 + 16, "Error")
        CheckAPPPackageName = False
        Exit Function
    Else
    
        CheckAPPPackageName = True
    
    End If
End Function


'Sub RunScript()
'
'
'    Dim DeviceName As String
'    Dim ClassName As String
'    Dim CaseName As String
'    Dim PackageName As String
'    Dim CaseNumber As Integer
'
'    ActiveWorkbook.Save
'    Application.Wait Now() + TimeValue("00:00:02") '暫緩2秒
'
'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Dim oFile As Object
'    Set oFile = fso.CreateTextFile("C:\TUTK_QA_TestTool\TestTool\Uiautomator.bat", True)
'
'    PackageName = Sheets("infor").Cells(2, "B").Text
'    ClassName = Sheets("infor").Cells(2, "C").Text
'
'    k = 2: CaseNumber = 0
'    Do While Sheets("infor").Cells(k, "D") <> ""
'        CaseNumber = CaseNumber + 1
'        k = k + 1
'    Loop
'
'    i = 2
'    Do
'        DeviceName = Sheets("infor").Cells(i, "A").Text
'        temp = ""
'        If CaseNumber > 1 Then
'            x = 2
'            Do
'                CaseName = "#" & Sheets("infor").Cells(x, "D").Text
'                temp = temp & " && " & "adb -s " & DeviceName & " shell am instrument -w -r   -e debug false -e class " & PackageName & "." & ClassName & CaseName & " " & PackageName & ".test/android.support.test.runner.AndroidJUnitRunner"
'                x = x + 1
'            Loop Until Sheets("infor").Cells(x, "D") = ""
'
'            oFile.WriteLine "start cmd /k " & Chr(34) & "echo Device Name:" & DeviceName & temp & Chr(34)
'        ElseIf CaseNumber = 1 Then
'            CaseName = "#" & Sheets("infor").Cells(2, "D").Text
'            oFile.WriteLine "start cmd /k " & Chr(34) & "echo Device Name:" & DeviceName & " && " & "adb -s " & DeviceName & " shell am instrument -w -r   -e debug false -e class " & PackageName & "." & ClassName & CaseName & " " & PackageName & ".test/android.support.test.runner.AndroidJUnitRunner" & Chr(34)
'
'        Else
'
'            oFile.WriteLine "start cmd /k " & Chr(34) & "echo Device Name:" & DeviceName & " && " & "adb -s " & DeviceName & " shell am instrument -w -r   -e debug false -e class " & PackageName & "." & ClassName & " " & PackageName & ".test/android.support.test.runner.AndroidJUnitRunner" & Chr(34)
'
'        End If
'        i = i + 1
'    Loop Until Sheets("infor").Cells(i, "A") = ""
'
'    oFile.WriteLine "exit"
'    oFile.Close
'    Set fso = Nothing
'    Set oFile = Nothing
'    Application.Wait Now() + TimeValue("00:00:02") '暫緩2秒
'    r = Shell(Environ("windir") & "\system32\cmd.exe cmd/k" & "C:\TUTK_QA_TestTool\TestTool\Uiautomator.bat", 1)    '啟動cmd
'
'End Sub
