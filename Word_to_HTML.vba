'**Written by: Dylan Lang**'

Sub Convert()
Dim alert As Integer
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("WScript.Network")

'Checks for WORD running before start up
Set Services = GetObject("winmgmts:")
    For Each process In Services.InstancesOf("Win32_Process")
        If process.Name = "WINWORD.EXE" Then
            alert = MsgBox("WORD is running. Please save and close all open WORD documents.", vbCritical, "Alert")
            Exit Sub
        End If
    Next


'Checks for app restart
Dim sRange As Range
Dim match As Range
Set sRange = Sheet1.Range("A:A")
Set match = sRange.Find("Pending")
If Not match Is Nothing Then
    match.Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Interior.ColorIndex = clNone
Else
    Range("B2").Select
End If
            
'Checks for directory in starting cell
If IsEmpty(ActiveCell) Then
    alert = MsgBox("Please enter a directory.", vbCritical, "Alert")
    Exit Sub
End If
      
'Starts Loop to convert and stop when an empty cell is reached.
Do Until IsEmpty(ActiveCell)
    
    'Changes Status to Pending
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Value = "Pending"
    ActiveCell.Offset(0, 1).Select
    strfilepath = ActiveCell
    
    'Gets filename from cell
    sfilename = fso.GetBaseName(strfilepath)
    
    'Check for Converted_Docs and creates if missing
    If Not fso.FolderExists("C:\Users\" + objNetwork.UserName + "\Desktop\Converted_Docs\") Then
        fso.CreateFolder ("C:\Users\" + objNetwork.UserName + "\Desktop\Converted_Docs\")
        MsgBox ("A folder called Converted_Docs has been created in your Desktop")
    End If
    
    'Build Directory
    spath = fso.GetParentFolderName(strfilepath)
    fullpath = ("C:\Users\" + objNetwork.UserName + "\Desktop\Converted_Docs\") & spath
    'fso.CreateFolder (fullpath)
    Set objShell = CreateObject("Wscript.Shell")
    objShell.Run "cmd /c mkdir " & Chr(34) & fullpath & Chr(34), 0, False
    
    'Convert File to filtered HTML
    Dim MyWord
    Dim MyDoc
    Set MyWord = CreateObject("Word.Application")
    
    On Error Resume Next
    Set MyDoc = MyWord.Documents.Open("\" & ActiveCell.Text)
    'Error Catch
    If Err.Number <> 0 Then
        MsgBox ("There was a problem opening file" & Chr(34) & sfilename & "." & Chr(34) & "Please review.")
        ActiveCell.Interior.ColorIndex = 3 'red
        MyDoc.Close
        MyWord.Quit
        Set MyDoc = Nothing
        Set MyWord = Nothing
        Exit Sub
    End If
        
    MyDoc.SaveAs (fullpath & "\" & sfilename & ".htm"), 10, , , , , , , , , , 65001
    ntRow = intRow + 1
        
        
    'Clean up
    MyDoc.Close
    MyWord.Quit
    Set MyDoc = Nothing
    Set MyWord = Nothing
    'wait for Word to close before moving on
    foundproc = False
    Do While foundproc = False
        For Each process In Services.InstancesOf("Win32_Process")
            If process.Name = "WINWORD.EXE" Then
                WScript.Sleep 5000
            Else
                foundproc = True
            End If
        Next
    Loop
    
    'Change status and step down
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Value = "Complete"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(1, 0).Select
Loop
    MsgBox ("DONE!")
End Sub
