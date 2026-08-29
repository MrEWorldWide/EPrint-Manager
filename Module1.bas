Attribute VB_Name = "Module1"
Option Explicit
Global SizeSmall, AppLoadChk As Boolean
Global FSO As New FileSystemObject
Global TXTread(2) As String
Global ScanVar, TravVar As Long
Global Scanstr, Travstr, ScanOrTrav As String

Public Function FileExists(FileLocation As String)
    Set FSO = New Scripting.FileSystemObject
    If FSO.FileExists(FileLocation) = False Then
        MsgBox "File does not exists!, vbNotify"
        GoTo errs
        Exit Function
    End If
    
Exit Function


End Function


Public Function GetDirList()
    
 On Error GoTo funcerr
    'do if the first time program is run
    If AppLoadChk = True Then
        Dim SNum As String
        
        Dim c As Integer
        
        c = 0
        'text file contains directory addresses
        Open App.Path & "\const.txt" For Input As #1
        
        While EOF(1) = False
            Line Input #1, SNum
            TXTread(c) = SNum
            c = c + 1
        Wend
        'scanned prints
        EPrint.ScanDir.Path = TXTread(0)
        'traveler navigation
        EPrint.TravDir.Path = TXTread(1)
        'files to display after swap from folder
        EPrint.TravFiles.Path = TXTread(2)
        Close #1
    End If

Exit Function

funcerr:
End Function

Public Function ReloadLists()
    
On Error GoTo funcerr
    
    If AppLoadChk = True Then
        Exit Function
    End If
    
  
        EPrint.ScanDir.Refresh
        EPrint.TravDir.Refresh
        EPrint.TravFiles.Refresh
        
Exit Function
funcerr:
    
End Function
