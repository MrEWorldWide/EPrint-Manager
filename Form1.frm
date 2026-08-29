VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form EPrint 
   AutoRedraw      =   -1  'True
   Caption         =   "EPrint Manager"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12690
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame prevframe 
      Height          =   6375
      Left            =   5040
      TabIndex        =   15
      Top             =   120
      Width           =   7575
      Begin AcroPDFLibCtl.AcroPDF PrevPDF 
         Height          =   6015
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   7335
         _cx             =   12938
         _cy             =   10610
      End
   End
   Begin VB.Frame travframe 
      Caption         =   "Equotes"
      Height          =   5175
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton movebackbut 
         Caption         =   "Move"
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton swapbut 
         Caption         =   "Files"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   4680
         Width           =   975
      End
      Begin VB.FileListBox TravFiles 
         Height          =   4170
         Left            =   120
         Pattern         =   "*.pdf"
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.DirListBox TravDir 
         Height          =   4140
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scanned Prints"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton renamebut 
         Caption         =   "Rename"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton movebut 
         Caption         =   "Move"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   4680
         Width           =   975
      End
      Begin VB.FileListBox ScanDir 
         Height          =   4170
         Left            =   120
         Pattern         =   "*.pdf"
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Optionsframe 
      Caption         =   "Options"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   4815
      Begin VB.CommandButton exitbut 
         Caption         =   "Exit"
         Height          =   375
         Left            =   3600
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox NumCopies 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   9
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton printbut 
         Caption         =   "Print"
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox sizelist 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         ItemData        =   "Form1.frx":0000
         Left            =   240
         List            =   "Form1.frx":000A
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.Label sizelbl 
         Caption         =   "Paper Size"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label numcopylbl 
         Caption         =   "Number of copies:"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "EPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub exitbut_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ScanOrTrav = "scan"
On Error GoTo errs
    'make sure its the first time the application is loaded, turn off when done.
    'this changes how functions react
    AppLoadChk = True
    GetDirList
    AppLoadChk = False
    
Exit Sub
errs:
    Select Case Err.Number
        Case 76
            
            MsgBox "Scan directory load failed. Check the CONST.txt in " & App.Path & " and edit line 1 with the correct location.", vbCritical, "ERROR!"
            MsgBox "Current path: " & ScanDir.Path
            Exit Sub
        Case Else
            
    End Select
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbchk = MsgBox("Are you sure you want to exit?", vbCritical + vbYesNo, "WAIT!")
    If mbchk = vbYes Then
        Exit Sub
    ElseIf mbchk = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub movebackbut_Click()
    
    'mbchk = MsgBox("Move " & TravFiles.List(TravFiles.ListIndex) & " to " & ScanDir.Path & "?", vbQuestion + vbOKCancel, "Confirm")
    'If mbchk = vbOK Then
        Dim newstr, oldstr As String
        newstr = ScanDir.Path & "\" & TravFiles.List(TravFiles.ListIndex())
        oldstr = TravDir.Path & "\" & TravFiles.List(TravFiles.ListIndex())
        
        FileCopy oldstr, newstr
        Kill oldstr
   ' ElseIf mbchk = vbCancel Then
    '    Exit Sub
    'End If
    ReloadLists
    
End Sub

Private Sub movebut_Click()
    
On Error GoTo errs
    'mbchk = MsgBox("Move " & ScanDir.List(ScanDir.ListIndex) & " to " & TravDir.Path & "?", vbQuestion + vbOKCancel, "Confirm")
    'If mbchk = vbOK Then
        Dim newstr, oldstr As String
        newstr = TravDir.Path & "\" & ScanDir.List(ScanDir.ListIndex())
        oldstr = ScanDir.Path & "\" & ScanDir.List(ScanDir.ListIndex())
        
        FileCopy oldstr, newstr
        Kill oldstr
   ' ElseIf mbchk = vbCancel Then
    '    Exit Sub
    'End If
errs:
ReloadLists
    
End Sub
Private Sub printbut_Click()
    If Val(NumCopies.Text) <= 0 Or IsNumeric(NumCopies.Text) = False Then
        MsgBox "Invalid number of copies or non-numeric character.", vbInformation
        Exit Sub
    End If
    
    If sizelist.Text = "8.5x11" Then
        Printer.PaperSize = vbPRPSLetter
    ElseIf sizelist.Text = "11x17" Then
        Printer.PaperSize = vbPRPSTabloid
    End If
    
    Printer.Copies = Val(NumCopies.Text)
    PrevPDF.printAll
    Printer.EndDoc
    
End Sub

Private Sub renamebut_Click()
    
    Dim secstr As String
    'rename a scanned file
    Dim newname, oldname As String
    newname = InputBox("Rename the file:", "Rename Scanned Print")
    If Len(newname) = 0 Then
        Exit Sub
    End If
On Error Resume Next
    secstr = newname
    oldname = ScanDir.Path & "\" & ScanDir.List(ScanDir.ListIndex())
    newname = ScanDir.Path & "\" & newname & ".pdf"
    Name oldname As newname
    ScanDir.ListIndex() = secstr
    ReloadLists
    
    
End Sub

Private Sub renamebut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 114 Then
        Call renamebut_Click
    End If
End Sub

Private Sub ScanDir_Click()

    'load selection
    PrevPDF.LoadFile (ScanDir.Path & "\" & ScanDir.List(ScanDir.ListIndex))
    PrevPDF.setZoom (100)
  
End Sub



Private Sub swapbut_Click()
    
    'based on what directory is loaded in the equotes list, it will swap to the files of
    'that directory. Click the button again to return to the other.
    
    If swapbut.Caption = "Folder" Then
        TravFiles.Visible = False
        TravDir.Visible = True
        swapbut.Caption = "Files"
    ElseIf swapbut.Caption = "Files" Then
        swapbut.Caption = "Folder"
        TravFiles.Visible = True
        TravDir.Visible = False
    End If
        TravFiles.Path = TravDir.Path
        ScanVar = ScanDir.List(ScanDir.ListIndex)
        
       ReloadLists
        
End Sub

Private Sub TravFiles_Click()
    
    'load selection
    PrevPDF.LoadFile (TravFiles.Path & "\" & TravFiles.List(TravFiles.ListIndex))
    PrevPDF.setZoom (100)
    
End Sub
