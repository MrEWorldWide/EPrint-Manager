VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form Form1 
   Caption         =   "PDF Preview"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame prevframe 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin AcroPDFLibCtl.AcroPDF PrevPDF 
         Height          =   6495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7335
         _cx             =   5080
         _cy             =   5080
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
