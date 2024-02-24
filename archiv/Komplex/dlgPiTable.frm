VERSION 5.00
Begin VB.Form dlgPiTable 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Pi-Tafel"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2655
   Icon            =   "dlgPiTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Auf die Zwischenablage"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ListBox lstPi 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2475
   End
End
Attribute VB_Name = "dlgPiTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tblPi(6) As Double

Private Sub cmdClipboard_Click()
    Clipboard.Clear
    Clipboard.SetText CStr(tblPi(lstPi.ListIndex))
End Sub

Private Sub Form_Load()
    Call InitTable
    Call InitList
End Sub
'----------------------------------------------------

Private Sub InitTable()
    tblPi(0) = Pi
    tblPi(1) = Pi * 2
    tblPi(2) = Pi / 2
    tblPi(3) = Pi / 3
    tblPi(4) = Pi / 4
    tblPi(5) = Pi / 5
    tblPi(6) = Pi / 6
End Sub

Private Sub InitList()
    With lstPi
        .AddItem "Pi"
        .AddItem "Pi * 2"
        .AddItem "Pi / 2"
        .AddItem "Pi / 3"
        .AddItem "Pi / 4"
        .AddItem "Pi / 5"
        .AddItem "Pi / 6"
    End With
End Sub
