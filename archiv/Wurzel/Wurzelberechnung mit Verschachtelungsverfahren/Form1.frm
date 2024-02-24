VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Intervallschachtelung"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame fraWerte 
      Caption         =   " Werte "
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtRechts 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         Height          =   240
         Left            =   3480
         TabIndex        =   11
         Text            =   "50"
         ToolTipText     =   "Genauigkeit der Berechnung"
         Top             =   1320
         Width           =   300
      End
      Begin VB.TextBox txtRadikant 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         Height          =   240
         Left            =   570
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "534"
         ToolTipText     =   "Hier Radikant eingeben"
         Top             =   360
         Width           =   900
      End
      Begin VB.TextBox txtGenauigkeit 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         Height          =   240
         Left            =   255
         TabIndex        =   5
         Text            =   "8"
         ToolTipText     =   "Genauigkeit der Berechnung"
         Top             =   1230
         Width           =   300
      End
      Begin VB.CommandButton cmdBerechnen 
         Caption         =   "Berechnen"
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtLinks 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         Height          =   240
         Left            =   3480
         TabIndex        =   3
         Text            =   "1"
         ToolTipText     =   "Genauigkeit der Berechnung"
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Startwert rechts:"
         Height          =   195
         Left            =   2280
         TabIndex        =   12
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label lblErgebnis 
         BackColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   1920
         TabIndex        =   10
         ToolTipText     =   "Das ist das Ergebnis"
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label lblIstGleich 
         Caption         =   "="
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   90
      End
      Begin VB.Line Line4 
         X1              =   1575
         X2              =   1575
         Y1              =   255
         Y2              =   390
      End
      Begin VB.Line Line3 
         X1              =   450
         X2              =   1590
         Y1              =   255
         Y2              =   255
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   450
         Y1              =   720
         Y2              =   240
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   360
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Label lblGenauigkeit 
         Caption         =   "Genauigkeit:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblStartwert 
         AutoSize        =   -1  'True
         Caption         =   "Startwert links:"
         Height          =   195
         Left            =   2280
         TabIndex        =   7
         Top             =   960
         Width           =   1035
      End
   End
   Begin VB.Frame fraZwischenergebnisse 
      Caption         =   " Zwischenergebnisse "
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   4095
      Begin VB.ListBox lstRechts 
         Appearance      =   0  '2D
         Height          =   1395
         ItemData        =   "Form1.frx":0000
         Left            =   2040
         List            =   "Form1.frx":0002
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.ListBox lstLinks 
         Appearance      =   0  '2D
         Height          =   1395
         ItemData        =   "Form1.frx":0004
         Left            =   120
         List            =   "Form1.frx":0006
         TabIndex        =   1
         ToolTipText     =   "Zwischenergebnisse"
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

'Autor: Markus Palme
'E-Mail: MarkusPalme@activevb.de

Option Explicit

Private n As Integer
Private x As Double
Private y As Double
Private a As Double

Private Sub cmdBerechnen_Click()
    'Eingabe prüfen
    If Len(txtRadikant.Text) = 0 _
        Or IsNumeric(txtRadikant.Text) = False Then
     
        Call MsgBox("Wert nicht zulässig", vbExclamation, "Fehler")
        Exit Sub
    End If
    
    cmdBerechnen.Enabled = False
    a = txtRadikant.Text
    x = txtLinks.Text
    y = txtRechts.Text
    lstLinks.Clear
    lstRechts.Clear
    
    Call Berechnen
End Sub

Private Sub Berechnen()
    Dim m As Double
    
    For n = 0 To txtGenauigkeit.Text
        
        m = (x + y) / 2
        
        If m * m < a Then
            x = m
        Else
            y = m
        End If
        
        lstLinks.AddItem x
        lstRechts.AddItem y
        
    Next n
    
    lblErgebnis.Caption = "[ " & x & " ; " & y & " ]"
    cmdBerechnen.Enabled = True
End Sub

Private Sub lstLinks_Click()
    lstRechts.ListIndex = lstLinks.ListIndex
End Sub

Private Sub lstRechts_Click()
    lstLinks.ListIndex = lstRechts.ListIndex
End Sub
