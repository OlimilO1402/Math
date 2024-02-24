VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Heronrechner"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7575
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtAuswahl2 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtAuswahlProdukt 
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtAuswahl1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtAnzahl 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      MaxLength       =   4
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuadratzahlen 
      Caption         =   "&Quadratzahlen"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox lstProdukt 
      Height          =   2775
      IntegralHeight  =   0   'False
      Left            =   6360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdWurzelZiehen 
      Caption         =   "&Wurzel ziehen"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstResultat2 
      Height          =   2775
      IntegralHeight  =   0   'False
      Left            =   3240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ListBox lstResultat1 
      Height          =   2775
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtEingabe 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblEchteWurzel 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   7335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label txtWurzel 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   7335
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
'Ansonsten viel Spaß und Erfolg mit diesem Source !

'Autor: Marc Ermshaus
'E-Mail: marcermshaus@aol.com

Option Explicit

'Private AusgangsZahl As Double
'Private Resultat1 As Double
'Private Resultat1ListIndex As Integer
'Private Resultat2 As Double
'Private Resultat2ListIndex As Integer
'Private Wurzel As Double
''Private Zahl As Integer
'Private Fertig As Boolean
'Private ProduktListIndex As Integer
'Private TooltipsZeigen As Boolean

Private Type TNum
    n  As Long
    n2 As Long
    n3 As Long
End Type
Private m_N23() As TNum

Private Sub Form_Load()
    txtWurzel.Caption = "In der ersten Liste steht die Ausgangszahl, " & _
                        "in der zweiten diese Zahl hoch 2 und" & _
                        "in der dritten hoch 3."
    
    lblEchteWurzel.Caption = ""
End Sub

Private Sub CreateN23()
    Dim n As Long
    If Not Long_TryParse(txtAnzahl.Text, n) Then
        txtEingabe.Text = "ERROR"
        Exit Sub
    End If
    ReDim m_N23(1 To n)
    Dim i As Long
    For i = 1 To n
        With m_N23(i)
            .n = i
            .n2 = i * i
            .n3 = .n2 * 1
        End With
    Next
End Sub

Private Function Long_TryParse(ByVal s As String, ByRef n_out As Long) As Boolean
Try: On Error GoTo Catch
    n_out = CLng(s)
    Long_TryParse = True
Catch:
End Function

Private Sub cmdQuadratzahlen_Click()
    
    txtWurzel.Caption = "Berechne ..."
    lblEchteWurzel.Caption = ""
    
    CreateN23
    
    View_Update
    
    txtWurzel.Caption = "In der ersten Liste steht die Ausgangszahl, " & _
                        "in der zweiten diese Zahl hoch 2 und" & _
                        "in der dritten hoch 3."
    
    lblEchteWurzel.Caption = ""
End Sub

Private Sub View_Clear()
    lstResultat1.Clear
    lstResultat2.Clear
    lstProdukt.Clear
End Sub

Private Sub View_Update()
    lstResultat1.Clear
    lstResultat2.Clear
    lstProdukt.Clear
    Dim i As Long
    For i = LBound(m_N23) To UBound(m_N23)
        With m_N23(i)
            lstResultat1.AddItem .n
            lstResultat2.AddItem .n2
            lstProdukt.AddItem .n3
        End With
    Next
End Sub

Private Sub View_Add(ByVal i As Long)
    lstResultat1.AddItem i
    lstResultat2.AddItem i ^ 2
    lstProdukt.AddItem i ^ 3
End Sub

Private Sub cmdWurzelZiehen_Click()
    Dim n As Long
    If Not Long_TryParse(txtEingabe.Text, n) Then
        MsgBox "Could not parse: " & txtEingabe.Text
        Exit Sub
    End If
    Dim w1 As Double: w1 = SqrH(n)
    Dim w2 As Double: w2 = VBA.Math.Sqr(Abs(n))
    MsgBox "SqrH(" & n & ")=" & w1 & "; VBA.Math.Sqr(" & n & ")=" & w2
    
'    Fertig = False
'
'    Resultat1 = 0
'    Resultat2 = 0
'    Resultat1ListIndex = 0
'    Resultat2ListIndex = 0
'    ProduktListIndex = 0
'
'    lstResultat1.Clear
'    lstResultat2.Clear
'    lstProdukt.Clear
'
'    AusgangsZahl = 0
'
'    If IsNumeric(txtEingabe.Text) = False _
'        Or txtEingabe.Text = "0" Then
'
'        txtEingabe.Text = "ERROR"
'        Exit Sub
'    End If
'
'    AusgangsZahl = txtEingabe.Text
'
'    Resultat1 = 1
'    Resultat2 = AusgangsZahl
'
'    lstResultat1.List(0) = Resultat1
'    lstResultat2.List(0) = AusgangsZahl
'    lstProdukt.List(0) = lstResultat1.List(0) * lstResultat2.List(0)
'
'    txtWurzel.Caption = "Berechne ..."
'
'    DoEvents
'
'    Call HeronMethode
'
End Sub

Private Sub HeronMethode()

Dim AusgangsZahl As Double
Dim Resultat1 As Double
Dim Resultat1ListIndex As Integer
Dim Resultat2 As Double
Dim Resultat2ListIndex As Integer
Dim Wurzel As Double
''Private Zahl As Integer
Dim Fertig As Boolean
'Dim ProduktListIndex As Integer
'Dim TooltipsZeigen As Boolean


    Do Until Fertig
        Resultat1 = (Resultat1 + Resultat2) / 2
        Resultat2 = AusgangsZahl / Resultat1
        
        Resultat1ListIndex = Resultat1ListIndex + 1
        lstResultat1.List(Resultat1ListIndex) = Resultat1
        Resultat2ListIndex = Resultat2ListIndex + 1
        lstResultat2.List(Resultat2ListIndex) = Resultat2
        ProduktListIndex = ProduktListIndex + 1
        
        lstProdukt.List(ProduktListIndex) = Resultat1 * Resultat2
        
        If Resultat1 - Resultat2 < 0.0000000009 Then Fertig = True
        If Resultat1 = Resultat2 Then Fertig = True
    Loop
    
    'Falls Zahlen verschieden sind, wird gemittelt
    If Resultat1 <> Resultat2 Then
        Resultat1 = (Resultat1 + Resultat2) / 2
        Resultat2 = Resultat1

        lstResultat1.AddItem "___________________________"
        lstResultat2.AddItem "___________________________"
        lstProdukt.AddItem "_________"
        lstResultat1.AddItem Resultat1
        lstResultat2.AddItem Resultat2
        
        lstProdukt.AddItem Resultat1 * Resultat2
    End If
  
    txtWurzel.Caption = "Die Wurzel aus " & AusgangsZahl & _
        " lautet " & Resultat1 & "."
                      
    lblEchteWurzel.Caption = "Die mit der sqr-Funktion " & _
        "ermittelte Wurzel der Zahl " & AusgangsZahl & " lautet " & _
        Sqr(AusgangsZahl) & "."
        
    With txtEingabe
        .SelStart = 0
        .SelLength = Len(txtEingabe.Text)
        .SetFocus
    End With
End Sub

Public Function SqrH(ByVal n As Double) As Double
    'SquareRoot due to Heron algorithm
    Dim s As Long:   s = Sgn(n): n = Abs(n)
    Dim r As Double: r = n
    Dim i As Long
    SqrH = 1
    Do
        i = i + 1
        SqrH = (SqrH + r) / 2
        r = n / SqrH
        If (SqrH - r) < 0.0000000000001 Then Exit Do
        If i = 20 Then Exit Do
    Loop
    Debug.Print i
End Function

Private Sub lstResultat1_Click()
    lstResultat2.ListIndex = lstResultat1.ListIndex
    lstProdukt.ListIndex = lstResultat1.ListIndex
    
    lstResultat1.TopIndex = lstResultat1.ListIndex
    
    lstResultat2.TopIndex = lstResultat1.TopIndex
    lstProdukt.TopIndex = lstResultat1.TopIndex
    
    txtAuswahl1.Text = lstResultat1.Text
    txtAuswahl2.Text = lstResultat2.Text
    txtAuswahlProdukt.Text = lstProdukt.Text
End Sub

Private Sub lstResultat2_Click()
    lstResultat1.ListIndex = lstResultat2.ListIndex
    lstProdukt.ListIndex = lstResultat2.ListIndex
    
    lstResultat2.TopIndex = lstResultat2.ListIndex
    
    lstResultat1.TopIndex = lstResultat2.TopIndex
    lstProdukt.TopIndex = lstResultat2.TopIndex
    
    txtAuswahl1.Text = lstResultat1.Text
    txtAuswahl2.Text = lstResultat2.Text
    txtAuswahlProdukt.Text = lstProdukt.Text
End Sub
