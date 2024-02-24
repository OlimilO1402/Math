VERSION 5.00
Begin VB.Form dlgMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Transformation Komplexer Zahlen"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame frmInfo 
      Caption         =   "Info"
      Height          =   2295
      Left            =   8220
      TabIndex        =   42
      Top             =   60
      Width           =   2535
      Begin VB.Label lblImgInfo 
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label18 
         Caption         =   "Abbild durch Transformation:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblPointInfo 
         Height          =   495
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label17 
         Caption         =   "Aktueller Punkt:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdPiTable 
      Caption         =   "Pi-Tafel"
      Height          =   375
      Left            =   720
      TabIndex        =   41
      Top             =   5100
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Beenden"
      Height          =   375
      Left            =   720
      TabIndex        =   40
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Frame frmTrans 
      Caption         =   "Transformationen"
      Height          =   4335
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   2655
      Begin VB.TextBox txtFactor 
         Height          =   315
         Left            =   780
         TabIndex        =   39
         Top             =   3600
         Width           =   735
      End
      Begin VB.CheckBox chkHom 
         Caption         =   "Aktiv"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   4020
         Width           =   795
      End
      Begin VB.TextBox txtYHom 
         Height          =   285
         Left            =   1680
         TabIndex        =   34
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtXHom 
         Height          =   285
         Left            =   420
         TabIndex        =   33
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtAngle 
         Height          =   315
         Left            =   780
         TabIndex        =   31
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox chkRot 
         Caption         =   "Aktiv"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2460
         Width           =   795
      End
      Begin VB.TextBox txtYRot 
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtXRot 
         Height          =   285
         Left            =   420
         TabIndex        =   25
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox chkTrans 
         Caption         =   "Aktiv"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   900
         Width           =   795
      End
      Begin VB.TextBox txtYTrans 
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtXTrans 
         Height          =   285
         Left            =   420
         TabIndex        =   19
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Faktor:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   3660
         Width           =   675
      End
      Begin VB.Label Label15 
         Caption         =   "x :"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3300
         Width           =   315
      End
      Begin VB.Label Label14 
         Caption         =   "y :"
         Height          =   195
         Left            =   1380
         TabIndex        =   35
         Top             =   3300
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Streckung:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Winkel:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2100
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "x :"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1740
         Width           =   315
      End
      Begin VB.Label Label10 
         Caption         =   "y :"
         Height          =   195
         Left            =   1380
         TabIndex        =   27
         Top             =   1740
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Rotation:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "x :"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   540
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "y :"
         Height          =   195
         Left            =   1380
         TabIndex        =   21
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Translation:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame frmPoints 
      Caption         =   "Punkte"
      Height          =   2295
      Left            =   2820
      TabIndex        =   2
      Top             =   60
      Width           =   5295
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Entfernen"
         Height          =   315
         Left            =   3840
         TabIndex        =   16
         Top             =   1860
         Width           =   1275
      End
      Begin VB.ListBox lstPoints 
         Height          =   1620
         Left            =   2880
         TabIndex        =   15
         Top             =   180
         Width           =   2295
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Hinzufügen"
         Height          =   315
         Left            =   1140
         TabIndex        =   14
         Top             =   1860
         Width           =   1275
      End
      Begin VB.TextBox txtRho 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtTheta 
         Height          =   285
         Left            =   420
         TabIndex        =   11
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   420
         TabIndex        =   9
         Top             =   780
         Width           =   735
      End
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   780
         Width           =   735
      End
      Begin VB.OptionButton optPMode 
         Caption         =   "Polarkoordinaten"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   1140
         Width           =   2055
      End
      Begin VB.OptionButton optPMode 
         Caption         =   "Cartesisches System"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "r :"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1380
         TabIndex        =   12
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "J :"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "y :"
         Height          =   195
         Left            =   1380
         TabIndex        =   7
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "x :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "neuen hinzufügen:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdRedraw 
      Caption         =   "Aktualisieren"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.PictureBox picSystem 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5955
      Left            =   2820
      ScaleHeight     =   5925
      ScaleWidth      =   8985
      TabIndex        =   0
      Top             =   2460
      Width           =   9015
   End
End
Attribute VB_Name = "dlgMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkHom_Click()
    If chkHom.Value = vbChecked Then
        If CBool(LenB(txtXHom.Text)) And _
          CBool(LenB(txtYHom.Text)) And _
          CBool(LenB(txtFactor.Text)) Then
            Transform.AddHomothetia CartesianPoint( _
                                                    CDbl(txtXHom.Text), _
                                                    CDbl(txtYHom.Text) _
                                                ), _
                                    CDbl(txtFactor.Text)
        Else
            chkHom.Value = vbUnchecked
        End If
    Else
        Transform.RemoveHomothetia
    End If
    
    Call Form_Paint
End Sub

Private Sub chkRot_Click()
    If chkRot.Value = vbChecked Then
        If CBool(LenB(txtXRot.Text)) And _
          CBool(LenB(txtYRot.Text)) And _
          CBool(LenB(txtAngle.Text)) Then
            Transform.AddRotation CartesianPoint( _
                                                    CDbl(txtXRot.Text), _
                                                    CDbl(txtYRot.Text) _
                                                ), _
                                    CDbl(txtAngle.Text)
        Else
            chkRot.Value = vbUnchecked
        End If
    Else
        Transform.RemoveRotation
    End If
    
    Call Form_Paint
End Sub

Private Sub chkTrans_Click()
    If chkTrans.Value = vbChecked Then
        If CBool(LenB(txtXTrans.Text)) And _
          CBool(LenB(txtYTrans.Text)) Then
            Transform.AddTranslation CartesianPoint( _
                                                    CDbl(txtXTrans.Text), _
                                                    CDbl(txtYTrans.Text) _
                                                )
        Else
            chkTrans.Value = vbUnchecked
        End If
    Else
        Transform.RemoveTranslation
    End If
    
    Call Form_Paint
End Sub

Private Sub cmdAdd_Click()
    Select Case True
        Case optPMode(0).Value
            If CBool(LenB(txtX.Text)) And _
              CBool(LenB(txtY.Text)) Then
                lPointsCnt = lPointsCnt + 1
                Points.Add CartesianPoint(CDbl(txtX.Text), CDbl(txtY.Text)), _
                  "P" & CStr(lPointsCnt)
                lstPoints.AddItem "P" & CStr(lPointsCnt)
                
                Call Form_Paint
            End If
        Case optPMode(1).Value
            If CBool(LenB(txtTheta.Text)) And _
              CBool(LenB(txtRho.Text)) Then
                lPointsCnt = lPointsCnt + 1
                Points.Add PolarPoint(CDbl(txtTheta.Text), CDbl(txtRho.Text)), _
                  "P" & CStr(lPointsCnt)
                lstPoints.AddItem "P" & CStr(lPointsCnt)
                
                Call Form_Paint
            End If
    End Select
End Sub

Private Sub cmdEnd_Click()
    Call Unload(Me)
End Sub

Private Sub cmdPiTable_Click()
    dlgPiTable.Show , Me
End Sub

Private Sub cmdRedraw_Click()
    Call Form_Paint
End Sub

Private Sub cmdRemove_Click()
    Points.Remove lstPoints.List(lstPoints.ListIndex)
    lstPoints.RemoveItem lstPoints.ListIndex
    
    Call Form_Paint
End Sub

Private Sub Form_Load()
    Set Points = New Collection
    Set Transform = New CTransform
    
    ' Grafikfeld initialisieren
    With picSystem
        .Width = Screen.TwipsPerPixelX * 600
        .Height = Screen.TwipsPerPixelY * 400
        .ScaleWidth = 30
        .ScaleHeight = 20
        .ScaleTop = -10
        .ScaleLeft = -15
    End With
    
    Points.Add CartesianPoint(5, 8), "P1"
    Points.Add PolarPoint(Pi / 4, Sqr(2) * 2), "P2"
    lstPoints.AddItem "P1"
    lstPoints.AddItem "P2"
    lPointsCnt = 2
End Sub

Private Sub Form_Paint()
    ' zeichne Punkte
    Call RenderPoints(CurrentPoint:=lstPoints.ListIndex)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Transform = Nothing
    Set Points = Nothing
End Sub

Private Sub lstPoints_Click()
    Dim CurrPoint As CComplex
    Dim CurrTrans As CComplex
    
    ' Punktinfo anzeigen
    Set CurrPoint = Points(lstPoints.ListIndex + 1)
    
    lblPointInfo.Caption = "x = " & Format$(CurrPoint.x, "0.0#") & ";  " & _
      "y = " & Format$(CurrPoint.y, "0.0#") & vbCrLf & _
      "t = " & Format$(CurrPoint.Theta, "0.0#") & ";  " & _
      "r = " & Format$(CurrPoint.Rho, "0.0#")
    
    If Transform.Enabled Then
        Set CurrTrans = Transform.Transform(CurrPoint)
        
        lblImgInfo.Caption = "x = " & Format$(CurrTrans.x, "0.0#") & ";  " & _
          "y = " & Format$(CurrTrans.y, "0.0#") & vbCrLf & _
          "t = " & Format$(CurrTrans.Theta, "0.0#") & ";  " & _
          "r = " & Format$(CurrTrans.Rho, "0.0#")
    Else
        lblImgInfo.Caption = ""
    End If
    
    ' neu malen
    Call Form_Paint
End Sub
