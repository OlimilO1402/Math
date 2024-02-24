Attribute VB_Name = "modGraphics"
Option Explicit

Public Sub RenderPoints(CurrentPoint As Integer)
    Dim Point   As CComplex
    Dim Image   As CComplex     ' nach Transformation
    Dim Colour  As OLE_COLOR
    Dim ImgCol  As OLE_COLOR
    Dim Cnt     As Long
    
    Call RenderField
    
    ' nette kleine Kreuzchen
    For Each Point In Points
        If Cnt = CurrentPoint Then _
            Colour = &HF0 _
        Else _
            Colour = RGB(20, 180, 0)
        
        dlgMain.picSystem.Line _
          (Point.x - 0.1, -Point.y - 0.1)-(Point.x + 0.15, -Point.y + 0.15), Colour
        dlgMain.picSystem.Line _
          (Point.x - 0.1, -Point.y + 0.1)-(Point.x + 0.15, -Point.y - 0.15), Colour
        
        If Transform.Enabled Then
            Set Image = Transform.Transform(Point)
            
            If Cnt = CurrentPoint Then _
                ImgCol = &HA000A0 _
            Else _
                ImgCol = &HA00000
            
            dlgMain.picSystem.Line _
              (Image.x - 0.1, -Image.y - 0.1)-(Image.x + 0.15, -Image.y + 0.15), ImgCol
            dlgMain.picSystem.Line _
              (Image.x - 0.1, -Image.y + 0.1)-(Image.x + 0.15, -Image.y - 0.15), ImgCol
        End If
        
        Cnt = Cnt + 1
    Next Point
End Sub

Public Sub RenderField()
    Dim Step As Integer
    
    dlgMain.picSystem.Cls

    ' zeichne Koordinatensystem
    
    ' x-Achsenmarkierung
    dlgMain.picSystem.Line (-15, 0)-(15, 0), vbBlack
    For Step = -14 To 14
        dlgMain.picSystem.Line (Step, -0.2)-(Step, 0.2), vbBlack
    Next Step
    
    ' Pfeilspitze
    dlgMain.picSystem.Line (14.6, -0.2)-(15, 0), vbBlack
    dlgMain.picSystem.Line (14.6, 0.2)-(15, 0), vbBlack
    
    ' Beschriftung
    dlgMain.picSystem.CurrentX = 14.5
    dlgMain.picSystem.CurrentY = 0.4
    dlgMain.picSystem.Print "x"
    
    ' y-Achsenmarkierung
    dlgMain.picSystem.Line (0, -10)-(0, 10), vbBlack
    For Step = -9 To 9
        dlgMain.picSystem.Line (-0.2, Step)-(0.2, Step), vbBlack
    Next Step
    
    ' Pfeilspitze
    dlgMain.picSystem.Line (-0.2, -9.6)-(0, -10), vbBlack
    dlgMain.picSystem.Line (0.2, -9.6)-(0, -10), vbBlack
    
    ' Beschriftung
    dlgMain.picSystem.CurrentX = -0.6
    dlgMain.picSystem.CurrentY = -9.9
    dlgMain.picSystem.Print "y"
End Sub
