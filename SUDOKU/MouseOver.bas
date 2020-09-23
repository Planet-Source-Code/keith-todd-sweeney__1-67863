Attribute VB_Name = "MouseOver"

Function MouseMove(FormName As Form, ControlName As Control, ControlLeft As Integer, ControlRight As Integer, ControlTop As Integer, ControlBottom As Integer)
    
' Return the mouse position if the screen
    Dim MousePos As POINTAPI
    Dim RetValue As Boolean
    RetValue = GetCursorPos(MousePos)
    
' Convert from Twips to Pixels
    Dim frmX, frmY
    frmX = MousePos.X - FormName.ScaleX(FormName.Left, vbTwips, vbPixels)
    frmY = MousePos.Y - FormName.ScaleY(FormName.Top, vbTwips, vbPixels)
    
' Hightlight  the Numbers ready for selection
    If frmX > ControlLeft And frmX < ControlRight Then
       If frmY > ControlTop And frmY < ControlBottom Then
         ControlName.ForeColor = vbRed
       Else
         ControlName.ForeColor = vbGreen
       End If
    Else
       ControlName.ForeColor = vbGreen
    End If
End Function


