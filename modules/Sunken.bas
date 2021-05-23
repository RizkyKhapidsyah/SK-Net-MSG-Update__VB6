Attribute VB_Name = "Sunken"
Option Explicit

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) _
        As Long
        
Const GWL_EXSTYLE = (-20)
Const WS_EX_CLIENTEDGE = &H200

Sub AddSunkenBorder(ObjX As Object)
    Dim lstStyle As Long
    ' Get objects extended window style
    lstStyle = GetWindowLong(ObjX.hwnd, GWL_EXSTYLE)
    ' Append the sunken border to the current extended window style
    lstStyle = lstStyle Or WS_EX_CLIENTEDGE
    ' Apply the change to the control
    Call SetWindowLong(ObjX.hwnd, GWL_EXSTYLE, lstStyle)
End Sub
'-- End --'

Public Function LockControl(ObjX As Object, cLock As Boolean)
   Dim i As Long
   If cLock Then
      ' This will lock the control
      LockWindowUpdate ObjX.hwnd
   Else
      ' This will unlock controls
      LockWindowUpdate 0
      ObjX.Refresh
   End If
End Function


