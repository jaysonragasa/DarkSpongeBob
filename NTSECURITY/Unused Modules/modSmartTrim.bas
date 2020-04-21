Attribute VB_Name = "modSmartTrim"
Option Explicit

Function TrimRight(ByVal sString As String)
     Dim i     As Integer
     
     If Len(sString) > 255 Then sString = Left$(sString, 255)
     
     For i = Len(sString) To 1 Step -1
          If Asc(Mid$(sString, i, 1)) <> 0 Then
               TrimRight = Left$(sString, i)
               Exit Function
          End If
     Next i
End Function
