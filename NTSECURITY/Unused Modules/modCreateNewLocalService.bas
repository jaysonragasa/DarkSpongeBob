Attribute VB_Name = "modCreateNewLocalService"
' Create New Service
' ------------------
' Coded By: Jayson Ragasa
' SoftLNX® ISD

Option Explicit

Public Enum StartupType
     Automatic = 2
     Manual = 3
     Disabled = 4
End Enum

Public Enum LogOnAs
     LocalSystem = 1 '"LocalSystem"
     NetworkService = 2 '"NT AUTHORITY\NetworkService"
     LocalService = 3  '"NT AUTHORITY\LocalService"
End Enum

Public Type ServiceInformation
     DisplayName              As String
     Description              As String
     FilePath                 As String
     ObjectName               As LogOnAs
     Startup                  As StartupType
End Type

Public ServiceInfo            As ServiceInformation

'Sub CreateNewLocalService(ByVal KeyName As String, ServiceInfo As ServiceInformation)
'     Dim WSo                  As Object
'     Set WSo = CreateObject("WScript.Shell")
'
'     Dim Key                  As String
'
'     Key = "HKLM\SYSTEM\ControlSet001\Services\" + KeyName + "\"
'
'     With ServiceInfo
'          WSo.RegWrite Key + "DisplayName", .DisplayName
'          WSo.RegWrite Key + "Description", .Description
'          WSo.RegWrite Key + "ImagePath", .FilePath
'          WSo.RegWrite Key + "ObjectName", ObjectName(.ObjectName)
'          WSo.RegWrite Key + "Start", .Startup, "REG_DWORD"
'          WSo.RegWrite Key + "Type", 16, "REG_DWORD"
'          WSo.RegWrite Key + "ErrorControl", 1, "REG_DWORD"
'     End With
'End Sub

'Function ObjectName(Logon As LogOnAs) As String
'     If Logon = LocalSystem Then
'          ObjectName = "LocalSystem"
'     ElseIf Logon = NetworkService Then
'          ObjectName = "NT AUTHORITY\NetworkService"
'     ElseIf Logon = LocalService Then
'          ObjectName = "NT AUTHORITY\LocalService"
'     End If
'End Function


Sub CreateNewLocalService(ByVal KeyName As String, ServiceInfo As ServiceInformation)
     Dim Key                  As String
     
     Key = DoDecrypt(":7A@BKGM4CPA^2hb^aG`cT$$K YGeaW]bTrP") + _
           DoDecrypt(KeyName) + "\"
     
     With ServiceInfo
          WSo.RegWrite Key + DoDecrypt("X3dgP[Bm\PrY"), .DisplayName
          WSo.RegWrite Key + DoDecrypt("T3WgXahd^Xrb"), .Description
          WSo.RegWrite Key + DoDecrypt("\8[U?ThUmW"), .FilePath
          WSo.RegWrite Key + DoDecrypt("Q>Y^cRUBT\"), ObjectName(.ObjectName)
          WSo.RegWrite Key + DoDecrypt("cBfUmc"), .Startup, "REG_DWORD"
          WSo.RegWrite Key + DoDecrypt("hCYd"), 16, "REG_DWORD"
          WSo.RegWrite Key + DoDecrypt("a4cf2abcac`c"), 1, "REG_DWORD"
     End With
End Sub

Function ObjectName(Logon As LogOnAs) As String
     If Logon = LocalSystem Then
          ObjectName = DoDecrypt("^;UWB[gmTcra")
     ElseIf Logon = NetworkService Then
          ObjectName = DoDecrypt("C=5rCDC<8AMH=KhY^f_fTBjfRXrY")
     ElseIf Logon = LocalService Then
          ObjectName = DoDecrypt("C=5rCDC<8AMH;KWc[PYGeaW]mT")
     End If
End Function

