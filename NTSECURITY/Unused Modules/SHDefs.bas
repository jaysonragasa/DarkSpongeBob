Attribute VB_Name = "mShellDefs"
Option Explicit

' Brought to you by Brad Martinez
'   http://members.aol.com/btmtz/vb
'   http://www.mvps.org/ccrp

' Code was written in and formatted for 8pt MS San Serif

' ====================================================================

Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

' Frees memory allocated by the shell (pidls)
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Public Const MAX_PATH = 260

' Defined as an HRESULT that corresponds to S_OK.
Public Const NOERROR = 0

' Converts an item identifier list to a file system path.
' Returns TRUE if successful or FALSE if an error occurs, for example,
' if the location specified by the pidl parameter is not part of the file system.
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

' ====================================================================
' SHGetSpecialFolderLocation

' Retrieves the location of a special (system) folder.
' Returns NOERROR if successful or an OLE-defined error result otherwise.
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As SHSpecialFolderIDs, pidl As Long) As Long

' Special folder values for SHGetSpecialFolderLocation and
' SHGetSpecialFolderPath (Shell32.dll v4.71)
Public Enum SHSpecialFolderIDs
     CSIDL_DESKTOP = &H0
     CSIDL_INTERNET = &H1
     CSIDL_PROGRAMS = &H2
     CSIDL_CONTROLS = &H3
     CSIDL_PRINTERS = &H4
     CSIDL_PERSONAL = &H5
     CSIDL_FAVORITES = &H6
     CSIDL_STARTUP = &H7
     CSIDL_RECENT = &H8
     CSIDL_SENDTO = &H9
     CSIDL_BITBUCKET = &HA
     CSIDL_STARTMENU = &HB
     CSIDL_DESKTOPDIRECTORY = &H10
     CSIDL_DRIVES = &H11
     CSIDL_NETWORK = &H12
     CSIDL_NETHOOD = &H13
     CSIDL_FONTS = &H14
     CSIDL_TEMPLATES = &H15
     CSIDL_COMMON_STARTMENU = &H16
     CSIDL_COMMON_PROGRAMS = &H17
     CSIDL_COMMON_STARTUP = &H18
     CSIDL_COMMON_DESKTOPDIRECTORY = &H19
     CSIDL_APPDATA = &H1A
     CSIDL_PRINTHOOD = &H1B
     CSIDL_ALTSTARTUP = &H1D                      ' ' DBCS
     CSIDL_COMMON_ALTSTARTUP = &H1E               ' ' DBCS
     CSIDL_COMMON_FAVORITES = &H1F
     CSIDL_INTERNET_CACHE = &H20
     CSIDL_COOKIES = &H21
     CSIDL_HISTORY = &H22
End Enum

' ====================================================================
' SHGetSpecialFolderLocation

' Retrieves information about an object in the file system, such as a file,
' a folder, a directory, or a drive root.
Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pidl As Any, ByVal dwFileAttributes As Long, psfib As SHFILEINFOBYTE, ByVal cbFileInfo As Long, ByVal uFlags As SHGFI_flags) As Long

' If pidl is invalid, SHGetFileInfo can very easily blow up when filling the
' szDisplayName and szTypeName string members of the SHFILEINFO'
' struct, so we'll define these members as Byte.
Public Type SHFILEINFOBYTE   ' sfib
     hIcon                         As Long
     iIcon                         As Long
     dwAttributes                  As Long
     szDisplayName(1 To MAX_PATH)  As Byte
     szTypeName(1 To 80)           As Byte
End Type

Enum SHGFI_flags
     SHGFI_LARGEICON = &H0              ' sfi.hIcon is large icon
     SHGFI_SMALLICON = &H1              ' sfi.hIcon is small icon
     SHGFI_OPENICON = &H2               ' sfi.hIcon is open icon
     SHGFI_SHELLICONSIZE = &H4          ' sfi.hIcon is shell size (not system size), rtns BOOL
     SHGFI_PIDL = &H8                   ' pszPath is pidl, rtns BOOL
     SHGFI_USEFILEATTRIBUTES = &H10     ' pretent pszPath exists, rtns BOOL
     SHGFI_ICON = &H100                 ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
     SHGFI_DISPLAYNAME = &H200          ' isf.szDisplayName is filled, rtns BOOL
     SHGFI_TYPENAME = &H400             ' isf.szTypeName is filled, rtns BOOL
     SHGFI_ATTRIBUTES = &H800           ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
     SHGFI_ICONLOCATION = &H1000        ' fills sfi.szDisplayName with filename
                                        ' containing the icon, rtns BOOL
     SHGFI_EXETYPE = &H2000             ' rtns two ASCII chars of exe type
     SHGFI_SYSICONINDEX = &H4000        ' sfi.iIcon is sys il icon index, rtns hImagelist
     SHGFI_LINKOVERLAY = &H8000         ' add shortcut overlay to sfi.hIcon
     SHGFI_SELECTED = &H10000           ' sfi.hIcon is selected icon
End Enum
'

' Returns an absolute pidl (realtive to the desktop) from a special folder's ID.
' (calling proc is responsible for freeing the pidl)

'   hOwner - handle of window that will own any displayed msg boxes
'   nFolder  - special folder ID

Public Function GetPIDLFromFolderID(hOwner As Long, nFolder As SHSpecialFolderIDs) As Long
     Dim pidl As Long
     
     If SHGetSpecialFolderLocation(hOwner, nFolder, pidl) = NOERROR Then
          GetPIDLFromFolderID = pidl
     End If
End Function

' If successful returns the specified absolute pidl's displayname,
' returns an empty string otherwise.

Public Function GetDisplayNameFromPIDL(pidl As Long) As String
     Dim sfib As SHFILEINFOBYTE
     
     If SHGetFileInfo(pidl, 0, sfib, Len(sfib), SHGFI_PIDL Or SHGFI_DISPLAYNAME) Then
          GetDisplayNameFromPIDL = GetStrFromBufferA(StrConv(sfib.szDisplayName, vbUnicode))
     End If
End Function

' Returns a path from only an absolute pidl (relative to the desktop)

Public Function GetPathFromPIDL(pidl As Long) As String
     Dim sPath As String * MAX_PATH
     
     If SHGetPathFromIDList(pidl, sPath) Then   ' rtns TRUE (1) if successful, FALSE (0) if not
          GetPathFromPIDL = GetStrFromBufferA(sPath)
     End If
End Function

' Returns the string before first null char encountered (if any) from an ANSII string.

Public Function GetStrFromBufferA(sz As String) As String
     If InStr(sz, vbNullChar) Then
          GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
     Else
          ' If sz had no null char, the Left$ function
          ' above would return a zero length string ("").
          GetStrFromBufferA = sz
     End If
End Function
