Attribute VB_Name = "mShellNotify"
Option Explicit

' Brought to you by Brad Martinez
'   http://members.aol.com/btmtz/vb
'   http://www.mvps.org/ccrp

' Code was written in and formatted for 8pt MS San Serif

' ====================================================================
' Demonstrates how to receive shell change notifications (ala "what happens when the
' SHChangeNotify API is called?")

' Interpretation of the shell's undocumented functions SHChangeNotifyRegister (ordinal 2)
' and SHChangeNotifyDeregister (ordinal 4) would not have been possible without the
' assistance of James Holderness. For a complete (and probably more accurate) overview
' of shell change notifcations, please refer to James' "Shell Notifications" page at
' http://www.geocities.com/SiliconValley/4942/
' ====================================================================

Private m_hSHNotify      As Long      ' the one and only shell change notification handle for the desktop folder
Private m_pidlDesktop    As Long   ' the desktop's pidl

' User defined notiication message sent to the specified window's window proc.
Public Const WM_SHNOTIFY = &H401

' ====================================================================

Public Type PIDLSTRUCT
  ' Fully qualified pidl (relative to the desktop folder) of the folder to monitor changes in.
  ' 0 can also be specifed for the desktop folder.
  pidl                   As Long
  ' Value specifying whether changes in the folder's subfolders trigger a change notification
  '  event (it's actually a Boolean, but we'll go Long because of VB's DWORD struct alignment).
  bWatchSubFolders       As Long
End Type

Declare Function SHChangeNotifyRegister Lib "shell32" Alias "#2" (ByVal hWnd As Long, ByVal uFlags As SHCN_ItemFlags, ByVal dwEventID As SHCN_EventIDs, ByVal uMsg As Long, ByVal cItems As Long, lpps As PIDLSTRUCT) As Long

' hWnd        - Handle of the window to receive the window message specified in uMsg.

' uFlags        - Flag that indicates the meaning of the dwItem1 and dwItem2 members of the
'                     SHNOTIFYSTRUCT (which is pointed to by the window procedure's wParam
'                     value when the specifed window message is received). This parameter can
'                     be one of the SHCN_ItemFlags enum values below.
'                     This interpretaion may be inaccurate as it appears pidls are almost alway returned
'                     in the SHNOTIFYSTRUCT despite this value. See James' site for more info...

' dwEventId - Combination of SHCN_EventIDs enum values that specifies what events the
'                     specified window will be notified of. See below.
                      
' uMsg          - Window message to be used to identify receipt of a shell change notification.
'                      The message should *not* be a value that lies within the specifed window's
'                      message range ( i.e. BM_ messages for a button window) or that window may
'                      not receive all (if not any) notifications sent by the shell!!!

' cItems         - Count of PIDLSTRUCT structures in the array pointed to by the lpps param.

' lpps             - Pointer to an array of PIDLSTRUCT structures indicating what folder(s) to monitor
'                      changes in, and whether to watch the specified folder's subfolder.

' If successful, returns a notification handle which must be passed to SHChangeNotifyDeregister
' when no longer used. Returns 0 otherwise.

' Once the specified message is registered with SHChangeNotifyRegister, the specified
' window's function proc will be notified by the shell of the specified event in (and under)
' the folder(s) speciifed in apidl. On message receipt, wParam points to a SHNOTIFYSTRUCT
' and lParam contains the event's ID value.

' The values in dwItem1 and dwItem2 are event specific. See the description of the values
' for the wEventId parameter of the documented SHChangeNotify API function.
Public Type SHNOTIFYSTRUCT
     dwItem1 As Long
     dwItem2 As Long
End Type

' Closes the notification handle returned from a call to SHChangeNotifyRegister.
' Returns True if succeful, False otherwise.
Declare Function SHChangeNotifyDeregister Lib "shell32" Alias "#4" (ByVal hNotify As Long) As Boolean

' ====================================================================

' This function should be called by any app that changes anything in the shell.
' The shell will then notify each "notification registered" window of this action.
Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As SHCN_EventIDs, ByVal uFlags As SHCN_ItemFlags, ByVal dwItem1 As Long, ByVal dwItem2 As Long)

' Shell notification event IDs

Public Enum SHCN_EventIDs
     SHCNE_RENAMEITEM = &H1             ' (D) A nonfolder item has been renamed.
     SHCNE_CREATE = &H2                 ' (D) A nonfolder item has been created.
     SHCNE_DELETE = &H4                 ' (D) A nonfolder item has been deleted.
     SHCNE_MKDIR = &H8                  ' (D) A folder item has been created.
     SHCNE_RMDIR = &H10                 ' (D) A folder item has been removed.
     SHCNE_MEDIAINSERTED = &H20         ' (G) Storage media has been inserted into a drive.
     SHCNE_MEDIAREMOVED = &H40          ' (G) Storage media has been removed from a drive.
     SHCNE_DRIVEREMOVED = &H80          ' (G) A drive has been removed.
     SHCNE_DRIVEADD = &H100             ' (G) A drive has been added.
     SHCNE_NETSHARE = &H200             ' A folder on the local computer is being shared via the network.
     SHCNE_NETUNSHARE = &H400           ' A folder on the local computer is no longer being shared via the network.
     SHCNE_ATTRIBUTES = &H800           ' (D) The attributes of an item or folder have changed.
     SHCNE_UPDATEDIR = &H1000           ' (D) The contents of an existing folder have changed, but the folder still exists and has not been renamed.
     SHCNE_UPDATEITEM = &H2000          ' (D) An existing nonfolder item has changed, but the item still exists and has not been renamed.
     SHCNE_SERVERDISCONNECT = &H4000    ' The computer has disconnected from a server.
     SHCNE_UPDATEIMAGE = &H8000&        ' (G) An image in the system image list has changed.
     SHCNE_DRIVEADDGUI = &H10000        ' (G) A drive has been added and the shell should create a new window for the drive.
     SHCNE_RENAMEFOLDER = &H20000       ' (D) The name of a folder has changed.
     SHCNE_FREESPACE = &H40000          ' (G) The amount of free space on a drive has changed.
     
     #If (WIN32_IE >= &H400) Then
          ' (G) SHCNE_EXTENDED_EVENT:  the extended event is identified in dwItem1,
          ' packed in LPITEMIDLIST format (same as SHCNF_DWORD packing).
          '
          ' Unlike the standard events, the extended events are ORDINALs, so we
          ' don't run out of bits.  Extended events follow the SHCNEE_* naming
          ' convention.
          '
          ' The dwItem2 parameter varies according to the extended event.
          SHCNE_EXTENDED_EVENT = &H4000000
     #End If     ' WIN32_IE >= &H0400
     
     SHCNE_ASSOCCHANGED = &H8000000     ' (G) A file type association has changed.
     
     SHCNE_DISKEVENTS = &H2381F         ' Specifies a combination of all of the disk event identifiers. (D)
     SHCNE_GLOBALEVENTS = &HC0581E0     ' Specifies a combination of all of the global event identifiers. (G)
     SHCNE_ALLEVENTS = &H7FFFFFFF
     SHCNE_INTERRUPT = &H80000000       ' The specified event occurred as a result of a system interrupt.
                                        ' It is stripped out before the clients of SHCNNotify_ see it.
End Enum

#If (WIN32_IE >= &H400) Then
     ' SHCNE_EXTENDED_EVENT extended events.  These events are ordinals. This is not a bitfield.
     
     Public Enum SHCNEE_flags
          SHCNEE_ORDERCHANGED = 2       ' dwItem2 is the pidl of the changed folder
          SHCNEE_MSI_CHANGE = 4         ' dwItem2 is the product code
          SHCNEE_MSI_UNINSTALL = 5      ' dwItem2 is the product code
     End Enum
#End If

' Notification flags

' uFlags & SHCNF_TYPE is an ID which indicates what dwItem1 and dwItem2 mean
Public Enum SHCN_ItemFlags
     SHCNF_IDLIST = &H0                 ' LPITEMIDLIST
     SHCNF_PATHA = &H1                  ' path name
     SHCNF_PRINTERA = &H2               ' printer friendly name
     SHCNF_DWORD = &H3                  ' DWORD
     SHCNF_PATHW = &H5                  ' path name
     SHCNF_PRINTERW = &H6               ' printer friendly name
     SHCNF_TYPE = &HFF
     ' Flushes the system event buffer. The function does not return until the system is
     ' finished processing the given event.
     SHCNF_FLUSH = &H1000
     ' Flushes the system event buffer. The function returns immediately regardless of
     ' whether the system is finished processing the given event.
     SHCNF_FLUSHNOWAIT = &H2000
     
     #If UNICODE Then
          SHCNF_PATH = SHCNF_PATHW
          SHCNF_PRINTER = SHCNF_PRINTERW
     #Else
          SHCNF_PATH = SHCNF_PATHA
          SHCNF_PRINTER = SHCNF_PRINTERA
     #End If
End Enum

' user-defined struct for SHCNF_DWORD
Public Type DWORDPACKEDSTRUCT
     cb As Integer
     dwItem1 As Long
     dwItem2 As Long
End Type
'

' Registers the one and only shell change notification.

Public Function SHNotify_Register(hWnd As Long) As Boolean
     Dim ps As PIDLSTRUCT
     
     ' If we don't already have a notification going...
     If (m_hSHNotify = 0) Then
     
          ' Get the pidl for the desktop folder.
          m_pidlDesktop = GetPIDLFromFolderID(0, CSIDL_DESKTOP)
          
          If m_pidlDesktop Then
     
               ' Fill the one and only PIDLSTRUCT, we're watching
               ' desktop and all of the it's subfolders, everything...
               ps.pidl = m_pidlDesktop
               ps.bWatchSubFolders = True
     
               ' Register the notification, specifying that we want the dwItem1 and dwItem2
               ' members of the SHNOTIFYSTRUCT to be pidls. We're watching all events.
               m_hSHNotify = SHChangeNotifyRegister(hWnd, SHCNF_TYPE Or SHCNF_IDLIST, SHCNE_ALLEVENTS Or SHCNE_INTERRUPT, WM_SHNOTIFY, 1, ps)
     
               SHNotify_Register = CBool(m_hSHNotify)
     
          Else
               ' If something went wrong...
               Call CoTaskMemFree(m_pidlDesktop)
     
          End If   ' m_pidlDesktop
     End If   ' (m_hSHNotify = 0)
End Function

' Unregisters the one and only shell change notification.

Public Function SHNotify_Unregister() As Boolean
  
  ' If we have a registered notification handle.
  If m_hSHNotify Then
    ' Unregister it. If the call is successful, zero the handle's variable,
    ' free and zero the the desktop's pidl.
    If SHChangeNotifyDeregister(m_hSHNotify) Then
      m_hSHNotify = 0
      Call CoTaskMemFree(m_pidlDesktop)
      m_pidlDesktop = 0
      SHNotify_Unregister = True
    End If
  End If

End Function

' Returns the event string associated with the specified event ID value.

Public Function SHNotify_GetEventStr(dwEventID As Long) As String
     Dim sEvent As String
     
     Select Case dwEventID
          Case SHCNE_CREATE: sEvent = "SHCNE_CREATE"                       ' = &H2"
          Case SHCNE_DELETE: sEvent = "SHCNE_DELETE"                       ' = &H4"
          Case SHCNE_MKDIR: sEvent = "SHCNE_MKDIR"                         ' = &H8"
          Case SHCNE_RMDIR: sEvent = "SHCNE_RMDIR"                         ' = &H10"
          Case SHCNE_MEDIAINSERTED: sEvent = "SHCNE_MEDIAINSERTED"         ' = &H20"
          Case SHCNE_MEDIAREMOVED: sEvent = "SHCNE_MEDIAREMOVED"           ' = &H40"
          Case SHCNE_DRIVEREMOVED: sEvent = "SHCNE_DRIVEREMOVED"           ' = &H80"
          Case SHCNE_DRIVEADD: sEvent = "SHCNE_DRIVEADD"                   ' = &H100"
          Case SHCNE_NETSHARE: sEvent = "SHCNE_NETSHARE"                   ' = &H200"
          Case SHCNE_NETUNSHARE: sEvent = "SHCNE_NETUNSHARE"               ' = &H400"
          Case SHCNE_ATTRIBUTES: sEvent = "SHCNE_ATTRIBUTES"               ' = &H800"
          Case SHCNE_UPDATEDIR: sEvent = "SHCNE_UPDATEDIR"                 ' = &H1000"
          Case SHCNE_UPDATEITEM: sEvent = "SHCNE_UPDATEITEM"               ' = &H2000"
          Case SHCNE_SERVERDISCONNECT: sEvent = "SHCNE_SERVERDISCONNECT"   ' = &H4000"
          Case SHCNE_UPDATEIMAGE: sEvent = "SHCNE_UPDATEIMAGE"             ' = &H8000&"
          Case SHCNE_DRIVEADDGUI: sEvent = "SHCNE_DRIVEADDGUI"             ' = &H10000"
          Case SHCNE_RENAMEFOLDER: sEvent = "SHCNE_RENAMEFOLDER"           ' = &H20000"
          Case SHCNE_FREESPACE: sEvent = "SHCNE_FREESPACE"                 ' = &H40000"
          
          #If (WIN32_IE >= &H400) Then
               Case SHCNE_EXTENDED_EVENT: sEvent = "SHCNE_EXTENDED_EVENT"  ' = &H4000000"
          #End If     ' WIN32_IE >= &H0400
          
          Case SHCNE_ASSOCCHANGED: sEvent = "SHCNE_ASSOCCHANGED"           ' = &H8000000"
          
          Case SHCNE_DISKEVENTS: sEvent = "SHCNE_DISKEVENTS"               ' = &H2381F"
          Case SHCNE_GLOBALEVENTS: sEvent = "SHCNE_GLOBALEVENTS"           ' = &HC0581E0"
          Case SHCNE_ALLEVENTS: sEvent = "SHCNE_ALLEVENTS"                 ' = &H7FFFFFFF"
          Case SHCNE_INTERRUPT: sEvent = "SHCNE_INTERRUPT"                 ' = &H80000000"
     End Select
     
     SHNotify_GetEventStr = sEvent
End Function
