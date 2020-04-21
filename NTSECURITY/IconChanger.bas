Attribute VB_Name = "IconChanger"
Option Explicit

Type DIB_HEADER
     Size                As Long
     Width               As Long
     Height              As Long
     Planes              As Integer
     Bitcount            As Integer
     Reserved            As Long
     ImageSize           As Long
End Type

Type ICON_DIR_ENTRY
     bWidth              As Byte
     bHeight             As Byte
     bColorCount         As Byte
     bReserved           As Byte
     wPlanes             As Integer
     wBitCount           As Integer
     dwBytesInRes        As Long
     dwImageOffset       As Long
End Type

Type ICON_DIR
     Reserved            As Integer
     Type                As Integer
     Count               As Integer
End Type

Type DIB_BITS
     Bits()              As Byte
End Type

Public Enum Errors
     FILE_CREATE_FAILED = 1000
     FILE_READ_FAILED
     INVALID_PE_SIGNATURE
     INVALID_ICO
     NO_RESOURCE_TREE
     NO_ICON_BRANCH
     CANT_HACK_HEADERS
End Enum

Private Type MEMICONDIRENTRY
     bWidth              As Byte        '// Width of the image
     bHeight             As Byte        '// Height of the image (times 2)
     bColorCount         As Byte        '// Number of colors in image (0 if >=8bpp)
     bReserved           As Byte        '// Reserved
     wPlanes             As Integer     '// Color Planes
     wBitCount           As Integer     '// Bits per pixel
     dwBytesInRes        As Long        '// how many bytes in this resource?
     nID                 As Integer     '// the ID
End Type

Private Const LOAD_LIBRARY_AS_DATAFILE = &H2&
Private Const RT_GROUP_ICON = 14&
Private Const RT_ICON = 3
Private Const GENERIC_WRITE = &H40000000
Private Const CREATE_ALWAYS = 2
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Any, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, lpName As Any, lpType As Any) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Private arIconNames()    As String        ' List of available resource names in the exe
Private intResCounter    As Integer       ' Used by callback procedure

Public Function ReplaceIcons(Source As String, Dest As String, Error As String) As Long
     On Error GoTo ErrHandler
     
     Dim IcoDir As ICON_DIR
     Dim IcoDirEntry As ICON_DIR_ENTRY
     Dim tBits As DIB_BITS
     Dim Icons() As IconDescriptor
     Dim lngRet As Long
     Dim BytesRead As Long
     Dim hSource As Long
     Dim hDest As Long
     Dim ResTree As Long
     
     hSource = CreateFile(Source, ByVal &H80000000, 0, ByVal 0&, 3, 0, ByVal 0)
     
     If hSource >= 0 Then
          If Valid_ICO(hSource) Then
               SetFilePointer hSource, 0, 0, 0
               ReadFile hSource, IcoDir, 6, BytesRead, ByVal 0&
               ReadFile hSource, IcoDirEntry, 16, BytesRead, ByVal 0&
               SetFilePointer hSource, IcoDirEntry.dwImageOffset, 0, 0
               
               ReDim tBits.Bits(IcoDirEntry.dwBytesInRes) As Byte
               
               ReadFile hSource, tBits.Bits(0), IcoDirEntry.dwBytesInRes, BytesRead, ByVal 0&
               CloseHandle hSource
               
               hDest = CreateFile(Dest, ByVal (&H80000000 Or &H40000000), 0, ByVal 0&, 3, 0, ByVal 0)
               
               If hDest >= 0 Then
                    If Valid_PE(hDest) Then
                         ResTree = GetResTreeOffset(hDest)
                         
                         If ResTree > 308 Then   ' Sanity check
                              lngRet = GetIconOffsets(hDest, ResTree, Icons)
                              
                              SetFilePointer hDest, Icons(1).Offset, 0, 0
                              WriteFile hDest, tBits.Bits(0), UBound(tBits.Bits), BytesRead, ByVal 0&
                              
                              If Not HackDirectories(hDest, ResTree, Icons(1).Offset, IcoDirEntry) Then
                                   Err.Raise CANT_HACK_HEADERS ', App.EXEName, "Unable to modify directories in target executable.  File may not contain any icon resources."
                              End If
                         Else
                              Err.Raise NO_RESOURCE_TREE ', App.EXEName, Dest & " does not contain a valid resource tree.  File may be corrupt."
                              CloseHandle hDest
                         End If
                    Else
                         Err.Raise INVALID_PE_SIGNATURE ', App.EXEName, Dest & " is not a valid Win32 executable."
                         CloseHandle hDest
                    End If
                    
                    CloseHandle hDest
               Else
                    Err.Raise FILE_CREATE_FAILED ', App.EXEName, "Failed to open " & Dest & ". Make sure file is not in use by another program."
               End If
          Else
               Err.Raise INVALID_ICO ', App.EXEName, Source & " is not a valid icon resource file."
               CloseHandle hSource
          End If
     Else
          Err.Raise FILE_CREATE_FAILED ', App.EXEName, "Failed to open " & Source & ". Make sure file is not in use by another program."
     End If
     
     ReplaceIcons = 0
     
     Exit Function
ErrHandler:
     ReplaceIcons = Err.Number
     Error = Err.Description
End Function

Public Function Valid_ICO(hFile As Long) As Boolean
     Dim tDir          As ICON_DIR
     Dim BytesRead     As Long
     
     If (hFile > 0) Then
          ReadFile hFile, tDir, Len(tDir), BytesRead, ByVal 0&
          
          If (tDir.Reserved = 0) And (tDir.Type = 1) And (tDir.Count > 0) Then
               Valid_ICO = True
          Else
               Valid_ICO = False
          End If
     Else
          Valid_ICO = False
     End If
End Function

'*******************************************************************************
' ExtractIcons: Takes two arguments:
'           1) strSource:  The full path of the executable file containing the
'                             icons to be extracted and saved as .ico files
'           2) strDest:    The full path of the .ico file which will be
'                             created and used to store extracted icons
'           Return Value:  Returns the error code
'*******************************************************************************

Public Function ExtractIcons(ByVal strSource As String, ByVal strDest As String) As Long
     On Error GoTo ErrHandler:
     
     ' Handles
     Dim hLib                 As Long
     Dim hResource            As Long
     Dim hLoaded              As Long
     Dim lPointer             As Long
     Dim hFile                As Long
     ' Icon Information Structures
     Dim SrcDir               As ICON_DIR
     Dim SrcEntries()         As ICON_DIR_ENTRY
     Dim SrcImages()          As DIB_BITS
     Dim MemEntry             As MEMICONDIRENTRY
     ' General use variables
     Dim arBytes()            As Byte
     Dim arID()               As Integer
     Dim lngBytesWritten      As Long
     Dim intI                 As Integer
     Dim intC                 As Integer
     Dim i                    As Integer
     Dim intBound             As Integer
     Dim intBaseOffset        As Integer
     Dim strTemp              As String
     
     ReDim arIconNames(0) As String
     intResCounter = 0
     
     ' Clear all memory structures
     hLib = 0: hResource = 0: hLoaded = 0: lPointer = 0: hFile = 0
     SrcDir.Count = 0: SrcDir.Reserved = 0: SrcDir.Type = 0
     
     ReDim SrcEntries(0) As ICON_DIR_ENTRY
     ReDim SrcImages(0) As DIB_BITS
     
     With MemEntry
          .bColorCount = 0: .bHeight = 0: .bWidth = 0: .bReserved = 0
          .wPlanes = 0: .wBitCount = 0: .dwBytesInRes = 0: .nID = 0
     End With
     
     ' Validate arguments
     If strSource = "" Or strDest = "" Then
          Err.Raise 1011 ', App.EXEName & ".SwapIcon.bas", "File not found"
     Else
          If Right$(strDest, 4) <> ".ico" Then strDest = strDest & ".ico"
     End If
     
     ' Load the executable into memory as a datafile
     hLib = LoadLibraryEx(strSource, ByVal 0&, LOAD_LIBRARY_AS_DATAFILE)
     
     If hLib = 0 Then Err.Raise 1011 ', App.EXEName & ".SwapIcon.bas", "File not found"
     
     ' Enumerate the resources in the library
     Call EnumResourceNames(hLib, RT_GROUP_ICON, AddressOf EnumResNameProc, 0)
     
     If UBound(arIconNames) < 0 Then Err.Raise 1002 ', App.EXEName & ".ExtractIcons", "No existing resources in source file"
     ' Loop through all resources found, copying the icons and writing them to file
     For intI = 0 To UBound(arIconNames)
          If Not arIconNames(intI) = "" Then
               ' Find, load, and lock the resource
               hResource = FindResource(hLib, ByVal arIconNames(intI), ByVal RT_GROUP_ICON)
               
               If hResource = 0 Then Err.Raise 1012 ', App.EXEName & ".SwapIcon.bas", "Failed to locate resource entry"
               
               hLoaded = LoadResource(hLib, hResource)
               
               If hLoaded = 0 Then Err.Raise 1013 ', App.EXEName & ".SwapIcon.bas", "Failed to load resource"
               
               lPointer = LockResource(hLoaded)
               
               If lPointer = 0 Then Err.Raise 1014 ', App.EXEName & ".SwapIcon.bas", "Failed to get pointer to resource data"
               
               ' Copy the icon directory structure from the file
               CopyMemory SrcDir, ByVal lPointer, Len(SrcDir)
               
               ' Check for icons in resource
               If SrcDir.Count > 0 Then
                    ' Copy all directory information into a byte array
                    ReDim SrcEntries(SrcDir.Count) As ICON_DIR_ENTRY
                    ReDim SrcImages(SrcDir.Count) As DIB_BITS
                    ReDim arID(SrcDir.Count) As Integer
                    
                    intBound = (Len(MemEntry) * (SrcDir.Count))
                    
                    ReDim arBytes(0 To intBound)
                    
                    ' Calculate the base offset for the icon bitmaps
                    intBaseOffset = (Len(SrcDir) + (SrcDir.Count * Len(SrcEntries(0))))
                    CopyMemory arBytes(0), ByVal (lPointer + Len(SrcDir)), intBound + 1
                    
                    ' For each icon in the resource, get the directory entry and the icon bits
                    For intC = 0 To (SrcDir.Count - 1)
                         ' Temporarily hold the data in the MemEntry structure
                         CopyMemory MemEntry, arBytes(intC * Len(MemEntry)), Len(MemEntry)
                         ' Add the icon's ID to the array
                         arID(intC) = MemEntry.nID
                         ' Copy the temp structure into the IconDirEntry structure
                         CopyMemory SrcEntries(intC), MemEntry, Len(MemEntry)
                         ' Assign the image offset
                         SrcEntries(intC).dwImageOffset = intBaseOffset
                         intBaseOffset = intBaseOffset + SrcEntries(intC).dwBytesInRes
                    Next intC
                    
                    ' Locate and copy the icon images
                    For intC = 0 To (SrcDir.Count - 1)
                         hResource = FindResource(hLib, ByVal "#" & CStr(arID(intC)), ByVal RT_ICON)
                         
                         If hResource > 0 Then
                              hLoaded = LoadResource(hLib, hResource)
                              
                              If hLoaded > 0 Then
                                   lPointer = LockResource(hLoaded)
                                   
                                   If lPointer > 0 Then
                                        ReDim Preserve SrcImages(intC).Bits(0 To SrcEntries(intC).dwBytesInRes)
                                        CopyMemory SrcImages(intC).Bits(0), ByVal lPointer, SrcEntries(intC).dwBytesInRes
                                   Else
                                        Err.Raise 1013 ', App.EXEName & ".ExtractIcons", "Failed to get resource address."
                                   End If
                              Else
                                   Err.Raise 1012 ', App.EXEName & ".ExtractIcons", "Failed to load resource."
                              End If
                         Else
                              Err.Raise 1011 ', App.EXEName & ".ExtractIcons", "Failed to locate resource."
                         End If
                    Next intC
                    
                    ' Append an index to the filename if more than one file will be created
                    If intI > 0 Then
                         strTemp = Left$(strDest, Len(strDest) - 4)
                         strTemp = strTemp & "(" & CStr(intI + 1) & DoDecrypt("{vW]m^") ' ).ico
                    Else
                         strTemp = strDest
                    End If

                    ' Create a new .ico file and write the complete icon resource
                    hFile = CreateFile(strTemp, GENERIC_WRITE, 0, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
                    
                    If hFile = 0 Then Err.Raise 1014 ', App.EXEName & ".IconSwap.bas", "Failed to get read/write handle"
                    
                    ' Write the directory
                    WriteFile hFile, SrcDir, Len(SrcDir), lngBytesWritten, ByVal 0&
                    
                    ' Write the directory entries
                    For intC = 0 To SrcDir.Count - 1
                         WriteFile hFile, SrcEntries(intC), Len(SrcEntries(intC)), lngBytesWritten, ByVal 0&
                    Next intC
                    
                    ' Write the icon bitmaps
                    For intC = 0 To SrcDir.Count - 1
                         WriteFile hFile, SrcImages(intC).Bits(0), SrcEntries(intC).dwBytesInRes, lngBytesWritten, ByVal 0&
                    Next intC
                    
                    ' Close the file
                    CloseHandle hFile
               End If
          End If
     Next intI
     
     ' Release the library and return the error code
     FreeLibrary (hLib)
     ExtractIcons = Err.Number
     Exit Function
     
ErrHandler:
     ExtractIcons = Err.Number
End Function

Sub RemoveExtractedIcon()
     If Len(Dir(dTmp + DoDecrypt("{wW]m^"))) <> 0 Then '*.ico
          Kill dTmp + DoDecrypt("{wW]m^") ' *.ico
          DoEvents
     End If
End Sub

'*******************************************************************************
' Private helper functions
'*******************************************************************************
Private Function EnumResNameProc(ByVal hModule As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByVal lParam As Long) As Long
     'BOOL CALLBACK EnumResNameProc(
     '    HANDLE hModule,  // resource-module handle
     '    LPCTSTR lpszType,   // pointer to resource type
     '    LPTSTR lpszName, // pointer to resource name
     '    LONG lParam   // application-defined parameter
     '   );
     Dim ResName As String
     Dim ResType As String
     Dim Continue As Boolean
     Dim Buffer As String
     Dim nRet As Long
     
     ' Retrieve resource ID.
     ResType = DecodeResTypeName(lpszType)
     ResName = DecodeResTypeName(lpszName)
     
     ' Add resource name to the array
     If ResName > "" Then
          intResCounter = intResCounter + 1
          
          ReDim Preserve arIconNames(intResCounter) As String
          
          arIconNames(intResCounter - 1) = ResName
          Continue = True
     Else
          Continue = False
     End If
     
     ' Continue enumeration?
     EnumResNameProc = Continue
End Function

Private Function DecodeResTypeName(ByVal lpszValue As Long) As String
     If HiWord(lpszValue) Then
          ' Pointers will always be >64K
          DecodeResTypeName = PointerToStringA(lpszValue)
     Else
          ' Otherwise we have an ID.
          DecodeResTypeName = "#" & CStr(lpszValue)
     End If
End Function

Private Function PointerToStringA(lpStringA As Long) As String
     Dim Buffer() As Byte
     Dim nLen As Long
     
     If lpStringA Then
          nLen = lstrlenA(ByVal lpStringA)
          
          If nLen Then
               ReDim Buffer(0 To (nLen - 1)) As Byte
               
               CopyMemory Buffer(0), ByVal lpStringA, nLen
               PointerToStringA = StrConv(Buffer, vbUnicode)
          End If
     End If
End Function

'*********************************************************************************
' Private utility functions
'*********************************************************************************
Private Function LoWord(LongIn As Long) As Integer
     Call CopyMemory(LoWord, LongIn, 2)
End Function

Private Function HiWord(LongIn As Long) As Integer
     Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function
