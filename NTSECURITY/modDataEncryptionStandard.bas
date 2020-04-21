Attribute VB_Name = "modDES"
' Encryption - Decryption
' Coded By: Jayson Ragasa
'
' Based on STI Books [ System Technology Institute, Inc. (STI College Baguio, Philippines) ]
'    - Title: Computer Networks
'    - Chapter 5: The Presentation and Application and Layers - Security
'    - Page 88 - 90
'
' I dont have enough time to explain each line

' but in the book, it said.
'
'  -.:: Figure 5-2 Example of Encryption Steps ::.--------------
' |                                                             |
' | 1. Divide the plain text into groups of eight characters.   |
' |    Pad with blanks at the end as necessary.                 |
' | 2. Rearrange the plain text characters by interchanging     |
' |    adjacent characters - that is making the first character |
' |    the second and the second the first, and so on.          |
' | 3. Translate each alphabetic character into an ordinal      |
' |    number -- that is, A becomes 1, B becomes 2, and so on.  |
' |    with a blank being a zero.                               |
' | 4. Select an 8-character encryption key.                    |
' | 5. Repeat step 3 for the encryption key.                    |
' | 6. Add the result of step 3 to the result of step 5.        |
' | 7. Divide the total by 27 and retain the remainder (which   |
' |    will be between 0 and 26                                 |
' | 8. Translate the remainder back into a character to yield   |
' |    the cipher text.                                         |
' |                                                             |
'  -------------------------------------------------------------
'
'  -.:: Figure 5-3 Example of Encryption ::.--------------------
' |                                                             |
' | 1. N    E    T    W    O    R    K    S                     |
' |    (separate into groups of eight)                          |
' | 2. E    N    W    T    R    O    S    K                     |
' |    (rearrange characters)                                   |
' | 3. 05   14   23   20   18   15   19   11                    |
' |    (translate characters to decimal)                        |
' | 4. P    R    O    T    O    C    O    L                     |
' |    (select an 8-character encryption key)                   |
' | 5. 16   18   15   20   15   03   15   12                    |
' |    (translate characters to decimal)                        |
' | 6. 05   14   23   20   18   15   19   11                    |
' |    16   18   15   20   15   03   15   12                    |
' |    --   --   --   --   --   --   --   --                    |
' |    21   32   38   40   33   18   34   23                    |
' |    (add values (message + key)                              |
' | 7. 21   05   11   13   06   18   07   23                    |
' |    (remainder after division by 27)                         |
' | 8. U    E    K    M    F    R    G    W                     |
' |                                                             |
'  --------------------------------------------------------------
'
'  -.:: Figure 5-4 Example if Decryption Steps ::.--------------
' |                                                             |
' | 1. Divide the cipher text into groups if eight characters.  |
' |    Pad with blanks at the end as necessary.                 |
' | 2. Translate each cipher text alphabetic character and the  |
' |    encryption key into an ordinal number: A becomes 1,      |
' |    B becomes 2, with a blank being a zero.                  |
' | 3. For each eight-character grouping, subtract the ordinal  |
' |    number of the key value from the ordinal number of the   |
' |    cipher text. Subtract the ordinal number of the first    |
' |    character of the key from the ordinal value of the       |
' |    first character of the eight-character group, and so on. |
' | 4. Some number in step 3 must be "normalized" by adding 27. |
' | 5. All numeric values will now be between 0 and 26. Trans-  |
' |    late these numbers back to their alphanumeric equiva-    |
' |    lents: 0 is a blank, 1 is in A, 2 is a B, and so on.     |
' | 6. Rearrange the text resulting from step 5 by              |
' |    interchanging adjacent characters: make the first        |
' |    character the second and the second the first; the third |
' |    character the fourth and the fourth character the third; |
' |    and so on.                                               |
' |                                                             |
'  -------------------------------------------------------------
'
'  -.:: Figure 5-5 Example of Decryption ::.--------------------
' |                                                             |
' | 1. U    E    K    M    F    R    G    W                     |
' |    (separate into groups of eight)                          |
' | 2. 21   05   11   13   06   18   07   23                    |
' |    (translate to an ordinal number)                         |
' | 3. 21   05   11   13   06   18   07   23                    |
' |    16   18   15   20   15   03   15   12                    |
' |    --   --   --   --   --   --   --   --                    |
' |    (subtract key value from cipher text)                    |
' | 4. 05  -13  -04  -07  -09   15  -08   11                    |
' |    05  -13  -04  -07  -09   15  -08   11                    |
' |    --   --   --   --   --   --   --   --                    |
' |         27   27   27   27        27                         |
' |    --   --   --   --   --   --   --   --                    |
' |    05   14   23   20   18   15   19   11                    |
' |    (add 27 to negative values)                              |
' | 5. E    N    W    T    R    O    S    K                     |
' |    (translate numbers to characters)                        |
' | 6. N    E    T    W    O    R    K    S                     |
' |    (rearrange characters)                                   |
' |                                                             |
'  -------------------------------------------------------------
'
' known weakness:
'    1. try to enter spaces (a lot!) as your plain text and the encryption key will
'       appear.
'
' -----------------------------------------------------------------
' i added 94 and 256 characters for more encryption enhancement. ;)
' -----------------------------------------------------------------

Option Explicit

Public Enum UseCharacter
     Standard27 = 27               ' A to Z (Converted to ordinal numbers, so: 1 to 27)
     NynTFor94 = 94                ' Space to Tilde (32 to 126)
     TuHndrdFftySx256 = 256        ' 0 to 256 characters
End Enum

Private Type CharacterDetails
     Character           As String
     OrdinalNo           As Integer
     CharFrKey           As String
     OrdNoFrKey          As Integer
     EncryptedChar       As String
End Type

'Function Encrypt(ByVal PlainText As String, ByVal KeyText As String, ByVal UseChar As UseCharacter) As String
'     Dim i          As Integer
'     Dim temp       As Integer
'
'     Dim res        As String
'     Dim IntrChng   As String
'
'     Dim CharDet()  As CharacterDetails
'
'     If UseChar = Standard27 Then
'          PlainText = UCase$(PlainText)
'          PlainText = RemoveInvalidCharacters(PlainText)
'     End If
'
'     If CBool(Len(PlainText) Mod 2) = True Then
'          PlainText = PlainText + Chr$(32)
'     End If
'
'     ReDim CharDet(Len(PlainText))
'
'     IntrChng = Interchange(PlainText)
'     KeyText = RepeatKey(Len(IntrChng), KeyText)
'
'     For i = 1 To Len(IntrChng)
'          With CharDet(i - 1)
'               .Character = Mid$(IntrChng, i, 1)
'               .OrdinalNo = GetOrdinalNumber(.Character, UseChar)
'               .CharFrKey = Mid$(KeyText, i, 1)
'               .OrdNoFrKey = GetOrdinalNumber(.CharFrKey, UseChar)
'
'               temp = .OrdinalNo + .OrdNoFrKey
'
'               temp = temp Mod UseChar
'
'               .EncryptedChar = GetCharacter(temp, UseChar)
'
'               res = res + .EncryptedChar
'          End With
'     Next i
'
'     Encrypt = res
'End Function

Function Decrypt(ByVal EncryptedText As String, ByVal KeyText As String, ByVal UseChar As UseCharacter)
     Dim i          As Integer
     Dim temp       As Integer
     
     Dim res        As String
     Dim CharDet()  As CharacterDetails
     
     ReDim CharDet(Len(EncryptedText))
     
     KeyText = RepeatKey(Len(EncryptedText), KeyText)
     
     For i = 1 To Len(EncryptedText)
          With CharDet(i - 1)
               .Character = Mid$(EncryptedText, i, 1)
               .OrdinalNo = GetOrdinalNumber(.Character, UseChar)
               .CharFrKey = Mid$(KeyText, i, 1)
               .OrdNoFrKey = GetOrdinalNumber(.CharFrKey, UseChar)
               
               temp = .OrdinalNo - .OrdNoFrKey
               
               If temp < 0 Then temp = temp + UseChar
               
               .EncryptedChar = GetCharacter(temp, UseChar)
               
               res = res + .EncryptedChar
          End With
     Next i
     
     res = Interchange(res)
     
     If UseChar = Standard27 Then res = RemoveInvalidCharacters(res)
     
     Decrypt = Trim$(res)
End Function

Function Interchange(ByVal sText As String)
     Dim i          As Integer
     Dim res        As String
     
     For i = 1 To Len(sText) Step 2
          res = res + StrReverse(Mid$(sText, i, 2))
     Next i
     
     Interchange = res
End Function

Function RepeatKey(ByVal LengthOfText As Long, ByVal KeyText As String) As String
     Dim i          As Integer
     Dim tmp        As Integer
     
     Dim res        As String
     
     For i = 1 To LengthOfText \ Len(KeyText)
          res = res + KeyText
     Next i
     
     If LengthOfText > Len(res) Then
          tmp = LengthOfText - Len(res)
          
          res = res + Left$(KeyText, tmp)
     End If
     
     RepeatKey = res
End Function

Function GetOrdinalNumber(ByVal Character As String, ByVal UseChar As UseCharacter) As Integer
     If UseChar = Standard27 Then
          ' Standard Algorithm
          ' -----------------------------------------------------------------
          If Character = Chr$(32) Then GetOrdinalNumber = 0: Exit Function
     
          GetOrdinalNumber = 26 - (90 - Asc(Character))
          
     ElseIf UseChar = NynTFor94 Then
          ' 64 Characters
          ' -----------------------------------------------------------------
          GetOrdinalNumber = 94 - (126 - Asc(Character))
          
     ElseIf UseChar = TuHndrdFftySx256 Then
          ' 256 Characters
          ' -----------------------------------------------------------------
          GetOrdinalNumber = Asc(Character)
          
     End If
End Function

Function GetCharacter(ByVal OrdinalNo As Integer, ByVal UseChar As UseCharacter) As String
     If UseChar = Standard27 Then
          ' Standard Algorithm
          ' -----------------------------------------------------------------
          If OrdinalNo = 0 Then GetCharacter = Chr$(32): Exit Function
     
          GetCharacter = Chr$(Abs(OrdinalNo) + 64)
          
     ElseIf UseChar = NynTFor94 Then
          ' 64 Characters
          ' -----------------------------------------------------------------
          GetCharacter = Chr$((Abs(OrdinalNo) + 32))
          
     ElseIf UseChar = TuHndrdFftySx256 Then
          ' 256 Characters
          ' -----------------------------------------------------------------
          GetCharacter = Chr$(Abs(OrdinalNo))
          
     End If
End Function

' this is function will be called when Standard27 encryption is used.
' -------------------------------------------------------------------
' this function will remove all 0 to 64 and 91 to 255 charcters and
' convert it to Space (chr$(32))
Function RemoveInvalidCharacters(ByVal sText As String) As String
     Dim i          As Integer
     Dim c          As String
     Dim res        As String
     
     For i = 1 To Len(sText)
          c = Mid$(sText, i, 1)

          If (Asc(c) < 65) Or (Asc(c) > 90) Then
               res = res + Chr$(32)
          Else
               res = res + c
          End If
     Next i
     
     RemoveInvalidCharacters = res
End Function
