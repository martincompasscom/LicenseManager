Attribute VB_Name = "CommonStrings"
Option Explicit

Private Const ASC_SPACE As Integer = 32
Private Const ASC_TAB As Integer = 9
Private Const ASC_CR As Integer = 13
Private Const ASC_LF As Integer = 10
Private Const ASC_FF As Integer = 12
Private Const ASC_VT As Integer = 8
Private Const ASC_BS As Integer = 11

Public Const SOCKET_STATE_UNKNOWN As String = "Unknown"
Public Const SOCKET_STATE_NA As String = "n/a"
Public Const SOCKET_STATE_OPEN As String = "Open"
Public Const SOCKET_STATE_CLOSED As String = "Closed"
Public Const SOCKET_STATE_CONNECTED As String = "Connected"
Public Const SOCKET_STATE_CLOSING As String = "Closing"
Public Const SOCKET_STATE_ERROR As String = "Error"

Public Const NINE_ZEROS As String = "000000000"
Public Const EIGHT_ZEROS As String = "00000000"
Public Const SEVEN_ZEROS As String = "0000000"
Public Const SIX_ZEROS As String = "000000"
Public Const FIVE_ZEROS As String = "00000"
Public Const FOUR_ZEROS As String = "0000"
Public Const THREE_ZEROS As String = "000"
Public Const TWO_ZEROS As String = "00"
Public Const ONE_ZERO As String = "0"

Type CRITICAL_SECTION

   Reserved1 As Long
   Reserved2 As Long
   Reserved3 As Long
   Reserved4 As Long
   Reserved5 As Long
   Reserved6 As Long
   
End Type

Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)

Private Declare Sub CopyMemory Lib "kernel32" Alias _
         "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


'Extracts a sub byte array from a byte array and converts it to a HexValueString
Public Function GetHexValueStringFromByteArray(bytData() As Byte, _
                                               lngStartPosition As Long, _
                                               lngNumBytes As Long) As String
                                               
   Dim bytArray() As Byte
                                               
   bytArray = GetByteArrayFromByteArray(bytData, lngStartPosition, lngNumBytes)
   GetHexValueStringFromByteArray = ConvertByteArrayToHexValueString(bytArray)
                                               
End Function


Public Function GetHexValueStringFromString(strMessage As String) As String
        
   If Len(strMessage) > 0 Then
      GetHexValueStringFromString = ConvertByteArrayToHexValueString(ConvertStringToByteArray(strMessage))
   Else
      GetHexValueStringFromString = vbNullString
   End If

End Function


Public Function AppendByteArrayToByteArray(bytDest() As Byte, _
                                           lngDestStartIndex As Long, _
                                           bytSource() As Byte, _
                                           Optional blnMainLineCode As Boolean = False) As Boolean

   If blnMainLineCode Then 'only log if we really want it
      gclsLogger.Log gclsLogger.LOG_WHAT_MESSAGING_EXTRA, "In CommonStrings::AppendByteArrayToByteArray()"
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::AppendByteArrayToByteArray()"
   End If

   On Error GoTo ERROR
   
   AppendByteArrayToByteArray = False
   
   Dim lngDestIndex As Long
   Dim lngSourceIndex As Long
   For lngDestIndex = lngDestStartIndex To lngDestStartIndex + UBound(bytSource)
      bytDest(lngDestIndex) = bytSource(lngSourceIndex)
      lngSourceIndex = lngSourceIndex + 1
   Next lngDestIndex
   
   AppendByteArrayToByteArray = True
   
EXIT_FUNC:
   If blnMainLineCode Then 'only log if we really want it
      gclsLogger.Log gclsLogger.LOG_WHAT_MESSAGING_EXTRA, "Out CommonStrings::AppendByteArrayToByteArray()"
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::AppendByteArrayToByteArray()"
   End If
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::AppendByteArrayToByteArray()", Err
   Resume EXIT_FUNC
      
End Function


' Returns the right hand part of a byte array. If lngNumBytesToDelete = -1, just take all rights chars
' to the end.
Public Function RightBytes(bytData() As Byte, lngStartPosition As Long, Optional lngNumBytesToKeep As Long = -1) As Byte()

   On Error GoTo ERROR
   
   Dim lngDataLen As Long
   lngDataLen = ArrayLen(bytData)
   
   Dim lngNumNewSize As Long
   lngNumNewSize = IIf(lngNumBytesToKeep = -1, lngDataLen - lngStartPosition, lngNumBytesToKeep)
   
   If lngNumNewSize > 0 Then
      If lngDataLen >= lngNumNewSize Then ' we really have something to do
   
         Dim bytNewData() As Byte
         ReDim bytNewData(lngNumNewSize - 1)
               
         Dim lngNewIndex As Long
         For lngNewIndex = 0 To lngNumNewSize - 1
            bytNewData(lngNewIndex) = bytData(lngStartPosition + lngNewIndex)
         Next lngNewIndex
         
         RightBytes = bytNewData
               
      End If
   End If
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::RightBytes()", Err
   Resume EXIT_FUNC

End Function


'extracts a sub byte array from a byte array
'if lngNumBytes = -1 ( i.e., no parameter passed ), get the rest of the bytes
'If the number of possible bytes to return is < specified lngNumBytes, just return
'   as many bytes as necessary.
'If start position is not in the array, return an empty array
Public Function GetByteArrayFromByteArray(bytSource() As Byte, _
                                          lngStartPosition As Long, _
                                          Optional lngNumBytes As Long = -1) As Byte()

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::GetByteArrayFromByteArray()"

   On Error GoTo ERROR
   
   Dim lngArrayLen As Long
   lngArrayLen = ArrayLen(bytSource)
   
   Dim lngActualMaxBytesToReturn As Long
   lngActualMaxBytesToReturn = lngArrayLen - lngStartPosition
   
   If lngActualMaxBytesToReturn > 0 Then
   
      If lngNumBytes = -1 Or lngActualMaxBytesToReturn < lngNumBytes Then 'calculate a real value
         lngNumBytes = lngActualMaxBytesToReturn
      End If
      
      Dim bytArray() As Byte
      
      If lngNumBytes > 0 Then
      
         ReDim bytArray(lngNumBytes - 1)
         
         Dim lngIndex As Long
         For lngIndex = 0 To lngNumBytes - 1
            bytArray(lngIndex) = bytSource(lngIndex + lngStartPosition)
         Next
         
      Else ' it's just an empty array, just pass back an empty array
         gclsLogger.Log gclsLogger.LOG_WHAT_INFO, "CommonStrings::GetByteArrayFromByteArray(), creating an empty array."
      End If
      
      GetByteArrayFromByteArray = bytArray
      
   ElseIf lngActualMaxBytesToReturn = 0 Then ' an "error" state, but one that this function "allows"
      gclsLogger.Log gclsLogger.LOG_WHAT_INFO, "CommonStrings::GetByteArrayFromByteArray()" & vbCrLf & _
                     "   start position """ & CStr(lngStartPosition) & """ after end of array, len = """ & CStr(lngArrayLen) & """."
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CommonStrings::GetByteArrayFromByteArray()" & vbCrLf & _
                     "   start position """ & CStr(lngStartPosition) & """ after end of array, len = """ & CStr(lngArrayLen) & """."
   End If
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::GetByteArrayFromByteArray()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::GetByteArrayFromByteArray()", Err
   Resume EXIT_FUNC
      
End Function


'removes a sub array and returns the modified original array
Public Function RemoveByteArrayFromByteArray(bytSource() As Byte, _
                                             lngStartPosition As Long, _
                                             lngNumBytesToRemove As Long) As Byte()

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::RemoveByteArrayFromByteArray()"

   On Error GoTo ERROR
   
   Dim lngOriginalArrayLen As Long
   lngOriginalArrayLen = ArrayLen(bytSource)
   
   'stupidity check
   If lngStartPosition + lngNumBytesToRemove <= lngOriginalArrayLen Then ' ok
   
      Dim bytTemp1() As Byte
   
      If lngStartPosition = 0 Or _
         lngStartPosition + lngNumBytesToRemove = lngOriginalArrayLen Then ' remove array at beginning or at end
         
         If lngOriginalArrayLen > lngNumBytesToRemove Then ' there will be some array left over
      
            gclsLogger.Log gclsLogger.LOG_WHAT_INFO, "CommonStrings::RemoveByteArrayFromByteArray(), removing array at beginning or end."
            
            ReDim bytTemp1((lngOriginalArrayLen - lngNumBytesToRemove) - 1)
            
            Dim lngIndex As Long
            For lngIndex = lngStartPosition To lngOriginalArrayLen - 1
               bytTemp1(lngIndex - lngStartPosition) = bytSource(lngIndex)
            Next lngIndex
            
            RemoveByteArrayFromByteArray = bytTemp1
            
         Else
            ' return an empty array
            gclsLogger.Log gclsLogger.LOG_WHAT_INFO, "CommonStrings::RemoveByteArrayFromByteArray(), returning empty array."
         End If
               
      Else ' somewhere in the middle of the array
      
         gclsLogger.Log gclsLogger.LOG_WHAT_INFO, "CommonStrings::RemoveByteArrayFromByteArray(), removing array in the middle."
      
         Dim bytTemp2() As Byte
         
         ReDim bytTemp1(lngStartPosition - 1) ' first half
         ReDim bytTemp2((lngOriginalArrayLen - (lngStartPosition + lngNumBytesToRemove)) - 1) ' second half
         
         For lngIndex = 0 To lngStartPosition - 1
            bytTemp1(lngIndex) = bytSource(lngIndex)
         Next lngIndex
         
         For lngIndex = lngStartPosition + lngNumBytesToRemove To lngOriginalArrayLen - 1
            bytTemp2(lngIndex - (lngStartPosition + lngNumBytesToRemove)) = bytSource(lngIndex)
         Next lngIndex
         
         Dim bytTempAll() As Byte
         ReDim bytTempAll((lngOriginalArrayLen - lngNumBytesToRemove) - 1)
         
         bytTempAll = Combine2ByteArrays(bytTemp1, bytTemp2)
         
         RemoveByteArrayFromByteArray = bytTempAll
         
         Erase bytTempAll
         Erase bytTemp2
         
      End If
      
      Erase bytTemp1
      
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
                     "Error in CommonStrings::RemoveByteArrayFromByteArray(), trying to remove bytes that don't exist."
   End If
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::RemoveByteArrayFromByteArray()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::RemoveByteArrayFromByteArray()", Err
   Resume EXIT_FUNC
      
End Function


'Should only be used if paired up with ConvertStringToByteArray(). Why?
'There are "endian" issues which are reversed when using ConvertStringToByteArray().
Public Function ConvertByteArrayToString(bytData() As Byte) As String

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::ConvertByteArrayToString()"

   On Error GoTo ERROR

   Dim lngArrayLen As Long
   lngArrayLen = ArrayLen(bytData)
   
   Dim intIntermediateValue As Integer
   
   Dim lngIndex As Long
   For lngIndex = 0 To lngArrayLen \ 2 - 1
   
      Dim byt2(1) As Byte
      byt2(0) = bytData(2 * lngIndex)
      byt2(1) = bytData(2 * lngIndex + 1)
   
      intIntermediateValue = CopyByteArrayIntoInteger(byt2)
      
      ConvertByteArrayToString = ConvertByteArrayToString & ChrW(intIntermediateValue)
            
   Next lngIndex
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::ConvertByteArrayToString()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::ConvertByteArrayToString()", Err
   Resume EXIT_FUNC

End Function


'Should only be used if paired up with ConvertStringToByteArray() Why?
'There are "endian" issues which are reversed when using ConvertByteArrayToString().
Public Function ConvertStringToByteArray(strString As String) As Byte()

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::ConvertStringToByteArray()"

   On Error GoTo ERROR
   
   Dim bytData() As Byte
   ReDim bytData(Len(strString) * 2 - 1)
   
   Dim intIntermediateValue As Integer
   
   Dim lngIndex As Long
   For lngIndex = 1 To Len(strString)
   
      intIntermediateValue = AscW(Mid$(strString, lngIndex, 1))
      
      Dim byt2() As Byte
      byt2 = CopyIntegerIntoByteArray(intIntermediateValue)
      
      bytData(lngIndex * 2 - 2) = byt2(0)
      bytData(lngIndex * 2 - 1) = byt2(1)
      
   Next lngIndex
   
   ConvertStringToByteArray = bytData
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::ConvertStringToByteArray()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::ConvertStringToByteArray()", Err
   Resume EXIT_FUNC
   
End Function


'legacy version - for instance, used with licenseManager
Public Function ConvertByteArrayToString2(bytRawMessage() As Byte) As String

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::ConvertByteArrayToString2()"

   On Error GoTo ERROR
   
   Dim iPos As Integer
   
   ConvertByteArrayToString2 = StrConv(bytRawMessage, vbUnicode)
   
   'get rid of possible leading 0 byte
   If Left$(ConvertByteArrayToString2, 1) = vbNullChar Then
      ConvertByteArrayToString2 = Mid$(ConvertByteArrayToString2, 2)
   End If
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::ConvertByteArrayToString2()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::ConvertByteArrayToString2()", Err
   Resume EXIT_FUNC

End Function


'legacy version - for instance, used with licenseManager
Public Function ConvertStringToByteArray2(strString As String) As Byte()

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::ConvertStringToByteArray2()"

   On Error GoTo ERROR
   
   Dim bytData() As Byte
   ReDim bytData(Len(strString) - 1)
   
   Dim intIndex As Integer
   For intIndex = 0 To Len(strString) - 1
      bytData(intIndex) = AscB(Mid$(strString, intIndex + 1, 1))
   Next intIndex
   
   ConvertStringToByteArray2 = bytData
      
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::ConvertStringToByteArray2()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::ConvertStringToByteArray2()", Err
   Resume EXIT_FUNC

End Function


'convert a string to a byte array that is made up of UTF8 chars
Public Function ConvertStringToUTF8ByteArray(strString As String) As Byte()

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::ConvertStringToUTF8ByteArray()"

   On Error GoTo ERROR

   Dim bytData() As Byte
   bytData = ConvertStringToByteArray(strString)

   Dim bytNewData() As Byte
   ReDim bytNewData(ArrayLen(bytData) \ 2 - 1)
   
   Dim lngIndex As Long
   For lngIndex = 0 To ArrayLen(bytData) - 2 Step 2
      bytNewData(lngIndex \ 2) = bytData(lngIndex)
   Next lngIndex
   
   ConvertStringToUTF8ByteArray = bytNewData
      
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::ConvertStringToUTF8ByteArray()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::ConvertStringToUTF8ByteArray()", Err
   Resume EXIT_FUNC
   
End Function


'Converts an array of bytes to a hex value representation in a string - each byte is
'converted to two characters, for example 255 -> "FF" and 11 -> "0B"
Public Function ConvertByteArrayToHexValueString(bytRawMessage() As Byte) As String

   gclsLogger.Log gclsLogger.LOG_WHAT_MESSAGING_EXTRA, "In CommonStrings::ConvertByteArrayToHexValueString()"

   On Error GoTo ERROR
   
   Dim lngArrayLen As Long
   lngArrayLen = ArrayLen(bytRawMessage)

   If lngArrayLen > 0 Then
   
      Dim intIndex As Integer
      For intIndex = 0 To lngArrayLen - 1
         ConvertByteArrayToHexValueString = ConvertByteArrayToHexValueString & GetHexValueAsString(bytRawMessage(intIndex))
      Next
      
   End If
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_MESSAGING_EXTRA, "Out CommonStrings::ConvertByteArrayToHexValueString()"
   Exit Function
   
ERROR:
   Resume EXIT_FUNC

End Function


'Does the opposite of the above - any two characters in a string are converted to a single
'byte value, for example "FF" -> 255, "0B" -> 11
'Assumptions: 2 characters is one byte.
'             each character represents a hex value
'             therefore, number of characters is divisible by 2
Public Function ConvertHexValueStringToByteArray(strString As String) As Byte()

   Dim bytArray() As Byte

   Dim lngLen As Long
   lngLen = Len(strString)
   
   If lngLen Mod 2 <> 0 Then
      ReDim bytArray(lngLen \ 2 + 1) ' all 0's
      ConvertHexValueStringToByteArray = bytArray
      GoTo EXIT_FUNC
   End If
   
   ReDim bytArray(lngLen \ 2 - 1)

   Dim lngIndex As Long
   For lngIndex = 0 To lngLen \ 2 - 1
      
      Dim bytMSB As Byte
      Dim bytLSB As Byte
      
      bytMSB = GetByteValue(Mid$(strString, 2 * lngIndex + 1, 1))
      bytLSB = GetByteValue(Mid$(strString, 2 * lngIndex + 2, 1))
   
      bytArray(lngIndex) = Get2HalfByteValue(bytMSB, bytLSB)
   
   Next

   ConvertHexValueStringToByteArray = bytArray

EXIT_FUNC:
   Exit Function

End Function


'Converts a string representation of a value to a real byte value
Private Function GetByteValue(strChar As String) As Byte

   Select Case strChar
   
      Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
         GetByteValue = CByte(strChar)
         
      Case "A", "B", "C", "D", "E", "F"
         GetByteValue = AscB(strChar) - 55
         
      Case "a", "b", "c", "d", "e", "f"
         GetByteValue = AscB(strChar) - 87
      
      Case Else
         GetByteValue = 0
         
   End Select
   
End Function


'Converts a string representation of a value to a real byte value
'String can be either 1 or 2 characters


'Converts a string representation of a value to a real byte value
'String can be either 1 or 2 characters
Public Function GetByteValue2(strChars As String, bytValue As Byte) As Boolean

   On Error GoTo ERROR
   
   GetByteValue2 = False

   If Len(strChars) = 1 Then strChars = "0" & strChars

   Dim strChar As String
   strChar = Mid$(strChars, 1, 1)

   Select Case strChar
   
      Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
         bytValue = CByte(strChar)
         
      Case "A", "B", "C", "D", "E", "F"
         bytValue = AscB(strChar) - 55
         
      Case "a", "b", "c", "d", "e", "f"
         bytValue = AscB(strChar) - 87
      
      Case Else
         gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "CommonStrings::GetByteValue2(), byte value 1 in """ & strChars & """"
         GoTo EXIT_FUNC
         
   End Select
   
   bytValue = bytValue * 16
   
   strChar = Mid$(strChars, 2, 1)
   
   Select Case strChar
   
      Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
         bytValue = bytValue + CByte(strChar)
         
      Case "A", "B", "C", "D", "E", "F"
         bytValue = bytValue + (AscB(strChar) - 55)
         
      Case "a", "b", "c", "d", "e", "f"
         bytValue = bytValue + (AscB(strChar) - 87)
      
      Case Else
         gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "CommonStrings::GetByteValue2(), byte value 2 in """ & strChars & """"
         GoTo EXIT_FUNC
         
   End Select
   
   GetByteValue2 = True
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   StandardErrorTrap "Common::GetByteValue2()", Err
   Resume EXIT_FUNC
   
End Function


'Converts a single byte value into it's two character hex representation
Public Function GetHexValueAsString(bytValue As Byte) As String

   GetHexValueAsString = CStr(Hex$(bytValue))
   
   If Len(GetHexValueAsString) = 1 Then
      GetHexValueAsString = "0" & GetHexValueAsString
   End If
   
End Function


'Adds up the half-byte values, using a most significant half-byte and a least significant
'half-byte to come up with a complete byte value
'This function is used to convert hex-value strings to real values. For example, with this
'value '1a', we would make a byte with the value of '1', a byte with the value of 'a', and then
'call this function to get an actual numeric byte value
Public Function Get2HalfByteValue(MSB As Byte, LSB As Byte) As Byte

   Get2HalfByteValue = MSB * 16 + LSB

End Function


'Adds up two full byte values to get a long value.
'If you're going to use this function, be sure that you won't create a value > 32767
'Since bytes can't have a negative value, returning -1 is a good way of saying
'something was wrong.
Public Function Get2ByteValue(MSB As Byte, LSB As Byte) As Long

   On Error Resume Next
   Get2ByteValue = CLng(MSB) * 256 + CLng(LSB)
   
   If Err.Number <> 0 Then
      Get2ByteValue = -1
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
                     "CommonStrings::Get2ByteValue(), overflowed maximum possible value"
   End If
   
End Function


'Handles space, vbCr, vbLf, vbTab, vbBack, vbVerticalTab, vbFormFeed
Public Function LTrimWS(strStr As String) As String

   LTrimWS = strStr
   
   Dim intAscFirstChar As Integer
   intAscFirstChar = Asc(LTrimWS)

   While intAscFirstChar = ASC_SPACE Or _
         intAscFirstChar = ASC_LF Or _
         intAscFirstChar = ASC_CR Or _
         intAscFirstChar = ASC_TAB Or _
         intAscFirstChar = ASC_VT Or _
         intAscFirstChar = ASC_FF Or _
         intAscFirstChar = ASC_BS
         
      LTrimWS = Mid$(LTrimWS, 2)
            
      intAscFirstChar = Asc(LTrimWS)
      
   Wend
   
End Function


'Handles space, vbCr, vbLf, vbTab, vbBack, vbVerticalTab, vbFormFeed
Public Function RTrimWS(strStr As String) As String

   RTrimWS = strStr
   
   Dim intAscLastChar As Integer
   intAscLastChar = Asc(Right$(strStr, 1))

   While intAscLastChar = ASC_SPACE Or _
         intAscLastChar = ASC_LF Or _
         intAscLastChar = ASC_CR Or _
         intAscLastChar = ASC_TAB Or _
         intAscLastChar = ASC_VT Or _
         intAscLastChar = ASC_FF Or _
         intAscLastChar = ASC_BS
         
      RTrimWS = Left$(RTrimWS, Len(RTrimWS) - 1)
            
      intAscLastChar = Asc(Right$(RTrimWS, 1))
      
   Wend
   
End Function


Public Function TrimWS(strStr As String) As String

   TrimWS = LTrimWS(RTrimWS(strStr))

End Function


'If uninitialized, error 9 occurs, which means array is not initialized, len = 0
Public Function ArrayLen(varArray As Variant) As Long

   On Error Resume Next
   
   ArrayLen = UBound(varArray) + 1
   If Err.Number <> 0 Then
      ArrayLen = 0
   End If
   
End Function


' SelectionSort for strings
Public Sub SortStrings(ListArray() As String, _
                       Optional blnAscending As Boolean = True, _
                       Optional blnCaseSensitive As Boolean = False)
                                    
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::SortStrings()"
   
   On Error GoTo ERROR
       
   Dim lngMin As Long
   Dim lngMax As Long
   
   lngMin = LBound(ListArray)
   lngMax = UBound(ListArray)
   
   If lngMin = lngMax Then
      Exit Sub
   End If
   
   ' Order Ascending or Descending?
   Dim lngOrder As Long
   lngOrder = IIf(blnAscending, -1, 1)
   
   ' Case sensitive search or not?
   Dim lngCompareType As Long
   lngCompareType = IIf(blnCaseSensitive, vbBinaryCompare, vbTextCompare)
   
   ' Loop through array swapping the smallest\largest (determined by lngOrder)
   ' item with the current item
   Dim lngCount1 As Long
   For lngCount1 = lngMin To lngMax - 1
   
      Dim strSmallest       As String
      Dim lngSmallest       As Long
   
      strSmallest = ListArray(lngCount1)
      lngSmallest = lngCount1
      
      ' Find the smallest\largest item in the array
      Dim lngCount2 As Long
      For lngCount2 = lngCount1 + 1 To lngMax
         If StrComp(ListArray(lngCount2), strSmallest, lngCompareType) = lngOrder Then
            strSmallest = ListArray(lngCount2)
            lngSmallest = lngCount2
         End If
      Next
   
      ' Just swap them, even if we are swapping it with itself,
      ' as it is generally quicker to do this than test first
      ' each time if we are already the smallest with a
      ' test like: If lSmallest <> lCount1 Then
      ListArray(lngSmallest) = ListArray(lngCount1)
      ListArray(lngCount1) = strSmallest
   
   Next
   
EXIT_SUB:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::SortStrings()"
   Exit Sub
   
ERROR:
   StandardErrorTrap "CommonStrings::SortStrings()", Err
   Resume EXIT_SUB
       
End Sub


Public Function RemoveDupeStringsFromArray(strArray() As String) As String()
                                    
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::RemoveDupeStringsFromArray()"
   
   On Error GoTo ERROR

   Dim lngNumEntries As Long
   lngNumEntries = ArrayLen(strArray)
   
   If lngNumEntries > 0 Then
   
      Dim lngIndex1 As Long
      For lngIndex1 = 0 To lngNumEntries - 2
         
         Dim lngIndex2 As Long
         For lngIndex2 = lngIndex1 + 1 To lngNumEntries - 1
         
            If strArray(lngIndex1) <> vbNullString Then
               If strArray(lngIndex1) = strArray(lngIndex2) Then
                  strArray(lngIndex2) = vbNullString
               End If
            End If
         
         Next lngIndex2
         
      Next lngIndex1
   
   End If
   
   'now remove the vbNullStrings
   'first, count them
   Dim lngNumNullStrings As Long
   For lngIndex1 = 0 To lngNumEntries - 1
      If strArray(lngIndex1) = vbNullString Then
         lngNumNullStrings = lngNumNullStrings + 1
      End If
   Next lngIndex1
   
   If lngNumNullStrings > 0 Then
   
      'next, make a new array and copy into it only valid strings
      Dim strNewArray() As String
      ReDim strNewArray(lngNumEntries - lngNumNullStrings - 1)
      
      Dim lngNewIndex As Long
      
      For lngIndex1 = 0 To lngNumEntries - 1
         If strArray(lngIndex1) <> vbNullString Then
            strNewArray(lngNewIndex) = strArray(lngIndex1)
            lngNewIndex = lngNewIndex + 1
         End If
      Next lngIndex1
      
      RemoveDupeStringsFromArray = strNewArray
      
   Else
      RemoveDupeStringsFromArray = strArray
   End If
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::RemoveDupeStringsFromArray()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::RemoveDupeStringsFromArray()", Err
   Resume EXIT_FUNC
   
End Function


'The array should NOT be order dependent, this will mess up the order
Public Sub RemoveStringFromStringArray(strToRemove As String, _
                                       strToRemoveFrom() As String, _
                                       Optional lngArrayLen = -1)
                                    
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::RemoveStringFromStringArray()"
   
   On Error GoTo ERROR

   Dim lngNumElements As Long
   
   If lngArrayLen = -1 Then 'using the full array
      lngNumElements = ArrayLen(strToRemoveFrom())
   Else 'using just a sub-array ( see DynamicStringArray::Remove() )
      lngNumElements = lngArrayLen
   End If
   
   'special case to watch out for, only one element in the array
   If lngNumElements = 1 Then
      If strToRemove = strToRemoveFrom(0) Then
         If lngArrayLen = -1 Then ' only if not a sub-array
            Erase strToRemoveFrom
         Else
            strToRemoveFrom(0) = vbNullString
         End If
      End If
   Else
      
      Dim lngIndex As Long
      For lngIndex = 0 To lngNumElements - 1
         If strToRemove = strToRemoveFrom(lngIndex) Then
         
            'rejigger the array
            strToRemoveFrom(lngIndex) = strToRemoveFrom(lngNumElements - 1) ' put the last one where the one to remove is
            
            'trim the array
            If lngArrayLen = -1 Then
               ReDim Preserve strToRemoveFrom(lngNumElements - 2)
            Else
               ' using just part of an array, don't resize
               strToRemoveFrom(lngNumElements - 1) = vbNullString ' don't want dupe entries
            End If
         
            Exit For
            
         End If
      Next lngIndex
      
   End If
   
EXIT_SUB:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::RemoveStringFromStringArray()"
   Exit Sub
   
ERROR:
   StandardErrorTrap "CommonStrings::RemoveStringFromStringArray()", Err
   Resume EXIT_SUB
   
End Sub


Public Sub AddStringToStringArray(strToAdd As String, strToAddTo() As String)
                                    
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::AddStringToStringArray()"
   
   On Error GoTo ERROR

   If ArrayLen(strToAddTo) = 0 Then
      ReDim strToAddTo(0)
   Else
      ReDim Preserve strToAddTo(UBound(strToAddTo) + 1)
   End If
   
   strToAddTo(UBound(strToAddTo)) = strToAdd
   
EXIT_SUB:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::AddStringToStringArray()"
   Exit Sub
   
ERROR:
   StandardErrorTrap "CommonStrings::AddStringToStringArray()", Err
   Resume EXIT_SUB
   
End Sub


Public Function Combine2StringArrays(strNewArray() As String, strArray1() As String, strArray2() As String) As Boolean

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::Combine2StringArrays()"

   On Error GoTo ERROR
   
   Combine2StringArrays = False
   
   Dim lngArray1Len As Long
   lngArray1Len = ArrayLen(strArray1)
   
   Dim lngArray2Len As Long
   lngArray2Len = ArrayLen(strArray2)
   
   Dim lngNewArrayLen As Long
   lngNewArrayLen = lngArray1Len + lngArray2Len
   
   ReDim strNewArray(lngNewArrayLen - 1)
   
   Dim lngIndex As Long
   For lngIndex = 0 To lngArray1Len - 1
      strNewArray(lngIndex) = strArray1(lngIndex)
   Next lngIndex
   
   Dim lngIndex2 As Long
   For lngIndex2 = 0 To lngArray2Len - 1
      strNewArray(lngIndex + lngIndex2) = strArray2(lngIndex2)
   Next lngIndex2
   
   Combine2StringArrays = True
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::Combine2StringArrays()"
   Exit Function
   
ERROR:
   Erase strNewArray
   StandardErrorTrap "CommonStrings::Combine2StringArrays()", Err
   Resume EXIT_FUNC
   
End Function


'array1 = array1 + array2
Public Function Combine2StringArrays2(strArray1() As String, strArray2() As String) As Boolean

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::Combine2StringArrays2()"

   On Error GoTo ERROR
   
   Combine2StringArrays2 = False
   
   Dim lngArray2Len As Long
   lngArray2Len = ArrayLen(strArray2)
   
   If lngArray2Len > 0 Then
   
      Dim lngArray1Len As Long
      lngArray1Len = ArrayLen(strArray1)
      
      Dim lngNewArrayLen As Long
      lngNewArrayLen = lngArray1Len + lngArray2Len
      
      ReDim Preserve strArray1(lngNewArrayLen - 1)
      
      Dim lngIndex As Long
      For lngIndex = lngArray1Len To lngNewArrayLen - 1
         strArray1(lngIndex) = strArray2(lngIndex - lngArray1Len)
      Next lngIndex
      
   End If
   
   Combine2StringArrays2 = True
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::Combine2StringArrays2()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::Combine2StringArrays2()", Err
   Resume EXIT_FUNC
   
End Function


'newArray = array1 + array2
Public Function Combine2ByteArrays(bytArray1() As Byte, bytArray2() As Byte) As Byte()

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::Combine2ByteArrays()"

   On Error GoTo ERROR
   
   Dim lngArray1Len As Long
   lngArray1Len = ArrayLen(bytArray1)
   
   Dim lngArray2Len As Long
   lngArray2Len = ArrayLen(bytArray2)
   
   Dim lngNewArrayLen As Long
   lngNewArrayLen = lngArray1Len + lngArray2Len
   
   Dim bytNewArray() As Byte
   ReDim bytNewArray(lngNewArrayLen - 1)
   
   Dim lngIndex As Long
   For lngIndex = 0 To lngArray1Len - 1
      bytNewArray(lngIndex) = bytArray1(lngIndex)
   Next lngIndex
   
   Dim lngIndex2 As Long
   For lngIndex2 = 0 To lngArray2Len - 1
      bytNewArray(lngIndex + lngIndex2) = bytArray2(lngIndex2)
   Next lngIndex2
   
   Combine2ByteArrays = bytNewArray
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::Combine2ByteArrays()"
   Exit Function
   
ERROR:
   Erase bytNewArray
   StandardErrorTrap "CommonStrings::Combine2ByteArrays()", Err
   Resume EXIT_FUNC
   
End Function


' Look for the first appearance of a byte array within another byte array.
' Return the index of the first character found in the source array if found, return
' -1 if not found.
' Using a pretty naive search algorithm. If the size of the searches is getting into the
' low thousands, a more sophisticated algorithm is recommended to avoid performance problems.
' Both arrays are bounded only with valid information ( neither is a buffer partially-filled ).

Public Function FindByteArrayInByteArray(bytSourceArray() As Byte, _
                                         bytTargetArray() As Byte, _
                                         Optional lngStartPosition As Long = 0) As Long

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::FindByteArrayInByteArray()"

   On Error GoTo ERROR
   
   FindByteArrayInByteArray = -1
   
   Dim lngTargetArrayLen As Long
   Dim lngSourceArrayLen As Long
   
   lngTargetArrayLen = ArrayLen(bytTargetArray)
   lngSourceArrayLen = ArrayLen(bytSourceArray)
   
   'stupidity check 1
   If lngTargetArrayLen = 0 Or lngSourceArrayLen = 0 Then
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
                     "Error in CommonStrings::Combine2ByteArrays(), target or source array length is 0." & vbCrLf & _
                     "   source len = " & lngSourceArrayLen & ", target len = " & lngTargetArrayLen & "."
      GoTo EXIT_SUB
   End If
   
   'stupidity check 2
   If lngTargetArrayLen > lngSourceArrayLen Then
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CommonStrings::Combine2ByteArrays(), target array larger than source array."
      GoTo EXIT_SUB
   End If
   
   Dim lngTargetIndex As Long
   Dim lngSourceIndex As Long
   
   For lngSourceIndex = lngStartPosition To lngSourceArrayLen - lngTargetArrayLen
   
      If bytSourceArray(lngSourceIndex) = bytTargetArray(lngTargetIndex) Then 'a possible match
      
         Dim lngPossibleStartIndex As Long
         lngPossibleStartIndex = lngSourceIndex 'save in case this is the correct answer
         
         Do
                  
            lngSourceIndex = lngSourceIndex + 1
            lngTargetIndex = lngTargetIndex + 1
            
            If lngTargetIndex = lngTargetArrayLen Then ' we found it ( because we already seached the whole target array
               Exit Do
            End If
         
         Loop While bytSourceArray(lngSourceIndex) = bytTargetArray(lngTargetIndex)
                  
         If lngTargetIndex = lngTargetArrayLen Then  'yup, we found it
            FindByteArrayInByteArray = lngPossibleStartIndex
            Exit For
         Else
            lngTargetIndex = 0 ' reset for more searching
            ' don't bother setting lngSourceIndex, it is exactly where we want it
         End If
         
      End If
   
   Next lngSourceIndex
   
EXIT_SUB:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::FindByteArrayInByteArray()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::FindByteArrayInByteArray()", Err
   Resume EXIT_SUB

End Function


Public Function GetSocketState(Socket As Winsock) As String

   On Error GoTo ERROR

   If Socket Is Nothing Then
      gclsLogger.Log gclsLogger.LOG_WHAT_INFO, "CommonStrings::GetSocketState, socket is unitialized."
      GetSocketState = SOCKET_STATE_NA
   Else
   
      Dim intState As Integer
      intState = Socket.State
      
      Select Case intState
      
         Case sckOpen
            GetSocketState = SOCKET_STATE_OPEN
         
         Case sckClosed
            GetSocketState = SOCKET_STATE_CLOSED
         
         Case sckConnected
            GetSocketState = SOCKET_STATE_CONNECTED
         
         Case sckClosing
            GetSocketState = SOCKET_STATE_CLOSING
         
         Case sckError
            GetSocketState = SOCKET_STATE_ERROR
         
         Case Else
            GetSocketState = CStr(intState)
      
      End Select
      
   End If
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::GetSocketState()", Err
   GetSocketState = SOCKET_STATE_UNKNOWN
   Resume EXIT_FUNC

End Function


Public Function Min(lng1 As Long, lng2 As Long) As Long

   Min = IIf(lng1 <= lng2, lng1, lng2)

End Function


'if =, return lng2 ( doesn't really matter, though, does it? )
Public Function Max(lng1 As Long, lng2 As Long)

   Max = IIf(lng1 > lng2, lng1, lng2)

End Function


Public Function CopyLongIntoByteArray(lngLong As Long) As Byte()

   Dim bytData() As Byte
   ReDim bytData(3)
   
   CopyMemory bytData(0), lngLong, 4
   
   CopyLongIntoByteArray = bytData

End Function


Public Function CopyIntegerIntoByteArray(intInteger As Integer) As Byte()

   Dim bytData() As Byte
   ReDim bytData(1)
   
   CopyMemory bytData(0), intInteger, 2
   
   CopyIntegerIntoByteArray = bytData
   
End Function


Public Function CopyByteArrayIntoLong2(byt1 As Byte, byt2 As Byte, byt3 As Byte, byt4 As Byte) As Long

   Dim bytArray(3) As Byte
   
   bytArray(0) = byt1
   bytArray(1) = byt2
   bytArray(2) = byt3
   bytArray(3) = byt4

   CopyByteArrayIntoLong2 = CopyByteArrayIntoLong(bytArray)

End Function


Public Function CopyByteArrayIntoLong(bytData() As Byte) As Long

   CopyMemory CopyByteArrayIntoLong, bytData(0), 4

End Function


Public Function CopyByteArrayIntoInteger2(byt1 As Byte, byt2 As Byte) As Integer

   Dim bytArray(1) As Byte
   
   bytArray(0) = byt1
   bytArray(1) = byt2

   CopyByteArrayIntoInteger2 = CopyByteArrayIntoInteger(bytArray)

End Function


Public Function CopyByteArrayIntoInteger(bytData() As Byte) As Integer

   CopyMemory CopyByteArrayIntoInteger, bytData(0), 2

End Function


Public Function ConvertIntToBitField(intColumns As Integer, intNumber As Integer) As String
   
   On Error GoTo ERROR

   Dim intValue As Integer
   Dim intIndex As Integer
   
   For intIndex = intColumns To 1 Step -1
   
      intValue = intNumber And 1
      ConvertIntToBitField = CStr(intValue) & ConvertIntToBitField
      
      'shift right 1
      intNumber = intNumber \ 2
   
   Next
   
EXIT_FUNC:
   Exit Function

ERROR:
   StandardErrorTrap "CommonStrings::ConvertIntToBitField()", Err, 0, "   all ""0""'s returned."
   ConvertIntToBitField = vbNullString
   For intIndex = 1 To intColumns
      ConvertIntToBitField = ConvertIntToBitField & "0"
   Next
   Resume EXIT_FUNC

End Function


Public Sub RemoveNothingsFromArray(varArray As Variant)

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::RemoveNothingsFromArray()"
   
   On Error GoTo ERROR

   Dim lngNumOrigObjects As Long
   lngNumOrigObjects = UBound(varArray) + 1
   
   Dim lngNumNothings As Long
   
   'move the "nothings" to the bottom
   Dim lngIndex As Long
   For lngIndex = 0 To lngNumOrigObjects - 1
   
      If varArray(lngIndex) Is Nothing Then
      
         'the last one is a special case that won't be moved but will be counted
         If lngIndex < lngNumOrigObjects - 1 Then
         
            Dim lngIndex2 As Long
            For lngIndex2 = lngIndex To lngNumOrigObjects - 2
               Set varArray(lngIndex2) = varArray(lngIndex2 + 1) 'move everything "below" up one
            Next lngIndex2
            
         End If
            
         lngNumNothings = lngNumNothings + 1
      
      End If
   
   Next lngIndex
   
   'save the "non-nothings", trim away the "nothings"
   Dim lngNumReals As Long
   lngNumReals = lngNumOrigObjects - lngNumNothings
   
   If lngNumReals > 0 Then
      ReDim Preserve varArray(lngNumReals - 1)
   Else
      Erase varArray
   End If
   
EXIT_SUB:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::RemoveNothingsFromArray()"
   Exit Sub
   
ERROR:
   StandardErrorTrap "CommonStrings::RemoveNothingsFromArray()", Err
   Resume EXIT_SUB

End Sub


Public Function UniqueInArray(strBaseArray() As String, strTestArray() As String) As String()

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::UniqueInArray()"
   
   On Error GoTo ERROR
   
   Dim lngNumTestStrings As Long
   lngNumTestStrings = ArrayLen(strTestArray)
   
   Dim lngNumBaseStrings As Long
   lngNumBaseStrings = ArrayLen(strBaseArray)
   
   Dim lngIndexTest As Long
   Dim lngIndexBase As Long
   
   Dim blnFoundMatch As Boolean
   blnFoundMatch = False
   
   Dim strUniqueValues() As String
   Dim lngNumUniqueValues As Long
   
   For lngIndexTest = 0 To lngNumTestStrings - 1
      
      For lngIndexBase = 0 To lngNumBaseStrings - 1
      
         If strTestArray(lngIndexTest) = strBaseArray(lngIndexBase) Then
            blnFoundMatch = True
            Exit For
         End If
      
      Next lngIndexBase
   
      If Not blnFoundMatch Then
         ReDim Preserve strUniqueValues(lngNumUniqueValues)
         strUniqueValues(lngNumUniqueValues) = strTestArray(lngIndexTest)
         lngNumUniqueValues = lngNumUniqueValues = 1
      Else
         blnFoundMatch = False
      End If
   
   Next lngIndexTest
   
   If lngNumUniqueValues > 0 Then
      UniqueInArray = strUniqueValues
   End If
   
EXIT_SUB:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::UniqueInArray()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::UniqueInArray()", Err
   Resume EXIT_SUB

End Function


Public Function IsHexString(strMessage As String) As Boolean

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::IsHexString()"
   
   On Error GoTo ERROR
   
   IsHexString = False
   
   Dim lngNumChars As Long
   lngNumChars = Len(strMessage)
   
   If lngNumChars > 0 Then
      
      Dim lngIndex As Long
      For lngIndex = 1 To lngNumChars
      
         Dim strSingleChar As String
         strSingleChar = Mid$(strMessage, lngIndex, 1)
         
         Dim blnIsHex As Boolean
         blnIsHex = False
         
         Select Case strSingleChar
            Case "0" To "9", "a" To "f", "A" To "F"
               blnIsHex = True
         End Select
         
         If Not blnIsHex Then
            GoTo EXIT_FUNC
         End If
         
      Next lngIndex
      
      IsHexString = True
            
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CommonStrings::IsHexString(), no data."
   End If

EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::IsHexString()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::IsHexString()", Err
   Resume EXIT_FUNC

End Function


Public Function CopyStringArray(strOrig() As String) As String()

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonStrings::CopyStringArray()"
   
   On Error GoTo ERROR
   
   Dim lngUBound As Long
   lngUBound = UBound(strOrig)
   
   Dim strNew() As String
   ReDim strNew(lngUBound)
   
   Dim lngIndex
   For lngIndex = 0 To lngUBound
      strNew(lngIndex) = strOrig(lngIndex) ' I hope this is a real copy and not a ref assign
   Next lngIndex
   
   CopyStringArray = strNew
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonStrings::CopyStringArray()"
   Exit Function
   
ERROR:
   CopyStringArray = strOrig ' return reference, yuck
   StandardErrorTrap "CommonStrings::CopyStringArray()", Err
   Resume EXIT_FUNC
   
End Function


Public Function AllPrintableChars(strMessage As String) As Boolean

   AllPrintableChars = True
   
End Function


'Searches a target string for one of any number of characters
Public Function FoundChars(strSearchTarget As String, strTargetChars() As String) As Boolean

   Dim lngNumTargetChars As Long
   lngNumTargetChars = ArrayLen(strTargetChars)
   
   Dim lngIndex
   For lngIndex = 0 To lngNumTargetChars - 1
      If InStr(strSearchTarget, strTargetChars(lngIndex)) > 0 Then
         FoundChars = True
         Exit For
      End If
   Next lngIndex
   
End Function
