Attribute VB_Name = "CommonAll"
Option Explicit

Public strVersionNumber As String

Public gstrInstanceName As String
Public glngConnectionInterval As Long
Public glngHeartbeatInterval As Long
Public gblnShowMessages As Boolean
Public gblnSuccessfulLoadConfigValues As Boolean
Public gblnRewriteConfigFile As Boolean
Public gclsXMLParser As XMLParser

Public glngParserPort As Long
Public glngParserMessagePort As Long
Public glngRealtimePort As Long
Public glngConfigPort As Long
Public glngHistoryPort As Long
Public glngMessagePort As Long

Public Const UNKNOWN_LATITUDE As String = "99.9999"
Public Const UNKNOWN_LONGITUDE As String = "999.9999"
Public Const UNKNOWN_ALTITUDE As String = "-99999"
Public Const UNKNOWN_SPEED As String = "9999"
Public Const UNKNOWN_HEADING As String = "9999"
Public Const UNKNOWN_DATASOURCE As String = "9"
Public Const UNKNOWN_TIME As String = "99999"
Public Const UNKNOWN_DATE As String = "99/99/9999"
Public Const UNKNOWN_UNIT_ID As String = "unknown"

Public Const MAX_INTEGER_VALUE As Integer = 32767
Public Const MAX_LONG_VALUE As Long = 2147483647

Public Const COMMA_SPACE_CHARS As String = ", "
Public Const COMMA_CHAR As String = ","
Public Const SPACE_CHAR As String = " "
Public Const HEARTBEAT_MESSAGE As String = ">heartbeat<"
Public Const MSG_SEPARATOR As String = vbNullChar
Public Const TRUE_STRING As String = "true"
Public Const FALSE_STRING As String = "false"

Public Const MAX_PORT As Long = 65535
Public Const MIN_PORT As Long = 1
Public Const DEFAULT_SHOW_MESSAGES As Boolean = True
Public Const DEFAULT_CONNECTION_INTERVAL As Long = 15
Public Const MIN_CONNECTION_INTERVAL As Long = 5
Public Const DEFAULT_HEARTBEAT_INTERVAL As Long = 15
Public Const MIN_HEARTBEAT_INTERVAL As Long = 5

Public Const END_OF_DAY = 86399
Public Const BEGINNING_OF_DAY = 0
Public Const SECONDS_PER_DAY As Long = 86400
Public Const SECONDS_PER_HOUR As Long = 3600
Public Const SECONDS_PER_MINUTE As Long = 60

'This is copied from CommonDB.bas
Public Const MAX_MESSAGE_ID_LEN As Integer = 30

Private Const MESSAGE_NUM_LINES As Long = 50
Private Const MESSAGE_LEN_LINES As Long = 87

Private blnCreatedBasePartOfMsgID As Boolean ' initial value is "false"
Private lngUnique As Long

Public Type typeFileTime

   lLowDateTime As Long
   lHighDateTime As Long
   
End Type

Public Type typeSystemTime

   iYear As Integer
   iMonth As Integer
   iDayOfWeek As Integer
   iDay As Integer
   iHour As Integer
   iMinute As Integer
   iSecond As Integer
   iMilliSecond As Integer
   
End Type

Declare Sub GetSystemTimeAsFileTime Lib "kernel32" _
   (ByRef FILETIME As typeFileTime)

Declare Function FileTimeToSystemTime Lib "kernel32" _
   (ByRef FILETIME As typeFileTime, _
    ByRef systemTime As typeSystemTime) As Boolean

Public Declare Sub GetLocalTime Lib "kernel32" (ByRef systemTime As typeSystemTime)


Public Function GetCurrentSecondsAfterMidnight() As Long

   On Error GoTo ERROR
   
   Dim strNow As String
   strNow = Format$(Now(), "hh:nn:ss")
   
   Dim strData() As String
   strData = Split(strNow, ":")
   
   GetCurrentSecondsAfterMidnight = SecondsAfterMidnight(strData(0), strData(1), strData(2))
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   GetCurrentSecondsAfterMidnight = UNKNOWN_TIME
   StandardErrorTrap "CommonStrings::GetCurrentSecondsAfterMidnight()", Err
   Resume EXIT_FUNC
   
End Function


Public Function GetUTCSecondsAfterMidnight() As Long

   On Error GoTo ERROR
   
   Dim rawUTCTime As typeFileTime
   GetSystemTimeAsFileTime rawUTCTime
   
   'Now convert to something easier to understand
   Dim utcTime As typeSystemTime
   Dim intRetVal As Integer
   
   intRetVal = FileTimeToSystemTime(rawUTCTime, utcTime)
   
   If intRetVal <> 0 Then ' true
      GetUTCSecondsAfterMidnight = CLng(utcTime.iHour) * SECONDS_PER_HOUR + _
                                   CLng(utcTime.iMinute) * SECONDS_PER_MINUTE + _
                                   CLng(utcTime.iSecond)
   Else ' use possible non-UTC time
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CommonAll::GetUTCSecondsAfterMidnight(), unable to calculate UTC time, using local time."
      GetUTCSecondsAfterMidnight = GetCurrentSecondsAfterMidnight()
   End If
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   GetUTCSecondsAfterMidnight = UNKNOWN_TIME
   StandardErrorTrap "CommonStrings::GetCurrentSecondsAfterMidnight()", Err
   Resume EXIT_FUNC
   
End Function


Public Function GetCurrentDate() As String

   On Error GoTo ERROR
   
   GetCurrentDate = Format$(Now(), "MM/DD/YYYY")
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::GetCurrentDate()", Err
   GetCurrentDate = UNKNOWN_DATE
   Resume EXIT_FUNC
   
End Function


Public Function GetUTCDate() As String

   On Error GoTo ERROR

   Dim rawUTCTime As typeFileTime
   GetSystemTimeAsFileTime rawUTCTime
   
   'Now convert to something easier to understand
   Dim utcTime As typeSystemTime
   
   FileTimeToSystemTime rawUTCTime, utcTime
   
   Dim strDay As String
   Dim strMonth As String
   
   With utcTime
   
      strDay = CStr(.iDay)
      If Len(strDay) = 2 Then
         If Left$(strDay, 1) = "0" Then
            strDay = Right$(strDay, 1)
         End If
      End If
      
      strMonth = CStr(.iMonth)
      If Len(strMonth) = 2 Then
         If Left$(strMonth, 1) = "0" Then
            strMonth = Right$(strMonth, 1)
         End If
      End If
      
      GetUTCDate = strMonth & "/" & strDay & "/" & CStr(.iYear)
   
   End With
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonStrings::GetUTCDate()", Err
   GetUTCDate = "9/9/1999"
   Resume EXIT_FUNC
      
End Function


Public Function SecondsAfterMidnight(strHours As String, strMinutes As String, strSeconds As String) As Long

   On Error GoTo ERROR

   SecondsAfterMidnight = CLng(strHours) * 3600 + CLng(strMinutes) * 60 + CLng(strSeconds)
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   SecondsAfterMidnight = 0
   StandardErrorTrap "CommonStrings::SecondsAfterMidnight()", Err, 0, "   hours = """ & strHours & """, minutes = """ & strMinutes & """, seconds = """ & strSeconds & """."
   Resume EXIT_FUNC

End Function


Public Function GetHHMMSSFromSAM(lngSAM As Long) As String

   On Error GoTo ERROR
   
   
   If lngSAM < 0 Or lngSAM >= SECONDS_PER_DAY Then
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CommonStrings::GetHHMMSSFromSAM(), invalid ""seconds after midnight"" value entered."
   Else
   
      Dim lngHours As Long
      Dim lngMinutes As Long
      Dim lngSeconds As Long
      
      lngHours = Floor(lngSAM / SECONDS_PER_HOUR)
      lngMinutes = Floor((lngSAM - lngHours * SECONDS_PER_HOUR) / SECONDS_PER_MINUTE)
      lngSeconds = lngSAM - ((lngHours * SECONDS_PER_HOUR) + (lngMinutes * SECONDS_PER_MINUTE))
      
      'in order to pad to two decimals
      Dim strHours As String
      Dim strMinutes As String
      Dim strSeconds As String
      
      strHours = CStr(lngHours)
      strMinutes = CStr(lngMinutes)
      strSeconds = CStr(lngSeconds)
      
      If Len(strHours) = 1 Then
         strHours = "0" & strHours
      End If
      If Len(strMinutes) = 1 Then
         strMinutes = "0" & strMinutes
      End If
      If Len(strSeconds) = 1 Then
         strSeconds = "0" & strSeconds
      End If
      
      GetHHMMSSFromSAM = strHours & ":" & strMinutes & ":" & strSeconds
      
   End If
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   GetHHMMSSFromSAM = "00:00:00"
   StandardErrorTrap "CommonStrings::GetHHMMSSFromSAM()", Err, 0, "   value = """ & CStr(lngSAM) & """."
   Resume EXIT_FUNC

End Function


'Return the floor of the number
'(the highest whole number less than or equal to the number)
Public Function Floor(dblNumber As Double) As Long

   Floor = CLng(dblNumber)

   If Floor > dblNumber Then
      Floor = Floor - 1
   End If
    
End Function



' Message Delimiting - there is only a separator between messages. There is no
' concept of a "start" and "stop" delimiter. This means that the first message
' received will not be preceded by the separator string.
Public Function ExtractMessage(ByRef strAllChars As String, _
                               ByRef blnCallAgain As Boolean, _
                               Optional strSeparator As String = MSG_SEPARATOR) As String
                               
   On Error GoTo ERROR
   
   If Len(strAllChars) > 0 Then 'being defensive, should never happen

      blnCallAgain = False
      
      Dim lngSeparatorPos As Long
      lngSeparatorPos = InStr(strAllChars, strSeparator)
   
      If lngSeparatorPos > 0 Then 'we have a complete message
      
         ExtractMessage = Left$(strAllChars, lngSeparatorPos - 1)
         
         If lngSeparatorPos = Len(ExtractMessage) Then 'separator at end of single msg
            strAllChars = vbNullString
         Else
            strAllChars = Mid$(strAllChars, lngSeparatorPos + 1)
         End If
         
         'check to see if another complete message exists
         lngSeparatorPos = InStr(strAllChars, strSeparator)
      
         If lngSeparatorPos > 0 Then 'we have another complete message
            blnCallAgain = True
         End If
   
      End If
   
   End If
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonAll::ExtractMessage()", Err
   Resume EXIT_FUNC
   
End Function


' If neither a start or stop delimiter is found, the data is tossed.
Public Function ExtractMessage2(ByRef strAllChars As String, _
                                strStartDelim As String, _
                                strStopDelim As String) As String
                               
   On Error GoTo ERROR
   
   If Len(strAllChars) > 0 Then 'being defensive, should never happen
      
      Dim lngStartDelimPos As Long
      lngStartDelimPos = InStr(strAllChars, strStartDelim)
   
      If lngStartDelimPos > 0 Then
   
         Dim lngStopDelimPos As Long
         lngStopDelimPos = InStr(strAllChars, strStopDelim)
         
         If lngStopDelimPos > 0 Then ' we have a complete message
            ExtractMessage2 = Mid$(strAllChars, lngStartDelimPos, lngStopDelimPos - lngStartDelimPos + 1)
            strAllChars = Mid$(strAllChars, lngStopDelimPos + 1) 'trim the original string to get rid of new message
         Else
            gclsLogger.Log gclsLogger.LOG_WHAT_INFO, "CommonAll::ExtractMessage2(), no stop delimiter, will wait for more data."
         End If
         
      Else
         gclsLogger.Log gclsLogger.LOG_WHAT_INFO, "CommonAll::ExtractMessage2(), no start delimiter, tossing all data."
         strAllChars = vbNullString
      End If
   
   End If
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonAll::ExtractMessage2()", Err
   Resume EXIT_FUNC
   
End Function


'alright, it's **relatively** unique
Public Function GetUniqueLong() As Long

   GetUniqueLong = lngUnique
   lngUnique = IncrementLong(lngUnique)

End Function


' If this is giving you problems compiling, you need to declare a listbox called lstMessages in MainForm.
' Don't initialize it, though
' Sorry for these stray references, will be fixed eventually.
Public Sub WriteMessage(ByVal strMessage As String)

   If gblnShowMessages Then
      If Not MainForm.lstMessages Is Nothing Then ' it might not be used, so it will be unitialized.
         With MainForm.lstMessages
         
            Dim strLine As String
            
            Do While .ListCount > MESSAGE_NUM_LINES
               .RemoveItem MESSAGE_NUM_LINES
            Loop
            
            Do
            
               strLine = Left$(strMessage, MESSAGE_LEN_LINES)
               strMessage = Mid$(strMessage, MESSAGE_LEN_LINES + 1)
               strMessage = SPACE_CHAR & strMessage
               .AddItem strLine, 0
               
            Loop While strMessage <> SPACE_CHAR
            
         End With
      End If
   End If
    
End Sub


'creates a string from 1 to 10 chars in length, just a long
Public Function CreateUniqueNumericMsgID() As String

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonAll::CreateUniqueNumericMsgID()"
   
   On Error GoTo ERROR
   
   If Not blnCreatedBasePartOfMsgID Then
      InitializeNumericMessageID
   End If
      
   CreateUniqueNumericMsgID = Left$(CStr(lngUnique), MAX_MESSAGE_ID_LEN)
   lngUnique = IncrementLong(lngUnique) ' after using, increment
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonAll::CreateUniqueNumericMsgID()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonAll::CreateUniqueNumericMsgID()", Err
   CreateUniqueNumericMsgID = "999999999" ' eventually we will figure it out
   Resume EXIT_FUNC

End Function


'make sure this does not exceed MAX_MESSAGE_ID_LEN
Public Function CreateUniqueCComMessageID(strInstanceName As String) As String
   
   If Not blnCreatedBasePartOfMsgID Then
      InitializeNumericMessageID
   End If

   Dim lngInstanceNameLen As Long
   lngInstanceNameLen = Len(strInstanceName)
   
   Dim strUniqueLong As String
   strUniqueLong = CStr(lngUnique)
   
   lngUnique = IncrementLong(lngUnique) ' after using, increment
   
   Dim lngLongLen As Long
   lngLongLen = Len(strUniqueLong)
   
   CreateUniqueCComMessageID = strUniqueLong & SPACE_CHAR & Left$(strInstanceName, MAX_MESSAGE_ID_LEN - lngInstanceNameLen - 1)

End Function
   
   
Private Sub InitializeNumericMessageID()

   Randomize
   lngUnique = CLng(Rnd() * CSng(MAX_LONG_VALUE))
   blnCreatedBasePartOfMsgID = True

End Sub


'Here to deal with "rolling over" which causes an overflow error if not prevented
Public Function IncrementLong(lngCurrValue As Long) As Long

   On Error GoTo ERROR

   IncrementLong = lngCurrValue + 1
   
EXIT_FUNC:
   Exit Function
   
ERROR: 'only possible cause is attempted rollover
   IncrementLong = 0
   Resume EXIT_FUNC
   
End Function
   

'Here to deal with "rolling over" which causes an overflow error if not prevented
Public Function IncrementInteger(intCurrValue As Integer) As Integer

   On Error GoTo ERROR

   IncrementInteger = intCurrValue + 1
   
EXIT_FUNC:
   Exit Function
   
ERROR: 'only possible cause is attempted rollover
   IncrementInteger = 0
   Resume EXIT_FUNC
   
End Function
   

'Here to deal with "rolling over" which causes an overflow error if not prevented
Public Function IncrementByte(bytCurrValue As Byte) As Byte

   On Error GoTo ERROR

   IncrementByte = bytCurrValue + 1
   
EXIT_FUNC:
   Exit Function
   
ERROR: 'only possible cause is attempted rollover
   IncrementByte = 0
   Resume EXIT_FUNC
   
End Function


Public Function SendStringMessageViaSocket(soc As Winsock, _
                                           strMessage As String, _
                                           blnUDP As Boolean, _
                                           Optional blnUseMsgSeparator As Boolean = True) As Boolean

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonAll::SendStringMessageViaSocket()"
   
   gclsLogger.Log gclsLogger.LOG_WHAT_MESSAGING_EXTRA, _
                  "CommonAll::SendStringMessageViaSocket(), msg = """ & strMessage & """."
   
   On Error GoTo ERROR
   
   SendStringMessageViaSocket = False
   
   If blnUseMsgSeparator Then
      strMessage = strMessage & MSG_SEPARATOR
   End If
   
   If Not soc Is Nothing Then
      If soc.State = IIf(blnUDP, sckOpen, sckConnected) Then
         soc.SendData strMessage
         SendStringMessageViaSocket = True
      Else
         gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CommonAll::SendStringMessageViaSocket(), socket NOT connected/open, state = """ & CStr(soc.State()) & """."
      End If
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CommonAll::SendStringMessageViaSocket(), socket is null."
   End If
      
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonAll::SendStringMessageViaSocket()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonAll::SendStringMessageViaSocket()", Err
   Resume EXIT_FUNC

End Function


Public Function SendByteArrayMessageViaSocket(soc As Winsock, _
                                              bytMessage() As Byte, _
                                              blnUDP As Boolean) As Boolean

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CommonAll::SendByteArrayMessageViaSocket()"
   
   If gclsLogger.GetVerboseMessagingLogging() Then
      gclsLogger.Log gclsLogger.LOG_WHAT_MESSAGING_EXTRA, _
                     "CommonAll::SendByteArrayMessageViaSocket()" & vbCrLf & _
                     "   bytes = " & vbCrLf & """" & ConvertByteArrayToHexValueString(bytMessage) & """"
   End If
      
   On Error GoTo ERROR
   
   SendByteArrayMessageViaSocket = False
   
   If Not soc Is Nothing Then
      If soc.State = IIf(blnUDP, sckOpen, sckConnected) Then
         soc.SendData bytMessage
         SendByteArrayMessageViaSocket = True
      Else
         gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CommonAll::SendByteArrayMessageViaSocket(), socket NOT connected/open, state = """ & CStr(soc.State()) & """."
      End If
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CommonAll::SendByteArrayMessageViaSocket(), socket is null."
   End If
      
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CommonAll::SendByteArrayMessageViaSocket()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CommonAll::SendByteArrayMessageViaSocket()", Err
   Resume EXIT_FUNC

End Function

