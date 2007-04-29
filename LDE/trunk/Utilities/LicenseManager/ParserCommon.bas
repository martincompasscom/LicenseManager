Attribute VB_Name = "ParserCommon"
'
'   ParserCommon.bas
'
' General function shared by all of the parsers


Option Explicit

Public Const HEARTBEAT As String = ">heartbeat<"
Public Const SPACE_CHAR As String = " "
Public Const COMMA_SPACE As String = ", "

Public Const MSG_SEPARATOR As String = vbNullChar

Public Const MAX_SOCKET_VALUE As Integer = 32767
Public Const MIN_SOCKET_VALUE As Integer = 1

'constants common to all parsers
Public gblnShowMessages As Boolean
Public gstrInstanceName As String
Public gblnSuccessfulLoadConfigValues As Boolean
Public gclsFDESessions As clsFDESessions
Public gclsXMLParser As clsXmlParser

'constants for adapters
Public gstrMessagingFDE_ID As String

'logging
Public gclsLogger As clsLogger
Public glngLoggingWhere As Long
Public glngLoggingWhat As Long
Public gstrLoggingFile As String
Public glngDaysLogFileHistory As Long

Public Const DEFAULT_LOGGING_WHAT As Long = 4 'LOG_WHAT_ERROR as defined in clsLogger
Public Const DEFAULT_LOGGING_WHERE As Long = 1 'LOG_WHERE_STATUS as defined in clsLogger
Public Const DEFAULT_LOGGING_FILE As String = "c:\CompassFDE.log"

Public Const FIX_STATUS_2D_GPS As String = "0"
Public Const FIX_STATUS_2D_DGPS As String = "1"
Public Const FIX_STATUS_3D_GPS As String = "2"
Public Const FIX_STATUS_3D_DGPS As String = "3"
Public Const FIX_STATUS_DEAD_RECKONING As String = "6"
Public Const FIX_STATUS_DEGRADED_DEAD_RECKONING As String = "8"
Public Const FIX_STATUS_UNKNOWN As String = "9"

Public Enum Fields
   MSG_TIME = 0
   MSG_DATE = 1
   MSG_LATITUDE = 2
   MSG_LONGITUDE = 3
   MSG_ALTITUDE = 4
   MSG_SPEED = 5
   MSG_HEADING = 6
   MSG_DISCRETES = 7
   MSG_ID = 8
   MSG_DATASOURCE = 9
End Enum

Public Const NUM_MSG_FIELDS As Integer = 10

Public Const STATUS_CONNECTED As Integer = 100
Public Const STATUS_UNCONNECTED As Integer = 101
Public Const STATUS_CONNECTING As Integer = 102

Type CRITICAL_SECTION_TYPE

   Reserved1 As Long
   Reserved2 As Long
   Reserved3 As Long
   Reserved4 As Long
   Reserved5 As Long
   Reserved6 As Long
   
End Type

Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION_TYPE)
Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION_TYPE)
Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION_TYPE)
Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION_TYPE)

Type typeFileTime
   lLowDateTime As Long
   lHighDateTime As Long
End Type

'Used for the 'GetLocalTime' call
Type typeSystemTime

   iYear As Integer
   iMonth As Integer
   iDayOfWeek As Integer
   iDay As Integer
   iHour As Integer
   iMinute As Integer
   iSecond As Integer
   iMilliSecond As Integer
   
End Type

Declare Sub GetLocalTime Lib "kernel32" (ByRef systemTime As typeSystemTime)

Declare Sub GetSystemTimeAsFileTime Lib "kernel32" _
   (ByRef FILETIME As typeFileTime)

Declare Function FileTimeToSystemTime Lib "kernel32" _
   (ByRef FILETIME As typeFileTime, _
    ByRef systemTime As typeSystemTime) As Boolean

Private Declare Sub Sleep Lib "kernel32" _
   (ByVal lMilliseconds As Long)
   
Declare Function SendMessageByNum Lib "user32" _
   Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
   wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const INVALID_HANDLE_VALUE = -1
Public Const ERROR_NO_MORE_FILES = 18&

Private Const STATUS_NUM_LINES As Long = 200
Private Const MESSAGE_NUM_LINES As Long = 50

'whitespace constants used by RTrimWS() and LTrimWS()
Private Const ASC_SPACE As Integer = 32
Private Const ASC_TAB As Integer = 9
Private Const ASC_CR As Integer = 13
Private Const ASC_LF As Integer = 10
Private Const ASC_FF As Integer = 12
Private Const ASC_VT As Integer = 8
Private Const ASC_BS As Integer = 11

'file API stuff
'Public Const MOVEFILE_REPLACE_EXISTING = &H1
'Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
'Public Const FILE_BEGIN = 0
Public Const FILE_SHARE_READ As Long = &H1
'Public Const FILE_SHARE_WRITE = &H2
'Public Const CREATE_NEW = 1
Public Const OPEN_ALWAYS As Long = 3
Public Const GENERIC_WRITE As Long = &H40000000

Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long


'If we are not visible, write to the event log, not the UI
Public Sub WriteStatus(strMessage As String)

   Static lngWidth As Long
   
   If InStr(strMessage, vbNullChar) > 0 Then  'multiline message
   
      Dim strAllStrings() As String
      strAllStrings = Split(strMessage, vbNullChar)
      
      Dim intIndex As Integer
      For intIndex = 0 To UBound(strAllStrings)
         WriteImpl frmMain.lstStatus, strAllStrings(intIndex), STATUS_NUM_LINES, lngWidth
      Next intIndex
   
   Else
      WriteImpl frmMain.lstStatus, strMessage, STATUS_NUM_LINES, lngWidth
   End If
    
End Sub


'If we are not visible, don't write. At this point, this only
'displays single lines, so we don't check for multiline
Public Sub WriteMessage(ByVal strMessage As String)

   Static lngWidth As Long
   
   WriteImpl frmMain.lstMessages, strMessage, MESSAGE_NUM_LINES, lngWidth
    
End Sub


'Writing to the UI
Private Sub WriteImpl(lstBox As ListBox, _
                      strMessage As String, _
                      intMaxNumLines As Integer, _
                      ByRef lngCurrWidth As Long)

   lstBox.AddItem strMessage, 0
   
   Do While lstBox.ListCount > intMaxNumLines
      lstBox.RemoveItem intMaxNumLines
   Loop
   
   'Do we need a horizontal scroll bar?
   With frmMain
      If lngCurrWidth < .TextWidth(strMessage & "   ") / Screen.TwipsPerPixelX Then
      
         lngCurrWidth = .TextWidth(strMessage & "   ") / Screen.TwipsPerPixelX
          
         If .ScaleMode = vbTwips Then
            SendMessageByNum lstBox.hwnd, LB_SETHORIZONTALEXTENT, lngCurrWidth, 0
         End If
         
      End If
   End With

End Sub


Private Sub WriteLog(strMessage As String)

   App.LogEvent strMessage, vbLogEventTypeInformation

End Sub


Public Function GetDate() As String

   Dim rawUTCTime As typeFileTime
   GetSystemTimeAsFileTime rawUTCTime
   
   'Now convert to something easier to understand
   Dim utcTime As typeSystemTime
   
   FileTimeToSystemTime rawUTCTime, utcTime
   
   Dim strDay As String
   Dim strMonth As String
   
   With utcTime
   
      strDay = .iDay
      If Len(strDay) = 2 Then
         If Left$(strDay, 1) = "0" Then
            strDay = Right$(strDay, 1)
         End If
      End If
      
      strMonth = .iMonth
      If Len(strMonth) = 2 Then
         If Left$(strMonth, 1) = "0" Then
            strMonth = Right$(strMonth, 1)
         End If
      End If
      
      GetDate = strMonth & "/" & strDay & "/" & .iYear
   
   End With
      
End Function


'set up for 4 least significant digits
Public Function ConvertToBitField(intDiscretes As Integer) As String
   
   If intDiscretes < 0 Or intDiscretes > 15 Then
      WriteStatus "Invalid discretes value = """ & CStr(intDiscretes) & """"
      ConvertToBitField = "0000"
   Else
      Select Case intDiscretes
      
         Case 0
            ConvertToBitField = "0000"
         
         Case 1
            ConvertToBitField = "1000"
         
         Case 2
            ConvertToBitField = "0100"
         
         Case 3
            ConvertToBitField = "1100"
         
         Case 4
            ConvertToBitField = "0010"
         
         Case 5
            ConvertToBitField = "1010"
         
         Case 6
            ConvertToBitField = "0110"
         
         Case 7
            ConvertToBitField = "1110"
         
         Case 8
            ConvertToBitField = "0001"
         
         Case 9
            ConvertToBitField = "1001"
         
         Case 10
            ConvertToBitField = "0101"
         
         Case 11
            ConvertToBitField = "1101"
         
         Case 12
            ConvertToBitField = "0011"
         
         Case 13
            ConvertToBitField = "1011"
         
         Case 14
            ConvertToBitField = "0111"
         
         Case 15
            ConvertToBitField = "1111"
         
      End Select
   End If
   
End Function


Public Sub Pause(sTimeInSeconds As Single)

   Sleep sTimeInSeconds * 1000
   
End Sub


Public Sub PauseWithEvents(sTotalTime As Single, sIntervalTime As Single)

   Dim intNumIntervals As Integer
   intNumIntervals = sTotalTime / sIntervalTime + 1
   
   Dim intIndex
   For intIndex = 0 To intNumIntervals - 1
      Pause sIntervalTime
      DoEvents
   Next

End Sub

' Message Delimiting - there is only a separator between messages. There is no
' concept of a "start" and "stop" delimiter. This means that the first message
' received will not be preceded by the separator string.
Public Function ExtractMessage(ByRef strAllChars As String, _
                               ByRef blnCallAgain As Boolean, _
                               Optional strSeparator As String = MSG_SEPARATOR) As String

   blnCallAgain = False
   Dim lngSeparatorPos As Long
   
   If Len(strAllChars) > 0 Then 'being defensive, should never happen
   
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
   
End Function
