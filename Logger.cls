VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intFileHandle As Integer

'predetermine these so we aren't constantly bit-twiddling
Private blnLogToFile As Boolean
Private blnLogToStatus As Boolean
Private blnLogError As Boolean
Private blnLogTrace As Boolean
Private blnLogInfo As Boolean
Private blnLogMessaging As Boolean
Private blnLogMessagingExtra As Boolean
Private blnLogDataStructures As Boolean
Private blnLogDatabase As Boolean
Private blnLogSockets As Boolean

'logging state
Private blnLoggingToFile As Boolean

'Pseudo-constants - These should be defined in this file
Public LOG_WHAT_TRACE As Long
Public LOG_WHAT_INFO As Long
Public LOG_WHAT_ERROR As Long
Public LOG_WHAT_MESSAGING As Long
Public LOG_WHAT_MESSAGING_EXTRA As Long
Public LOG_WHAT_DATA_STRUCTURES As Long
Public LOG_WHAT_DATABASE As Long
Public LOG_WHAT_SOCKETS As Long

Public LOG_WHERE_STATUS As Long
Public LOG_WHERE_FILE As Long

Private Const PRIVATE_LOG_WHAT_TRACE As Long = 1
Private Const PRIVATE_LOG_WHAT_INFO As Long = 2
Private Const PRIVATE_LOG_WHAT_ERROR As Long = 4
Private Const PRIVATE_LOG_WHAT_MESSAGING As Long = 8
Private Const PRIVATE_LOG_WHAT_MESSAGING_EXTRA As Long = 16 'generally timed events
Private Const PRIVATE_LOG_DATA_STRUCTURES As Long = 32 'watching for data in and out of data structures
Private Const PRIVATE_LOG_DATABASE As Long = 64 'watching for data in and out of data structures
Private Const PRIVATE_LOG_SOCKETS As Long = 128 'watching for data in and out of data structures

Private Const PRIVATE_LOG_WHERE_STATUS As Long = 1
Private Const PRIVATE_LOG_WHERE_FILE As Long = 2

Private strCurrDate As String
Private datCurrDate As Date
Private fsoFileSystem As FileSystemObject

Private Const DATE_FORMAT_STRING As String = "mm-dd-yyyy"
   

Private Sub Class_Initialize()

   'initialize the "constants"
   LOG_WHAT_TRACE = PRIVATE_LOG_WHAT_TRACE
   LOG_WHAT_INFO = PRIVATE_LOG_WHAT_INFO
   LOG_WHAT_ERROR = PRIVATE_LOG_WHAT_ERROR
   LOG_WHAT_MESSAGING = PRIVATE_LOG_WHAT_MESSAGING
   LOG_WHAT_MESSAGING_EXTRA = PRIVATE_LOG_WHAT_MESSAGING_EXTRA
   LOG_WHAT_DATA_STRUCTURES = PRIVATE_LOG_DATA_STRUCTURES
   LOG_WHAT_DATABASE = PRIVATE_LOG_DATABASE
   LOG_WHAT_SOCKETS = PRIVATE_LOG_SOCKETS
   
   LOG_WHERE_STATUS = PRIVATE_LOG_WHERE_STATUS
   LOG_WHERE_FILE = PRIVATE_LOG_WHERE_FILE
   
   blnLoggingToFile = False
   
   Set fsoFileSystem = New FileSystemObject
   
   datCurrDate = Date$
   strCurrDate = Format$(datCurrDate, DATE_FORMAT_STRING)

End Sub


Private Sub Class_Terminate()

   If blnLoggingToFile Then
      StopLoggingToFile
   End If
   
   Set fsoFileSystem = Nothing
   
End Sub


Public Sub Initialize()

   blnLogTrace = IIf((glngLoggingWhat And LOG_WHAT_TRACE) = LOG_WHAT_TRACE, True, False)
   blnLogInfo = IIf((glngLoggingWhat And LOG_WHAT_INFO) = LOG_WHAT_INFO, True, False)
   blnLogError = IIf((glngLoggingWhat And LOG_WHAT_ERROR) = LOG_WHAT_ERROR, True, False)
   blnLogMessaging = IIf((glngLoggingWhat And LOG_WHAT_MESSAGING) = LOG_WHAT_MESSAGING, True, False)
   blnLogMessagingExtra = IIf((glngLoggingWhat And LOG_WHAT_MESSAGING_EXTRA) = LOG_WHAT_MESSAGING_EXTRA, True, False)
   blnLogDataStructures = IIf((glngLoggingWhat And LOG_WHAT_DATA_STRUCTURES) = LOG_WHAT_DATA_STRUCTURES, True, False)
   blnLogDatabase = IIf((glngLoggingWhat And LOG_WHAT_DATABASE) = LOG_WHAT_DATABASE, True, False)
   blnLogSockets = IIf((glngLoggingWhat And LOG_WHAT_SOCKETS) = LOG_WHAT_SOCKETS, True, False)
   
   blnLogToStatus = IIf((glngLoggingWhere And LOG_WHERE_STATUS) = LOG_WHERE_STATUS, True, False)
   blnLogToFile = IIf((glngLoggingWhere And LOG_WHERE_FILE) = LOG_WHERE_FILE, True, False)
   
   If blnLogToFile Then 'start logging to file
      If Not StartLoggingToFile Then
         blnLogToFile = False
      End If
   End If
   
End Sub


Public Function StartLoggingToFile() As Boolean

   On Error GoTo ERROR
   
   StartLoggingToFile = False

   If Not blnLoggingToFile Then
   
      If Len(gstrLoggingFile) > 0 Then
      
         Dim strFullFileName As String
         strFullFileName = GetCurrentLogFileName()
         
         intFileHandle = FreeFile()
               
         On Error Resume Next
         If InStr(strFullFileName, ":") > 0 Then 'absolute path, no mods necessary
            Open strFullFileName For Append Lock Write As #intFileHandle
         Else 'relative or non-based path
            strFullFileName = fsoFileSystem.GetAbsolutePathName(strFullFileName)
            Open strFullFileName For Append Lock Write As #intFileHandle
         End If
         
         If Err.Number <> 0 Then
            If Err.Number = 76 Then ' path not found, try to make it;
               
               Dim lngLastSeparator As Long
               lngLastSeparator = InStrRev(strFullFileName, "\")
               
               If lngLastSeparator > 0 Then
                  MsgBox "Logging path """ & Left$(strFullFileName, lngLastSeparator) & """ does not exist. Logging to file is disabled."""
               Else 'do the best we can
                  MsgBox "Logging path """ & strFullFileName & """ does not exist. Logging to file is disabled."""
               End If
            
               StopLoggingToFile
               blnLogToFile = False
               GoTo EXIT_FUNC
               
            Else
               StopLoggingToFile
               blnLogToFile = False
               GoTo ERROR
            End If
         End If
      
         blnLoggingToFile = True
         
         'just making sure we can open and close successfully
         Close #intFileHandle
         
         'reopen
         Open strFullFileName For Append Lock Write As #intFileHandle
         
         'time and date stamp opening
         Print #intFileHandle, strCurrDate & SPACE_CHAR & Time$ & SPACE_CHAR & "Opened log" & vbCrLf & vbCrLf
         
      Else
         GoTo EXIT_FUNC
      End If
      
   End If

   StartLoggingToFile = True
   
EXIT_FUNC:
   Exit Function
   
ERROR:
   WriteStatus "Error in Logger::StartLoggingToFile()" & MSG_SEPARATOR & _
               "   num = """ & Err.Number & """, desc = """ & Err.Description & """" & MSG_SEPARATOR & _
               "   filename = """ & gstrLoggingFile & """ or """ & strFullFileName & """"
   Resume EXIT_FUNC

End Function


Public Sub StopLoggingToFile()

   On Error GoTo ERROR
   
   If blnLogToFile Then
   
      ' in the case where we couldn't open a file and caught an error, this func is called. Set this value first
      ' so that there will be no more attempts to log to file. The following calls will throw an error, but it
      ' will be caught and all will be rosy after that.
      blnLoggingToFile = False
      
      'time and date stamp closing
      Print #intFileHandle, strCurrDate & SPACE_CHAR & Time$ & SPACE_CHAR & "Closed log" & vbCrLf & vbCrLf
      Close #intFileHandle
      
   End If
   
EXIT_SUB:
   Exit Sub
   
ERROR:
   WriteStatus "Error in Logger::StopLoggingToFile()" & MSG_SEPARATOR & _
               "   num = """ & Err.Number & """, desc = """ & Err.Description & """."
   Resume EXIT_SUB

End Sub


Private Sub RestartLogging()

   StopLoggingToFile
   StartLoggingToFile
   
End Sub


'Standard logging call.
'  lngLogLevel is the requested logging level.
Public Sub Log(lngLogLevel As Long, strMessage As String)

   On Error GoTo ERROR
         
   If ((lngLogLevel And LOG_WHAT_INFO) = LOG_WHAT_INFO And blnLogInfo) Or _
      ((lngLogLevel And LOG_WHAT_TRACE) = LOG_WHAT_TRACE And blnLogTrace) Or _
      ((lngLogLevel And LOG_WHAT_MESSAGING) = LOG_WHAT_MESSAGING And blnLogMessaging) Or _
      ((lngLogLevel And LOG_WHAT_MESSAGING_EXTRA) = LOG_WHAT_MESSAGING_EXTRA And blnLogMessagingExtra) Or _
      ((lngLogLevel And LOG_WHAT_ERROR) = LOG_WHAT_ERROR And blnLogError) Or _
      ((lngLogLevel And LOG_WHAT_DATA_STRUCTURES) = LOG_WHAT_DATA_STRUCTURES And blnLogDataStructures) Or _
      ((lngLogLevel And LOG_WHAT_DATABASE) = LOG_WHAT_DATABASE And blnLogDatabase) Or _
      ((lngLogLevel And LOG_WHAT_SOCKETS) = LOG_WHAT_SOCKETS And blnLogSockets) Then

      If blnLogToStatus Then
         WriteStatus strMessage
      End If
      
      If Date$ <> datCurrDate Then
         RolloverLogFile
      End If
      
      If blnLogToFile Then
         Print #intFileHandle, strCurrDate & SPACE_CHAR & Time$ & SPACE_CHAR & strMessage & vbCrLf & vbCrLf
     End If
      
   End If
            
EXIT_SUB:
   Exit Sub
            
ERROR:
   WriteStatus "Error in FormLogging::Log()" & MSG_SEPARATOR & _
               "   num = """ & Err.Number & """,   desc = """ & Err.Description & """" & MSG_SEPARATOR & _
               "   handle = """ & intFileHandle & """, filename = """ & gstrLoggingFile & """"
   Resume EXIT_SUB
   
End Sub


'This call writes to the destination specified, as long as the user allowed it in the config file
Public Sub LogDirect1(lngLogLevel As Long, lngDestFlag As Long, strMessage As String)

   On Error GoTo ERROR
   
   If ((lngLogLevel And LOG_WHAT_INFO) = LOG_WHAT_INFO And blnLogInfo) Or _
      ((lngLogLevel And LOG_WHAT_TRACE) = LOG_WHAT_TRACE And blnLogTrace) Or _
      ((lngLogLevel And LOG_WHAT_MESSAGING) = LOG_WHAT_MESSAGING And blnLogMessaging) Or _
      ((lngLogLevel And LOG_WHAT_MESSAGING_EXTRA) = LOG_WHAT_MESSAGING_EXTRA And blnLogMessagingExtra) Or _
      ((lngLogLevel And LOG_WHAT_ERROR) = LOG_WHAT_ERROR And blnLogError) Or _
      ((lngLogLevel And LOG_WHAT_DATA_STRUCTURES) = LOG_WHAT_DATA_STRUCTURES And blnLogDataStructures) Or _
      ((lngLogLevel And LOG_WHAT_DATABASE) = LOG_WHAT_DATABASE And blnLogDatabase) Or _
      ((lngLogLevel And LOG_WHAT_SOCKETS) = LOG_WHAT_SOCKETS And blnLogSockets) Then
            
      If (lngDestFlag And LOG_WHERE_FILE) = LOG_WHERE_FILE Then
         LogDirect2 strMessage
      End If
      
      If (lngDestFlag And LOG_WHERE_STATUS) = LOG_WHERE_STATUS Then
         If blnLogToStatus Then
            WriteStatus strMessage
         End If
      End If
   
   End If
      
EXIT_SUB:
   Exit Sub
   
ERROR:
   WriteStatus "Error in FormLogging::LogDirect()" & MSG_SEPARATOR & _
               "   num = """ & Err.Number & """,   desc = """ & Err.Description & """"
   Resume EXIT_SUB

End Sub


'This call ignores all logging flags and only writes to the file, if it is open.
Public Sub LogDirect2(strMessage As String)

   On Error GoTo ERROR
               
   If blnLogToFile Then
   
      If Date$ <> datCurrDate Then
         RolloverLogFile
      End If
      
      Print #intFileHandle, strCurrDate & SPACE_CHAR & Time$ & SPACE_CHAR & strMessage & vbCrLf & vbCrLf
      
   End If
      
EXIT_SUB:
   Exit Sub
   
ERROR:
   WriteStatus "Error in FormLogging::LogDirect2()" & MSG_SEPARATOR & _
               "   num = """ & Err.Number & """,   desc = """ & Err.Description & """"
   Resume EXIT_SUB

End Sub


'Put enough characters into the file buffer to flush out any useful data from the buffer to the file.
'Assuming that the file buffer is 2k.
Public Sub Flush()

   If blnLogToFile Then 'just being careful
   
      Dim strFlusher As String
      strFlusher = String$(2048, SPACE_CHAR) & vbCrLf
      
      Print #intFileHandle, strFlusher
      
   End If

End Sub


Public Function GetLoggingToFile() As Boolean

   GetLoggingToFile = blnLogToFile
   
End Function


Public Sub SetLoggingToFile(blnNewLogToFile As Boolean)

   If blnLogToFile <> blnNewLogToFile Then 'it changed
   
      blnLogToFile = blnNewLogToFile
      
      If blnLogToFile Then
         glngLoggingWhere = glngLoggingWhere Or LOG_WHERE_FILE
      Else
         glngLoggingWhere = glngLoggingWhere And (Not LOG_WHERE_FILE)
      End If
      
   End If
      
End Sub


Public Function GetLoggingToUI() As Boolean

   GetLoggingToUI = blnLogToStatus
   
End Function


Public Sub SetLoggingToUI(blnNewLogToStatus As Boolean)

   If blnLogToStatus <> blnNewLogToStatus Then
   
      blnLogToStatus = blnNewLogToStatus
      
      If blnLogToStatus Then
         glngLoggingWhere = glngLoggingWhere Or LOG_WHERE_STATUS
      Else
         glngLoggingWhere = glngLoggingWhere And (Not LOG_WHERE_STATUS)
      End If
      
   End If
         
End Sub


Public Function GetTraceLogging() As Boolean

   GetTraceLogging = blnLogTrace
   
End Function


Public Sub SetTraceLogging(blnNewLogTrace As Boolean)

   If blnLogTrace <> blnNewLogTrace Then
   
      blnLogTrace = blnNewLogTrace
      
      If blnLogTrace Then
         glngLoggingWhat = glngLoggingWhat Or LOG_WHAT_TRACE
      Else
         glngLoggingWhat = glngLoggingWhat And (Not LOG_WHAT_TRACE)
      End If
      
   End If
   
End Sub


Public Function GetInfoLogging() As Boolean

   GetInfoLogging = blnLogInfo
   
End Function


Public Sub SetInfoLogging(blnNewLogInfo As Boolean)

   If blnLogInfo <> blnNewLogInfo Then
   
      blnLogInfo = blnNewLogInfo
      
      If blnLogInfo Then
         glngLoggingWhat = glngLoggingWhat Or LOG_WHAT_INFO
      Else
         glngLoggingWhat = glngLoggingWhat And (Not LOG_WHAT_INFO)
      End If
      
   End If
   
End Sub


Public Function GetMessagingLogging() As Boolean

   GetMessagingLogging = blnLogMessaging
   
End Function


Public Sub SetMessagingLogging(blnNewLogMessaging As Boolean)

   If blnLogMessaging <> blnNewLogMessaging Then
   
      blnLogMessaging = blnNewLogMessaging
      
      If blnLogMessaging Then
         glngLoggingWhat = glngLoggingWhat Or LOG_WHAT_MESSAGING
      Else
         glngLoggingWhat = glngLoggingWhat And (Not LOG_WHAT_MESSAGING)
      End If
      
   End If
   
End Sub


Public Function GetVerboseMessagingLogging() As Boolean

   GetVerboseMessagingLogging = blnLogMessagingExtra
   
End Function


Public Sub SetVerboseMessagingLogging(blnNewLogMessagingExtra As Boolean)

   If blnLogMessagingExtra <> blnNewLogMessagingExtra Then
   
      blnLogMessagingExtra = blnNewLogMessagingExtra
      
      If blnLogMessagingExtra Then
         glngLoggingWhat = glngLoggingWhat Or LOG_WHAT_MESSAGING_EXTRA
      Else
         glngLoggingWhat = glngLoggingWhat And (Not LOG_WHAT_MESSAGING_EXTRA)
      End If
      
   End If
   
End Sub


Public Function GetErrorLogging() As Boolean
   
   GetErrorLogging = blnLogError
   
End Function


Public Sub SetErrorLogging(blnNewLogError As Boolean)
   
   If blnLogError <> blnNewLogError Then
   
      blnLogError = blnNewLogError
      
      If blnLogError Then
         glngLoggingWhat = glngLoggingWhat Or LOG_WHAT_ERROR
      Else
         glngLoggingWhat = glngLoggingWhat And (Not LOG_WHAT_ERROR)
      End If
      
   End If
   
End Sub


Public Function GetDataStructuresLogging() As Boolean

   GetDataStructuresLogging = blnLogDataStructures
   
End Function


Public Sub SetDataStructuresLogging(blnNewLogDataStructures As Boolean)

   If blnLogDataStructures <> blnNewLogDataStructures Then
   
      blnLogDataStructures = blnNewLogDataStructures
      
      If blnLogDataStructures Then
         glngLoggingWhat = glngLoggingWhat Or LOG_WHAT_DATA_STRUCTURES
      Else
         glngLoggingWhat = glngLoggingWhat And (Not LOG_WHAT_DATA_STRUCTURES)
      End If
      
   End If
   
End Sub


Public Function GetDatabaseLogging() As Boolean

   GetDatabaseLogging = blnLogDatabase
   
End Function


Public Sub SetDatabaseLogging(blnNewLogDatabase As Boolean)

   If blnLogDatabase <> blnNewLogDatabase Then
   
      blnLogDatabase = blnNewLogDatabase
      
      If blnLogDatabase Then
         glngLoggingWhat = glngLoggingWhat Or LOG_WHAT_DATABASE
      Else
         glngLoggingWhat = glngLoggingWhat And (Not LOG_WHAT_DATABASE)
      End If
      
   End If
   
End Sub


Public Function GetSocketsLogging() As Boolean

   GetSocketsLogging = blnLogSockets
   
End Function


Public Sub SetSocketsLogging(blnNewLogSockets As Boolean)

   If blnLogSockets <> blnNewLogSockets Then
   
      blnLogSockets = blnNewLogSockets
      
      If blnLogSockets Then
         glngLoggingWhat = glngLoggingWhat Or LOG_WHAT_SOCKETS
      Else
         glngLoggingWhat = glngLoggingWhat And (Not LOG_WHAT_SOCKETS)
      End If
      
   End If
   
End Sub


'merely insert the date into the file name suggested by the user
Private Function GetCurrentLogFileName() As String

   If datCurrDate <> Date$ Then
      datCurrDate = Date$
      strCurrDate = Format$(datCurrDate, DATE_FORMAT_STRING)
   End If

   Dim lngPoint As Long
   lngPoint = InStrRev(gstrLoggingFile, ".")
   
   If lngPoint > 0 Then
      GetCurrentLogFileName = Left$(gstrLoggingFile, lngPoint - 1) & _
                              "_" & _
                              strCurrDate & _
                              Mid$(gstrLoggingFile, lngPoint)
   Else 'no point, just append
      GetCurrentLogFileName = gstrLoggingFile & strCurrDate
   End If

End Function


Private Sub RolloverLogFile()

   On Error GoTo ERROR

   datCurrDate = Date$
   strCurrDate = Format$(datCurrDate, DATE_FORMAT_STRING)
   
   'close current file, open new one
   RestartLogging
   
   'get all appropriate log files
   Dim colFileNames As Collection
   Set colFileNames = GetAllLogFiles()
   
   If Not colFileNames Is Nothing Then
   
      'count how many log files exist
      Dim lngNumLogFiles As Long
      lngNumLogFiles = colFileNames.Count()
      
      'pare count down to the correct value
      Dim lngNumFilesToDelete As Long
      lngNumFilesToDelete = (lngNumLogFiles - 1) - glngDaysLogFileHistory ' don't count current file
      
      If lngNumFilesToDelete > 0 Then
         
         'create an array for sorting
         Dim strFileNames() As String
         ReDim strFileNames(lngNumLogFiles - 1)
         
         Dim lngIndex As Long
         For lngIndex = 0 To lngNumLogFiles - 1
            strFileNames(lngIndex) = colFileNames.Item(lngIndex + 1)
         Next lngIndex
      
         'sort
         SortStrings strFileNames
      
         'the oldest ones are at the top
         For lngIndex = 0 To lngNumFilesToDelete - 1
            fsoFileSystem.DeleteFile strFileNames(lngIndex)
         Next lngIndex
         
      End If
      
   Else
      Log LOG_WHAT_ERROR, _
          "Error in FormLogging::RolloverLogFile()" & vbCrLf & _
          "   unable to get list of log file names"
   End If
      
EXIT_SUB:
   Erase strFileNames
   If Not colFileNames Is Nothing Then
      Set colFileNames = Nothing
   End If
   Exit Sub
   
ERROR:
   Log LOG_WHAT_ERROR, _
       "Error in FormLogging::RolloverLogFile()" & vbCrLf & _
       "   num = """ & Err.Number & """,   desc = """ & Err.Description & """"
   Resume EXIT_SUB

End Sub


Private Function GetAllLogFiles() As Collection

   On Error GoTo ERROR

   Dim lngPoint As Long
   lngPoint = InStrRev(gstrLoggingFile, ".")

   Dim strLogFileRoot As String
   Dim strLogFileExt As String
   If lngPoint > 0 Then
      strLogFileRoot = Left$(gstrLoggingFile, lngPoint - 1)
      strLogFileExt = Mid$(gstrLoggingFile, lngPoint)
   Else
      strLogFileRoot = gstrLoggingFile
      strLogFileExt = vbNullString
   End If
   
   'save this cuz findfiles will not give it back to you
   Dim strLogFilePath As String
   lngPoint = InStrRev(gstrLoggingFile, "\")
   If lngPoint > 0 Then
      strLogFilePath = Left$(gstrLoggingFile, lngPoint)
   Else
      strLogFilePath = vbNullString
   End If
   
   Dim strFindString As String
   strFindString = strLogFileRoot & "_??-??-????" & strLogFileExt
   
   Dim lngLenFileName As Long
   lngLenFileName = Len(strFindString)

   Dim lngFindHandle As Long
   Dim typFindData As WIN32_FIND_DATA
   
   lngFindHandle = FindFirstFile(strFindString, typFindData)
   If lngFindHandle <= 0 Then '2 = file not found
      Log LOG_WHAT_ERROR, _
          "Error in FormLogging::RolloverLogFile(), FindFirstFile" & vbCrLf & _
          "   num = """ & Err.LastDllError() & """"
      GoTo EXIT_FUNC
   End If
   
   Set GetAllLogFiles = New Collection
   GetAllLogFiles.Add strLogFilePath & typFindData.cFileName
   
   Dim lngRetVal As Long
   lngRetVal = FindNextFile(lngFindHandle, typFindData)
   
   Do While lngRetVal <> 0
      GetAllLogFiles.Add strLogFilePath & typFindData.cFileName
      lngRetVal = FindNextFile(lngFindHandle, typFindData)
   Loop
   
   'lngRetVal now = 0
   If Err.LastDllError() <> ERROR_NO_MORE_FILES Then
   
      Log LOG_WHAT_ERROR, _
          "Error in FormLogging::RolloverLogFile(), FindNextFile" & vbCrLf & _
          "   num = """ & Err.LastDllError() & """"
          
      Set GetAllLogFiles = Nothing
      GoTo EXIT_FUNC
   
   End If
   
   lngRetVal = FindClose(lngFindHandle)
   If lngRetVal = 0 Then
      Log LOG_WHAT_ERROR, _
          "Error in FormLogging::RolloverLogFile(), FindClose" & vbCrLf & _
          "   num = """ & Err.LastDllError() & """"
   End If
      
EXIT_FUNC:
   Exit Function
   
ERROR:
   If Not GetAllLogFiles Is Nothing Then
      Set GetAllLogFiles = Nothing
   End If
   Log LOG_WHAT_ERROR, _
       "Error in FormLogging::GetAllLogFiles()" & vbCrLf & _
       "   num = """ & Err.Number & """,   desc = """ & Err.Description & """"
   Resume EXIT_FUNC

End Function


'this function will only work if we are logging to file
Public Sub DumpDictionaryToLog(strName As String, _
                               map As Dictionary, _
                               Optional blnSupportsGetString As Boolean = False)

   If blnLoggingToFile Then
      If map.Count() > 0 Then
   
         Dim strMessage As String
         strMessage = "Dictionary " & strName & " has count of " & map.Count() & vbCrLf
   
         Dim varKeys As Variant
         varKeys = map.Keys()
         
         Dim key As Variant
         For Each key In varKeys
            If blnSupportsGetString Then
               strMessage = strMessage & "   key = """ & key & """" & _
                                         "   val = """ & map.Item(key).GetString() & """" & vbCrLf
            Else
               strMessage = strMessage & "   key = """ & key & """"
               
               On Error Resume Next 'catch if there is nothing as an item, a normal case btw
               Dim strTemp As String
               strTemp = CStr(map.Item(key))
               If Err.Number = 0 Then 'no error
                  strMessage = strMessage & "   val = """ & strTemp & """" & vbCrLf
               Else
                  strMessage = strMessage & "   val = unknown or unprintable" & vbCrLf
               End If
            End If
         Next
         
         LogDirect2 strMessage
   
      Else
         LogDirect2 "Dictionary " & strName & " has count of 0."
      End If
   End If
   
End Sub


'this function will only work if we are logging to file
Public Sub DumpCollectionToLog(strName As String, _
                               col As Collection, _
                               Optional blnSupportsGetString As Boolean = False)

   If blnLoggingToFile Then
      If col.Count() > 0 Then
   
         Dim strMessage As String
         strMessage = "Collection " & strName & " has count of " & col.Count() & vbCrLf
   
         Dim varItem As Variant
         
         For Each varItem In col
            If blnSupportsGetString Then
               strMessage = strMessage & "   """ & varItem.GetString() & """" & vbCrLf
            Else
            
               On Error Resume Next 'catch if there is nothing as an item, a normal case btw
               Dim strTemp As String
               strTemp = CStr(varItem)
               
               If Err.Number = 0 Then 'no error
                  strMessage = strMessage & "   """ & CStr(varItem) & """" & vbCrLf
               Else
                  strMessage = strMessage & "   unknown Or unprintable" & vbCrLf
               End If
               
            End If
         Next
                  
         LogDirect2 strMessage
   
      Else
         LogDirect2 "Collection " & strName & " has count of 0."
      End If
   End If
   
End Sub


'Just clear the log file out
Public Sub Clear()

   On Error GoTo ERROR

   If blnLoggingToFile Then
   
      'just making sure we can open and close successfully
      Close #intFileHandle
   
      Dim strFullFileName As String
      strFullFileName = GetCurrentLogFileName()
   
      'open in such a way as to delete all data
      Open strFullFileName For Output Lock Write As #intFileHandle
   
      'close so we can reopen in append mode
      Close #intFileHandle
   
      'reopen
      Open strFullFileName For Append Lock Write As #intFileHandle
      
      'time and date stamp opening
      Print #intFileHandle, strCurrDate & SPACE_CHAR & Time$ & SPACE_CHAR & "Opened log" & vbCrLf & vbCrLf
      
   End If
   
EXIT_SUB:
   Exit Sub
   
ERROR:
   Log LOG_WHAT_ERROR, _
       "Error in FormLogging::Clear()" & vbCrLf & _
       "   num = """ & Err.Number & """,   desc = """ & Err.Description & """"
   Resume EXIT_SUB

End Sub
