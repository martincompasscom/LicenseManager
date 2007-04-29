Attribute VB_Name = "CommonLogging"
Option Explicit

Public Const MAX_PATH = 260
Public Const ERROR_NO_MORE_FILES = 18&

Public Const DEFAULT_LOGGING_WHAT As Long = 16 'LOG_WHAT_ERROR as defined in clsLogger
Public Const DEFAULT_LOGGING_WHERE As Long = 2 'LOG_WHERE_STATUS as defined in clsLogger
Public Const DEFAULT_DAYS_LOG_FILE_HISTORY As Long = 3
Public Const MAX_DAYS_LOG_FILE As Long = 30

Public glngLoggingWhat As Long
Public glngLoggingWhere As Long
Public gstrLoggingFile As String
Public glngDaysLogFileHistory As Long
Public gclsLogger As Logger

Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type


Public Declare Function FindClose Lib "kernel32" _
   (ByVal hFindFile As Long) As Long
   
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
   (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
   (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long


Public Sub StandardErrorTrap(strLocation As String, _
                             objErr As ErrObject, _
                             Optional lngExtraFlags As Long = 0, _
                             Optional strExtraLine As String = vbNullString)
                             
   If Not gclsLogger Is Nothing Then
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR + lngExtraFlags, _
                     "Error in " & strLocation & vbCrLf & _
                     "   num = """ & objErr.Number & """, desc = """ & objErr.Description & """" & vbCrLf & _
                     strExtraLine
   Else
      WriteStatus "Error in " & strLocation
      WriteStatus "   num = """ & objErr.Number & """, desc = """ & objErr.Description & """"
      WriteStatus strExtraLine
   End If

End Sub


