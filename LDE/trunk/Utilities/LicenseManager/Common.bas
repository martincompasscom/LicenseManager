Attribute VB_Name = "Common"
Option Explicit

Public Const DATE_SEPARATOR As String = "/"
Public Const STD_DATE_FORMAT As String = "mm" & DATE_SEPARATOR & "dd" & DATE_SEPARATOR & "yyyy"

'**********
'These are defined because they are used in CC_Common, and we don't
'want to drag in these four extra files.

Public Const INCIDENT_MESSAGE = vbNullString 'originally defined in CC_INCIDENT
Public Const STATUS_MESSAGE = vbNullString 'originally defined in CC_STATUS
Public Const TEXT_MESSAGE = vbNullString 'originally defined in CC_TEXT
Public Const SYSTEM_MESSAGE = vbNullString 'originally defined in CC_SYSTEM

'**********


Public Sub WriteStatus(strMessage As String)

   MsgBox strMessage

End Sub

'
'Public Sub WriteMessage(strMessage As String)
'
'   MsgBox strMessage
'
'End Sub
'
