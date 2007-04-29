Attribute VB_Name = "Common"
Option Explicit

Public Const DATE_SEPARATOR As String = "/"
Public Const STD_DATE_FORMAT As String = "mm" & DATE_SEPARATOR & "dd" & DATE_SEPARATOR & "yyyy"


Public Sub WriteStatus(strMessage As String)

   MsgBox strMessage

End Sub


Public Sub WriteMessage(strMessage As String)

   MsgBox strMessage

End Sub

