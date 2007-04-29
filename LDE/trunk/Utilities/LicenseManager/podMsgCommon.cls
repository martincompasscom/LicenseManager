VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "podMsgCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private podDestinations() As podMsgDestination
Public strMessageID As String
Public strOriginator As String
Public strDate As String
Public strTime As String
Public blnSentFromMM As Boolean


Private Sub Class_Initialize()

   blnSentFromMM = False
   
End Sub


Public Sub SetDestinations(podNewDestinations() As podMsgDestination)

   podDestinations = podNewDestinations

End Sub


Public Function GetDestinations() As podMsgDestination()

   GetDestinations = podDestinations

End Function


Public Function GetString() As String

   GetString = "MessageID: " & strMessageID & vbCrLf & _
               "Originator: " & strOriginator & vbCrLf & _
               "Date: " & strDate & vbCrLf & _
               "Time: " & strTime & vbCrLf & _
               "Sent from MM: " & CStr(blnSentFromMM) & vbCrLf
               
   Dim lngCount As Long
   lngCount = UBound(podDestinations) + 1
   
   Dim lngIndex As Long
   For lngIndex = 0 To lngCount - 1
   
      If lngIndex = 0 Then
         GetString = GetString & "Destinations: " & vbCrLf
      End If
      
      GetString = GetString & podDestinations(lngIndex).GetString()
      
   Next lngIndex

End Function


