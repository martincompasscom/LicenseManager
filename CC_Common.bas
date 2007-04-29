Attribute VB_Name = "CC_Common"
Option Explicit


Public Enum MessageType
   REALTIME_X = 125
   HISTORIC_X = 126
End Enum

Public Const TYPE_REALTIME As String = "REALTIME"
Public Const TYPE_HISTORIC As String = "HISTORIC"
Public Const MM_ATTRIBUTE As String = " MM="""""

Public Const ACK_NONE As String = "none"
Public Const ACK_APP As String = "app"
Public Const ACK_USER As String = "user"

Public Const UNKNOWN_MSG As Integer = 12
Public Const INCIDENT_MSG As Integer = 13
Public Const STATUS_MSG As Integer = 14
Public Const TEXT_MSG As Integer = 15
Public Const SYSTEM_MSG As Integer = 16

Private Const ORIG_TAG_OPEN As String = "<Originator>"
Private Const ORIG_TAG_CLOSE As String = "</Originator>"
Private Const ORIG_TAG_OPEN_LEN As Long = 12


Public Function GetMessageType(strMessage As String) As Integer

   If InStr(strMessage, INCIDENT_MESSAGE) > 0 Then
      GetMessageType = INCIDENT_MSG
   ElseIf InStr(strMessage, STATUS_MESSAGE) > 0 Then
      GetMessageType = STATUS_MSG
   ElseIf InStr(strMessage, TEXT_MESSAGE) > 0 Then
      GetMessageType = TEXT_MSG
   ElseIf InStr(strMessage, SYSTEM_MESSAGE) > 0 Then
      GetMessageType = SYSTEM_MSG
   Else
      GetMessageType = UNKNOWN_MSG
   End If

End Function


'only checks that it is "well-formed" and "valid", I think. See XML docs for definitions
Public Function IsValidXMLMessage(strMessage As String) As Boolean
                                  
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Common::IsValidXMLMessage()"
                              
   IsValidXMLMessage = False
   
   Dim domDoc As DOMDocument30
   Set domDoc = New DOMDocument30
   
   domDoc.async = False
   domDoc.setProperty "SelectionLanguage", "XPath"
   
   If domDoc.loadXML(strMessage) Then
      IsValidXMLMessage = True
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
                     "Error in CC_ConfigResponse::ExtractConfigResponseData(), loading document" & vbCrLf & _
                     "   num = """ & domDoc.parseError.errorCode & """, desc = """ & domDoc.parseError.reason & """."
   End If
   

EXIT_FUNC:
   Set domDoc = Nothing
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Common::IsValidXMLMessage()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CC_Common::IsValidXMLMessage()", Err
   Resume EXIT_FUNC

End Function


'this function gets all of the common data for all explicitly routable messages.
Public Function ExtractCommonData(domDoc As DOMDocument30, _
                                  node As MSXML2.IXMLDOMNode, _
                                  nodes As MSXML2.IXMLDOMNodeList, _
                                  blnGetDestinations As Boolean, _
                                  podCommonData As podMsgCommon) As Boolean
                                  
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Common::ExtractCommonData()"
                              
   ExtractCommonData = False
   
   Dim strAttemptedNode As String

   On Error GoTo ERROR

   If blnGetDestinations Then
      If GetDestinationsImpl(domDoc, node, nodes, podCommonData) = False Then
         gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CC_Common::ExtractCommonData(), unable to extract destinations."
         GoTo EXIT_FUNC
      End If
   End If
                  
   strAttemptedNode = "MessageID"
   Set node = domDoc.documentElement.selectSingleNode(strAttemptedNode)
   If Len(node.Text()) = 0 Then
      GoTo EMPTY_NODE_ERROR
   Else
      podCommonData.strMessageID = node.Text()
   End If

   strAttemptedNode = "Originator"
   Set node = domDoc.documentElement.selectSingleNode(strAttemptedNode)
   If Len(node.Text()) = 0 Then
      GoTo EMPTY_NODE_ERROR
   Else
   
      podCommonData.strOriginator = node.Text()
         
      Dim strDummy As String
      
      strAttemptedNode = "./@MM"
      On Error Resume Next
      strDummy = node.selectSingleNode(strAttemptedNode).Text()
      If Err.Number <> 0 Then
         podCommonData.blnSentFromMM = False
      Else
         podCommonData.blnSentFromMM = True
      End If
      
      On Error GoTo ERROR
         
   End If

   strAttemptedNode = "Date"
   Set node = domDoc.documentElement.selectSingleNode(strAttemptedNode)
   If Len(node.Text()) = 0 Then
      GoTo EMPTY_NODE_ERROR
   Else
      podCommonData.strDate = node.Text()
   End If

   strAttemptedNode = "Time"
   Set node = domDoc.documentElement.selectSingleNode(strAttemptedNode)
   If Len(node.Text()) = 0 Then
      GoTo EMPTY_NODE_ERROR
   Else
      podCommonData.strTime = node.Text()
   End If
   
   ExtractCommonData = True

EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Common::ExtractCommonData()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CC_Common::ExtractCommonData()", Err, 0, "   attempted node = """ & strAttemptedNode & """"
   Resume EXIT_FUNC
   
EMPTY_NODE_ERROR:
   gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
                  "Error in CC_Common::ExtractCommonData(), empty node when attempting to get """ & strAttemptedNode & """"
   Resume EXIT_FUNC

End Function


Public Function StripDTD(strMessage As String) As String
                                  
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Common::StripDTD()"
   
   On Error GoTo ERROR
   
   Dim lngEndDTDPos As Long
   lngEndDTDPos = InStr(strMessage, "]>")
   
   If lngEndDTDPos > 0 Then
      StripDTD = Mid$(strMessage, lngEndDTDPos + 2)
      StripDTD = LTrimWS(StripDTD) ' there may or may not be vbcrlf, just be sure and get rid of it
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "CC_Common::StripDTD(), can't find DTD marker, is there a DTD?, msg = " & vbCrLf & strMessage
   End If
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Common::StripDTD()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CC_Common::StripDTD()", Err
   Resume EXIT_FUNC

End Function


Public Function GetDestinations(strMessage As String) As String()
                                  
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Common::GetDestinations()"
   
   On Error GoTo ERROR
   
   Dim domDoc As DOMDocument30
   Set domDoc = New DOMDocument30
   
   domDoc.async = False
   domDoc.setProperty "SelectionLanguage", "XPath"
   
   If Not domDoc.loadXML(strMessage) Then
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
                     "Error in CC_Common::GetDestinations() loading document" & vbCrLf & _
                     "   num = """ & domDoc.parseError.errorCode & """, desc = """ & domDoc.parseError.reason & """."
      GoTo EXIT_FUNC
   End If
   
   Dim podCommonData As podMsgCommon
   Set podCommonData = New podMsgCommon

   Dim node As MSXML2.IXMLDOMNode
   Dim nodes As MSXML2.IXMLDOMNodeList

   If GetDestinationsImpl(domDoc, node, nodes, podCommonData) = True Then
         
      Dim podDestMsgs() As podMsgDestination
      podDestMsgs = podCommonData.GetDestinations()
      
      Dim lngNumDests As Long
      lngNumDests = ArrayLen(podDestMsgs)
   
      If lngNumDests > 0 Then
            
         Dim strDests() As String
         ReDim strDests(lngNumDests - 1)
         
         Dim lngIndex As Long
         For lngIndex = 0 To lngNumDests - 1
            strDests(lngIndex) = podDestMsgs(lngIndex).strDestination
            Set podDestMsgs(lngIndex) = Nothing
         Next lngIndex
                  
         GetDestinations = strDests
         Erase podDestMsgs
         
      Else
         gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CC_Common::GetDestinations(), no destinations."
      End If
   
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CC_Common::GetDestinations(), unable to extract destinations."
   End If
   
EXIT_FUNC:
   If Not domDoc Is Nothing Then
      Set domDoc = Nothing
   End If
   If Not node Is Nothing Then
      Set node = Nothing
   End If
   If Not nodes Is Nothing Then
      Set nodes = Nothing
   End If
   If Not podCommonData Is Nothing Then
      Set podCommonData = Nothing
   End If
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Common::GetDestinations()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CC_Common::GetDestinations()", Err
   Resume EXIT_FUNC

End Function


Private Function GetDestinationsImpl(domDoc As DOMDocument30, _
                                     node As MSXML2.IXMLDOMNode, _
                                     nodes As MSXML2.IXMLDOMNodeList, _
                                     podCommonData As podMsgCommon) As Boolean
                                  
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Common::GetDestinationsImpl()"
   
   On Error GoTo ERROR
   
   GetDestinationsImpl = False
   
   Dim strAttemptedNode As String
   
   'direct destinations
   strAttemptedNode = "Destinations/Destination"
   Set nodes = domDoc.documentElement.selectNodes(strAttemptedNode)
   
   If Not nodes Is Nothing Then
   
      Set node = nodes.nextNode()
      
      If Not node Is Nothing Then
      
         'getting the destinations/acks first
         Dim destinationData() As podMsgDestination
         ReDim destinationData(nodes.length() - 1)
         
         Dim lngIndex As Long
         For lngIndex = 0 To nodes.length() - 1
         
            Dim typDestination As podMsgDestination
            Set typDestination = New podMsgDestination
         
            strAttemptedNode = "destination"
            typDestination.strDestination = node.selectSingleNode(".").Text()
            
            strAttemptedNode = "acknowledge"
            typDestination.strAck = node.selectSingleNode("./@Acknowledge").Text()
            
            If Len(typDestination.strDestination) = 0 Or Len(typDestination.strAck) = 0 Then
               gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, "Error in CC_Common::GetDestinationsImpl(), missing destination or destination ack."
               GoTo EXIT_FUNC
            End If
            
            Set destinationData(lngIndex) = typDestination
            
            Set node = nodes.nextNode()
            
         Next lngIndex
         
         podCommonData.SetDestinations destinationData
         
         GetDestinationsImpl = True
         
      End If
      
   End If
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Common::GetDestinationsImpl()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CC_Common::GetDestinationsImpl()", Err, 0, "   attempted node = """ & strAttemptedNode & """."
   Resume EXIT_FUNC

End Function


Public Function ExtractOriginator(strMessage As String) As String

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Common::ExtractOriginator()"
   
   ExtractOriginator = ExtractSingleField(strMessage, _
                                          ORIG_TAG_OPEN, _
                                          ORIG_TAG_CLOSE, _
                                          ORIG_TAG_OPEN_LEN)

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Common::ExtractOriginator()"
                                          
End Function


'This is a here so that we don't have to go through the huge process of extracting all data.
Public Function ExtractSingleField(strMessage As String, _
                                   strOpenTag As String, _
                                   strCloseTag As String, _
                                   lngOpenTagLen) As String

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Common::ExtractSingleField()"

   On Error GoTo ERROR
   
   Dim lngOpenTagStart As Long
   lngOpenTagStart = InStr(strMessage, strOpenTag)
   
   Dim lngCloseTagStart As Long
   lngCloseTagStart = InStr(strMessage, strCloseTag)
   
   Dim lngLenField As Long
   lngLenField = lngCloseTagStart - (lngOpenTagStart + lngOpenTagLen)

   ExtractSingleField = Trim$(Mid$(strMessage, lngOpenTagStart + lngOpenTagLen, lngLenField))
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Common::ExtractSingleField()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CC_Common::ExtractSingleField()", Err
   Resume EXIT_FUNC
                                          
End Function

