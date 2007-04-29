Attribute VB_Name = "CC_Position"
Option Explicit

Public Const POSITION_REPORT As String = "CC_Position"

Public Const UNKNOWN_ALL_DISCRETES As String = "99999999"
Public Const UNKNOWN_FOUR_DISCRETES As String = "9999"
Public Const UNKNOWN_SINGLE_DISCRETE As String = "9"

Private Const DISCRETES_OPEN_TAG As String = "<Discretes>"
Private Const LEN_DISCRETES_OPEN_TAG As Long = 11
Private Const DISCRETES_CLOSE_TAG As String = "</Discretes>"

Private Const UNITID_TAG_OPEN As String = "<UnitID>"
Private Const UNITID_TAG_CLOSE As String = "</UnitID>"
Private Const UNITID_TAG_OPEN_LEN As Long = 8

Public Enum PosMsgFields
   POS_MSG_TIME = 0
   POS_MSG_DATE = 1
   POS_MSG_LATITUDE = 2
   POS_MSG_LONGITUDE = 3
   POS_MSG_ALTITUDE = 4
   POS_MSG_SPEED = 5
   POS_MSG_HEADING = 6
   POS_MSG_DISCRETES = 7
   POS_MSG_UNIT_ID = 8
   POS_MSG_DATASOURCE = 9
End Enum

Public Const NUM_POS_MSG_FIELDS As Integer = 10

Private Const strPositionDTD As String = _
   "<?xml version=""1.0""?>" & vbCrLf & _
   "<!DOCTYPE CC_Position [" & vbCrLf & _
   "<!ELEMENT CC_Position (UnitID, Date, Time, Latitude, Longitude, Altitude?, Speed, Heading, DataSource?, Discretes?)>" & vbCrLf & _
   "<!ATTLIST CC_Position Version CDATA #REQUIRED SubType (REALTIME | HISTORIC) #REQUIRED>" & vbCrLf & _
   "<!ELEMENT UnitID (#PCDATA)>" & vbCrLf & _
   "<!ELEMENT Date (#PCDATA)>" & vbCrLf & _
   "<!ELEMENT Time (#PCDATA)>" & vbCrLf & _
   "<!ELEMENT Latitude (#PCDATA)>" & vbCrLf & _
   "<!ELEMENT Longitude (#PCDATA)>" & vbCrLf & _
   "<!ELEMENT Altitude (#PCDATA)>" & vbCrLf & _
   "<!ELEMENT Speed (#PCDATA)>" & vbCrLf & _
   "<!ELEMENT Heading (#PCDATA)>" & vbCrLf & _
   "<!ELEMENT DataSource (#PCDATA)>" & vbCrLf & _
   "<!ELEMENT Discretes (#PCDATA)>" & vbCrLf & _
   "]>"

Private Const strMsgPositionFrag1 As String = _
   "<CC_Position Version=""1.0"" SubType="""
   
Private Const strMsgPositionFrag2  As String = _
                                             """>" & vbCrLf & _
   "   <UnitID>"
      
Private Const strMsgPositionFrag3 As String = _
               "</UnitID>" & vbCrLf & _
   "   <Date>"
      
Private Const strMsgPositionFrag4 As String = _
               "</Date>" & vbCrLf & _
   "   <Time>"
      
Private Const strMsgPositionFrag5 As String = _
               "</Time>" & vbCrLf & _
   "   <Latitude>"
      
Private Const strMsgPositionFrag6 As String = _
                  "</Latitude>" & vbCrLf & _
   "   <Longitude>"
      
Private Const strMsgPositionFrag7 As String = _
                  "</Longitude>" & vbCrLf & _
   "   <Altitude>"
      
Private Const strMsgPositionFrag8 As String = _
                  "</Altitude>" & vbCrLf & _
   "   <Speed>"
      
Private Const strMsgPositionFrag9 As String = _
               "</Speed>" & vbCrLf & _
   "   <Heading>"
      
Private Const strMsgPositionFrag10 As String = _
               "</Heading>" & vbCrLf & _
   "   <DataSource>"
      
Private Const strMsgPositionFrag11 As String = _
               "</DataSource>" & vbCrLf & _
   "   <Discretes>"
      
Private Const strMsgPositionFrag12 As String = _
                  "</Discretes>" & vbCrLf & _
   "</CC_Position>" & vbCrLf


Public Function CreatePositionMessage(pType As MessageType, _
                                      strUnitID As String, _
                                      strDate As String, _
                                      strTime As String, _
                                      strLon As String, _
                                      strLat As String, _
                                      strAltitude As String, _
                                      strSpeed As String, _
                                      strHeading As String, _
                                      strDataSource As String, _
                                      strDiscretes As String, _
                                      Optional blnLineFeeds As Boolean = True) As String

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Position::CreatePositionMessage()"
                                      
   On Error GoTo ERROR
   
   CreatePositionMessage = strMsgPositionFrag1 & _
                           IIf(pType = REALTIME_X, TYPE_REALTIME, TYPE_HISTORIC) & _
                           strMsgPositionFrag2 & strUnitID & _
                           strMsgPositionFrag3 & strDate & _
                           strMsgPositionFrag4 & strTime & _
                           strMsgPositionFrag5 & strLat & _
                           strMsgPositionFrag6 & strLon & _
                           strMsgPositionFrag7 & strAltitude & _
                           strMsgPositionFrag8 & strSpeed & _
                           strMsgPositionFrag9 & strHeading & _
                           strMsgPositionFrag10 & strDataSource & _
                           strMsgPositionFrag11 & strDiscretes & _
                           strMsgPositionFrag12
                           
   If Not blnLineFeeds Then
      CreatePositionMessage = Replace(CreatePositionMessage, vbCrLf, vbNullString)
   End If
    
EXIT_FUNC:
   If gclsLogger.GetVerboseMessagingLogging() Then
      gclsLogger.Log gclsLogger.LOG_WHAT_MESSAGING_EXTRA, "CC_Position::CreatePositionMessage(), msg =" & vbCrLf & """" & CreatePositionMessage & """"
   End If
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Position::CreatePositionMessage()"
   Exit Function
   
ERROR:
   CreatePositionMessage = vbNullString
   StandardErrorTrap "CC_Position::CreatePositionMessage()", Err
   Resume EXIT_FUNC
         
End Function


'Assumption: array is the correct size and has data
Public Function CreatePositionMessage2(pType As MessageType, _
                                      strDataFields() As String, _
                                      Optional blnLineFeeds As Boolean = True) As String

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Position::CreatePositionMessage2()"
                                      
   On Error GoTo ERROR
   
   CreatePositionMessage2 = strMsgPositionFrag1 & _
                            IIf(pType = REALTIME_X, TYPE_REALTIME, TYPE_HISTORIC) & _
                            strMsgPositionFrag2 & strDataFields(POS_MSG_UNIT_ID) & _
                            strMsgPositionFrag3 & strDataFields(POS_MSG_DATE) & _
                            strMsgPositionFrag4 & strDataFields(POS_MSG_TIME) & _
                            strMsgPositionFrag5 & strDataFields(POS_MSG_LATITUDE) & _
                            strMsgPositionFrag6 & strDataFields(POS_MSG_LONGITUDE) & _
                            strMsgPositionFrag7 & strDataFields(POS_MSG_ALTITUDE) & _
                            strMsgPositionFrag8 & strDataFields(POS_MSG_SPEED) & _
                            strMsgPositionFrag9 & strDataFields(POS_MSG_HEADING) & _
                            strMsgPositionFrag10 & strDataFields(POS_MSG_DATASOURCE) & _
                            strMsgPositionFrag11 & strDataFields(POS_MSG_DISCRETES) & _
                            strMsgPositionFrag12
                           
   ' not good performance, but this very rarely is needed
   If Not blnLineFeeds Then
      CreatePositionMessage2 = Replace(CreatePositionMessage2, vbCrLf, vbNullString)
   End If
    
EXIT_FUNC:
   If gclsLogger.GetVerboseMessagingLogging() Then
      gclsLogger.Log gclsLogger.LOG_WHAT_MESSAGING_EXTRA, "CC_Position::CreatePositionMessage2(), msg =" & vbCrLf & """" & CreatePositionMessage2 & """"
   End If
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Position::CreatePositionMessage2()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CC_Position::CreatePositionMessage2()", Err
   Resume EXIT_FUNC
         
End Function


'If it came from a parser, all fields will be filled in. It is the parser's responsibility
'to put default values in unused fields
Public Function ExtractPositionData(strMessage As String, _
                                    ByRef strUnitID As String, _
                                    ByRef strDate As String, _
                                    ByRef strTime As String, _
                                    ByRef strLon As String, _
                                    ByRef strLat As String, _
                                    ByRef strAltitude As String, _
                                    ByRef strSpeed As String, _
                                    ByRef strHeading As String, _
                                    ByRef strDataSource As String, _
                                    ByRef strDiscretes As String) As Boolean

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Position::ExtractPositionData()"
                                    
   On Error GoTo ERROR

   Dim domDoc As DOMDocument30
   Set domDoc = New DOMDocument30
   
   domDoc.async = False
   domDoc.setProperty "SelectionLanguage", "XPath"
   
   If domDoc.loadXML(strMessage) Then
      
      strUnitID = Trim$(domDoc.documentElement.selectSingleNode("UnitID").Text())
      strDate = Trim$(domDoc.documentElement.selectSingleNode("Date").Text())
      strTime = Trim$(domDoc.documentElement.selectSingleNode("Time").Text())
      strLon = Trim$(domDoc.documentElement.selectSingleNode("Longitude").Text())
      strLat = Trim$(domDoc.documentElement.selectSingleNode("Latitude").Text())
      strAltitude = Trim$(domDoc.documentElement.selectSingleNode("Altitude").Text())
      strSpeed = Trim$(domDoc.documentElement.selectSingleNode("Speed").Text())
      strHeading = Trim$(domDoc.documentElement.selectSingleNode("Heading").Text())
      strDataSource = Trim$(domDoc.documentElement.selectSingleNode("DataSource").Text())
      strDiscretes = Trim$(domDoc.documentElement.selectSingleNode("Discretes").Text())
      
      ExtractPositionData = True
      
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
                     "Error in CC_Position::ExtractPositionData(), loading document" & vbCrLf & _
                     "   num = """ & domDoc.parseError.errorCode & """, desc = """ & domDoc.parseError.reason & """."
      ExtractPositionData = False
   End If
      
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Position::ExtractPositionData()"
   If Not domDoc Is Nothing Then
      Set domDoc = Nothing
   End If
   Exit Function
   
ERROR:
   StandardErrorTrap "CC_Position::ExtractPositionData()", Err
   Resume EXIT_FUNC
   
End Function


Public Function ExtractVehicleID(strMessage As String) As String

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Position::ExtractVehicleID()"

   On Error GoTo ERROR

   ExtractVehicleID = ExtractSingleField(strMessage, _
                                         UNITID_TAG_OPEN, _
                                         UNITID_TAG_CLOSE, _
                                         UNITID_TAG_OPEN_LEN)
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Position::ExtractVehicleID()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CC_Position::ExtractVehicleID()", Err
   Resume EXIT_FUNC
   
End Function


Public Function ExtractDiscretes(strMessage As String) As String

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Position::ExtractDiscretes()"

   On Error GoTo ERROR

   Dim domDoc As DOMDocument30
   Set domDoc = New DOMDocument30
   
   domDoc.async = False
   domDoc.setProperty "SelectionLanguage", "XPath"
   
   If domDoc.loadXML(strMessage) Then
      ExtractDiscretes = Trim$(domDoc.documentElement.selectSingleNode("Discretes").Text())
   Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
                     "Error in CC_Position::ExtractDiscretes() loading document" & vbCrLf & _
                     "   num = """ & domDoc.parseError.errorCode & """, desc = """ & domDoc.parseError.reason & """."
      ExtractDiscretes = vbNullString
   End If
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Position::ExtractDiscretes()"
   If Not domDoc Is Nothing Then
      Set domDoc = Nothing
   End If
   Exit Function
   
ERROR:
   StandardErrorTrap "CC_Position::ExtractDiscretes()", Err
   Resume EXIT_FUNC
   
End Function


Public Function InsertDiscretes(strMessage As String, strNewDiscretes As String) As Boolean

   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "In CC_Position::InsertDiscretes()"

   On Error GoTo ERROR
   
   InsertDiscretes = False

   Dim lngStart As Long
   Dim lngEnd As Long
   
   lngStart = InStr(strMessage, DISCRETES_OPEN_TAG) + LEN_DISCRETES_OPEN_TAG
   lngEnd = InStrRev(strMessage, DISCRETES_CLOSE_TAG)

   'we're doing trims because I'm unsure of what whitespace lurks. An extra
   'vbCrLf could cause the xml parser to throw an error ( one is okay, 2+ is bad )
   Dim strFirstSubString As String
   strFirstSubString = RTrimWS(Left$(strMessage, lngStart - 1))
      
   Dim strLastSubString As String
   strLastSubString = LTrimWS(Mid$(strMessage, lngEnd))
   
   strMessage = strFirstSubString & strNewDiscretes & strLastSubString
   
   InsertDiscretes = True
   
EXIT_FUNC:
   gclsLogger.Log gclsLogger.LOG_WHAT_TRACE, "Out CC_Position::InsertDiscretes()"
   Exit Function
   
ERROR:
   StandardErrorTrap "CC_Position::InsertDiscretes()", Err
   Resume EXIT_FUNC
   
End Function





Public Sub DisplayCurrentPositionData(strMessage As String)

   On Error GoTo ERROR

   'extract the fields from the xml
   Dim strUnitID As String
   Dim strDate As String
   Dim strTime As String
   Dim strLat As String
   Dim strLon As String
   Dim strAltitude As String
   Dim strSpeed As String
   Dim strHeading As String
   Dim strDataSource As String
   Dim strDiscretes As String
   
   Dim blnGotPositionData As Boolean
   
   blnGotPositionData = ExtractPositionData(strMessage, _
                                            strUnitID, _
                                            strDate, _
                                            strTime, _
                                            strLon, _
                                            strLat, _
                                            strAltitude, _
                                            strSpeed, _
                                            strHeading, _
                                            strDataSource, _
                                            strDiscretes)
                                            
   If blnGotPositionData Then
      WritePositionDataImpl strUnitID, _
                            strDate, _
                            strTime, _
                            strLon, _
                            strLat, _
                            strAltitude, _
                            strSpeed, _
                            strHeading, _
                            strDataSource, _
                            strDiscretes
   End If
   
EXIT_SUB:
   Exit Sub
   
ERROR:
   StandardErrorTrap "CC_Position::DisplayCurrentPositionData()", Err
   Resume EXIT_SUB
   
End Sub


Public Sub DisplayCurrentPositionData2(strUnitID As String, _
                                       strDate As String, _
                                       strTime As String, _
                                       strLon As String, _
                                       strLat As String, _
                                       strAltitude As String, _
                                       strSpeed As String, _
                                       strHeading As String, _
                                       strDataSource As String, _
                                       strDiscretes As String)

   On Error GoTo ERROR
   
   WritePositionDataImpl strUnitID, _
                         strDate, _
                         strTime, _
                         strLon, _
                         strLat, _
                         strAltitude, _
                         strSpeed, _
                         strHeading, _
                         strDataSource, _
                         strDiscretes
   
EXIT_SUB:
   Exit Sub
   
ERROR:
   StandardErrorTrap "CC_Position::DisplayCurrentPositionData2()", Err
   Resume EXIT_SUB
   
End Sub


Public Sub DisplayCurrentPositionData3(strData() As String)

   On Error GoTo ERROR
   
   WritePositionDataImpl strData(POS_MSG_UNIT_ID), _
                         strData(POS_MSG_DATE), _
                         strData(POS_MSG_TIME), _
                         strData(POS_MSG_LONGITUDE), _
                         strData(POS_MSG_LATITUDE), _
                         strData(POS_MSG_ALTITUDE), _
                         strData(POS_MSG_SPEED), _
                         strData(POS_MSG_HEADING), _
                         strData(POS_MSG_DATASOURCE), _
                         strData(POS_MSG_DISCRETES)
   
EXIT_SUB:
   Exit Sub
   
ERROR:
   StandardErrorTrap "CC_Position::DisplayCurrentPositionData3()", Err
   Resume EXIT_SUB
   
End Sub


Private Sub WritePositionDataImpl(strUnitID As String, _
                                  strDate As String, _
                                  strTime As String, _
                                  strLon As String, _
                                  strLat As String, _
                                  strAltitude As String, _
                                  strSpeed As String, _
                                  strHeading As String, _
                                  strDataSource As String, _
                                  strDiscretes As String)

   WriteMessage strUnitID & SPACE_CHAR & _
                strDate & SPACE_CHAR & _
                strTime & SPACE_CHAR & _
                strLon & SPACE_CHAR & _
                strLat & SPACE_CHAR & _
                strAltitude & SPACE_CHAR & _
                strSpeed & SPACE_CHAR & _
                strHeading & SPACE_CHAR & _
                strDataSource & SPACE_CHAR & _
                strDiscretes

End Sub


