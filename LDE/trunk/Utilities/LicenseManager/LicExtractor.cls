VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LicExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private bytLicenseKey() As Byte
Private clsCRC As CRC16


Private Sub Class_Initialize()

   Set clsCRC = New CRC16

End Sub


Private Sub Class_Terminate()

   Set clsCRC = Nothing
   
End Sub


Public Function Extract(strLicenseString As String, _
                        ByRef datGenDate As Date, _
                        ByRef datExpDate As Date, _
                        ByRef intNumClients As Integer, _
                        ByRef intNumVehicles As Integer) As Boolean
                   
   On Error GoTo ERROR
   
   Extract = False
                                         
   Erase bytLicenseKey
   bytLicenseKey = ConvertHexValueStringToByteArray(strLicenseString)
   
   Debug.Assert ArrayLen(bytLicenseKey) = LIC_LENGTH
   
   'generate crc, and check validity
   Dim bytEmbeddedCRC() As Byte
   bytEmbeddedCRC = GetCRCDataFromArray()
   PutPlaceholdersInCRCSlots bytLicenseKey

   If ValidateCRC(bytLicenseKey, bytEmbeddedCRC) Then 'all's well
      'put CRC back in license string
      PlaceCRCDataInArray bytLicenseKey, bytEmbeddedCRC
   Else
#If LIC_MANAGER_PROJECT Then
      MsgBox "Invalid CRC in license string, license invalid"
#Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
                     "Error in LicenseExtractor::Extract()" & vbCrLf & _
                     "   invalid CRC in license string, license invalid"
#End If
      GoTo EXIT_FUNC
   End If
                  
   'extract coded values from array
   Dim bytGenDate() As Byte
   Dim bytExpDate() As Byte
   Dim bytNumClients() As Byte
   Dim bytNumVehicles() As Byte
   
   bytGenDate = GetGenDateFromArray()
   bytExpDate = GetExpDateFromArray()
   bytNumClients = GetNumClientsFromArray()
   bytNumVehicles = GetNumVehiclesFromArray()
   
   'decode coded values
   bytGenDate = Decode(bytGenDate)
   bytExpDate = Decode(bytExpDate)
   bytNumClients = Decode(bytNumClients)
   bytNumVehicles = Decode(bytNumVehicles)
   
   'convert to correct types
   datGenDate = GetDateFromByteArray(bytGenDate)
   datExpDate = GetDateFromByteArray(bytExpDate)
   intNumClients = GetIntegerFromByteArray(bytNumClients)
   intNumVehicles = GetIntegerFromByteArray(bytNumVehicles)
   
   Extract = True
   
EXIT_FUNC:
   Exit Function

ERROR:
#If LIC_MANAGER_PROJECT Then
   MsgBox "Error in LicenseExtractor::Extract()" & vbCrLf & _
          "   num = """ & Err.Number & """, desc = """ & Err.Description & """"
#Else
   gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
                  "Error in LicenseExtractor::Extract()" & vbCrLf & _
                  "   num = """ & Err.Number & """, desc = """ & Err.Description & """"
#End If
   Resume EXIT_FUNC

End Function



Private Function GetGenDateFromArray() As Byte()

   Dim bytGenDate() As Byte
   ReDim bytGenDate(NUM_DATE_FIELDS - 1)

   bytGenDate(0) = bytLicenseKey(POS_GEN_DATE_M1)
   bytGenDate(1) = bytLicenseKey(POS_GEN_DATE_M2)
   bytGenDate(2) = 0 'just a place holder
   bytGenDate(3) = bytLicenseKey(POS_GEN_DATE_D1)
   bytGenDate(4) = bytLicenseKey(POS_GEN_DATE_D2)
   bytGenDate(5) = 0 'just a place holder
   bytGenDate(6) = bytLicenseKey(POS_GEN_DATE_Y1)
   bytGenDate(7) = bytLicenseKey(POS_GEN_DATE_Y2)
   bytGenDate(8) = bytLicenseKey(POS_GEN_DATE_Y3)
   bytGenDate(9) = bytLicenseKey(POS_GEN_DATE_Y4)
   
   GetGenDateFromArray = bytGenDate

End Function


Private Function GetExpDateFromArray() As Byte()

   Dim bytExpDate() As Byte
   ReDim bytExpDate(NUM_DATE_FIELDS - 1)

   bytExpDate(0) = bytLicenseKey(POS_EXP_DATE_M1)
   bytExpDate(1) = bytLicenseKey(POS_EXP_DATE_M2)
   bytExpDate(2) = 0 'just a place holder
   bytExpDate(3) = bytLicenseKey(POS_EXP_DATE_D1)
   bytExpDate(4) = bytLicenseKey(POS_EXP_DATE_D2)
   bytExpDate(5) = 0 'just a place holder
   bytExpDate(6) = bytLicenseKey(POS_EXP_DATE_Y1)
   bytExpDate(7) = bytLicenseKey(POS_EXP_DATE_Y2)
   bytExpDate(8) = bytLicenseKey(POS_EXP_DATE_Y3)
   bytExpDate(9) = bytLicenseKey(POS_EXP_DATE_Y4)
   
   GetExpDateFromArray = bytExpDate

End Function


Private Function GetNumClientsFromArray() As Byte()

   Dim bytNumClients() As Byte
   ReDim bytNumClients(NUM_LIMIT_FIELDS - 1)

   bytNumClients(0) = bytLicenseKey(POS_CLIENT_LIM_1)
   bytNumClients(1) = bytLicenseKey(POS_CLIENT_LIM_2)
   bytNumClients(2) = bytLicenseKey(POS_CLIENT_LIM_3)
   bytNumClients(3) = bytLicenseKey(POS_CLIENT_LIM_4)
   
   GetNumClientsFromArray = bytNumClients

End Function


Private Function GetNumVehiclesFromArray() As Byte()

   Dim bytNumVehicles() As Byte
   ReDim bytNumVehicles(NUM_LIMIT_FIELDS - 1)

   bytNumVehicles(0) = bytLicenseKey(POS_VEHICLE_LIM_1)
   bytNumVehicles(1) = bytLicenseKey(POS_VEHICLE_LIM_2)
   bytNumVehicles(2) = bytLicenseKey(POS_VEHICLE_LIM_3)
   bytNumVehicles(3) = bytLicenseKey(POS_VEHICLE_LIM_4)
   
   GetNumVehiclesFromArray = bytNumVehicles

End Function
                    

Private Function GetCRCDataFromArray() As Byte()

   Dim bytCRCData() As Byte
   ReDim bytCRCData(NUM_CRC_FIELDS - 1)

   bytCRCData(0) = bytLicenseKey(POS_CRC_1)
   bytCRCData(1) = bytLicenseKey(POS_CRC_2)
   bytCRCData(2) = bytLicenseKey(POS_CRC_3)
   bytCRCData(3) = bytLicenseKey(POS_CRC_4)
   
   GetCRCDataFromArray = bytCRCData

End Function


Private Function Decode(bytAllData() As Byte) As Byte()

   Dim lngLen As Long
   lngLen = ArrayLen(bytAllData)
   
   Dim lngIndex
   For lngIndex = 0 To lngLen - 1
      
      Dim bytData As Byte
      bytData = bytAllData(lngIndex)
      
      'shift a few
      bytData = LeftShiftRotateByte(bytData, 2)
      
      'invert
      bytData = Not bytData
      
      bytAllData(lngIndex) = bytData
      
   Next lngIndex
   
   Decode = bytAllData

End Function



Private Function GetDateFromByteArray(bytDateData() As Byte) As Date

   bytDateData(2) = AscB(DATE_SEPARATOR)
   bytDateData(5) = AscB(DATE_SEPARATOR)

   GetDateFromByteArray = CDate(ConvertByteArrayToString2(bytDateData))
   
End Function


Private Function GetIntegerFromByteArray(bytIntData() As Byte) As Integer
   
   GetIntegerFromByteArray = CInt(ConvertByteArrayToString2(bytIntData))
                             
End Function


Private Function LeftShiftRotateByte(bytOrig As Byte, intNumPlaces As Integer) As Byte

   LeftShiftRotateByte = bytOrig

   If intNumPlaces > 0 Then
      
      Dim bytFillMask As Byte
      bytFillMask = 1 ' binary 00000001, unsigned
      
      Dim bytRotateMask As Byte
      bytRotateMask = 128 ' binary 10000000, unsigned
      
      Dim bytOverflowMask As Byte
      bytOverflowMask = 127 ' binary 01111111, unsigned
      
      Dim bytTemp As Byte
      
      Dim intIndex As Integer
      For intIndex = 1 To intNumPlaces
         
         bytTemp = LeftShiftRotateByte 'keep orig value to see if we need to shift
         
         'we would normally just do a "* 2" but that could lead to overflow if MSB is set.
         'Instead, we will mask out the MSB, and *then* do a "* 2"
         LeftShiftRotateByte = (LeftShiftRotateByte And bytOverflowMask)
         LeftShiftRotateByte = LeftShiftRotateByte * 2
         
         If (bytTemp And bytRotateMask) = bytRotateMask Then 'fill on the right
            LeftShiftRotateByte = (LeftShiftRotateByte Or bytFillMask)
         End If
         
      Next intIndex
      
   End If

End Function


Private Function ValidateCRC(bytLicenseKey() As Byte, bytEmbeddedCRC() As Byte) As Boolean

   ValidateCRC = False

   Dim bytCRCData() As Byte
   bytCRCData = GenerateCRC(bytLicenseKey)
   
   If ArrayLen(bytCRCData) = 4 Then
      If bytCRCData(0) = bytEmbeddedCRC(0) And _
         bytCRCData(1) = bytEmbeddedCRC(1) And _
         bytCRCData(2) = bytEmbeddedCRC(2) And _
         bytCRCData(3) = bytEmbeddedCRC(3) Then
                  
         ValidateCRC = True
         
      End If
   End If

End Function



Private Function GenerateCRC(bytLicenseKey() As Byte) As Byte()

   Dim bytCRCData() As Byte
   
   If clsCRC.GenerateCRC(bytLicenseKey, bytCRCData) = False Then
#If LIC_MANAGER_PROJECT Then
      MsgBox "LicenseKey::GenerateCRC(), Problem getting CRC Value"
#Else
      gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
                     "LicenseExtractor::GenerateCRC()" & vbCrLf & _
                     "   Problem generating a CRC Value"
#End If
   End If
   
   GenerateCRC = bytCRCData
   
End Function

