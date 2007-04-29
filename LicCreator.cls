VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LicCreator"
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




'assumptions: data is convertible to strings, and is within reasonable limits
Public Function Generate(datGenDate As Date, _
                         datExpDate As Date, _
                         intNumClients As Integer, _
                         intNumVehicles As Integer) As String
                         
   On Error GoTo ERROR
   
   Erase bytLicenseKey
   ReDim bytLicenseKey(LIC_LENGTH - 1)
   
   'convert all data into strings, the byte arrays
   Dim strNumClients As String
   Dim strNumVehicles As String
   
   strNumClients = CStr(intNumClients)
   While Len(strNumClients) < NUM_LIMIT_FIELDS
      strNumClients = "0" & strNumClients
   Wend
   
   strNumVehicles = CStr(intNumVehicles)
   While Len(strNumVehicles) < NUM_LIMIT_FIELDS
      strNumVehicles = "0" & strNumVehicles
   Wend
   
   Dim bytGenDate() As Byte
   Dim bytExpDate() As Byte
   Dim bytNumClients() As Byte
   Dim bytNumVehicles() As Byte
   
   bytGenDate = ConvertStringToByteArray2(Format$(CStr(datGenDate), STD_DATE_FORMAT))
   bytExpDate = ConvertStringToByteArray2(Format$(CStr(datExpDate), STD_DATE_FORMAT))
   bytNumClients = ConvertStringToByteArray2(strNumClients)
   bytNumVehicles = ConvertStringToByteArray2(strNumVehicles)
   
   Debug.Assert ArrayLen(bytGenDate) = NUM_DATE_FIELDS
   Debug.Assert ArrayLen(bytExpDate) = NUM_DATE_FIELDS
   Debug.Assert ArrayLen(bytNumClients) = NUM_LIMIT_FIELDS
   Debug.Assert ArrayLen(bytNumVehicles) = NUM_LIMIT_FIELDS

   'do coding operations
   bytGenDate = Code(bytGenDate)
   bytExpDate = Code(bytExpDate)
   bytNumClients = Code(bytNumClients)
   bytNumVehicles = Code(bytNumVehicles)
   
   'fill array with random numbers
   RandomizeArray
   
   'place coded values in the array
   PlaceGenDateInArray bytGenDate
   PlaceExpDateInArray bytExpDate
   PlaceNumClientsInArray bytNumClients
   PlaceNumVehiclesInArray bytNumVehicles
   
   'generate crc from array and place crc values in array
   PutPlaceholdersInCRCSlots bytLicenseKey
   Dim bytCRC() As Byte
   bytCRC = GenerateCRC(bytLicenseKey)
   If ArrayLen(bytCRC) > 0 Then
      PlaceCRCDataInArray bytLicenseKey, bytCRC
   Else
      Generate = vbNullString
      GoTo EXIT_FUNC 'already had a message box by here
   End If
   
   Generate = ConvertByteArrayToHexValueString(bytLicenseKey)
   
EXIT_FUNC:
   Exit Function

ERROR:
   MsgBox "Error in LicenseKey::Generate()" & vbCrLf & _
          "   num = """ & Err.Number & """, desc = """ & Err.Description & """"
   Generate = vbNullString
   Resume EXIT_FUNC
                         
End Function


Private Function Code(bytAllData() As Byte) As Byte()

   Dim lngLen As Long
   lngLen = ArrayLen(bytAllData)
   
   Dim lngIndex
   For lngIndex = 0 To lngLen - 1
      
      Dim bytData As Byte
      bytData = bytAllData(lngIndex)
      
      'invert
      bytData = Not bytData
      
      'shift a few
      bytData = RightShiftRotateByte(bytData, 2)
      
      bytAllData(lngIndex) = bytData
      
   Next lngIndex
   
   Code = bytAllData

End Function


Private Sub RandomizeArray()

   Randomize Now()
      
   Dim lngIndex As Long
   For lngIndex = 0 To LIC_LENGTH - 1
   
      Dim bytData As Byte
      bytData = CByte(Int(Rnd() * 256))
      
      bytLicenseKey(lngIndex) = bytData
      
   Next lngIndex
   
End Sub


Private Sub PlaceGenDateInArray(bytGenDate() As Byte)

   bytLicenseKey(POS_GEN_DATE_M1) = bytGenDate(0)
   bytLicenseKey(POS_GEN_DATE_M2) = bytGenDate(1)
   bytLicenseKey(POS_GEN_DATE_D1) = bytGenDate(3)
   bytLicenseKey(POS_GEN_DATE_D2) = bytGenDate(4)
   bytLicenseKey(POS_GEN_DATE_Y1) = bytGenDate(6)
   bytLicenseKey(POS_GEN_DATE_Y2) = bytGenDate(7)
   bytLicenseKey(POS_GEN_DATE_Y3) = bytGenDate(8)
   bytLicenseKey(POS_GEN_DATE_Y4) = bytGenDate(9)

End Sub


Private Sub PlaceExpDateInArray(bytExpDate() As Byte)

   bytLicenseKey(POS_EXP_DATE_M1) = bytExpDate(0)
   bytLicenseKey(POS_EXP_DATE_M2) = bytExpDate(1)
   bytLicenseKey(POS_EXP_DATE_D1) = bytExpDate(3)
   bytLicenseKey(POS_EXP_DATE_D2) = bytExpDate(4)
   bytLicenseKey(POS_EXP_DATE_Y1) = bytExpDate(6)
   bytLicenseKey(POS_EXP_DATE_Y2) = bytExpDate(7)
   bytLicenseKey(POS_EXP_DATE_Y3) = bytExpDate(8)
   bytLicenseKey(POS_EXP_DATE_Y4) = bytExpDate(9)
   
End Sub


Private Sub PlaceNumClientsInArray(bytNumClients() As Byte)

   bytLicenseKey(POS_CLIENT_LIM_1) = bytNumClients(0)
   bytLicenseKey(POS_CLIENT_LIM_2) = bytNumClients(1)
   bytLicenseKey(POS_CLIENT_LIM_3) = bytNumClients(2)
   bytLicenseKey(POS_CLIENT_LIM_4) = bytNumClients(3)
   
End Sub


Private Sub PlaceNumVehiclesInArray(bytNumVehicles() As Byte)

   bytLicenseKey(POS_VEHICLE_LIM_1) = bytNumVehicles(0)
   bytLicenseKey(POS_VEHICLE_LIM_2) = bytNumVehicles(1)
   bytLicenseKey(POS_VEHICLE_LIM_3) = bytNumVehicles(2)
   bytLicenseKey(POS_VEHICLE_LIM_4) = bytNumVehicles(3)
   
End Sub


Private Function RightShiftRotateByte(bytOrig As Byte, intNumPlaces As Integer) As Byte

   RightShiftRotateByte = bytOrig
   
   If intNumPlaces > 0 Then
      
      Dim bytFillMask As Byte
      bytFillMask = 128 ' binary 10000000, unsigned
      
      Dim bytRotateMask As Byte
      bytRotateMask = 1 ' binary 00000001, unsigned
      
      Dim bytTemp As Byte
      
      Dim intIndex As Integer
      For intIndex = 1 To intNumPlaces
         
         bytTemp = RightShiftRotateByte 'keep orig value to see if we need to shift
         RightShiftRotateByte = bytTemp \ 2
         
         If (bytTemp And bytRotateMask) = bytRotateMask Then  ' fill on the left
            RightShiftRotateByte = (RightShiftRotateByte Or bytFillMask)
         End If
         
      Next intIndex
      
   End If

End Function


Private Function GenerateCRC(bytLicenseKey() As Byte) As Byte()

   Dim bytCRCData() As Byte
   
   If clsCRC.GenerateCRC(bytLicenseKey, bytCRCData) = False Then
      MsgBox "LicenseKey::GenerateCRC(), Problem getting CRC Value"
   End If
   
   GenerateCRC = bytCRCData
   
End Function

