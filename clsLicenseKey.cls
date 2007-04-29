VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLicenseKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The license string is 254 bytes long. Each two bytes is a string of a single hex value, 0-255
'actually "00" to "ff", which encodes a single character. So while we say that a license string
'is 127 long ( or whatever ), in reality it is 2 * 127 long. This is one of the levels of encoding.
'
'Encoding process:
' 1) Each piece of data is converted to a string
' 2) Each string is converted to it's hexstring equivalent ( doubling it's size )
' 3) Each byte within the hexstring equivalent is inverted and then right-shifted ( with fill ), 2 places
' 4) Each piece of data has it's individual encoded hexstring equivalents placed in different
'    places within the license string.
' 5) A CRC of the resultant string is calculated, encoded, and placed back in the string, again, broken
'    up and placed in different places within the array.
'
'Decoding is exactly the reverse.

'Positions for data within the license string (0-based array of bytes)
'Numbering is from left to right ( 1 -> 4 )

Private Const POS_GEN_DATE_M1 As Integer = 12
Private Const POS_GEN_DATE_M2 As Integer = 55
Private Const POS_GEN_DATE_D1 As Integer = 102
Private Const POS_GEN_DATE_D2 As Integer = 119
Private Const POS_GEN_DATE_Y1 As Integer = 34
Private Const POS_GEN_DATE_Y2 As Integer = 2
Private Const POS_GEN_DATE_Y3 As Integer = 36
Private Const POS_GEN_DATE_Y4 As Integer = 71

Private Const NUM_DATE_FIELDS As Long = 10

Private Const POS_EXP_DATE_M1 As Integer = 15
Private Const POS_EXP_DATE_M2 As Integer = 3
Private Const POS_EXP_DATE_D1 As Integer = 121
Private Const POS_EXP_DATE_D2 As Integer = 88
Private Const POS_EXP_DATE_Y1 As Integer = 91
Private Const POS_EXP_DATE_Y2 As Integer = 100
Private Const POS_EXP_DATE_Y3 As Integer = 17
Private Const POS_EXP_DATE_Y4 As Integer = 6

Private Const POS_CLIENT_LIM_1 As Integer = 117
Private Const POS_CLIENT_LIM_2 As Integer = 69
Private Const POS_CLIENT_LIM_3 As Integer = 41
Private Const POS_CLIENT_LIM_4 As Integer = 10

Private Const POS_VEHICLE_LIM_1 As Integer = 50
Private Const POS_VEHICLE_LIM_2 As Integer = 99
Private Const POS_VEHICLE_LIM_3 As Integer = 13
Private Const POS_VEHICLE_LIM_4 As Integer = 25

Private Const NUM_LIMIT_FIELDS As Long = 4

Private Const POS_CRC_1 As Integer = 103
Private Const POS_CRC_2 As Integer = 78
Private Const POS_CRC_3 As Integer = 37
Private Const POS_CRC_4 As Integer = 9

Private Const KNOWN_INITIAL_VALUE As String = "x"
Private Const KNOWN_CRC_VALUE As String = "y"


Private bytLicenseKey() As Byte
Private clsCRC As clsCRC16


Private Sub Class_Initialize()

   Set clsCRC = New clsCRC16

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
   
   bytGenDate = ConvertStringToByteArray(Format$(CStr(datGenDate), STD_DATE_FORMAT))
   bytExpDate = ConvertStringToByteArray(Format$(CStr(datExpDate), STD_DATE_FORMAT))
   bytNumClients = ConvertStringToByteArray(strNumClients)
   bytNumVehicles = ConvertStringToByteArray(strNumVehicles)
   
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
   PutPlaceholdersInCRCSlots
   Dim bytCRC() As Byte
   bytCRC = GenerateCRC(bytLicenseKey)
   If ArrayLen(bytCRC) > 0 Then
      PlaceCRCDataInArray bytCRC
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
   PutPlaceholdersInCRCSlots

   If ValidateCRC(bytLicenseKey, bytEmbeddedCRC) Then 'all's well
      'put CRC back in license string
      PlaceCRCDataInArray bytEmbeddedCRC
   Else
      MsgBox "Invalid CRC in license string, license invalid"
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
   MsgBox "Error in LicenseKey::Extract()" & vbCrLf & _
          "   num = """ & Err.Number & """, desc = """ & Err.Description & """"
   Resume EXIT_FUNC

End Function


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


Private Sub PlaceCRCDataInArray(bytCRCData() As Byte)

   bytLicenseKey(POS_CRC_1) = bytCRCData(0)
   bytLicenseKey(POS_CRC_2) = bytCRCData(1)
   bytLicenseKey(POS_CRC_3) = bytCRCData(2)
   bytLicenseKey(POS_CRC_4) = bytCRCData(3)
   
End Sub


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


Private Sub PutPlaceholdersInCRCSlots()

   bytLicenseKey(POS_CRC_1) = AscB(KNOWN_INITIAL_VALUE)
   bytLicenseKey(POS_CRC_2) = AscB(KNOWN_INITIAL_VALUE)
   bytLicenseKey(POS_CRC_3) = AscB(KNOWN_INITIAL_VALUE)
   bytLicenseKey(POS_CRC_4) = AscB(KNOWN_INITIAL_VALUE)

End Sub

                    

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

   GetDateFromByteArray = CDate(ConvertByteArrayToString(bytDateData))
   
End Function


Private Function GetIntegerFromByteArray(bytIntData() As Byte) As Integer
   
   GetIntegerFromByteArray = CInt(ConvertByteArrayToString(bytIntData))
                             
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


Private Function GenerateCRC(bytLicenseKey() As Byte) As Byte()

   Dim bytCRCData() As Byte
   
   If clsCRC.GenerateCRC(bytLicenseKey, bytCRCData) = False Then
      MsgBox "LicenseKey::GenerateCRC(), Problem getting CRC Value"
   End If
   
   GenerateCRC = bytCRCData
   
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


















