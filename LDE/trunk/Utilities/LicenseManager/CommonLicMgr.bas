Attribute VB_Name = "CommonLicMgr"
Option Explicit

Public Const POS_GEN_DATE_M1 As Integer = 12
Public Const POS_GEN_DATE_M2 As Integer = 55
Public Const POS_GEN_DATE_D1 As Integer = 102
Public Const POS_GEN_DATE_D2 As Integer = 119
Public Const POS_GEN_DATE_Y1 As Integer = 34
Public Const POS_GEN_DATE_Y2 As Integer = 2
Public Const POS_GEN_DATE_Y3 As Integer = 36
Public Const POS_GEN_DATE_Y4 As Integer = 71


Public Const POS_EXP_DATE_M1 As Integer = 15
Public Const POS_EXP_DATE_M2 As Integer = 3
Public Const POS_EXP_DATE_D1 As Integer = 121
Public Const POS_EXP_DATE_D2 As Integer = 88
Public Const POS_EXP_DATE_Y1 As Integer = 91
Public Const POS_EXP_DATE_Y2 As Integer = 100
Public Const POS_EXP_DATE_Y3 As Integer = 17
Public Const POS_EXP_DATE_Y4 As Integer = 6

Public Const POS_CLIENT_LIM_1 As Integer = 117
Public Const POS_CLIENT_LIM_2 As Integer = 69
Public Const POS_CLIENT_LIM_3 As Integer = 41
Public Const POS_CLIENT_LIM_4 As Integer = 10

Public Const POS_VEHICLE_LIM_1 As Integer = 50
Public Const POS_VEHICLE_LIM_2 As Integer = 99
Public Const POS_VEHICLE_LIM_3 As Integer = 13
Public Const POS_VEHICLE_LIM_4 As Integer = 25

Public Const POS_CRC_1 As Integer = 103
Public Const POS_CRC_2 As Integer = 78
Public Const POS_CRC_3 As Integer = 37
Public Const POS_CRC_4 As Integer = 9

Public Const NUM_DATE_FIELDS As Long = 10
Public Const NUM_LIMIT_FIELDS As Long = 4
Public Const NUM_CRC_FIELDS As Long = 4

Public Const KNOWN_INITIAL_VALUE As String = "x"

Public Const UNLIMITED_CLIENTS As Integer = 9999  'string needs to be four digits
Public Const UNLIMITED_VEHICLES As Integer = 9999 '
Public Const UNLIMITED_EXP_DATE As String = "11" & DATE_SEPARATOR & "18" & DATE_SEPARATOR & "5000"


' License string length
Public Const LIC_LENGTH As Integer = 127



Public Sub PlaceCRCDataInArray(bytLicenseArray() As Byte, bytCRCData() As Byte)

   bytLicenseArray(POS_CRC_1) = bytCRCData(0)
   bytLicenseArray(POS_CRC_2) = bytCRCData(1)
   bytLicenseArray(POS_CRC_3) = bytCRCData(2)
   bytLicenseArray(POS_CRC_4) = bytCRCData(3)
   
End Sub


Public Sub PutPlaceholdersInCRCSlots(bytLicenseArray() As Byte)

   bytLicenseArray(POS_CRC_1) = AscB(KNOWN_INITIAL_VALUE)
   bytLicenseArray(POS_CRC_2) = AscB(KNOWN_INITIAL_VALUE)
   bytLicenseArray(POS_CRC_3) = AscB(KNOWN_INITIAL_VALUE)
   bytLicenseArray(POS_CRC_4) = AscB(KNOWN_INITIAL_VALUE)

End Sub

