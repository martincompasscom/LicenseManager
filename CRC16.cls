VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRC16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' basCRC16: Calculates CRC-16 checksum for a given message string
' Version 1. Published 6 May 2001.
'************************COPYRIGHT NOTICE*************************
' Copyright (C) 2001 DI Management Services Pty Ltd,
' Sydney Australia <www.di-mgt.com.au>. All rights reserved.
' This code was originally written in Visual Basic by David Ireland.
' You are free to use this code in your applications without liability
' or compensation, but the courtesy of both notification of use and
' inclusion of due credit are requested. You must keep this copyright
' notice intact.
' It is PROHIBITED to distribute or reproduce this code for profit
' or otherwise, on any web site, ftp server or BBS, or by any
' other means, including CD-ROM or other physical media, without the
' EXPRESS WRITTEN PERMISSION of the author.
' Use at your own risk.
' David Ireland and DI Management Services Pty Limited
' offer no warranty of its fitness for any purpose whatsoever,
' and accept no liability whatsoever for any loss or damage
' incurred by its use.
' If you use it, or found it useful, or can suggest an improvement
' please let us know at <code@di-mgt.com.au>.
'*****************************************************************

Private aCRC16Table(255) As Integer


Private Sub Class_Initialize()

   CRC16Setup
   
End Sub


Private Sub Class_Terminate()

   '

End Sub

'
'Public Function TestCRC16()
'
'' Test suite answers:
''CRC16(123456789) = BB3D
''CRC16(hello world)=39C1
''CRC16(Hello world)=F96A
''CRC16(a) = E8C1
''CRC16() = D801
'
'   Dim sMessage As String
'   Dim iCRC As Integer
'
'   sMessage = "123456789"
'   iCRC = ValidateCRC16(sMessage)
'   Debug.Print "crc16(" & sMessage & ")=" & Hex(iCRC)
'
'   sMessage = "hello world"
'   iCRC = ValidateCRC16(sMessage)
'   Debug.Print "crc16(" & sMessage & ")=" & Hex(iCRC)
'
'   sMessage = "Hello world"
'   iCRC = ValidateCRC16(sMessage)
'   Debug.Print "crc16(" & sMessage & ")=" & Hex(iCRC)
'
'   sMessage = "a"
'   iCRC = ValidateCRC16(sMessage)
'   Debug.Print "crc16(" & sMessage & ")=" & Hex(iCRC)
'
'   sMessage = " "
'   iCRC = ValidateCRC16(sMessage)
'   Debug.Print "crc16(" & sMessage & ")=" & Hex(iCRC)
'
'End Function


Public Function ValidateCRC16(sMessage As String) As Boolean

   ValidateCRC16 = True
    
'   Dim iCRC As Integer
'   Dim i As Integer
'   Dim bytT As Byte
'   Dim bytC As Byte
'   Dim ia As Integer
'
'   'extract candidate CRC value
'
'
'   ' Set iCRC = 0
'   ' For each byte in message
'   '   calculate iCRC = (iCRC >> 8) ^ Table[(iCRC & 0xFF) ^ byte]
'   ' Return iCRC
'
'   iCRC = 0
'
'   For i = 1 To Len(sMessage)
'
'      bytC = Asc(Mid(sMessage, i, 1))
'      bytT = (iCRC And &HFF) Xor bytC
'      ia = uiShiftRightBy8(iCRC)
'      iCRC = ia Xor aCRC16Table(bytT)
'
'   Next
'
'   'compare calculated CRC to candidate

End Function


'extract candidate CRC value
' Set iCRC = 0
' For each byte in message
'   calculate iCRC = (iCRC >> 8) ^ Table[(iCRC & 0xFF) ^ byte]
' Return iCRC
Public Function GenerateCRC(bytData() As Byte, bytCRCData() As Byte) As Boolean

   On Error GoTo ERROR

   GenerateCRC = False
    
   Dim bytT As Byte
   Dim bytC As Byte
   Dim intShifted As Integer

   Dim intCRC As Integer
   intCRC = 0
   
   Dim intArrayLen As Integer
   intArrayLen = CInt(ArrayLen(bytData))

   Dim intIndex As Integer
   For intIndex = 0 To intArrayLen - 1

      bytC = bytData(intIndex)
      bytT = (intCRC And &HFF) Xor bytC
      intShifted = uiShiftRightBy8(intCRC)
      intCRC = intShifted Xor aCRC16Table(bytT)

   Next intIndex

   bytCRCData = GetCRCByteValues(intCRC)

   GenerateCRC = True
   
EXIT_FUNC:
   Exit Function
   
ERROR:
#If LIC_MANAGER_PROJECT Then
   MsgBox "Error in CRC16::GenerateCRC()" And vbCrLf & _
          "   num = """ & Err.Number & """, desc = """ & Err.Description & """"
#Else
   gclsLogger.Log gclsLogger.LOG_WHAT_ERROR, _
              "Error in CRC16::GenerateCRC()" & vbCrLf & _
              "   num = """ & Err.Number & """, desc = """ & Err.Description & """"
#End If
   Resume EXIT_FUNC

End Function


Private Function uiShiftRightBy8(intParam As Integer) As Integer
    ' Shift 16-bit integer value to right by 8 bits
    ' Avoiding problem with sign bit
    
   Dim intNew As Integer
   intNew = (intParam And &H7FFF) \ 256
   
   If (intParam And &H8000) <> 0 Then
      intNew = intNew Or &H80
   End If
   
   uiShiftRightBy8 = intNew
   
End Function


Private Function GetCRCByteValues(intCRC As Integer) As Byte()

   Dim bytData() As Byte
   ReDim bytData(NUM_CRC_FIELDS - 1)
   
   Dim intMask As Integer
   intMask = 15 ' 0000000000001111

   bytData(3) = CByte(intCRC And intMask)

   ' right shift intCRC over by 4 bits
   intCRC = intCRC \ 16
   bytData(2) = CByte(intCRC And intMask)

   ' right shift intCRC over by 4 bits
   intCRC = intCRC \ 16
   bytData(1) = CByte(intCRC And intMask)

   ' right shift intCRC over by 4 bits
   intCRC = intCRC \ 16
   bytData(0) = CByte(intCRC And intMask)
   
   GetCRCByteValues = bytData

End Function


Private Function CRC16Setup()

   Dim vntA As Variant
   Dim intIndex As Integer

   ' Use variant array kludge to set up table
   vntA = Array( _
      &H0, &HC0C1, &HC181, &H140, &HC301, &H3C0, &H280, &HC241, _
      &HC601, &H6C0, &H780, &HC741, &H500, &HC5C1, &HC481, &H440, _
      &HCC01, &HCC0, &HD80, &HCD41, &HF00, &HCFC1, &HCE81, &HE40, _
      &HA00, &HCAC1, &HCB81, &HB40, &HC901, &H9C0, &H880, &HC841, _
      &HD801, &H18C0, &H1980, &HD941, &H1B00, &HDBC1, &HDA81, &H1A40, _
      &H1E00, &HDEC1, &HDF81, &H1F40, &HDD01, &H1DC0, &H1C80, &HDC41, _
      &H1400, &HD4C1, &HD581, &H1540, &HD701, &H17C0, &H1680, &HD641, _
      &HD201, &H12C0, &H1380, &HD341, &H1100, &HD1C1, &HD081, &H1040)
       
   For intIndex = 0 To 63
      aCRC16Table(intIndex) = vntA(intIndex - 0)
   Next intIndex
   
   vntA = Array( _
      &HF001, &H30C0, &H3180, &HF141, &H3300, &HF3C1, &HF281, &H3240, _
      &H3600, &HF6C1, &HF781, &H3740, &HF501, &H35C0, &H3480, &HF441, _
      &H3C00, &HFCC1, &HFD81, &H3D40, &HFF01, &H3FC0, &H3E80, &HFE41, _
      &HFA01, &H3AC0, &H3B80, &HFB41, &H3900, &HF9C1, &HF881, &H3840, _
      &H2800, &HE8C1, &HE981, &H2940, &HEB01, &H2BC0, &H2A80, &HEA41, _
      &HEE01, &H2EC0, &H2F80, &HEF41, &H2D00, &HEDC1, &HEC81, &H2C40, _
      &HE401, &H24C0, &H2580, &HE541, &H2700, &HE7C1, &HE681, &H2640, _
      &H2200, &HE2C1, &HE381, &H2340, &HE101, &H21C0, &H2080, &HE041)

   For intIndex = 64 To 127
      aCRC16Table(intIndex) = vntA(intIndex - 64)
   Next intIndex
   
   vntA = Array( _
      &HA001, &H60C0, &H6180, &HA141, &H6300, &HA3C1, &HA281, &H6240, _
      &H6600, &HA6C1, &HA781, &H6740, &HA501, &H65C0, &H6480, &HA441, _
      &H6C00, &HACC1, &HAD81, &H6D40, &HAF01, &H6FC0, &H6E80, &HAE41, _
      &HAA01, &H6AC0, &H6B80, &HAB41, &H6900, &HA9C1, &HA881, &H6840, _
      &H7800, &HB8C1, &HB981, &H7940, &HBB01, &H7BC0, &H7A80, &HBA41, _
      &HBE01, &H7EC0, &H7F80, &HBF41, &H7D00, &HBDC1, &HBC81, &H7C40, _
      &HB401, &H74C0, &H7580, &HB541, &H7700, &HB7C1, &HB681, &H7640, _
      &H7200, &HB2C1, &HB381, &H7340, &HB101, &H71C0, &H7080, &HB041)

   For intIndex = 128 To 191
      aCRC16Table(intIndex) = vntA(intIndex - 128)
   Next intIndex
   
   vntA = Array( _
      &H5000, &H90C1, &H9181, &H5140, &H9301, &H53C0, &H5280, &H9241, _
      &H9601, &H56C0, &H5780, &H9741, &H5500, &H95C1, &H9481, &H5440, _
      &H9C01, &H5CC0, &H5D80, &H9D41, &H5F00, &H9FC1, &H9E81, &H5E40, _
      &H5A00, &H9AC1, &H9B81, &H5B40, &H9901, &H59C0, &H5880, &H9841, _
      &H8801, &H48C0, &H4980, &H8941, &H4B00, &H8BC1, &H8A81, &H4A40, _
      &H4E00, &H8EC1, &H8F81, &H4F40, &H8D01, &H4DC0, &H4C80, &H8C41, _
      &H4400, &H84C1, &H8581, &H4540, &H8701, &H47C0, &H4680, &H8641, _
      &H8201, &H42C0, &H4380, &H8341, &H4100, &H81C1, &H8081, &H4040)

   For intIndex = 192 To 255
      aCRC16Table(intIndex) = vntA(intIndex - 192)
   Next intIndex
    
End Function







