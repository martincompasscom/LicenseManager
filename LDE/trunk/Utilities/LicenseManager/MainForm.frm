VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "CompassCom License Manager"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnClearFields 
      Caption         =   "Clear Fields"
      Height          =   225
      Left            =   4275
      TabIndex        =   15
      Top             =   1935
      Width           =   1380
   End
   Begin VB.CheckBox chkExpDateUnlimited 
      Caption         =   "Unlimited"
      Height          =   195
      Left            =   4680
      TabIndex        =   11
      Top             =   1485
      Width           =   1050
   End
   Begin VB.CheckBox chkVehicleUnlimited 
      Caption         =   "Unlimited"
      Height          =   195
      Left            =   4680
      TabIndex        =   10
      Top             =   990
      Width           =   1050
   End
   Begin VB.CheckBox chkClientUnlimited 
      Caption         =   "Unlimited"
      Height          =   195
      Left            =   4680
      TabIndex        =   9
      Top             =   450
      Width           =   1050
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "Ok"
      Height          =   225
      Left            =   4860
      TabIndex        =   8
      Top             =   3555
      Width           =   855
   End
   Begin VB.TextBox txtLicenseKey 
      Height          =   285
      Left            =   210
      TabIndex        =   7
      Top             =   3030
      Width           =   5490
   End
   Begin VB.CommandButton btnDecode 
      Caption         =   "Decode"
      Height          =   225
      Left            =   3600
      TabIndex        =   6
      Top             =   2610
      Width           =   1380
   End
   Begin VB.CommandButton btnGenerate 
      Caption         =   "Generate"
      Height          =   225
      Left            =   900
      TabIndex        =   5
      Top             =   2610
      Width           =   1380
   End
   Begin VB.TextBox txtExpirationDate 
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtVehicleLimit 
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Top             =   945
      Width           =   975
   End
   Begin VB.TextBox txtClientLimit 
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   420
      Width           =   975
   End
   Begin VB.TextBox txtGeneratedDate 
      Height          =   285
      Left            =   210
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Expiration Date:"
      Height          =   195
      Left            =   2385
      TabIndex        =   14
      Top             =   1485
      Width           =   1140
   End
   Begin VB.Label Label3 
      Caption         =   "Vehicle limit:"
      Height          =   195
      Left            =   2430
      TabIndex        =   13
      Top             =   990
      Width           =   1050
   End
   Begin VB.Label Label2 
      Caption         =   "Client limit:"
      Height          =   195
      Left            =   2430
      TabIndex        =   12
      Top             =   450
      Width           =   1050
   End
   Begin VB.Line Line1 
      X1              =   210
      X2              =   5670
      Y1              =   2385
      Y2              =   2385
   End
   Begin VB.Label Label1 
      Caption         =   "Date Generated:"
      Height          =   225
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   1800
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsLicKeyCreator As LicCreator
Private clsLicKeyExtractor As LicExtractor

Public lstMessages As ListBox


Private Sub Form_Load()

   Set clsLicKeyCreator = New LicCreator
   Set clsLicKeyExtractor = New LicExtractor
   
   txtClientLimit.Enabled = False
   txtVehicleLimit.Enabled = False
   txtExpirationDate.Enabled = False
   
   chkClientUnlimited.Value = 1
   chkVehicleUnlimited.Value = 1
   chkExpDateUnlimited.Value = 1
   
   glngLoggingWhat = 0
   glngLoggingWhere = 0
   gstrLoggingFile = vbNullString
   glngDaysLogFileHistory = 3
   Set gclsLogger = New Logger
   
   Set gclsXMLParser = New XMLParser
   
End Sub


Private Sub Form_Unload(intCancel As Integer)

   Set clsLicKeyCreator = Nothing
   Set clsLicKeyExtractor = Nothing
   
   Set gclsLogger = Nothing
   Set gclsXMLParser = Nothing
   
End Sub


Private Sub btnDecode_Click()

   If Len(txtLicenseKey.Text()) = LIC_LENGTH * 2 Then
   
      Dim intNumClients As Integer
      Dim intNumVehicles As Integer
      Dim datGenDate As Date
      Dim datExpDate As Date
      
      Dim blnValidLicenseKey As Boolean
      blnValidLicenseKey = clsLicKeyExtractor.Extract(txtLicenseKey.Text(), _
                                                      datGenDate, _
                                                      datExpDate, _
                                                      intNumClients, _
                                                      intNumVehicles)
      
      If blnValidLicenseKey Then
      
         txtGeneratedDate.Text = Format$(datGenDate, STD_DATE_FORMAT)
         
         If Format$(datExpDate, STD_DATE_FORMAT) = UNLIMITED_EXP_DATE Then
            chkExpDateUnlimited.Value = 1
            txtExpirationDate.Enabled = False
         Else
            chkExpDateUnlimited.Value = 0
            txtExpirationDate.Enabled = True
            txtExpirationDate.Text = Format$(datExpDate, STD_DATE_FORMAT)
         End If
         
         If intNumClients = UNLIMITED_CLIENTS Then
            chkClientUnlimited.Value = 1
            txtClientLimit.Enabled = False
         Else
            chkClientUnlimited.Value = 0
            txtClientLimit.Enabled = True
            txtClientLimit.Text = CStr(intNumClients)
         End If
         
         If intNumVehicles = UNLIMITED_VEHICLES Then
            chkVehicleUnlimited.Value = 1
            txtVehicleLimit.Enabled = False
         Else
            chkVehicleUnlimited.Value = 0
            txtVehicleLimit.Enabled = True
            txtVehicleLimit.Text = CStr(intNumVehicles)
         End If
         
      End If
      
   Else
      MsgBox "Invalid license key, invalid length = " & Len(txtLicenseKey.Text()) & vbCrLf & _
             "   length should be " & CStr(LIC_LENGTH * 2)
   End If

End Sub


Private Sub btnGenerate_Click()

   Dim strAttemptedInput As String

   'validate all input values
   Dim intNumClients As Integer
   Dim intNumVehicles As Integer
   Dim datExpiration As Date
   
   If chkClientUnlimited.Value() = 0 Then
   
      On Error Resume Next
      intNumClients = CInt(txtClientLimit.Text())
      
      If Err.Number = 0 Then
         If intNumClients < 1 Then
            MsgBox "Invalid clients limit = """ & txtClientLimit.Text() & """"
            GoTo EXIT_SUB
         End If
      Else
         MsgBox "Invalid clients limit = """ & txtClientLimit.Text() & """"
         GoTo EXIT_SUB
      End If
      
   Else 'unlimited
      intNumClients = UNLIMITED_CLIENTS
   End If
   
   If chkVehicleUnlimited.Value() = 0 Then
   
      On Error Resume Next
      intNumVehicles = CInt(txtVehicleLimit.Text())
      
      If Err.Number = 0 Then
         If intNumVehicles < 1 Then
            MsgBox "Invalid vehicles limit = """ & txtVehicleLimit.Text() & """"
            GoTo EXIT_SUB
         End If
      Else
         MsgBox "Invalid vehicles limit = """ & txtVehicleLimit.Text() & """"
         GoTo EXIT_SUB
      End If
      
   Else 'unlimited
      intNumVehicles = UNLIMITED_VEHICLES
   End If
   
   If chkExpDateUnlimited.Value() = 0 Then
   
      On Error Resume Next
      datExpiration = CDate(Format$(txtExpirationDate.Text(), STD_DATE_FORMAT))
      
      If Err.Number = 0 Then
         If datExpiration <= Now() Then
            MsgBox "Invalid date, not in the future, = """ & txtExpirationDate.Text() & """"
            GoTo EXIT_SUB
         End If
      Else
         MsgBox "Invalid date = """ & txtExpirationDate.Text() & """"
         GoTo EXIT_SUB
      End If
      
   Else
      datExpiration = CDate(UNLIMITED_EXP_DATE)
   End If
   
   Dim strLicKey As String
   strLicKey = clsLicKeyCreator.Generate(Now(), datExpiration, intNumClients, intNumVehicles)
   
   If strLicKey <> vbNullString Then
      txtLicenseKey.Text = strLicKey
   Else
      txtLicenseKey.Text = "Problems generating license key"
   End If
      
EXIT_SUB:
   Exit Sub

End Sub


Private Sub btnClearFields_Click()

   txtGeneratedDate.Text = vbNullString
   txtClientLimit.Text = vbNullString
   txtVehicleLimit.Text = vbNullString
   txtExpirationDate.Text = vbNullString
   
   chkClientUnlimited.Value = 0
   chkVehicleUnlimited.Value = 0
   chkExpDateUnlimited.Value = 0
   
End Sub


Private Sub btnOK_Click()

   Unload Me

End Sub


Private Sub chkClientUnlimited_Click()

   txtClientLimit.Enabled = IIf(chkClientUnlimited.Value() = 0, True, False)

End Sub


Private Sub chkExpDateUnlimited_Click()

   txtExpirationDate.Enabled = IIf(chkExpDateUnlimited.Value() = 0, True, False)

End Sub


Private Sub chkVehicleUnlimited_Click()

   txtVehicleLimit.Enabled = IIf(chkVehicleUnlimited.Value() = 0, True, False)
   
End Sub
