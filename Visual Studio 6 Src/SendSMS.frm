VERSION 5.00
Begin VB.Form SendSmsFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send SMS"
   ClientHeight    =   5460
   ClientLeft      =   2865
   ClientTop       =   2085
   ClientWidth     =   8910
   DrawWidth       =   10
   Icon            =   "SendSMS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox maskNameTxt 
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Text            =   "Your Mask"
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox mobileNumbers 
      Height          =   1815
      Left            =   5040
      TabIndex        =   9
      Text            =   "01811419557"
      Top             =   2880
      Width           =   3615
   End
   Begin VB.CommandButton exitBtn 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7320
      TabIndex        =   8
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton sendBulkSms 
      Caption         =   "Send Bulk"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton sendBatchSms 
      Caption         =   "Send Batch"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton sendSingleSms 
      Caption         =   "Send"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox smsText 
      Height          =   1815
      Left            =   360
      TabIndex        =   4
      Text            =   "SMS Test from Vb6"
      Top             =   2880
      Width           =   4455
   End
   Begin VB.CommandButton connectBtn 
      Caption         =   "Connect"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox password 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Text            =   "Your Password"
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox userName 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Text            =   "User Name"
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox apiUrl 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "http://api2.onnorokomsms.com/SendSms.asmx?wsdl"
      Top             =   360
      Width           =   8415
   End
   Begin VB.Label statusLbl 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   4200
      TabIndex        =   10
      Top             =   960
      Width           =   4455
   End
End
Attribute VB_Name = "SendSmsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim client As New SoapClient
Dim SmsType As String


Private Sub connectBtn_Click()
    
    SmsType = "TEXT"
    'SmsType = "UCS" 'For Bangla SMS
    
    client.mssoapinit (apiUrl.Text)
    
    Dim balance As Variant
    balance = client.GetBalance(userName.Text, password.Text)
    statusLbl.Caption = "Connected. Your Balance is " & balance
    
End Sub

Private Sub exitBtn_Click()
End
End Sub

Private Sub sendBatchSms_Click()

Dim header As New messageHeader
header.CampingName = ""
header.MarskText = maskNameTxt.Text
header.userName = userName.Text
header.UserPassword = password.Text

Dim smsList(2) As New WsSms

smsList(0).MobileNumber = "01811419557"
smsList(0).smsText = "SMS FROM VB6 BATCH"
smsList(0).SmsType = "TEXT"

smsList(1).MobileNumber = "01811419556"
smsList(1).smsText = "SMS FROM VB6 BATCH"
smsList(1).SmsType = "TEXT"


Dim rsp As Variant
    rsp = client.OneToOneBulk(header, smsList)
    statusLbl.Caption = rsp
End Sub

Private Sub sendBulkSms_Click()
Dim rsp As Variant
    rsp = client.OneToMany(userName.Text, password.Text, smsText.Text, mobileNumbers.Text, SmsType, maskNameTxt.Text, "")
    statusLbl.Caption = rsp
End Sub

Private Sub sendSingleSms_Click()
    Dim rsp As Variant
    rsp = client.OneToOne(userName.Text, password.Text, mobileNumbers.Text, smsText.Text, SmsType, maskNameTxt.Text, "")
    statusLbl.Caption = rsp
End Sub
