VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WsSms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarMobileNumber As String 'local copy
Private mvarSmsText As String 'local copy
Private mvarSmsType As String 'local copy
Public Property Let SmsType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SmsType = 5
    mvarSmsType = vData
End Property


Public Property Get SmsType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SmsType
    SmsType = mvarSmsType
End Property

Public Property Let smsText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SmsText = 5
    mvarSmsText = vData
End Property


Public Property Get smsText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SmsText
    smsText = mvarSmsText
End Property



Public Property Let MobileNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MobileNumber = 5
    mvarMobileNumber = vData
End Property


Public Property Get MobileNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MobileNumber
    MobileNumber = mvarMobileNumber
End Property




