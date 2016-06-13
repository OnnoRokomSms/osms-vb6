# osms-vb6
Send SMS from VB6 code using OnnoRokom SMS API

Register free demo account from https://onnorokomsms.com

Install prerequisits first then use given sample project source code.

'SmsType = "TEXT"
'SmsType = "UCS" 'For Bangla SMS

Dim client As New SoapClient
Dim SmsType As String

// Check balance example

Private Sub checkBalanceBtn_Click()

    client.mssoapinit (http://api2.onnorokomsms.com/SendSms.asmx?wsdl)
    
    Dim balance As Variant
    balance = client.GetBalance("your username", "your password")
    statusLbl.Caption = "Connected. Your Balance is " & balance
    
End Sub

// Bulk SMS Sending example

Private Sub sendBulkSms_Click()
Dim rsp As Variant
    rsp = client.OneToMany(userName.Text, password.Text, smsText.Text, mobileNumbers.Text, SmsType, maskNameTxt.Text, "")
    statusLbl.Caption = rsp
End Sub

/// Single SMS Sending example

Private Sub sendSingleSms_Click()
    Dim rsp As Variant
    rsp = client.OneToOne(userName.Text, password.Text, mobileNumbers.Text, smsText.Text, SmsType, maskNameTxt.Text, "")
    statusLbl.Caption = rsp
End Sub



