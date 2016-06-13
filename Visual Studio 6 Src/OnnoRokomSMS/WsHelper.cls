VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "messageHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarUserName As String 'local copy
Private mvarUserPassword As String 'local copy
Private mvarMarskText As String 'local copy
Private mvarCampingName As String 'local copy
Public Property Let CampingName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CampingName = 5
    mvarCampingName = vData
End Property


Public Property Get CampingName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CampingName
    CampingName = mvarCampingName
End Property



Public Property Let MarskText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MarskText = 5
    mvarMarskText = vData
End Property


Public Property Get MarskText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MarskText
    MarskText = mvarMarskText
End Property



Public Property Let UserPassword(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserPassword = 5
    mvarUserPassword = vData
End Property


Public Property Get UserPassword() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserPassword
    UserPassword = mvarUserPassword
End Property



Public Property Let userName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserName = 5
    mvarUserName = vData
End Property


Public Property Get userName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    userName = mvarUserName
End Property



      
