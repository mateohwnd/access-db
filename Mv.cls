VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private sCustomer, sPlate, sMVFile, sCOC, sName, sAddress, sChassis, sEngine, sBodyType As String
Private dDateIssued, dDateCoverStart, dDateCoverEnd As Date
Private dAmount As Double

Public Property Let Customer(value As String)
    sCustomer = value
End Property
Public Property Get Customer() As String
    Customer = sCustomer
End Property

Public Property Let PlateNumber(value As String)
    sPlate = value
End Property
Public Property Get PlateNumber() As String
    PlateNumber = sPlate
End Property

Public Property Let MvFileNumber(value As String)
    sMVFile = value
End Property
Public Property Get MvFileNumber() As String
    MvFileNumber = sMVFile
End Property

Public Property Let cocNumber(value As String)
    sCOC = value
End Property
Public Property Get cocNumber() As String
    cocNumber = sCOC
End Property

Public Property Let Name(value As String)
    sName = value
End Property
Public Property Get Name() As String
    Name = sName
End Property

Public Property Let Address(value As String)
    sAddress = value
End Property
Public Property Get Address() As String
    Address = sAddress
End Property

Public Property Let ChassisNumber(value As String)
    sChassis = value
End Property
Public Property Get ChassisNumber() As String
    ChassisNumber = sChassis
End Property

Public Property Let EngineNumber(value As String)
    sEngine = value
End Property
Public Property Get EngineNumber() As String
    EngineNumber = sEngine
End Property

Public Property Let BodyType(value As String)
    sBodyType = value
End Property
Public Property Get BodyType() As String
    BodyType = sBodyType
End Property

 Public Property Let DateIssued(value As Date)
    dDateIssued = value
End Property
Public Property Get DateIssued() As Date
    DateIssued = dDateIssued
End Property

Public Property Let DateCoverStart(value As Date)
    dDateCoverStart = value
End Property
Public Property Get DateCoverStart() As Date
    DateCoverStart = dDateCoverStart
End Property

Public Property Let DateCoverEnd(value As Date)
    dDateCoverEnd = value
End Property
Public Property Get DateCoverEnd() As Date
    DateCoverEnd = dDateCoverEnd
End Property

Public Property Let amount(value As Double)
    dAmount = value
End Property
Public Property Get amount() As Double
    amount = dAmount
End Property

Private Sub Class_Initialize()
    Customer = ""
    PlateNumber = ""
    MvFileNumber = ""
    cocNumber = ""
    Name = ""
    Address = ""
    ChassisNumber = ""
    EngineNumber = ""
    DateIssued = Date
    DateCoverStart = Date
    DateCoverEnd = Date
    amount = 0#
End Sub
