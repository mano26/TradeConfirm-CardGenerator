VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TradeInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ContractCodes As Collection
Public strikes As Collection
Public OptionTypes As Collection
Public StrikeOverrides As Collection

Public Strategy As String
Public IsStraddle As Boolean
Public IsStrangle As Boolean
Public IsSingleOption As Boolean
Public Volume As Long
Public Premium As Double
Public IsCVD As Boolean
Public CVDPrice As Double
Public CVDHasOverride As Boolean
Public CVDOverrideSide As String
Public DirectionSide As String
Public DeltaPercent As Double
Public DeltaOverride As String
Public IsCallCentric As Boolean
Public IsPutCentric As Boolean
Public ratio1 As Long
Public ratio2 As Long
Public legCount As Long
Public IsStupid As Boolean
Public SuppressPremium As Boolean

Private Sub Class_Initialize()
    Set ContractCodes = New Collection
    Set strikes = New Collection
    Set OptionTypes = New Collection
    Set StrikeOverrides = New Collection
    Strategy = ""
    IsStraddle = False
    IsStrangle = False
    IsSingleOption = False
    Volume = 0
    Premium = 0
    IsCVD = False
    CVDPrice = 0
    CVDHasOverride = False
    CVDOverrideSide = ""
    DirectionSide = ""
    DeltaPercent = 0
    DeltaOverride = ""
    IsCallCentric = False
    IsPutCentric = False
    ratio1 = 0
    ratio2 = 0
    legCount = 0
    IsStupid = False
    SuppressPremium = False
End Sub
