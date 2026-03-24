Attribute VB_Name = "ContractMap"
Option Explicit

Private Type PackMap
    code As String ' e.g. "0QZ5"
    Pack As String ' S0, S2, S3
    offset As Long ' 1, 2, 3
End Type
Private Const MAP_COUNT As Long = 12
Private PackTable(1 To MAP_COUNT) As PackMap
Private packInitDone As Boolean

Private Sub InitPackTable()
    If packInitDone Then Exit Sub

    Dim i As Long: i = 1

    ' RED PACK (S0)
    PackTable(i).code = "0QZ5": PackTable(i).Pack = "S0": PackTable(i).offset = 1: i = i + 1
    PackTable(i).code = "0QF6": PackTable(i).Pack = "S0": PackTable(i).offset = 1: i = i + 1
    PackTable(i).code = "0QG6": PackTable(i).Pack = "S0": PackTable(i).offset = 1: i = i + 1
    PackTable(i).code = "0QH6": PackTable(i).Pack = "S0": PackTable(i).offset = 1: i = i + 1

    ' GREEN PACK (S2)
    PackTable(i).code = "2QM6": PackTable(i).Pack = "S2": PackTable(i).offset = 2: i = i + 1
    PackTable(i).code = "2QN6": PackTable(i).Pack = "S2": PackTable(i).offset = 2: i = i + 1
    PackTable(i).code = "2QQ6": PackTable(i).Pack = "S2": PackTable(i).offset = 2: i = i + 1
    PackTable(i).code = "2QU6": PackTable(i).Pack = "S2": PackTable(i).offset = 2: i = i + 1

    ' BLUE PACK (S3)
    PackTable(i).code = "3QU6": PackTable(i).Pack = "S3": PackTable(i).offset = 3: i = i + 1
    PackTable(i).code = "3QV6": PackTable(i).Pack = "S3": PackTable(i).offset = 3: i = i + 1
    PackTable(i).code = "3QX6": PackTable(i).Pack = "S3": PackTable(i).offset = 3: i = i + 1
    PackTable(i).code = "3QZ6": PackTable(i).Pack = "S3": PackTable(i).offset = 3: i = i + 1

    packInitDone = True
End Sub

Public Function IsShortDatedContract(ByVal token As String) As Boolean
    Call InitPackTable
    Dim i As Long
    token = UCase$(token)
    For i = 1 To MAP_COUNT
        If PackTable(i).code = token Then
            IsShortDatedContract = True
            Exit Function
        End If
    Next i
End Function

Public Function PackCodeFromShortDated(ByVal token As String) As String
    Call InitPackTable
    Dim i As Long
    token = UCase$(token)
    For i = 1 To MAP_COUNT
        If PackTable(i).code = token Then
            PackCodeFromShortDated = PackTable(i).Pack
            Exit Function
        End If
    Next i
End Function

Public Function PackOffsetFromShortDated(ByVal token As String) As Long
    Call InitPackTable
    Dim i As Long
    token = UCase$(token)
    For i = 1 To MAP_COUNT
        If PackTable(i).code = token Then
            PackOffsetFromShortDated = PackTable(i).offset
            Exit Function
        End If
    Next i
End Function


