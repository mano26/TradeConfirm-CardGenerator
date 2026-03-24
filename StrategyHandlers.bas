Attribute VB_Name = "StrategyHandlers"
Option Explicit

Public Function GetExpiry(code As String, Optional isFuture As Boolean = False) As String
    Dim suffix As String, monthCode As String, yearCode As String
    Dim monthName As String, yearNum As Integer, digit As Integer
    Dim currentYear As Integer

    currentYear = year(Date)
    suffix = Right$(code, 2)
    monthCode = UCase$(Left$(suffix, 1))
    yearCode = Right$(suffix, 1)

    Select Case monthCode
        Case "F": monthName = "JAN"
        Case "G": monthName = "FEB"
        Case "H": monthName = "MAR"
        Case "J": monthName = "APR"
        Case "K": monthName = "MAY"
        Case "M": monthName = "JUN"
        Case "N": monthName = "JUL"
        Case "Q": monthName = "AUG"
        Case "U": monthName = "SEP"
        Case "V": monthName = "OCT"
        Case "X": monthName = "NOV"
        Case "Z": monthName = "DEC"
        Case Else
            MsgBox "Unrecognised month code '" & monthCode & "' in contract '" & code & "'." & vbNewLine & _
                   "Expected F G H J K M N Q U V X Z.", vbCritical
            monthName = "???"
    End Select

    digit = Val(yearCode)
    yearNum = 2020 + digit
    If digit < 5 Then yearNum = yearNum + 10

    If yearNum < currentYear Then
        Select Case monthCode
            Case "H", "M", "U", "Z": yearNum = yearNum + 10
            Case Else:                yearNum = yearNum + 1
        End Select
    End If
    If yearNum > currentYear + 10 Then yearNum = currentYear + 10

    If isFuture Then
        Select Case monthCode
            Case "F", "G", "H": monthName = "MAR"
            Case "J", "K", "M": monthName = "JUN"
            Case "N", "Q", "U": monthName = "SEP"
            Case "V", "X", "Z": monthName = "DEC"
        End Select

        Dim offset As Integer: offset = 0
        If Len(code) = 4 And IsNumeric(Left$(UCase$(code), 1)) Then
            offset = CInt(Left$(UCase$(code), 1))
            If offset = 0 Then offset = 1
            yearNum = yearNum + offset
            If yearNum > currentYear + 10 Then yearNum = currentYear + 10
        End If
    End If

    GetExpiry = monthName & Right$(CStr(yearNum), 2)
End Function

Public Function GetContractType(code As String, Optional isFuture As Boolean = False) As String
    Dim u As String: u = UCase$(code)

    If isFuture Then
        GetContractType = "SR3"
        Exit Function
    End If

    If Left$(u, 3) = "SR3" Or Left$(u, 3) = "SFR" Then
        GetContractType = "SR3"
        Exit Function
    End If

    If Len(u) = 4 And IsNumeric(Left$(u, 1)) Then
        Select Case Left$(u, 1)
            Case "0": GetContractType = "S0": Exit Function
            Case "2": GetContractType = "S2": Exit Function
            Case "3": GetContractType = "S3": Exit Function
        End Select
    End If

    If Left$(u, 2) = "S0" Then
        GetContractType = "S0"
    ElseIf Left$(u, 2) = "S2" Then
        GetContractType = "S2"
    ElseIf Left$(u, 2) = "S3" Then
        GetContractType = "S3"
    Else
        MsgBox "Unrecognised contract code: '" & code & "'." & vbNewLine & _
               "Check your trade input and try again.", vbCritical
        GetContractType = "ERR"
    End If
End Function

Public Function GetCardMoCode(code As String, Optional isFuture As Boolean = False) As String
    If Not isFuture Then
        GetCardMoCode = UCase$(code)
        Exit Function
    End If

    Dim suffix As String:    suffix = Right$(code, 2)
    Dim monthCode As String: monthCode = UCase$(Left$(suffix, 1))
    Dim yearCode As String:  yearCode = Right$(suffix, 1)
    Dim qtrLetter As String
    Dim digit As Integer, yearNum As Integer
    Dim currentYear As Integer: currentYear = year(Date)

    Select Case monthCode
        Case "F", "G", "H": qtrLetter = "H"
        Case "J", "K", "M": qtrLetter = "M"
        Case "N", "Q", "U": qtrLetter = "U"
        Case "V", "X", "Z": qtrLetter = "Z"
        Case Else:           qtrLetter = monthCode
    End Select

    digit = Val(yearCode)
    yearNum = 2020 + digit
    If digit < 5 Then yearNum = yearNum + 10

    If yearNum < currentYear Then
        Select Case monthCode
            Case "H", "M", "U", "Z": yearNum = yearNum + 10
            Case Else:                yearNum = yearNum + 1
        End Select
    End If
    If yearNum > currentYear + 10 Then yearNum = currentYear + 10

    If Len(code) = 4 And IsNumeric(Left$(UCase$(code), 1)) Then
        Dim offset As Integer: offset = CInt(Left$(UCase$(code), 1))
        If offset = 0 Then offset = 1
        yearNum = yearNum + offset
        If yearNum > currentYear + 10 Then yearNum = currentYear + 10
    End If

    GetCardMoCode = "SFR" & qtrLetter & CStr(yearNum Mod 10)
End Function

Public Function ResolveStrikeSide(trade As TradeInput, _
                                  strikeIdx As Integer, _
                                  defaultSide As String) As String
    If Not trade.StrikeOverrides Is Nothing Then
        If strikeIdx <= trade.StrikeOverrides.Count Then
            Dim ov As String: ov = CStr(trade.StrikeOverrides(strikeIdx))
            If ov = "+" Then
                ResolveStrikeSide = "B"
                Exit Function
            ElseIf ov = "-" Then
                ResolveStrikeSide = "S"
                Exit Function
            End If
        End If
    End If
    ResolveStrikeSide = defaultSide
End Function

Private Sub PrintLeg(ws As Worksheet, rowOut As Long, side As String, _
                     vol As Double, code As String, trade As TradeInput, _
                     Optional isFuture As Boolean = False, _
                     Optional strike As Variant = "", _
                     Optional optType As String = "", _
                     Optional priceTicks As Variant = "")
    Dim expiryCode As String, typeCode As String

    expiryCode = trade.ContractCodes(1)
    If trade.ContractCodes.Count > 1 Then
        typeCode = trade.ContractCodes(2)
    Else
        typeCode = expiryCode
    End If

    ws.Cells(rowOut, 3).Value = side
    ws.Cells(rowOut, 4).Value = vol
    ws.Cells(rowOut, 5).Value = "CME"
    ws.Cells(rowOut, 6).Value = GetContractType(typeCode, isFuture)
    ws.Cells(rowOut, 7).Value = GetExpiry(expiryCode, isFuture)

    If IsMissing(strike) Or strike = "" Then
        ws.Cells(rowOut, 8).ClearContents
    Else
        ws.Cells(rowOut, 8).Value = Format$(CDbl(strike), "0.0000")
    End If

    If optType = "" Then
        ws.Cells(rowOut, 9).ClearContents
    Else
        ws.Cells(rowOut, 9).Value = UCase$(optType)
    End If

    If IsMissing(priceTicks) Or priceTicks = "" Then
        ws.Cells(rowOut, 10).ClearContents
    Else
        If isFuture Then
            ws.Cells(rowOut, 10).Value = CDbl(priceTicks)
        Else
            ws.Cells(rowOut, 10).Value = Round(CDbl(priceTicks), 4)
        End If
    End If

    ws.Cells(rowOut, 18).Value = "AXIS"

    ' Col T: raw card MO code — white font so invisible to user
    ws.Cells(rowOut, 20).Value = GetCardMoCode(expiryCode, isFuture)
    ws.Cells(rowOut, 20).Font.Color = RGB(255, 255, 255)
End Sub

Public Function BuildStraddle(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim k As Double: k = trade.strikes(1)
    Call PrintLeg(ws, startRow, trade.DirectionSide, trade.Volume, trade.ContractCodes(1), trade, False, k, "P")
    Call PrintLeg(ws, startRow + 1, trade.DirectionSide, trade.Volume, trade.ContractCodes(1), trade, False, k, "C")
    BuildStraddle = startRow + 2
End Function

Public Function BuildStrangle(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim loStr As Double, hiStr As Double
    loStr = WorksheetFunction.Min(trade.strikes(1), trade.strikes(2))
    hiStr = WorksheetFunction.Max(trade.strikes(1), trade.strikes(2))
    Call PrintLeg(ws, startRow, trade.DirectionSide, trade.Volume, trade.ContractCodes(1), trade, False, loStr, "P")
    Call PrintLeg(ws, startRow + 1, trade.DirectionSide, trade.Volume, trade.ContractCodes(1), trade, False, hiStr, "C")
    BuildStrangle = startRow + 2
End Function

Public Function BuildCallSpread(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim loStr As Double, hiStr As Double

    If trade.strikes Is Nothing Or trade.strikes.Count = 0 Then
        MsgBox "No strikes available for call spread.", vbCritical
        BuildCallSpread = startRow: Exit Function
    ElseIf trade.strikes.Count = 1 Then
        loStr = CDbl(trade.strikes(1))
        hiStr = CDbl(trade.strikes(1))
    Else
        loStr = WorksheetFunction.Min(trade.strikes(1), trade.strikes(2))
        hiStr = WorksheetFunction.Max(trade.strikes(1), trade.strikes(2))
    End If

    Dim firstVol As Double: firstVol = trade.Volume
    Dim secondVol As Double: secondVol = trade.Volume
    If trade.ratio1 > 0 Then firstVol = trade.ratio1 * trade.Volume
    If trade.ratio2 > 0 Then secondVol = trade.ratio2 * trade.Volume

    Dim side1 As String: side1 = ResolveStrikeSide(trade, 1, IIf(trade.DirectionSide = "S", "S", "B"))
    Dim side2 As String: side2 = ResolveStrikeSide(trade, 2, IIf(trade.DirectionSide = "S", "B", "S"))

    Call PrintLeg(ws, startRow, side1, firstVol, trade.ContractCodes(1), trade, False, loStr, "C")
    Call PrintLeg(ws, startRow + 1, side2, secondVol, trade.ContractCodes(1), trade, False, hiStr, "C")
    BuildCallSpread = startRow + 2
End Function

Public Function BuildPutSpread(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim loStr As Double, hiStr As Double

    If trade.strikes Is Nothing Or trade.strikes.Count = 0 Then
        MsgBox "No strikes available for put spread.", vbCritical
        BuildPutSpread = startRow: Exit Function
    ElseIf trade.strikes.Count = 1 Then
        loStr = CDbl(trade.strikes(1))
        hiStr = CDbl(trade.strikes(1))
    Else
        loStr = WorksheetFunction.Min(trade.strikes(1), trade.strikes(2))
        hiStr = WorksheetFunction.Max(trade.strikes(1), trade.strikes(2))
    End If

    Dim firstVol As Double: firstVol = trade.Volume
    Dim secondVol As Double: secondVol = trade.Volume
    If trade.ratio1 > 0 Then firstVol = trade.ratio1 * trade.Volume
    If trade.ratio2 > 0 Then secondVol = trade.ratio2 * trade.Volume

    Dim side1 As String: side1 = ResolveStrikeSide(trade, 1, IIf(trade.DirectionSide = "S", "S", "B"))
    Dim side2 As String: side2 = ResolveStrikeSide(trade, 2, IIf(trade.DirectionSide = "S", "B", "S"))

    Call PrintLeg(ws, startRow, side1, firstVol, trade.ContractCodes(1), trade, False, hiStr, "P")
    Call PrintLeg(ws, startRow + 1, side2, secondVol, trade.ContractCodes(1), trade, False, loStr, "P")
    BuildPutSpread = startRow + 2
End Function

Public Function BuildRiskReversal(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim loStr As Double, hiStr As Double
    loStr = WorksheetFunction.Min(trade.strikes(1), trade.strikes(2))
    hiStr = WorksheetFunction.Max(trade.strikes(1), trade.strikes(2))

    Dim putSide As String, callSide As String
    If trade.IsCallCentric Then
        putSide = IIf(trade.DirectionSide = "B", "S", "B")
        callSide = trade.DirectionSide
    ElseIf trade.IsPutCentric Then
        putSide = trade.DirectionSide
        callSide = IIf(trade.DirectionSide = "B", "S", "B")
    Else
        putSide = "S": callSide = "B"
    End If

    putSide = ResolveStrikeSide(trade, 1, putSide)
    callSide = ResolveStrikeSide(trade, 2, callSide)

    Call PrintLeg(ws, startRow, putSide, trade.Volume, trade.ContractCodes(1), trade, False, loStr, "P")
    Call PrintLeg(ws, startRow + 1, callSide, trade.Volume, trade.ContractCodes(1), trade, False, hiStr, "C")
    BuildRiskReversal = startRow + 2
End Function

Public Function BuildIronCondor(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim k(1 To 4) As Double: Dim i As Integer
    For i = 1 To 4: k(i) = trade.strikes(i): Next i
    Call PrintLeg(ws, startRow, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, k(1), "P")
    Call PrintLeg(ws, startRow + 1, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, k(2), "P")
    Call PrintLeg(ws, startRow + 2, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, k(3), "C")
    Call PrintLeg(ws, startRow + 3, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, k(4), "C")
    BuildIronCondor = startRow + 4
End Function

Public Function BuildIronButterfly(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim k(1 To 3) As Double: Dim i As Integer
    For i = 1 To 3: k(i) = trade.strikes(i): Next i
    Call PrintLeg(ws, startRow, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, k(1), "P")
    Call PrintLeg(ws, startRow + 1, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, k(2), "P")
    Call PrintLeg(ws, startRow + 2, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, k(2), "C")
    Call PrintLeg(ws, startRow + 3, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, k(3), "C")
    BuildIronButterfly = startRow + 4
End Function

Public Function BuildCallButterfly(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim s(1 To 3) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2): s(3) = trade.strikes(3)
    Call PrintLeg(ws, startRow, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, s(1), "C")
    Call PrintLeg(ws, startRow + 1, IIf(trade.DirectionSide = "B", "S", "B"), 2 * trade.Volume, trade.ContractCodes(1), trade, False, s(2), "C")
    Call PrintLeg(ws, startRow + 2, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, s(3), "C")
    BuildCallButterfly = startRow + 3
End Function

Public Function BuildPutButterfly(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim s(1 To 3) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2): s(3) = trade.strikes(3)
    Call PrintLeg(ws, startRow, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, s(1), "P")
    Call PrintLeg(ws, startRow + 1, IIf(trade.DirectionSide = "B", "S", "B"), 2 * trade.Volume, trade.ContractCodes(1), trade, False, s(2), "P")
    Call PrintLeg(ws, startRow + 2, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, s(3), "P")
    BuildPutButterfly = startRow + 3
End Function

Public Function BuildCallCondor(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim s(1 To 4) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2)
    s(3) = trade.strikes(3): s(4) = trade.strikes(4)
    Dim temp As Double, i As Integer, j As Integer
    For i = 1 To 3
        For j = i + 1 To 4
            If s(i) > s(j) Then temp = s(i): s(i) = s(j): s(j) = temp
        Next j
    Next i
    Call PrintLeg(ws, startRow, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, s(1), "C")
    Call PrintLeg(ws, startRow + 1, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, s(2), "C")
    Call PrintLeg(ws, startRow + 2, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, s(3), "C")
    Call PrintLeg(ws, startRow + 3, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, s(4), "C")
    BuildCallCondor = startRow + 4
End Function

Public Function BuildPutCondor(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim s(1 To 4) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2)
    s(3) = trade.strikes(3): s(4) = trade.strikes(4)
    Dim temp As Double, i As Integer, j As Integer
    For i = 1 To 3
        For j = i + 1 To 4
            If s(i) < s(j) Then temp = s(i): s(i) = s(j): s(j) = temp
        Next j
    Next i
    Call PrintLeg(ws, startRow, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, s(1), "P")
    Call PrintLeg(ws, startRow + 1, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, s(2), "P")
    Call PrintLeg(ws, startRow + 2, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, s(3), "P")
    Call PrintLeg(ws, startRow + 3, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, s(4), "P")
    BuildPutCondor = startRow + 4
End Function

Public Function BuildCallChristmasTree(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim s(1 To 3) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2): s(3) = trade.strikes(3)
    Call PrintLeg(ws, startRow, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, s(1), "C")
    Call PrintLeg(ws, startRow + 1, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, s(2), "C")
    Call PrintLeg(ws, startRow + 2, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, s(3), "C")
    BuildCallChristmasTree = startRow + 3
End Function

Public Function BuildPutChristmasTree(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim s(1 To 3) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2): s(3) = trade.strikes(3)
    Call PrintLeg(ws, startRow, IIf(trade.DirectionSide = "B", "B", "S"), trade.Volume, trade.ContractCodes(1), trade, False, s(1), "P")
    Call PrintLeg(ws, startRow + 1, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, s(2), "P")
    Call PrintLeg(ws, startRow + 2, IIf(trade.DirectionSide = "B", "S", "B"), trade.Volume, trade.ContractCodes(1), trade, False, s(3), "P")
    BuildPutChristmasTree = startRow + 3
End Function

Public Function BuildBoxSpread(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim kLow As Double, kHigh As Double
    kLow = WorksheetFunction.Min(trade.strikes(1), trade.strikes(2))
    kHigh = WorksheetFunction.Max(trade.strikes(1), trade.strikes(2))
    Call PrintLeg(ws, startRow, "B", trade.Volume, trade.ContractCodes(1), trade, False, kLow, "C")
    Call PrintLeg(ws, startRow + 1, "S", trade.Volume, trade.ContractCodes(1), trade, False, kHigh, "C")
    Call PrintLeg(ws, startRow + 2, "S", trade.Volume, trade.ContractCodes(1), trade, False, kLow, "P")
    Call PrintLeg(ws, startRow + 3, "B", trade.Volume, trade.ContractCodes(1), trade, False, kHigh, "P")
    BuildBoxSpread = startRow + 4
End Function

Public Function BuildCvdOverlay(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim baseVol As Double, futVol As Double
    Dim sideFut As String, px As Double

    baseVol = trade.Volume
    futVol = Round(baseVol * trade.DeltaPercent / 100, 0)
    px = trade.CVDPrice

    If futVol = 0 Then
        MsgBox "Futures hedge quantity calculated as zero." & vbNewLine & _
               "Check your delta percentage (D token).", vbExclamation
    End If

    If trade.Strategy = "single" Then
        Dim optType As String
        If Not trade.OptionTypes Is Nothing And trade.OptionTypes.Count > 0 Then
            optType = UCase$(trade.OptionTypes(1))
        Else
            optType = ""
        End If
        If optType = "P" Then
            sideFut = trade.DirectionSide
        Else
            sideFut = IIf(trade.DirectionSide = "B", "S", "B")
        End If
    ElseIf trade.IsPutCentric Then
        sideFut = IIf(trade.DirectionSide = "B", "B", "S")
    ElseIf trade.IsCallCentric Then
        sideFut = IIf(trade.DirectionSide = "B", "S", "B")
    ElseIf trade.Strategy = "ps" Then
        sideFut = IIf(trade.DirectionSide = "B", "B", "S")
    ElseIf trade.Strategy = "cs" Then
        sideFut = IIf(trade.DirectionSide = "B", "S", "B")
    Else
        sideFut = IIf(trade.DirectionSide = "B", "S", "B")
    End If

    ' CVD price override — absolute
    If trade.CVDHasOverride Then
        If trade.CVDOverrideSide = "+" Then sideFut = "B"
        If trade.CVDOverrideSide = "-" Then sideFut = "S"
    End If

    ' Delta direction override — absolute, final precedence
    If trade.DeltaOverride = "B" Then sideFut = "B"
    If trade.DeltaOverride = "S" Then sideFut = "S"

    Call PrintLeg(ws, startRow, sideFut, futVol, trade.ContractCodes(1), trade, True, , , px)
    BuildCvdOverlay = startRow + 1
End Function

Public Function BuildSingleOption(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim px As Double
    Dim legSide As String

    legSide = ResolveStrikeSide(trade, 1, trade.DirectionSide)

    If trade.SuppressPremium Then
        Call PrintLeg(ws, startRow, legSide, trade.Volume, trade.ContractCodes(1), _
                      trade, False, trade.strikes(1), trade.OptionTypes(1))
    Else
        px = Round(trade.Premium, 4)
        Call PrintLeg(ws, startRow, legSide, trade.Volume, trade.ContractCodes(1), _
                      trade, False, trade.strikes(1), trade.OptionTypes(1), px)
    End If
    BuildSingleOption = startRow + 1
End Function

