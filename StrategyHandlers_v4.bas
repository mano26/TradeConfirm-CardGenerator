Attribute VB_Name = "StrategyHandlers_v4"

Option Explicit

Public Function GetExpiry(code As String, Optional isFuture As Boolean = False) As String
    Dim suffix As String, monthCode As String, yearCode As String
    Dim monthName As String, yearNum As Integer, digit As Integer
    Dim currentYear As Integer
    currentYear = Year(Date)
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
        Case Else: monthName = "???"
    End Select
    digit = Val(yearCode)
    yearNum = 2020 + digit
    If digit < 5 Then yearNum = yearNum + 10
    If yearNum < currentYear Then
        Select Case monthCode
            Case "H", "M", "U", "Z": yearNum = yearNum + 10
            Case Else: yearNum = yearNum + 1
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
        GetContractType = "ERR"
    End If
End Function

Public Function GetCardMoCode(code As String, Optional isFuture As Boolean = False) As String
    If Not isFuture Then
        GetCardMoCode = UCase$(code)
        Exit Function
    End If
    Dim suffix As String: suffix = Right$(code, 2)
    Dim monthCode As String: monthCode = UCase$(Left$(suffix, 1))
    Dim yearCode As String: yearCode = Right$(suffix, 1)
    Dim qtrLetter As String
    Dim digit As Integer, yearNum As Integer
    Dim currentYear As Integer: currentYear = Year(Date)
    Select Case monthCode
        Case "F", "G", "H": qtrLetter = "H"
        Case "J", "K", "M": qtrLetter = "M"
        Case "N", "Q", "U": qtrLetter = "U"
        Case "V", "X", "Z": qtrLetter = "Z"
        Case Else: qtrLetter = monthCode
    End Select
    digit = Val(yearCode)
    yearNum = 2020 + digit
    If digit < 5 Then yearNum = yearNum + 10
    If yearNum < currentYear Then
        Select Case monthCode
            Case "H", "M", "U", "Z": yearNum = yearNum + 10
            Case Else: yearNum = yearNum + 1
        End Select
    End If
    If yearNum > currentYear + 10 Then yearNum = currentYear + 10
    If Len(code) = 4 And IsNumeric(Left$(UCase$(code), 1)) Then
        Dim o As Integer: o = CInt(Left$(UCase$(code), 1))
        If o = 0 Then o = 1
        yearNum = yearNum + o
        If yearNum > currentYear + 10 Then yearNum = currentYear + 10
    End If
    GetCardMoCode = "SFR" & qtrLetter & CStr(yearNum Mod 10)
End Function

Public Function ResolveStrikeSide(trade As TradeInput, strikeIdx As Integer, defaultSide As String) As String
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
    ws.Cells(rowOut, S1_COL_SIDE).Value = side
    ws.Cells(rowOut, S1_COL_VOL).Value = vol
    ws.Cells(rowOut, S1_COL_MARKET).Value = "CME"
    ws.Cells(rowOut, S1_COL_CONTRACT).Value = GetContractType(typeCode, isFuture)
    ws.Cells(rowOut, S1_COL_EXPIRY).Value = GetExpiry(expiryCode, isFuture)
    If IsMissing(strike) Or strike = "" Then
        ws.Cells(rowOut, S1_COL_STRIKE).ClearContents
    Else
        ws.Cells(rowOut, S1_COL_STRIKE).Value = Format$(CDbl(strike), "0.0000")
    End If
    If optType = "" Then
        ws.Cells(rowOut, S1_COL_OPTTYPE).ClearContents
    Else
        ws.Cells(rowOut, S1_COL_OPTTYPE).Value = UCase$(optType)
    End If
    If IsMissing(priceTicks) Or priceTicks = "" Then
        ws.Cells(rowOut, S1_COL_PRICE).ClearContents
    Else
        If isFuture Then
            ws.Cells(rowOut, S1_COL_PRICE).Value = CDbl(priceTicks)
        Else
            ws.Cells(rowOut, S1_COL_PRICE).Value = Round(CDbl(priceTicks), 4)
        End If
    End If
    ws.Cells(rowOut, S1_COL_BROKER_STAMP).Value = "AXIS"
    ws.Cells(rowOut, S1_COL_MO_CARD).Value = GetCardMoCode(expiryCode, isFuture)
    ws.Cells(rowOut, S1_COL_MO_CARD).Font.Color = RGB(255, 255, 255)
End Sub

Public Function BuildStraddle(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim k As Double: k = trade.strikes(1)
    Call PrintLeg(ws, startRow, trade.DirectionSide, trade.Volume, trade.ContractCodes(1), trade, False, k, "P")
    Call PrintLeg(ws, startRow + 1, trade.DirectionSide, trade.Volume, trade.ContractCodes(1), trade, False, k, "C")
    BuildStraddle = startRow + 2
End Function

Public Function BuildStrangle(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim lo As Double, hi As Double
    lo = WorksheetFunction.Min(trade.strikes(1), trade.strikes(2))
    hi = WorksheetFunction.Max(trade.strikes(1), trade.strikes(2))
    Call PrintLeg(ws, startRow, trade.DirectionSide, trade.Volume, trade.ContractCodes(1), trade, False, lo, "P")
    Call PrintLeg(ws, startRow + 1, trade.DirectionSide, trade.Volume, trade.ContractCodes(1), trade, False, hi, "C")
    BuildStrangle = startRow + 2
End Function

Public Function BuildCallSpread(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim lo As Double, hi As Double
    If trade.strikes Is Nothing Or trade.strikes.Count = 0 Then
        MsgBox "No strikes available for call spread.", vbCritical
        BuildCallSpread = startRow
        Exit Function
    ElseIf trade.strikes.Count = 1 Then
        lo = CDbl(trade.strikes(1))
        hi = CDbl(trade.strikes(1))
    Else
        lo = WorksheetFunction.Min(trade.strikes(1), trade.strikes(2))
        hi = WorksheetFunction.Max(trade.strikes(1), trade.strikes(2))
    End If
    Dim firstVol As Double: firstVol = trade.Volume
    Dim secondVol As Double: secondVol = trade.Volume
    If trade.ratio1 > 0 Then firstVol = trade.ratio1 * trade.Volume
    If trade.ratio2 > 0 Then secondVol = trade.ratio2 * trade.Volume
    Dim side1 As String
    If trade.DirectionSide = "S" Then
        side1 = "S"
    Else
        side1 = "B"
    End If
    side1 = ResolveStrikeSide(trade, 1, side1)
    Dim side2 As String
    If trade.DirectionSide = "S" Then
        side2 = "B"
    Else
        side2 = "S"
    End If
    side2 = ResolveStrikeSide(trade, 2, side2)
    Call PrintLeg(ws, startRow, side1, firstVol, trade.ContractCodes(1), trade, False, lo, "C")
    Call PrintLeg(ws, startRow + 1, side2, secondVol, trade.ContractCodes(1), trade, False, hi, "C")
    BuildCallSpread = startRow + 2
End Function

Public Function BuildPutSpread(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim lo As Double, hi As Double
    If trade.strikes Is Nothing Or trade.strikes.Count = 0 Then
        MsgBox "No strikes available for put spread.", vbCritical
        BuildPutSpread = startRow
        Exit Function
    ElseIf trade.strikes.Count = 1 Then
        lo = CDbl(trade.strikes(1))
        hi = CDbl(trade.strikes(1))
    Else
        lo = WorksheetFunction.Min(trade.strikes(1), trade.strikes(2))
        hi = WorksheetFunction.Max(trade.strikes(1), trade.strikes(2))
    End If
    Dim firstVol As Double: firstVol = trade.Volume
    Dim secondVol As Double: secondVol = trade.Volume
    If trade.ratio1 > 0 Then firstVol = trade.ratio1 * trade.Volume
    If trade.ratio2 > 0 Then secondVol = trade.ratio2 * trade.Volume
    Dim side1 As String
    If trade.DirectionSide = "S" Then
        side1 = "S"
    Else
        side1 = "B"
    End If
    side1 = ResolveStrikeSide(trade, 1, side1)
    Dim side2 As String
    If trade.DirectionSide = "S" Then
        side2 = "B"
    Else
        side2 = "S"
    End If
    side2 = ResolveStrikeSide(trade, 2, side2)
    Call PrintLeg(ws, startRow, side1, firstVol, trade.ContractCodes(1), trade, False, hi, "P")
    Call PrintLeg(ws, startRow + 1, side2, secondVol, trade.ContractCodes(1), trade, False, lo, "P")
    BuildPutSpread = startRow + 2
End Function

Public Function BuildRiskReversal(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim lo As Double, hi As Double
    lo = WorksheetFunction.Min(trade.strikes(1), trade.strikes(2))
    hi = WorksheetFunction.Max(trade.strikes(1), trade.strikes(2))
    Dim putSide As String, callSide As String
    If trade.IsCallCentric Then
        putSide = IIf(trade.DirectionSide = "B", "S", "B")
        callSide = trade.DirectionSide
    ElseIf trade.IsPutCentric Then
        putSide = trade.DirectionSide
        callSide = IIf(trade.DirectionSide = "B", "S", "B")
    Else
        putSide = "S"
        callSide = "B"
    End If
    putSide = ResolveStrikeSide(trade, 1, putSide)
    callSide = ResolveStrikeSide(trade, 2, callSide)
    Call PrintLeg(ws, startRow, putSide, trade.Volume, trade.ContractCodes(1), trade, False, lo, "P")
    Call PrintLeg(ws, startRow + 1, callSide, trade.Volume, trade.ContractCodes(1), trade, False, hi, "C")
    BuildRiskReversal = startRow + 2
End Function

Public Function BuildIronCondor(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim k(1 To 4) As Double
    Dim i As Integer
    For i = 1 To 4: k(i) = trade.strikes(i): Next i
    Dim buySide As String, sellSide As String
    If trade.DirectionSide = "B" Then
        buySide = "B"
        sellSide = "S"
    Else
        buySide = "S"
        sellSide = "B"
    End If
    Call PrintLeg(ws, startRow, buySide, trade.Volume, trade.ContractCodes(1), trade, False, k(1), "P")
    Call PrintLeg(ws, startRow + 1, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, k(2), "P")
    Call PrintLeg(ws, startRow + 2, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, k(3), "C")
    Call PrintLeg(ws, startRow + 3, buySide, trade.Volume, trade.ContractCodes(1), trade, False, k(4), "C")
    BuildIronCondor = startRow + 4
End Function

Public Function BuildIronButterfly(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim k(1 To 3) As Double
    Dim i As Integer
    For i = 1 To 3: k(i) = trade.strikes(i): Next i
    Dim buySide As String, sellSide As String
    If trade.DirectionSide = "B" Then
        buySide = "B"
        sellSide = "S"
    Else
        buySide = "S"
        sellSide = "B"
    End If
    Call PrintLeg(ws, startRow, buySide, trade.Volume, trade.ContractCodes(1), trade, False, k(1), "P")
    Call PrintLeg(ws, startRow + 1, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, k(2), "P")
    Call PrintLeg(ws, startRow + 2, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, k(2), "C")
    Call PrintLeg(ws, startRow + 3, buySide, trade.Volume, trade.ContractCodes(1), trade, False, k(3), "C")
    BuildIronButterfly = startRow + 4
End Function

Public Function BuildCallButterfly(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim s(1 To 3) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2): s(3) = trade.strikes(3)
    Dim buySide As String, sellSide As String
    If trade.DirectionSide = "B" Then
        buySide = "B"
        sellSide = "S"
    Else
        buySide = "S"
        sellSide = "B"
    End If
    Call PrintLeg(ws, startRow, buySide, trade.Volume, trade.ContractCodes(1), trade, False, s(1), "C")
    Call PrintLeg(ws, startRow + 1, sellSide, 2 * trade.Volume, trade.ContractCodes(1), trade, False, s(2), "C")
    Call PrintLeg(ws, startRow + 2, buySide, trade.Volume, trade.ContractCodes(1), trade, False, s(3), "C")
    BuildCallButterfly = startRow + 3
End Function

Public Function BuildPutButterfly(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim s(1 To 3) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2): s(3) = trade.strikes(3)
    Dim buySide As String, sellSide As String
    If trade.DirectionSide = "B" Then
        buySide = "B"
        sellSide = "S"
    Else
        buySide = "S"
        sellSide = "B"
    End If
    Call PrintLeg(ws, startRow, buySide, trade.Volume, trade.ContractCodes(1), trade, False, s(1), "P")
    Call PrintLeg(ws, startRow + 1, sellSide, 2 * trade.Volume, trade.ContractCodes(1), trade, False, s(2), "P")
    Call PrintLeg(ws, startRow + 2, buySide, trade.Volume, trade.ContractCodes(1), trade, False, s(3), "P")
    BuildPutButterfly = startRow + 3
End Function

Public Function BuildCallCondor(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim s(1 To 4) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2)
    s(3) = trade.strikes(3): s(4) = trade.strikes(4)
    Dim temp As Double, i As Integer, j As Integer
    For i = 1 To 3
        For j = i + 1 To 4
            If s(i) > s(j) Then temp = s(i): s(i) = s(j): s(j) = temp
        Next j
    Next i
    Dim buySide As String, sellSide As String
    If trade.DirectionSide = "B" Then
        buySide = "B"
        sellSide = "S"
    Else
        buySide = "S"
        sellSide = "B"
    End If
    Call PrintLeg(ws, startRow, buySide, trade.Volume, trade.ContractCodes(1), trade, False, s(1), "C")
    Call PrintLeg(ws, startRow + 1, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, s(2), "C")
    Call PrintLeg(ws, startRow + 2, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, s(3), "C")
    Call PrintLeg(ws, startRow + 3, buySide, trade.Volume, trade.ContractCodes(1), trade, False, s(4), "C")
    BuildCallCondor = startRow + 4
End Function

Public Function BuildPutCondor(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim s(1 To 4) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2)
    s(3) = trade.strikes(3): s(4) = trade.strikes(4)
    Dim temp As Double, i As Integer, j As Integer
    For i = 1 To 3
        For j = i + 1 To 4
            If s(i) < s(j) Then temp = s(i): s(i) = s(j): s(j) = temp
        Next j
    Next i
    Dim buySide As String, sellSide As String
    If trade.DirectionSide = "B" Then
        buySide = "B"
        sellSide = "S"
    Else
        buySide = "S"
        sellSide = "B"
    End If
    Call PrintLeg(ws, startRow, buySide, trade.Volume, trade.ContractCodes(1), trade, False, s(1), "P")
    Call PrintLeg(ws, startRow + 1, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, s(2), "P")
    Call PrintLeg(ws, startRow + 2, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, s(3), "P")
    Call PrintLeg(ws, startRow + 3, buySide, trade.Volume, trade.ContractCodes(1), trade, False, s(4), "P")
    BuildPutCondor = startRow + 4
End Function

Public Function BuildCallChristmasTree(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim s(1 To 3) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2): s(3) = trade.strikes(3)
    Dim buySide As String, sellSide As String
    If trade.DirectionSide = "B" Then
        buySide = "B"
        sellSide = "S"
    Else
        buySide = "S"
        sellSide = "B"
    End If
    Call PrintLeg(ws, startRow, buySide, trade.Volume, trade.ContractCodes(1), trade, False, s(1), "C")
    Call PrintLeg(ws, startRow + 1, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, s(2), "C")
    Call PrintLeg(ws, startRow + 2, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, s(3), "C")
    BuildCallChristmasTree = startRow + 3
End Function

Public Function BuildPutChristmasTree(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim s(1 To 3) As Double
    s(1) = trade.strikes(1): s(2) = trade.strikes(2): s(3) = trade.strikes(3)
    Dim buySide As String, sellSide As String
    If trade.DirectionSide = "B" Then
        buySide = "B"
        sellSide = "S"
    Else
        buySide = "S"
        sellSide = "B"
    End If
    Call PrintLeg(ws, startRow, buySide, trade.Volume, trade.ContractCodes(1), trade, False, s(1), "P")
    Call PrintLeg(ws, startRow + 1, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, s(2), "P")
    Call PrintLeg(ws, startRow + 2, sellSide, trade.Volume, trade.ContractCodes(1), trade, False, s(3), "P")
    BuildPutChristmasTree = startRow + 3
End Function

Public Function BuildBoxSpread(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
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
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim futVol As Double, sideFut As String, px As Double
    futVol = Round(trade.Volume * trade.DeltaPercent / 100, 0)
    px = trade.CVDPrice
    
    If trade.Strategy = "single" Then
        Dim optType As String: optType = ""
        If Not trade.OptionTypes Is Nothing And trade.OptionTypes.Count > 0 Then
            optType = UCase$(trade.OptionTypes(1))
        End If
        If optType = "P" Then
            sideFut = trade.DirectionSide
        Else
            If trade.DirectionSide = "B" Then
                sideFut = "S"
            Else
                sideFut = "B"
            End If
        End If
    ElseIf trade.IsPutCentric Then
        If trade.DirectionSide = "B" Then
            sideFut = "B"
        Else
            sideFut = "S"
        End If
    ElseIf trade.IsCallCentric Then
        If trade.DirectionSide = "B" Then
            sideFut = "S"
        Else
            sideFut = "B"
        End If
    ElseIf trade.Strategy = "ps" Then
        If trade.DirectionSide = "B" Then
            sideFut = "B"
        Else
            sideFut = "S"
        End If
    ElseIf trade.Strategy = "cs" Then
        If trade.DirectionSide = "B" Then
            sideFut = "S"
        Else
            sideFut = "B"
        End If
    Else
        If trade.DirectionSide = "B" Then
            sideFut = "S"
        Else
            sideFut = "B"
        End If
    End If
    
    If trade.CVDHasOverride Then
        If trade.CVDOverrideSide = "+" Then sideFut = "B"
        If trade.CVDOverrideSide = "-" Then sideFut = "S"
    End If
    If trade.DeltaOverride = "B" Then sideFut = "B"
    If trade.DeltaOverride = "S" Then sideFut = "S"
    
    Call PrintLeg(ws, startRow, sideFut, futVol, trade.ContractCodes(1), trade, True, , , px)
    BuildCvdOverlay = startRow + 1
End Function

Public Function BuildSingleOption(trade As TradeInput, startRow As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH1_NAME)
    Dim legSide As String
    legSide = ResolveStrikeSide(trade, 1, trade.DirectionSide)
    If trade.SuppressPremium Then
        Call PrintLeg(ws, startRow, legSide, trade.Volume, trade.ContractCodes(1), _
                      trade, False, trade.strikes(1), trade.OptionTypes(1))
    Else
        Dim px As Double: px = Round(trade.Premium, 4)
        Call PrintLeg(ws, startRow, legSide, trade.Volume, trade.ContractCodes(1), _
                      trade, False, trade.strikes(1), trade.OptionTypes(1), px)
    End If
    BuildSingleOption = startRow + 1
End Function

