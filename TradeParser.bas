Attribute VB_Name = "TradeParser"
Option Explicit

Private Function IsPackHelperCode(ByVal code As String) As Boolean
    Dim u As String: u = UCase$(Trim$(code))
    IsPackHelperCode = (u = "S0" Or u = "S2" Or u = "S3" Or u = "SR3")
End Function

Private Sub AddPackHelperIfShortDated(ByVal t As TradeInput, ByVal code As String)
    Dim u As String: u = UCase$(Trim$(code))
    If Len(u) = 4 Then
        Select Case Left$(u, 1)
            Case "0": t.ContractCodes.Add "S0"
            Case "2": t.ContractCodes.Add "S2"
            Case "3": t.ContractCodes.Add "S3"
        End Select
    End If
End Sub

Private Function CloneTradeInput(ByVal src As TradeInput) As TradeInput
    Dim dst As New TradeInput
    Dim i As Long

    Set dst.ContractCodes = New Collection
    For i = 1 To src.ContractCodes.Count
        dst.ContractCodes.Add src.ContractCodes(i)
    Next i

    Set dst.strikes = New Collection
    For i = 1 To src.strikes.Count
        dst.strikes.Add src.strikes(i)
    Next i

    Set dst.OptionTypes = New Collection
    If Not src.OptionTypes Is Nothing Then
        For i = 1 To src.OptionTypes.Count
            dst.OptionTypes.Add src.OptionTypes(i)
        Next i
    End If

    Set dst.StrikeOverrides = New Collection
    If Not src.StrikeOverrides Is Nothing Then
        For i = 1 To src.StrikeOverrides.Count
            dst.StrikeOverrides.Add src.StrikeOverrides(i)
        Next i
    End If

    dst.Strategy = src.Strategy
    dst.IsStraddle = src.IsStraddle
    dst.IsStrangle = src.IsStrangle
    dst.IsSingleOption = src.IsSingleOption
    dst.Volume = src.Volume
    dst.Premium = src.Premium
    dst.IsCVD = src.IsCVD
    dst.CVDPrice = src.CVDPrice
    dst.CVDHasOverride = src.CVDHasOverride
    dst.CVDOverrideSide = src.CVDOverrideSide
    dst.DirectionSide = src.DirectionSide
    dst.DeltaPercent = src.DeltaPercent
    dst.DeltaOverride = src.DeltaOverride
    dst.IsCallCentric = src.IsCallCentric
    dst.IsPutCentric = src.IsPutCentric
    dst.ratio1 = src.ratio1
    dst.ratio2 = src.ratio2
    dst.legCount = src.legCount
    dst.IsStupid = src.IsStupid
    dst.SuppressPremium = src.SuppressPremium

    Set CloneTradeInput = dst
End Function

Private Function ValidateStrikes(t As TradeInput) As Boolean
    Dim expected As String
    Select Case t.Strategy
        Case "cs", "ps", "strangle", "rr"
            ValidateStrikes = (t.strikes.Count = 2) Or (t.strikes.Count = 1)
            expected = "1 or 2"
        Case "bflyc", "bflyp", "ctree", "ptree"
            ValidateStrikes = (t.strikes.Count = 3)
            expected = "3"
        Case "condorc", "condorp", "ic"
            ValidateStrikes = (t.strikes.Count = 4)
            expected = "4"
        Case "straddle", "c", "p", "single"
            ValidateStrikes = (t.strikes.Count >= 1)
            expected = "at least 1"
        Case Else
            ValidateStrikes = (t.strikes.Count >= 1)
            expected = "at least 1"
    End Select
    If Not ValidateStrikes Then
        MsgBox "Wrong number of strikes for strategy '" & t.Strategy & "'." & vbNewLine & _
               "Expected: " & expected & ", found: " & t.strikes.Count & "." & vbNewLine & _
               "Check your strike inputs.", vbCritical
    End If
End Function

Private Function SetStrategy(t As TradeInput, token As String, _
                             i As Long, tokens() As String) As Long
    Dim uToken As String: uToken = UCase$(Trim$(token))
    Dim NextTok As String
    Dim peek As Long
    SetStrategy = 0

    Select Case uToken

        Case "C"
            peek = i + 1
            If peek <= UBound(tokens) Then
                Select Case UCase$(Trim$(tokens(peek)))
                    Case "FLY":  t.Strategy = "bflyc":  SetStrategy = 1
                    Case "CON":  t.Strategy = "condorc": SetStrategy = 1
                    Case "TREE": t.Strategy = "ctree":   SetStrategy = 1
                End Select
            End If

        Case "P"
            peek = i + 1
            If peek <= UBound(tokens) Then
                Select Case UCase$(Trim$(tokens(peek)))
                    Case "FLY":  t.Strategy = "bflyp":  SetStrategy = 1
                    Case "CON":  t.Strategy = "condorp": SetStrategy = 1
                    Case "TREE": t.Strategy = "ptree":   SetStrategy = 1
                End Select
            End If

        Case "IRON", "IRONCONDOR", "IRONCOND", "IC"
            If uToken = "IC" Then
                t.Strategy = "ic"
            Else
                peek = i + 1
                If peek <= UBound(tokens) Then
                    Select Case UCase$(Trim$(tokens(peek)))
                        Case "CONDOR", "CON": t.Strategy = "ic":    SetStrategy = 1
                        Case "FLY":           t.Strategy = "ibfly": SetStrategy = 1
                        Case Else:            t.Strategy = "ic"
                    End Select
                Else
                    t.Strategy = "ic"
                End If
            End If

        Case "STUPID"
            t.IsStupid = True

        Case "CS", "CALLSPREAD", "CALLSP", "CSPD"
            t.Strategy = "cs"
        Case "PS", "PUTSPREAD", "PUTSP", "PSPD"
            t.Strategy = "ps"
        Case "RR", "RISKREV", "RISKREVERSE", "RV"
            t.Strategy = "rr"

        Case "CVD"
            t.IsCVD = True
            If i + 1 <= UBound(tokens) Then
                Dim cvdTok As String: cvdTok = Trim$(tokens(i + 1))
                Dim cvdNum As String: cvdNum = cvdTok
                Dim cvdOv  As String: cvdOv = ""
                Dim cvdP   As Integer: cvdP = InStr(cvdTok, "(")
                If cvdP > 0 Then
                    Dim cvdQ As Integer: cvdQ = InStr(cvdTok, ")")
                    If cvdQ > cvdP Then
                        cvdOv = Mid$(cvdTok, cvdP + 1, cvdQ - cvdP - 1)
                        cvdNum = Left$(cvdTok, cvdP - 1)
                    End If
                End If
                If IsNumeric(cvdNum) Then
                    t.CVDPrice = CDbl(cvdNum)
                    If cvdOv = "+" Or cvdOv = "-" Then
                        t.CVDHasOverride = True
                        t.CVDOverrideSide = cvdOv
                    End If
                    SetStrategy = 1
                Else
                    MsgBox "CVD token found but no valid price follows it." & vbNewLine & _
                           "Expected a number after CVD, got: '" & cvdTok & "'", vbCritical
                End If
            Else
                MsgBox "CVD token found but no price follows it. Add a futures price after CVD.", vbCritical
            End If

        Case "D"
            If i + 1 <= UBound(tokens) Then
                If IsNumeric(tokens(i + 1)) Then
                    t.DeltaPercent = CDbl(tokens(i + 1))
                    SetStrategy = 1
                Else
                    MsgBox "Delta token 'D' found but no quantity follows it." & vbNewLine & _
                           "Expected a number after D, got: '" & tokens(i + 1) & "'", vbCritical
                End If
            Else
                MsgBox "Delta token 'D' found at end of input with no quantity following it.", vbCritical
            End If

        Case "(CALLS)": t.IsCallCentric = True
        Case "(PUTS)":  t.IsPutCentric = True

        Case "CONDORC", "CONC", "CALLCONDOR"
            t.Strategy = "condorc"
        Case "CONDORP", "CONP", "PUTCONDOR"
            t.Strategy = "condorp"
        Case "CON", "CONDOR"
            peek = i + 1
            If peek <= UBound(tokens) Then
                NextTok = UCase$(Trim$(tokens(peek)))
                If NextTok = "C" Or NextTok = "CALL" Then
                    t.Strategy = "condorc": SetStrategy = 1
                ElseIf NextTok = "P" Or NextTok = "PUT" Then
                    t.Strategy = "condorp": SetStrategy = 1
                End If
            End If

        Case "BFLY", "BUTTERFLY", "FLY"
            If t.Strategy = "" Then t.Strategy = "bfly"
        Case "BFLYC", "CALLBFLY", "CALLFLY", "BUTTERFLYC"
            t.Strategy = "bflyc"
        Case "BFLYP", "PUTBFLY", "BUTTERFLYP"
            t.Strategy = "bflyp"

        Case "TREE", "CALLTREE", "CTREE", "TREEC", "XMAS", "CHRISTMAS", "CALLXMAS", "XMASC"
            t.Strategy = "ctree"
        Case "PUTTREE", "PTREE", "TREEP", "PUTXMAS", "PUTCHRISTMAS"
            t.Strategy = "ptree"

        Case "1X2", "1BY2": t.ratio1 = 1: t.ratio2 = 2
        Case "1X3", "1BY3": t.ratio1 = 1: t.ratio2 = 3
        Case "2X3", "2BY3": t.ratio1 = 2: t.ratio2 = 3
        Case "2X1", "2BY1": t.ratio1 = 2: t.ratio2 = 1

    End Select
End Function

Private Function IsContractCode(token As String) As Boolean
    Dim u As String: u = UCase$(token)
    If Len(u) = 4 Then
        Dim prefix As String: prefix = Left$(u, 1)
        Dim month  As String: month = Mid$(u, 2, 1)
        Dim yr     As String: yr = Right$(u, 1)
        If IsNumeric(prefix) And prefix >= "0" And prefix <= "3" And _
           (month = "F" Or month = "G" Or month = "H" Or month = "J" Or _
            month = "K" Or month = "M" Or month = "N" Or month = "Q" Or _
            month = "U" Or month = "V" Or month = "X" Or month = "Z") And _
           IsNumeric(yr) Then
            IsContractCode = True
            Exit Function
        End If
    End If
    IsContractCode = (Left$(u, 3) = "SR3" Or Left$(u, 3) = "SFR" _
                   Or Left$(u, 2) = "S0" Or Left$(u, 2) = "S2" _
                   Or Left$(u, 2) = "S3")
End Function

Private Function LegContainsCode(t As TradeInput, code As String) As Boolean
    Dim i As Long
    For i = 1 To t.ContractCodes.Count
        If UCase$(t.ContractCodes(i)) = UCase$(code) Then
            LegContainsCode = True: Exit Function
        End If
    Next i
End Function

Private Sub StripOverride(token As String, numStr As String, ovSide As String)
    numStr = token
    ovSide = ""
    Dim pOpen As Integer: pOpen = InStr(token, "(")
    If pOpen > 0 Then
        Dim pClose As Integer: pClose = InStr(token, ")")
        If pClose > pOpen Then
            ovSide = Mid$(token, pOpen + 1, pClose - pOpen - 1)
            numStr = Trim$(Left$(token, pOpen - 1))
            If ovSide <> "+" And ovSide <> "-" Then ovSide = ""
        End If
    End If
End Sub

Private Function ParseSingleLeg(tokens() As String) As TradeInput
    Dim t As New TradeInput
    Dim i As Long, token As String, uToken As String
    Dim consumed As Long
    Dim hasCall As Boolean, hasPut As Boolean

    For i = LBound(tokens) To UBound(tokens)
        token = Trim$(tokens(i))
        If token = "" Then GoTo NextToken
        uToken = UCase$(token)

        If uToken = "CVD" Then
            consumed = SetStrategy(t, token, i, tokens)
            If consumed > 0 Then i = i + consumed
            GoTo NextToken
        End If

        If uToken = "D" Then
            If i + 1 <= UBound(tokens) Then
                Dim dTok As String: dTok = Trim$(tokens(i + 1))
                If IsNumeric(dTok) Then
                    t.DeltaPercent = CDbl(dTok)
                    i = i + 1
                    If i + 1 <= UBound(tokens) Then
                        Dim ovTok As String: ovTok = Trim$(tokens(i + 1))
                        If ovTok = "(+)" Or ovTok = "(-)" Then
                            t.DeltaOverride = IIf(ovTok = "(+)", "B", "S")
                            i = i + 1
                        End If
                    End If
                Else
                    MsgBox "Delta token 'D' found but no quantity follows it." & vbNewLine & _
                           "Expected a number, got: '" & dTok & "'", vbCritical
                End If
            Else
                MsgBox "Delta token 'D' found at end of input with no quantity following it.", vbCritical
            End If
            GoTo NextToken
        End If

        If IsContractCode(uToken) Then
            If t.ContractCodes Is Nothing Then Set t.ContractCodes = New Collection
            t.ContractCodes.Add token
            If Len(uToken) = 4 And Left$(uToken, 1) Like "[023]" Then
                Select Case Left$(uToken, 1)
                    Case "0": t.ContractCodes.Add "S0"
                    Case "2": t.ContractCodes.Add "S2"
                    Case "3": t.ContractCodes.Add "S3"
                End Select
            End If
            GoTo NextToken
        End If

        Dim numStr As String, ovSide As String
        Call StripOverride(token, numStr, ovSide)

        If IsNumeric(numStr) And InStr(numStr, ".") > 0 Then
            Dim prevTok As String: prevTok = ""
            If i > LBound(tokens) Then prevTok = UCase$(Trim$(tokens(i - 1)))
            If prevTok <> "CVD" And prevTok <> "D" Then
                If t.strikes Is Nothing Then Set t.strikes = New Collection
                If t.StrikeOverrides Is Nothing Then Set t.StrikeOverrides = New Collection
                t.strikes.Add CDbl(numStr)
                t.StrikeOverrides.Add ovSide
            End If
            GoTo NextToken
        End If

        If uToken = "^" Then
            t.IsStraddle = True: t.Strategy = "straddle"
            If t.OptionTypes Is Nothing Then Set t.OptionTypes = New Collection
            t.OptionTypes.Add "P": t.OptionTypes.Add "C"
            GoTo NextToken
        End If

        If uToken = "^^" Then
            t.IsStrangle = True: t.Strategy = "strangle"
            If t.OptionTypes Is Nothing Then Set t.OptionTypes = New Collection
            t.OptionTypes.Add "P": t.OptionTypes.Add "C"
            GoTo NextToken
        End If

        ' Standalone (+)/(-) — attach to most recently added strike
        If uToken = "(-)" Or uToken = "(+)" Then
            If Not t.StrikeOverrides Is Nothing Then
                If t.StrikeOverrides.Count > 0 Then
                    Dim newOv As String: newOv = IIf(uToken = "(+)", "+", "-")
                    Dim tempColl As New Collection
                    Dim si As Long
                    For si = 1 To t.StrikeOverrides.Count - 1
                        tempColl.Add t.StrikeOverrides(si)
                    Next si
                    tempColl.Add newOv
                    Set t.StrikeOverrides = tempColl
                End If
            End If
            GoTo NextToken
        End If

        If uToken = "C" Then hasCall = True
        If uToken = "P" Then hasPut = True
        If uToken = "CALL" Then hasCall = True
        If uToken = "PUT" Then hasPut = True

        consumed = SetStrategy(t, token, i, tokens)
        If consumed > 0 Then i = i + consumed

NextToken:
    Next i

    If t.Strategy = "bfly" Or t.Strategy = "ctree" Or t.Strategy = "ptree" Or _
       t.Strategy = "condor" Or t.Strategy = "condorc" Or t.Strategy = "condorp" Or _
       t.Strategy = "ibfly" Or t.Strategy = "cs" Or t.Strategy = "ps" Then

        If t.OptionTypes Is Nothing Then Set t.OptionTypes = New Collection

        If hasPut And Not hasCall Then
            If t.Strategy = "bfly" Or t.Strategy = "bflyc" Then t.Strategy = "bflyp"
            If t.Strategy = "ctree" Then t.Strategy = "ptree"
            If t.Strategy = "condor" Or t.Strategy = "condorc" Then t.Strategy = "condorp"
            If t.OptionTypes.Count = 0 Then t.OptionTypes.Add "P"
        Else
            If t.Strategy = "bfly" Or t.Strategy = "bflyp" Then t.Strategy = "bflyc"
            If t.Strategy = "ptree" Then t.Strategy = "ctree"
            If t.Strategy = "condor" Or t.Strategy = "condorp" Then t.Strategy = "condorc"
            If t.OptionTypes.Count = 0 Then t.OptionTypes.Add "C"
        End If
    End If

    If t.Strategy = "" And t.ratio1 > 0 And t.ratio2 > 0 Then
        If hasCall And Not hasPut Then
            t.Strategy = "cs"
            If t.OptionTypes Is Nothing Then Set t.OptionTypes = New Collection
            t.OptionTypes.Add "C": t.IsCallCentric = True
        ElseIf hasPut And Not hasCall Then
            t.Strategy = "ps"
            If t.OptionTypes Is Nothing Then Set t.OptionTypes = New Collection
            t.OptionTypes.Add "P": t.IsPutCentric = True
        End If
    End If

    If t.Strategy = "bflyc" Or t.Strategy = "ctree" Or t.Strategy = "condorc" Or _
       t.Strategy = "ibfly" Or t.Strategy = "cs" Then t.IsCallCentric = True
    If t.Strategy = "bflyp" Or t.Strategy = "ptree" Or t.Strategy = "condorp" Or _
       t.Strategy = "ps" Then t.IsPutCentric = True

    If t.Strategy = "" Then
        If hasCall And Not hasPut Then
            t.Strategy = "single"
            If t.OptionTypes Is Nothing Then Set t.OptionTypes = New Collection
            t.OptionTypes.Add "C"
        ElseIf hasPut And Not hasCall Then
            t.Strategy = "single"
            If t.OptionTypes Is Nothing Then Set t.OptionTypes = New Collection
            t.OptionTypes.Add "P"
        Else
            MsgBox "Could not determine strategy from input." & vbNewLine & _
                   "No recognised option type (C/P) or strategy token found." & vbNewLine & _
                   "Check your input syntax.", vbCritical
            Set ParseSingleLeg = Nothing
            Exit Function
        End If
    End If

    If t.ContractCodes Is Nothing Or t.ContractCodes.Count = 0 Then
        MsgBox "No contract code found in input." & vbNewLine & _
               "Expected a code like SFRM6, 2QK6, SR3, etc.", vbCritical
        Set ParseSingleLeg = Nothing: Exit Function
    End If

    If t.strikes Is Nothing Or t.strikes.Count = 0 Then
        MsgBox "No strikes found in input." & vbNewLine & _
               "Expected at least one decimal number as a strike price.", vbCritical
        Set ParseSingleLeg = Nothing: Exit Function
    End If

    Set ParseSingleLeg = t
End Function

Public Function ParseTradeInput(inputLine As String) As Collection
    Dim parts() As String, raw As String
    Dim leg1 As TradeInput, leg2 As TradeInput
    Dim i As Long, vsIndex As Long, overrideCode As String
    Dim tradeParts As New Collection
    Dim parsedVolume As Long, parsedPremium As Double, parsedSide As String

    inputLine = Trim$(inputLine)
    Debug.Print "ParseTradeInput: Input = " & inputLine
    If inputLine = "" Then
        Set ParseTradeInput = tradeParts: Exit Function
    End If

    ' Extract contract override — skip if content is + or -
    If InStr(inputLine, "(") > 0 And InStr(inputLine, ")") > 0 Then
        Dim pOpen  As Integer: pOpen = InStr(inputLine, "(")
        Dim pClose As Integer: pClose = InStr(inputLine, ")")
        If pClose > pOpen Then
            Dim pContent As String
            pContent = Mid$(inputLine, pOpen + 1, pClose - pOpen - 1)
            If pContent <> "+" And pContent <> "-" Then
                Dim beforeParen As String
                beforeParen = Trim$(Left$(inputLine, pOpen - 1))
                Dim lastSpace As Integer: lastSpace = InStrRev(beforeParen, " ")
                Dim prevToken As String
                prevToken = Mid$(beforeParen, lastSpace + 1)
                If Not IsNumeric(prevToken) Then
                    overrideCode = UCase$(Trim$(pContent))
                    inputLine = Trim$(Replace(inputLine, "(" & pContent & ")", ""))
                End If
            End If
        End If
    End If

    inputLine = Replace(inputLine, " @ ", "@")
    inputLine = Replace(inputLine, " / ", "/")
    Do While InStr(inputLine, "  ") > 0
        inputLine = Replace(inputLine, "  ", " ")
    Loop
    inputLine = Trim$(inputLine)

    parts = Split(inputLine, " ")
    vsIndex = -1
    parsedSide = ""
    parsedVolume = 0
    parsedPremium = 0

    For i = LBound(parts) To UBound(parts)
        raw = Trim$(parts(i))
        If raw = "" Then GoTo SkipToken

        If InStr(raw, "/") > 0 Then
            Dim pq() As String: pq = Split(raw, "/")
            If UBound(pq) = 1 And IsNumeric(Trim$(pq(0))) And IsNumeric(Trim$(pq(1))) Then
                parsedSide = "B"
                parsedPremium = Round(CDbl(Trim$(pq(0))) * 0.01, 4)
                parsedVolume = CLng(Trim$(pq(1)))
            End If
        ElseIf InStr(raw, "@") > 0 Then
            Dim qp() As String: qp = Split(raw, "@")
            If UBound(qp) = 1 And IsNumeric(Trim$(qp(0))) And IsNumeric(Trim$(qp(1))) Then
                parsedSide = "S"
                parsedPremium = Round(CDbl(Trim$(qp(1))) * 0.01, 4)
                parsedVolume = CLng(Trim$(qp(0)))
            End If
        End If

        If UCase$(raw) = "VS" Then vsIndex = i
SkipToken:
    Next i

    If parsedVolume = 0 Then
        MsgBox "No volume found in trade string." & vbNewLine & _
               "Use price/qty format (e.g. 4/500) or qty@price format (e.g. 500@4).", vbCritical
        Set ParseTradeInput = tradeParts: Exit Function
    End If

    If parsedSide <> "B" And parsedSide <> "S" Then
        MsgBox "Could not determine buy/sell direction." & vbNewLine & _
               "Use price/qty for a debit (buy) or qty@price for a credit (sell).", vbCritical
        Set ParseTradeInput = tradeParts: Exit Function
    End If

    Dim cleanParts() As String
    ReDim cleanParts(0 To UBound(parts))
    Dim cleanCount As Long: cleanCount = 0
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then
            cleanParts(cleanCount) = Trim$(parts(i))
            cleanCount = cleanCount + 1
        End If
    Next i
    ReDim Preserve cleanParts(0 To cleanCount - 1)

    ' Single leg
    If vsIndex = -1 Then
        Set leg1 = ParseSingleLeg(cleanParts)
        If leg1 Is Nothing Then
            Set ParseTradeInput = tradeParts: Exit Function
        End If

        leg1.Volume = parsedVolume
        leg1.Premium = parsedPremium
        leg1.DirectionSide = parsedSide

        If overrideCode <> "" And LegContainsCode(leg1, overrideCode) Then
            Set leg1.ContractCodes = New Collection
            leg1.ContractCodes.Add overrideCode
            AddPackHelperIfShortDated leg1, overrideCode
        End If

        Dim realCodes As New Collection
        For i = 1 To leg1.ContractCodes.Count
            If Not IsPackHelperCode(leg1.ContractCodes(i)) Then
                realCodes.Add leg1.ContractCodes(i)
            End If
        Next i

        If realCodes.Count > 1 Then
            If leg1.IsStupid Then
                For i = 1 To realCodes.Count
                    Dim tDup As TradeInput
                    Set tDup = CloneTradeInput(leg1)
                    Set tDup.ContractCodes = New Collection
                    tDup.ContractCodes.Add realCodes(i)
                    AddPackHelperIfShortDated tDup, realCodes(i)
                    tDup.DirectionSide = leg1.DirectionSide
                    tradeParts.Add tDup
                Next i
                Set ParseTradeInput = tradeParts: Exit Function
            End If

            If (leg1.Strategy = "cs" Or leg1.Strategy = "ps") And leg1.strikes.Count = 1 Then
                For i = 1 To realCodes.Count
                    Dim tCal As TradeInput
                    Set tCal = CloneTradeInput(leg1)
                    tCal.Strategy = "single"
                    Set tCal.OptionTypes = New Collection
                    If leg1.Strategy = "ps" Then
                        tCal.OptionTypes.Add "P"
                        tCal.IsPutCentric = True: tCal.IsCallCentric = False
                    Else
                        tCal.OptionTypes.Add "C"
                        tCal.IsCallCentric = True: tCal.IsPutCentric = False
                    End If
                    Set tCal.ContractCodes = New Collection
                    tCal.ContractCodes.Add realCodes(i)
                    AddPackHelperIfShortDated tCal, realCodes(i)
                    tCal.DirectionSide = IIf(i = 1, parsedSide, IIf(parsedSide = "B", "S", "B"))
                    tradeParts.Add tCal
                Next i
                Set ParseTradeInput = tradeParts: Exit Function
            End If
        End If

        If Not ValidateStrikes(leg1) Then
            Set ParseTradeInput = tradeParts: Exit Function
        End If

        tradeParts.Add leg1

    ' VS two legs
    Else
        If vsIndex = 0 Then
            MsgBox "VS found at start of input with no left leg.", vbCritical
            Set ParseTradeInput = tradeParts: Exit Function
        End If
        If vsIndex = UBound(parts) Then
            MsgBox "VS found at end of input with no right leg.", vbCritical
            Set ParseTradeInput = tradeParts: Exit Function
        End If

        Dim tokens1() As String, tokens2() As String
        ReDim tokens1(0 To vsIndex - 1)
        ReDim tokens2(0 To UBound(parts) - vsIndex - 1)
        For i = 0 To vsIndex - 1: tokens1(i) = parts(i): Next i
        For i = vsIndex + 1 To UBound(parts): tokens2(i - vsIndex - 1) = parts(i): Next i

        Set leg1 = ParseSingleLeg(tokens1)
        Set leg2 = ParseSingleLeg(tokens2)

        If leg1 Is Nothing Then
            MsgBox "Could not parse left side of VS trade. Check input before VS.", vbCritical
            Set ParseTradeInput = tradeParts: Exit Function
        End If
        If leg2 Is Nothing Then
            MsgBox "Could not parse right side of VS trade. Check input after VS.", vbCritical
            Set ParseTradeInput = tradeParts: Exit Function
        End If

        leg1.Volume = parsedVolume: leg1.Premium = parsedPremium
        leg2.Volume = parsedVolume: leg2.Premium = parsedPremium
        leg1.SuppressPremium = True
        leg2.SuppressPremium = True

        Dim stupidMode As Boolean
        stupidMode = (leg1.IsStupid Or leg2.IsStupid)

        If overrideCode <> "" And LegContainsCode(leg1, overrideCode) Then
            Set leg1.ContractCodes = New Collection
            leg1.ContractCodes.Add overrideCode
            AddPackHelperIfShortDated leg1, overrideCode
            leg1.DirectionSide = parsedSide
            leg2.DirectionSide = IIf(stupidMode, parsedSide, IIf(parsedSide = "B", "S", "B"))
        ElseIf overrideCode <> "" And LegContainsCode(leg2, overrideCode) Then
            Set leg2.ContractCodes = New Collection
            leg2.ContractCodes.Add overrideCode
            AddPackHelperIfShortDated leg2, overrideCode
            leg2.DirectionSide = parsedSide
            leg1.DirectionSide = IIf(stupidMode, parsedSide, IIf(parsedSide = "B", "S", "B"))
        Else
            leg1.DirectionSide = parsedSide
            leg2.DirectionSide = IIf(stupidMode, parsedSide, IIf(parsedSide = "B", "S", "B"))
        End If

        ' Apply per-strike absolute overrides
        If leg1.StrikeOverrides.Count > 0 Then
            If CStr(leg1.StrikeOverrides(1)) = "+" Then leg1.DirectionSide = "B"
            If CStr(leg1.StrikeOverrides(1)) = "-" Then leg1.DirectionSide = "S"
        End If
        If leg2.StrikeOverrides.Count > 0 Then
            If CStr(leg2.StrikeOverrides(1)) = "+" Then leg2.DirectionSide = "B"
            If CStr(leg2.StrikeOverrides(1)) = "-" Then leg2.DirectionSide = "S"
        End If

        If Not ValidateStrikes(leg1) Then
            Set ParseTradeInput = tradeParts: Exit Function
        End If
        If Not ValidateStrikes(leg2) Then
            Set ParseTradeInput = tradeParts: Exit Function
        End If

        tradeParts.Add leg1
        tradeParts.Add leg2
    End If

    Debug.Print "ParseTradeInput: Returning " & tradeParts.Count & " legs."
    Dim dbg As Integer
    For dbg = 1 To tradeParts.Count
        Dim tdbg As TradeInput
        Set tdbg = tradeParts(dbg)
        Debug.Print "Leg " & dbg & ": DirectionSide=" & tdbg.DirectionSide & _
                    " StrikeOverrides.Count=" & tdbg.StrikeOverrides.Count
        If tdbg.StrikeOverrides.Count > 0 Then
            Debug.Print "  StrikeOverride(1)=" & CStr(tdbg.StrikeOverrides(1))
        End If
        Debug.Print "  DeltaOverride=" & tdbg.DeltaOverride
    Next dbg
    Set ParseTradeInput = tradeParts
End Function

