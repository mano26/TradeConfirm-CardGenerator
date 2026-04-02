Attribute VB_Name = "TradeValidation_v4"

Option Explicit

Public Function ValidateBeforeGenerate() As Boolean
    ValidateBeforeGenerate = False
    
    Dim ws1 As Worksheet: Set ws1 = ThisWorkbook.Sheets(SH1_NAME)
    Dim ws2 As Worksheet: Set ws2 = ThisWorkbook.Sheets(SH2_NAME)
    
    ' Check trade legs exist on Sheet 1
    Dim baseOptVol As Double: baseOptVol = 0
    Dim legCount As Integer: legCount = 0
    Dim legRows(1 To 50) As Long
    Dim optTypes(1 To 50) As String
    Dim strikes(1 To 50) As String
    Dim prices(1 To 50) As String
    Dim r As Long: r = S1_CONF_START
    Dim blankRun As Integer: blankRun = 0
    
    Do While r <= 200
        If ws1.Cells(r, S1_COL_VOL).Value <> "" Then
            blankRun = 0
            legCount = legCount + 1
            legRows(legCount) = r
            optTypes(legCount) = Trim$(CStr(ws1.Cells(r, S1_COL_OPTTYPE).Value))
            strikes(legCount) = Trim$(CStr(ws1.Cells(r, S1_COL_STRIKE).Value))
            prices(legCount) = Trim$(CStr(ws1.Cells(r, S1_COL_PRICE).Value))
            Dim isFut As Boolean
            isFut = (optTypes(legCount) = "" And strikes(legCount) = "")
            If Not isFut And baseOptVol = 0 Then
                baseOptVol = CDbl(ws1.Cells(r, S1_COL_VOL).Value)
            End If
        Else
            blankRun = blankRun + 1
            If blankRun >= 2 Then Exit Do
        End If
        r = r + 1
    Loop
    
    If legCount = 0 Then
        MsgBox "No trade legs found on " & SH1_NAME & "." & vbNewLine & _
               "Please process a trade first.", vbExclamation
        Exit Function
    End If
    
    ' Check House / Account on Sheet 2
    Dim house As String
    house = Trim$(CStr(ws2.Cells(S2_HEADER_ROW, S2_HOUSE_COL).Value))
    Dim account As String
    account = Trim$(CStr(ws2.Cells(S2_HEADER_ROW, S2_ACCOUNT_COL).Value))
    
    Dim missingHeader As String: missingHeader = ""
    If house = "" Then
        missingHeader = missingHeader & "  - House (A7)" & vbNewLine
    End If
    If account = "" Then
        missingHeader = missingHeader & "  - Account (B7)" & vbNewLine
    End If
    
    If Len(missingHeader) > 0 Then
        MsgBox "Please fill the following on '" & SH2_NAME & "':" & vbNewLine & _
               vbNewLine & missingHeader, vbExclamation
        Exit Function
    End If
    
    ' Check counterparty rows on Sheet 2
    Dim cpQty(1 To 20) As Double
    Dim cpCount As Integer: cpCount = 0
    Dim missingFields As String: missingFields = ""
    Dim hasAnyCP As Boolean: hasAnyCP = False
    Dim i As Integer
    
    For i = 0 To (S2_CP_DATA_END - S2_CP_DATA_START)
        Dim rn As Long: rn = S2_CP_DATA_START + i
        
        Dim hasQ As Boolean
        hasQ = (Trim$(CStr(ws2.Cells(rn, S2_CP_COL_QTY).Value)) <> "")
        Dim hasBrk As Boolean
        hasBrk = (Trim$(CStr(ws2.Cells(rn, S2_CP_COL_BROKER).Value)) <> "")
        Dim hasS As Boolean
        hasS = (Trim$(CStr(ws2.Cells(rn, S2_CP_COL_SYMBOL).Value)) <> "")
        Dim hasB As Boolean
        hasB = (Trim$(CStr(ws2.Cells(rn, S2_CP_COL_BRACKET).Value)) <> "")
        
        If hasQ Or hasBrk Or hasS Or hasB Then
            hasAnyCP = True
            Dim rowMiss As String: rowMiss = ""
            
            If Not hasQ Then
                rowMiss = rowMiss & "Qty, "
                ws2.Cells(rn, S2_CP_COL_QTY).Interior.Color = RGB(255, 235, 0)
            End If
            If Not hasBrk Then
                rowMiss = rowMiss & "Broker, "
                ws2.Cells(rn, S2_CP_COL_BROKER).Interior.Color = RGB(255, 235, 0)
            End If
            If Not hasS Then
                rowMiss = rowMiss & "Symbol, "
                ws2.Cells(rn, S2_CP_COL_SYMBOL).Interior.Color = RGB(255, 235, 0)
            End If
            If Not hasB Then
                rowMiss = rowMiss & "Bracket, "
                ws2.Cells(rn, S2_CP_COL_BRACKET).Interior.Color = RGB(255, 235, 0)
            End If
            
            If Len(rowMiss) > 0 Then
                rowMiss = Left$(rowMiss, Len(rowMiss) - 2)
                missingFields = missingFields & "  Row " & rn & ": " & rowMiss & vbNewLine
            Else
                ws2.Cells(rn, S2_CP_COL_QTY).Interior.ColorIndex = xlNone
                ws2.Cells(rn, S2_CP_COL_BROKER).Interior.ColorIndex = xlNone
                ws2.Cells(rn, S2_CP_COL_SYMBOL).Interior.ColorIndex = xlNone
                ws2.Cells(rn, S2_CP_COL_BRACKET).Interior.ColorIndex = xlNone
                cpCount = cpCount + 1
                cpQty(cpCount) = CDbl(ws2.Cells(rn, S2_CP_COL_QTY).Value)
            End If
        End If
    Next i
    
    If Not hasAnyCP Then
        MsgBox "No counterparties entered on '" & SH2_NAME & "'." & vbNewLine & _
               vbNewLine & "Please fill at least one row (Qty, Broker, Symbol, Bracket).", vbExclamation
        Exit Function
    End If
    
    If Len(missingFields) > 0 Then
        MsgBox "Incomplete counterparty rows:" & vbNewLine & vbNewLine & _
               missingFields & vbNewLine & _
               "Please fill highlighted cells.", vbExclamation
        Exit Function
    End If
    
    ' Validate qty splits
    Dim cpTotal As Double: cpTotal = 0
    For i = 1 To cpCount
        cpTotal = cpTotal + cpQty(i)
    Next i
    
    If baseOptVol = 0 Then
        r = S1_CONF_START
        blankRun = 0
        Do While r <= 200
            If ws1.Cells(r, S1_COL_VOL).Value <> "" Then
                baseOptVol = CDbl(ws1.Cells(r, S1_COL_VOL).Value)
                Exit Do
            End If
            r = r + 1
        Loop
    End If
    
    If baseOptVol > 0 And Abs(cpTotal - baseOptVol) > 0.01 Then
        MsgBox "Counterparty qty split does not match trade size:" & vbNewLine & _
               vbNewLine & _
               "  Base leg volume: " & Format$(baseOptVol, "#,##0") & vbNewLine & _
               "  CP total: " & Format$(cpTotal, "#,##0") & vbNewLine & _
               "  Difference: " & Format$(Abs(cpTotal - baseOptVol), "#,##0") & vbNewLine & _
               vbNewLine & "Please correct quantities.", vbExclamation
        Exit Function
    End If
    
    ' Check all legs have prices filled
    Dim missingPrices As String: missingPrices = ""
    Dim k As Integer
    For k = 1 To legCount
        If Trim$(prices(k)) = "" Then
            ws1.Cells(legRows(k), S1_COL_PRICE).Interior.Color = RGB(255, 235, 0)
            missingPrices = missingPrices & "  Row " & legRows(k) & vbNewLine
        Else
            ws1.Cells(legRows(k), S1_COL_PRICE).Interior.ColorIndex = xlNone
        End If
    Next k
    
    If Len(missingPrices) > 0 Then
        MsgBox "Missing prices (col J):" & vbNewLine & vbNewLine & _
               missingPrices & vbNewLine & _
               "Please fill all leg prices before generating.", vbExclamation
        Exit Function
    End If
    
    ' Price reconciliation
    If Not ValidatePriceReconciliation(ws1, legCount, legRows, optTypes, strikes, prices) Then
        Exit Function
    End If
    
    ValidateBeforeGenerate = True
End Function

Private Function ValidatePriceReconciliation(ws As Worksheet, legCount As Integer, _
        legRows() As Long, optTypes() As String, _
        strikes() As String, prices() As String) As Boolean
    ValidatePriceReconciliation = True
    
    ' Collect unique package premiums
    Dim segPrems(1 To 50) As Double
    Dim segCount As Integer: segCount = 0
    Dim k As Integer, s As Integer, found As Boolean
    
    For k = 1 To legCount
        If optTypes(k) = "" And Trim$(strikes(k)) = "" Then GoTo NL1
        Dim pp As Double: pp = 0
        On Error Resume Next
        pp = CDbl(ws.Cells(legRows(k), S1_COL_PKG_PREM).Value)
        On Error GoTo 0
        found = False
        For s = 1 To segCount
            If segPrems(s) = pp Then
                found = True
                Exit For
            End If
        Next s
        If Not found Then
            segCount = segCount + 1
            segPrems(segCount) = pp
        End If
NL1:
    Next k
    
    If segCount = 0 Then Exit Function
    
    Dim seg As Integer
    For seg = 1 To segCount
        Dim pps As Double: pps = segPrems(seg)
        Dim ss(1 To 50) As String, sV(1 To 50) As Double, sp(1 To 50) As Double
        Dim slc As Integer: slc = 0
        Dim allFilled As Boolean: allFilled = True
        
        For k = 1 To legCount
            If optTypes(k) = "" And Trim$(strikes(k)) = "" Then GoTo NL2
            Dim lp As Double: lp = 0
            On Error Resume Next
            lp = CDbl(ws.Cells(legRows(k), S1_COL_PKG_PREM).Value)
            On Error GoTo 0
            If lp = pps Then
                If Trim$(prices(k)) = "" Then
                    allFilled = False
                Else
                    slc = slc + 1
                    ss(slc) = ws.Cells(legRows(k), S1_COL_SIDE).Value
                    sV(slc) = CDbl(ws.Cells(legRows(k), S1_COL_VOL).Value)
                    sp(slc) = CDbl(prices(k))
                End If
            End If
NL2:
        Next k
        
        If Not allFilled Or slc = 0 Then GoTo NS
        
        Dim bv As Double: bv = sV(1)
        Dim j As Integer
        For j = 2 To slc
            If sV(j) < bv Then bv = sV(j)
        Next j
        If bv = 0 Then GoTo NS
        
        Dim np As Double: np = 0
        Dim legSign As Double
        For j = 1 To slc
            If ss(j) = "S" Then
                legSign = 1
            Else
                legSign = -1
            End If
            np = np + legSign * (sV(j) / bv) * sp(j)
        Next j
        
        If Abs(Abs(np) - pps) > 0.000001 Then
            MsgBox "Price reconciliation failed for package " & Format$(pps, "0.0000") & "." & _
                   vbNewLine & vbNewLine & _
                   "Expected net: " & Format$(pps, "0.0000") & vbNewLine & _
                   "Calculated net: " & Format$(Abs(np), "0.0000") & vbNewLine & _
                   "Discrepancy: " & Format$(Abs(Abs(np) - pps), "0.0000") & vbNewLine & vbNewLine & _
                   "Please check your leg prices in column J.", vbCritical
            ValidatePriceReconciliation = False
            Exit Function
        End If
NS:
    Next seg
End Function

