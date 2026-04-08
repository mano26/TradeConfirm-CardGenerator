Attribute VB_Name = "CardGenerator_v4"
Option Explicit

Public Function GenerateCardsFile() As String
    GenerateCardsFile = ""
    
    Dim ws1 As Worksheet: Set ws1 = ThisWorkbook.Sheets(SH1_NAME)
    Dim ws2 As Worksheet: Set ws2 = ThisWorkbook.Sheets(SH2_NAME)
    
    ' Collect legs from Sheet 1
    Dim sides(1 To 50) As String, vols(1 To 50) As Double
    Dim moCards(1 To 50) As String, strikes(1 To 50) As String
    Dim optTypes(1 To 50) As String, prices(1 To 50) As String
    Dim tickets(1 To 50) As String, legRows(1 To 50) As Long
    Dim legCount As Integer: legCount = 0
    Dim r As Long: r = S1_CONF_START
    Dim blankRun As Integer: blankRun = 0
    
    Do While r <= 200
        If ws1.Cells(r, S1_COL_VOL).Value <> "" Then
            blankRun = 0
            legCount = legCount + 1
            legRows(legCount) = r
            sides(legCount) = CStr(ws1.Cells(r, S1_COL_SIDE).Value)
            vols(legCount) = CDbl(ws1.Cells(r, S1_COL_VOL).Value)
            optTypes(legCount) = CStr(ws1.Cells(r, S1_COL_OPTTYPE).Value)
            prices(legCount) = CStr(ws1.Cells(r, S1_COL_PRICE).Value)
            tickets(legCount) = Trim$(CStr(ws1.Cells(r, S1_COL_TICKET).Value))
            
            Dim moVal As String
            moVal = Trim$(CStr(ws1.Cells(r, S1_COL_MO_CARD).Value))
            If moVal = "" Then
                MsgBox "MO code missing row " & r, vbExclamation
                Exit Function
            End If
            moCards(legCount) = moVal
            
            If ws1.Cells(r, S1_COL_STRIKE).Value = "" Then
                strikes(legCount) = ""
            Else
                Dim sd As Double: sd = CDbl(ws1.Cells(r, S1_COL_STRIKE).Value)
                Dim ss As String: ss = CStr(sd)
                If InStr(ss, ".") = 0 Then
                    strikes(legCount) = ss & ".00"
                ElseIf Len(ss) - InStr(ss, ".") < 2 Then
                    strikes(legCount) = ss & "0"
                Else
                    strikes(legCount) = ss
                End If
            End If
        Else
            blankRun = blankRun + 1
            If blankRun >= 2 Then Exit Do
        End If
        r = r + 1
    Loop
    
    If legCount = 0 Then
        MsgBox "No legs found.", vbExclamation
        Exit Function
    End If
    
    ' Delta ratio
    Dim totalOptVol As Double: totalOptVol = 0
    Dim totalFutVol As Double: totalFutVol = 0
    Dim k As Integer
    For k = 1 To legCount
        If optTypes(k) = "" And Trim$(strikes(k)) = "" Then
            totalFutVol = vols(k)
        Else
            If totalOptVol = 0 Then totalOptVol = vols(k)
        End If
    Next k
    If totalOptVol = 0 Then totalOptVol = 1
    Dim deltaRatio As Double: deltaRatio = totalFutVol / totalOptVol
    
    ' Read counterparties from Sheet 2
    Dim cpQty(1 To 20) As Double, cpSym(1 To 20) As String
    Dim cpBkt(1 To 20) As String, cpBrkr(1 To 20) As String
    Dim cpCount As Integer: cpCount = 0
    Dim i As Integer
    
    For i = 0 To (S2_CP_DATA_END - S2_CP_DATA_START)
        Dim rn As Long: rn = S2_CP_DATA_START + i
        Dim sym As String: sym = Trim$(CStr(ws2.Cells(rn, S2_CP_COL_SYMBOL).Value))
        If sym <> "" Then
            cpCount = cpCount + 1
            If ws2.Cells(rn, S2_CP_COL_QTY).Value = "" Then
                cpQty(cpCount) = 0
            Else
                cpQty(cpCount) = CDbl(ws2.Cells(rn, S2_CP_COL_QTY).Value)
            End If
            cpSym(cpCount) = sym
            cpBkt(cpCount) = Trim$(UCase$(CStr(ws2.Cells(rn, S2_CP_COL_BRACKET).Value)))
            cpBrkr(cpCount) = Trim$(UCase$(CStr(ws2.Cells(rn, S2_CP_COL_BROKER).Value)))
        End If
    Next i
    
    Dim tradeDate As String: tradeDate = Format$(Now(), "MM/DD/YY")
    Dim isMultiLeg As Boolean: isMultiLeg = (legCount > 1)
    
    ' Unique bracket+broker combinations
    Dim grpBkt(1 To 40) As String, grpBrk(1 To 40) As String
    Dim grpCount As Integer: grpCount = 0
    Dim b As Integer, bFound As Boolean
    
    For i = 1 To cpCount
        bFound = False
        For b = 1 To grpCount
            If grpBkt(b) = cpBkt(i) And grpBrk(b) = cpBrkr(i) Then
                bFound = True
                Exit For
            End If
        Next b
        If Not bFound And cpBkt(i) <> "" And cpBrkr(i) <> "" Then
            grpCount = grpCount + 1
            grpBkt(grpCount) = cpBkt(i)
            grpBrk(grpCount) = cpBrkr(i)
        End If
    Next i
    
    If grpCount = 0 Then
        MsgBox "No bracket/broker combinations found.", vbExclamation
        Exit Function
    End If
    
    ' Build HTML
    Dim html As String: html = BuildHTMLHeader(tradeDate)
    
    For b = 1 To grpCount
        Dim thisBkt As String: thisBkt = grpBkt(b)
        Dim thisBrk As String: thisBrk = grpBrk(b)
        Dim printBkt As String
        If isMultiLeg Then
            printBkt = thisBkt & "6"
        Else
            printBkt = thisBkt
        End If
        Dim bQty(1 To 20) As Double, bSym(1 To 20) As String, bBrkr(1 To 20) As String
        Dim bCnt As Integer: bCnt = 0
        For i = 1 To cpCount
            If cpBkt(i) = thisBkt And cpBrkr(i) = thisBrk Then
                bCnt = bCnt + 1
                bQty(bCnt) = cpQty(i)
                bSym(bCnt) = cpSym(i)
                bBrkr(bCnt) = cpBrkr(i)
            End If
        Next i
        Dim pages As Integer: pages = Int((bCnt - 1) / 5) + 1
        Dim pg As Integer
        For pg = 1 To pages
            Dim cpFrom As Integer: cpFrom = (pg - 1) * 5 + 1
            Dim cpTo As Integer
            If pg * 5 <= bCnt Then
                cpTo = pg * 5
            Else
                cpTo = bCnt
            End If
            For k = 1 To legCount
                html = html & BuildCardHTML(sides(k), vols(k), moCards(k), _
                       strikes(k), optTypes(k), prices(k), _
                       bQty, bSym, bBrkr, cpFrom, cpTo, _
                       printBkt, tradeDate, deltaRatio)
            Next k
        Next pg
    Next b
    
    html = html & "</div></body></html>"
    
    ' Save to permanent dated folder
    Dim filePath As String
    filePath = GetOutputFolder() & "\GFI_Cards_" & Format$(Now(), "YYYYMMDD_HHMMSS") & ".html"
    
    Dim fNum As Integer: fNum = FreeFile
    On Error GoTo CardFileErr
    Open filePath For Output As #fNum
    Print #fNum, html
    Close #fNum
    On Error GoTo 0
    
    Shell "cmd /c start """" """ & filePath & """", vbHide
    
    GenerateCardsFile = filePath
    Exit Function
    
CardFileErr:
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
    MsgBox "Error saving cards file: " & filePath & vbNewLine & Err.Description, vbCritical
End Function

Private Function ValidatePrices(ws As Worksheet, legCount As Integer, _
        legRows() As Long, optTypes() As String, _
        strikes() As String, prices() As String) As Boolean
    ValidatePrices = True
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
            MsgBox "Price reconciliation failed for package " & Format$(pps, "0.0000"), vbCritical
            ValidatePrices = False
            Exit Function
        End If
NS:
    Next seg
End Function

Private Function BuildHTMLHeader(tradeDate As String) As String
    Dim s As String
    s = "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    s = s & "<title>GFI Trading Cards " & tradeDate & "</title>" & vbNewLine
    s = s & "<style>" & vbNewLine
    s = s & "* { box-sizing:border-box; margin:0; padding:0; }" & vbNewLine
    s = s & "body { font-family:Arial,Helvetica,sans-serif; background:#e0e0e0; padding:0.3in; }" & vbNewLine
    s = s & ".cards-wrap { display:flex; flex-wrap:wrap; gap:0.15in; justify-content:flex-start; }" & vbNewLine
    s = s & ".card { width:3.5in; height:5.5in; border-radius:10px; overflow:hidden; border:1.5px solid; "
    s = s & "page-break-inside:avoid; display:flex; flex-direction:column; }" & vbNewLine
    s = s & ".card-header { padding:6px 10px 0 10px; flex-shrink:0; }" & vbNewLine
    s = s & ".card-top-row { display:flex; justify-content:space-between; align-items:baseline; }" & vbNewLine
    s = s & ".card-type { font-size:19px; font-weight:900; letter-spacing:1px; }" & vbNewLine
    s = s & ".card-broker { font-size:19px; font-weight:900; letter-spacing:2px; text-align:center; flex:1; }" & vbNewLine
    s = s & ".card-role { font-size:12px; font-weight:700; margin-top:2px; padding-bottom:4px; }" & vbNewLine
    s = s & ".card-rule { border:none; border-top:1px solid; margin:0; flex-shrink:0; }" & vbNewLine
    s = s & ".col-headers { display:flex; flex-shrink:0; border-bottom:1.5px solid; }" & vbNewLine
    s = s & ".col-headers div { font-size:11px; font-weight:700; text-align:center; padding:3px 1px; }" & vbNewLine
    s = s & ".slots { flex:1; display:flex; flex-direction:column; min-height:0; }" & vbNewLine
    s = s & ".slot { flex:1; display:flex; border-bottom:0.5px solid; min-height:0; }" & vbNewLine
    s = s & ".slot:last-child { border-bottom:none; }" & vbNewLine
    s = s & ".cell { display:flex; align-items:center; justify-content:center; font-size:14px; "
    s = s & "border-right:0.5px solid; overflow:hidden; }" & vbNewLine
    s = s & ".cell:last-child { border-right:none; }" & vbNewLine
    s = s & ".cp-cell { display:flex; flex-direction:column; border-right:0.5px solid; overflow:hidden; }" & vbNewLine
    s = s & ".cp-top { flex:1; display:flex; align-items:center; justify-content:center; font-size:14px; "
    s = s & "font-weight:700; color:#007700; border-bottom:0.5px solid; overflow:hidden; }" & vbNewLine
    s = s & ".cp-bot { flex:1; display:flex; align-items:center; justify-content:center; "
    s = s & "font-size:14px; color:#005500; overflow:hidden; }" & vbNewLine
    s = s & ".w-qty { width:13%; } .w-mo { width:16%; } .w-str { width:16%; }" & vbNewLine
    s = s & ".w-pr { width:13%; } .w-cp { width:32%; } .w-bkt { width:10%; }" & vbNewLine
    s = s & ".card-footer { font-size:7px; text-align:center; padding:4px; border-top:1px solid; flex-shrink:0; }" & vbNewLine
    s = s & "@media print { body { background:white; padding:0; margin:0; }" & vbNewLine
    s = s & "@page { size:letter portrait; margin:0.35in; }" & vbNewLine
    s = s & ".cards-wrap { gap:0.15in; }" & vbNewLine
    s = s & ".card { width:3.5in; height:5.5in; border:1.5px solid !important; "
    s = s & "-webkit-print-color-adjust:exact; print-color-adjust:exact; } }" & vbNewLine
    s = s & "</style></head><body><div class='cards-wrap'>" & vbNewLine
    BuildHTMLHeader = s
End Function

Private Function BuildCardHTML(side As String, vol As Double, _
        moCode As String, strike As String, optType As String, price As String, _
        bQty() As Double, bSym() As String, bBroker() As String, _
        cpFrom As Integer, cpTo As Integer, _
        bracket As String, tradeDate As String, deltaRatio As Double) As String
    
    Dim isFut As Boolean
    isFut = (optType = "" And Trim$(strike) = "")
    
    Dim cardType As String, cardRole As String, cpRole As String
    Dim bgColor As String, ink As String
    
    If isFut Then
        cardType = "FUTURES"
        If side = "B" Then
            cardRole = "BUYER"
            cpRole = "SELLER"
        Else
            cardRole = "SELLER"
            cpRole = "BUYER"
        End If
        bgColor = "#fefce8"
    ElseIf UCase$(optType) = "C" Then
        cardType = "CALL"
        If side = "S" Then
            cardRole = "SELLER"
            cpRole = "BUYER"
        Else
            cardRole = "BUYER"
            cpRole = "SELLER"
        End If
        bgColor = "#ffffff"
    Else
        cardType = "PUT"
        If side = "S" Then
            cardRole = "SELLER"
            cpRole = "BUYER"
        Else
            cardRole = "BUYER"
            cpRole = "SELLER"
        End If
        bgColor = "#f5f0c8"
    End If
    
    If cardRole = "BUYER" Then
        ink = "#1f4e79"
    Else
        ink = "#cc2222"
    End If
    
    Dim brokerName As String: brokerName = ""
    If cpFrom >= 1 And cpFrom <= UBound(bBroker) Then brokerName = bBroker(cpFrom)
    
    Dim qLbl As String, sLbl As String, pLbl As String, bLbl As String
    If isFut Then
        qLbl = "CARS"
        sLbl = ""
        pLbl = "PRICE"
        bLbl = "BK"
    Else
        qLbl = "QTY."
        sLbl = "STRIKE"
        pLbl = "PREM."
        bLbl = "BKT."
    End If
    
    Dim h As String
    h = "<div class='card' style='background:" & bgColor & ";border-color:" & ink & ";'>" & vbNewLine
    h = h & "<div class='card-header'><div class='card-top-row'>"
    h = h & "<div class='card-type' style='color:" & ink & "'>" & cardType & "</div>"
    h = h & "<div class='card-broker' style='color:" & ink & "'>" & brokerName & "</div></div>"
    h = h & "<div class='card-role' style='color:" & ink & "'>" & cardRole & "</div></div>"
    h = h & "<hr class='card-rule' style='border-color:" & ink & "'>"
    
    h = h & "<div class='col-headers' style='border-color:" & ink & ";color:" & ink & "'>"
    h = h & "<div class='w-qty' style='border-right:0.5px solid " & ink & "'>" & qLbl & "</div>"
    h = h & "<div class='w-mo' style='border-right:0.5px solid " & ink & "'>MO.</div>"
    h = h & "<div class='w-str' style='border-right:0.5px solid " & ink & "'>" & sLbl & "</div>"
    h = h & "<div class='w-pr' style='border-right:0.5px solid " & ink & "'>" & pLbl & "</div>"
    h = h & "<div class='w-cp' style='border-right:0.5px solid " & ink & "'>" & cpRole & "</div>"
    h = h & "<div class='w-bkt'>" & bLbl & "</div></div>"
    
    h = h & "<div class='slots'>" & vbNewLine
    Dim slot As Integer
    For slot = 1 To 5
        Dim cpIdx As Integer: cpIdx = cpFrom + slot - 1
        h = h & "<div class='slot' style='border-color:" & ink & "'>"
        If cpIdx <= cpTo Then
            Dim dq As Long
            If isFut Then
                dq = CLng(Round(bQty(cpIdx) * deltaRatio, 0))
            Else
                dq = CLng(bQty(cpIdx))
            End If
            h = h & "<div class='cell w-qty' style='color:" & ink & ";border-color:" & ink & "'>" & CStr(dq) & "</div>"
            h = h & "<div class='cell w-mo' style='color:" & ink & ";border-color:" & ink & "'>" & UCase$(moCode) & "</div>"
            If isFut Then
                h = h & "<div class='cell w-str' style='border-color:" & ink & "'>&nbsp;</div>"
            Else
                h = h & "<div class='cell w-str' style='color:" & ink & ";border-color:" & ink & "'>" & strike & "</div>"
            End If
            h = h & "<div class='cell w-pr' style='color:" & ink & ";border-color:" & ink & "'>" & price & "</div>"
            
            Dim rawS As String: rawS = bSym(cpIdx)
            Dim slashP As Integer: slashP = InStr(rawS, "/")
            Dim sT As String, sB As String
            If slashP > 0 Then
                sT = Trim$(Left$(rawS, slashP - 1))
                sB = Trim$(Mid$(rawS, slashP + 1))
            Else
                sT = Trim$(rawS)
                sB = "&nbsp;"
            End If
            h = h & "<div class='cp-cell w-cp' style='border-color:" & ink & "'>"
            h = h & "<div class='cp-top' style='border-color:" & ink & "'>" & sT & "</div>"
            h = h & "<div class='cp-bot'>" & sB & "</div></div>"
            h = h & "<div class='cell w-bkt' style='color:" & ink & ";border-right:none'>" & bracket & "</div>"
        Else
            h = h & "<div class='cell w-qty' style='border-color:" & ink & "'>&nbsp;</div>"
            h = h & "<div class='cell w-mo' style='border-color:" & ink & "'>&nbsp;</div>"
            h = h & "<div class='cell w-str' style='border-color:" & ink & "'>&nbsp;</div>"
            h = h & "<div class='cell w-pr' style='border-color:" & ink & "'>&nbsp;</div>"
            h = h & "<div class='cp-cell w-cp' style='border-color:" & ink & "'>"
            h = h & "<div class='cp-top' style='border-color:" & ink & "'>&nbsp;</div>"
            h = h & "<div class='cp-bot'>&nbsp;</div></div>"
            h = h & "<div class='cell w-bkt' style='border-right:none'>&nbsp;</div>"
        End If
        h = h & "</div>" & vbNewLine
    Next slot
    h = h & "</div>"
    h = h & "<div class='card-footer' style='color:" & ink & ";border-color:" & ink & "'>"
    h = h & "</div>"
    h = h & "</div>" & vbNewLine
    BuildCardHTML = h
End Function
