Attribute VB_Name = "CardGenerator"
Option Explicit

Private Const CP_HDR_ROW      As Long = 12
Private Const CP_DATA_START   As Long = 13
Private Const CP_DATA_END     As Long = 32
Private Const COL_QTY         As Long = 4
Private Const COL_SYMBOL      As Long = 5
Private Const COL_EXEC_BROKER As Long = 6
Private Const COL_BRACKET     As Long = 7
Private Const COL_NOTES       As Long = 8
Private Const COL_TICKET      As Long = 19
Private Const COL_MO_CARD     As Long = 20
Private Const COL_PKG_PREM    As Long = 21

Public Sub GenerateCards()
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets("GFI Upload Template")

    Dim sides(1 To 50)    As String
    Dim vols(1 To 50)     As Double
    Dim moCards(1 To 50)  As String
    Dim strikes(1 To 50)  As String
    Dim optTypes(1 To 50) As String
    Dim prices(1 To 50)   As String
    Dim tickets(1 To 50)  As String
    Dim legRows(1 To 50)  As Long
    Dim legCount As Integer: legCount = 0

    Dim r As Long: r = 5
    Dim blankRun As Integer: blankRun = 0
    Do While r <= 200
        If ws1.Cells(r, 4).Value <> "" Then
            blankRun = 0
            legCount = legCount + 1
            legRows(legCount) = r
            sides(legCount) = CStr(ws1.Cells(r, 3).Value)
            vols(legCount) = CDbl(ws1.Cells(r, 4).Value)
            optTypes(legCount) = CStr(ws1.Cells(r, 9).Value)
            prices(legCount) = CStr(ws1.Cells(r, 10).Value)
            tickets(legCount) = Trim$(CStr(ws1.Cells(r, COL_TICKET).Value))

            Dim moVal As String: moVal = Trim$(CStr(ws1.Cells(r, COL_MO_CARD).Value))
            If moVal = "" Then
                MsgBox "MO code missing for row " & r & "." & vbNewLine & _
                       "Please re-process the trade and try again.", vbExclamation
                Exit Sub
            End If
            moCards(legCount) = moVal

            If ws1.Cells(r, 8).Value = "" Then
                strikes(legCount) = ""
            Else
                Dim strikeDbl As Double: strikeDbl = CDbl(ws1.Cells(r, 8).Value)
                Dim strikeStr As String: strikeStr = CStr(strikeDbl)
                If InStr(strikeStr, ".") = 0 Then
                    strikes(legCount) = strikeStr & ".00"
                ElseIf Len(strikeStr) - InStr(strikeStr, ".") < 2 Then
                    strikes(legCount) = strikeStr & "0"
                Else
                    strikes(legCount) = strikeStr
                End If
            End If
        Else
            blankRun = blankRun + 1
            If blankRun >= 2 Then Exit Do
        End If
        r = r + 1
    Loop

    If legCount = 0 Then
        MsgBox "No trade legs found. Please process a trade first.", vbExclamation
        Exit Sub
    End If

    ' Ticket check
    Dim missingTicket As String: missingTicket = ""
    Dim k As Integer
    For k = 1 To legCount
        If tickets(k) = "" Then
            ws1.Cells(legRows(k), COL_TICKET).Interior.Color = RGB(255, 235, 0)
            missingTicket = missingTicket & "  Row " & legRows(k) & vbNewLine
        Else
            ws1.Cells(legRows(k), COL_TICKET).Interior.ColorIndex = xlNone
        End If
    Next k
    If Len(missingTicket) > 0 Then
        MsgBox "Missing ticket number (col S) for:" & vbNewLine & vbNewLine & _
               missingTicket & vbNewLine & "Please fill ticket numbers and try again.", vbExclamation
        Exit Sub
    End If

    ' Futures price check
    ws1.Range("J5:J1000").Interior.ColorIndex = xlNone
    Dim missingPrice As String: missingPrice = ""
    For k = 1 To legCount
        Dim isFutLeg As Boolean
        isFutLeg = (optTypes(k) = "" And Trim$(strikes(k)) = "")
        If isFutLeg And Trim$(prices(k)) = "" Then
            ws1.Cells(legRows(k), 10).Interior.Color = RGB(255, 235, 0)
            missingPrice = missingPrice & "  Row " & legRows(k) & " - Futures " & moCards(k) & vbNewLine
        End If
    Next k
    If Len(missingPrice) > 0 Then
        MsgBox "Missing futures prices (col J) - please fill:" & vbNewLine & vbNewLine & _
               missingPrice & vbNewLine & "Cards not generated.", vbExclamation
        Exit Sub
    End If

    ' Price reconciliation check
    If Not ValidatePrices(ws1, legCount, legRows, optTypes, strikes, prices) Then
        Exit Sub
    End If

    ' Delta ratio for futures qty per counterparty
    Dim totalOptVol As Double: totalOptVol = 0
    Dim totalFutVol As Double: totalFutVol = 0
    For k = 1 To legCount
        If optTypes(k) = "" And Trim$(strikes(k)) = "" Then
            totalFutVol = vols(k)
        Else
            If totalOptVol = 0 Then totalOptVol = vols(k)
        End If
    Next k
    If totalOptVol = 0 Then totalOptVol = 1
    Dim deltaRatio As Double: deltaRatio = totalFutVol / totalOptVol

    ' Read counterparty table
    Dim cpQty(1 To 20)    As Double
    Dim cpSym(1 To 20)    As String
    Dim cpBroker(1 To 20) As String
    Dim cpBkt(1 To 20)    As String
    Dim cpCount As Integer: cpCount = 0
    Dim i As Integer

    For i = 1 To 20
        Dim rn As Long: rn = CP_DATA_START + i - 1
        Dim sym As String: sym = Trim$(CStr(ws1.Cells(rn, COL_SYMBOL).Value))
        If sym <> "" Then
            cpCount = cpCount + 1
            cpQty(cpCount) = IIf(ws1.Cells(rn, COL_QTY).Value = "", 0, CDbl(ws1.Cells(rn, COL_QTY).Value))
            cpSym(cpCount) = sym
            Dim brkr As String: brkr = UCase$(Trim$(CStr(ws1.Cells(rn, COL_EXEC_BROKER).Value)))
            If brkr = "" Then
                MsgBox "No executing broker entered for counterparty '" & sym & "' (row " & rn & ")." & vbNewLine & _
                       "Please fill column F before generating cards.", vbExclamation
                Exit Sub
            End If
            cpBroker(cpCount) = brkr
            cpBkt(cpCount) = Trim$(UCase$(CStr(ws1.Cells(rn, COL_BRACKET).Value)))
        End If
    Next i

    If cpCount = 0 Then
        MsgBox "Please enter at least one counterparty.", vbExclamation
        Exit Sub
    End If

    Dim tradeDate As String
    Dim dtVal As Variant: dtVal = ws1.Cells(CP_HDR_ROW, 3).Value
    tradeDate = IIf(IsDate(dtVal), Format$(CDate(dtVal), "MM/DD/YY"), Format$(Now(), "MM/DD/YY"))

    Dim isMultiLeg As Boolean: isMultiLeg = (legCount > 1)

    Dim bktList(1 To 20) As String
    Dim bktCount As Integer: bktCount = 0
    Dim b As Integer, bFound As Boolean

    For i = 1 To cpCount
        bFound = False
        For b = 1 To bktCount
            If bktList(b) = cpBkt(i) Then bFound = True: Exit For
        Next b
        If Not bFound And cpBkt(i) <> "" Then
            bktCount = bktCount + 1
            bktList(bktCount) = cpBkt(i)
        End If
    Next i

    If bktCount = 0 Then
        MsgBox "Please enter a bracket for at least one counterparty.", vbExclamation
        Exit Sub
    End If

    Dim html As String
    html = BuildHTMLHeader(tradeDate)

    For b = 1 To bktCount
        Dim thisBkt As String: thisBkt = bktList(b)
        Dim printBkt As String: printBkt = thisBkt & IIf(isMultiLeg, "6", "")

        Dim bQty(1 To 20)    As Double
        Dim bSym(1 To 20)    As String
        Dim bBroker(1 To 20) As String
        Dim bCount As Integer: bCount = 0

        For i = 1 To cpCount
            If cpBkt(i) = thisBkt Then
                bCount = bCount + 1
                bQty(bCount) = cpQty(i)
                bSym(bCount) = cpSym(i)
                bBroker(bCount) = cpBroker(i)
            End If
        Next i

        Dim pages As Integer: pages = Int((bCount - 1) / 5) + 1
        Dim pg As Integer

        For pg = 1 To pages
            Dim cpFrom As Integer: cpFrom = (pg - 1) * 5 + 1
            Dim cpTo   As Integer: cpTo = IIf(pg * 5 <= bCount, pg * 5, bCount)
            For k = 1 To legCount
                html = html & BuildCardHTML(sides(k), vols(k), moCards(k), _
                    strikes(k), optTypes(k), prices(k), _
                    bQty, bSym, bBroker, cpFrom, cpTo, _
                    printBkt, tradeDate, deltaRatio)
            Next k
        Next pg
    Next b

    html = html & "</div></body></html>"

    Dim filePath As String
    filePath = ThisWorkbook.Path & "\GFI_Cards_" & Format$(Now(), "YYYYMMDD_HHMMSS") & ".html"

    Dim fNum As Integer: fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, html
    Close #fNum

    Shell "cmd /c start """" """ & filePath & """", vbHide
    MsgBox "Cards opened in browser - Ctrl+P to print or save as PDF." & vbNewLine & _
           legCount & " leg(s), " & bktCount & " bracket(s)." & vbNewLine & vbNewLine & _
           "Tip: Set paper to Letter, no scaling.", vbInformation
End Sub


Private Function ValidatePrices(ws As Worksheet, legCount As Integer, _
    legRows() As Long, optTypes() As String, _
    strikes() As String, prices() As String) As Boolean

    ValidatePrices = True

    Dim segPrems(1 To 50) As Double
    Dim segCount As Integer: segCount = 0
    Dim k As Integer, s As Integer
    Dim found As Boolean

    For k = 1 To legCount
        If optTypes(k) = "" And Trim$(strikes(k)) = "" Then GoTo NextLeg

        Dim pkgPrem As Double: pkgPrem = 0
        On Error Resume Next
        pkgPrem = CDbl(ws.Cells(legRows(k), COL_PKG_PREM).Value)
        On Error GoTo 0

        found = False
        For s = 1 To segCount
            If segPrems(s) = pkgPrem Then found = True: Exit For
        Next s
        If Not found Then
            segCount = segCount + 1
            segPrems(segCount) = pkgPrem
        End If
NextLeg:
    Next k

    If segCount = 0 Then Exit Function

    Dim seg As Integer
    For seg = 1 To segCount
        Dim pkgPremSeg As Double: pkgPremSeg = segPrems(seg)

        Dim segSides(1 To 50)  As String
        Dim segVols(1 To 50)   As Double
        Dim segPrices(1 To 50) As Double
        Dim segLegCount As Integer: segLegCount = 0
        Dim allPricesFilled As Boolean: allPricesFilled = True

        For k = 1 To legCount
            If optTypes(k) = "" And Trim$(strikes(k)) = "" Then GoTo NextLeg2

            Dim lp As Double: lp = 0
            On Error Resume Next
            lp = CDbl(ws.Cells(legRows(k), COL_PKG_PREM).Value)
            On Error GoTo 0

            If lp = pkgPremSeg Then
                If Trim$(prices(k)) = "" Then
                    allPricesFilled = False
                Else
                    segLegCount = segLegCount + 1
                    segSides(segLegCount) = ws.Cells(legRows(k), 3).Value
                    segVols(segLegCount) = CDbl(ws.Cells(legRows(k), 4).Value)
                    segPrices(segLegCount) = CDbl(prices(k))
                End If
            End If
NextLeg2:
        Next k

        If Not allPricesFilled Then GoTo NextSeg
        If segLegCount = 0 Then GoTo NextSeg

        Dim baseVol As Double: baseVol = segVols(1)
        Dim j As Integer
        For j = 2 To segLegCount
            If segVols(j) < baseVol Then baseVol = segVols(j)
        Next j
        If baseVol = 0 Then GoTo NextSeg

        Dim netPrem As Double: netPrem = 0
        For j = 1 To segLegCount
            Dim legSign As Double
            legSign = IIf(segSides(j) = "S", 1, -1)
            netPrem = netPrem + legSign * (segVols(j) / baseVol) * segPrices(j)
        Next j

        Dim calcNet As Double: calcNet = Abs(netPrem)
        Dim diff As Double:    diff = Abs(calcNet - pkgPremSeg)

        If diff > 0.000001 Then
            MsgBox "Price reconciliation failed for segment with package price " & _
                   Format$(pkgPremSeg, "0.0000") & "." & vbNewLine & vbNewLine & _
                   "Expected net: " & Format$(pkgPremSeg, "0.0000") & vbNewLine & _
                   "Calculated net: " & Format$(calcNet, "0.0000") & vbNewLine & _
                   "Discrepancy: " & Format$(diff, "0.0000") & vbNewLine & vbNewLine & _
                   "Please check your leg prices in column J.", vbCritical
            ValidatePrices = False
            Exit Function
        End If
NextSeg:
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
    s = s & "@media print {" & vbNewLine
    s = s & "  body { background:white; padding:0; margin:0; }" & vbNewLine
    s = s & "  @page { size:letter portrait; margin:0.35in; }" & vbNewLine
    s = s & "  .cards-wrap { gap:0.15in; }" & vbNewLine
    s = s & "  .card { width:3.5in; height:5.5in; border:1.5px solid !important; "
    s = s & "-webkit-print-color-adjust:exact; print-color-adjust:exact; }" & vbNewLine
    s = s & "}" & vbNewLine
    s = s & "</style></head><body><div class='cards-wrap'>" & vbNewLine

    BuildHTMLHeader = s
End Function


Private Function BuildCardHTML(side As String, vol As Double, _
    moCode As String, strike As String, optType As String, price As String, _
    bQty() As Double, bSym() As String, bBroker() As String, _
    cpFrom As Integer, cpTo As Integer, _
    bracket As String, tradeDate As String, _
    deltaRatio As Double) As String

    Dim isFut As Boolean
    isFut = (optType = "" And Trim$(strike) = "")

    Dim cardType As String, cardRole As String, cpRole As String
    Dim bgColor As String, ink As String

    If isFut Then
        cardType = "FUTURES"
        cardRole = IIf(side = "B", "BUYER", "SELLER")
        cpRole = IIf(side = "B", "SELLER", "BUYER")
        bgColor = "#fefce8"
    ElseIf UCase$(optType) = "C" Then
        cardType = "CALL"
        cardRole = IIf(side = "S", "SELLER", "BUYER")
        cpRole = IIf(side = "S", "BUYER", "SELLER")
        bgColor = "#ffffff"
    Else
        cardType = "PUT"
        cardRole = IIf(side = "S", "SELLER", "BUYER")
        cpRole = IIf(side = "S", "BUYER", "SELLER")
        bgColor = "#f5f0c8"
    End If

    ink = IIf(cardRole = "BUYER", "#1f4e79", "#cc2222")

    Dim brokerName As String: brokerName = ""
    If cpFrom >= 1 And cpFrom <= UBound(bBroker) Then
        brokerName = bBroker(cpFrom)
    End If

    Dim h As String

    h = "<div class='card' style='background:" & bgColor & ";border-color:" & ink & ";'>" & vbNewLine
    h = h & "<div class='card-header'>" & vbNewLine
    h = h & "<div class='card-top-row'>" & vbNewLine
    h = h & "<div class='card-type' style='color:" & ink & "'>" & cardType & "</div>" & vbNewLine
    h = h & "<div class='card-broker' style='color:" & ink & "'>" & brokerName & "</div>" & vbNewLine
    h = h & "</div>" & vbNewLine
    h = h & "<div class='card-role' style='color:" & ink & "'>" & cardRole & "</div>" & vbNewLine
    h = h & "</div>" & vbNewLine
    h = h & "<hr class='card-rule' style='border-color:" & ink & "'>" & vbNewLine

    Dim qtyLbl As String, strLbl As String, prLbl As String, bktLbl As String
    If isFut Then
        qtyLbl = "CARS": strLbl = "": prLbl = "PRICE": bktLbl = "BK"
    Else
        qtyLbl = "QTY.": strLbl = "STRIKE": prLbl = "PREM.": bktLbl = "BKT."
    End If

    h = h & "<div class='col-headers' style='border-color:" & ink & ";color:" & ink & "'>" & vbNewLine
    h = h & "<div class='w-qty' style='border-right:0.5px solid " & ink & "'>" & qtyLbl & "</div>" & vbNewLine
    h = h & "<div class='w-mo' style='border-right:0.5px solid " & ink & "'>MO.</div>" & vbNewLine
    h = h & "<div class='w-str' style='border-right:0.5px solid " & ink & "'>" & strLbl & "</div>" & vbNewLine
    h = h & "<div class='w-pr' style='border-right:0.5px solid " & ink & "'>" & prLbl & "</div>" & vbNewLine
    h = h & "<div class='w-cp' style='border-right:0.5px solid " & ink & "'>" & cpRole & "</div>" & vbNewLine
    h = h & "<div class='w-bkt'>" & bktLbl & "</div>" & vbNewLine
    h = h & "</div>" & vbNewLine
    h = h & "<div class='slots'>" & vbNewLine

    Dim slot As Integer
    For slot = 1 To 5
        Dim cpIdx As Integer: cpIdx = cpFrom + slot - 1
        h = h & "<div class='slot' style='border-color:" & ink & "'>" & vbNewLine

        If cpIdx <= cpTo Then
            Dim displayQty As Long
            If isFut Then
                displayQty = CLng(Round(bQty(cpIdx) * deltaRatio, 0))
            Else
                displayQty = CLng(bQty(cpIdx))
            End If

            h = h & "<div class='cell w-qty' style='color:" & ink & ";border-color:" & ink & "'>"
            h = h & CStr(displayQty) & "</div>" & vbNewLine
            h = h & "<div class='cell w-mo' style='color:" & ink & ";border-color:" & ink & "'>"
            h = h & UCase$(moCode) & "</div>" & vbNewLine

            If isFut Then
                h = h & "<div class='cell w-str' style='border-color:" & ink & "'>&nbsp;</div>" & vbNewLine
            Else
                h = h & "<div class='cell w-str' style='color:" & ink & ";border-color:" & ink & "'>"
                h = h & strike & "</div>" & vbNewLine
            End If

            h = h & "<div class='cell w-pr' style='color:" & ink & ";border-color:" & ink & "'>"
            h = h & price & "</div>" & vbNewLine

            Dim rawSym As String: rawSym = bSym(cpIdx)
            Dim slashPos As Integer: slashPos = InStr(rawSym, "/")
            Dim symTop As String, symBot As String
            If slashPos > 0 Then
                symTop = Trim$(Left$(rawSym, slashPos - 1))
                symBot = Trim$(Mid$(rawSym, slashPos + 1))
            Else
                symTop = Trim$(rawSym)
                symBot = "&nbsp;"
            End If

            h = h & "<div class='cp-cell w-cp' style='border-color:" & ink & "'>" & vbNewLine
            h = h & "<div class='cp-top' style='border-color:" & ink & "'>" & symTop & "</div>" & vbNewLine
            h = h & "<div class='cp-bot'>" & symBot & "</div>" & vbNewLine
            h = h & "</div>" & vbNewLine
            h = h & "<div class='cell w-bkt' style='color:" & ink & ";border-right:none'>"
            h = h & bracket & "</div>" & vbNewLine

        Else
            h = h & "<div class='cell w-qty' style='border-color:" & ink & "'>&nbsp;</div>" & vbNewLine
            h = h & "<div class='cell w-mo' style='border-color:" & ink & "'>&nbsp;</div>" & vbNewLine
            h = h & "<div class='cell w-str' style='border-color:" & ink & "'>&nbsp;</div>" & vbNewLine
            h = h & "<div class='cell w-pr' style='border-color:" & ink & "'>&nbsp;</div>" & vbNewLine
            h = h & "<div class='cp-cell w-cp' style='border-color:" & ink & "'>" & vbNewLine
            h = h & "<div class='cp-top' style='border-color:" & ink & "'>&nbsp;</div>" & vbNewLine
            h = h & "<div class='cp-bot'>&nbsp;</div>" & vbNewLine
            h = h & "</div>" & vbNewLine
            h = h & "<div class='cell w-bkt' style='border-right:none'>&nbsp;</div>" & vbNewLine
        End If

        h = h & "</div>" & vbNewLine
    Next slot

    h = h & "</div>" & vbNewLine
    h = h & "<div class='card-footer' style='color:" & ink & ";border-color:" & ink & "'>"
    h = h & "TC S-P OPT.&nbsp;&nbsp;&nbsp;LAZARE Printing Co., Inc.&nbsp;&nbsp;&nbsp;(773) 871-2500</div>" & vbNewLine
    h = h & "</div>" & vbNewLine

    BuildCardHTML = h
End Function

