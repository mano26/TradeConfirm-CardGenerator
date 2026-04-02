Attribute VB_Name = "TicketGenerator_v4"
Option Explicit

Private Type TicketLeg
    side As String
    optType As String
    qty As String
    mo As String
    strike As String
    price As String
End Type

Public Function GenerateTicketFile(ticketNum As Long) As String
    GenerateTicketFile = ""
    
    Dim ws1 As Worksheet: Set ws1 = ThisWorkbook.Sheets(SH1_NAME)
    Dim ws2 As Worksheet: Set ws2 = ThisWorkbook.Sheets(SH2_NAME)
    
    Dim legs() As TicketLeg
    ReDim legs(1 To 50)
    Dim legCount As Integer: legCount = 0
    Dim r As Long: r = S1_CONF_START
    Dim blankRun As Integer: blankRun = 0
    
    Do While r <= 200
        If ws1.Cells(r, S1_COL_VOL).Value <> "" Then
            blankRun = 0
            legCount = legCount + 1
            
            Dim rawSide As String: rawSide = CStr(ws1.Cells(r, S1_COL_SIDE).Value)
            Dim rawOpt As String: rawOpt = Trim$(CStr(ws1.Cells(r, S1_COL_OPTTYPE).Value))
            Dim rawStrike As String
            
            If ws1.Cells(r, S1_COL_STRIKE).Value = "" Then
                rawStrike = ""
            Else
                Dim sd As Double: sd = CDbl(ws1.Cells(r, S1_COL_STRIKE).Value)
                Dim ss As String: ss = CStr(sd)
                If InStr(ss, ".") = 0 Then
                    rawStrike = ss & ".00"
                ElseIf Len(ss) - InStr(ss, ".") < 2 Then
                    rawStrike = ss & "0"
                Else
                    rawStrike = ss
                End If
            End If
            
            Dim isFut As Boolean
            isFut = (rawOpt = "" And Trim$(rawStrike) = "")
            
            If isFut Then
                legs(legCount).optType = "FUT"
            ElseIf UCase$(rawOpt) = "C" Then
                legs(legCount).optType = "CALL"
            ElseIf UCase$(rawOpt) = "P" Then
                legs(legCount).optType = "PUT"
            Else
                legs(legCount).optType = "CALL"
            End If
            
            If rawSide = "B" Then
                legs(legCount).side = "BUY"
            Else
                legs(legCount).side = "SELL"
            End If
            
            legs(legCount).qty = CStr(CLng(CDbl(ws1.Cells(r, S1_COL_VOL).Value)))
            
            Dim moVal As String
            moVal = Trim$(CStr(ws1.Cells(r, S1_COL_MO_CARD).Value))
            If moVal = "" Then moVal = Trim$(CStr(ws1.Cells(r, S1_COL_EXPIRY).Value))
            legs(legCount).mo = UCase$(moVal)
            
            legs(legCount).strike = rawStrike
            legs(legCount).price = Trim$(CStr(ws1.Cells(r, S1_COL_PRICE).Value))
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
    
    ' Read bracket from Sheet 2 (first one found)
    Dim bracket As String: bracket = ""
    Dim i As Integer
    For i = S2_CP_DATA_START To S2_CP_DATA_END
        Dim bk As String
        bk = Trim$(UCase$(CStr(ws2.Cells(i, S2_CP_COL_BRACKET).Value)))
        If bk <> "" Then
            bracket = bk
            Exit For
        End If
    Next i
    
    ' Read all unique brokers from counterparty rows
    Dim broker As String: broker = ""
    Dim brokerList(1 To 20) As String
    Dim brokerCount As Integer: brokerCount = 0
    Dim bi As Integer
    For bi = S2_CP_DATA_START To S2_CP_DATA_END
        Dim bVal As String
        bVal = Trim$(UCase$(CStr(ws2.Cells(bi, S2_CP_COL_BROKER).Value)))
        If bVal <> "" Then
            Dim bExists As Boolean: bExists = False
            Dim bx As Integer
            For bx = 1 To brokerCount
                If brokerList(bx) = bVal Then
                    bExists = True
                    Exit For
                End If
            Next bx
            If Not bExists Then
                brokerCount = brokerCount + 1
                brokerList(brokerCount) = bVal
            End If
        End If
    Next bi
    
    Dim bj As Integer
    For bj = 1 To brokerCount
        If bj > 1 Then broker = broker & " / "
        broker = broker & brokerList(bj)
    Next bj
    
    ' Max rows per type per side
    Dim maxRows As Integer: maxRows = 1
    Dim bc As Integer, bp As Integer, bf As Integer
    Dim sc As Integer, sp As Integer, sf As Integer
    bc = 0: bp = 0: bf = 0: sc = 0: sp = 0: sf = 0
    
    Dim k As Integer
    For k = 1 To legCount
        If legs(k).side = "BUY" Then
            Select Case legs(k).optType
                Case "CALL": bc = bc + 1
                Case "PUT": bp = bp + 1
                Case "FUT": bf = bf + 1
            End Select
        Else
            Select Case legs(k).optType
                Case "CALL": sc = sc + 1
                Case "PUT": sp = sp + 1
                Case "FUT": sf = sf + 1
            End Select
        End If
    Next k
    
    If bc > maxRows Then maxRows = bc
    If bp > maxRows Then maxRows = bp
    If bf > maxRows Then maxRows = bf
    If sc > maxRows Then maxRows = sc
    If sp > maxRows Then maxRows = sp
    If sf > maxRows Then maxRows = sf
    If maxRows > 4 Then maxRows = 4
    
    ' Build HTML
    Dim html As String
    html = BuildTicketHTMLHeader(maxRows)
    html = html & BuildTicketHTML(ticketNum, legs, legCount, maxRows, bracket, broker)
    html = html & "</div></body></html>"
    
    ' Save to permanent dated folder
    Dim filePath As String
    filePath = GetOutputFolder() & "\AXIS_Ticket_" & Format$(ticketNum, "0000") & _
               "_" & Format$(Now(), "YYYYMMDD_HHMMSS") & ".html"
    
    Dim fNum As Integer: fNum = FreeFile
    On Error GoTo TicketFileErr
    Open filePath For Output As #fNum
    Print #fNum, html
    Close #fNum
    On Error GoTo 0
    
    Shell "cmd /c start """" """ & filePath & """", vbHide
    
    GenerateTicketFile = filePath
    Exit Function
    
TicketFileErr:
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
    MsgBox "Error saving ticket file: " & filePath & vbNewLine & Err.Description, vbCritical
End Function
Private Sub CollectLegs(legs() As TicketLeg, legCount As Integer, _
        targetSide As String, targetType As String, _
        outQty() As String, outMo() As String, _
        outStr() As String, outPr() As String, outCount As Integer)
    outCount = 0
    Dim k As Integer
    For k = 1 To legCount
        If legs(k).side = targetSide And legs(k).optType = targetType Then
            outCount = outCount + 1
            If outCount <= 4 Then
                outQty(outCount) = legs(k).qty
                outMo(outCount) = legs(k).mo
                outStr(outCount) = legs(k).strike
                outPr(outCount) = legs(k).price
            End If
        End If
    Next k
End Sub

Private Function BuildTicketHTMLHeader(maxRows As Integer) As String
    Dim cF As Integer, tF As Integer, sf As Integer, lF As Integer
    Select Case maxRows
        Case 1: cF = 14: tF = 24: sf = 20: lF = 13
        Case 2: cF = 12: tF = 22: sf = 18: lF = 12
        Case 3: cF = 10: tF = 20: sf = 16: lF = 11
        Case Else: cF = 9: tF = 18: sf = 15: lF = 10
    End Select
    
    Dim s As String
    s = "<!DOCTYPE html><html><head><meta charset='utf-8'><title>AXIS Ticket</title>" & vbNewLine
    s = s & "<style>" & vbNewLine
    s = s & "* { box-sizing:border-box; margin:0; padding:0; }" & vbNewLine
    s = s & "body { font-family:Arial,Helvetica,sans-serif; background:#e0e0e0; padding:0.4in; }" & vbNewLine
    s = s & ".tickets-wrap { display:flex; flex-wrap:wrap; gap:0.25in; justify-content:center; }" & vbNewLine
    s = s & ".ticket { width:8in; height:5.5in; border:1.5px solid #000; background:#fff; "
    s = s & "padding:14px 18px; display:flex; flex-direction:column; page-break-inside:avoid; }" & vbNewLine
    s = s & ".tkt-header { display:flex; justify-content:space-between; align-items:flex-start; "
    s = s & "margin-bottom:4px; }" & vbNewLine
    s = s & ".tkt-num { font-size:15px; color:#cc2222; font-weight:700; font-family:monospace; }" & vbNewLine
    s = s & ".tkt-title { font-size:" & tF & "px; font-weight:900; letter-spacing:5px; "
    s = s & "text-align:center; flex:1; }" & vbNewLine
    s = s & ".tkt-acct { text-align:right; font-size:10px; }" & vbNewLine
    s = s & ".tkt-acct-box { border:1px solid #888; width:80px; height:20px; margin-top:2px; }" & vbNewLine
    s = s & ".tkt-body { display:flex; flex:1; gap:0; border-top:1.5px solid #000; }" & vbNewLine
    s = s & ".tkt-side { flex:1; display:flex; flex-direction:column; padding:5px 8px; }" & vbNewLine
    s = s & ".tkt-side + .tkt-side { border-left:1.5px solid #000; }" & vbNewLine
    s = s & ".side-title { font-size:" & sf & "px; font-weight:900; text-align:center; "
    s = s & "letter-spacing:4px; margin-bottom:3px; }" & vbNewLine
    s = s & ".opt-section { display:flex; align-items:stretch; margin-bottom:1px; }" & vbNewLine
    s = s & ".opt-label { font-size:" & lF & "px; font-weight:700; width:40px; "
    s = s & "display:flex; align-items:center; flex-shrink:0; }" & vbNewLine
    s = s & ".opt-grid { flex:1; display:grid; grid-template-columns:1fr 1.3fr 1fr 1fr; }" & vbNewLine
    s = s & ".opt-cell-group { border:0.5px solid #888; display:flex; flex-direction:column; }" & vbNewLine
    s = s & ".opt-entry { flex:1; display:flex; align-items:center; justify-content:center; "
    s = s & "font-size:" & cF & "px; font-weight:600; padding:1px 2px; text-align:center; "
    s = s & "min-height:18px; }" & vbNewLine
    s = s & ".col-hdrs { display:flex; margin-left:40px; }" & vbNewLine
    s = s & ".col-hdr { font-size:7px; font-weight:700; text-align:center; color:#555; "
    s = s & "padding:0 1px; }" & vbNewLine
    s = s & ".col-hdr:nth-child(1){flex:1}" & vbNewLine
    s = s & ".col-hdr:nth-child(2){flex:1.3}" & vbNewLine
    s = s & ".col-hdr:nth-child(3){flex:1}" & vbNewLine
    s = s & ".col-hdr:nth-child(4){flex:1}" & vbNewLine
    s = s & ".con-cxl { display:flex; align-items:center; margin-top:3px; }" & vbNewLine
    s = s & ".con-cxl-label { font-size:10px; font-weight:700; width:40px; line-height:1.1; }" & vbNewLine
    s = s & ".con-cxl-arrow { font-size:14px; margin-left:4px; }" & vbNewLine
    s = s & ".tkt-footer { margin-top:auto; padding-top:6px; border-top:1px solid #aaa; "
    s = s & "text-align:center; }" & vbNewLine
    s = s & ".bracket-row { display:flex; gap:3px; justify-content:center; flex-wrap:wrap; "
    s = s & "font-size:11px; font-weight:700; margin-bottom:5px; }" & vbNewLine
    s = s & ".bkt-letter { width:15px; height:15px; display:flex; align-items:center; "
    s = s & "justify-content:center; }" & vbNewLine
    s = s & ".bkt-letter.circled { border:2px solid #cc2222; border-radius:50%; "
    s = s & "color:#cc2222; }" & vbNewLine
    s = s & ".footer-row { display:flex; align-items:center; justify-content:space-between; "
    s = s & "font-size:9px; margin-top:4px; }" & vbNewLine
    s = s & ".footer-section { display:flex; align-items:center; gap:10px; }" & vbNewLine
    s = s & ".check-box { display:inline-block; width:9px; height:9px; border:0.5px solid #888; "
    s = s & "margin-right:2px; }" & vbNewLine
    s = s & ".broker-box { border:1px solid #888; padding:2px 12px; font-size:10px; "
    s = s & "text-align:center; min-width:70px; }" & vbNewLine
    s = s & ".broker-label { font-size:7px; color:#666; }" & vbNewLine
    s = s & ".slmq-box { display:flex; flex-direction:column; align-items:center; "
    s = s & "font-size:10px; font-weight:700; border:0.5px solid #888; padding:2px 6px; "
    s = s & "line-height:1.2; }" & vbNewLine
    s = s & ".lazare { font-size:7px; color:#999; margin-top:5px; }" & vbNewLine
    s = s & "@media print {" & vbNewLine
    s = s & "  body { background:white; padding:0; margin:0; }" & vbNewLine
    s = s & "  @page { size:8in 5.5in; margin:0; }" & vbNewLine
    s = s & "  .ticket { width:8in; height:5.5in; border:1.5px solid #000 !important; "
    s = s & "-webkit-print-color-adjust:exact; print-color-adjust:exact; }" & vbNewLine
    s = s & "}" & vbNewLine
    s = s & "</style></head><body><div class='tickets-wrap'>" & vbNewLine
    BuildTicketHTMLHeader = s
End Function

Private Function BuildTicketHTML(ticketNum As Long, legs() As TicketLeg, _
        legCount As Integer, maxRows As Integer, _
        bracket As String, broker As String) As String
    Dim h As String
    h = "<div class='ticket'>" & vbNewLine
    h = h & "<div class='tkt-header'>"
    h = h & "<div class='tkt-num'>" & Format$(ticketNum, "0000") & "</div>"
    h = h & "<div class='tkt-title'>A X I S</div>"
    h = h & "<div class='tkt-acct'>Account No.<div class='tkt-acct-box'></div></div>"
    h = h & "</div>" & vbNewLine
    h = h & "<div class='tkt-body'>" & vbNewLine
    h = h & BuildSideHTML(legs, legCount, "BUY", maxRows)
    h = h & BuildSideHTML(legs, legCount, "SELL", maxRows)
    h = h & "</div>" & vbNewLine
    h = h & "<div class='tkt-footer'>" & vbNewLine
    h = h & BuildBracketRow(bracket)
    h = h & "<div class='footer-row'>"
    h = h & "<div class='footer-section'>"
    h = h & "<span class='check-box'></span> INITIAL &nbsp;&nbsp;&nbsp;"
    h = h & "<span class='check-box'></span> CLOSING</div>"
    h = h & "<div class='slmq-box'>S<br>L<br>M<br>Q</div>"
    h = h & "<div style='text-align:center'>"
    h = h & "<div class='broker-box'>" & broker & "</div>"
    h = h & "<div class='broker-label'>Broker No.</div></div>"
    h = h & "<div class='footer-section'>"
    h = h & "<span class='check-box'></span> INITIAL &nbsp;&nbsp;&nbsp;"
    h = h & "<span class='check-box'></span> CLOSING</div>"
    h = h & "</div>" & vbNewLine
    h = h & "<div class='lazare'>LAZARE Printing Co., Inc.&nbsp;&nbsp;&nbsp;(773) 871-2500</div>" & vbNewLine
    h = h & "</div>" & vbNewLine
    h = h & "</div>" & vbNewLine
    BuildTicketHTML = h
End Function

Private Function BuildSideHTML(legs() As TicketLeg, legCount As Integer, _
        sideName As String, maxRows As Integer) As String
    Dim h As String
    h = "<div class='tkt-side'>"
    h = h & "<div class='side-title'>" & sideName & "</div>" & vbNewLine
    h = h & BuildTypeSection(legs, legCount, sideName, "CALL", maxRows)
    h = h & "<div class='col-hdrs'>"
    h = h & "<div class='col-hdr'>QUANTITY</div>"
    h = h & "<div class='col-hdr'>CONTRACT/MONTH</div>"
    h = h & "<div class='col-hdr'>STRIKE</div>"
    h = h & "<div class='col-hdr'>PREMIUM</div></div>" & vbNewLine
    h = h & BuildTypeSection(legs, legCount, sideName, "PUT", maxRows)
    h = h & BuildTypeSection(legs, legCount, sideName, "FUT", maxRows)
    h = h & "<div class='con-cxl'>"
    h = h & "<div class='con-cxl-label'>CON<br>CXL</div>"
    h = h & "<div class='con-cxl-arrow'>&#9655;</div></div>"
    h = h & "</div>" & vbNewLine
    BuildSideHTML = h
End Function

Private Function BuildTypeSection(legs() As TicketLeg, legCount As Integer, _
        sideName As String, typeName As String, maxRows As Integer) As String
    Dim lQ(1 To 4) As String, lM(1 To 4) As String
    Dim lS(1 To 4) As String, lp(1 To 4) As String
    Dim cnt As Integer
    Call CollectLegs(legs, legCount, sideName, typeName, lQ, lM, lS, lp, cnt)
    
    Dim n As Integer: n = maxRows
    If n < 1 Then n = 1
    
    Dim h As String, j As Integer
    Dim cellVal As String
    
    h = "<div class='opt-section'>"
    h = h & "<div class='opt-label'>" & typeName & "</div>"
    h = h & "<div class='opt-grid'>" & vbNewLine
    
    ' QTY column
    h = h & "<div class='opt-cell-group'>"
    For j = 1 To n
        If j <= cnt Then
            cellVal = lQ(j)
        Else
            cellVal = "&nbsp;"
        End If
        h = h & "<div class='opt-entry'>" & cellVal & "</div>"
    Next j
    h = h & "</div>" & vbNewLine
    
    ' CONTRACT/MONTH column
    h = h & "<div class='opt-cell-group'>"
    For j = 1 To n
        If j <= cnt Then
            cellVal = lM(j)
        Else
            cellVal = "&nbsp;"
        End If
        h = h & "<div class='opt-entry'>" & cellVal & "</div>"
    Next j
    h = h & "</div>" & vbNewLine
    
    ' STRIKE column
    h = h & "<div class='opt-cell-group'>"
    For j = 1 To n
        If j <= cnt Then
            cellVal = lS(j)
        Else
            cellVal = "&nbsp;"
        End If
        h = h & "<div class='opt-entry'>" & cellVal & "</div>"
    Next j
    h = h & "</div>" & vbNewLine
    
    ' PREMIUM column
    h = h & "<div class='opt-cell-group'>"
    For j = 1 To n
        If j <= cnt Then
            cellVal = lp(j)
        Else
            cellVal = "&nbsp;"
        End If
        h = h & "<div class='opt-entry'>" & cellVal & "</div>"
    Next j
    h = h & "</div>" & vbNewLine
    
    h = h & "</div></div>" & vbNewLine
    BuildTypeSection = h
End Function

Private Function BuildBracketRow(activeBracket As String) As String
    Dim letters As Variant
    letters = Array("$", "A", "B", "C", "D", "E", "F", "G", "H", "I", _
                    "J", "K", "L", "M", "N", "O", "P", "Q", " ", _
                    "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
                    "2", "3", "4", "5", "6", "7", "8", "9", "%")
    
    Dim h As String
    h = "<div class='bracket-row'>" & vbNewLine
    
    Dim i As Integer
    For i = LBound(letters) To UBound(letters)
        Dim l As String: l = CStr(letters(i))
        If l = " " Then
            h = h & "<div style='width:8px'></div>" & vbNewLine
        Else
            Dim c As String: c = "bkt-letter"
            If UCase$(l) = UCase$(activeBracket) Then c = c & " circled"
            h = h & "<div class='" & c & "'>" & l & "</div>" & vbNewLine
        End If
    Next i
    
    h = h & "</div>" & vbNewLine
    BuildBracketRow = h
End Function
