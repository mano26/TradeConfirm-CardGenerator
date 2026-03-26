Attribute VB_Name = "TicketGenerator"

Option Explicit

' ---------------------------------------------------------------------------
'  AXIS Ticket Generator v2
'  - One ticket per trade, all legs shown
'  - Up to 4 entries per type (CALL/PUT/FUT) per side
'  - Dynamic font sizing based on max rows needed
'  - No extra row borders or repeated labels — cells split vertically
'  - Ticket size: 8" x 5.5"
'  - Sequential numbering 0001-9999 in hidden cell W1
' ---------------------------------------------------------------------------

Private Const TKT_COUNTER_CELL As String = "W1"
Private Const MAX_TICKET As Long = 9999

Private Type TicketLeg
    side As String      ' "BUY" or "SELL"
    optType As String   ' "CALL", "PUT", or "FUT"
    qty As String
    mo As String
    strike As String
    price As String
End Type

Public Sub GenerateTickets()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GFI Upload Template")
    
    Dim legs() As TicketLeg
    ReDim legs(1 To 50)
    Dim legCount As Integer: legCount = 0
    
    Dim r As Long: r = 5
    Dim blankRun As Integer: blankRun = 0
    
    Do While r <= 200
        If ws.Cells(r, 4).Value <> "" Then
            blankRun = 0
            legCount = legCount + 1
            
            Dim rawSide As String: rawSide = CStr(ws.Cells(r, 3).Value)
            Dim rawOpt As String: rawOpt = Trim$(CStr(ws.Cells(r, 9).Value))
            Dim rawStrike As String
            
            If ws.Cells(r, 8).Value = "" Then
                rawStrike = ""
            Else
                Dim strikeDbl As Double: strikeDbl = CDbl(ws.Cells(r, 8).Value)
                Dim strikeStr As String: strikeStr = CStr(strikeDbl)
                If InStr(strikeStr, ".") = 0 Then
                    rawStrike = strikeStr & ".00"
                ElseIf Len(strikeStr) - InStr(strikeStr, ".") < 2 Then
                    rawStrike = strikeStr & "0"
                Else
                    rawStrike = strikeStr
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
            
            legs(legCount).side = IIf(rawSide = "B", "BUY", "SELL")
            legs(legCount).qty = CStr(CLng(CDbl(ws.Cells(r, 4).Value)))
            
            Dim moVal As String
            moVal = Trim$(CStr(ws.Cells(r, 20).Value))
            If moVal = "" Then moVal = Trim$(CStr(ws.Cells(r, 7).Value))
            legs(legCount).mo = UCase$(moVal)
            
            legs(legCount).strike = rawStrike
            legs(legCount).price = Trim$(CStr(ws.Cells(r, 10).Value))
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
    
    Dim bracket As String: bracket = ""
    Dim i As Integer
    For i = 13 To 32
        Dim bkt As String: bkt = Trim$(UCase$(CStr(ws.Cells(i, 7).Value)))
        If bkt <> "" Then bracket = bkt: Exit For
    Next i
    
    Dim broker As String: broker = ""
    For i = 13 To 32
        Dim brk As String: brk = Trim$(UCase$(CStr(ws.Cells(i, 6).Value)))
        If brk <> "" Then broker = brk: Exit For
    Next i
    
    Dim ticketNum As Long
    ticketNum = GetNextTicketNumber(ws)
    
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
    
    Dim html As String
    html = BuildTicketHTMLHeader(maxRows)
    html = html & BuildTicketHTML(ticketNum, legs, legCount, maxRows, bracket, broker)
    html = html & "</div></body></html>"
    
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\AXIS_Ticket_" & Format$(ticketNum, "0000") & _
               "_" & Format$(Now(), "YYYYMMDD_HHMMSS") & ".html"
    
    Dim fNum As Integer: fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, html
    Close #fNum
    
    Shell "cmd /c start """" """ & filePath & """", vbHide
    
    MsgBox "Ticket #" & Format$(ticketNum, "0000") & " generated." & vbNewLine & _
           legCount & " leg(s) on ticket." & vbNewLine & vbNewLine & _
           "Ctrl+P to print. Set paper to Letter, no scaling.", vbInformation
    
End Sub

Private Function GetNextTicketNumber(ws As Worksheet) As Long
    Dim current As Long
    ws.Columns("W").Hidden = True
    
    On Error Resume Next
    current = CLng(ws.Range(TKT_COUNTER_CELL).Value)
    On Error GoTo 0
    
    If current < 1 Or current > MAX_TICKET Then current = 0
    current = current + 1
    If current > MAX_TICKET Then current = 1
    
    ws.Range(TKT_COUNTER_CELL).Value = current
    ws.Range(TKT_COUNTER_CELL).Font.Color = RGB(255, 255, 255)
    
    GetNextTicketNumber = current
End Function

Private Sub CollectLegs(legs() As TicketLeg, legCount As Integer, _
                        targetSide As String, targetType As String, _
                        outQty() As String, outMo() As String, _
                        outStr() As String, outPr() As String, _
                        outCount As Integer)
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
    Dim cellFont As Integer
    Select Case maxRows
        Case 1: cellFont = 14
        Case 2: cellFont = 12
        Case 3: cellFont = 10
        Case Else: cellFont = 9
    End Select
    
    Dim titleFont As Integer
    Select Case maxRows
        Case 1: titleFont = 24
        Case 2: titleFont = 22
        Case 3: titleFont = 20
        Case Else: titleFont = 18
    End Select
    
    Dim sideTitleFont As Integer
    Select Case maxRows
        Case 1: sideTitleFont = 20
        Case 2: sideTitleFont = 18
        Case 3: sideTitleFont = 16
        Case Else: sideTitleFont = 15
    End Select
    
    Dim labelFont As Integer
    Select Case maxRows
        Case 1: labelFont = 13
        Case 2: labelFont = 12
        Case 3: labelFont = 11
        Case Else: labelFont = 10
    End Select
    
    Dim s As String
    s = "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    s = s & "<title>AXIS Ticket</title>" & vbNewLine
    s = s & "<style>" & vbNewLine
    s = s & "* { box-sizing:border-box; margin:0; padding:0; }" & vbNewLine
    s = s & "body { font-family:Arial,Helvetica,sans-serif; background:#e0e0e0; padding:0.4in; }" & vbNewLine
    s = s & ".tickets-wrap { display:flex; flex-wrap:wrap; gap:0.25in; justify-content:center; }" & vbNewLine
    s = s & ".ticket { width:8in; height:5.5in; border:1.5px solid #000; background:#fff; "
    s = s & "padding:14px 18px; display:flex; flex-direction:column; page-break-inside:avoid; }" & vbNewLine
    s = s & ".tkt-header { display:flex; justify-content:space-between; align-items:flex-start; "
    s = s & "margin-bottom:4px; }" & vbNewLine
    s = s & ".tkt-num { font-size:15px; color:#cc2222; font-weight:700; font-family:monospace; }" & vbNewLine
    s = s & ".tkt-title { font-size:" & titleFont & "px; font-weight:900; letter-spacing:5px; text-align:center; flex:1; }" & vbNewLine
    s = s & ".tkt-acct { text-align:right; font-size:10px; }" & vbNewLine
    s = s & ".tkt-acct-box { border:1px solid #888; width:80px; height:20px; margin-top:2px; }" & vbNewLine
    s = s & ".tkt-body { display:flex; flex:1; gap:0; border-top:1.5px solid #000; }" & vbNewLine
    s = s & ".tkt-side { flex:1; display:flex; flex-direction:column; padding:5px 8px; }" & vbNewLine
    s = s & ".tkt-side + .tkt-side { border-left:1.5px solid #000; }" & vbNewLine
    s = s & ".side-title { font-size:" & sideTitleFont & "px; font-weight:900; text-align:center; "
    s = s & "letter-spacing:4px; margin-bottom:3px; }" & vbNewLine
    s = s & ".opt-section { display:flex; align-items:stretch; margin-bottom:1px; }" & vbNewLine
    s = s & ".opt-label { font-size:" & labelFont & "px; font-weight:700; width:40px; "
    s = s & "display:flex; align-items:center; flex-shrink:0; }" & vbNewLine
    s = s & ".opt-grid { flex:1; display:grid; grid-template-columns:1fr 1.3fr 1fr 1fr; }" & vbNewLine
    s = s & ".opt-cell-group { border:0.5px solid #888; display:flex; flex-direction:column; }" & vbNewLine
    s = s & ".opt-entry { flex:1; display:flex; align-items:center; justify-content:center; "
    s = s & "font-size:" & cellFont & "px; font-weight:600; padding:1px 2px; text-align:center; "
    s = s & "min-height:18px; }" & vbNewLine
    s = s & ".col-hdrs { display:flex; margin-left:40px; }" & vbNewLine
    s = s & ".col-hdr { font-size:7px; font-weight:700; text-align:center; color:#555; padding:0 1px; }" & vbNewLine
    s = s & ".col-hdr:nth-child(1) { flex:1; }" & vbNewLine
    s = s & ".col-hdr:nth-child(2) { flex:1.3; }" & vbNewLine
    s = s & ".col-hdr:nth-child(3) { flex:1; }" & vbNewLine
    s = s & ".col-hdr:nth-child(4) { flex:1; }" & vbNewLine
    s = s & ".con-cxl { display:flex; align-items:center; margin-top:3px; }" & vbNewLine
    s = s & ".con-cxl-label { font-size:10px; font-weight:700; width:40px; line-height:1.1; }" & vbNewLine
    s = s & ".con-cxl-arrow { font-size:14px; margin-left:4px; }" & vbNewLine
    s = s & ".tkt-footer { margin-top:auto; padding-top:6px; border-top:1px solid #aaa; text-align:center; }" & vbNewLine
    s = s & ".bracket-row { display:flex; gap:3px; justify-content:center; flex-wrap:wrap; "
    s = s & "font-size:11px; font-weight:700; margin-bottom:5px; }" & vbNewLine
    s = s & ".bkt-letter { width:15px; height:15px; display:flex; align-items:center; justify-content:center; }" & vbNewLine
    s = s & ".bkt-letter.circled { border:2px solid #cc2222; border-radius:50%; color:#cc2222; }" & vbNewLine
    s = s & ".footer-row { display:flex; align-items:center; justify-content:space-between; "
    s = s & "font-size:9px; margin-top:4px; }" & vbNewLine
    s = s & ".footer-section { display:flex; align-items:center; gap:10px; }" & vbNewLine
    s = s & ".check-box { display:inline-block; width:9px; height:9px; border:0.5px solid #888; margin-right:2px; }" & vbNewLine
    s = s & ".broker-box { border:1px solid #888; padding:2px 12px; font-size:10px; text-align:center; min-width:70px; }" & vbNewLine
    s = s & ".broker-label { font-size:7px; color:#666; }" & vbNewLine
    s = s & ".slmq-box { display:flex; flex-direction:column; align-items:center; "
    s = s & "font-size:10px; font-weight:700; border:0.5px solid #888; padding:2px 6px; line-height:1.2; }" & vbNewLine
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
    h = h & "<div class='tkt-header'>" & vbNewLine
    h = h & "  <div class='tkt-num'>" & Format$(ticketNum, "0000") & "</div>" & vbNewLine
    h = h & "  <div class='tkt-title'>A X I S</div>" & vbNewLine
    h = h & "  <div class='tkt-acct'>Account No.<div class='tkt-acct-box'></div></div>" & vbNewLine
    h = h & "</div>" & vbNewLine
    h = h & "<div class='tkt-body'>" & vbNewLine
    h = h & BuildSideHTML(legs, legCount, "BUY", maxRows)
    h = h & BuildSideHTML(legs, legCount, "SELL", maxRows)
    h = h & "</div>" & vbNewLine
    h = h & "<div class='tkt-footer'>" & vbNewLine
    h = h & BuildBracketRow(bracket)
    h = h & "  <div class='footer-row'>" & vbNewLine
    h = h & "    <div class='footer-section'>"
    h = h & "<span class='check-box'></span> INITIAL &nbsp;&nbsp;&nbsp; "
    h = h & "<span class='check-box'></span> CLOSING</div>" & vbNewLine
    h = h & "    <div class='slmq-box'>S<br>L<br>M<br>Q</div>" & vbNewLine
    h = h & "    <div style='text-align:center'>"
    h = h & "<div class='broker-box'>" & broker & "</div>"
    h = h & "<div class='broker-label'>Broker No.</div></div>" & vbNewLine
    h = h & "    <div class='footer-section'>"
    h = h & "<span class='check-box'></span> INITIAL &nbsp;&nbsp;&nbsp; "
    h = h & "<span class='check-box'></span> CLOSING</div>" & vbNewLine
    h = h & "  </div>" & vbNewLine
    h = h & "  <div class='lazare'>LAZARE Printing Co., Inc.&nbsp;&nbsp;&nbsp;(773) 871-2500</div>" & vbNewLine
    h = h & "</div>" & vbNewLine
    h = h & "</div>" & vbNewLine
    BuildTicketHTML = h
End Function

Private Function BuildSideHTML(legs() As TicketLeg, legCount As Integer, _
                                sideName As String, maxRows As Integer) As String
    Dim h As String
    h = "<div class='tkt-side'>" & vbNewLine
    h = h & "  <div class='side-title'>" & sideName & "</div>" & vbNewLine
    h = h & BuildTypeSection(legs, legCount, sideName, "CALL", maxRows)
    h = h & "  <div class='col-hdrs'>"
    h = h & "<div class='col-hdr'>QUANTITY</div>"
    h = h & "<div class='col-hdr'>CONTRACT/MONTH</div>"
    h = h & "<div class='col-hdr'>STRIKE</div>"
    h = h & "<div class='col-hdr'>PREMIUM</div></div>" & vbNewLine
    h = h & BuildTypeSection(legs, legCount, sideName, "PUT", maxRows)
    h = h & BuildTypeSection(legs, legCount, sideName, "FUT", maxRows)
    h = h & "  <div class='con-cxl'><div class='con-cxl-label'>CON<br>CXL</div>"
    h = h & "<div class='con-cxl-arrow'>&#9655;</div></div>" & vbNewLine
    h = h & "</div>" & vbNewLine
    BuildSideHTML = h
End Function

Private Function BuildTypeSection(legs() As TicketLeg, legCount As Integer, _
                                   sideName As String, typeName As String, _
                                   maxRows As Integer) As String
    Dim lQty(1 To 4) As String, lMo(1 To 4) As String
    Dim lStr(1 To 4) As String, lPr(1 To 4) As String
    Dim cnt As Integer
    Call CollectLegs(legs, legCount, sideName, typeName, lQty, lMo, lStr, lPr, cnt)
    
    Dim numEntries As Integer
    numEntries = maxRows
    If numEntries < 1 Then numEntries = 1
    
    Dim h As String
    h = "  <div class='opt-section'>" & vbNewLine
    h = h & "    <div class='opt-label'>" & typeName & "</div>" & vbNewLine
    h = h & "    <div class='opt-grid'>" & vbNewLine
    
    Dim j As Integer
    
    h = h & "      <div class='opt-cell-group'>"
    For j = 1 To numEntries
        h = h & "<div class='opt-entry'>" & IIf(j <= cnt, lQty(j), "&nbsp;") & "</div>"
    Next j
    h = h & "</div>" & vbNewLine
    
    h = h & "      <div class='opt-cell-group'>"
    For j = 1 To numEntries
        h = h & "<div class='opt-entry'>" & IIf(j <= cnt, lMo(j), "&nbsp;") & "</div>"
    Next j
    h = h & "</div>" & vbNewLine
    
    h = h & "      <div class='opt-cell-group'>"
    For j = 1 To numEntries
        h = h & "<div class='opt-entry'>" & IIf(j <= cnt, lStr(j), "&nbsp;") & "</div>"
    Next j
    h = h & "</div>" & vbNewLine
    
    h = h & "      <div class='opt-cell-group'>"
    For j = 1 To numEntries
        h = h & "<div class='opt-entry'>" & IIf(j <= cnt, lPr(j), "&nbsp;") & "</div>"
    Next j
    h = h & "</div>" & vbNewLine
    
    h = h & "    </div>" & vbNewLine
    h = h & "  </div>" & vbNewLine
    BuildTypeSection = h
End Function

Private Function BuildBracketRow(activeBracket As String) As String
    Dim letters As Variant
    letters = Array("$", "A", "B", "C", "D", "E", "F", "G", "H", "I", _
                    "J", "K", "L", "M", "N", "O", "P", "Q", _
                    " ", _
                    "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
                    "2", "3", "4", "5", "6", "7", "8", "9", "%")
    
    Dim h As String
    h = "  <div class='bracket-row'>" & vbNewLine
    
    Dim i As Integer
    For i = LBound(letters) To UBound(letters)
        Dim ltr As String: ltr = CStr(letters(i))
        If ltr = " " Then
            h = h & "    <div style='width:8px'></div>" & vbNewLine
        Else
            Dim cls As String: cls = "bkt-letter"
            If UCase$(ltr) = UCase$(activeBracket) Then cls = cls & " circled"
            h = h & "    <div class='" & cls & "'>" & ltr & "</div>" & vbNewLine
        End If
    Next i
    
    h = h & "  </div>" & vbNewLine
    BuildBracketRow = h
End Function

