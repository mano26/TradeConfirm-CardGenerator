Attribute VB_Name = "ConfirmationEngine_v4"

Option Explicit

Public Sub ClearPage()
    Call ClearConfirmationOutput
End Sub

Public Sub ProcessTrade()
    Call GenerateConfirmation
End Sub

Sub ClearConfirmationOutput()
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets(SH1_NAME)
    
    ws1.Range("C6").ClearContents
    ws1.Range("C7:R1000").ClearContents
    ws1.Range("S7:S1000").ClearContents
    ws1.Range("T7:T1000").ClearContents
    ws1.Range("U7:U1000").ClearContents
    ws1.Range("J7:J1000").Interior.ColorIndex = xlNone
    ws1.Range("S7:S1000").Interior.ColorIndex = xlNone
    ws1.Columns("T").Hidden = True
    ws1.Columns("U").Hidden = True
    
    Dim ws2 As Worksheet
    Dim sheetExists As Boolean: sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = SH2_NAME Then
            sheetExists = True
            Exit For
        End If
    Next sh
    
    If sheetExists Then
        Set ws2 = ThisWorkbook.Sheets(SH2_NAME)
        ws2.Range("B1").MergeArea.ClearContents
        ws2.Cells(S2_HEADER_ROW, S2_HOUSE_COL).ClearContents
        ws2.Cells(S2_HEADER_ROW, S2_ACCOUNT_COL).ClearContents
        Dim cpR As Long
        For cpR = S2_CP_DATA_START To S2_CP_DATA_END
            ws2.Cells(cpR, S2_CP_COL_QTY).ClearContents
            ws2.Cells(cpR, S2_CP_COL_BROKER).ClearContents
            ws2.Cells(cpR, S2_CP_COL_SYMBOL).ClearContents
            ws2.Cells(cpR, S2_CP_COL_BRACKET).ClearContents
            ws2.Cells(cpR, S2_CP_COL_NOTES).ClearContents
            ws2.Range(ws2.Cells(cpR, S2_CP_COL_QTY), ws2.Cells(cpR, S2_CP_COL_NOTES)).Interior.ColorIndex = xlNone
        Next cpR
    End If
End Sub
Sub GenerateConfirmation()
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets(SH1_NAME)
    
    Dim inputLine As String
    inputLine = Trim$(ws1.Range("D3").Value)
    If inputLine = "" Then
        MsgBox "Please enter a trade in cell D3.", vbExclamation
        Exit Sub
    End If
    
    Dim tradeParts As Collection
    Dim subParts As Collection
    Dim i As Long, j As Long
    Set tradeParts = New Collection
    
    If Left$(inputLine, 1) = "[" Then
        Set tradeParts = ParseGenericWrapper(inputLine)
    Else
        Dim parts() As String: parts = Split(inputLine, ",")
        For i = LBound(parts) To UBound(parts)
            Dim seg As String: seg = Trim$(parts(i))
            If Len(seg) > 0 Then
                Set subParts = ParseTradeInput(seg)
                If Not subParts Is Nothing Then
                    For j = 1 To subParts.Count
                        tradeParts.Add subParts(j)
                    Next j
                End If
            End If
        Next i
    End If
    
    If tradeParts Is Nothing Or tradeParts.Count = 0 Then
        MsgBox "No valid trade legs found. Check your input syntax.", vbCritical
        Exit Sub
    End If
    
    Call ClearConfirmationOutput
    
    Dim nextRow As Long: nextRow = S1_CONF_START
    Dim trade As TradeInput
    Dim segStartRow As Long
    
    For i = 1 To tradeParts.Count
        Set trade = tradeParts(i)
        segStartRow = nextRow
        
        Select Case trade.Strategy
            Case "straddle": nextRow = BuildStraddle(trade, nextRow)
            Case "strangle": nextRow = BuildStrangle(trade, nextRow)
            Case "cs": nextRow = BuildCallSpread(trade, nextRow)
            Case "ps": nextRow = BuildPutSpread(trade, nextRow)
            Case "rr": nextRow = BuildRiskReversal(trade, nextRow)
            Case "bflyc": nextRow = BuildCallButterfly(trade, nextRow)
            Case "bflyp": nextRow = BuildPutButterfly(trade, nextRow)
            Case "ctree": nextRow = BuildCallChristmasTree(trade, nextRow)
            Case "ptree": nextRow = BuildPutChristmasTree(trade, nextRow)
            Case "condorc": nextRow = BuildCallCondor(trade, nextRow)
            Case "condorp": nextRow = BuildPutCondor(trade, nextRow)
            Case "ic": nextRow = BuildIronCondor(trade, nextRow)
            Case "ibfly": nextRow = BuildIronButterfly(trade, nextRow)
            Case "box": nextRow = BuildBoxSpread(trade, nextRow)
            Case "single", "c", "p": nextRow = BuildSingleOption(trade, nextRow)
            Case Else
                MsgBox "Strategy not recognised: '" & trade.Strategy & "'", vbCritical
        End Select
        
        If (trade.IsCVD = True) Or (trade.CVDPrice <> 0) Then
            nextRow = BuildCvdOverlay(trade, nextRow)
        End If
        
        Dim stampRow As Long
        For stampRow = segStartRow To nextRow - 1
            If ws1.Cells(stampRow, S1_COL_VOL).Value <> "" Then
                If ws1.Cells(stampRow, S1_COL_OPTTYPE).Value <> "" Then
                    ws1.Cells(stampRow, S1_COL_PKG_PREM).Value = trade.Premium
                    ws1.Cells(stampRow, S1_COL_PKG_PREM).Font.Color = RGB(255, 255, 255)
                End If
            End If
        Next stampRow
        
        nextRow = nextRow + 1
    Next i
    
    ws1.Columns("T").ColumnWidth = 0.5
    ws1.Columns("U").ColumnWidth = 0.5
    
    ' Get ticket number
    Dim ticketNum As Long
    ticketNum = GetNextTicketNumber()
    
    ' Stamp ticket number on all confirmation rows
    Dim r As Long
    For r = S1_CONF_START To nextRow - 1
        If ws1.Cells(r, S1_COL_VOL).Value <> "" Then
            ws1.Cells(r, S1_COL_TICKET).NumberFormat = "@"
            ws1.Cells(r, S1_COL_TICKET).Value = Format$(ticketNum, "0000")
        End If
    Next r
    
    ' Show trade input on Sheet 2 header (just the raw string)
    ' Store trade input on Sheet 2 header for reference
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets(SH2_NAME)
    ws2.Range("B1").Value = UCase$(inputLine)
    ws2.Range("B1").Font.Size = 14
    
    ' Duplicate trade input to C6 on Sheet 1
    ws1.Range("C6").Value = inputLine
    
    MsgBox "Trade processed - " & (nextRow - S1_CONF_START) & " leg(s)." & vbNewLine & _
           "Fill prices in column J, then go to '" & SH2_NAME & "' to generate.", vbInformation
    
End Sub
Private Sub AppendToOrderLog(ws1 As Worksheet, ticketNum As Long)
    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets(SH3_NAME)
    
    Dim logRow As Long: logRow = GetNextLogRow()
    
    ' Add blank separator if not first entry
    If logRow > 3 Then
        If ws3.Cells(logRow - 1, S3_COL_SIDE).Value <> "" Then
            logRow = logRow + 1
        End If
    End If
    
    Dim r As Long: r = S1_CONF_START
    Dim blankRun As Integer: blankRun = 0
    
    Do While r <= 200
        If ws1.Cells(r, S1_COL_VOL).Value <> "" Then
            blankRun = 0
            ws3.Cells(logRow, S3_COL_SIDE).Value = ws1.Cells(r, S1_COL_SIDE).Value
            ws3.Cells(logRow, S3_COL_VOL).Value = ws1.Cells(r, S1_COL_VOL).Value
            ws3.Cells(logRow, S3_COL_MARKET).Value = ws1.Cells(r, S1_COL_MARKET).Value
            ws3.Cells(logRow, S3_COL_CONTRACT).Value = ws1.Cells(r, S1_COL_CONTRACT).Value
            ws3.Cells(logRow, S3_COL_EXPIRY).Value = ws1.Cells(r, S1_COL_EXPIRY).Value
            ws3.Cells(logRow, S3_COL_STRIKE).Value = ws1.Cells(r, S1_COL_STRIKE).Value
            ws3.Cells(logRow, S3_COL_OPTTYPE).Value = ws1.Cells(r, S1_COL_OPTTYPE).Value
            ws3.Cells(logRow, S3_COL_PRICE).Value = ws1.Cells(r, S1_COL_PRICE).Value
            ws3.Cells(logRow, S3_COL_TICKET).NumberFormat = "@"
            ws3.Cells(logRow, S3_COL_TICKET).Value = Format$(ticketNum, "0000")
            logRow = logRow + 1
        Else
            blankRun = blankRun + 1
            If blankRun >= 2 Then Exit Do
        End If
        r = r + 1
    Loop
End Sub
Private Sub CopyTradeToSheet2(ws1 As Worksheet, inputLine As String)
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets(SH2_NAME)
    
    ' Row 1: raw trade input
    ws2.Range("B1").Value = inputLine
    ws2.Range("B1").Font.Size = 12
    
    ' Row 2: leg summary
    Dim summary As String: summary = ""
    Dim r As Long: r = S1_CONF_START
    Dim blankRun As Integer: blankRun = 0
    
    Do While r <= 200
        If ws1.Cells(r, S1_COL_VOL).Value <> "" Then
            blankRun = 0
            If summary <> "" Then summary = summary & "  |  "
            summary = summary & ws1.Cells(r, S1_COL_SIDE).Value & " "
            summary = summary & ws1.Cells(r, S1_COL_VOL).Value & " "
            summary = summary & ws1.Cells(r, S1_COL_CONTRACT).Value & " "
            summary = summary & ws1.Cells(r, S1_COL_EXPIRY).Value
            If ws1.Cells(r, S1_COL_STRIKE).Value <> "" Then
                summary = summary & " " & ws1.Cells(r, S1_COL_STRIKE).Value
                summary = summary & " " & ws1.Cells(r, S1_COL_OPTTYPE).Value
            End If
        Else
            blankRun = blankRun + 1
            If blankRun >= 2 Then Exit Do
        End If
        r = r + 1
    Loop
    
    ws2.Range("B2").Value = summary
    ws2.Range("B2").Font.Size = 9
    ws2.Range("B2").Font.Color = RGB(100, 100, 100)
End Sub

Public Function ParseGenericWrapper(inputLine As String) As Collection
    Dim result As New Collection
    Set ParseGenericWrapper = result
    
    Dim closePos As Integer: closePos = InStr(inputLine, "]")
    If closePos = 0 Then
        MsgBox "[] syntax error: missing ]", vbCritical
        Exit Function
    End If
    
    Dim inner As String: inner = Trim$(Mid$(inputLine, 2, closePos - 2))
    Dim trailer As String: trailer = Trim$(Mid$(inputLine, closePos + 1))
    
    If Len(inner) = 0 Then
        MsgBox "[] syntax error: empty brackets", vbCritical
        Exit Function
    End If
    If Len(trailer) = 0 Then
        MsgBox "[] syntax error: no price/qty after ]", vbCritical
        Exit Function
    End If
    
    Dim pkgSide As String
    pkgSide = IIf(InStr(trailer, "@") > 0, "S", "B")
    Dim pkgVol As Double: pkgVol = 0
    Dim pkgPrem As Double: pkgPrem = 0
    Dim slashPos As Integer: slashPos = InStr(trailer, "/")
    Dim atPos As Integer: atPos = InStr(trailer, "@")
    
    On Error GoTo WrapErr
    If slashPos > 0 Then
        pkgPrem = CDbl(Trim$(Left$(trailer, slashPos - 1)))
        pkgVol = CDbl(Trim$(Mid$(trailer, slashPos + 1)))
    ElseIf atPos > 0 Then
        pkgVol = CDbl(Trim$(Left$(trailer, atPos - 1)))
        pkgPrem = CDbl(Trim$(Mid$(trailer, atPos + 1)))
    End If
    On Error GoTo 0
    
    If pkgVol = 0 Then
        MsgBox "[] syntax error: volume = 0", vbCritical
        Exit Function
    End If
    
    Dim segments() As String: segments = Split(inner, ",")
    Dim i As Long, j2 As Long
    For i = LBound(segments) To UBound(segments)
        Dim s As String: s = Trim$(segments(i))
        If Len(s) > 0 Then
            Dim sp As Collection
            Set sp = ParseTradeInput(s)
            If Not sp Is Nothing Then
                For j2 = 1 To sp.Count
                    Dim t As TradeInput
                    Set t = sp(j2)
                    t.DirectionSide = pkgSide
                    t.SuppressPremium = True
                    If t.Volume = 0 And pkgVol > 0 Then t.Volume = CLng(pkgVol)
                    If t.Premium = 0 And pkgPrem > 0 Then t.Premium = pkgPrem
                    result.Add t
                Next j2
            End If
        End If
    Next i
    Exit Function
    
WrapErr:
    MsgBox "[] syntax error parsing trailer: '" & trailer & "'", vbCritical
    On Error GoTo 0
End Function

Public Sub StampLogMetadata()
    Dim ws2 As Worksheet: Set ws2 = ThisWorkbook.Sheets(SH2_NAME)
    Dim ws3 As Worksheet: Set ws3 = ThisWorkbook.Sheets(SH3_NAME)
    
    Dim house As String
    house = Trim$(CStr(ws2.Cells(S2_HEADER_ROW, S2_HOUSE_COL).Value))
    Dim account As String
    account = Trim$(CStr(ws2.Cells(S2_HEADER_ROW, S2_ACCOUNT_COL).Value))
    
    Dim tktNum As Long: tktNum = PeekCurrentTicketNumber()
    Dim tktStr As String: tktStr = Format$(tktNum, "0000")
    
    Dim lastRow As Long
    lastRow = ws3.Cells(ws3.Rows.Count, S3_COL_TICKET).End(xlUp).Row
    If lastRow < 3 Then Exit Sub
    
    Dim firstFound As Boolean: firstFound = False
    Dim r As Long
    For r = 3 To lastRow
        Dim cellTkt As String
        cellTkt = Trim$(CStr(ws3.Cells(r, S3_COL_TICKET).Value))
        If cellTkt = tktStr Then
            If Not firstFound Then
                ws3.Cells(r, S3_COL_HOUSE).Value = house
                ws3.Cells(r, S3_COL_ACCOUNT).Value = account
                firstFound = True
            End If
        End If
    Next r
End Sub
Public Sub StampLogLink(filePath As String, linkLabel As String)
    Dim ws3 As Worksheet: Set ws3 = ThisWorkbook.Sheets(SH3_NAME)
    
    Dim tktNum As Long: tktNum = PeekCurrentTicketNumber()
    Dim tktStr As String: tktStr = Format$(tktNum, "0000")
    
    Dim lastRow As Long
    lastRow = ws3.Cells(ws3.Rows.Count, S3_COL_TICKET).End(xlUp).Row
    If lastRow < 3 Then Exit Sub
    
    Dim r As Long
    For r = 3 To lastRow
        Dim cellTkt As String
        cellTkt = Trim$(CStr(ws3.Cells(r, S3_COL_TICKET).Value))
        If cellTkt = tktStr Then
            Dim existingLink As String
            existingLink = Trim$(CStr(ws3.Cells(r, S3_COL_LINKS).Value))
            If existingLink = "" Then
                ws3.Hyperlinks.Add _
                    Anchor:=ws3.Cells(r, S3_COL_LINKS), _
                    Address:=filePath, _
                    TextToDisplay:=linkLabel
                Exit Sub
            End If
        End If
    Next r
    
    For r = 3 To lastRow
        cellTkt = Trim$(CStr(ws3.Cells(r, S3_COL_TICKET).Value))
        If cellTkt = tktStr Then
            ws3.Cells(r, S3_COL_LINKS).ClearContents
            ws3.Hyperlinks.Add _
                Anchor:=ws3.Cells(r, S3_COL_LINKS), _
                Address:=filePath, _
                TextToDisplay:=linkLabel
            Exit Sub
        End If
    Next r
End Sub

Public Sub GenerateCardsAndTickets()
    ' Shared validation
    If Not TradeValidation_v4.ValidateBeforeGenerate() Then Exit Sub
    
    Dim ws1 As Worksheet: Set ws1 = ThisWorkbook.Sheets(SH1_NAME)
    
    ' Get the ticket number assigned during ProcessTrade
    Dim ticketNum As Long: ticketNum = PeekCurrentTicketNumber()
    Dim tktStr As String: tktStr = Format$(ticketNum, "0000")
    
    ' Check for existing log block (re-generation) and delete it
    Dim existingRow As Long
    existingRow = FindLogBlockByTicket(ticketNum)
    If existingRow > 0 Then
        Call DeleteLogBlock(existingRow)
    End If
    
    ' Append to Order Log NOW (prices are filled at this point)
    Call AppendToOrderLog(ws1, ticketNum)
    
    ' Stamp House/Account/Broker from Sheet 2 onto the log
    Call StampLogMetadata
    
    ' Generate ticket HTML
    Dim ticketPath As String
    ticketPath = GenerateTicketFile(ticketNum)
    
    ' Generate cards HTML
    Dim cardsPath As String
    cardsPath = GenerateCardsFile()
    
    ' Stamp links to Order Log
    If ticketPath <> "" Then
        Call StampLogLink(ticketPath, "Ticket #" & tktStr)
    End If
    If cardsPath <> "" Then
        Call StampLogLink(cardsPath, "Cards")
    End If
    
    MsgBox "Ticket #" & tktStr & " and cards generated." & vbNewLine & _
           "Both opened in browser - Ctrl+P to print.", vbInformation
End Sub

