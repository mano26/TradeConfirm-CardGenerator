Attribute VB_Name = "WorkbookSetup_v4"
Option Explicit

' =========================================================================
'  WORKBOOK SETUP & SHARED CONSTANTS  (v4)
'  - Creates/validates the 3-sheet structure
'  - All layout constants referenced by every other module
'  - Ticket counter, Order Log helpers
' =========================================================================

' Sheet names
Public Const SH1_NAME As String = "Trade Entry-Confirmation"
Public Const SH2_NAME As String = "Cards & Tickets"
Public Const SH3_NAME As String = "Order Log"

' ── Sheet 1 layout ──
Public Const S1_INPUT_ROW As Long = 3
Public Const S1_INPUT_COL As Long = 4          ' D3 = trade input
Public Const S1_CONF_START As Long = 7         ' confirmation output begins row 7
Public Const S1_COL_SIDE As Long = 3           ' C
Public Const S1_COL_VOL As Long = 4            ' D
Public Const S1_COL_MARKET As Long = 5         ' E
Public Const S1_COL_CONTRACT As Long = 6       ' F
Public Const S1_COL_EXPIRY As Long = 7         ' G
Public Const S1_COL_STRIKE As Long = 8         ' H
Public Const S1_COL_OPTTYPE As Long = 9        ' I
Public Const S1_COL_PRICE As Long = 10         ' J
Public Const S1_COL_BROKER_STAMP As Long = 18  ' R  (AXIS stamp)
Public Const S1_COL_TICKET As Long = 19        ' S  (ticket #)
Public Const S1_COL_MO_CARD As Long = 20       ' T  (card MO code, hidden)
Public Const S1_COL_PKG_PREM As Long = 21      ' U  (package premium, hidden)

' Sheet 2 layout constants
Public Const S2_TRADE_REF_ROW As Long = 1
Public Const S2_HEADER_ROW As Long = 7          ' data entry row for house/account
Public Const S2_HOUSE_COL As Long = 1            ' A7
Public Const S2_ACCOUNT_COL As Long = 2          ' B7

Public Const S2_CP_HDR_ROW As Long = 9
Public Const S2_CP_DATA_START As Long = 10
Public Const S2_CP_DATA_END As Long = 29
Public Const S2_CP_COL_QTY As Long = 1          ' A
Public Const S2_CP_COL_BROKER As Long = 2       ' B
Public Const S2_CP_COL_SYMBOL As Long = 3       ' C
Public Const S2_CP_COL_BRACKET As Long = 4      ' D
Public Const S2_CP_COL_NOTES As Long = 5        ' E

' ── Sheet 3 (Order Log) columns ──
Public Const S3_COL_HOUSE As Long = 1          ' A
Public Const S3_COL_ACCOUNT As Long = 2        ' B
Public Const S3_COL_SIDE As Long = 3           ' C
Public Const S3_COL_VOL As Long = 4            ' D
Public Const S3_COL_MARKET As Long = 5         ' E
Public Const S3_COL_CONTRACT As Long = 6       ' F
Public Const S3_COL_EXPIRY As Long = 7         ' G
Public Const S3_COL_STRIKE As Long = 8         ' H
Public Const S3_COL_OPTTYPE As Long = 9        ' I
Public Const S3_COL_PRICE As Long = 10         ' J
Public Const S3_COL_BROKER As Long = 11        ' K
Public Const S3_COL_TICKET As Long = 12        ' L
Public Const S3_COL_LINKS As Long = 13         ' M

' Ticket counter
Public Const TKT_COUNTER_CELL As String = "W1"
Public Const MAX_TICKET As Long = 9999

' =========================================================================
'  Initialize workbook with 3 sheets
' =========================================================================
Public Sub InitializeWorkbook()
    Application.ScreenUpdating = False

    Dim ws1 As Worksheet: Set ws1 = EnsureSheet(SH1_NAME)
    Dim ws2 As Worksheet: Set ws2 = EnsureSheet(SH2_NAME)
    Dim ws3 As Worksheet: Set ws3 = EnsureSheet(SH3_NAME)

    SetupSheet1 ws1
    SetupSheet2 ws2
    SetupSheet3 ws3

    On Error Resume Next
    ws1.Move Before:=Sheets(1)
    ws2.Move After:=ws1
    ws3.Move After:=ws2
    On Error GoTo 0

    ws1.Activate
    ws1.Range("D3").Select
    Application.ScreenUpdating = True

    MsgBox "Workbook initialized:" & vbNewLine & _
           "  1) " & SH1_NAME & vbNewLine & _
           "  2) " & SH2_NAME & vbNewLine & _
           "  3) " & SH3_NAME, vbInformation
End Sub

Private Function EnsureSheet(shName As String) As Worksheet
    Dim ws As Worksheet
    Dim found As Boolean: found = False
    
    ' Check if sheet already exists
    Dim s As Worksheet
    For Each s In ThisWorkbook.Sheets
        If s.Name = shName Then
            Set ws = s
            found = True
            Exit For
        End If
    Next s
    
    ' Create if not found
    If Not found Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = shName
    End If
    
    Set EnsureSheet = ws
End Function
Private Sub SetupSheet1(ws As Worksheet)
    ws.Cells.HorizontalAlignment = xlCenter
    ws.Cells.VerticalAlignment = xlCenter
    ws.Columns("G").NumberFormat = "@"
    ws.Columns("H").NumberFormat = "0.0000"
    ws.Columns("J").NumberFormat = "0.0000"
    ws.Range("A1:C1").Interior.Color = RGB(0, 32, 96)
    ws.Range("A1:C1").Font.Color = RGB(255, 255, 255)
    ws.Range("H1:S1").Interior.Color = RGB(0, 32, 96)
    ws.Range("H1:S1").Font.Color = RGB(255, 255, 255)
    ws.Range("D1").Interior.Color = RGB(247, 255, 79)
    ws.Range("D1").Font.Color = RGB(0, 0, 0)
    ws.Range("C3:G3").Interior.Color = RGB(247, 255, 79)
    ws.Range("C3:G3").Font.Color = RGB(0, 0, 0)
    ' Instruction in D1
    ws.Range("D1").Value = "Enter trade in D3"
    ws.Range("D1").Font.Bold = True
    
    ' Title
    ws.Range("F1").Value = "AXIS TRADE ENTRY"
    ws.Range("F1").Font.Size = 22
    ws.Range("F1").Font.Bold = True
    
    ' Headers in row 5
    ws.Cells(5, 2).Value = "ACCOUNT"
    ws.Cells(5, 3).Value = "B/S"
    ws.Cells(5, 4).Value = "VOLUME"
    ws.Cells(5, 5).Value = "MARKET"
    ws.Cells(5, 6).Value = "CONTRACT"
    ws.Cells(5, 7).Value = "EXPIRY"
    ws.Cells(5, 8).Value = "STRIKE"
    ws.Cells(5, 9).Value = "C/P"
    ws.Cells(5, 10).Value = "PRICE"
    ws.Cells(5, 11).Value = "ORDER"
    ws.Cells(5, 12).Value = "FLOOR"
    ws.Cells(5, 13).Value = "COMMENT"
    ws.Cells(5, 14).Value = "MEMBER"
    ws.Cells(5, 15).Value = "TYPE"
    ws.Cells(5, 16).Value = "STRATEGY"
    ws.Cells(5, 17).Value = "TIME IN"
    ws.Cells(5, 18).Value = "BROKER"
    ws.Cells(5, 19).Value = "TICKET #"
    
    Dim c As Long
    For c = 2 To 19
        ws.Cells(5, c).Font.Bold = True
        ws.Cells(5, c).Font.Size = 12
    Next c
    
    ' Hide utility columns
    ws.Columns("T").Hidden = True
    ws.Columns("U").Hidden = True
    ws.Columns("W").Hidden = True
End Sub

Private Sub SetupSheet2(ws As Worksheet)
    Dim rng2 As Range
    Set rng2 = ws.Range("D" & S2_CP_DATA_START & ":D" & S2_CP_DATA_END)
    With rng2.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Sheet1!$C$4:$C$32"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Dim valRange As String
    valRange = "='" & SH1_NAME & "'!$A$4:$A$54"
    Dim rng As Range
    Set rng = ws.Range("C" & S2_CP_DATA_START & ":C" & S2_CP_DATA_END)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Sheet1!$A$4:$A$54"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    ws.Range("B1:E1").Merge
    ws.Range("B1:E1").Interior.Color = RGB(247, 255, 79)
    ws.Range("B1:E1").Font.Color = RGB(255, 0, 0)
    ws.Range("A6:B6").Interior.Color = RGB(0, 32, 96)
    ws.Range("A6:B6").Font.Color = RGB(255, 255, 255)
    ws.Range("A7:B7").Interior.Color = RGB(247, 255, 79)
    ws.Range("A7:B7").Font.Color = RGB(0, 0, 0)
    ws.Range("A9:E9").Interior.Color = RGB(0, 32, 96)
    ws.Range("A9:E9").Font.Color = RGB(255, 255, 255)
    ws.Cells.HorizontalAlignment = xlCenter
    ws.Cells.VerticalAlignment = xlCenter
    ws.Range("A1").Value = "CURRENT TRADE:"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    
    ' B1 populated by ProcessTrade at 14pt
    
    ws.Cells(4, 1).Value = "COUNTERPARTIES"
    ws.Cells(4, 1).Font.Bold = True
    ws.Cells(4, 1).Font.Size = 14
    
    ws.Cells(6, 1).Value = "HOUSE"
    ws.Cells(6, 2).Value = "ACCOUNT"
    ws.Cells(6, 1).Font.Bold = True
    ws.Cells(6, 2).Font.Bold = True
    ws.Cells(6, 1).Font.Size = 12
    ws.Cells(6, 2).Font.Size = 12
    
    ws.Cells(S2_CP_HDR_ROW, S2_CP_COL_QTY).Value = "QTY"
    ws.Cells(S2_CP_HDR_ROW, S2_CP_COL_BROKER).Value = "BROKER"
    ws.Cells(S2_CP_HDR_ROW, S2_CP_COL_SYMBOL).Value = "OPPOSITE/HOUSE"
    ws.Cells(S2_CP_HDR_ROW, S2_CP_COL_BRACKET).Value = "BRACKET"
    ws.Cells(S2_CP_HDR_ROW, S2_CP_COL_NOTES).Value = "NOTES"
    
    Dim c As Integer
    For c = S2_CP_COL_QTY To S2_CP_COL_NOTES
        ws.Cells(S2_CP_HDR_ROW, c).Font.Bold = True
        ws.Cells(S2_CP_HDR_ROW, c).Font.Size = 12
    Next c
    
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 14
End Sub
Private Sub SetupSheet3(ws As Worksheet)
    ws.Range("A1:N1").Interior.Color = RGB(0, 32, 96)
    ws.Range("A1:N1").Font.Color = RGB(255, 255, 255)
    ws.Columns("G").NumberFormat = "@"
    ws.Columns("H").NumberFormat = "0.0000"
    ws.Columns("J").NumberFormat = "0.0000"
    ws.Cells.HorizontalAlignment = xlCenter
    ws.Cells.VerticalAlignment = xlCenter
    If ws.Cells(1, 1).Value <> "" Then Exit Sub

    Dim h As Variant
    h = Array("HOUSE", "ACCOUNT", "B/S", "VOLUME", "MARKET", _
              "CONTRACT", "EXPIRY", "STRIKE", "C/P", "PRICE", _
              "BROKER", "TICKET #", "LINKS")
    Dim i As Integer
    For i = LBound(h) To UBound(h)
        ws.Cells(1, i + 1).Value = h(i)
        ws.Cells(1, i + 1).Font.Bold = True
        ws.Cells(1, i + 1).Font.Size = 12
    Next i

    ' Row 2: blank separator between headers and first data
    
    ws.Activate
    ws.Rows("3:3").Select
    ActiveWindow.FreezePanes = True

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 5
    ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 8
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("G").ColumnWidth = 9
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 5
    ws.Columns("J").ColumnWidth = 9
    ws.Columns("K").ColumnWidth = 9
    ws.Columns("L").ColumnWidth = 10
    ws.Columns("M").ColumnWidth = 35
End Sub
' =========================================================================
'  Order Log helpers
' =========================================================================
Public Function GetNextLogRow() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH3_NAME)
    GetNextLogRow = ws.Cells(ws.Rows.Count, S3_COL_SIDE).End(xlUp).Row + 1
    If GetNextLogRow <= 2 Then GetNextLogRow = 3  ' skip header + blank row
End Function

Public Function FindLogBlockByTicket(ticketNum As Long) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH3_NAME)
    Dim tktStr As String: tktStr = Format$(ticketNum, "0000")
    Dim r As Long
    For r = 3 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If Trim$(CStr(ws.Cells(r, S3_COL_TICKET).Value)) = tktStr Then
            FindLogBlockByTicket = r
            Exit Function
        End If
    Next r
    FindLogBlockByTicket = 0
End Function

Public Sub DeleteLogBlock(startRow As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH3_NAME)
    Dim endRow As Long: endRow = startRow
    Do While endRow <= ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If ws.Cells(endRow, S3_COL_SIDE).Value = "" And _
           ws.Cells(endRow, S3_COL_HOUSE).Value = "" Then
            Exit Do
        End If
        endRow = endRow + 1
    Loop
    ws.Rows(startRow & ":" & endRow).Delete Shift:=xlUp
End Sub

' =========================================================================
'  Ticket counter  (stored in Sheet 1 hidden cell W1)
' =========================================================================
Public Function GetNextTicketNumber() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH1_NAME)
    ws.Columns("W").Hidden = True

    Dim current As Long
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

' =========================================================================
'  Read current ticket number WITHOUT incrementing (for re-processing check)
' =========================================================================
Public Function PeekCurrentTicketNumber() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH1_NAME)
    On Error Resume Next
    PeekCurrentTicketNumber = CLng(ws.Range(TKT_COUNTER_CELL).Value)
    On Error GoTo 0
End Function
Public Function GetOutputFolder() As String
    Dim basePath As String
    basePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\AXIS_Output"
    
    Dim datePath As String
    datePath = basePath & "\" & Format$(Now(), "MMDDYYYY")
    
    If Dir(basePath, vbDirectory) = "" Then
        MkDir basePath
    End If
    
    If Dir(datePath, vbDirectory) = "" Then
        MkDir datePath
    End If
    
    GetOutputFolder = datePath
End Function
