Attribute VB_Name = "ConfirmationEngine"
Option Explicit

Sub ClearConfirmationOutput()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GFI Upload Template")
    ws.Range("C5:R1000").ClearContents
    ws.Range("S5:S1000").ClearContents
    ws.Range("T5:T1000").ClearContents
    ws.Range("U5:U1000").ClearContents
    ws.Range("J5:J1000").Interior.ColorIndex = xlNone
    ws.Range("S5:S1000").Interior.ColorIndex = xlNone

    ' Hide and narrow cols T and U
    ws.Columns("T").ColumnWidth = 0.5
    ws.Columns("U").ColumnWidth = 0.5

    ' Restore counterparty section headers
    ws.Range("B12").Value = "DATE"
    ws.Range("C12").Formula = "=TODAY()"
    ws.Range("D12").Value = "QTY"
    ws.Range("E12").Value = "OPPOSITE/HOUSE"
    ws.Range("F12").Value = "EXECUTING BROKER"
    ws.Range("G12").Value = "BRACKET"
    ws.Range("H12").Value = "NOTES"

    On Error Resume Next
    ThisWorkbook.Sheets("GFI Upload Template").Shapes("btnGenerateCards").Visible = False
    On Error GoTo 0
End Sub

Public Sub ClearPage()
    Call ClearConfirmationOutput
End Sub

Public Sub ProcessTrade()
    Call GenerateConfirmation
End Sub

Sub GenerateConfirmation()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GFI Upload Template")

    Dim inputLine As String
    inputLine = Trim$(ws.Range("D3").Value)

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
        Dim parts() As String
        parts = Split(inputLine, ",")
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

    Dim nextRow As Long: nextRow = 5
    Dim trade As TradeInput
    Dim segStartRow As Long

    For i = 1 To tradeParts.Count
        Set trade = tradeParts(i)
        segStartRow = nextRow

        Select Case trade.Strategy
            Case "straddle":          nextRow = BuildStraddle(trade, nextRow)
            Case "strangle":          nextRow = BuildStrangle(trade, nextRow)
            Case "cs":                nextRow = BuildCallSpread(trade, nextRow)
            Case "ps":                nextRow = BuildPutSpread(trade, nextRow)
            Case "rr":                nextRow = BuildRiskReversal(trade, nextRow)
            Case "bflyc":             nextRow = BuildCallButterfly(trade, nextRow)
            Case "bflyp":             nextRow = BuildPutButterfly(trade, nextRow)
            Case "ctree":             nextRow = BuildCallChristmasTree(trade, nextRow)
            Case "ptree":             nextRow = BuildPutChristmasTree(trade, nextRow)
            Case "condorc":           nextRow = BuildCallCondor(trade, nextRow)
            Case "condorp":           nextRow = BuildPutCondor(trade, nextRow)
            Case "ic":                nextRow = BuildIronCondor(trade, nextRow)
            Case "ibfly":             nextRow = BuildIronButterfly(trade, nextRow)
            Case "box":               nextRow = BuildBoxSpread(trade, nextRow)
            Case "single", "c", "p": nextRow = BuildSingleOption(trade, nextRow)
            Case Else
                MsgBox "Strategy not recognised: '" & trade.Strategy & "'" & vbNewLine & _
                       "Check your strategy token in D3.", vbCritical
        End Select

        If (trade.IsCVD = True) Or (trade.CVDPrice <> 0) Then
            nextRow = BuildCvdOverlay(trade, nextRow)
        End If

        ' Stamp package premium to col U for option rows in this segment
        ' White font so it is invisible to user but readable by VBA
        Dim stampRow As Long
        For stampRow = segStartRow To nextRow - 1
            If ws.Cells(stampRow, 4).Value <> "" Then
                If ws.Cells(stampRow, 9).Value <> "" Then
                    ws.Cells(stampRow, 21).Value = trade.Premium
                    ws.Cells(stampRow, 21).Font.Color = RGB(255, 255, 255)
                End If
            End If
        Next stampRow

        nextRow = nextRow + 1
    Next i

    ' Ensure cols T and U stay narrow and hidden after writing
    ws.Columns("T").ColumnWidth = 0.5
    ws.Columns("U").ColumnWidth = 0.5

    On Error Resume Next
    ws.Shapes("btnGenerateCards").Visible = True
    On Error GoTo 0
End Sub

Public Function ParseGenericWrapper(inputLine As String) As Collection
    Dim result As New Collection
    Set ParseGenericWrapper = result

    Dim closePos As Integer: closePos = InStr(inputLine, "]")
    If closePos = 0 Then
        MsgBox "[] syntax error: missing closing bracket in: " & inputLine, vbCritical
        Exit Function
    End If

    Dim inner   As String: inner = Trim$(Mid$(inputLine, 2, closePos - 2))
    Dim trailer As String: trailer = Trim$(Mid$(inputLine, closePos + 1))

    If Len(Trim$(inner)) = 0 Then
        MsgBox "[] syntax error: no content inside brackets.", vbCritical
        Exit Function
    End If

    If Len(Trim$(trailer)) = 0 Then
        MsgBox "[] syntax error: missing price/qty after closing bracket.", vbCritical
        Exit Function
    End If

    Dim pkgSide As String
    pkgSide = IIf(InStr(trailer, "@") > 0, "S", "B")

    Dim pkgVol  As Double: pkgVol = 0
    Dim pkgPrem As Double: pkgPrem = 0
    Dim slashPos As Integer: slashPos = InStr(trailer, "/")
    Dim atPos    As Integer: atPos = InStr(trailer, "@")

    On Error GoTo ParseWrapperError
    If slashPos > 0 Then
        pkgPrem = CDbl(Trim$(Left$(trailer, slashPos - 1)))
        pkgVol = CDbl(Trim$(Mid$(trailer, slashPos + 1)))
    ElseIf atPos > 0 Then
        pkgVol = CDbl(Trim$(Left$(trailer, atPos - 1)))
        pkgPrem = CDbl(Trim$(Mid$(trailer, atPos + 1)))
    End If
    On Error GoTo 0

    If pkgVol = 0 Then
        MsgBox "[] syntax error: could not parse volume from trailer: '" & trailer & "'", vbCritical
        Exit Function
    End If

    Dim segments() As String: segments = Split(inner, ",")
    Dim i As Long, j As Long

    For i = LBound(segments) To UBound(segments)
        Dim seg As String: seg = Trim$(segments(i))
        If Len(seg) > 0 Then
            Dim subParts As Collection
            Set subParts = ParseTradeInput(seg)
            If Not subParts Is Nothing Then
                For j = 1 To subParts.Count
                    Dim t As TradeInput
                    Set t = subParts(j)
                    t.DirectionSide = pkgSide
                    t.SuppressPremium = True
                    If t.Volume = 0 And pkgVol > 0 Then t.Volume = CLng(pkgVol)
                    If t.Premium = 0 And pkgPrem > 0 Then t.Premium = pkgPrem
                    result.Add t
                Next j
            End If
        End If
    Next i
    Exit Function

ParseWrapperError:
    MsgBox "[] syntax error: could not parse volume/premium from trailer: '" & trailer & "'", vbCritical
    On Error GoTo 0
End Function

