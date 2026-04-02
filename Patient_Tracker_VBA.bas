Attribute VB_Name = "Module1"
Sub GenerateMonthlySummaryAndKPIs()

    Dim ws As Worksheet
    Dim countsWS As Worksheet, KPIsWS As Worksheet
    Dim lastRow As Long
    Dim monthName As String
    Dim patientCount As Long, esseCount As Long, inpatientCount As Long, turndownCount As Long
    Dim colKPI As Long, rowSheet As Long, rowProcType As Long
    Dim procTypes As Variant, procType As Variant
    Dim validMonths As Variant, m As Variant, isMonthSheet As Boolean
    Dim echoDate As Variant, procDate As Variant
    Dim i As Long, count As Long, totalDays As Double

    procTypes = Array("TAVR", "TAVR/PCI", "mTEER", "redo mTEER", "SAVR", "SAVR/CABG", "TMVR")
    validMonths = Array("January", "February", "March", "April", "May", "June", _
                        "July", "August", "September", "October", "November", "December")

    ' Create or clear MonthlyCounts sheet
    On Error Resume Next
    Set countsWS = ThisWorkbook.Worksheets("MonthlyCounts")
    If countsWS Is Nothing Then
        Set countsWS = ThisWorkbook.Worksheets.Add
        countsWS.Name = "MonthlyCounts"
    Else
        countsWS.Cells.Clear
    End If

    ' Create or clear EchoKPIs sheet
    Set KPIsWS = ThisWorkbook.Worksheets("EchoKPIs")
    If KPIsWS Is Nothing Then
        Set KPIsWS = ThisWorkbook.Worksheets.Add
        KPIsWS.Name = "EchoKPIs"
    Else
        KPIsWS.Cells.Clear
    End If
    On Error GoTo 0

    ' MonthlyCounts header
    With countsWS
        .Range("A1:E1").Value = Array("Month", "Total Patients", "ESSE Patients", "Inpatient", "Surgical Turndowns")
    End With

    ' EchoKPIs header
    KPIsWS.Cells(1, 1).Value = "Procedure Type"
    Dim sectionTitles As Variant
    sectionTitles = Array("Echo to Procedure, Average Days", "Eval to Procedure, Average Days", "Eval to Gated CTA, Average Days")
    Dim sectionStartRow As Long: sectionStartRow = 2

    ' Write section titles and procedure types in column A
    For i = 0 To UBound(sectionTitles)
        KPIsWS.Cells(sectionStartRow, 1).Value = sectionTitles(i)
        For rowProcType = 0 To UBound(procTypes)
            KPIsWS.Cells(sectionStartRow + rowProcType + 1, 1).Value = procTypes(rowProcType)
        Next rowProcType
        sectionStartRow = sectionStartRow + UBound(procTypes) + 2
    Next i

    colKPI = 2 ' Start writing months in column B

    ' Loop through monthly sheets
    For Each ws In ThisWorkbook.Worksheets
        monthName = Trim(ws.Name)
        isMonthSheet = False

        For Each m In validMonths
            If InStr(1, monthName, m, vbTextCompare) > 0 Then
                isMonthSheet = True
                Exit For
            End If
        Next m

        If isMonthSheet Then
            lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
            KPIsWS.Cells(1, colKPI).Value = monthName

            For rowProcType = 0 To UBound(procTypes)
                procType = procTypes(rowProcType)

                ' Echo ? Procedure (L to AA)
                totalDays = 0: count = 0
                For rowSheet = 2 To lastRow
                    If Trim(ws.Cells(rowSheet, "Y").Value) = procType Then
                        echoDate = ws.Cells(rowSheet, "L").Value
                        procDate = ws.Cells(rowSheet, "AA").Value
                        If IsDate(echoDate) And IsDate(procDate) And procDate > echoDate Then
                            totalDays = totalDays + (procDate - echoDate)
                            count = count + 1
                        End If
                    End If
                Next rowSheet
                If count > 0 Then
                    KPIsWS.Cells(2 + rowProcType + 1, colKPI).Value = Round(totalDays / count, 1)
                Else
                    KPIsWS.Cells(2 + rowProcType + 1, colKPI).Value = ""
                End If

                ' Eval ? Procedure (I to AA)
                totalDays = 0: count = 0
                For rowSheet = 2 To lastRow
                    If Trim(ws.Cells(rowSheet, "Y").Value) = procType Then
                        echoDate = ws.Cells(rowSheet, "I").Value
                        procDate = ws.Cells(rowSheet, "AA").Value
                        If IsDate(echoDate) And IsDate(procDate) And procDate > echoDate Then
                            totalDays = totalDays + (procDate - echoDate)
                            count = count + 1
                        End If
                    End If
                Next rowSheet
                If count > 0 Then
                    KPIsWS.Cells(2 + UBound(procTypes) + 2 + rowProcType + 1, colKPI).Value = Round(totalDays / count, 1)
                Else
                    KPIsWS.Cells(2 + UBound(procTypes) + 2 + rowProcType + 1, colKPI).Value = ""
                End If

                ' Eval ? Gated CTA (I to U)
                totalDays = 0: count = 0
                For rowSheet = 2 To lastRow
                    If Trim(ws.Cells(rowSheet, "Y").Value) = procType Then
                        echoDate = ws.Cells(rowSheet, "I").Value
                        procDate = ws.Cells(rowSheet, "U").Value
                        If IsDate(echoDate) And IsDate(procDate) And procDate > echoDate Then
                            totalDays = totalDays + (procDate - echoDate)
                            count = count + 1
                        End If
                    End If
                Next rowSheet
                If count > 0 Then
                    KPIsWS.Cells(2 + (UBound(procTypes) + 2) * 2 + rowProcType + 1, colKPI).Value = Round(totalDays / count, 1)
                Else
                    KPIsWS.Cells(2 + (UBound(procTypes) + 2) * 2 + rowProcType + 1, colKPI).Value = ""
                End If
            Next rowProcType

            ' Count metrics for MonthlyCounts
            patientCount = 0: esseCount = 0: inpatientCount = 0: turndownCount = 0
            For rowSheet = 2 To lastRow
                If Trim(ws.Cells(rowSheet, "A").Value) <> "" Then patientCount = patientCount + 1
                If UCase(Trim(ws.Cells(rowSheet, "E").Value)) = "YES" Then esseCount = esseCount + 1
                If UCase(Trim(ws.Cells(rowSheet, "H").Value)) = "INPT" Then inpatientCount = inpatientCount + 1
                If UCase(Trim(ws.Cells(rowSheet, "X").Value)) = "YES" Then turndownCount = turndownCount + 1
            Next rowSheet

            ' Write to MonthlyCounts
            countsWS.Cells(countsWS.Cells(countsWS.Rows.count, "A").End(xlUp).Row + 1, "A").Resize(1, 5).Value = _
                Array(monthName, patientCount, esseCount, inpatientCount, turndownCount)

            colKPI = colKPI + 1
        End If
    Next ws

    ' Add YTD Average column
    Dim lastCol As Long, ytdCol As Long
    Dim currentMonthIndex As Long, monthCell As Range, monthIndex As Long

    currentMonthIndex = Month(Date) - 1
    lastCol = KPIsWS.Cells(1, KPIsWS.Columns.count).End(xlToLeft).Column
    ytdCol = lastCol + 1
    KPIsWS.Cells(1, ytdCol).Value = "YTD Avg"

    For i = 0 To 2 ' Three KPI sections
        Dim rowStart As Long, rowEnd As Long
        rowStart = 2 + i * (UBound(procTypes) + 2) + 1
        rowEnd = rowStart + UBound(procTypes)

        Dim r As Long, c As Long, sum As Double
        For r = rowStart To rowEnd
            sum = 0: count = 0

            For c = 2 To lastCol
                Set monthCell = KPIsWS.Cells(1, c)
                If Not IsError(Application.Match(monthCell.Value, validMonths, 0)) Then
                    monthIndex = Application.Match(monthCell.Value, validMonths, 0)
                    If monthIndex <= currentMonthIndex Then
                        If IsNumeric(KPIsWS.Cells(r, c).Value) Then
                            sum = sum + KPIsWS.Cells(r, c).Value
                            count = count + 1
                        End If
                    End If
                End If
            Next c

            If count > 0 Then
                KPIsWS.Cells(r, ytdCol).Value = Round(sum / count, 1)
            Else
                KPIsWS.Cells(r, ytdCol).Value = ""
            End If
        Next r
    Next i

    ' Add Totals row to MonthlyCounts
    Dim lastCountRow As Long
    lastCountRow = countsWS.Cells(countsWS.Rows.count, "A").End(xlUp).Row + 1

    With countsWS
        .Cells(lastCountRow, "A").Value = "TOTAL"
        .Cells(lastCountRow, "B").Formula = "=SUM(B2:B" & lastCountRow - 1 & ")"
        .Cells(lastCountRow, "C").Formula = "=SUM(C2:C" & lastCountRow - 1 & ")"
        .Cells(lastCountRow, "D").Formula = "=SUM(D2:D" & lastCountRow - 1 & ")"
        .Cells(lastCountRow, "E").Formula = "=SUM(E2:E" & lastCountRow - 1 & ")"
        .Range("A" & lastCountRow & ":E" & lastCountRow).Font.Bold = True
        .Range("A" & lastCountRow & ":E" & lastCountRow).Interior.Color = RGB(230, 230, 230)
    End With

    ' -------------------------
    ' FORMATTING FOR ECHOKPIS
    ' -------------------------

    ' Add month labels to rows 10 and 18
    Dim monthLabels As Variant
    monthLabels = Array("January", "February", "March", "April", "May", "June", _
                        "July", "August", "September", "October", "November", "December", "YTD Avg")

    For i = 0 To UBound(monthLabels)
        KPIsWS.Cells(10, i + 2).Value = monthLabels(i)
        KPIsWS.Cells(18, i + 2).Value = monthLabels(i)
    Next i

    ' Apply formatting to EchoKPIs
    With KPIsWS
        Dim lastKPICol As Long, lastKPIRow As Long
        lastKPICol = .Cells(1, .Columns.count).End(xlToLeft).Column
        lastKPIRow = .Cells(.Rows.count, 1).End(xlUp).Row

        ' Outline all used cells
        .Range(.Cells(1, 1), .Cells(lastKPIRow, lastKPICol)).Borders.LineStyle = xlContinuous

        ' Header row color
        .Range(.Cells(1, 1), .Cells(1, lastKPICol)).Interior.Color = RGB(77, 147, 217)

        ' Month label rows
        .Range(.Cells(10, 2), .Cells(10, lastKPICol)).Interior.Color = RGB(77, 147, 217)
        .Range(.Cells(18, 2), .Cells(18, lastKPICol)).Interior.Color = RGB(77, 147, 217)

        ' Column A formatting
        Dim rKPI As Long
        For rKPI = 1 To lastKPIRow
            Select Case rKPI
                Case 1, 10, 18
                    .Cells(rKPI, 1).Interior.Color = RGB(77, 147, 217)
                Case Else
                    .Cells(rKPI, 1).Interior.Color = RGB(166, 201, 236)
            End Select
        Next rKPI

        ' Numeric cells
        Dim cKPI As Long
        For rKPI = 2 To lastKPIRow
            For cKPI = 2 To lastKPICol
                If IsNumeric(.Cells(rKPI, cKPI).Value) Then
                    .Cells(rKPI, cKPI).Interior.Color = RGB(218, 233, 248)
                End If
            Next cKPI
        Next rKPI

        ' Auto-size columns
        .Columns("A:" & Split(.Cells(1, lastKPICol).Address(False, False), "1")(0)).AutoFit
    End With

    ' -------------------------
    ' FORMATTING FOR MONTHLYCOUNTS
    ' -------------------------

    With countsWS
        Dim lastCountCol As Long, lastRowCount As Long
        lastCountCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        lastRowCount = .Cells(.Rows.count, "A").End(xlUp).Row

        ' Outline all used cells
        .Range(.Cells(1, 1), .Cells(lastRowCount, lastCountCol)).Borders.LineStyle = xlContinuous

        ' Header row color
        .Range(.Cells(1, 1), .Cells(1, lastCountCol)).Interior.Color = RGB(77, 147, 217)

        ' Numeric cells
        Dim rCount As Long, cCount As Long
        For rCount = 2 To lastRowCount
            For cCount = 2 To lastCountCol
                If IsNumeric(.Cells(rCount, cCount).Value) Then
                    .Cells(rCount, cCount).Interior.Color = RGB(218, 233, 248)
                End If
            Next cCount
        Next rCount

        ' Auto-size columns
        .Columns("A:" & Split(.Cells(1, lastCountCol).Address(False, False), "1")(0)).AutoFit
    End With

End Sub


