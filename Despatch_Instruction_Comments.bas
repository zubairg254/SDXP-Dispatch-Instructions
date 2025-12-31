Attribute VB_Name = "Despatch_Instruction_Comments"
Option Explicit

Public Sub ProcessDispatchInstructions()
    ' Suppress screen updating to speed up macro execution
    Application.ScreenUpdating = False

    ' Prompt user to select the Dispatch Instruction Report
    Dim sourceFilePath As Variant
    sourceFilePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xls), *.xls", _
        Title:="Please select the Dispatch Instruction Report file", _
        MultiSelect:=False)

    ' Check if the user cancelled the dialog
    If sourceFilePath = False Then
        MsgBox "Operation cancelled by the user.", vbInformation
        Exit Sub
    End If

    ' Suppress alerts to prevent the file format mismatch warning
    Application.DisplayAlerts = False

    ' Open the selected workbook in read-only mode
    Dim sourceWb As Workbook
    Set sourceWb = Application.Workbooks.Open(sourceFilePath, ReadOnly:=True)

    ' Re-enable alerts now that the file is open
    Application.DisplayAlerts = True

    ' --- Header Detection and Column Mapping ---
    Dim sourceWs As Worksheet
    Set sourceWs = sourceWb.Worksheets(1) ' Assuming the data is on the first sheet

    Dim headerRow As Long
    headerRow = 1 ' Assume the first row is always the header

    Dim columnMap As Object ' Dictionary
    Set columnMap = MapColumns(sourceWs, headerRow)

    If columnMap Is Nothing Then
        ' The MapColumns function will have already shown a specific error message.
        sourceWb.Close SaveChanges:=False
        Exit Sub
    End If

    ' --- Data Processing and Validation ---
    Dim dispatchData As Collection
    Set dispatchData = ProcessData(sourceWs, columnMap, headerRow)

    If dispatchData.Count = 0 Then
        MsgBox "No valid dispatch instructions found matching the criteria.", vbInformation
        ' Note: The plan is to create an empty report, so we don't exit here.
    Else
        MsgBox dispatchData.Count & " valid dispatch instructions were processed.", vbInformation
    End If

    ' --- Output Workbook Generation ---
    Dim outputWb As Workbook
    Set outputWb = Application.Workbooks.Add
    Dim outputWs As Worksheet
    Set outputWs = outputWb.Worksheets(1)

    GenerateTimeline outputWs, dispatchData

    ' --- Remark Construction and Placement ---
    PlaceRemarksAndFormat outputWs, dispatchData

    ' --- Finalization ---
    outputWs.Columns.AutoFit
    outputWs.Rows.AutoFit

    MsgBox "Dispatch instruction report has been successfully generated.", vbInformation


    ' Clean up
    sourceWb.Close SaveChanges:=False
    Set sourceWb = Nothing

    Application.ScreenUpdating = True
End Sub

Private Function MapColumns(ws As Worksheet, headerRow As Long) As Object
    ' Maps the required column headers to their column index
    Dim columnMap As Object
    Set columnMap = CreateObject("Scripting.Dictionary")

    Dim logicalFields As Object
    Set logicalFields = CreateObject("Scripting.Dictionary")
    logicalFields("NotificationDateTime") = Array("Notification Date & Time", "Notification Time", "Notification Date")
    logicalFields("TargetTime") = Array("Target Date & Time", "Target Time")
    logicalFields("TargetDemand") = Array("Target Demand", "Demand", "MW")
    logicalFields("ActualComplianceTime") = Array("Actual Date & Time", "Actual Compliance", "Actual Time")
    logicalFields("DemandType") = Array("Demand Type", "Instruction Type", "Load Type")

    Dim key As Variant
    Dim variants As Variant
    Dim variantName As Variant
    Dim headerCell As Range
    Dim found As Boolean

    For Each key In logicalFields.Keys
        found = False
        variants = logicalFields(key)
        For Each variantName In variants
            For Each headerCell In ws.Rows(headerRow).Cells
                If LCase(Trim(Replace(headerCell.Value, vbLf, " "))) = LCase(variantName) Then
                    columnMap(key) = headerCell.Column
                    found = True
                    Exit For
                End If
            Next headerCell
            If found Then Exit For
        Next variantName

        If Not found Then
            MsgBox "Error: Required column for '" & key & "' could not be found.", vbCritical
            Set MapColumns = Nothing
            Exit Function
        End If
    Next key

    Set MapColumns = columnMap
End Function

Private Function ProcessData(ws As Worksheet, columnMap As Object, headerRow As Long) As Collection
    Set ProcessData = New Collection

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = headerRow + 1 To lastRow
        Dim demandType As String
        demandType = ws.Cells(i, columnMap("DemandType")).Value

        If InStr(1, demandType, "Increase Load", vbTextCompare) > 0 Or InStr(1, demandType, "Decrease Load", vbTextCompare) > 0 Then
            ' Validation Checks
            If Not IsDate(ws.Cells(i, columnMap("NotificationDateTime")).Value) Then GoTo SkipRow
            If Not IsDate(ws.Cells(i, columnMap("TargetTime")).Value) Then GoTo SkipRow
            If Not IsDate(ws.Cells(i, columnMap("ActualComplianceTime")).Value) Then GoTo SkipRow
            If Not IsNumeric(ws.Cells(i, columnMap("TargetDemand")).Value) Then GoTo SkipRow

            ' If all checks pass, add to collection
            Dim dataRow As Object
            Set dataRow = CreateObject("Scripting.Dictionary")
            dataRow("NotificationDateTime") = CDate(ws.Cells(i, columnMap("NotificationDateTime")).Value)
            dataRow("TargetTime") = CDate(ws.Cells(i, columnMap("TargetTime")).Value)
            dataRow("ActualComplianceTime") = CDate(ws.Cells(i, columnMap("ActualComplianceTime")).Value)
            dataRow("TargetDemand") = CDbl(ws.Cells(i, columnMap("TargetDemand")).Value)

            ProcessData.Add dataRow
        End If
SkipRow:
    Next i
End Function

Private Sub GenerateTimeline(ws As Worksheet, data As Collection)
    ws.Columns("A").NumberFormat = "@" ' Text format for the timeline
    ws.Columns("B").NumberFormat = "@" ' Text format for remarks

    ws.Cells(1, 1).Value = "Date / Time"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 2).Value = "Dispatch Remarks"
    ws.Cells(1, 2).Font.Bold = True

    If data.Count = 0 Then
        ws.Cells(2, 1).Value = "No data matching filtering criteria."
        Exit Sub
    End If

    ' Get unique dates
    Dim uniqueDates As Object
    Set uniqueDates = CreateObject("Scripting.Dictionary")
    Dim item As Object
    For Each item In data
        uniqueDates(Int(item("NotificationDateTime"))) = 1
    Next item

    Dim sortedDates As Object
    Set sortedDates = CreateObject("System.Collections.ArrayList")
    Dim d As Variant
    For Each d In uniqueDates.Keys
        sortedDates.Add d
    Next d
    sortedDates.Sort

    Dim currentRow As Long
    currentRow = 2

    For Each d In sortedDates
        ' Date Header
        ws.Cells(currentRow, 1).Value = Format(d, "dd-mmm-yyyy")
        With ws.Cells(currentRow, 1).Font
            .Bold = True
            .Color = vbBlack
        End With
        ws.Cells(currentRow, 1).Interior.Color = RGB(255, 255, 200)
        currentRow = currentRow + 1

        ' Hourly Slots
        Dim h As Long
        For h = 0 To 23
            Dim endTime As String
            If h = 23 Then
                endTime = "24:00"
            Else
                endTime = Format(TimeSerial(h + 1, 0, 0), "hh:mm")
            End If
            ws.Cells(currentRow + h, 1).Value = Format(TimeSerial(h, 0, 0), "hh:mm") & " – " & endTime
        Next h
        currentRow = currentRow + 24
    Next d
End Sub

Private Sub PlaceRemarksAndFormat(ws As Worksheet, data As Collection)
    If data.Count = 0 Then Exit Sub

    Dim item As Object
    For Each item In data
        Dim remark As String
        remark = BuildRemark(item)

        Dim targetRow As Long
        targetRow = FindRowForRemark(ws, item("NotificationDateTime"))

        If targetRow > 0 Then
            Dim targetCell As Range
            Set targetCell = ws.Cells(targetRow, 2)

            Dim startPos As Long
            If Len(targetCell.Value) > 0 Then
                ' Position after the existing text and the newline character
                startPos = Len(targetCell.Value) + 2
                targetCell.Value = targetCell.Value & vbLf & remark
            Else
                startPos = 1
                targetCell.Value = remark
            End If

            ' --- Apply Formatting to the newly added text block ---

            ' Rule 1: Target Demand > 320 MW
            If item("TargetDemand") > 320 Then
                Dim fcblPos As Long
                fcblPos = InStr(startPos, targetCell.Value, "FCBL")
                If fcblPos > 0 Then
                    With targetCell.Characters(fcblPos, 4).Font
                        .Bold = True
                        .Color = vbBlue
                    End With
                End If
            End If

            ' Rule 2: Compliance Delay
            If item("ActualComplianceTime") > item("TargetTime") Then
                Dim actualLineStart As Long
                actualLineStart = InStr(startPos, targetCell.Value, "Actual Compliance:")

                If actualLineStart > 0 Then
                    ' Find the end of the line for the current remark
                    Dim actualLineEnd As Long
                    actualLineEnd = InStr(actualLineStart, targetCell.Value, vbLf)

                    Dim lineLength As Long
                    If actualLineEnd > 0 And actualLineEnd > actualLineStart Then
                        lineLength = actualLineEnd - actualLineStart
                    Else
                        ' If no newline, it's the last line in the cell
                        lineLength = Len(targetCell.Value) - actualLineStart + 1
                    End If

                    If lineLength > 0 Then
                        With targetCell.Characters(actualLineStart, lineLength).Font
                            .Bold = True
                            .Color = vbRed
                        End With
                    End If
                End If
            End If
        End If
    Next item
End Sub

Private Function BuildRemark(item As Object) As String
    Dim notifTime As String
    notifTime = "Notification Time: " & Format(item("NotificationDateTime"), "hh:mm")

    Dim targetTime As String
    If Int(item("TargetTime")) <> Int(item("NotificationDateTime")) Then
        targetTime = "Target Time: " & Format(item("TargetTime"), "hh:mm (dd.mmm.yy)")
    Else
        targetTime = "Target Time: " & Format(item("TargetTime"), "hh:mm")
    End If

    Dim targetDemand As String
    If item("TargetDemand") > 320 Then
        targetDemand = "Target Demand: FCBL"
    Else
        targetDemand = "Target Demand: " & Format(item("TargetDemand"), "#,##0.00")
    End If

    Dim actualTime As String
    If Int(item("ActualComplianceTime")) <> Int(item("NotificationDateTime")) Then
        actualTime = "Actual Compliance: " & Format(item("ActualComplianceTime"), "hh:mm (dd.mmm.yy)")
    Else
        actualTime = "Actual Compliance: " & Format(item("ActualComplianceTime"), "hh:mm")
    End If

    BuildRemark = notifTime & vbLf & targetTime & "; " & targetDemand & vbLf & actualTime
End Function

Private Function FindRowForRemark(ws As Worksheet, notificationTime As Date) As Long
    Dim searchDate As String
    searchDate = Format(notificationTime, "dd-mmm-yyyy")

    Dim dateCell As Range
    Set dateCell = ws.Columns(1).Find(What:=searchDate, LookIn:=xlValues, LookAt:=xlWhole)

    If Not dateCell Is Nothing Then
        Dim hourSlot As Integer
        hourSlot = Hour(notificationTime)
        FindRowForRemark = dateCell.Row + 1 + hourSlot
    Else
        FindRowForRemark = 0
    End If
End Function
