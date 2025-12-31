Attribute VB_Name = "Despatch_Instruction_Comments"
Option Explicit

Public Sub ProcessDispatchInstructions()
    Dim priorScreenUpdating As Boolean
    Dim priorDisplayAlerts As Boolean
    Dim priorEnableEvents As Boolean

    priorScreenUpdating = Application.ScreenUpdating
    priorDisplayAlerts = Application.DisplayAlerts
    priorEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanFail

    Dim sourceFilePath As Variant
    sourceFilePath = GetSourceFilePath()
    If sourceFilePath = False Then
        GoTo CleanExit
    End If

    Application.DisplayAlerts = False

    Dim sourceWb As Workbook
    Set sourceWb = Application.Workbooks.Open(Filename:=sourceFilePath, ReadOnly:=True)

    Application.DisplayAlerts = priorDisplayAlerts

    Dim sourceWs As Worksheet
    Set sourceWs = sourceWb.Worksheets(1)

    Dim columnMap As Object
    Set columnMap = MapColumns(sourceWs)
    If columnMap Is Nothing Then
        GoTo CleanExit
    End If

    Dim dateMap As Object
    Set dateMap = CreateObject("Scripting.Dictionary")

    Dim dispatchData As Collection
    Set dispatchData = CollectDispatchData(sourceWs, columnMap, dateMap)

    Dim outputWb As Workbook
    Set outputWb = Application.Workbooks.Add

    Dim outputWs As Worksheet
    Set outputWs = outputWb.Worksheets(1)
    outputWs.Name = "Dispatch Timeline"

    Dim timelineMap As Object
    Set timelineMap = GenerateTimeline(outputWs, dateMap)

    Dim remarksWritten As Long
    remarksWritten = PlaceRemarksAndFormat(outputWs, dispatchData, timelineMap)

    outputWs.Columns.AutoFit
    outputWs.Rows.AutoFit

    If remarksWritten = 0 Then
        MsgBox "No dispatch instructions matched the filtering criteria.", vbInformation
    Else
        MsgBox "Dispatch instruction report has been successfully generated.", vbInformation
    End If

CleanExit:
    On Error Resume Next
    If Not sourceWb Is Nothing Then
        sourceWb.Close SaveChanges:=False
    End If

    Application.DisplayAlerts = priorDisplayAlerts
    Application.ScreenUpdating = priorScreenUpdating
    Application.EnableEvents = priorEnableEvents
    On Error GoTo 0
    Exit Sub

CleanFail:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Function GetSourceFilePath() As Variant
    Dim dialog As FileDialog
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)

    With dialog
        .Title = "Please select the Dispatch Instruction Report file"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm"
        If .Show <> -1 Then
            GetSourceFilePath = False
            Exit Function
        End If
        GetSourceFilePath = .SelectedItems(1)
    End With
End Function

Private Function NormalizeHeader(ByVal value As Variant) As String
    Dim normalized As String
    normalized = CStr(value)
    normalized = Replace(normalized, vbCr, " ")
    normalized = Replace(normalized, vbLf, " ")
    normalized = LCase(Trim(normalized))

    Do While InStr(normalized, "  ") > 0
        normalized = Replace(normalized, "  ", " ")
    Loop

    NormalizeHeader = normalized
End Function

Private Function MapColumns(ws As Worksheet) As Object
    Dim columnMap As Object
    Set columnMap = CreateObject("Scripting.Dictionary")

    Dim logicalFields As Object
    Set logicalFields = CreateObject("Scripting.Dictionary")
    logicalFields("NotificationDateTime") = Array("notification date & time", "notification time", "notification date")
    logicalFields("TargetTime") = Array("target date & time", "target time")
    logicalFields("TargetDemand") = Array("target demand", "demand", "mw")
    logicalFields("ActualComplianceTime") = Array("actual date & time", "actual compliance", "actual time")
    logicalFields("DemandType") = Array("demand type", "instruction type", "load type")

    Dim headerRow As Long
    headerRow = 1

    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim col As Long
    Dim headerText As String
    Dim key As Variant
    Dim variantName As Variant

    For Each key In logicalFields.Keys
        columnMap(key) = 0
    Next key

    For col = 1 To lastCol
        headerText = NormalizeHeader(ws.Cells(headerRow, col).Value)
        If Len(headerText) > 0 Then
            For Each key In logicalFields.Keys
                If columnMap(key) = 0 Then
                    For Each variantName In logicalFields(key)
                        If headerText = variantName Then
                            columnMap(key) = col
                            Exit For
                        End If
                    Next variantName
                End If
            Next key
        End If
    Next col

    For Each key In logicalFields.Keys
        If columnMap(key) = 0 Then
            MsgBox "Error: Required column for '" & key & "' could not be found.", vbCritical
            Set MapColumns = Nothing
            Exit Function
        End If
    Next headerName

    If missingHeaders.Count > 0 Then
        Dim message As String
        message = "Error: Required column(s) missing from row 1:" & vbLf
        For Each headerName In missingHeaders
            message = message & "- " & headerName & vbLf
        Next headerName
        MsgBox message, vbCritical
        Set MapColumns = Nothing
        Exit Function
    End If

    Set MapColumns = columnMap
End Function

Private Function TryParseDateTime(ByVal value As Variant, ByRef result As Date) As Boolean
    Dim textValue As String
    textValue = Trim(CStr(value))
    If Len(textValue) = 0 Then Exit Function

    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")

    If IsDate(textValue) Then
        result = CDate(textValue)
        TryParseDateTime = True
        Exit Function
    End If

    Dim parts As Variant
    parts = Split(textValue, " ")
    If UBound(parts) >= 1 Then
        If IsDate(parts(0)) And IsDate(parts(1)) Then
            result = DateValue(parts(0)) + TimeValue(parts(1))
            TryParseDateTime = True
        End If
    End If
End Function

Private Function TryParseDemand(ByVal value As Variant, ByRef result As Double) As Boolean
    Dim textValue As String
    textValue = Trim(CStr(value))
    If Len(textValue) = 0 Then Exit Function

    textValue = Replace(textValue, ",", "")
    textValue = Replace(textValue, " ", "")

    If IsNumeric(textValue) Then
        result = CDbl(textValue)
        TryParseDemand = True
    End If
End Function

Private Function CollectDispatchData(ws As Worksheet, columnMap As Object, dateMap As Object) As Collection
    Dim results As New Collection

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        Dim demandType As String
        demandType = LCase(Trim(CStr(ws.Cells(rowIndex, columnMap("DemandType")).Value)))

        If InStr(1, demandType, "increase", vbTextCompare) = 0 And InStr(1, demandType, "decrease", vbTextCompare) = 0 Then
            GoTo NextRow
        End If

        Dim notificationDateTime As Date
        Dim targetTime As Date
        Dim actualComplianceTime As Date
        Dim targetDemand As Double

        If Not TryParseDateTime(ws.Cells(rowIndex, columnMap("NotificationDateTime")).Value, notificationDateTime) Then GoTo NextRow

        dateMap(Format(Int(notificationDateTime), "yyyy-mm-dd")) = True

        If Not TryParseDateTime(ws.Cells(rowIndex, columnMap("TargetTime")).Value, targetTime) Then GoTo NextRow
        If Not TryParseDateTime(ws.Cells(rowIndex, columnMap("ActualComplianceTime")).Value, actualComplianceTime) Then GoTo NextRow
        If Not TryParseDemand(ws.Cells(rowIndex, columnMap("TargetDemand")).Value, targetDemand) Then GoTo NextRow

        Dim dataRow As Object
        Set dataRow = CreateObject("Scripting.Dictionary")
        dataRow("NotificationDateTime") = notificationDateTime
        dataRow("TargetTime") = targetTime
        dataRow("ActualComplianceTime") = actualComplianceTime
        dataRow("TargetDemand") = targetDemand

        results.Add dataRow

NextRow:
    Next rowIndex

    Set CollectDispatchData = results
End Function

Private Function GenerateTimeline(ws As Worksheet, dateMap As Object) As Object
    Dim dateRowMap As Object
    Set dateRowMap = CreateObject("Scripting.Dictionary")

    ws.Columns("A").NumberFormat = "@"
    ws.Columns("B").NumberFormat = "@"

    If dateMap.Count = 0 Then
        Set GenerateTimeline = dateRowMap
        Exit Function
    End If

    Dim sortedDates As Object
    Set sortedDates = CreateObject("System.Collections.ArrayList")

    Dim d As Variant
    For Each d In dateMap.Keys
        sortedDates.Add DateValue(d)
    Next d
    sortedDates.Sort

    Dim currentRow As Long
    currentRow = 1

    Dim d As Variant
    For Each d In sortedDates
        ws.Cells(currentRow, 1).Value = Format(d, "dd-mmm-yyyy")
        With ws.Cells(currentRow, 1)
            .Font.Bold = True
            .Interior.Color = RGB(255, 255, 200)
        End With

        dateRowMap(Format(d, "yyyy-mm-dd")) = currentRow
        currentRow = currentRow + 1

        Dim h As Long
        For h = 0 To 23
            Dim startLabel As String
            Dim endLabel As String
            startLabel = Format(TimeSerial(h, 0, 0), "hh:mm")
            If h = 23 Then
                endLabel = "24:00"
            Else
                endLabel = Format(TimeSerial(h + 1, 0, 0), "hh:mm")
            End If
            ws.Cells(currentRow, 1).Value = startLabel & " – " & endLabel
            currentRow = currentRow + 1
        Next h
    Next d

    Set GenerateTimeline = dateRowMap
End Function

Private Function PlaceRemarksAndFormat(ws As Worksheet, data As Collection, dateRowMap As Object) As Long
    Dim remarksWritten As Long
    remarksWritten = 0

    Dim item As Object
    For Each item In data
        Dim dateKey As String
        dateKey = Format(Int(item("NotificationDateTime")), "yyyy-mm-dd")

        If Not dateRowMap.Exists(dateKey) Then
            GoTo NextItem
        End If

        Dim hourIndex As Long
        hourIndex = Hour(item("NotificationDateTime"))

        Dim targetRow As Long
        targetRow = dateRowMap(dateKey) + 1 + hourIndex

        Dim targetCell As Range
        Set targetCell = ws.Cells(targetRow, 2)

        Dim remark As String
        remark = BuildRemark(item)

        Dim remarkStart As Long
        If Len(targetCell.Value) > 0 Then
            targetCell.Value = targetCell.Value & vbLf & remark
            remarkStart = Len(targetCell.Value) - Len(remark) + 1
        Else
            targetCell.Value = remark
            remarkStart = 1
        End If

        ApplyRemarkFormatting targetCell, item, remark, remarkStart
        remarksWritten = remarksWritten + 1

NextItem:
    Next item

    PlaceRemarksAndFormat = remarksWritten
End Function

Private Function BuildRemark(item As Object) As String
    Dim notifTime As String
    notifTime = "Notification Time: " & Format(item("NotificationDateTime"), "hh:mm")

    Dim targetTime As String
    If Int(item("TargetTime")) <> Int(item("NotificationDateTime")) Then
        targetTime = "Target Time: " & Format(item("TargetTime"), "hh:mm") & " (" & Format(item("TargetTime"), "dd.mmm.yy") & ")"
    Else
        targetTime = "Target Time: " & Format(item("TargetTime"), "hh:mm")
    End If

    Dim targetDemand As String
    If item("TargetDemand") > 320 Then
        targetDemand = "Target Demand: FCBL"
    Else
        targetDemand = "Target Demand: " & Format(item("TargetDemand"), "0.00")
    End If

    Dim actualTime As String
    If Int(item("ActualComplianceTime")) <> Int(item("NotificationDateTime")) Then
        actualTime = "Actual Compliance: " & Format(item("ActualComplianceTime"), "hh:mm") & " (" & Format(item("ActualComplianceTime"), "dd.mmm.yy") & ")"
    Else
        actualTime = "Actual Compliance: " & Format(item("ActualComplianceTime"), "hh:mm")
    End If

    BuildRemark = notifTime & vbLf & targetTime & "; " & targetDemand & vbLf & actualTime
End Function

Private Sub ApplyRemarkFormatting(targetCell As Range, item As Object, remark As String, remarkStart As Long)
    Dim localFcblPos As Long
    localFcblPos = InStr(1, remark, "FCBL", vbTextCompare)
    If localFcblPos > 0 Then
        With targetCell.Characters(remarkStart + localFcblPos - 1, 4).Font
            .Bold = True
            .Color = vbBlue
        End With
    End If

    If item("ActualComplianceTime") > item("TargetTime") Then
        Dim localActualStart As Long
        localActualStart = InStr(1, remark, "Actual Compliance:", vbTextCompare)
        If localActualStart > 0 Then
            Dim localActualEnd As Long
            localActualEnd = InStr(localActualStart, remark, vbLf)
            Dim lineLength As Long
            If localActualEnd > 0 Then
                lineLength = localActualEnd - localActualStart
            Else
                lineLength = Len(remark) - localActualStart + 1
            End If

            If lineLength > 0 Then
                With targetCell.Characters(remarkStart + localActualStart - 1, lineLength).Font
                    .Bold = True
                    .Color = vbRed
                End With
            End If
        End If
    End If
End Sub
