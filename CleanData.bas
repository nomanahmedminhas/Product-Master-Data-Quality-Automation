Attribute VB_Name = "Module1"

Sub RefreshCleanMasterData()

    Dim wsRaw As Worksheet
    Dim wsClean As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set wsRaw = ThisWorkbook.Worksheets("Raw Product Data")
    Set wsClean = ThisWorkbook.Worksheets("Clean Master Data")

    'Find the last row containing a SKU
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, "A").End(xlUp).Row

    'Clear previous values without removing formatting
    wsClean.Range("A:G").ClearContents

    'Copy values only, preserving destination formatting
    wsClean.Range("A1:G" & lastRow).Value = _
    wsRaw.Range("A1:G" & lastRow).Value
    
    'Clean text and standardize categories
    For r = 2 To lastRow

        wsClean.Cells(r, "A").Value = _
            UCase(Trim(wsClean.Cells(r, "A").Value))

        wsClean.Cells(r, "C").Value = _
            Trim(wsClean.Cells(r, "C").Value)

        Select Case LCase(wsClean.Cells(r, "C").Value)
            Case "electronic", "electronics"
                wsClean.Cells(r, "C").Value = "Electronics"
            Case "office"
                wsClean.Cells(r, "C").Value = "Office"
            Case "home"
                wsClean.Cells(r, "C").Value = "Home"
        End Select

'Standardize product names
Select Case LCase(Trim(wsClean.Cells(r, "B").Value))
    Case "usb c hub", "usb-c hub"
        wsClean.Cells(r, "B").Value = "USB-C Hub"
End Select

'Standardize supplier names
Select Case LCase(Trim(wsClean.Cells(r, "E").Value))
    Case "acme"
        wsClean.Cells(r, "E").Value = "Acme"
    Case "northstar"
        wsClean.Cells(r, "E").Value = "Northstar"
    Case "visiontek"
        wsClean.Cells(r, "E").Value = "VisionTek"
    Case "paper co."
        wsClean.Cells(r, "E").Value = "Paper Co."
    Case "bright co"
        wsClean.Cells(r, "E").Value = "Bright Co"
    Case "lift lab"
        wsClean.Cells(r, "E").Value = "Lift Lab"
End Select

'Standardize colors
Select Case LCase(Trim(wsClean.Cells(r, "D").Value))
    Case "black"
        wsClean.Cells(r, "D").Value = "Black"
    Case "blue"
        wsClean.Cells(r, "D").Value = "Blue"
    Case "white"
        wsClean.Cells(r, "D").Value = "White"
    Case "silver"
        wsClean.Cells(r, "D").Value = "Silver"
    Case "grey"
        wsClean.Cells(r, "D").Value = "Grey"
    Case "green"
        wsClean.Cells(r, "D").Value = "Green"
End Select

'Standardize status
Select Case LCase(Trim(wsClean.Cells(r, "G").Value))
    Case "active"
        wsClean.Cells(r, "G").Value = "Active"
    Case "discontinued"
        wsClean.Cells(r, "G").Value = "Discontinued"
End Select

    Next r

'STEP 8: Build Review Queue and handle duplicates

Dim wsReview As Worksheet
Dim seen As Object
Dim reviewRow As Long
Dim cleanLastRow As Long
Dim firstRow As Long
Dim sku As String
Dim fieldName As Variant
Dim c As Long
Dim conflictDetails As String

Set wsReview = ThisWorkbook.Worksheets("Review Queue")
Set seen = CreateObject("Scripting.Dictionary")
seen.CompareMode = vbTextCompare

'Refresh the Review Queue
wsReview.Range("A:E").ClearContents
wsReview.Range("A1:E1").Value = Array( _
    "Source Row", "SKU", "Issue Type", _
    "Details", "Review Status")

reviewRow = 2

'Detect duplicates from bottom to top
For r = lastRow To 2 Step -1

    sku = Trim(CStr(wsClean.Cells(r, "A").Value))

    If sku <> "" Then

        If seen.Exists(sku) Then

            firstRow = seen(sku)
            conflictDetails = ""

            'Compare fields B through G
            For c = 2 To 7

                If StrComp( _
                    Trim(CStr(wsClean.Cells(r, c).Value)), _
                    Trim(CStr(wsClean.Cells(firstRow, c).Value)), _
                    vbTextCompare) <> 0 Then

                    conflictDetails = conflictDetails & _
                        wsClean.Cells(1, c).Value & "; "

                End If

            Next c

            wsReview.Cells(reviewRow, 1).Value = r
            wsReview.Cells(reviewRow, 2).Value = sku
            wsReview.Cells(reviewRow, 3).Value = "Duplicate SKU"

            If conflictDetails = "" Then
                wsReview.Cells(reviewRow, 4).Value = _
                    "Identical standardized record"
                wsReview.Cells(reviewRow, 5).Value = _
                    "Duplicate - Auto Removed"
            Else
                wsReview.Cells(reviewRow, 4).Value = _
                    "Differences in: " & conflictDetails
                wsReview.Cells(reviewRow, 5).Value = _
                    "Pending Review"
            End If

            reviewRow = reviewRow + 1

            'Remove duplicate from provisional master.
            'Any conflicts remain visible in Review Queue.
            wsClean.Range("A" & r & ":G" & r).Delete Shift:=xlUp

            'Adjust stored row numbers after deletion
            Dim key As Variant
            For Each key In seen.Keys
                If seen(key) > r Then
                    seen(key) = seen(key) - 1
                End If
            Next key

        Else
            seen.Add sku, r
        End If

    End If

Next r

'STEP 9: Identify missing required fields

cleanLastRow = wsClean.Cells( _
    wsClean.Rows.Count, "A").End(xlUp).Row

For r = 2 To cleanLastRow

    For c = 1 To 7

        If Trim(CStr(wsClean.Cells(r, c).Value)) = "" Then

            wsReview.Cells(reviewRow, 1).Value = _
                "Master Row " & r

            wsReview.Cells(reviewRow, 2).Value = _
                wsClean.Cells(r, "A").Value

            wsReview.Cells(reviewRow, 3).Value = _
                "Missing Data"

            wsReview.Cells(reviewRow, 4).Value = _
                "Missing field: " & _
                wsClean.Cells(1, c).Value

            wsReview.Cells(reviewRow, 5).Value = _
                "Pending Verification"

            reviewRow = reviewRow + 1

        End If

    Next c

Next r

    'Maintain the Clean Master Data Excel Table
    Dim tbl As ListObject
    Dim finalRow As Long
    
    finalRow = wsClean.Cells(wsClean.Rows.Count, "A").End(xlUp).Row
    
    On Error Resume Next
    Set tbl = wsClean.ListObjects("tblCleanMaster")
    On Error GoTo 0
    
    If tbl Is Nothing Then
    
        Set tbl = wsClean.ListObjects.Add( _
            xlSrcRange, _
            wsClean.Range("A1:G" & finalRow), _
            , xlYes)
    
        tbl.Name = "tblCleanMaster"
    
    Else
    
        tbl.Resize wsClean.Range("A1:G" & finalRow)
    
    End If
    
    tbl.TableStyle = "TableStyleMedium2"
    tbl.ShowTableStyleRowStripes = True

    Call GenerateExceptionLog

    MsgBox "Initial data cleaning completed.", vbInformation

End Sub


Sub GenerateExceptionLog()

    Dim wsRaw As Worksheet
    Dim wsLog As Worksheet
    Dim wsReview As Worksheet
    Dim lastRow As Long
    Dim logRow As Long
    Dim r As Long, c As Long
    Dim rawValue As String
    Dim cleanValue As String
    Dim sku As String
    Dim issue As String
    Dim actionTaken As String
    Dim resolution As String
    Dim fieldName As String
    Dim reviewLast As Long

    Set wsRaw = ThisWorkbook.Worksheets("Raw Product Data")
    Set wsLog = ThisWorkbook.Worksheets("Exception Log")
    Set wsReview = ThisWorkbook.Worksheets("Review Queue")

    wsLog.Range("A:F").ClearContents

    wsLog.Range("A1:F1").Value = Array( _
        "Original Row", "SKU", "Field", _
        "Issue Found", "Action Taken", _
        "Resolution Status")

    logRow = 2

    lastRow = wsRaw.Cells(wsRaw.Rows.Count, "A").End(xlUp).Row

    'Compare each raw field with the approved standardization
    'rules. Preserve original source row references.
    For r = 2 To lastRow

        sku = UCase(Trim(CStr(wsRaw.Cells(r, 1).Value)))

        For c = 1 To 7

            fieldName = CStr(wsRaw.Cells(1, c).Value)
            rawValue = CStr(wsRaw.Cells(r, c).Value)
            cleanValue = Trim(rawValue)

            Select Case c

                Case 1
                    cleanValue = UCase(cleanValue)

                Case 2
                    Select Case LCase(cleanValue)
                        Case "usb c hub", "usb-c hub"
                            cleanValue = "USB-C Hub"
                    End Select

                Case 3
                    Select Case LCase(cleanValue)
                        Case "electronic", "electronics"
                            cleanValue = "Electronics"
                        Case "office"
                            cleanValue = "Office"
                        Case "home"
                            cleanValue = "Home"
                    End Select

                Case 4
                    Select Case LCase(cleanValue)
                        Case "black": cleanValue = "Black"
                        Case "blue": cleanValue = "Blue"
                        Case "white": cleanValue = "White"
                        Case "silver": cleanValue = "Silver"
                        Case "grey": cleanValue = "Grey"
                        Case "green": cleanValue = "Green"
                    End Select

                Case 5
                    Select Case LCase(cleanValue)
                        Case "acme": cleanValue = "Acme"
                        Case "northstar": cleanValue = "Northstar"
                        Case "visiontek": cleanValue = "VisionTek"
                        Case "paper co.": cleanValue = "Paper Co."
                        Case "bright co": cleanValue = "Bright Co"
                        Case "lift lab": cleanValue = "Lift Lab"
                    End Select

                Case 7
                    Select Case LCase(cleanValue)
                        Case "active": cleanValue = "Active"
                        Case "discontinued"
                            cleanValue = "Discontinued"
                    End Select

            End Select

            If Trim(rawValue) = "" Then

                issue = "Missing Data"
                actionTaken = "Awaiting source verification"
                resolution = "Pending"

            ElseIf StrComp(rawValue, cleanValue, _
                           vbBinaryCompare) <> 0 Then

                issue = "Inconsistent value: " & rawValue
                actionTaken = "Standardized to: " & cleanValue
                resolution = "Resolved"

            Else
                issue = ""
            End If

            If issue <> "" Then

                wsLog.Cells(logRow, 1).Value = r
                wsLog.Cells(logRow, 2).Value = sku
                wsLog.Cells(logRow, 3).Value = fieldName
                wsLog.Cells(logRow, 4).Value = issue
                wsLog.Cells(logRow, 5).Value = actionTaken
                wsLog.Cells(logRow, 6).Value = resolution

                logRow = logRow + 1

            End If

        Next c

    Next r

    'Append duplicate exceptions from Review Queue
    reviewLast = wsReview.Cells( _
        wsReview.Rows.Count, "B").End(xlUp).Row

    For r = 2 To reviewLast

        If wsReview.Cells(r, 3).Value = "Duplicate SKU" Then

            wsLog.Cells(logRow, 1).Value = _
                wsReview.Cells(r, 1).Value

            wsLog.Cells(logRow, 2).Value = _
                wsReview.Cells(r, 2).Value

            wsLog.Cells(logRow, 3).Value = "SKU"
            wsLog.Cells(logRow, 4).Value = "Duplicate record"
            wsLog.Cells(logRow, 5).Value = _
                wsReview.Cells(r, 4).Value

            If wsReview.Cells(r, 5).Value = _
                "Duplicate - Auto Removed" Then
                wsLog.Cells(logRow, 6).Value = "Resolved"
            Else
                wsLog.Cells(logRow, 6).Value = "Pending"
            End If

            logRow = logRow + 1

        End If

    Next r

    MsgBox "Exception Log generated: " & _
           (logRow - 2) & " entries.", vbInformation

End Sub


