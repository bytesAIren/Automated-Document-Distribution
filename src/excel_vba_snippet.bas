' -----------------------------------------------------------------------------
' Project:     Automated PDF Generator from Excel
' Description: Reads data from Excel, opens a Word template, replaces
'              placeholders, and saves output as PDF.
'              Includes error handling, duplicate prevention, process cleanup,
'              and visual feedback via status bar.
'
' Excel Column Setup:
'   A: Company Name | B: Record Code | C: Version (V1/V2/V3) | D: Email | E: Date | F: Status | G: PDF Filename (auto-filled)
'
' Word Template Placeholders:
'   <<CODE>> | <<COMPANY>> | <<EMAIL>> | <<DATE>>
'
' Template files expected in the same folder as this workbook:
'   V1.docx, V2.docx, V3.docx
'
' Output PDFs are saved to: <WorkbookFolder>\Generated_PDFs\
' -----------------------------------------------------------------------------

Option Explicit

Sub GeneratePDFs()

    ' -------------------------------------------------------------------------
    ' DECLARATIONS
    ' -------------------------------------------------------------------------
    Dim wdApp       As Object
    Dim wdDoc       As Object
    Dim i           As Long
    Dim lastRow     As Long

    Dim basePath        As String
    Dim outputFolder    As String
    Dim templatePath    As String
    Dim pdfFilePath     As String

    Dim companyName     As String
    Dim safeCompanyName As String
    Dim recordCode      As String
    Dim versionLabel    As String
    Dim emailAddress    As String
    Dim recordDate      As String
    Dim statusCell      As String

    Dim pdfFileName     As String
    Dim illegalChars    As Variant
    Dim c               As Long
    Dim successCount    As Long
    Dim skipCount       As Long
    Dim errorCount      As Long

    On Error GoTo ErrorHandler

    ' -------------------------------------------------------------------------
    ' SETUP
    ' -------------------------------------------------------------------------
    basePath     = ThisWorkbook.Path
    outputFolder = basePath & "\Generated_PDFs\"
    illegalChars = Array("/", "\", ":", "*", "?", """", "<", ">", "|", "'")

    ' Create output folder if it doesn't exist
    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder

    ' Find last row with data in column A
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data found. Please check your spreadsheet.", vbExclamation
        Exit Sub
    End If

    ' Start Word (hidden)
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    successCount = 0
    skipCount    = 0
    errorCount   = 0

    ' -------------------------------------------------------------------------
    ' MAIN LOOP
    ' -------------------------------------------------------------------------
    For i = 2 To lastRow

        Application.StatusBar = "Processing row " & i & " of " & lastRow & "..."

        ' Read cell values
        companyName  = Trim(Cells(i, 1).Value)
        recordCode   = Trim(Cells(i, 2).Value)
        versionLabel = UCase(Trim(Cells(i, 3).Value))
        emailAddress = Trim(Cells(i, 4).Value)
        recordDate   = Trim(Cells(i, 5).Value)
        statusCell   = Trim(Cells(i, 6).Value)

        ' --- Skip rows with missing required data ---
        If companyName = "" Or recordCode = "" Then
            Debug.Print "Row " & i & " skipped: missing Company Name or Record Code."
            skipCount = skipCount + 1
            GoTo NextRow
        End If

        ' --- Skip rows already marked as done ---
        If InStr(1, statusCell, "Done", vbTextCompare) > 0 Then
            Debug.Print "Row " & i & " skipped: already processed (" & statusCell & ")."
            skipCount = skipCount + 1
            GoTo NextRow
        End If

        ' --- Select template based on version ---
        Select Case versionLabel
            Case "V1": templatePath = basePath & "\V1.docx"
            Case "V2": templatePath = basePath & "\V2.docx"
            Case "V3": templatePath = basePath & "\V3.docx"
            Case Else
                Debug.Print "Row " & i & " skipped: unrecognised version '" & versionLabel & "'."
                Cells(i, 6).Value = "ERROR - Unknown version: " & versionLabel
                skipCount = skipCount + 1
                GoTo NextRow
        End Select

        ' --- Verify template exists ---
        If Dir(templatePath) = "" Then
            Debug.Print "Row " & i & ": template not found at " & templatePath
            Cells(i, 6).Value = "ERROR - Template not found: " & versionLabel
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        ' --- Sanitise company name for use in file name ---
        safeCompanyName = companyName
        For c = LBound(illegalChars) To UBound(illegalChars)
            safeCompanyName = Replace(safeCompanyName, illegalChars(c), "-")
        Next c

        ' Build canonical PDF filename and write it to column G (source of truth for GAS)
        pdfFileName = recordCode & "_" & safeCompanyName & ".pdf"
        Cells(i, 7).Value = pdfFileName

        ' Build full output path
        pdfFilePath = outputFolder & pdfFileName

        ' --- Open Word template (read-write, on a fresh copy each time) ---
        Set wdDoc = Nothing
        On Error Resume Next
        Set wdDoc = wdApp.Documents.Open(templatePath, ReadOnly:=False)
        On Error GoTo ErrorHandler

        If wdDoc Is Nothing Then
            Debug.Print "Row " & i & ": failed to open template " & templatePath
            Cells(i, 6).Value = "ERROR - Could not open template"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        ' --- Replace placeholders ---
        Call ReplacePlaceholder(wdDoc, "<<CODE>>",    recordCode)
        Call ReplacePlaceholder(wdDoc, "<<COMPANY>>", companyName)
        Call ReplacePlaceholder(wdDoc, "<<EMAIL>>",   emailAddress)
        Call ReplacePlaceholder(wdDoc, "<<DATE>>",    recordDate)

        ' --- Export to PDF ---
        On Error Resume Next
        wdDoc.ExportAsFixedFormat OutputFileName:=pdfFilePath, ExportFormat:=17 ' 17 = wdExportFormatPDF
        If Err.Number <> 0 Then
            Debug.Print "Row " & i & ": PDF export failed - " & Err.Description
            Cells(i, 6).Value = "ERROR - PDF export failed"
            Err.Clear
            wdDoc.Close SaveChanges:=False
            Set wdDoc = Nothing
            errorCount = errorCount + 1
            On Error GoTo ErrorHandler
            GoTo NextRow
        End If
        On Error GoTo ErrorHandler

        ' --- Close document WITHOUT saving changes to the template ---
        wdDoc.Close SaveChanges:=False
        Set wdDoc = Nothing

        ' --- Mark row as done with timestamp ---
        Cells(i, 6).Value = "Done - " & Format(Now(), "yyyy-mm-dd hh:mm")
        successCount = successCount + 1

NextRow:
    Next i

    ' -------------------------------------------------------------------------
    ' CLEANUP & SUMMARY
    ' -------------------------------------------------------------------------
CleanUp:
    Application.StatusBar = False   ' Restore default status bar

    ' Safely close any document still open
    If Not wdDoc Is Nothing Then
        On Error Resume Next
        wdDoc.Close SaveChanges:=False
        Set wdDoc = Nothing
        On Error GoTo 0
    End If

    ' Quit Word
    If Not wdApp Is Nothing Then
        On Error Resume Next
        wdApp.Quit
        Set wdApp = Nothing
        On Error GoTo 0
    End If

    ' Final summary message
    MsgBox "Process complete." & vbCrLf & vbCrLf & _
           "  Successful :  " & successCount & vbCrLf & _
           "  Skipped     :  " & skipCount & vbCrLf & _
           "  Errors       :  " & errorCount & vbCrLf & vbCrLf & _
           "PDFs saved to: " & outputFolder, _
           vbInformation, "PDF Generator - Summary"

    Exit Sub

    ' -------------------------------------------------------------------------
    ' ERROR HANDLER
    ' -------------------------------------------------------------------------
ErrorHandler:
    Dim errMsg As String
    errMsg = "Unexpected error at row " & i & ":" & vbCrLf & _
             "Error " & Err.Number & " - " & Err.Description

    Debug.Print errMsg

    ' Try to log error in the Status column
    On Error Resume Next
    If i >= 2 Then Cells(i, 6).Value = "ERROR - " & Err.Description
    On Error GoTo 0

    errorCount = errorCount + 1
    Resume CleanUp

End Sub


' -----------------------------------------------------------------------------
' HELPER: ReplacePlaceholder
' Replaces all occurrences of a placeholder tag in the entire Word document,
' including headers, footers, and text boxes.
' -----------------------------------------------------------------------------
Private Sub ReplacePlaceholder(ByVal doc As Object, _
                                ByVal placeholder As String, _
                                ByVal replacement As String)

    Dim rng As Object

    ' --- Main body ---
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Text              = placeholder
        .Replacement.Text  = replacement
        .Forward           = True
        .Wrap              = 1          ' wdFindContinue
        .MatchCase         = False
        .MatchWholeWord    = False
        .Execute Replace:=2             ' wdReplaceAll
    End With

    ' --- Headers and Footers (all sections) ---
    Dim sec As Object
    Dim hf  As Object
    For Each sec In doc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then
                Set rng = hf.Range
                With rng.Find
                    .ClearFormatting
                    .Text             = placeholder
                    .Replacement.Text = replacement
                    .Forward          = True
                    .Wrap             = 1
                    .Execute Replace:=2
                End With
            End If
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then
                Set rng = hf.Range
                With rng.Find
                    .ClearFormatting
                    .Text             = placeholder
                    .Replacement.Text = replacement
                    .Forward          = True
                    .Wrap             = 1
                    .Execute Replace:=2
                End With
            End If
        Next hf
    Next sec

    ' --- Text Boxes / Shapes ---
    Dim shp As Object
    For Each shp In doc.Shapes
        On Error Resume Next
        Set rng = shp.TextFrame.TextRange
        If Not rng Is Nothing Then
            With rng.Find
                .ClearFormatting
                .Text             = placeholder
                .Replacement.Text = replacement
                .Forward          = True
                .Wrap             = 1
                .Execute Replace:=2
            End With
        End If
        On Error GoTo 0
    Next shp

End Sub
