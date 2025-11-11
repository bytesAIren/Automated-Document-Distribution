' -----------------------------------------------------------------------------
' Project: Automated PDF Generator from Excel
' Author: Jenny Marchioro (vibecoding with ChatGPT)
' Description:
'   This VBA macro reads data from an Excel worksheet, opens corresponding Word
'   templates, replaces placeholders with values from the sheet, and exports
'   each filled document as a PDF file. Designed for non-developers who want to
'   automate repetitive document creation tasks.
'
' How to use:
'   1. Place this VBA code inside an Excel module.
'   2. In your Excel file, create a table with these columns:
'        A: Company Name
'        B: Unique Code or ID
'        C: Version (e.g., V1, V2, V3)
'   3. Store your Word templates (V1.docx, V2.docx, V3.docx) in the same folder
'      as the Excel file. Inside each template, include placeholders like:
'        <<CODE>>  and  <<COMPANY>>
'   4. Run the macro. It will create a folder called “Generated_PDFs” and save
'      one PDF for each Excel row.
'
' Notes:
'   - This macro runs locally and does not send any data online.
'   - Works with Microsoft Word and Excel on Windows.
'   - Edit placeholders or variable names as needed for your workflow.
' -----------------------------------------------------------------------------

Sub GeneratePDFs()
    Dim wdApp As Object, wdDoc As Object
    Dim i As Long
    Dim templatePath As String, outputFolder As String, pdfFileName As String
    Dim recordCode As String, companyName As String, versionLabel As String
    Dim basePath As String
    Dim findCode As Object, findCompany As Object

    ' Base path (folder where the Excel file is located)
    basePath = ThisWorkbook.Path
    outputFolder = basePath & "\Generated_PDFs\"

    ' Create output folder if it doesn't exist
    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder

    ' Start Word (hidden)
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    ' Loop through all rows starting from row 2
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        companyName = Trim(Cells(i, 1).Value)       ' Column A - Company Name
        recordCode = Trim(Cells(i, 2).Value)        ' Column B - Unique Code / ID
        versionLabel = UCase(Trim(Cells(i, 3).Value)) ' Column C - Version or Template Type

        ' Select the correct Word template based on version
        Select Case versionLabel
            Case "V1": templatePath = basePath & "\V1.docx"
            Case "V2": templatePath = basePath & "\V2.docx"
            Case "V3": templatePath = basePath & "\V3.docx"
            Case Else
                MsgBox "Unrecognized version at row " & i & ": " & versionLabel, vbExclamation
                GoTo NextRow
        End Select

        ' Clean illegal characters from company name (for file name safety)
        companyName = Replace(companyName, "/", "-")
        companyName = Replace(companyName, "\", "-")
        companyName = Replace(companyName, ":", "-")
        companyName = Replace(companyName, "*", "")
        companyName = Replace(companyName, "?", "")
        companyName = Replace(companyName, """", "")
        companyName = Replace(companyName, "<", "")
        companyName = Replace(companyName, ">", "")
        companyName = Replace(companyName, "|", "")

        ' Open Word template
        Set wdDoc = wdApp.Documents.Open(templatePath)

        ' Replace placeholders with Excel data values
        Set findCode = wdDoc.Content
        With findCode.Find
            .Text = "<<CODE>>"
            .Replacement.Text = recordCode
            .Execute Replace:=2
        End With

        Set findCompany = wdDoc.Content
        With findCompany.Find
            .Text = "<<COMPANY>>"
            .Replacement.Text = companyName
            .Execute Replace:=2
        End With

        ' Save the filled document as PDF
        pdfFileName = outputFolder & recordCode & "_" & companyName & ".pdf"
        wdDoc.ExportAsFixedFormat OutputFileName:=pdfFileName, ExportFormat:=17
        wdDoc.Close False

NextRow:
    Next i

    ' Close Word instance
    wdApp.Quit

    ' Confirmation message
    MsgBox "PDFs successfully generated in: " & outputFolder, vbInformation
End Sub
