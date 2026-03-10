' -----------------------------------------------------------------------------
' Project: Automated PDF Generator from Excel
' Author: Jenny Marchioro (Fixed Version)
' Description:
'   Legge i dati da Excel, apre template Word, sostituisce i placeholder
'   e salva in PDF. Include gestione errori e pulizia processi.
'
' Setup Colonne Excel:
'   A: Company Name | B: Code | C: Version | D: Email | E: Date
' -----------------------------------------------------------------------------

Sub GeneratePDFs()
    Dim wdApp As Object, wdDoc As Object
    Dim i As Long
    Dim templatePath As String, outputFolder As String, pdfFileName As String
    Dim recordCode As String, companyName As String, versionLabel As String
    Dim emailAddress As String, recordDate As String
    Dim basePath As String
    Dim findRange As Object
    
    On Error GoTo ErrorHandler

    ' Percorso base (cartella del file Excel)
    basePath = ThisWorkbook.Path
    outputFolder = basePath & "\Generated_PDFs\"

    ' Crea cartella di output se non esiste
    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder

    ' Avvia Word (nascosto)
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    ' Ciclo sulle righe partendo dalla 2
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        companyName = Trim(Cells(i, 1).Value)         ' Colonna A
        recordCode = Trim(Cells(i, 2).Value)          ' Colonna B
        versionLabel = UCase(Trim(Cells(i, 3).Value))  ' Colonna C
        emailAddress = Trim(Cells(i, 4).Value)        ' Colonna D
        recordDate = Trim(Cells(i, 5).Value)          ' Colonna E

        ' Selezione template in base alla versione
        Select Case versionLabel
            Case "V1": templatePath = basePath & "\V1.docx"
            Case "V2": templatePath = basePath & "\V2.docx"
            Case "V3": templatePath = basePath & "\V3.docx"
            Case Else
                Debug.Print "Versione non riconosciuta alla riga " & i & ": " & versionLabel
                GoTo NextRow
        End Select

        ' Verifica esistenza template
        If Dir(templatePath) = "" Then
            Debug.Print "Template non trovato: " & templatePath
            GoTo NextRow
        End If

        ' Pulizia caratteri illegali per il nome file
        Dim safeCompanyName As String
        safeCompanyName = companyName
        Dim chars As Variant, c As Long
        chars = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
        For c = LBound(chars) To UBound(chars)
            safeCompanyName = Replace(safeCompanyName, chars(c), "-")
        Next c

        ' Apre il template Word
        Set wdDoc = wdApp.Documents.Open(templatePath, ReadOnly:=True)

        ' Sostituzione Placeholder
        ' Utilizziamo una funzione helper o cicliamo sul contenuto
        Set findRange = wdDoc.Content
        
        ' <<CODE>>
        With findRange.Find
            .Text = "<<CODE>>"
            .Replacement.Text = recordCode
            .Execute Replace:=2 ' wdReplaceAll
        End With
        
        ' <<COMPANY>>
        Set findRange = wdDoc.Content
        With findRange.Find
            .Text = "<<COMPANY>>"
            .Replacement.Text = companyName
            .Execute Replace:=2
        End With

        ' <<EMAIL>>
        Set findRange = wdDoc.Content
        With findRange.Find
            .Text = "<<EMAIL>>"
            .Replacement.Text = emailAddress
            .Execute Replace:=2
        End With

        ' <<DATE>>
        Set findRange = wdDoc.Content
        With findRange.Find
            .Text = "<<DATE>>"
            .Replacement.Text = recordDate
            .Execute Replace:=2
        End With

        ' Esporta come PDF
        pdfFileName = outputFolder & recordCode & "_" & safeCompanyName & ".pdf"
        wdDoc.ExportAsFixedFormat OutputFileName:=pdfFileName, ExportFormat:=17 ' 17 = wdExportFormatPDF
        
        ' Chiude senza salvare modifiche al template
        wdDoc.Close False
        Set wdDoc = Nothing

NextRow:
    Next i

CleanUp:
    If Not wdApp Is Nothing Then
        wdApp.Quit
        Set wdApp = Nothing
    End If
    MsgBox "Processo completato. PDF generati in: " & outputFolder, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Errore alla riga " & i & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub
