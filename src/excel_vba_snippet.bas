Sub GeneraPDF_Trentino_FINAL()
    Dim wdApp As Object, wdDoc As Object
    Dim i As Long
    Dim modello As String, cartellaOutput As String, pdfName As String
    Dim protocollo As String, spettle As String, versione As String
    Dim basePath As String
    Dim rngProt As Object, rngSpettle As Object
    
    ' Percorso base (dove si trova l'Excel)
    basePath = ThisWorkbook.Path
    cartellaOutput = basePath & "\PDF_Generati\"
    If Dir(cartellaOutput, vbDirectory) = "" Then MkDir cartellaOutput
    
    ' Avvia Word
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    
    ' Scorri tutte le righe (dalla 2 in poi)
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        spettle = Trim(Cells(i, 1).Value)       ' Colonna A - SPETTLE
        protocollo = Trim(Cells(i, 2).Value)    ' Colonna B - PROT
        versione = UCase(Trim(Cells(i, 3).Value)) ' Colonna C - Versione
        
        ' Seleziona il modello giusto
        Select Case versione
            Case "V1": modello = basePath & "\V1.docx"
            Case "V2": modello = basePath & "\V2.docx"
            Case "VGEAS": modello = basePath & "\VGeas.docx"
            Case Else
                MsgBox "Versione non riconosciuta alla riga " & i & ": " & versione, vbExclamation
                GoTo Prossimo
        End Select
        
        ' Pulisci nome file da caratteri illegali
        spettle = Replace(spettle, "/", "-")
        spettle = Replace(spettle, "\", "-")
        spettle = Replace(spettle, ":", "-")
        spettle = Replace(spettle, "*", "")
        spettle = Replace(spettle, "?", "")
        spettle = Replace(spettle, """", "")
        spettle = Replace(spettle, "<", "")
        spettle = Replace(spettle, ">", "")
        spettle = Replace(spettle, "|", "")
        
        ' Apri il modello
        Set wdDoc = wdApp.Documents.Open(modello)
        
        ' --- Inserisci i valori nei punti corretti ---
        Set rngProt = wdDoc.Content
        With rngProt.Find
            .Text = "<<PROT>>"
            .Replacement.Text = protocollo
            .Execute Replace:=2
        End With
        
        Set rngSpettle = wdDoc.Content
        With rngSpettle.Find
            .Text = "<<SPETTLE>>"
            .Replacement.Text = spettle
            .Execute Replace:=2
        End With
        ' ----------------------------------------------
        
        ' Salva come PDF
        pdfName = cartellaOutput & protocollo & "_" & spettle & ".pdf"
        wdDoc.ExportAsFixedFormat OutputFileName:=pdfName, ExportFormat:=17
        wdDoc.Close False
        
Prossimo:
    Next i
    
    wdApp.Quit
    MsgBox "PDF generati correttamente in: " & cartellaOutput, vbInformation
End Sub

