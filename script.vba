Sub SendConfirmationEmails()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    
    ' Angabe des Blattnamens, in dem die Daten enthalten sind
    Set ws = ThisWorkbook.Sheets("Sendungen")
    
    ' Definieren der letzten Zeile in Spalte A
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Überprüfen, ob Outlook geöffnet ist
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    
    ' Falls Outlook nicht geöffnet ist, öffne es
    If OutApp Is Nothing Then
        Set OutApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    ' Schleife zum Durchgehen der E-Mail-Adressen
    For i = 2 To LastRow ' Beginn ab Zeile 2, da erste Zeile Überschrift enthält
        ' Überprüfen, ob in Spalte C bereits eine E-Mail gesendet wurde
        If ws.Cells(i, "C").Value = "" Then
            ' Erstellen einer neuen E-Mail
            Set OutMail = OutApp.CreateItem(0)
            
            ' E-Mail-Adresse aus Spalte A
            OutMail.To = ws.Cells(i, "A").Value
            
            ' Betreff und Text der E-Mail
            OutMail.Subject = "Versandbestätigung für Bestellung " & ws.Cells(i, "B").Value
            OutMail.Body = "Sehr geehrter Kunde," & vbCrLf & vbCrLf & _
                           "Ihre Bestellung mit der Nummer " & ws.Cells(i, "B").Value & " wurde versendet."
                           
            ' E-Mail senden
            OutMail.Send
            
            ' Markieren, dass die E-Mail gesendet wurde
            ws.Cells(i, "C").Value = "Gesendet"
            
            ' Freigabe des E-Mail-Objekts
            Set OutMail = Nothing
        End If
    Next i
    
    ' Freigabe des Outlook-Objekts
    Set OutApp = Nothing
    
    ' Erfolgsmeldung anzeigen
    MsgBox "Versandbestätigungen wurden erfolgreich gesendet.", vbInformation
End Sub
