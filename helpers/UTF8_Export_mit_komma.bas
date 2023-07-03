Attribute VB_Name = "UTF8_Export"
Sub Erstelle_UTF8()
'Call daten.Komma_Punkt
Dim Datenbereich As String
Datenbereich = ActiveSheet.UsedRange.Address
MsgBox Datenbereich
Datenbereich = InputBox("Datenbereich eingeben")
Range(Datenbereich).Select
Dim fsT As Object, sFilename As Variant, tmpStr As String
Dim lS As Long, lZ As Long, l As Long
Dim SrcRg As Range

'Pfad und Name der zu erstellenden Datei
sFilename = Application.GetSaveAsFilename("", "CSV File (*.csv), *.csv")

If sFilename <> False Then ' Überprüfe, ob eine Datei ausgewählt wurde
        If Dir(sFilename) <> "" Then ' Überprüfe, ob die Datei bereits existiert
            ' Warnung anzeigen und Benutzerentscheidung treffen
            If MsgBox("Die ausgewählte Datei existiert bereits. Möchten Sie die Datei überschreiben?", vbQuestion + vbYesNo) = vbNo Then
                Exit Sub ' Abbruch, ohne zu speichern
            End If
        End If
    Else
        Exit Sub ' Abbruch, wenn keine Datei ausgewählt wurde
    End If



If Selection.Cells.Count > 1 Then
    Set SrcRg = Selection
Else
    Set SrcRg = ActiveSheet.UsedRange
End If

With SrcRg
    For lZ = 1 To .Rows.Count
        For lS = 1 To .Columns.Count
            tmpStr = tmpStr & "" & .Cells(lZ, lS) & ","
            Debug.Print tmpStr
        Next lS
        tmpStr = Left(tmpStr, Len(tmpStr) - 1) & vbCrLf
        Debug.Print tmpStr
    Next lZ
End With

Set fsT = CreateObject("ADODB.Stream")
fsT.Type = 2                'Stream-Typ: Text/String
fsT.Charset = "utf-8"       'Zeichensatz
fsT.Open                    'Stream öffnen
fsT.WriteText tmpStr        'Daten schreiben
fsT.SaveToFile sFilename, 2 'Datei speichern
Set fsT = Nothing
End Sub

