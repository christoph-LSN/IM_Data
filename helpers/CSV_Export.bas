Attribute VB_Name = "Modul1"
Sub Komma_Punkt_Optimiert()
    Dim c As Range
    Dim firstAddress As String
    Dim StartRow As Long, EndRow As Long
    Dim Werte As String
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Finde die Zelle mit dem Inhalt "Value"
    Set c = ws.UsedRange.Find("Value", LookIn:=xlValues)
    If c Is Nothing Then
        MsgBox """Value"" wurde nicht gefunden."
        Exit Sub
    End If

    firstAddress = c.Address
    StartRow = c.Row + 1
    EndRow = ws.Cells(c.Row, c.Column).End(xlDown).Row

    ' Schleife über die relevanten Zellen
    Dim i As Long
    For i = StartRow To EndRow
        With ws.Cells(i, c.Column)
            .NumberFormat = "@"
            Werte = .Value
            Debug.Print "Vorher: " & Werte
            Werte = Replace(Werte, ",", ".")
            .Value = Werte
            Debug.Print "Nachher: " & Werte
        End With
    Next i

    MsgBox "Fertig: Kommas wurden durch Punkte ersetzt."
End Sub

Sub Erstelle_UTF8_prozent()
    Dim Datenbereich As String
    Dim fsT As Object, sFilename As Variant, tmpStr As String
    Dim lS As Long, lZ As Long
    Dim SrcRg As Range
    Dim Prozent As Integer, LetzterProzent As Integer

    ' Aktuellen Datenbereich anzeigen
    Datenbereich = ActiveSheet.UsedRange.Address
    MsgBox "Aktueller Datenbereich: " & Datenbereich

    ' Benutzerdefinierten Bereich abfragen
    Datenbereich = InputBox("Datenbereich eingeben (z. B. A1:D100):", "CSV Export", Datenbereich)

    ' Bereich prüfen und setzen
    On Error Resume Next
    Set SrcRg = Range(Datenbereich)
    On Error GoTo 0
    If SrcRg Is Nothing Then
        MsgBox "Ungültiger Bereich. Vorgang abgebrochen.", vbExclamation
        Exit Sub
    End If

    ' Dateiname abfragen
    sFilename = Application.GetSaveAsFilename("", "CSV File (*.csv), *.csv")
    If VarType(sFilename) = vbBoolean Then Exit Sub ' Abbruch bei Abbrechen

    ' Warnung bei existierender Datei
    If Dir(sFilename) <> "" Then
        If MsgBox("Die Datei existiert bereits. Überschreiben?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    ' Stream vorbereiten
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2                ' Text-Stream
    fsT.Charset = "utf-8"       ' UTF-8 Zeichensatz
    fsT.Open                    ' Stream öffnen

    ' Fortschrittsanzeige vorbereiten
    Application.StatusBar = "Export beginnt..."
    DoEvents
    LetzterProzent = -1

    ' Daten zeilenweise schreiben
    With SrcRg
        For lZ = 1 To .Rows.Count
            tmpStr = ""
            For lS = 1 To .Columns.Count
                If IsEmpty(.Cells(lZ, lS)) Then
                    tmpStr = tmpStr & ","
                Else
                    tmpStr = tmpStr & .Cells(lZ, lS).Text & ","
                End If
            Next lS
            tmpStr = Left(tmpStr, Len(tmpStr) - 1) & vbCrLf
            fsT.WriteText tmpStr

            ' Fortschritt in Prozent anzeigen
            Prozent = Int((lZ / .Rows.Count) * 100)
            If Prozent <> LetzterProzent Then
                Application.StatusBar = "Fortschritt: " & Prozent & "%"
                LetzterProzent = Prozent
                DoEvents
            End If
        Next lZ
    End With

    ' Datei speichern
    fsT.SaveToFile sFilename, 2 ' 2 = Überschreiben
    fsT.Close
    Set fsT = Nothing

    ' Statusleiste zurücksetzen
    Application.StatusBar = False

    MsgBox "CSV-Datei erfolgreich erstellt unter:" & vbCrLf & sFilename, vbInformation
End Sub

