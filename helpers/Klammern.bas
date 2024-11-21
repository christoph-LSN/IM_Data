Attribute VB_Name = "Klammern"
Sub KlammernUndSchraegstrichEntfernenMitFormel()
    Dim I As Integer
    Dim c As Range
    Dim firstAddress As String
    Dim daten As String
    Dim Werte As String
    Dim MaxZ As Long, MaxS As Long

    ' Letzten Bereich ermitteln
    With ActiveSheet
        MaxZ = .Cells.Find("*", , , , xlByRows, xlPrevious).Row
        MaxS = .Cells.Find("*", , , , xlByColumns, xlPrevious).Column

        ' Bereich für die Daten setzen
        daten = .Range(.Range("A1"), .Cells(MaxZ, MaxS)).Address(0, 0)
    End With

    ' Datenbereich durchsuchen
    With ActiveSheet.Range(daten)
        ' Suche nach der Zelle, die "Value" enthält
        Set c = .Find("Value", LookIn:=xlValues)
        If Not c Is Nothing Then
            firstAddress = c.Address
        End If

        ' Beginnt eine Zeile unter der Zelle mit "Value"
        Set c = Range(firstAddress).Offset(1, 0)

        ' Berechnung der Anzahl der Zeilen bis zur letzten Zeile
        For I = c.Row To MaxZ
            ' Zelleninhalt als String formatieren, um sicherzustellen, dass er korrekt verarbeitet wird
            c.NumberFormat = "@"

            ' Aktuellen Wert in der Zelle abfragen
            Werte = c.Value

            ' Logik der Formel implementieren:
            If InStr(Werte, "/") > 0 Then
                ' Wenn ein Schrägstrich vorhanden ist, schreibe "gesperrt" in die Zelle links
                c.Offset(0, -1).Value = "gesperrt"
                ' Entferne den Schrägstrich aus der aktuellen Zelle
                Werte = Replace(Werte, "/", "")
                c.Value = Werte
            ElseIf InStr(Werte, "(") > 0 And InStr(Werte, ")") > 0 Then
                ' Wenn sowohl eine öffnende als auch eine schließende Klammer vorhanden ist
                c.Offset(0, -1).Value = "Statistisch unsicher"
                ' Entferne die Klammern aus der aktuellen Zelle
                Werte = Replace(Werte, "(", "")
                Werte = Replace(Werte, ")", "")
                c.Value = Werte
            Else
                ' Wenn keine der Bedingungen erfüllt ist, bleibt die Zelle links leer
                c.Offset(0, -1).Value = ""
            End If

            ' Zur nächsten Zelle springen
            Set c = c.Offset(1, 0)
        Next I

    End With
End Sub

