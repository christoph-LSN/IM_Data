Attribute VB_Name = "Modul2"
Sub ZeileNachUntenZiehen()
    Dim selectedRow As Integer
    Dim Zeilen As Integer
    
    ' �berpr�fen, ob eine Zeile ausgew�hlt ist
    If Selection.Rows.Count <> 1 Then
        MsgBox "Bitte w�hlen Sie genau eine Zeile aus, die Sie nach unten ziehen m�chten.", vbExclamation
        Exit Sub
    End If
    
    Zeilen = InputBox("Anzahl Zeilen", "Zeilen")
    
    ' Die ausgew�hlte Zeile ermitteln
    selectedRow = Selection.Row
    
    ' Die ausgew�hlte Zeile 50-mal nach unten kopieren
    Rows(selectedRow & ":" & selectedRow).Copy
    Rows(selectedRow + 1 & ":" & selectedRow + Zeilen).Insert Shift:=xlDown
    Application.CutCopyMode = False
    
    'Zur lezte Zeile gehen
    Cells(selectedRow + Zeilen, 1).Select
    
   
    
    
End Sub
