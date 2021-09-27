Attribute VB_Name = "Daten"
Sub Komma_Punkt()
    Dim I, Z As Integer
    Dim c  As Range
    Dim firstAddress, daten, Werte As String

daten = ActiveSheet.UsedRange.Address


Dim MaxZ As Long, MaxS As Long

With ActiveSheet        'Anpassen

   'letzte Zeile / Spalte
   MaxZ = .Cells.Find("*", , , , xlByRows, xlPrevious).Row
   MaxS = .Cells.Find("*", , , , xlByColumns, xlPrevious).Column

   'letzte Zeile / Spalte
   'MaxZ = .Cells.Find("*", , , , xlByRows, xlPrevious).Row
   'MaxS = .Cells.Find("*", , , , xlByColumns, xlPrevious).Column

   'Dein gesuchter Bereich:
  daten = .Range(.Range("A1"), .Cells(MaxZ, MaxS)).Address(0, 0)

End With



    With ActiveSheet.Range(daten)
        Set c = .Find("Value", LookIn:=xlValues)
        If Not c Is Nothing Then
            firstAddress = c.Address
        End If

  Range(firstAddress).End(xlDown).Offset(0, 0).Select
  Debug.Print Selection.Row
  Z = Selection.Row
  Debug.Print Z

    Werte = Range(firstAddress).Offset(1, 0).Select
    Selection.NumberFormat = "@"
 For I = 0 To Z
   Selection.NumberFormat = "@"

   Werte = Selection.Value
          Debug.Print Werte
   Werte = Replace(Werte, ",", ".")
   Selection.Value = Werte

   Selection.Offset(1, 0).Select



Next I

    End With

End Sub
