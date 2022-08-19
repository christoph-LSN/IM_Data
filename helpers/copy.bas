Dim Zeile_Start, Zeile_Stopp As Integer
Dim Spalte_Start, Spalte_Stopp As String
Dim rng_Copy As String

Spalte_Start = "A"
Spalte_Stopp = "I"

Zeile_Start = 8738
Zeile_Stopp = 626 + 624 + 624 + 624 + 624 + 624 + 624 + 624 + 624 + 624 + 624 + 624 + 624 + 624 + 624

rng_Copy = Spalte_Start & Zeile_Start & ":" & Spalte_Stopp & Zeile_Stopp

Debug.Print rng_Copy

    Selection.AutoFill Destination:=Range(rng_Copy), Type:=xlFillDefault
    Range(rng_Copy).Select

'Selection.AutoFill Destination:=Range("A2:I626"), Type:=xlFillDefault
 '   Range("A2:I626").Select
End Sub
