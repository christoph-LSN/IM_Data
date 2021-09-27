Attribute VB_Name = "Modul1"
Sub Format_Downloadtabelle()
Attribute Format_Downloadtabelle.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("B1").Value = "Migration und Teilhabe in Niedersachsen - Integrationsmonitoring 2021"
    Rows("1:4").Select
    Selection.RowHeight = 15
    Range("B1").Select
    Range("B1").Select
    With Selection.Font
        .Name = "NDSFrutiger 55 Roman"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "NDSFrutiger 55 Roman"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("B3").Select
    With Selection.Font
        .Name = "NDSFrutiger 55 Roman"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "NDSFrutiger 55 Roman"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("B4").Select
    With Selection.Font
        .Name = "NDSFrutiger 55 Roman"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Columns("A:A").Select
    Selection.ColumnWidth = 5
    ActiveWindow.DisplayGridlines = False
End Sub
