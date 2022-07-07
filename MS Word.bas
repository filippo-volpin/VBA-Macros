Sub Outline_Level1()
'
' Outline_Level1 Macro
'
'
    With Selection.ParagraphFormat
        .OutlineLevel = wdOutlineLevel1
    End With
End Sub
Sub Outline_Level2()
'
' Outline_Level2 Macro
'
'
    With Selection.ParagraphFormat
        .OutlineLevel = wdOutlineLevel2
    End With
End Sub
Sub Outline_Level3()
'
' Outline_Level3 Macro
'
'
    With Selection.ParagraphFormat
        .OutlineLevel = wdOutlineLevel3
    End With
End Sub
Sub Outline_Level4()
'
' Outline_Level4 Macro
'
'
    With Selection.ParagraphFormat
        .OutlineLevel = wdOutlineLevel4
    End With
End Sub
Sub Outline_BodyText()
'
' Outline_BodyText Macro
'
'
    With Selection.ParagraphFormat
        .OutlineLevel = wdOutlineLevelBodyText
    End With
End Sub
Sub Make_UPPERCASE()
'
' Make_UPPERCASE Macro
'
'
    Selection.Range.Case = wdUpperCase
End Sub
Sub Make_Capitalise()
'
' Make_Capitalise Macro
'
'
    Selection.Range.Case = wdTitleWord
End Sub
Sub Format_Blue()
'
' Format_Blue Macro
' Keyboard Shortcut: Alt + B
'
    Selection.Font.Bold = True
    Selection.Font.TextColor.RGB = RGB(0, 112, 192)
    Selection.Shading.BackgroundPatternColor = RGB(222, 234, 246)
End Sub
Sub Format_Green()
'
' Format_Green Macro
' Keyboard Shortcut: Alt + G
'
    Selection.Font.Bold = True
    Selection.Font.TextColor.RGB = RGB(0, 176, 80)
    Selection.Shading.BackgroundPatternColor = RGB(237, 245, 231)
End Sub
Sub Format_Orange()
'
' Format_Orange Macro
' Keyboard Shortcut: Alt + O
'
    Selection.Font.Bold = True
    Selection.Font.TextColor.RGB = RGB(237, 125, 49)
    Selection.Shading.BackgroundPatternColor = RGB(251, 228, 214)
End Sub
Sub Format_Purple()
'
' Format_Purple Macro
' Keyboard Shortcut: Alt + P
'
    Selection.Font.Bold = True
    Selection.Font.TextColor.RGB = RGB(204, 0, 255)
    Selection.Shading.BackgroundPatternColor = RGB(255, 221, 255)
End Sub
Sub Format_Red()
'
' Format_Red Macro
' Keyboard Shortcut: Alt + R
'
    Selection.Font.Bold = True
    Selection.Font.TextColor.RGB = RGB(255, 0, 0)
End Sub
Sub Format_Yellow()
'
' Format_Yellow Macro
' Keyboard Shortcut: Alt + Y
'
    Selection.Shading.BackgroundPatternColor = RGB(255, 229, 153)
End Sub
Sub Format_Normal()
'
' Format_Normal Macro
' Keyboard Shortcut: Alt + N
'
    Selection.Font.Bold = False
    Selection.Font.TextColor = wdColorAutomatic
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = wdColorAutomatic
End Sub
Sub Curly_Braces()
'
' Curly_Braces Macro (wraps selection in curly braces)
' Keyboard Shortcut: Ctrl + Alt + [
'
    Selection.Text = "{" & Selection.Text & "}"
End Sub
Sub Wrap_Quotation_Marks()
'
' Wrap_Quotation_Marks Macro
' Keyboard Shortcut: Alt + Shift + :
'
    strQuote = Chr$(34)
    Selection.Text = strQuote & Selection.Text & strQuote
End Sub
Sub Summation_Operator()
'
' Summation_Operator Macro
' Keyboard Shortcut: Ctrl + Alt + ]
'
    Selection.TypeText Text:="\sum_{i=0}^{T} {x}"
End Sub
Sub Sigma()
'
' Sigma Macro
' Keyboard Shortcut: Alt + i
'
    Selection.TypeText Text:="\sigma^{2}"
End Sub
Sub Distribution_Convergence_Arrow()
'
' Distribution_Convergence_Arrow Macro
' Keyboard Shortcut: Alt + #
'
    Selection.TypeText Text:="\longrightarrow\above{D}"
End Sub
Sub OneColumn()
'
' OneColumn Macro
'
'
    With Selection.PageSetup.TextColumns
        .SetCount NumColumns:=1
        .EvenlySpaced = True
    End With
End Sub
Sub TwoColumns()
'
' TwoColumns Macro
'
'
    With Selection.PageSetup.TextColumns
        .SetCount NumColumns:=2
        .EvenlySpaced = True
    End With
End Sub
Sub Delete_NewLines()
'
' Removing all newlines in selected text, adding a whitespace
' Keyboard Shortcut: Alt + S
'
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Sub Delete_NewLines_Place_Comma()
'
' Removing all newlines in selected text, adding a comma
' Keyboard Shortcut: Alt + W
'
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = ", "
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
End Sub
