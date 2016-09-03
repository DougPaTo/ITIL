Attribute VB_Name = "Horas"
Sub teste1()
    arquivo = ActiveDocument.Name
    
'   Captura de horario no relatório mensal
    


Selection.HomeKey Unit:=wdStory

Do Until Selection = "Resumo"
    Selection.HomeKey Unit:=wdLine
    Selection.MoveRight Unit:=wdCharacter, Count:=7, Extend:=wdExtend
    If Selection = "Recurso" Then
        Selection.HomeKey Unit:=wdLine
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.Copy
        Windows("Horas").Activate
        Selection.PasteAndFormat (wdFormatPlainText)
        Selection.TypeText Text:=vbTab
        Windows(arquivo).Activate
        
        Selection.HomeKey Unit:=wdLine
        Selection.MoveDown Unit:=wdLine, Count:=2
        Selection.MoveRight Unit:=wdCharacter, Count:=7
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.Copy
        Windows("Horas").Activate
        Selection.PasteAndFormat (wdFormatPlainText)
        Selection.TypeText Text:=vbTab
        Windows(arquivo).Activate
        
        Selection.HomeKey Unit:=wdLine
        Selection.MoveDown Unit:=wdLine, Count:=1
        Selection.MoveRight Unit:=wdCharacter, Count:=8
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.Copy
        Windows("Horas").Activate
        Selection.PasteAndFormat (wdFormatPlainText)
        Selection.TypeParagraph
        Windows(arquivo).Activate
    End If
    Selection.HomeKey Unit:=wdLine
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend
Loop

End Sub


