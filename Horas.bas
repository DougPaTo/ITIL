Attribute VB_Name = "Horas"
Sub teste1()
                            '###########################################################
                            '###########################################################
                            '##   Captura de horario no relatório mensal              ##
                            '##   É necessário copiar os dados dos cards;             ##
                            '##   Colar no arquivo de relatório;                      ##
                            '##   Abrir um novo arquivo do word com o nome "Horas";   ##
                            '##   Pressionar Alt + F11;                               ##
                            '##   Selecionar o arquivo onde os dados do trello estão; ##
                            '##   Rodar a Macro.                                      ##
                            '###########################################################
                            '###########################################################


'### Captura o nome do arquivo atual ###
arquivo = ActiveDocument.Name

'### Coloca o cursor no início da página ###
Selection.HomeKey Unit:=wdStory

'### Varredura no arquivo de relatório até encontrar a frase "Resumo" ###   '##############################################################################################
Do Until Selection = "Resumo"                                               '##############################################################################################
    Selection.HomeKey Unit:=wdLine                                          '### Coloca o cursor no início da página                                                    ###
    Selection.MoveRight Unit:=wdCharacter, Count:=7, Extend:=wdExtend       '### Move o cursor 7 caracteres para a direita                                              ###
    If Selection = "Recurso" Then                                           '### Se a seleção contiver "Recurso"                                                        ###
        Selection.HomeKey Unit:=wdLine                                      '### Coloca o cursor no início da linha atual                                               ###
        Selection.MoveUp Unit:=wdLine, Count:=1                             '### Move o cursor para a linha de cima                                                     ###
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend                     '### Seleciona o conteúdo da linha (Data do atendimento)                                    ###
        Selection.Copy                                                      '### Copia as informações selecionadas                                                      ###
        Windows("Horas").Activate                                           '### Abre o arquivo "Horas"                                                                 ###
        Selection.PasteAndFormat (wdFormatPlainText)                        '### Cola os dados copiados (Data do atendimento)                                           ###
        Selection.TypeText Text:=vbTab                                      '### Da um "Tab", para que as informações fiquem no formato correto ao copiar para o Excel  ###
        Windows(arquivo).Activate                                           '### Volta para o arquivo onde estão as informações do relatório                            ###
                                                                            '##############################################################################################
        Selection.HomeKey Unit:=wdLine                                      '### Coloca o cursor no início da página                                                    ###
        Selection.MoveDown Unit:=wdLine, Count:=2                           '### Move o cursor duas linhas para baixo                                                   ###
        Selection.MoveRight Unit:=wdCharacter, Count:=7                     '### Move o cursor 7 caracteres para a direita                                              ###
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend                     '### Seleciona o conteúdo da linha (Hora de início)                                         ###
        Selection.Copy                                                      '### Copia as informações selecionadas                                                      ###
        Windows("Horas").Activate                                           '### Abre o arquivo "Horas"                                                                 ###
        Selection.PasteAndFormat (wdFormatPlainText)                        '### Cola os dados copiados (Hora do início do atendimento)                                 ###
        Selection.TypeText Text:=vbTab                                      '### Da um "Tab", para que as informações fiquem no formato correto ao copiar para o Excel  ###
        Windows(arquivo).Activate                                           '### Volta para o arquivo onde estão as informações do relatório                            ###
                                                                            '##############################################################################################
        Selection.HomeKey Unit:=wdLine                                      '### Coloca o cursor no início da página                                                    ###
        Selection.MoveDown Unit:=wdLine, Count:=1                           '### Move o cursor uma linha para baixo                                                     ###
        Selection.MoveRight Unit:=wdCharacter, Count:=8                     '### Move o cursor 8 caracteres para a direita                                              ###
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend                     '### Seleciona o conteúdo da linha (Hora de término)                                        ###
        Selection.Copy                                                      '### Copia as informações selecionadas                                                      ###
        Windows("Horas").Activate                                           '### Abre o arquivo "Horas"                                                                 ###
        Selection.PasteAndFormat (wdFormatPlainText)                        '### Cola os dados copiados (Data do atendimento)                                           ###
        Selection.TypeParagraph                                             '### Pula uma linha para inserir as informações de outra data de atendimento                ###
        Windows(arquivo).Activate                                           '### Volta para o arquivo onde estão as informações do relatório                            ###
    End If                                                                  '##############################################################################################
    Selection.HomeKey Unit:=wdLine                                          '### Coloca o cursor no início da linha                                                     ###
    Selection.MoveDown Unit:=wdLine, Count:=1                               '### Move o cursor para a linha de baixo                                                    ###
    Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend       '### Seleciona os 6 primeiros caracteres                                                    ###
Loop                                                                        '### Faz o processo novamente até que a seleção acima encontre a palavra "Resumo"           ###
                                                                            '##############################################################################################
                                                                            '##############################################################################################
End Sub



