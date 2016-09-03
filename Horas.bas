Attribute VB_Name = "Horas"
Sub teste1()
                            '###########################################################
                            '###########################################################
                            '##   Captura de horario no relat�rio mensal              ##
                            '##   � necess�rio copiar os dados dos cards;             ##
                            '##   Colar no arquivo de relat�rio;                      ##
                            '##   Abrir um novo arquivo do word com o nome "Horas";   ##
                            '##   Pressionar Alt + F11;                               ##
                            '##   Selecionar o arquivo onde os dados do trello est�o; ##
                            '##   Rodar a Macro.                                      ##
                            '###########################################################
                            '###########################################################


'### Captura o nome do arquivo atual ###
arquivo = ActiveDocument.Name

'### Coloca o cursor no in�cio da p�gina ###
Selection.HomeKey Unit:=wdStory

'### Varredura no arquivo de relat�rio at� encontrar a frase "Resumo" ###   '##############################################################################################
Do Until Selection = "Resumo"                                               '##############################################################################################
    Selection.HomeKey Unit:=wdLine                                          '### Coloca o cursor no in�cio da p�gina                                                    ###
    Selection.MoveRight Unit:=wdCharacter, Count:=7, Extend:=wdExtend       '### Move o cursor 7 caracteres para a direita                                              ###
    If Selection = "Recurso" Then                                           '### Se a sele��o contiver "Recurso"                                                        ###
        Selection.HomeKey Unit:=wdLine                                      '### Coloca o cursor no in�cio da linha atual                                               ###
        Selection.MoveUp Unit:=wdLine, Count:=1                             '### Move o cursor para a linha de cima                                                     ###
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend                     '### Seleciona o conte�do da linha (Data do atendimento)                                    ###
        Selection.Copy                                                      '### Copia as informa��es selecionadas                                                      ###
        Windows("Horas").Activate                                           '### Abre o arquivo "Horas"                                                                 ###
        Selection.PasteAndFormat (wdFormatPlainText)                        '### Cola os dados copiados (Data do atendimento)                                           ###
        Selection.TypeText Text:=vbTab                                      '### Da um "Tab", para que as informa��es fiquem no formato correto ao copiar para o Excel  ###
        Windows(arquivo).Activate                                           '### Volta para o arquivo onde est�o as informa��es do relat�rio                            ###
                                                                            '##############################################################################################
        Selection.HomeKey Unit:=wdLine                                      '### Coloca o cursor no in�cio da p�gina                                                    ###
        Selection.MoveDown Unit:=wdLine, Count:=2                           '### Move o cursor duas linhas para baixo                                                   ###
        Selection.MoveRight Unit:=wdCharacter, Count:=7                     '### Move o cursor 7 caracteres para a direita                                              ###
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend                     '### Seleciona o conte�do da linha (Hora de in�cio)                                         ###
        Selection.Copy                                                      '### Copia as informa��es selecionadas                                                      ###
        Windows("Horas").Activate                                           '### Abre o arquivo "Horas"                                                                 ###
        Selection.PasteAndFormat (wdFormatPlainText)                        '### Cola os dados copiados (Hora do in�cio do atendimento)                                 ###
        Selection.TypeText Text:=vbTab                                      '### Da um "Tab", para que as informa��es fiquem no formato correto ao copiar para o Excel  ###
        Windows(arquivo).Activate                                           '### Volta para o arquivo onde est�o as informa��es do relat�rio                            ###
                                                                            '##############################################################################################
        Selection.HomeKey Unit:=wdLine                                      '### Coloca o cursor no in�cio da p�gina                                                    ###
        Selection.MoveDown Unit:=wdLine, Count:=1                           '### Move o cursor uma linha para baixo                                                     ###
        Selection.MoveRight Unit:=wdCharacter, Count:=8                     '### Move o cursor 8 caracteres para a direita                                              ###
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend                     '### Seleciona o conte�do da linha (Hora de t�rmino)                                        ###
        Selection.Copy                                                      '### Copia as informa��es selecionadas                                                      ###
        Windows("Horas").Activate                                           '### Abre o arquivo "Horas"                                                                 ###
        Selection.PasteAndFormat (wdFormatPlainText)                        '### Cola os dados copiados (Data do atendimento)                                           ###
        Selection.TypeParagraph                                             '### Pula uma linha para inserir as informa��es de outra data de atendimento                ###
        Windows(arquivo).Activate                                           '### Volta para o arquivo onde est�o as informa��es do relat�rio                            ###
    End If                                                                  '##############################################################################################
    Selection.HomeKey Unit:=wdLine                                          '### Coloca o cursor no in�cio da linha                                                     ###
    Selection.MoveDown Unit:=wdLine, Count:=1                               '### Move o cursor para a linha de baixo                                                    ###
    Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend       '### Seleciona os 6 primeiros caracteres                                                    ###
Loop                                                                        '### Faz o processo novamente at� que a sele��o acima encontre a palavra "Resumo"           ###
                                                                            '##############################################################################################
                                                                            '##############################################################################################
End Sub



