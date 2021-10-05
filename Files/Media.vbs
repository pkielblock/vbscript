'Declarar Variaveis Locais
Dim n1, n2, n3, media, situacao, resp, audio

Call carregarAudio
sub carregarAudio ()
set audio = CreateObject("SAPI.SPVOICE") 'M�dulo de Voz - Recurso do Windows
audio.Volume = 100 'Volume do Audio
audio.Rate = 2 'Velocidade - Positivo Mais Rapido, Negativo Mais Lento
Call entradaNotas
End Sub

sub entradaNotas ()
'Entradas de Dados
'Conversao de para Double | CDbl para Double - CInt para Int - CCurl para Monetarios - CStr para String - CDate para Datas
n1 = CDbl(InputBox("Digite a N1: ","AVISO"))
n2 = CDbl(InputBox("Digite a N2: ","AVISO"))
n3 = CDbl(InputBox("Digite a N3: ","AVISO"))

'Processamento
'Comando Round delimita casas decimais
media = Round((n1 + n2 + n3) / 3, 1) 

If media < 4 Then
    situacao = "Reprovado"
ElseIf media >= 4 and media < 6 Then
    situacao = "Exame"
Else
    situacao = "Aprovado"
End If

'Saida de Dados Por Voz
audio.Speak("Rendimento do Aluno" + vbNewLine &_
             "Media do Aluno: "& media & "" + vbNewLine &_
             "Situacao do Aluno: "& situacao & "")
'Saida de Dados Por Mensagem
MsgBox("Rendimento do Aluno" + vbNewLine &_
       "Media do Aluno: "& media & "" + vbNewLine &_
       "Situacao do Aluno: "& situacao & ""), vbInformation + vbOKOnly, "AVISO"
Call pergunta
End Sub

sub pergunta()
resp = MsgBox("Deseja Realizar Novo Calculo?", vbQuestion + vbYesNo, "ATENCAO")
if resp = vbYes Then
    call entradaNotas
Else
    WScript.Quit
End If
End Sub