Dim palavras(30), audio, placar, resultado, n1, n2, n3, nome, nv

Call carregarAudio
sub carregarAudio ()
set audio = CreateObject("SAPI.SPVOICE")
audio.Volume = 100
audio.Rate = 2
End Sub

Call jogador
Call inicio

Sub inicio()
'Tive que colocar as palvras sem acento, pois com acento ele não aceitava as respostas
palavras(1) = "abacate"
palavras(2) = "bala"
palavras(3) = "alicate"
palavras(4) = "abafado"
palavras(5) = "ajuda"
palavras(6) = "azeite"
palavras(7) = "azeitona"
palavras(8) = "ave"
palavras(9) = "bolo"
palavras(10) = "boneca"
palavras(11) = "hipopotamo"
palavras(12) = "gaviao"
palavras(13) = "epoca"
palavras(14) = "egua"
palavras(15) = "domino"
palavras(16) = "duvida"
palavras(17) = "duzia"
palavras(18) = "divida"
palavras(19) = "cafune"
palavras(20) = "botao"
palavras(21) = "acrimonia"
palavras(22) = "consociacão "
palavras(23) = "corolario "
palavras(24) = "elucubracoes"
palavras(25) = "efluvio"
palavras(26) = "fosmeo"
palavras(27) = "homizio"
palavras(28) = "implicito"
palavras(29) = "inocuo"
palavras(30) = "jaez"

If placar = 15 Then
    Call acabar
ElseIf placar < 5 Then
    Call nivel1
ElseIf placar >= 5 and placar < 10 Then
    Call nivel2
ElseIf placar > 5 and placar < 15 Then
    Call nivel3
End If
End Sub

Sub jogador()
nome = InputBox("Digite Seu Nome: ", "AVISO")
End Sub

Sub nivel1()
nv = "Nivel 1"

Call sorteio

If n1 = n1 Then
    Call sorteio
End If

audio.Speak(palavras(n1))

resultado = InputBox("NIVEL 1 - Digite a Palavra Ouvida Sem Acentos: ", nome)

If resultado = palavras(n1) Then
    Call venceu
Else
    Call perdeu
End If
End Sub

Sub nivel2()
nv = "Nivel 2"

Call sorteio

If n2 = n2 Then
    Call sorteio
End If

audio.Speak(palavras(n2))

resultado = InputBox("NIVEL 2 - Digite a Palavra Ouvida Sem Acentos: ", nome)

If resultado = palavras(n2) Then
    Call venceu
Else
    Call perdeu
End If
End Sub

Sub nivel3()
nv = "Nivel 3"

Call sorteio

If n3 = n3 Then
    Call sorteio
End If

audio.Speak(palavras(n3))

resultado = InputBox("NIVEL 3 - Digite a Palavra Ouvida Sem Acentos: ", nome)

If resultado = palavras(n3) Then
    Call venceu
Else
    Call perdeu
End If
End Sub

Sub sorteio()
Randomize(second(time))
'n1 = int (rnd * 10) + 1
n1 = int((10 - 1 + 1) * Rnd + 1)
n2 = int((20 - 10 + 1) * Rnd + 10)
n3 = int((30 - 20 + 1) * Rnd + 20)
End Sub

Sub venceu()
placar = placar + 1
MsgBox("Parabens Voce Venceu!" + vbNewLine &_
       "Voce Esta No " & nv &"" + vbNewLine &_
       "Seu placar e de: " & placar &""), vbInformation + vbOKOnly, "ATENCAO"
Call inicio
End Sub

Sub acabar()
MsgBox("Parabens Voce Venceu!" + vbNewLine &_
       "Voce Chegou No " & nv &"" + vbNewLine &_
       "Seu placar e de: " & placar &""), vbInformation + vbOKOnly, "ATENCAO"
resp=msgbox("Deseja Jogar Novamente?",vbquestion + vbyesno,"ATENCAO")
If resp=vbyes Then
    placar = 0
    Call jogador
    Call inicio
Else
    WScript.Quit
End If
End Sub

Sub perdeu()
MsgBox("Voce Perdeu!" + vbNewLine &_
       "Voce Chegou No " & nv &"" + vbNewLine &_
       "Com o Placar de: " & placar &""), vbInformation + vbOKOnly, "ATENCAO"
resp=msgbox("Deseja Continuar?",vbquestion + vbyesno,"ATENCAO")
If resp=vbyes Then
    Call jogador
    placar = 0
    Call inicio
Else
    WScript.Quit
End If
End Sub