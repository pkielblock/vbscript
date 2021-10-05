Dim resp, i, placar, resultado, n1, n2, op, conta

Call inicio

Sub inicio()
Call sorteio
Select Case op
    Case 1:
        conta = n1 + n2
        resultado = CInt(InputBox(n1 & " + " & n2 & " = ?","AVISO"))
    Case 2:
        conta = n1 - n2
        resultado = CInt(InputBox(n1 & " - " & n2 & " = ?","AVISO"))
    Case 3:
        conta = n1 * n2
        resultado = CInt(InputBox(n1 & " * " & n2 & " = ?","AVISO"))
    Case Else
        MsgBox("Erro"), vbExclamation + vbOKOnly, "Atencao"
        Call inicio
    End Select

If resultado = conta Then
    placar = placar + 1
    Call venceu
Else
    Call perdeu
End If
End Sub

Sub sorteio()
Randomize(second(time))
n1 = int (rnd * 10) + 1
n2 = int (rnd * 10) + 1
op = int (rnd * 3) + 1
End Sub

Sub venceu()
MsgBox("Parabens Voce Venceu!" + vbNewLine &_
       "Seu placar e de: " & placar &""), vbInformation + vbOKOnly, "ATENCAO"
Call inicio
End Sub

Sub perdeu()
MsgBox("Voce Perdeu!" + vbNewLine &_
       "Com o Placar de: " & placar &""), vbInformation + vbOKOnly, "ATENCAO"
resp=msgbox("Deseja Continuar?",vbquestion + vbyesno,"ATENCAO")
If resp=vbyes Then
    placar = 0
    Call inicio
Else
    WScript.Quit
End If
End Sub