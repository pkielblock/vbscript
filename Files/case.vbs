Dim cor, op, resp

Call inicio

Sub inicio()
op = CInt(InputBox("[1] Verde" + vbNewLine &_
                   "[2] Amarelo" + vbNewLine &_
                   "[3] Vermelho" + vbNewLine &_
                   "[0/10] Encerrar o Script","Cores do Semaforo"))
Select Case op
    Case 1:
        cor = "Verde - Siga em Frente"
    Case 2:
        cor = "Amarelo - Atencao"
    Case 3:
        cor = "Vermelho - PARE"
    Case 0,10:
        resp=MsgBox("Deseja Encerrar?", vbQuestion + vbYesNo, "Atencao")
        If resp = vbYes Then
            WScript.Quit
        Else
            Call inicio
        End If
    Case Else
        MsgBox("Erro, Digite um Numero Valido"), vbExclamation + vbOKOnly, "Atencao"
        Call inicio
End Select
MsgBox("Voce Selecionou: "& cor &""), vbInformation + vbOKOnly, "Cores Semaforo"
End Sub