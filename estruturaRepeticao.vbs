Dim numeroInicial, numeroFinal, i, tipo, resp, n

Call inicio

Sub inicio()
numeroInicial = CInt(InputBox("Digite um Numero Inteiro Inicial", "Aviso"))
numeroFinal = CInt(InputBox("Digite um Numero Inteiro Final", "Aviso"))
i = CInt(InputBox("Digite o Incremento", "Aviso"))

tipo = InputBox("[F] Estrutura de Repeticao For" + vbNewLine &_
                "[W] Estrutura de Repeticao While" + vbNewLine &_
                "[E] Encerrar Script", "Selecione Uma Opcao")

Select Case tipo
    Case "F","f":
        For n = numeroInicial to numeroFinal step i
            MsgBox(n)
        Next
    Case "W","w":
        n = numeroInicial
        Do While n <= numeroFinal
            MsgBox(n)
            n = n + i
        Loop
    Case "E","e":
        resp=MsgBox("Deseja Encerrar?", vbQuestion + vbYesNo, "Atencao")
        If resp = vbYes Then
            WScript.Quit
        Else
            Call inicio
        End If
    Case Else
        MsgBox("Erro"), vbExclamation + vbOKOnly, "Atencao"
        Call inicio
End Select
End Sub
