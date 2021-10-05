Dim n,i,resp

Call sorteio

Sub sorteio()
i=1
Do While i <=10
    'Exemplo Numeros Randumicos
	Randomize(second(time)) 'Segundos da hora do S.O
	n = int (rnd * 100) + 1
	msgbox("Quantidade de Sorteios: "& i &"" + vbnewline &_
	       "Numero Sorteado: "& n &""),vbinformation + vbOKOnly,"AVISO"
    i=i+1
Loop
   resp=msgbox("Fim do Laco!!!" + vbnewline &_
               "Deseja realizar novo Sorteio?",vbquestion + vbyesno,"ATENCAO")
   If resp=vbyes Then
      Call sorteio
   Else
      wscript.quit
   End If
End Sub