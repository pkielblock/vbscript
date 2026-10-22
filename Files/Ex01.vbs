Dim n1, n2, n3, audio, resp, respAudio, maior

Call perguntaAudio
sub perguntaAudio ()
respAudio = MsgBox("Deseja Ativar Recurso de Voz?", vbQuestion + vbYesNo, "ATEN��O")
If respAudio = vbYes Then
    Call exercicioAudio
Else
    Call exercicio
End If
End Sub

sub exercicioAudio ()
set audio = CreateObject("SAPI.SPVOICE")
audio.Volume = 100
audio.Rate = 2

n1 = CInt(InputBox("Digite o N1: ","AVISO"))
n2 = CInt(InputBox("Digite o N2: ","AVISO"))
n3 = CInt(InputBox("Digite o N3: ","AVISO"))

If n1 > n2 and n1 > n3 Then
    maior = n1
ElseIf n2 > n1 and n2 > n3 Then
    maior = n2
ElseIf n3 > n2 and n3 > n1 Then
    maior = n3
End If

audio.Speak("Maior Numero: "& maior &"")

Call pergunta
End Sub

sub exercicio ()
n1 = CInt(InputBox("Digite o N1: ","AVISO"))
n2 = CInt(InputBox("Digite o N2: ","AVISO"))
n3 = CInt(InputBox("Digite o N3: ","AVISO"))

If n1 > n2 and n1 > n3 Then
    maior = n1
ElseIf n2 > n1 and n2 > n3 Then
    maior = n2
ElseIf n3 > n2 and n3 > n1 Then
    maior = n3
End If

MsgBox("Maior Numero: "& maior &"")

Call pergunta
End Sub

sub pergunta()
resp = MsgBox("Deseja Realizar Novo Calculo?", vbQuestion + vbYesNo, "ATENCAO")
If resp = vbYes Then
    Call perguntaAudio
Else
    WScript.Quit
End If
End Sub