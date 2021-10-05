Dim n1, resp, respAudio, sucessor, antecessor

Call perguntaAudio

sub perguntaAudio ()
respAudio = MsgBox("Deseja Ativar Recurso de Voz?", vbQuestion + vbYesNo, "ATENCAO")
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
sucessor = n1 + 1
antecessor = n1 - 1

audio.Speak("Sucessor: "& sucessor &"" + vbNewLine &_
            "Antecessor: "& antecessor &"")
End Sub

sub exercicio()
n1 = CInt(InputBox("Digite o N1: ","AVISO"))
sucessor = n1 + 1
antecessor = n1 - 1

MsgBox("Sucessor: "& sucessor &"" + vbNewLine &_
       "Antecessor: "& antecessor &"")
End Sub

sub pergunta()
resp = MsgBox("Deseja Realizar Novo Calculo?", vbQuestion + vbYesNo, "ATENCAO")
If resp = vbYes Then
    Call perguntaAudio
Else
    WScript.Quit
End If
End Sub