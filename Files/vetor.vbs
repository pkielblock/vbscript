Dim n, cor(5), i, audio

Call carregarAudio
sub carregarAudio ()
set audio = CreateObject("SAPI.SPVOICE")
audio.Volume = 100
audio.Rate = 2
End Sub

Call carregarCores

Sub carregarCores()
cor(1) = "Azul"
cor(2) = "Vermelho"
cor(3) = "Verde"
cor(4) = "Amarelo"
cor(5) = "Preto"

n = 1

Do While n <= 5
    Randomize(second(time))
    i = int(rnd * 5) + 1
    MsgBox(cor(i))
    audio.Speak(cor(i))
    n = n + 1
Loop
    MsgBox("Fim do Laco")
    audio.Speak("Fim do Laco")
End Sub