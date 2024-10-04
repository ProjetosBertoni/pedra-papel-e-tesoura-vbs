Option Explicit

Dim jogador, computador
Dim opcoes(2)

opcoes(0) = "Pedra"
opcoes(1) = "Papel"
opcoes(2) = "Tesoura"


Function EscolherComputador()
    Randomize
    EscolherComputador = Int(Rnd * 3)
End Function


jogador = InputBox("Escolha: Pedra, Papel ou Tesoura")


If jogador = "" Then
    WScript.Echo "Você não fez uma escolha."
    WScript.Quit
End If

jogador = LCase(jogador)

Select Case jogador
    Case "pedra"
        jogador = 0
    Case "papel"
        jogador = 1
    Case "tesoura"
        jogador = 2
    Case Else
        WScript.Echo "Escolha inválida! Tente novamente."
        WScript.Quit
End Select


computador = EscolherComputador()


WScript.Echo "Você escolheu: " & opcoes(jogador)
WScript.Echo "O computador escolheu: " & opcoes(computador)


If jogador = computador Then
    WScript.Echo "Empate!"
ElseIf (jogador = 0 And computador = 2) Or _
       (jogador = 1 And computador = 0) Or _
       (jogador = 2 And computador = 1) Then
    WScript.Echo "Você ganhou!"
Else
    WScript.Echo "O computador ganhou!"
End If