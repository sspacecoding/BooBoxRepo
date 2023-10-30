Option Explicit
Dim messages(6), randomMessage, delayTime, i

' Defina as mensagens: 
messages(0) = "Estou te observando"
messages(1) = "Coé!"
messages(2) = "Eu te vejo..."
messages(3) = "Tu não podes se esconder!"
messages(4) = "Estou atrás de ti..."
messages(5) = "Cuidado com os fantasmas..."
messages(6) = "Eu voltarei..."

' Mostre as mensagens aleatoriamente com atrasos em loop infinito
Do
    randomMessage = messages(Int((6 - 0 + 1) * Rnd + 0)) ' Escolha uma mensagem aleatória
    delayTime = Int((5 - 2 + 1) * Rnd + 2) * 1000 ' Tempo de atraso entre 2 e 5 segundos
    
    MsgBox randomMessage, 0, "Halloween"
    WScript.Sleep delayTime
Loop
