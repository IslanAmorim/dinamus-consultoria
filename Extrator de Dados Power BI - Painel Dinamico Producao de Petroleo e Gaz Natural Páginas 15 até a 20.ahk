#Persistent
#SingleInstance Force
SetTitleMatchMode, 2  ; Para garantir a correspondência parcial do título da janela

Sleep, 20000  ; Espera 20 segundos antes de iniciar

Loop  ; Loop principal
{
    ; 1. Ativa a janela do Power BI
    IfWinExist, Microsoft Power BI
    {
        WinActivate
        Sleep, 1000
    }
    else
    {
        MsgBox, Power BI não encontrado. Verifique se está aberto.
        ExitApp
    }

    ; 2. Seleciona exatamente 50 linhas no Power BI
    Send, {Shift Down}
    Loop, 50
    {
        Send, {Down}
        Sleep, 150  ; Aumenta o tempo para garantir que todas as linhas sejam selecionadas corretamente
    }
    Send, {Shift Up}  ; Solta o Shift após selecionar as 50 linhas
    Sleep, 500

    ; 3. Copiar a Seleção com o menu de contexto (botão direito)
    Send, {AppsKey}  ; Abre o menu de contexto (equivalente ao botão direito)
    Sleep, 500
    Send, {Right}  ; Navega para a direita no menu
    Sleep, 200
    Send, {Down}  ; Seta para baixo para selecionar "Copiar Seleção"
    Sleep, 200
    Send, {Enter}  ; Executa o comando de copiar
    Sleep, 1500  ; Tempo para copiar

    ; 4. Alterna para o PlanMaker
    IfWinExist, PlanMaker
    {
        WinActivate
        Sleep, 1000
    }
    else
    {
        MsgBox, PlanMaker não encontrado. Verifique se está aberto.
        ExitApp
    }

    ; 5. Cola no PlanMaker
    Send, ^v  ; Ctrl + V para colar
    Sleep, 1500  ; Aumenta o tempo para garantir que a colagem seja completa

    ; 6. Mover uma linha para cima para excluir a linha de cima
    Send, {Up}  ; Sobe uma linha para posicionar na linha de cima (cabeçalho)
    Sleep, 500

    ; 7. Excluir a linha de cima (primeira vez)
    Send, {AppsKey}  ; Abre o menu de contexto (botão direito)
    Sleep, 500
    Loop, 9
    {
        Send, {Down}  ; Navega 9 vezes para baixo no menu
        Sleep, 100
    }
    Send, {Enter}  ; Seleciona a opção de deletar a linha
    Sleep, 500
    Send, {Down 2}  ; Move para baixo duas linhas para garantir que a exclusão está completa
    Sleep, 200
    Send, {Enter}  ; Deleta a linha
    Sleep, 500

    ; 8. Excluir a segunda linha (logo após a primeira exclusão)
    Send, {Up}  ; Sobe novamente para a linha que ficou após a exclusão
    Sleep, 500
    Send, {AppsKey}  ; Abre o menu de contexto novamente (botão direito)
    Sleep, 500
    Loop, 9
    {
        Send, {Down}  ; Navega 9 vezes para baixo no menu
        Sleep, 100
    }
    Send, {Enter}  ; Seleciona a opção de deletar a segunda linha
    Sleep, 500
    Send, {Down 2}  ; Move para baixo duas linhas para garantir que a exclusão está completa
    Sleep, 200
    Send, {Enter}  ; Deleta a segunda linha
    Sleep, 500

    ; 9. Excluir a terceira linha (sem mover para baixo)
    Send, {AppsKey}  ; Abre o menu de contexto novamente (botão direito)
    Sleep, 500
    Loop, 9
    {
        Send, {Down}  ; Navega 9 vezes para baixo no menu
        Sleep, 100
    }
    Send, {Enter}  ; Seleciona a opção de deletar a terceira linha
    Sleep, 500
    Send, {Down 2}  ; Move para baixo duas linhas
    Sleep, 200
    Send, {Enter}  ; Deleta a terceira linha
    Sleep, 500

    ; 10. Contar exatamente 52 linhas coladas, sem pular ou deixar linhas vazias
    Loop, 52    {
        Send, {Down}
        Sleep, 100  ; Tempo para garantir que o movimento seja feito linha por linha
    }

    ; 11. Mover uma linha para baixo (linha livre) para continuar o ciclo
    Send, {Down}
    Sleep, 100

    ; 12. Alterna de volta para o Power BI
    IfWinExist, Microsoft Power BI
    {
        WinActivate
        Sleep, 1000
    }
    else
    {
        MsgBox, Power BI não encontrado. Verifique se está aberto.
        ExitApp
    }

    ; 13. Mover uma linha para baixo e pressionar Enter para confirmar a nova linha
    Send, {Down}  ; Move uma linha para baixo
    Sleep, 500
    Send, {Enter}  ; Pressiona Enter para confirmar a linha
    Sleep, 500

    ; O ciclo continua com a próxima seleção de 50 linhas
}
