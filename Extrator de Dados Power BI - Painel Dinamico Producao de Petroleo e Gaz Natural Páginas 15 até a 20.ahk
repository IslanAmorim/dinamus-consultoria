#Persistent
#SingleInstance Force
SetTitleMatchMode, 2  ; Para garantir a correspond�ncia parcial do t�tulo da janela

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
        MsgBox, Power BI n�o encontrado. Verifique se est� aberto.
        ExitApp
    }

    ; 2. Seleciona exatamente 50 linhas no Power BI
    Send, {Shift Down}
    Loop, 50
    {
        Send, {Down}
        Sleep, 150  ; Aumenta o tempo para garantir que todas as linhas sejam selecionadas corretamente
    }
    Send, {Shift Up}  ; Solta o Shift ap�s selecionar as 50 linhas
    Sleep, 500

    ; 3. Copiar a Sele��o com o menu de contexto (bot�o direito)
    Send, {AppsKey}  ; Abre o menu de contexto (equivalente ao bot�o direito)
    Sleep, 500
    Send, {Right}  ; Navega para a direita no menu
    Sleep, 200
    Send, {Down}  ; Seta para baixo para selecionar "Copiar Sele��o"
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
        MsgBox, PlanMaker n�o encontrado. Verifique se est� aberto.
        ExitApp
    }

    ; 5. Cola no PlanMaker
    Send, ^v  ; Ctrl + V para colar
    Sleep, 1500  ; Aumenta o tempo para garantir que a colagem seja completa

    ; 6. Mover uma linha para cima para excluir a linha de cima
    Send, {Up}  ; Sobe uma linha para posicionar na linha de cima (cabe�alho)
    Sleep, 500

    ; 7. Excluir a linha de cima (primeira vez)
    Send, {AppsKey}  ; Abre o menu de contexto (bot�o direito)
    Sleep, 500
    Loop, 9
    {
        Send, {Down}  ; Navega 9 vezes para baixo no menu
        Sleep, 100
    }
    Send, {Enter}  ; Seleciona a op��o de deletar a linha
    Sleep, 500
    Send, {Down 2}  ; Move para baixo duas linhas para garantir que a exclus�o est� completa
    Sleep, 200
    Send, {Enter}  ; Deleta a linha
    Sleep, 500

    ; 8. Excluir a segunda linha (logo ap�s a primeira exclus�o)
    Send, {Up}  ; Sobe novamente para a linha que ficou ap�s a exclus�o
    Sleep, 500
    Send, {AppsKey}  ; Abre o menu de contexto novamente (bot�o direito)
    Sleep, 500
    Loop, 9
    {
        Send, {Down}  ; Navega 9 vezes para baixo no menu
        Sleep, 100
    }
    Send, {Enter}  ; Seleciona a op��o de deletar a segunda linha
    Sleep, 500
    Send, {Down 2}  ; Move para baixo duas linhas para garantir que a exclus�o est� completa
    Sleep, 200
    Send, {Enter}  ; Deleta a segunda linha
    Sleep, 500

    ; 9. Excluir a terceira linha (sem mover para baixo)
    Send, {AppsKey}  ; Abre o menu de contexto novamente (bot�o direito)
    Sleep, 500
    Loop, 9
    {
        Send, {Down}  ; Navega 9 vezes para baixo no menu
        Sleep, 100
    }
    Send, {Enter}  ; Seleciona a op��o de deletar a terceira linha
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
        MsgBox, Power BI n�o encontrado. Verifique se est� aberto.
        ExitApp
    }

    ; 13. Mover uma linha para baixo e pressionar Enter para confirmar a nova linha
    Send, {Down}  ; Move uma linha para baixo
    Sleep, 500
    Send, {Enter}  ; Pressiona Enter para confirmar a linha
    Sleep, 500

    ; O ciclo continua com a pr�xima sele��o de 50 linhas
}
