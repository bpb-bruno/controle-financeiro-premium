Sub RESETAR_PLANILHA()
    ' Limpa os intervalos de dados mantendo as fórmulas e estrutura
    Sheets("RECEITAS").Range("C6:N13").ClearContents
    Sheets("DESPESAS MENSAIS").Range("C7:N22").ClearContents
End Sub
