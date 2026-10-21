let
    comandoSql = "SELECT * "
        & "FROM tabela_a "
        & "WHERE DATA >= (SELECT MAX(DATA) - INTERVAL 60 DAY FROM tabela_a) "
        & "AND DATA >= '2015-01-01' "
        & "AND COD_SETOR > '0'",
    faltaBruta = Odbc.Query(DsnErp, comandoSql),
    #"Linhas Classificadas" = Table.Sort(faltaBruta,{{"SETOR", Order.Ascending}})
in
    #"Linhas Classificadas"
