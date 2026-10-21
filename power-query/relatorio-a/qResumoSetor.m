let
    resumoSetor = Odbc.Query(
        DsnErp,
        "SELECT DATA, COD_SETOR, SETOR, COUNT(*) AS TOTAL_LINHAS "
        & "FROM tabela_a "
        & "WHERE DATA = '2026-01-01' "
        & "GROUP BY DATA, COD_SETOR, SETOR "
        & "ORDER BY COD_SETOR"
    ),
    #"Linhas Filtradas" = Table.SelectRows(resumoSetor, each [SETOR] <> "SETOR X - UNIDADE Y" and [SETOR] <> "SETOR Y - UNIDADE Y")
in
    #"Linhas Filtradas"
