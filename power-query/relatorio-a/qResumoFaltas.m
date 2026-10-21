let
    comandoSql = "SELECT A.DATA, A.COD_SETOR, COUNT(*) AS RegistrosFalta, SUM(A.QTD) AS SomaQtd "
        & "FROM tabela_b A "
        & "WHERE A.DATA BETWEEN '2015-01-01' AND CURDATE() "
        & "AND A.COD_SETOR > '0' "
        & "GROUP BY A.DATA, A.COD_SETOR",
    faltasPorSetor = Odbc.Query(DsnErp, comandoSql),
    faltasTexto = Table.TransformColumnTypes(faltasPorSetor, {{"COD_SETOR", type text}}),
    setores = Table.TransformColumnTypes(
        Excel.CurrentWorkbook(){[Name = "tblSetores"]}[Content],
        {{"CodSetor", type text}}
    ),
    juncao = Table.NestedJoin(
        faltasTexto, {"COD_SETOR"},
        setores, {"CodSetor"},
        "Setor", JoinKind.Inner
    ),
    expandida = Table.ExpandTableColumn(juncao, "Setor", {"SetorPadrao"}),
    agrupada = Table.Group(
        expandida,
        {"DATA", "SetorPadrao"},
        {
            {"Faltas", each List.Sum([RegistrosFalta]), type number},
            {"SomaQtd", each List.Sum([SomaQtd]), type number}
        }
    ),
    renomeada = Table.RenameColumns(
        agrupada,
        {{"DATA", "Data"}, {"SetorPadrao", "Setor"}}
    ),
    tipada = Table.TransformColumnTypes(
        renomeada,
        {{"Data", type date}, {"Faltas", type number}, {"SomaQtd", type number}}
    ),
    ordenada = Table.Sort(tipada,{{"Setor", Order.Ascending}, {"Data", Order.Ascending}})
in
    ordenada
