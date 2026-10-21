let
    bruta = Table.SelectColumns(
        Table.TransformColumnTypes(qFaltaRefBruta, {{"COD_SETOR", type text}}),
        {"DATA", "COD_SETOR", "AUTORIZADO", "EFETIVO"}
    ),
    setores = Table.TransformColumnTypes(
        Excel.CurrentWorkbook(){[Name = "tblSetores"]}[Content],
        {{"CodSetor", type text}}
    ),
    juncao = Table.NestedJoin(
        bruta, {"COD_SETOR"},
        setores, {"CodSetor"},
        "Setor", JoinKind.Inner
    ),
    expandida = Table.ExpandTableColumn(juncao, "Setor", {"SetorPadrao"}),
    agrupada = Table.Group(
        expandida,
        {"DATA", "SetorPadrao"},
        {
            {"Autorizados", each List.Sum([AUTORIZADO]), type number},
            {"Efetivo", each List.Sum([EFETIVO]), type number},
            {"LinhasOrigem", each Table.RowCount(_), type number}
        }
    ),
    renomeada = Table.RenameColumns(
        agrupada,
        {{"DATA", "Data"}, {"SetorPadrao", "Setor"}}
    ),
    tipada = Table.TransformColumnTypes(renomeada, {{"Data", type date}})
in
    tipada
