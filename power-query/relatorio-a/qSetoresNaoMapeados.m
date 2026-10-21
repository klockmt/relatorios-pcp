let
    bruta = Table.Distinct(
        Table.SelectColumns(
            Table.TransformColumnTypes(qFaltaRefBruta, {{"COD_SETOR", type text}}),
            {"COD_SETOR"}
        )
    ),
    setores = Table.TransformColumnTypes(
        Excel.CurrentWorkbook(){[Name = "tblSetores"]}[Content],
        {{"CodSetor", type text}}
    ),
    orfaos = Table.SelectRows(
        bruta,
        each not List.Contains(setores[CodSetor], [COD_SETOR])
    ),
    #"Linhas Classificadas" = Table.Sort(orfaos,{{"COD_SETOR", Order.Ascending}})
in
    #"Linhas Classificadas"
