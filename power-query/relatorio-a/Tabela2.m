let
    Fonte = Excel.Workbook(File.Contents(CaminhoPremissasTabela2), null, true),
    Ref1 = Fonte{[Name="Ref"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Ref1, [PromoteAllScalars=true]),

    // Transforma colunas de mês em pares Mês/Valor
    #"Não Dinamizado" = Table.UnpivotOtherColumns(
        #"Cabeçalhos Promovidos",
        {"CTG", "Mês de competência:", "Column3"},
        "MêsSerial",
        "Valor"
    ),

    // Força Valor como number; MêsSerial permanece flexível
    #"Valor Como Numero" = Table.TransformColumns(
        #"Não Dinamizado",
        {{"Valor", each try Number.From(_) otherwise null, type number}}
    ),

    // Remove linhas sem mês válido
    #"Linha Sem Mes" = Table.SelectRows(#"Valor Como Numero", each [MêsSerial] <> null),

    // Converte mês para data, aceitando texto de data ou serial numérico
    #"Mês Data" = Table.AddColumn(
        #"Linha Sem Mes",
        "Mês",
        each
            let textoMes = Text.From([MêsSerial]) in
            try Date.From(textoMes)
            otherwise try Date.From(#date(1899, 12, 30) + #duration(Number.From(textoMes), 0, 0, 0))
            otherwise null,
        type date
    ),

    // Remove linhas onde a conversão de mês falhou
    #"Linha Com Mes" = Table.SelectRows(#"Mês Data", each [Mês] <> null),

    #"Colunas Removidas" = Table.RemoveColumns(#"Linha Com Mes", {"MêsSerial", "CTG", "Mês de competência:"}),
    #"Renomeado" = Table.RenameColumns(#"Colunas Removidas", {{"Column3", "Item"}}),

    // ---- Itens de mapeamento direto (grupo10 sai daqui) ----
    Mapeamento = #table(
        {"Item", "Categoria", "Etiqueta"},
        {
            {"Item 1/mês", "setor_d", "setor_d"},
            {"Item 2/mês", "setor_b", "setor_b"},
            {"Item 3 mês", "categoria_f", "categoria_f"},
            {"Item 4 para Setor C", "setor_a", "setor_c"},
            {"Item 5/mês TOTAL", "setor_c", "setor_cs"},
            {"ITEM 6", "setor_c", "item6"},
            {"Qtde. Itens Acabados", "setor_c", "item_acabado"},
            {"P.Item 7", "setor_e", "item7"},
            {"P.Item 8 mês", "setor_e", "item8"}
        }
    ),

    #"Direto Mesclado" = Table.NestedJoin(#"Renomeado", {"Item"}, Mapeamento, {"Item"}, "MapInfo", JoinKind.Inner),
    #"Direto Expandido" = Table.ExpandTableColumn(#"Direto Mesclado", "MapInfo", {"Categoria", "Etiqueta"}),
    #"Direto Final" = Table.SelectColumns(#"Direto Expandido", {"Categoria", "Etiqueta", "Mês", "Valor"}),

    // ---- Grupo9: soma das 3 linhas por mês ----
    #"Grupo9 Filtrado" = Table.SelectRows(#"Renomeado", each List.Contains({"P.Item 9A", "P.Item 9B", "P.Item 9C"}, [Item])),
    #"Grupo9 Agrupado" = Table.Group(#"Grupo9 Filtrado", {"Mês"}, {{"Valor", each List.Sum([Valor]), type number}}),
    #"Grupo9 Final2" = Table.SelectColumns(
        Table.AddColumn(Table.AddColumn(#"Grupo9 Agrupado", "Categoria", each "setor_e"), "Etiqueta", each "grupo9"),
        {"Categoria", "Etiqueta", "Mês", "Valor"}
    ),

    // ---- Setor A grupo10: soma de 2 linhas por mês ----
    #"Grupo10 Filtrado" = Table.SelectRows(#"Renomeado", each List.Contains({"Item 10A", "Item 10B"}, [Item])),
    #"Grupo10 Agrupado" = Table.Group(#"Grupo10 Filtrado", {"Mês"}, {{"Valor", each List.Sum([Valor]), type number}}),
    #"Grupo10 Final" = Table.SelectColumns(
        Table.AddColumn(Table.AddColumn(#"Grupo10 Agrupado", "Categoria", each "setor_a"), "Etiqueta", each "grupo10"),
        {"Categoria", "Etiqueta", "Mês", "Valor"}
    ),

    // ---- Grupo11: Total/mês - (grupo9 + item7 + item8) por mês ----
    #"Total Itens" = Table.SelectRows(#"Renomeado", each [Item] = "Total/mês"),
    #"Total Itens Sel" = Table.SelectColumns(Table.RenameColumns(#"Total Itens", {{"Valor", "TotalItens"}}), {"Mês", "TotalItens"}),

    #"Componentes" = Table.Combine({
        #"Grupo9 Final2",
        Table.SelectRows(#"Direto Final", each [Etiqueta] = "item7" or [Etiqueta] = "item8")
    }),
    #"Componentes Agrupado" = Table.Group(#"Componentes", {"Mês"}, {{"SomaComponentes", each List.Sum([Valor]), type number}}),

    #"Grupo11 Base" = Table.NestedJoin(#"Total Itens Sel", {"Mês"}, #"Componentes Agrupado", {"Mês"}, "Comp", JoinKind.Inner),
    #"Grupo11 Expandido" = Table.ExpandTableColumn(#"Grupo11 Base", "Comp", {"SomaComponentes"}),
    #"Grupo11 Valor" = Table.AddColumn(#"Grupo11 Expandido", "Valor", each [TotalItens] - [SomaComponentes], type number),
    #"Grupo11 Final2" = Table.SelectColumns(
        Table.AddColumn(Table.AddColumn(#"Grupo11 Valor", "Categoria", each "setor_e"), "Etiqueta", each "grupo11"),
        {"Categoria", "Etiqueta", "Mês", "Valor"}
    ),

    // ---- Junta tudo ----
    #"Resultado Final" = Table.Combine({#"Direto Final", #"Grupo9 Final2", #"Grupo10 Final", #"Grupo11 Final2"}),

    // Cria uma coluna de texto com o mês formatado, pra virar cabeçalho legível
    #"Mês Texto" = Table.AddColumn(#"Resultado Final", "MêsTexto", each Date.ToText([Mês], "MMM/yy"), type text),
    #"Colunas Base Removidas" = Table.RemoveColumns(#"Mês Texto", {"Mês"}),

    // Pivota: cada mês vira uma coluna, valores preenchidos com soma
    #"Pivotado" = Table.Pivot(
        #"Colunas Base Removidas",
        List.Distinct(#"Colunas Base Removidas"[MêsTexto]),
        "MêsTexto",
        "Valor",
        List.Sum
    ),

    // Força tipo number nas colunas de mês
    #"Colunas Mes Tipo Numero" = Table.TransformColumns(
        #"Pivotado",
        List.Transform(
            List.RemoveItems(Table.ColumnNames(#"Pivotado"), {"Categoria", "Etiqueta"}),
            each {_, each _, type number}
        )
    ),

    // Garante Categoria e Etiqueta como as duas primeiras colunas
    #"Ordenado" = Table.ReorderColumns(
        #"Colunas Mes Tipo Numero",
        {"Categoria", "Etiqueta"} & List.RemoveItems(Table.ColumnNames(#"Colunas Mes Tipo Numero"), {"Categoria", "Etiqueta"})
    )
in
    #"Ordenado"
