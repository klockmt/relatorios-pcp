# Relatórios: VBA e Power Query

Exemplo de automação de relatórios em Excel. Os nomes de setores, itens, tabelas e caminhos foram trocados por valores ilustrativos. A estrutura e a lógica são as originais.

## Estrutura

| Caminho | Conteúdo |
|---|---|
| `vba/modRelRef.bas` | `GerarRelatorioMensal`: monta a aba `RelRef` por setor, dia, semana e totais |
| `vba/modRelTabela2.bas` | `GerarAbaRelTabela2`: monta a aba `RelTabela2` por categoria a partir da tabela `Tabela2` |
| `vba/modExportacao.bas` | `ExportarLancamentos`: gera CSV com saldo e faltas por setor |
| `power-query/relatorio-a/*.m` | Uma consulta por arquivo, com o nome da consulta no Power Query |
| `power-query/fonte-b/Ref.m` | Consulta da fonte B a partir da planilha de origem |

## Power Query

Cada arquivo `.m` é uma consulta. Crie no Power Query uma consulta em branco com o mesmo nome e cole o conteúdo.

As consultas `qResumoRH`, `qResumoFaltas` e `qSetoresNaoMapeados` dependem de `qFaltaRefBruta` e de `Tabela2`. Crie todas antes de atualizar.

### Parâmetros

| Nome | Valor | Usado em |
|---|---|---|
| `DsnErp` | `dsn=<nome do DSN ODBC>` | `relatorio-a` |
| `CaminhoPremissasSig` | Caminho completo da planilha de origem | `relatorio-a` e `fonte-b` |

Sem esses parâmetros, as consultas falham ao atualizar.

## VBA

### Parâmetros no Excel

- `modRelRef` lê o mês e o ano de `GerarRelatorio!B2` e `GerarRelatorio!B3`. Usa as tabelas `tblResumoRH`, `tblResumoFaltas` e `tblAjustesManuais`.
- `modExportacao` pede a data do lançamento e grava o CSV na pasta da planilha, com o nome `integracao_lancamentos_AAAAMMDD.csv`. Depende da tabela `tblCentrosCusto`, com as colunas `Setor` e `CentroCusto`.

### Importar os módulos

1. Abra o editor VBA (Alt+F11).
2. Arquivo → Importar arquivo, e selecione os `.bas` de `vba/`.
3. Salve como `.xlsm`.

## Pontos de atenção

- **Acentos:** os `.bas` foram exportados com caracteres substituídos (`�`). Isso afeta só textos da interface. Reexporte pelo editor do Excel para corrigir.
- **Nomes ilustrativos:** `SETOR A`…`SETOR E`, `Item 1`…`Item 10`, `UNIDADE X` e `UNIDADE Y` são exemplos. Troque pelos nomes reais na sua planilha e nos `Case` das funções de normalização.
- **Dependência da tabela `Tabela2`:** o módulo procura a tabela pelo nome. Ajuste `ObterTabelaTabela2` se ela tiver outro nome.
