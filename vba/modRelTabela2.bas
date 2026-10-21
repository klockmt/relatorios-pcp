Attribute VB_Name = "modRelTabela2"
Option Explicit

Public Function GerarAbaRelTabela2(ByVal mesAlvo As Long, ByVal anoAlvo As Long) As Boolean
    Dim tabelaTabela2 As ListObject
    Set tabelaTabela2 = ObterTabelaTabela2("Tabela2") ' Ajuste para o nome real da tabela do TABELA2.

    If tabelaTabela2 Is Nothing Then
        MsgBox "N�o achei a tabela do TABELA2.", vbExclamation
        Exit Function
    End If

    If tabelaTabela2.DataBodyRange Is Nothing Then
        MsgBox "A tabela do TABELA2 est� vazia.", vbExclamation
        Exit Function
    End If

    Dim matrizTabela2 As Variant
    matrizTabela2 = tabelaTabela2.DataBodyRange.Value

    Dim serialAlvo As Long
    serialAlvo = CLng(DateSerial(anoAlvo, mesAlvo, 1))

    Dim indiceColuna As Long
    Dim seriaisMes As Object
    Set seriaisMes = CreateObject("Scripting.Dictionary")

    For indiceColuna = 3 To tabelaTabela2.ListColumns.Count
        If MesTextoParaSerial(tabelaTabela2.ListColumns(indiceColuna).Name) = serialAlvo Then
            seriaisMes(indiceColuna) = serialAlvo
        End If
    Next indiceColuna

    If seriaisMes.Count = 0 Then
        MsgBox "Nenhuma coluna de m�s para " & mesAlvo & "/" & anoAlvo & " na tabela do TABELA2.", vbExclamation
        Exit Function
    End If

    Dim ordemColunas As Variant
    ordemColunas = OrdenarPorSerial(seriaisMes)

    Dim abaTabela2 As Worksheet
    Set abaTabela2 = ObterOuCriarAbaTabela2()

    abaTabela2.Cells.Clear
    EscreverCabecalhoTabela2 abaTabela2, ordemColunas, tabelaTabela2

    Dim categoriasOrdem As Variant
    categoriasOrdem = Array("setor_a", "aduela maci�a", "setor_b", "setor_c", "m�dulo", "setor_e")

    Dim indiceCategoria As Long
    Dim linhaAtual As Long
    linhaAtual = 2

    For indiceCategoria = LBound(categoriasOrdem) To UBound(categoriasOrdem)
        linhaAtual = EscreverBlocoCategoriaTabela2( _
            abaTabela2, _
            CStr(categoriasOrdem(indiceCategoria)), _
            matrizTabela2, _
            ordemColunas, _
            linhaAtual)
    Next indiceCategoria

    GerarAbaRelTabela2 = True
End Function

Private Function ObterOuCriarAbaTabela2() As Worksheet
    Dim indiceAba As Long
    Dim nomeAtual As String
    Dim abaDepois As Worksheet

    Application.DisplayAlerts = False

    For indiceAba = ThisWorkbook.Worksheets.Count To 1 Step -1
        nomeAtual = UCase(Trim(ThisWorkbook.Worksheets(indiceAba).Name))

        If nomeAtual = "RELTABELA2" Then
            ThisWorkbook.Worksheets(indiceAba).Delete
        End If
    Next indiceAba

    Application.DisplayAlerts = True

    On Error Resume Next
    Set abaDepois = ThisWorkbook.Worksheets("RelRef")
    On Error GoTo 0

    If abaDepois Is Nothing Then
        Set abaDepois = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    End If

    Set ObterOuCriarAbaTabela2 = ThisWorkbook.Sheets.Add(After:=abaDepois)
    ObterOuCriarAbaTabela2.Name = "RelTabela2"
End Function

Private Sub EscreverCabecalhoTabela2(ByVal abaTabela2 As Worksheet, ByVal ordemColunas As Variant, ByVal tabelaTabela2 As ListObject)
    Dim posicao As Long
    abaTabela2.Cells(1, 1).Value = "Categoria / Etiqueta"
    For posicao = LBound(ordemColunas) To UBound(ordemColunas)
        abaTabela2.Cells(1, posicao + 2).Value = tabelaTabela2.ListColumns(CLng(ordemColunas(posicao))).Name
    Next posicao
    abaTabela2.Range(abaTabela2.Cells(1, 1), abaTabela2.Cells(1, UBound(ordemColunas) + 2)).Interior.Color = RGB(237, 237, 237)
    abaTabela2.Rows(1).Font.Bold = True
End Sub

Private Function EscreverBlocoCategoriaTabela2(ByVal abaTabela2 As Worksheet, ByVal nomeCategoria As String, ByVal matrizTabela2 As Variant, ByVal ordemColunas As Variant, ByVal linhaInicial As Long) As Long
    Dim linhaMatriz As Long, linhaAtual As Long
    If ContarLinhasCategoria(matrizTabela2, nomeCategoria) = 0 Then
        EscreverBlocoCategoriaTabela2 = linhaInicial
        Exit Function
    End If
    abaTabela2.Cells(linhaInicial, 1).Value = StrConv(nomeCategoria, vbProperCase)
    linhaAtual = linhaInicial + 1
    For linhaMatriz = 1 To UBound(matrizTabela2, 1)
        If LCase(Trim(CStr(matrizTabela2(linhaMatriz, 1)))) = nomeCategoria Then
            abaTabela2.Cells(linhaAtual, 1).Value = matrizTabela2(linhaMatriz, 2)
            EscreverValoresLinhaTabela2 abaTabela2, linhaAtual, matrizTabela2, linhaMatriz, ordemColunas
            linhaAtual = linhaAtual + 1
        End If
    Next linhaMatriz
    abaTabela2.Cells(linhaInicial, 1).Interior.Color = RGB(146, 208, 80)
    abaTabela2.Cells(linhaInicial, 1).Font.Bold = True
    abaTabela2.Range(abaTabela2.Cells(linhaInicial + 1, 1), abaTabela2.Cells(linhaAtual - 1, 1)).Interior.Color = RGB(219, 219, 219)
    EscreverBlocoCategoriaTabela2 = linhaAtual + 1
End Function

Private Function ContarLinhasCategoria(ByVal matrizTabela2 As Variant, ByVal nomeCategoria As String) As Long
    Dim linhaMatriz As Long
    For linhaMatriz = 1 To UBound(matrizTabela2, 1)
        If LCase(Trim(CStr(matrizTabela2(linhaMatriz, 1)))) = nomeCategoria Then
            ContarLinhasCategoria = ContarLinhasCategoria + 1
        End If
    Next linhaMatriz
End Function

Private Sub EscreverValoresLinhaTabela2(ByVal abaTabela2 As Worksheet, ByVal linhaAlvo As Long, ByVal matrizTabela2 As Variant, ByVal linhaMatriz As Long, ByVal ordemColunas As Variant)
    Dim posicao As Long
    For posicao = LBound(ordemColunas) To UBound(ordemColunas)
        If IsNumeric(matrizTabela2(linhaMatriz, CLng(ordemColunas(posicao)))) Then
            abaTabela2.Cells(linhaAlvo, posicao + 2).Value = matrizTabela2(linhaMatriz, CLng(ordemColunas(posicao)))
        End If
    Next posicao
End Sub

Private Function MesTextoParaSerial(ByVal textoMes As String) As Long
    Dim partes As Variant
    Dim mesesAbrev As Variant
    Dim indiceMes As Long
    partes = Split(Trim(textoMes), "/")
    If UBound(partes) < 1 Then Exit Function
    mesesAbrev = Array("jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez")
    For indiceMes = 0 To 11
        If LCase(Trim(CStr(partes(0)))) = mesesAbrev(indiceMes) Then
            MesTextoParaSerial = CLng(DateSerial(2000 + CLng(partes(1)), indiceMes + 1, 1))
            Exit Function
        End If
    Next indiceMes
End Function

Private Function OrdenarPorSerial(ByVal seriaisMes As Object) As Variant
    Dim chaves As Variant
    Dim i As Long, j As Long, tmp As Variant
    chaves = seriaisMes.Keys
    For i = LBound(chaves) To UBound(chaves) - 1
        For j = i + 1 To UBound(chaves)
            If seriaisMes(chaves(i)) > seriaisMes(chaves(j)) Then
                tmp = chaves(i): chaves(i) = chaves(j): chaves(j) = tmp
            End If
        Next j
    Next i
    OrdenarPorSerial = chaves
End Function

Private Function ObterTabelaTabela2(nome As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If UCase(Trim(lo.Name)) = UCase(Trim(nome)) Then
                Set ObterTabelaTabela2 = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function
