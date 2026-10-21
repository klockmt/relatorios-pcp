Attribute VB_Name = "modRelRef"
Option Explicit

Public Sub GerarRelatorioMensal(Optional ByVal incluirTabela2 As Boolean = True)
    Static emExec As Boolean
    If emExec Then Exit Sub
    emExec = True
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Finalizar
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("GerarRelatorio")
    Dim mes As Long: mes = LerParametroMes()
    Dim ano As Long: ano = LerParametroAno()
    If mes = 0 Or ano = 0 Then GoTo Finalizar
    
    Dim loRh As ListObject: Set loRh = ObterTabela("tblResumoRH")
    Dim loF As ListObject: Set loF = ObterTabela("tblResumoFaltas")
    If loRh Is Nothing Then MsgBox "N�o achei tblResumoRH": GoTo Finalizar
    
    Dim dictRh As Object: Set dictRh = CreateObject("Scripting.Dictionary")
    Dim dictF As Object: Set dictF = CreateObject("Scripting.Dictionary")
    Dim dictDias As Object: Set dictDias = CreateObject("Scripting.Dictionary")
    Dim dictAjustes As Object: Set dictAjustes = CriarDictAjustes()
    Dim dictSemana As Object
    Dim dictSomaSemana As Object: Set dictSomaSemana = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, dt As Date, setor As String, chave As String
    
    For i = 1 To loRh.DataBodyRange.Rows.Count
        dt = CDate(loRh.DataBodyRange.Cells(i, 1).Value2)
        If Month(dt) = mes And Year(dt) = ano Then
            setor = Trim(CStr(loRh.DataBodyRange.Cells(i, 2).Value2))
            chave = setor & "|" & CLng(dt)
            dictRh(chave) = Array(loRh.DataBodyRange.Cells(i, 3).Value2, loRh.DataBodyRange.Cells(i, 4).Value2)
            dictDias(CLng(dt)) = True
        End If
    Next i
    
    If Not loF Is Nothing Then
        For i = 1 To loF.DataBodyRange.Rows.Count
            dt = CDate(loF.DataBodyRange.Cells(i, 1).Value2)
            If Month(dt) = mes And Year(dt) = ano Then
                setor = Trim(CStr(loF.DataBodyRange.Cells(i, 2).Value2))
                chave = setor & "|" & CLng(dt)
                dictF(chave) = loF.DataBodyRange.Cells(i, 3).Value2
                dictDias(CLng(dt)) = True
            End If
        Next i
    End If
    
    If dictDias.Count = 0 Then MsgBox "Nenhum dia encontrado para " & mes & "/" & ano, vbExclamation: GoTo Finalizar
    
    Dim diasOrdenados As Variant
    diasOrdenados = dictDias.Keys
    Dim a As Long, b As Long, tmp As Variant
    For a = LBound(diasOrdenados) To UBound(diasOrdenados) - 1
        For b = a + 1 To UBound(diasOrdenados)
            If CLng(diasOrdenados(a)) > CLng(diasOrdenados(b)) Then
                tmp = diasOrdenados(a)
                diasOrdenados(a) = diasOrdenados(b)
                diasOrdenados(b) = tmp
            End If
        Next b
    Next a
    
    Set dictSemana = CriarDictSemana(diasOrdenados)
    Dim totalSemanas As Long: totalSemanas = dictSemana(CLng(diasOrdenados(UBound(diasOrdenados))))
    Dim qtdSemana As Variant
    ReDim qtdSemana(1 To totalSemanas) As Long
    For i = LBound(diasOrdenados) To UBound(diasOrdenados)
        qtdSemana(dictSemana(CLng(diasOrdenados(i)))) = qtdSemana(dictSemana(CLng(diasOrdenados(i)))) + 1
    Next i
    
    Dim abaRel As Worksheet: Set abaRel = ObterAbaRelatorio()
    
    Dim setores As Variant: setores = Array("Setor A", "Setor B", "Setor C", "Setor D", "Setor E", "UNIDADE X")
    Dim metricasManuais As Variant: metricasManuais = Array("HE", "EMPRESTIMO")
    Dim s As Long, linha As Long, col As Long, diaLong As Variant
    Dim r As Long, semanaAtual As Long, numDias As Long
    Dim temCarga As Boolean
    Dim somaMetrica(1 To 6) As Double
    Dim valoresDia As Variant
    linha = 1
    
    For s = LBound(setores) To UBound(setores)
        abaRel.Cells(linha, 1).Value = setores(s)
        abaRel.Cells(linha + 1, 1).Value = "Autorizado"
        abaRel.Cells(linha + 2, 1).Value = "Efetivo"
        abaRel.Cells(linha + 3, 1).Value = "Faltas"
        abaRel.Cells(linha + 4, 1).Value = "HE"
        abaRel.Cells(linha + 5, 1).Value = "Empr�stimo"
        abaRel.Cells(linha + 6, 1).Value = "Total Trabalhou"
        
        For r = 1 To 6
            somaMetrica(r) = 0
        Next r
        
        col = 2
        For Each diaLong In diasOrdenados
            dt = CDate(CLng(diaLong))
            abaRel.Cells(linha, col).Value = dt
            abaRel.Cells(linha, col).NumberFormat = "dd/mm"
            
            chave = setores(s) & "|" & CLng(dt)
            semanaAtual = dictSemana(CLng(dt))
            valoresDia = ValoresDoDia(setores(s), CLng(dt), dictRh, dictF, dictAjustes, metricasManuais)
            temCarga = dictRh.Exists(chave)
            If setores(s) = "UNIDADE X" Then temCarga = ExisteCargaUnidade(dictRh, CLng(dt))
            
            If temCarga Then
                abaRel.Cells(linha + 1, col).Value = valoresDia(1)
                abaRel.Cells(linha + 2, col).Value = valoresDia(2)
                abaRel.Cells(linha + 3, col).Value = valoresDia(3)
                abaRel.Cells(linha + 6, col).Value = valoresDia(6)
            End If
            For r = 4 To 5
                abaRel.Cells(linha + r, col).Value = valoresDia(r)
            Next r
            For r = 1 To 6
                somaMetrica(r) = somaMetrica(r) + valoresDia(r)
                dictSomaSemana(setores(s) & "|" & r & "|" & semanaAtual) = dictSomaSemana(setores(s) & "|" & r & "|" & semanaAtual) + valoresDia(r)
            Next r
            col = col + 1
        Next diaLong
        
        numDias = UBound(diasOrdenados) - LBound(diasOrdenados) + 1
        EscreverColunaTotal abaRel, linha, somaMetrica, numDias, col
        EscreverColunasSemanas abaRel, linha, setores(s), totalSemanas, qtdSemana, dictSomaSemana, col + 5
        AplicarFormatacaoBloco abaRel, linha, col + 1, col + 5, totalSemanas
        linha = linha + 10
    Next s
    
    If incluirTabela2 Then Call GerarAbaRelTabela2(mes, ano)
    MsgBox "Gerado! " & dictDias.Count & " dias em " & mes & "/" & ano, vbInformation

Finalizar:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    emExec = False
    If Err.Number <> 0 Then MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Function LerParametroMes() As Long
    Dim valorLido As Variant
    valorLido = ThisWorkbook.Worksheets("GerarRelatorio").Range("B2").Value
    If IsNumeric(valorLido) And valorLido >= 1 And valorLido <= 12 Then
        LerParametroMes = CLng(valorLido)
        Exit Function
    End If
    MsgBox "Mes invalido em GerarRelatorio!B2.", vbExclamation
End Function

Private Function LerParametroAno() As Long
    Dim valorLido As Variant
    valorLido = ThisWorkbook.Worksheets("GerarRelatorio").Range("B3").Value
    If IsNumeric(valorLido) And valorLido >= 2015 And valorLido <= 2100 Then
        LerParametroAno = CLng(valorLido)
        Exit Function
    End If
    MsgBox "Ano invalido em GerarRelatorio!B3.", vbExclamation
End Function

Private Function CriarDictSemana(ByVal diasOrdenados As Variant) As Object
    Dim dictSemana As Object
    Dim indice As Long
    Dim semanaAtual As Long
    Set dictSemana = CreateObject("Scripting.Dictionary")
    semanaAtual = 1
    For indice = LBound(diasOrdenados) To UBound(diasOrdenados)
        If indice > LBound(diasOrdenados) Then
            If CLng(diasOrdenados(indice)) - CLng(diasOrdenados(indice - 1)) > 2 Then semanaAtual = semanaAtual + 1
        End If
        dictSemana(CLng(diasOrdenados(indice))) = semanaAtual
    Next indice
    Set CriarDictSemana = dictSemana
End Function

Private Function CriarDictAjustes() As Object
    Dim tabelaAjustes As ListObject
    Dim matrizAjustes As Variant
    Dim indiceLinha As Long
    Set CriarDictAjustes = CreateObject("Scripting.Dictionary")
    Set tabelaAjustes = ObterTabela("tblAjustesManuais")
    If tabelaAjustes Is Nothing Then Exit Function
    If tabelaAjustes.DataBodyRange Is Nothing Then Exit Function
    matrizAjustes = tabelaAjustes.DataBodyRange.Value
    For indiceLinha = 1 To UBound(matrizAjustes, 1)
        CriarDictAjustes(Trim(CStr(matrizAjustes(indiceLinha, 2))) & "|" & UCase(Trim(CStr(matrizAjustes(indiceLinha, 3)))) & "|" & CLng(CDate(matrizAjustes(indiceLinha, 1)))) = CDbl(matrizAjustes(indiceLinha, 4))
    Next indiceLinha
End Function

Private Function ValorAjuste(ByVal dictAjustes As Object, ByVal nomeSetor As String, ByVal nomeMetrica As String, ByVal serialDia As Long) As Double
    Dim chaveAjuste As String
    chaveAjuste = nomeSetor & "|" & nomeMetrica & "|" & serialDia
    If dictAjustes.Exists(chaveAjuste) Then ValorAjuste = dictAjustes(chaveAjuste)
End Function

Private Function ListaSetoresBase() As Variant
    ' Lista oficial de setores; UNIDADE X eh sempre derivada desta lista
    ListaSetoresBase = Array("Setor A", "Setor B", "Setor C", "Setor D", "Setor E")
End Function

Private Function ValoresDoDia(ByVal nomeSetor As String, ByVal serialDia As Long, ByVal dictRh As Object, ByVal dictF As Object, ByVal dictAjustes As Object, ByVal metricasManuais As Variant) As Variant
    Dim valores(1 To 6) As Double
    Dim r As Long
    Dim chaveDia As String
    chaveDia = nomeSetor & "|" & serialDia
    If nomeSetor = "UNIDADE X" Then
        valores(1) = SomaRhUnidade(dictRh, serialDia, 0)
        valores(2) = SomaRhUnidade(dictRh, serialDia, 1)
        valores(3) = SomaFaltasUnidade(dictF, serialDia)
    ElseIf dictRh.Exists(chaveDia) Then
        valores(1) = dictRh(chaveDia)(0)
        valores(2) = dictRh(chaveDia)(1)
        If dictF.Exists(chaveDia) Then valores(3) = dictF(chaveDia)
    End If
    For r = 4 To 5
        valores(r) = ValorManualDia(dictAjustes, nomeSetor, metricasManuais(r - 4), serialDia)
    Next r
    valores(6) = valores(2) + valores(4) + valores(5) - valores(3)
    ValoresDoDia = valores
End Function

Private Function ValorManualDia(ByVal dictAjustes As Object, ByVal nomeSetor As String, ByVal nomeMetrica As String, ByVal serialDia As Long) As Double
    If nomeSetor = "UNIDADE X" Then
        ValorManualDia = SomaAjusteUnidade(dictAjustes, nomeMetrica, serialDia)
    Else
        ValorManualDia = ValorAjuste(dictAjustes, nomeSetor, nomeMetrica, serialDia)
    End If
End Function

Private Function SomaRhUnidade(ByVal dictRh As Object, ByVal serialDia As Long, ByVal indiceValor As Long) As Double
    Dim setoresBase As Variant
    Dim iBase As Long
    Dim chaveBase As String
    setoresBase = ListaSetoresBase()
    For iBase = LBound(setoresBase) To UBound(setoresBase)
        chaveBase = setoresBase(iBase) & "|" & serialDia
        If dictRh.Exists(chaveBase) Then SomaRhUnidade = SomaRhUnidade + dictRh(chaveBase)(indiceValor)
    Next iBase
End Function

Private Function SomaFaltasUnidade(ByVal dictF As Object, ByVal serialDia As Long) As Double
    Dim setoresBase As Variant
    Dim iBase As Long
    Dim chaveBase As String
    setoresBase = ListaSetoresBase()
    For iBase = LBound(setoresBase) To UBound(setoresBase)
        chaveBase = setoresBase(iBase) & "|" & serialDia
        If dictF.Exists(chaveBase) Then SomaFaltasUnidade = SomaFaltasUnidade + dictF(chaveBase)
    Next iBase
End Function

Private Function SomaAjusteUnidade(ByVal dictAjustes As Object, ByVal nomeMetrica As String, ByVal serialDia As Long) As Double
    Dim setoresBase As Variant
    Dim iBase As Long
    setoresBase = ListaSetoresBase()
    For iBase = LBound(setoresBase) To UBound(setoresBase)
        SomaAjusteUnidade = SomaAjusteUnidade + ValorAjuste(dictAjustes, setoresBase(iBase), nomeMetrica, serialDia)
    Next iBase
End Function

Private Function ExisteCargaUnidade(ByVal dictRh As Object, ByVal serialDia As Long) As Boolean
    Dim setoresBase As Variant
    Dim iBase As Long
    setoresBase = ListaSetoresBase()
    For iBase = LBound(setoresBase) To UBound(setoresBase)
        If dictRh.Exists(setoresBase(iBase) & "|" & serialDia) Then
            ExisteCargaUnidade = True
            Exit Function
        End If
    Next iBase
End Function

Private Sub EscreverColunaTotal(ByVal abaAlvo As Worksheet, ByVal linhaInicial As Long, ByVal somaMetrica As Variant, ByVal numDias As Long, ByVal colInicial As Long)
    Dim r As Long
    abaAlvo.Cells(linhaInicial, colInicial).Value = "TOTAL"
    abaAlvo.Cells(linhaInicial, colInicial + 1).Value = "MEDIA"
    For r = 1 To 6
        abaAlvo.Cells(linhaInicial + r, colInicial).Value = somaMetrica(r)
        abaAlvo.Cells(linhaInicial + r, colInicial + 1).Value = somaMetrica(r) / numDias
        abaAlvo.Cells(linhaInicial + r, colInicial + 1).NumberFormat = "0.0"
    Next r
End Sub

Private Sub EscreverColunasSemanas(ByVal abaAlvo As Worksheet, ByVal linhaInicial As Long, ByVal nomeSetor As String, ByVal totalSemanas As Long, ByVal qtdSemana As Variant, ByVal dictSomaSemana As Object, ByVal colInicial As Long)
    Dim r As Long, w As Long
    Dim colSem As Long
    Dim valorSemana As Double
    colSem = colInicial
    For w = 1 To totalSemanas
        abaAlvo.Cells(linhaInicial, colSem).Value = "Sem " & w & " Tot"
        abaAlvo.Cells(linhaInicial, colSem + 1).Value = "Sem " & w & " Med"
        For r = 1 To 6
            valorSemana = CDbl(dictSomaSemana(nomeSetor & "|" & r & "|" & w))
            abaAlvo.Cells(linhaInicial + r, colSem).Value = valorSemana
            abaAlvo.Cells(linhaInicial + r, colSem + 1).Value = valorSemana / qtdSemana(w)
            abaAlvo.Cells(linhaInicial + r, colSem + 1).NumberFormat = "0.0"
        Next r
        colSem = colSem + 3
    Next w
End Sub

Private Function ObterAbaRelatorio() As Worksheet
    Dim indiceAba As Long
    Dim nomeAtual As String
    Application.DisplayAlerts = False
    For indiceAba = ThisWorkbook.Worksheets.Count To 1 Step -1
        nomeAtual = UCase(Trim(ThisWorkbook.Worksheets(indiceAba).Name))
        If nomeAtual = "RELREF" Or nomeAtual = "RELATORIOMENSAL" Then
            ThisWorkbook.Worksheets(indiceAba).Delete
        End If
    Next indiceAba
    Application.DisplayAlerts = True
    Set ObterAbaRelatorio = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Worksheets("GerarRelatorio"))
    ObterAbaRelatorio.Name = "RelRef"
End Function

Private Function ObterTabela(nome As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If UCase(Trim(lo.Name)) = UCase(Trim(nome)) Then
                Set ObterTabela = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Sub AplicarFormatacaoBloco(ByVal abaAlvo As Worksheet, ByVal linhaInicial As Long, ByVal ultimaColunaContinua As Long, ByVal colSemanaInicial As Long, ByVal totalSemanas As Long)
    Dim corSetor As Long
    Dim corDatas As Long
    Dim corRotulos As Long
    Dim w As Long
    Dim colSem As Long
    corSetor = RGB(146, 208, 80)
    corDatas = RGB(237, 237, 237)
    corRotulos = RGB(219, 219, 219)
    abaAlvo.Cells(linhaInicial, 1).Interior.Color = corSetor
    abaAlvo.Cells(linhaInicial, 1).Font.Bold = True
    abaAlvo.Range(abaAlvo.Cells(linhaInicial, 2), abaAlvo.Cells(linhaInicial, ultimaColunaContinua)).Interior.Color = corDatas
    abaAlvo.Range(abaAlvo.Cells(linhaInicial + 1, 1), abaAlvo.Cells(linhaInicial + 6, 1)).Interior.Color = corRotulos
    For w = 1 To totalSemanas
        colSem = colSemanaInicial + ((w - 1) * 3)
        abaAlvo.Cells(linhaInicial, colSem).Interior.Color = corDatas
        abaAlvo.Cells(linhaInicial, colSem + 1).Interior.Color = corDatas
    Next w
End Sub
Public Sub GerarRelatorioRefMensal()
    GerarRelatorioMensal False
End Sub

Public Sub GerarRelatorioTabela2Mensal()
    Static emExec As Boolean
    If emExec Then Exit Sub
    emExec = True

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Finalizar

    Dim mes As Long, ano As Long

    mes = LerParametroMes()
    If mes = 0 Then GoTo Finalizar

    ano = LerParametroAno()
    If ano = 0 Then GoTo Finalizar

    If GerarAbaRelTabela2(mes, ano) Then
        MsgBox "Relat�rio TABELA2 gerado para " & mes & "/" & ano & ".", vbInformation
    End If

Finalizar:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    emExec = False

    If Err.Number <> 0 Then
        MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
    End If
End Sub
