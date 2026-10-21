Attribute VB_Name = "modExportacao"
Option Explicit

Public Sub ExportarLancamentosIntegracao()
    Static emExec As Boolean
    If emExec Then Exit Sub
    emExec = True

    On Error GoTo Finalizar
    Application.ScreenUpdating = False

    Dim dataLancamento As Date
    If Not PerguntarDataLancamentoIntegracao(dataLancamento) Then GoTo Finalizar

    Dim tabelaRh As ListObject
    Dim tabelaFaltas As ListObject
    Set tabelaRh = ObterTabelaPorNomeIntegracao("tblResumoRH")
    Set tabelaFaltas = ObterTabelaPorNomeIntegracao("tblResumoFaltas")

    If tabelaRh Is Nothing Then
        MsgBox "N�o encontrei a tabela tblResumoRH.", vbCritical
        GoTo Finalizar
    End If

    Dim resumoRhPorSetor As Object
    Dim faltasPorSetor As Object
    Set resumoRhPorSetor = CreateObject("Scripting.Dictionary")
    Set faltasPorSetor = CreateObject("Scripting.Dictionary")

    CarregarResumoRhPorDataIntegracao tabelaRh, dataLancamento, resumoRhPorSetor

    If Not tabelaFaltas Is Nothing Then
        CarregarFaltasPorDataIntegracao tabelaFaltas, dataLancamento, faltasPorSetor
    End If

    Dim caminhoSaida As String
    caminhoSaida = CaminhoArquivoSaidaIntegracao(dataLancamento)
    If caminhoSaida = vbNullString Then GoTo Finalizar

    EscreverCsvLancamentosIntegracao caminhoSaida, dataLancamento, resumoRhPorSetor, faltasPorSetor

    Application.ScreenUpdating = True
    MsgBox "Arquivo gerado com sucesso:" & vbCrLf & caminhoSaida, vbInformation

Finalizar:
    Application.ScreenUpdating = True
    emExec = False
    If Err.Number <> 0 Then
        MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
    End If
End Sub

Private Function PerguntarDataLancamentoIntegracao(ByRef dataLancamento As Date) As Boolean
    Dim resposta As Variant
    resposta = Application.InputBox( _
        Prompt:="Informe a data do lan�amento (dd/mm/aaaa):", _
        Title:="Integracao", _
        Default:=Format(Date - 1, "dd/mm/yyyy"), _
        Type:=2)

    If VarType(resposta) = vbBoolean Then
        Exit Function
    End If

    Dim textoData As String
    textoData = Trim(CStr(resposta))

    If Not IsDate(textoData) Then
        MsgBox "Data inv�lida: " & textoData, vbExclamation
        Exit Function
    End If

    dataLancamento = DateValue(textoData)
    PerguntarDataLancamentoIntegracao = True
End Function

Private Function ObterTabelaPorNomeIntegracao(ByVal nomeTabela As String) As ListObject
    Dim aba As Worksheet
    Dim tabela As ListObject

    For Each aba In ThisWorkbook.Worksheets
        For Each tabela In aba.ListObjects
            If UCase(Trim(tabela.Name)) = UCase(Trim(nomeTabela)) Then
                Set ObterTabelaPorNomeIntegracao = tabela
                Exit Function
            End If
        Next tabela
    Next aba
End Function

Private Sub CarregarResumoRhPorDataIntegracao( _
    ByVal tabelaRh As ListObject, _
    ByVal dataLancamento As Date, _
    ByVal resumoRhPorSetor As Object)

    If tabelaRh.DataBodyRange Is Nothing Then Exit Sub

    Dim indiceLinha As Long
    Dim valorData As Variant
    Dim nomeSetorNormalizado As String
    Dim autorizado As Double
    Dim efetivo As Double
    Dim valoresRh As Variant

    For indiceLinha = 1 To tabelaRh.DataBodyRange.Rows.Count
        valorData = tabelaRh.DataBodyRange.Cells(indiceLinha, 1).Value2

        If Not (IsError(valorData) Or IsNull(valorData)) Then
            If IsDate(valorData) Then
                If DateValue(CDate(valorData)) = dataLancamento Then
                    nomeSetorNormalizado = NormalizarSetorIntegracao( _
                        TextoCelulaIntegracao(tabelaRh.DataBodyRange.Cells(indiceLinha, 2).Value2))

                    If nomeSetorNormalizado <> vbNullString Then
                        autorizado = ValorNumericoIntegracao(tabelaRh.DataBodyRange.Cells(indiceLinha, 3).Value2)
                        efetivo = ValorNumericoIntegracao(tabelaRh.DataBodyRange.Cells(indiceLinha, 4).Value2)

                        If resumoRhPorSetor.Exists(nomeSetorNormalizado) Then
                            valoresRh = resumoRhPorSetor(nomeSetorNormalizado)
                            valoresRh(0) = valoresRh(0) + autorizado
                            valoresRh(1) = valoresRh(1) + efetivo
                            resumoRhPorSetor(nomeSetorNormalizado) = valoresRh
                        Else
                            resumoRhPorSetor.Add nomeSetorNormalizado, Array(autorizado, efetivo)
                        End If
                    End If
                End If
            End If
        End If
    Next indiceLinha
End Sub

Private Sub CarregarFaltasPorDataIntegracao( _
    ByVal tabelaFaltas As ListObject, _
    ByVal dataLancamento As Date, _
    ByVal faltasPorSetor As Object)

    If tabelaFaltas.DataBodyRange Is Nothing Then Exit Sub

    Dim indiceLinha As Long
    Dim valorData As Variant
    Dim nomeSetorNormalizado As String
    Dim faltas As Double

    For indiceLinha = 1 To tabelaFaltas.DataBodyRange.Rows.Count
        valorData = tabelaFaltas.DataBodyRange.Cells(indiceLinha, 1).Value2

        If Not (IsError(valorData) Or IsNull(valorData)) Then
            If IsDate(valorData) Then
                If DateValue(CDate(valorData)) = dataLancamento Then
                    nomeSetorNormalizado = NormalizarSetorIntegracao( _
                        TextoCelulaIntegracao(tabelaFaltas.DataBodyRange.Cells(indiceLinha, 2).Value2))

                    If nomeSetorNormalizado <> vbNullString Then
                        faltas = ValorNumericoIntegracao(tabelaFaltas.DataBodyRange.Cells(indiceLinha, 3).Value2)

                        If faltasPorSetor.Exists(nomeSetorNormalizado) Then
                            faltasPorSetor(nomeSetorNormalizado) = CDbl(faltasPorSetor(nomeSetorNormalizado)) + faltas
                        Else
                            faltasPorSetor.Add nomeSetorNormalizado, faltas
                        End If
                    End If
                End If
            End If
        End If
    Next indiceLinha
End Sub

Private Function TextoCelulaIntegracao(ByVal valorCelula As Variant) As String
    If IsError(valorCelula) Or IsNull(valorCelula) Or IsEmpty(valorCelula) Then
        Exit Function
    End If

    TextoCelulaIntegracao = Trim(CStr(valorCelula))
End Function

Private Function ValorNumericoIntegracao(ByVal valorCelula As Variant) As Double
    If IsError(valorCelula) Or IsNull(valorCelula) Or IsEmpty(valorCelula) Then
        Exit Function
    End If

    If IsNumeric(valorCelula) Then
        ValorNumericoIntegracao = CDbl(valorCelula)
    End If
End Function

Private Function NormalizarSetorIntegracao(ByVal nomeSetor As String) As String
    Dim textoSetor As String
    textoSetor = UCase(Trim(nomeSetor))
    textoSetor = Replace(textoSetor, vbTab, " ")
    textoSetor = Replace(textoSetor, Chr(160), " ")
    textoSetor = CStr(Application.WorksheetFunction.Trim(textoSetor))

    Select Case True
        Case InStr(textoSetor, "SETOR A") > 0
            NormalizarSetorIntegracao = "SETOR A"

        Case InStr(textoSetor, "SETOR B") > 0
            NormalizarSetorIntegracao = "SETOR B"

        Case InStr(textoSetor, "SETOR C") > 0
            NormalizarSetorIntegracao = "SETOR C"

        Case InStr(textoSetor, "SETOR D") > 0 Or InStr(textoSetor, "SETOR D") > 0
            NormalizarSetorIntegracao = "SETOR D"

        Case InStr(textoSetor, "SETOR E") > 0
            NormalizarSetorIntegracao = "SETOR E"

        Case Else
            NormalizarSetorIntegracao = vbNullString
    End Select
End Function

Private Function CaminhoArquivoSaidaIntegracao(ByVal dataLancamento As Date) As String
    Dim pastaDestino As String
    pastaDestino = ThisWorkbook.Path

    If pastaDestino = vbNullString Then
        pastaDestino = Environ$("USERPROFILE") & "\Desktop"
    End If

    CaminhoArquivoSaidaIntegracao = pastaDestino & _
        "\integracao_lancamentos_" & Format(dataLancamento, "yyyymmdd") & ".csv"
End Function

Private Sub EscreverCsvLancamentosIntegracao( _
    ByVal caminhoArquivo As String, _
    ByVal dataLancamento As Date, _
    ByVal resumoRhPorSetor As Object, _
    ByVal faltasPorSetor As Object)

    Dim setores As Variant
    setores = ListaSetoresIntegracao()

    Dim numeroArquivo As Integer
    numeroArquivo = FreeFile

    Dim arquivoAberto As Boolean
    Dim indiceSetor As Long
    Dim nomeSetor As String
    Dim centroCusto As String
    Dim autorizado As Double
    Dim efetivo As Double
    Dim faltas As Double
    Dim saldoRef As Double
    Dim valoresRh As Variant

    On Error GoTo ErroArquivo

    Open caminhoArquivo For Output As #numeroArquivo
    arquivoAberto = True

    Print #numeroArquivo, "data_falta,filial,setor_referencia,centro_custo,saldo_ref,quantidade_faltas"

    For indiceSetor = LBound(setores) To UBound(setores)
        nomeSetor = setores(indiceSetor)
        centroCusto = CentroCustoPorSetorIntegracao(nomeSetor)

        autorizado = 0#
        efetivo = 0#
        faltas = 0#

        If resumoRhPorSetor.Exists(nomeSetor) Then
            valoresRh = resumoRhPorSetor(nomeSetor)
            autorizado = CDbl(valoresRh(0))
            efetivo = CDbl(valoresRh(1))
        End If

        If faltasPorSetor.Exists(nomeSetor) Then
            faltas = CDbl(faltasPorSetor(nomeSetor))
        End If

        saldoRef = efetivo - autorizado

        Print #numeroArquivo, _
    Format(dataLancamento, "dd/mm/yyyy") & "," & _
    "Unidade X" & "," & _
    nomeSetor & "," & _
    centroCusto & "," & _
    TextoInteiroIntegracao(saldoRef) & "," & _
    TextoInteiroIntegracao(faltas)
    Next indiceSetor

    Close #numeroArquivo
    arquivoAberto = False
    Exit Sub

ErroArquivo:
    If arquivoAberto Then Close #numeroArquivo
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ListaSetoresIntegracao() As Variant
    ListaSetoresIntegracao = Array("SETOR A", "SETOR B", "SETOR C", "SETOR D", "SETOR E")
End Function

Private Function CentroCustoPorSetorIntegracao(ByVal nomeSetorNormalizado As String) As String
    Dim tabelaCentros As ListObject
    Dim indiceLinha As Long
    Set tabelaCentros = ObterTabelaPorNomeIntegracao("tblCentrosCusto")
    If tabelaCentros Is Nothing Then Exit Function
    If tabelaCentros.DataBodyRange Is Nothing Then Exit Function

    For indiceLinha = 1 To tabelaCentros.DataBodyRange.Rows.Count
        If UCase(Trim(CStr(tabelaCentros.DataBodyRange.Cells(indiceLinha, 1).Value2))) = nomeSetorNormalizado Then
            CentroCustoPorSetorIntegracao = Trim(CStr(tabelaCentros.DataBodyRange.Cells(indiceLinha, 2).Value2))
            Exit Function
        End If
    Next indiceLinha
End Function

Private Function TextoInteiroIntegracao(ByVal valorNumerico As Double) As String
    ' Converte sempre para inteiro, arredondando se vier fracionado.
    ' Usamos CLng em vez de CInt porque CInt estoura em valores maiores.
    ' O resultado sai como "2", "-3", "0" sem ponto nem v�rgula.
    TextoInteiroIntegracao = CStr(CLng(Round(valorNumerico, 0)))
End Function
