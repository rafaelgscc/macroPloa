Option Explicit 'determina que todas as variáveis precisam ser declaradas

'Declarando as variáveis string que controlam informações do Workbook ativo e do Arquivo Processado
Public sWBMacro As String, sPlanAtiva As String, sCaminho As String, sArquivo As String, sPlanWBArquivo As String, _
       sFuncional As String, sContaCorrente As String, sCodMT As String, sFonteSOF As String, sCelulaNatureza As String, _
       sAno As String, sRP As String, sNatureza As String, sUGResp As String, sData As String
 
'Declarando as Variáveis Inteiras relacionadas ao Workbook Macro no processamento dos arquivos
Public iLinhaCabecalho As Integer, iLinha As Long, lSeq As Long, sMes As Integer

'inicializando a Macro
Sub Macro_PLOAWEB()

    Dim Confirma
    Confirma = MsgBox("Deseja iniciar o procedimento: Formatar arquivos do PLOAWEB?", vbOKCancel, "Confirmação")
    If Confirma = vbCancel Then
        Call Cancelada
        Exit Sub
    End If
        
    'Inicializando as Variáveis
    Call InicializandoVariaveis
    
    'Inicializando o processamento
    Call TB_IMP_EXECUCAO(4)
    Call TB_IMP_EXECUCAO_PTRES(5)
    Call TB_IMP_EXECUCAO_RP(6)
    Call TB_IMP_BLOQUEIO_NATUREZA(7)
    Call TB_IMP_DESTAQUE(8)
    Call TB_IMP_DESTAQUE_PTRES(9)
    Call TB_IMP_INDISPONIVEL(10)
    Call TB_IMP_LOA_DETALHADA(11)
    Call TB_IMP_RECEITAS(12)
    Call TB_IMP_HISTORICO_PGTO(13)
    Call TB_IMP_HISTORICO_PGTO2(14)
    Call TB_IMP_SRE_ADM(15)
    Call TB_IMP_PGTO(16)
    Call TB_IMP_RAP_EMPENHO(17)
    Call TB_IMP_LIMITE(18)
    Call TB_IMP_RECEBIDO(19)
    Call TB_IMP_RECEBIDO_PTRES(20)
    Call TB_IMP_RECEBIDO_RP(21)
    Call TB_IMP_CREDITOS(22)
    Call TB_IMP_UG_DESTAQUE(23)
    Call TB_IMP_NATUREZADETALHADA(24)
    Call TB_IMP_DISPONIVELSRE(25)
    Call TB_IMP_CORFIN_NE_DIARIO(26)
    Call TB_IMP_NATUREZA_REC(27)
    Call TB_IMP_NATUREZA_PTRES(28)
    Call TB_IMP_NE_2022_0(29)
    Call TB_IMP_NE_2022_1(30)
    Call TB_IMP_NE_2023_0(31)
    Call TB_IMP_NE_2023_1(32)
    Call JUNTARVALORESEXECUCAO(33)
    Call JUNTARVALORESEXECUCAOPTRES(34)
    Call JUNTARVALORESEXECUCAORP(35)
    Call TB_IMP_NE_2018_6(36)
    Call TB_IMP_NE_2019_4(37)
    Call TB_IMP_NE_2020_2(38)
    Call TB_IMP_NE_2020_3(39)
    Call TB_IMP_NE_2021_0(40)
    Call TB_IMP_NE_2021_1(41)
    Call TB_IMP_UGEXECUTORA(42)
    Call TB_IMP_BLOQUEIO(43)
    Call JUNTARVALORESNATUREZA(44)
    
    
    'Finalizando a macro
    Call Encerrar

End Sub

Sub Encerrar()
    ''' Encerrar a Execução da Macro'''
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Operação concluída!"
End Sub

Sub Cancelada()
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Operação Cancelada pelo Uruário!"
End Sub

Sub InicializandoVariaveis()
    'Inicializando as variáveis que controlam informações do Workbook ativo
    sWBMacro = ActiveWorkbook.Name
    sPlanAtiva = ActiveSheet.Name
    sCaminho = ActiveWorkbook.Path
            
    ''' Alterando o status da tela antes de iniciar o processamento '''
    Application.ScreenUpdating = False
    Application.StatusBar = "Iniciando..."
End Sub


Sub TB_IMP_EXECUCAO(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                   
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano,
        '3)9(VL_Orc_Ini), 4)10(VL_Cred_Sup), 5)11(VL_Cred_Ext), 6)12(VL_Cred_Esp), 7)13(VL_Orc_Aut), 8)14(VL_Cred_Can),
        '9)19(VL_Disponivel), 10)20(VL_Contido), 11)23(VL_Empenhado), 12)25(VL_Liquidado), 13)26(VL_Pago),
        '14)35(RP_Pro_Inscrito), 15)36(RP_Pro_Reinscrito), 16)37(RP_Pro_Cancelado), 17)38(RP_Pro_Pago),
        '18)39(RP_Pro_APagar), 19)40(RP_NPro_Inscrito), 20)41(RP_NPro_Reinscrito), 21)42(RP_NPro_Cancelado), 22)43(RP_NPro_ALiquidar),
        '23)44(RP_NPro_Liquidado), 24)45(RP_NPro_Liq_APagar), 25)46(RP_NPro_Pago), 26)47(RP_NPro_APagar), 27)48(RP_NPro_Bloquedo),
        '28)Externo
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
                                    
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
               
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
   
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value <> 9 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "VL_Orc_Ini"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value <> 10 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "VL_Cred_Sup"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> 11 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Cred_Esp"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> 12 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Cred_Ext"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> 13 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Orc_Aut"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value <> 14 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Cred_Can"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> 19 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Disponivel"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> 20 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 23 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 25 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 28 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> "Externo" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "Externo"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "0"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
     
: Return
   
FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "VL_Orc_Ini"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "VL_Cred_Sup"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Cred_Esp"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Cred_Ext"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Orc_Aut"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Cred_Can"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Disponivel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Contido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Bloquedo"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "Externo"
        
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_EXECUCAO_PTRES(lm As Integer)
    Dim sAutor As String, sDescricaoPO As String
    
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
       'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                    
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano, 3)PTRES, 4)PO, 5)Descrição do PO, 6)RP, 7)Nº Emenda, 8)Autor,
        '9)9(VL_Orc_Ini), 10)10(VL_Cred_Sup), 11)11(VL_Cred_Ext), 12)12(VL_Cred_Esp), 13)13(VL_Orc_Aut),
        '14)14(VL_Cred_Can), 15)19(VL_Disponivel), 16)20(VL_Contido), 17)23(VL_Empenhado), 18)25(VL_Liquidado),
        '19)26(VL_Pago), 20)35(RP_Pro_Inscrito), 21)36(RP_Pro_Reinscrito), 22)37(RP_Pro_Cancelado), 23)38(RP_Pro_Pago),
        '24)39(RP_Pro_APagar), 25)40(RP_NPro_Inscrito), 26)41(RP_NPro_Reinscrito), 27)42(RP_NPro_Cancelado), 28)43(RP_NPro_ALiquidar),
        '29)44(RP_NPro_Liquidado), 30)45(RP_NPro_Liq_APagar), 31)46(RP_NPro_Pago), 32)47(RP_NPro_APagar), 33)48(RP_NPro_Bloquedo)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sDescricaoPO = Trim(Left(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 255))
            sAutor = Trim(Left(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 80))
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = sDescricaoPO
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sAutor
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> 9 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Orc_Ini"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> 10 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Cred_Sup"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 11 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Esp"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 12 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Ext"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 13 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Orc_Aut"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 14 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Can"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 19 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Disponivel"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 20 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 23 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 25 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 28 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "PO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Descrição do PO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "Nº Emenda"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Autor Emenda"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Orc_Ini"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Cred_Sup"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Esp"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Ext"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Orc_Aut"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Can"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Disponivel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Contido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = "RP_NPro_Bloquedo"
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_EXECUCAO_RP(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                   
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano, 3)RPrimario, 4)UG Executora,
        '5)19(VL_Disponivel), 6)20(VL_Contido), 7)23(VL_Empenhado), 8)25(VL_Liquidado), 9)26(VL_Pago),
        '10)35(RP_Pro_Inscrito), 11)36(RP_Pro_Reinscrito), 12)37(RP_Pro_Cancelado), 13)38(RP_Pro_Pago),
        '14)39(RP_Pro_APagar), 15)40(RP_NPro_Inscrito), 16)41(RP_NPro_Reinscrito), 17)42(RP_NPro_Cancelado),
        '18)43(RP_NPro_ALiquidar), 19)44(RP_NPro_Liquidado), 20)45(RP_NPro_Liq_APagar), 21)46(RP_NPro_Pago),
        '22)47(RP_NPro_APagar), 23)48(RP_NPro_Bloquedo)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> 19 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Disponivel"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> 20 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> 23 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value <> 25 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> 28 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    
: Return

FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "UG Executora"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Disponivel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Contido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Bloquedo"
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub


Sub TB_IMP_BLOQUEIO(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        
        'Estrutura do arquivo:
        '1)PT, 2)Ano Lançamento, 3)VL_Cred_Bloq_Rema(622120101), 4)VL_Cred_Bloq_Contido(622120104), 5)VL_Cred_Bloq_SOF(622120105),
        '6)VL_Cred_Bloq_RemaSOF(622120106), 7)VL_Cred_Bloq_RP(622120108), 8)VL_Cred_Bloq_PreEmp(622120200)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
   
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value <> "622120101" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "VL_Cred_Bloq_Rema"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value <> "622120104" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "VL_Cred_Bloq_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> "622120105" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Cred_Bloq_SOF"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> "622120106" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Cred_Bloq_RemaSOF"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> "622120107" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Cred_Bloq_Controle"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value <> "622120108" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Cred_Bloq_RP"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> "622120200" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Cred_Bloq_PreEmp"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    
: Return

FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "VL_Cred_Bloq_Rema"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "VL_Cred_Bloq_Contido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Cred_Bloq_SOF"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Cred_Bloq_RemaSOF"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Cred_Bloq_Controle"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Cred_Bloq_RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Cred_Bloq_PreEmp"
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_BLOQUEIO_NATUREZA(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        
        'Estrutura do arquivo:
        '1)PT, 2)Ano Lançamento, 3)PTRES, 4) RPrimario, 5) Natureza, 6) FonteSOF, 7) CatEco, 8) Grupo de Despesa(GND),
        '9) Modalidade de Despesa, 10) Elemento, 11) VL_Cred_Bloq_Rema, 12) VL_Cred_Bloq_Contido, 13) VL_Cred_Bloq_SOF
        '14) VL_Cred_Bloq_RemaSOF, 15) VL_Cred_Bloq_Controle, 16) VL_Cred_Bloq_RP, 17) VL_Cred_Bloq_PreEmp
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 5, 2)
            sFonteSOF = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value, 1, 4) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value, 5, 6)
                        
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = sFonteSOF
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
   'apenas formatando a coluna Fonte Sof
    Columns(6).NumberFormat = "@"
    Columns(6).HorizontalAlignment = xlHAlignLeft
    Columns(6).VerticalAlignment = xlVAlignTop
   
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> "622120101" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Bloq_Rema"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> "622120104" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Bloq_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> "622120105" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Bloq_SOF"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> "622120106" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Bloq_RemaSOF"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> "622120107" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Cred_Bloq_RP"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> "622120108" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Cred_Bloq_RP"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> "622120200" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Cred_Bloq_PreEmp"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    
: Return

FormatarCabecalho:
       
        '14) VL_Cred_Bloq_RemaSOF, 15) VL_Cred_Bloq_Controle, 16) VL_Cred_Bloq_RP, 17) VL_Cred_Bloq_PreEmp
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "FonteSOF"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CatEco"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "GND"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "Modalidade"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "Elemento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Bloq_Rema"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Bloq_Contido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Bloq_SOF"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Bloq_RemaSOF"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Cred_Bloq_Controle"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Cred_Bloq_RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Cred_Bloq_PreEmp"
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_DESTAQUE(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
                      
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        
        'Estrutura do arquivo:
        '1)PT, 2)Ano Lançamento, 3)PROVISAO RECEBIDA, 4)PROVISAO CONCEDIDA, 5)DESTAQUE RECEBIDO, 6)DESTAQUE CONCEDIDO
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
   
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value <> "PROVISAO RECEBIDA" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PROVISAO RECEBIDA"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value <> "PROVISAO CONCEDIDA" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "PROVISAO CONCEDIDA"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> "DESTAQUE RECEBIDO" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "DESTAQUE RECEBIDO"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> "DESTAQUE CONCEDIDO" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "DESTAQUE CONCEDIDO"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
        
: Return

FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PROVISAO RECEBIDA"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "PROVISAO CONCEDIDA"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "DESTAQUE RECEBIDO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "DESTAQUE CONCEDIDO"
   
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_DESTAQUE_PTRES(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
                      
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        
        'Estrutura do arquivo:
        '1)PT, 2)Ano Lançamento, 3)PTRES, 4)PROVISAO RECEBIDA, 5)PROVISAO CONCEDIDA, 6)DESTAQUE RECEBIDO, 7)DESTAQUE CONCEDIDO
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
   
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value <> "PROVISAO RECEBIDA" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "PROVISAO RECEBIDA"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> "PROVISAO CONCEDIDA" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "PROVISAO CONCEDIDA"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> "DESTAQUE RECEBIDO" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "DESTAQUE RECEBIDO"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> "DESTAQUE CONCEDIDO" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "DESTAQUE CONCEDIDO"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
        
: Return

FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "PROVISAO RECEBIDA"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "PROVISAO CONCEDIDA"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "DESTAQUE RECEBIDO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "DESTAQUE CONCEDIDO"
   
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub


Sub TB_IMP_INDISPONIVEL(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        
        'Estrutura do arquivo:
        '1)PT, 2)Ano Lançamento, 3) PTRES, 4) UO, 5) Modalidade, 6) GND, 7) Fonte, 8) RP
        '9)VL_Cred_Bloq_Rema(622120101), 10)VL_Cred_Bloq_Contido(622120104), 11)VL_Cred_Bloq_SOF(622120105),
        '12)VL_Cred_Bloq_RemaSOF(622120106), 13)VL_Cred_Bloq_Controle(622120107), 14)VL_Cred_Bloq_RP(622120108),
        '15)VL_Cred_Bloq_PreEmp(622120200)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sFonteSOF = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, 1, 4) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, 5, 6)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = CStr(sFonteSOF)
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
   'apenas formatando a coluna Fonte Sof
    Columns(7).NumberFormat = "@"
    Columns(7).HorizontalAlignment = xlHAlignLeft
    Columns(7).VerticalAlignment = xlVAlignTop
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> "622120101" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Cred_Bloq_Rema"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> "622120104" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Cred_Bloq_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> "622120105" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Bloq_SOF"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> "622120106" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Bloq_RemaSOF"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> "622120107" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Bloq_Controle"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> "622120108" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Bloq_RP"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> "622120200" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Cred_Bloq_PreEmp"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    
: Return

FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "UO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Modalidade"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "GND"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "Fonte"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Cred_Bloq_Rema"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Cred_Bloq_Contido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Bloq_SOF"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Bloq_RemaSOF"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Bloq_Controle"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Bloq_RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Cred_Bloq_PreEmp"
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_LOA_DETALHADA(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                   
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano, 3) PTRES, 4) UO, 5) Modalidade, 6) GND, 7)Fonte, 8) RP
        '9) 8 VL_PLOA, 10)9(VL_Orc_Ini), 11)10(VL_Cred_Sup), 12)11(VL_Cred_Esp), 13)12(VL_Cred_Ext), 14)13(VL_Orc_Aut), 15)14(VL_Cred_Can),
        '16)19(VL_Disponivel), 17)29(VL_Empenhado), 18)31(VL_Liquidado), 19)34(VL_Pago),
                
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sFonteSOF = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, 1, 4) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, 5, 6)
                                    
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = CStr(sFonteSOF)
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    'apenas formatando a coluna Fonte Sof
    Columns(7).NumberFormat = "@"
    Columns(7).HorizontalAlignment = xlHAlignLeft
    Columns(7).VerticalAlignment = xlVAlignTop
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> 8 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_PLOA"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignLeft
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> 9 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Orc_Ini"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignLeft
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 10 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Sup"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 11 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Esp"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 12 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Ext"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 13 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Orc_Aut"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 14 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Cred_Can"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 19 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Disponivel"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
     
: Return
   
FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "UO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Modalidade"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "GND"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "Fonte"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_PLOA"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Orc_Ini"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Sup"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Esp"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Ext"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Orc_Aut"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Cred_Can"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Disponivel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "VL_Pago"
        
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub



Sub TB_IMP_HISTORICO_PGTO(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
                      
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        
        'Estrutura do arquivo:
        '1)ContaCorrente, 2)Ano, 3)Mês, 4)DESPESAS PAGAS (CONTROLE EMPENHO), 5)RESTOS A PAGAR PAGOS (PROC E N PROC)
        
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        lSeq = 0             'Zerando a sequencia de registro no primeiro arquivo da série
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
            
            'Armazenar a informação da Nota de Empenho
            sContaCorrente = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value
            If sContaCorrente <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                sAno = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value, 5, 4)
                sMes = FormatarMes(Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value, 1, 3))
                
                'Armazenar as informações modificadas
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = sContaCorrente
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = sAno
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sMes
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
            End If
        
        'Fim do Loop 1.1
        Loop Until Len(sContaCorrente) = 0
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).NumberFormat = "@"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).HorizontalAlignment = xlHAlignLeft
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop

: Return
    
ConferirColunasValores:
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> "DESPESAS PAGAS (CONTROLE EMPENHO)" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> "RESTOS A PAGAR PAGOS (PROC E N PROC)" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "RP_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
            
: Return

FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequência"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Mês"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "RP_Pago"
       
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub


Sub TB_IMP_HISTORICO_PGTO2(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
                      
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        
        'Estrutura do arquivo:
        '1)ContaCorrente, 2)Ano, 3)Mês, 4)DESPESAS PAGAS (CONTROLE EMPENHO), 5)RESTOS A PAGAR PAGOS (PROC E N PROC)
        
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
            
            'Armazenar a informação da Nota de Empenho
            sContaCorrente = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value
            If sContaCorrente <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                sAno = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value, 5, 4)
                sMes = FormatarMes(Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value, 1, 3))
                
                'Armazenar as informações modificadas
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = sContaCorrente
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = sAno
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sMes
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
            End If
        
        'Fim do Loop 1.1
        Loop Until Len(sContaCorrente) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).NumberFormat = "@"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).HorizontalAlignment = xlHAlignLeft
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop

: Return
    
ConferirColunasValores:
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> "DESPESAS PAGAS (CONTROLE EMPENHO)" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> "RESTOS A PAGAR PAGOS (PROC E N PROC)" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "RP_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
            
: Return

FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequência"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Mês"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "RP_Pago"
       
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_SRE_ADM(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
                      
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        
        'Estrutura do arquivo:
        '1)PT, 2)Ano Lançamento, 3)PTRES, 4)CodMT, 5)UG_Executora, 6)Mês, 7)VL_Empenhado (DESPESAS EMPENHADAS), 8)VL_Liquidado (LIQUIDACOES TOTAIS (EXERCICIO E RPNP))
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value, 0)
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sCodMT
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> "DESPESAS EMPENHADAS" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value <> "LIQUIDACOES TOTAIS (EXERCICIO E RPNP)" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
            
: Return

FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "UG Executora"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "Mês/Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Liquidado"
   
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_PGTO(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
                      
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
         
        'Estrutura do arquivo:
        '1)PT, 2)Ano Lançamento, 3)Mês/Ano, 4)RPrimario, 5)Plano Orçamentario, 6)PAGAMENTOS TOTAIS (EXERCICIO E RAP)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> "PAGAMENTOS TOTAIS (EXERCICIO E RAP)" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
                
: Return

FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "Mês/Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "PO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Pago"
      
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_RAP_EMPENHO(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows("2:3").Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
       
        'Estrutura do arquivo:
        '1)ContaCorrente, 2)Ano, 3)Mês/Ano, 4)35(RP_ExAnt), 5)43(RP_Saldo), 6) 44(RP_Liquidado), 7) 48(RP_Bloquado),
        '8)50(RP_Inscrito), 9)51(RP_Cancelado), 10)52(RP_Pago)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sContaCorrente = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value
            If sContaCorrente <> "" Then
                'Armazenar a informação da Nota de Empenho da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sContaCorrente
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados
        Loop Until Len(sContaCorrente) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:

    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value <> "35" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "RP_ExAnt"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> "43" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "RP_Saldo"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> "44" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "RP_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> "48" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "RP_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value <> "50" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "RP_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> "51" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RP_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> "52" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "RP_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
        
: Return

FormatarCabecalho:
   
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "Mês/Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "RP_ExAnt"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "RP_Saldo"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "RP_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "RP_Bloqueado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "RP_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RP_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "RP_Pago"
    
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_LIMITE(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
                      
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        
        'Estrutura do arquivo:
        '1)PT, 2)Ano Lançamento, 3)RPrimario, 4)Valor Indisponível
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
       
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value <> "CREDITO INDISPONIVEL" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "VL_Indisponivel"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
            
: Return

FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "VL_Indisponivel"
       
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_RECEBIDO(lm As Integer)
    
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                   
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano,
        '3)9(VL_Orc_Ini), 4)10(VL_Cred_Sup), 5)11(VL_Cred_Ext), 6)12(VL_Cred_Esp), 7)13(VL_Orc_Aut), 8)14(VL_Cred_Can),
        '9)19(VL_Disponivel), 10)20(VL_Contido), 11)23(VL_Empenhado), 12)25(VL_Liquidado), 13)26(VL_Pago),
        '14)35(RP_Pro_Inscrito), 15)36(RP_Pro_Reinscrito), 16)37(RP_Pro_Cancelado), 17)38(RP_Pro_Pago),
        '18)39(RP_Pro_APagar), 19)40(RP_NPro_Inscrito), 20)41(RP_NPro_Reinscrito), 21)42(RP_NPro_Cancelado), 22)43(RP_NPro_ALiquidar),
        '23)44(RP_NPro_Liquidado), 24)45(RP_NPro_Liq_APagar), 25)46(RP_NPro_Pago), 26)47(RP_NPro_APagar), 27)48(RP_NPro_Bloquedo),
        '28)Externo
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
                                    
            If sFuncional <> "" And sFuncional <> "26.122.0032.2000.0001" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
               
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 1
            ElseIf sFuncional = "26.122.0032.2000.0001" Then
                'Deletar a linha da Funcional
                Rows(iLinha).Delete
                iLinha = iLinha - 1 'Voltar 1 no contador de linha
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
   
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value <> 9 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "VL_Orc_Ini"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).NumberFormat = "@"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).HorizontalAlignment = xlHAlignLeft
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value <> 10 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "VL_Cred_Sup"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> 11 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Cred_Esp"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> 12 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Cred_Ext"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> 13 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Orc_Aut"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value <> 14 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Cred_Can"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> 19 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Disponivel"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> 20 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 23 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 25 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 28 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> "Externo" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "Externo"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "0"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
     
: Return
   
FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "VL_Orc_Ini"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "VL_Cred_Sup"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Cred_Esp"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Cred_Ext"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Orc_Aut"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Cred_Can"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Disponivel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Contido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Bloquedo"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "Externo"
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_RECEBIDO_PTRES(lm As Integer)
    Dim sAutor As String, sDescricaoPO As String
    
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
       'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                    
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano, 3)PTRES, 4)PO, 5)Descrição do PO, 6)RP, 7)Nº Emenda, 8)Autor,
        '9)9(VL_Orc_Ini), 10)10(VL_Cred_Sup), 11)11(VL_Cred_Ext), 12)12(VL_Cred_Esp), 13)13(VL_Orc_Aut),
        '14)14(VL_Cred_Can), 15)19(VL_Disponivel), 16)20(VL_Contido), 17)23(VL_Empenhado), 18)25(VL_Liquidado),
        '19)26(VL_Pago), 20)35(RP_Pro_Inscrito), 21)36(RP_Pro_Reinscrito), 22)37(RP_Pro_Cancelado), 23)38(RP_Pro_Pago),
        '24)39(RP_Pro_APagar), 25)40(RP_NPro_Inscrito), 26)41(RP_NPro_Reinscrito), 27)42(RP_NPro_Cancelado), 28)43(RP_NPro_ALiquidar),
        '29)44(RP_NPro_Liquidado), 30)45(RP_NPro_Liq_APagar), 31)46(RP_NPro_Pago), 32)47(RP_NPro_APagar), 33)48(RP_NPro_Bloquedo)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sDescricaoPO = Trim(Left(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 255))
            sAutor = Trim(Left(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 80))
            
            If sFuncional <> "" And sFuncional <> "26.122.0032.2000.0001" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = sDescricaoPO
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sAutor
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = 0
            ElseIf sFuncional = "26.122.0032.2000.0001" Then
                'Deletar a linha da Funcional
                Rows(iLinha).Delete
                iLinha = iLinha - 1 'Voltar 1 no contador de linha
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> 9 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Orc_Ini"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "@"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignLeft
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> 10 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Cred_Sup"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 11 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Esp"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 12 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Ext"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 13 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Orc_Aut"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 14 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Can"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 19 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Disponivel"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 20 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 23 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 25 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 28 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "PO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Descrição do PO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "Nº Emenda"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Autor Emenda"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Orc_Ini"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Cred_Sup"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Esp"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Ext"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Orc_Aut"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Can"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Disponivel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Contido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = "RP_NPro_Bloquedo"
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_RECEBIDO_RP(lm As Integer)
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                   
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano, 3)RPrimario, 4)UG Executora,
        '5)19(VL_Disponivel), 6)20(VL_Contido), 7)23(VL_Empenhado), 8)25(VL_Liquidado), 9)26(VL_Pago),
        '10)35(RP_Pro_Inscrito), 11)36(RP_Pro_Reinscrito), 12)37(RP_Pro_Cancelado), 13)38(RP_Pro_Pago),
        '14)39(RP_Pro_APagar), 15)40(RP_NPro_Inscrito), 16)41(RP_NPro_Reinscrito), 17)42(RP_NPro_Cancelado),
        '18)43(RP_NPro_ALiquidar), 19)44(RP_NPro_Liquidado), 20)45(RP_NPro_Liq_APagar), 21)46(RP_NPro_Pago),
        '22)47(RP_NPro_APagar), 23)48(RP_NPro_Bloquedo)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            If sFuncional <> "" And sFuncional <> "26.122.0032.2000.0001" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
            ElseIf sFuncional = "26.122.0032.2000.0001" Then
                'Deletar a linha da Funcional
                Rows(iLinha).Delete
                iLinha = iLinha - 1 'Voltar 1 no contador de linha
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> 19 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Disponivel"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> 20 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> 23 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value <> 25 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> 28 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    
: Return

FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "UG Executora"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Disponivel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Contido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Bloquedo"
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_CREDITOS(lm As Integer)
        
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                   
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano, 3)PTRES, 4)GND, 5)FonteSOF, 6)RPrimario, 7)8(VL_PLOA), 8)9(VL_Orc_Ini), 9)10(VL_Cred_Sup),
        '10)11(VL_Cred_Esp), 11)12(VL_Cred_Ext), 12)13(VL_Orc_Aut), 13)14(VL_Cred_Can)
                
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            If iLinha = 168 Then
                iLinha = 168
            End If
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sAno = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value
            sFonteSOF = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 1, 4) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 5, 6)
            sRP = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value
                        
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = CStr(sFonteSOF)
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = sRP
               
               
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        GoSub LimparArquivo
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    'apenas formatando a coluna Fonte Sof
    Columns(5).NumberFormat = "@"
    Columns(5).HorizontalAlignment = xlHAlignLeft
    Columns(5).VerticalAlignment = xlVAlignTop
    'Formatando a coluna de crédito cancelado que recebe valor negativo
    Columns(7).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Columns(7).HorizontalAlignment = xlHAlignRight
    Columns(7).VerticalAlignment = xlVAlignTop
    Columns(8).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Columns(8).HorizontalAlignment = xlHAlignRight
    Columns(8).VerticalAlignment = xlVAlignTop
    Columns(9).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Columns(9).HorizontalAlignment = xlHAlignRight
    Columns(9).VerticalAlignment = xlVAlignTop
    Columns(10).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Columns(10).HorizontalAlignment = xlHAlignRight
    Columns(10).VerticalAlignment = xlVAlignTop
    Columns(11).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Columns(11).HorizontalAlignment = xlHAlignRight
    Columns(11).VerticalAlignment = xlVAlignTop
    Columns(12).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Columns(12).HorizontalAlignment = xlHAlignRight
    Columns(12).VerticalAlignment = xlVAlignTop
    Columns(13).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Columns(13).HorizontalAlignment = xlHAlignRight
    Columns(13).VerticalAlignment = xlVAlignTop
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> 8 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_PLOA"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value <> 9 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Orc_Ini"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> 10 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Cred_Sup"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> 11 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Cred_Esp"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 12 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Ext"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 13 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Orc_Aut"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 14 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Can"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
     
: Return
   
FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "GND"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "FonteSOF"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_PLOA"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "VL_Orc_Ini"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "VL_Cred_Sup"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "VL_Cred_Esp"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Cred_Ext"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Orc_Aut"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Can"
        
: Return

LimparArquivo:
    Range(Cells(iLinha, 1), Cells(Rows.Count, 1)).EntireRow.Select
    Selection.Delete Shift:=xlDown
    Range(Cells(1, 14), Cells(1, Columns.Count)).EntireColumn.Select
    Selection.Delete Shift:=xlToRight
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub



Sub TB_IMP_EMPREEND_ACAO(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
       
        'Estrutura do arquivo:
        '1)PT, 2)Ano Lançamento, 3)CodMT, 4)DOTAÇÃO ATUALIZADA, 5)DESPESA EMPENHADA, 6)DESPESA LIQUIDADA, 7)DESPESA PAGA
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value, 0)
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = sCodMT
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value <> "DOTAÇÃO ATUALIZADA" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "VL_Orc_Aut"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> "DESPESA EMPENHADA" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> "DESPESA LIQUIDADA" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> "DESPESA PAGA" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    
: Return

FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "VL_Orc_Aut"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "VL_Pago"
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_UG_DESTAQUE(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
       'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                    
        'Estrutura do arquivo:
        '1)PT, 2)Ano Lançamento, 3)PTRES, 4)Dotação Atualizada 5)Provisão Concedida, 6)Destaque Concedido, 7)Crédito Disponivel,
        '8)Crédito Indisponível, 9)Despesas Pre-Empenhadas a Empenhar, 10)Despesas Empenhadas (Controle e empenho), 11)Despesas Liquidadas (Controle e Empenho), 12)Despesas Pagas (Controle e Empenho),
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sAno = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = sAno
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
            
ConferirColunasValores:

    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value <> 13 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "DOTACAO ATUALIZADA"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> 16 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "PROVISAO CONCEDIDA"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> 18 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "DESTAQUE CONCEDIDO"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> 19 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CREDITO DISPONIVEL"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value <> 20 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "CREDITO INDISPONIVEL"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> 22 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "DESPESAS PRE-EMPENHADAS A EMPENHAR"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "DESPESAS EMPENHADAS (CONTROLE EMPENHO)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "DESPESAS LIQUIDADAS (CONTROLE EMPENHO)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "DESPESAS PAGAS (CONTROLE EMPENHO)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    
: Return

FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Dotação Atualizada"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Provisão Concedida"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "Destaque Concedido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "Crédito Disponível"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Crédito Indisponível"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "Despesas Pre-Empenhadas a Empenhar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "Despesas Empenhadas (Controle Empenho)"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Despesas Liquidadas (Controle Empenho)"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "Despesas Pagas (Controle Empenho)"
    
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    
    Exit Sub
End Sub


Sub TB_IMP_NATUREZADETALHADA(lm As Integer)
    Dim sNaturezaDetalhada As String, sPI As String
    
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
       'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                    
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano, 3)PTRES, 4)UG_Executora, 5) PO, 6) DescriçãoPO, 7)NaturezaDetalhada, 8)DetalhamentoNatureza,
        '9)PI, 10) Descrição do PI, 11)23(VL_Empenhado), 12)25(VL_Liquidado), 13)26(VL_Pago), 14)35(RP_Pro_Inscrito), 15)36(RP_Pro_Reinscrito),
        '16)37(RP_Pro_Cancelado), 17)38(RP_Pro_Pago), 18)40(RP_NPro_Inscrito), 19)41(RP_NPro_Reinscrito), 20)42(RP_NPro_Cancelado),
        '21)44(RP_NPro_Liquidado), 22)46(RP_NPro_Pago), 23)48(RP_NPro_Bloquedo)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sNaturezaDetalhada = FormatarNaturezaDetalhada(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value)
            sPI = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value, 0)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = sNaturezaDetalhada
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = sPI
                 
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoSub LimparArquivo
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "PO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Descrição PO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "UG_Executora"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Detalhamento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "PI"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "Descrição do PI"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Bloquedo"
: Return

LimparArquivo:
    Columns(10).EntireColumn.Select
    Selection.Delete Shift:=xlDown

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_UGEXECUTORA(lm As Integer)
        
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
       'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                    
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano, 3)PTRES, 4)UG_Executora, 5)GND, 6)15 Provisão Recebida,
        '7) 16 Provisão Concedida, 8) 18 Destaque Concedido, 9) 29 Credito Empenhado, 10) 31 Credito Liquidado,
        '11) 34 Valor Pago
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
                       
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                               
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> 15 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "Provisão Recebida"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> 16 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "Provisão Concedida"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value <> 18 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Destaque Concedido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "Valor Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "Valor Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 10), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Valor Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "UG_Executora"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "GND"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "Provisão Recebida"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "Provisão Concedida"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Destaque Concedido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "Valor Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "Valor Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Valor Pago"
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_DISPONIVELSRE(lm As Integer)
    Dim sDescricaoPO As String, sPI As String
    
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
       'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                    
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano, 3)PTRES, 4) UG_Executora, 5)RP, 6)GND, 7)PO, 8)Descrição do PO, 9)PI, 10) Descricao do PI
        '11)15(VL_Provisao_Recebida), 12)16(VL_Provisao_Concedida), 13)18(VL_Destacado), 14)19(VL_Disponivel),
        '15)20(VL_Contido)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sDescricaoPO = Trim(Left(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 255))
            sPI = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value)
                        
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sDescricaoPO
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = sPI
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        GoSub LimparArquivo
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 15 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Provisao_Recebida"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 16 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Provisao_Concedida"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 18 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Destacado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 19 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Disponivel"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 20 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
   
: Return
   
FormatarCabecalho:
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "UG Executora"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "GND"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "PO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Descrição do PO"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "PI"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "Descrição do PI"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Provisao_Recebida"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Provisao_Concedida"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Destacado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Disponivel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Contido"
   
: Return

LimparArquivo:
    Columns(10).EntireColumn.Select
    Selection.Delete Shift:=xlDown
    
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub


Sub TB_IMP_CORFIN_NE_DIARIO(lm As Integer)
    Dim sLinha As Integer
    
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'APAGANDO A QUARTA COLUNA
        Columns(4).Delete
                       
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                    
        'Estrutura do arquivo:
        '1)Ano, 2)DT_Emissao, 3)CodEspecie, 4)NomeEspecie, 5)VL_Empenhado
        
        GoSub FormatarCabecalho
        Columns(2).NumberFormat = "yyyy/mm/dd"
        Columns(2).HorizontalAlignment = xlHAlignLeft
        Columns(2).VerticalAlignment = xlVAlignTop
        Columns(3).NumberFormat = "@"
        Columns(3).HorizontalAlignment = xlHAlignLeft
        Columns(3).VerticalAlignment = xlVAlignTop
        Columns(4).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Columns(4).HorizontalAlignment = xlHAlignRight
        Columns(4).VerticalAlignment = xlVAlignTop
                
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
                                                            
            'se existe valor empenhado vamos guardar
            If (Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "-9") Or (Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "'-9") Then
               Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "1"
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value) = 0
                        
        GoSub LimparArquivo
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Emissão"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "Código"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "VL_Empenhado"
            
: Return

LimparArquivo:
    Range(Cells(iLinha, 1), Cells(Rows.Count, 1)).EntireRow.Select
    Selection.Delete Shift:=xlDown
    Range(Cells(1, 5), Cells(1, Columns.Count)).EntireColumn.Select
    Selection.Delete Shift:=xlToRight
    
: Return

SalvarArquivo:
    
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_EMPENHO_DETALHE_13(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                    
        'Estrutura do arquivo:
        '1)Sequencial, 2)ContaCorrente, 3)Funcional, 4)Ano, 5)PTRES, 6)CodMT, 7)RPrimario, 8)UG Responsável, 9)UG Executora, 10)Natureza
        '11)29(VL_EMPENHADO), 12)31(VL_Liquidado), 13)34(VL_Pago),
        '14)35(RP_Pro_Inscrito), 15)36(RP_Pro_Reinscrito), 16)37(RP_Pro_Cancelado), 17)38(RP_Pro_Pago),
        '18)39(RP_Pro_APagar), 19)40(RP_NPro_Inscrito), 20)41(RP_NPro_Reinscrito), 21)42(RP_NPro_Cancelado), 22)43(RP_NPro_ALiquidar),
        '23)44(RP_NPro_Liquidado), 24)45(RP_NPro_Liq_APagar), 25)46(RP_NPro_Pago), 26)47(RP_NPro_APagar), 27)48(RP_NPro_Bloquedo)
        
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
                                    
            'Armazenar a informação da rubrica da linha em questão
            sAno = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value, 12, 4)
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
            sRP = IIf(Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) > 0, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, 0)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value, sAno, CInt(sRP))
            sUGResp = FormatarUGResponsavel(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 5, 2)

            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sAno
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).NumberFormat = "@"
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = sCodMT
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sUGResp
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = sNatureza
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        'Deletando a coluna da UG Executora
        Columns(9).Delete
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).NumberFormat = "@"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).HorizontalAlignment = xlHAlignLeft
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 4), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop

    'apenas formatando a coluna 8
    Columns(8).NumberFormat = "@"
    Columns(8).HorizontalAlignment = xlHAlignLeft
    Columns(8).VerticalAlignment = xlVAlignTop

: Return
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "UG Responsavel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "UG Executora"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Bloquedo"
        
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_NATUREZA_REC(lm As Integer)
        
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                   
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano, 3)PTRES, 4)RPrimario, 5)Natureza, 6)FonteSOF 7)CatEco, 8)GND, 9)ModDesp, 10)Elemento,
        '11)9(VL_Orc_Ini), 12)10(VL_Cred_Sup), 13)11(VL_Cred_Ext), 14)12(VL_Cred_Esp), 15)13(VL_Orc_Aut), 16)14(VL_Cred_Can),
        '17)19(VL_Disponivel), 18)20(VL_Contido), 19)23(VL_Empenhado), 20)25(VL_Liquidado), 21)26(VL_Pago),
        '22)35(RP_Pro_Inscrito), 23)36(RP_Pro_Reinscrito), 24)37(RP_Pro_Cancelado), 25)38(RP_Pro_Pago), 26)39(RP_Pro_APagar),
        '27)40(RP_NPro_Inscrito), 28)41(RP_NPro_Reinscrito), 29)42(RP_NPro_Cancelado), 30)43(RP_NPro_ALiquidar),
        '31)44(RP_NPro_Liquidado), 32)45(RP_NPro_Liq_APagar), 33)46(RP_NPro_Pago), 34)47(RP_NPro_APagar), 35)48(RP_NPro_Bloquedo)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sAno = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value
            sRP = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 5, 2)
            sFonteSOF = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value, 1, 4) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value, 5, 6)
                        
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sRP
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = CStr(sFonteSOF)
               
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    'apenas formatando a coluna Fonte Sof
    Columns(6).NumberFormat = "@"
    Columns(6).HorizontalAlignment = xlHAlignLeft
    Columns(6).VerticalAlignment = xlVAlignTop
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 9 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Orc_Ini"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 10 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Sup"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 11 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Est"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 12 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Ext"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 13 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Orc_Aut"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 14 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Cred_Can"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 19 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Disponivel"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 20 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "VL_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 23 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 25 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 28 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 34), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 34), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 34), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 35), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 35), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 35), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
     
: Return
   
FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "FonteSOF"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CatEco"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "GND"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "ModDesp"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "Elemento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Orc_Ini"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Sup"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Esp"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Ext"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Orc_Aut"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Cred_Can"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Disponivel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "VL_Contido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Value = "RP_NPro_Bloquedo"
        
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_NATUREZA_PTRES(lm As Integer)
        
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                   
        'Estrutura do arquivo:
        '1)Funcional, 2)Ano, 3)PTRES, 4)RPrimario, 5)Natureza, 6)FonteSOF 7)CatEco, 8)GND, 9)ModDesp, 10)Elemento,
        '11)9(VL_Orc_Ini), 12)10(VL_Cred_Sup), 13)11(VL_Cred_Ext), 14)12(VL_Cred_Esp), 15)13(VL_Orc_Aut), 16)14(VL_Cred_Can),
        '17)19(VL_Disponivel), 18)20(VL_Contido), 19)23(VL_Empenhado), 20)25(VL_Liquidado), 21)26(VL_Pago),
        '22)35(RP_Pro_Inscrito), 23)36(RP_Pro_Reinscrito), 24)37(RP_Pro_Cancelado), 25)38(RP_Pro_Pago), 26)39(RP_Pro_APagar),
        '27)40(RP_NPro_Inscrito), 28)41(RP_NPro_Reinscrito), 29)42(RP_NPro_Cancelado), 30)43(RP_NPro_ALiquidar),
        '31)44(RP_NPro_Liquidado), 32)45(RP_NPro_Liq_APagar), 33)46(RP_NPro_Pago), 34)47(RP_NPro_APagar), 35)48(RP_NPro_Bloquedo)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value)
            sAno = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value
            sRP = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value, 5, 2)
            sFonteSOF = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value, 1, 4) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value, 5, 6)
                        
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sRP
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = CStr(sFonteSOF)
               
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    'apenas formatando a coluna Fonte Sof
    Columns(6).NumberFormat = "@"
    Columns(6).HorizontalAlignment = xlHAlignLeft
    Columns(6).VerticalAlignment = xlVAlignTop
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value <> 9 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Orc_Ini"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 11), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value <> 10 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Sup"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 12), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value <> 11 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Esp"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 13), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 12 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Ext"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 13 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Orc_Aut"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 14 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Cred_Can"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 19 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Disponivel"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 20 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "VL_Contido"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 23 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 25 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 28 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 31), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 32), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 33), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 34), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 34), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 34), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 35), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 35), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 35), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
     
: Return
   
FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "RP"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "FonteSOF"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CatEco"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "GND"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "ModDesp"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "Elemento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "VL_Orc_Ini"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "VL_Cred_Sup"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "VL_Cred_Esp"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Cred_Ext"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Orc_Aut"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Cred_Can"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "VL_Disponivel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "VL_Contido"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 31).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 32).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 33).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 34).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 35).Value = "RP_NPro_Bloquedo"
        
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_RECEITAS(lm As Integer)
    Dim sMesAno As String
    
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
                   
        'Estrutura do arquivo:
        '1)MesAno, 2)Fonte, 3)Natureza, 4)Especie, 5)Previsao(2), 6)Receita Bruta(3) 7)Deducao Receita(4), 8)Recita Liquida(5), 9)Receita Executada(89)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sMesAno = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value
            sFonteSOF = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value, 1, 4) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value, 5, 6)
            sNatureza = FormatarNaturezaDetalhada(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
           
                        
            If sMesAno <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = CStr(sFonteSOF)
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = CStr(sNatureza)
                               
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sMesAno) = 0
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
    
    'Formatando Numérico
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    'Formatando texto
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 2), Selection.End(xlDown)).NumberFormat = "@"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 2), Selection.End(xlDown)).HorizontalAlignment = xlHAlignLeft
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 2), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value <> 2 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Previsão Receita"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 5), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value <> 3 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "Receita Bruta"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 6), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value <> 4 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "Dedução Receita"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 7), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value <> 5 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Receita Liquida"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 8), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value <> 89 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "Receita Executada"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 9), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
     
: Return
   
FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "MêsAno"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "FonteSOF"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Espécie"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Previsão Receita"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "Receita Bruta"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "Dedução Receita"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Receita Liquida"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "Receita Executada"
    
: Return

SalvarArquivo:
    'Apagar colunas
    Columns("J:J").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft

    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_EMPENHO_SALDO(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows("2:3").Delete
                
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        
        'Estrutura do arquivo:
        '1)ContaCorrente, 2)RP_NPro_ALiquidar(631100000), 3)RP_NPro_Bloqueado(631510000)
        
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sContaCorrente = Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value
            If sContaCorrente <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = sContaCorrente
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sContaCorrente) = 0
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
ConferirColunasValores:
   
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value <> "43" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 2), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 2), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 2), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value <> "48" Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 3), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    
: Return

FormatarCabecalho:
       
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "RP_NPro_Bloqueado"
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
        
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_NE_2021_0(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        'lSeq = 0             'Zerando a sequencia de registro no primeiro arquivo da série
                    
        'Estrutura do arquivo:
        '1)Sequencial, 2)AnoLaçamento, 3)ContaCorrente, 4)Funcional, 5)Ano, 6)PTRES, 7)CodMT, 8)Natureza, 9)RPrimario, 10)UG Responsável, 11) NE Dia Emissão
        '12) cnpj, 13)Empresa 14)29(VL_EMPENHADO), 15)31(VL_Liquidado), 16)34(VL_Pago),
        '17)35(RP_Pro_Inscrito), 18)36(RP_Pro_Reinscrito), 19)37(RP_Pro_Cancelado), 20)38(RP_Pro_Pago),
        '21)39(RP_Pro_APagar), 22)40(RP_NPro_Inscrito), 23)41(RP_NPro_Reinscrito), 24)42(RP_NPro_Cancelado), 25)43(RP_NPro_ALiquidar),
        '26)44(RP_NPro_Liquidado), 27)45(RP_NPro_Liq_APagar), 28)46(RP_NPro_Pago), 29)47(RP_NPro_APagar), 30)48(RP_NPro_Bloquedo)
                
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
               
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
                                    
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value)
            sRP = IIf(Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) > 0, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, 0)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, "2018", CInt(sRP))
            sUGResp = FormatarUGResponsavel(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 5, 2)
            'sData = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 7, 4) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 4, 2) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 1, 2)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).NumberFormat = "@"
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = sCodMT
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = sUGResp
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).NumberFormat = "mm/dd/yyyy"
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
   
: Return
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
       
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "AnoLançamento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RPrimario"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "UG Responsavel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Data de Emissao"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "CNPJ"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "Empresa"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloquedo"
           
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub

End Sub

Sub TB_IMP_NE_2021_1(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        'lSeq = 0             'Zerando a sequencia de registro no primeiro arquivo da série
                    
        'Estrutura do arquivo:
        '1)Sequencial, 2)AnoLaçamento, 3)ContaCorrente, 4)Funcional, 5)Ano, 6)PTRES, 7)CodMT, 8)Natureza, 9)RPrimario, 10)UG Responsável, 11) NE Dia Emissão
        '12) cnpj, 13)Empresa 14)29(VL_EMPENHADO), 15)31(VL_Liquidado), 16)34(VL_Pago),
        '17)35(RP_Pro_Inscrito), 18)36(RP_Pro_Reinscrito), 19)37(RP_Pro_Cancelado), 20)38(RP_Pro_Pago),
        '21)39(RP_Pro_APagar), 22)40(RP_NPro_Inscrito), 23)41(RP_NPro_Reinscrito), 24)42(RP_NPro_Cancelado), 25)43(RP_NPro_ALiquidar),
        '26)44(RP_NPro_Liquidado), 27)45(RP_NPro_Liq_APagar), 28)46(RP_NPro_Pago), 29)47(RP_NPro_APagar), 30)48(RP_NPro_Bloquedo)
                
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
               
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
                                    
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value)
            sRP = IIf(Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) > 0, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, 0)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, "2018", CInt(sRP))
            sUGResp = FormatarUGResponsavel(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 5, 2)
            'sData = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 7, 4) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 4, 2) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 1, 2)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).NumberFormat = "@"
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = sCodMT
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = sUGResp
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).NumberFormat = "mm/dd/yyyy"
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
   
: Return
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
       
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "AnoLançamento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RPrimario"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "UG Responsavel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Data de Emissao"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "CNPJ"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "Empresa"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloquedo"
           
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub

End Sub


Sub TB_IMP_NE_2018_6(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        lSeq = 0             'Zerando a sequencia de registro no primeiro arquivo da série
                    
        'Estrutura do arquivo:
        '1)Sequencial, 2)AnoLaçamento, 3)ContaCorrente, 4)Funcional, 5)Ano, 6)PTRES, 7)CodMT, 8)Natureza, 9)RPrimario, 10)UG Responsável, 11) NE Dia Emissão
        '12) cnpj, 13)Empresa 14)29(VL_EMPENHADO), 15)31(VL_Liquidado), 16)34(VL_Pago),
        '17)35(RP_Pro_Inscrito), 18)36(RP_Pro_Reinscrito), 19)37(RP_Pro_Cancelado), 20)38(RP_Pro_Pago),
        '21)39(RP_Pro_APagar), 22)40(RP_NPro_Inscrito), 23)41(RP_NPro_Reinscrito), 24)42(RP_NPro_Cancelado), 25)43(RP_NPro_ALiquidar),
        '26)44(RP_NPro_Liquidado), 27)45(RP_NPro_Liq_APagar), 28)46(RP_NPro_Pago), 29)47(RP_NPro_APagar), 30)48(RP_NPro_Bloquedo)
                
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
               
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
                                    
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value)
            sRP = IIf(Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) > 0, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, 0)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, "2018", CInt(sRP))
            sUGResp = FormatarUGResponsavel(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 5, 2)
            'sData = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 7, 4) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 4, 2) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 1, 2)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).NumberFormat = "@"
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = sCodMT
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = sUGResp
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).NumberFormat = "mm/dd/yyyy"
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
   
: Return
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
       
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "AnoLançamento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RPrimario"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "UG Responsavel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Data de Emissao"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "CNPJ"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "Empresa"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloquedo"
           
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_NE_2023_0(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        'lSeq = 0             'Zerando a sequencia de registro no primeiro arquivo da série
                    
        'Estrutura do arquivo:
        '1)Sequencial, 2)AnoLaçamento, 3)ContaCorrente, 4)Funcional, 5)Ano, 6)PTRES, 7)CodMT, 8)Natureza, 9)RPrimario, 10)UG Responsável, 11) NE Dia Emissão
        '12) cnpj, 13)Empresa 14)29(VL_EMPENHADO), 15)31(VL_Liquidado), 16)34(VL_Pago),
        '17)35(RP_Pro_Inscrito), 18)36(RP_Pro_Reinscrito), 19)37(RP_Pro_Cancelado), 20)38(RP_Pro_Pago),
        '21)39(RP_Pro_APagar), 22)40(RP_NPro_Inscrito), 23)41(RP_NPro_Reinscrito), 24)42(RP_NPro_Cancelado), 25)43(RP_NPro_ALiquidar),
        '26)44(RP_NPro_Liquidado), 27)45(RP_NPro_Liq_APagar), 28)46(RP_NPro_Pago), 29)47(RP_NPro_APagar), 30)48(RP_NPro_Bloquedo)
                
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
               
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
                                    
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value)
            sRP = IIf(Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) > 0, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, 0)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, "2018", CInt(sRP))
            sUGResp = FormatarUGResponsavel(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 5, 2)
            'sData = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 7, 4) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 4, 2) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 1, 2)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).NumberFormat = "@"
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = sCodMT
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = sUGResp
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).NumberFormat = "mm/dd/yyyy"
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
   
: Return
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
       
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "AnoLançamento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RPrimario"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "UG Responsavel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Data de Emissao"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "CNPJ"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "Empresa"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloquedo"
           
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_NE_2019_4(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        'lSeq = 0             'Zerando a sequencia de registro no primeiro arquivo da série
                    
        'Estrutura do arquivo:
        '1)Sequencial, 2)AnoLaçamento, 3)ContaCorrente, 4)Funcional, 5)Ano, 6)PTRES, 7)CodMT, 8)Natureza, 9)RPrimario, 10)UG Responsável, 11) NE Dia Emissão
        '12) cnpj, 13)Empresa 14)29(VL_EMPENHADO), 15)31(VL_Liquidado), 16)34(VL_Pago),
        '17)35(RP_Pro_Inscrito), 18)36(RP_Pro_Reinscrito), 19)37(RP_Pro_Cancelado), 20)38(RP_Pro_Pago),
        '21)39(RP_Pro_APagar), 22)40(RP_NPro_Inscrito), 23)41(RP_NPro_Reinscrito), 24)42(RP_NPro_Cancelado), 25)43(RP_NPro_ALiquidar),
        '26)44(RP_NPro_Liquidado), 27)45(RP_NPro_Liq_APagar), 28)46(RP_NPro_Pago), 29)47(RP_NPro_APagar), 30)48(RP_NPro_Bloquedo)
                
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
               
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
                                    
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value)
            sRP = IIf(Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) > 0, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, 0)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, "2018", CInt(sRP))
            sUGResp = FormatarUGResponsavel(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 5, 2)
            'sData = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 7, 4) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 4, 2) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 1, 2)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).NumberFormat = "@"
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = sCodMT
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = sUGResp
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).NumberFormat = "mm/dd/yyyy"
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
   
: Return
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
       
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "AnoLançamento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RPrimario"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "UG Responsavel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Data de Emissao"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "CNPJ"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "Empresa"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloquedo"
           
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_NE_2023_1(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        'lSeq = 0             'Zerando a sequencia de registro no primeiro arquivo da série
                    
        'Estrutura do arquivo:
        '1)Sequencial, 2)AnoLaçamento, 3)ContaCorrente, 4)Funcional, 5)Ano, 6)PTRES, 7)CodMT, 8)Natureza, 9)RPrimario, 10)UG Responsável, 11) NE Dia Emissão
        '12) cnpj, 13)Empresa 14)29(VL_EMPENHADO), 15)31(VL_Liquidado), 16)34(VL_Pago),
        '17)35(RP_Pro_Inscrito), 18)36(RP_Pro_Reinscrito), 19)37(RP_Pro_Cancelado), 20)38(RP_Pro_Pago),
        '21)39(RP_Pro_APagar), 22)40(RP_NPro_Inscrito), 23)41(RP_NPro_Reinscrito), 24)42(RP_NPro_Cancelado), 25)43(RP_NPro_ALiquidar),
        '26)44(RP_NPro_Liquidado), 27)45(RP_NPro_Liq_APagar), 28)46(RP_NPro_Pago), 29)47(RP_NPro_APagar), 30)48(RP_NPro_Bloquedo)
                
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
               
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
                                    
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value)
            sRP = IIf(Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) > 0, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, 0)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, "2018", CInt(sRP))
            sUGResp = FormatarUGResponsavel(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 5, 2)
            'sData = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 7, 4) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 4, 2) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 1, 2)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).NumberFormat = "@"
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = sCodMT
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = sUGResp
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).NumberFormat = "mm/dd/yyyy"
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
   
: Return
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
       
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "AnoLançamento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RPrimario"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "UG Responsavel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Data de Emissao"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "CNPJ"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "Empresa"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloquedo"
           
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub


Sub TB_IMP_NE_2020_2(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        'lSeq = 0             'Zerando a sequencia de registro no primeiro arquivo da série
                    
        'Estrutura do arquivo:
        '1)Sequencial, 2)AnoLaçamento, 3)ContaCorrente, 4)Funcional, 5)Ano, 6)PTRES, 7)CodMT, 8)Natureza, 9)RPrimario, 10)UG Responsável, 11) NE Dia Emissão
        '12) cnpj, 13)Empresa 14)29(VL_EMPENHADO), 15)31(VL_Liquidado), 16)34(VL_Pago),
        '17)35(RP_Pro_Inscrito), 18)36(RP_Pro_Reinscrito), 19)37(RP_Pro_Cancelado), 20)38(RP_Pro_Pago),
        '21)39(RP_Pro_APagar), 22)40(RP_NPro_Inscrito), 23)41(RP_NPro_Reinscrito), 24)42(RP_NPro_Cancelado), 25)43(RP_NPro_ALiquidar),
        '26)44(RP_NPro_Liquidado), 27)45(RP_NPro_Liq_APagar), 28)46(RP_NPro_Pago), 29)47(RP_NPro_APagar), 30)48(RP_NPro_Bloquedo)
                
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
               
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
                                    
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value)
            sRP = IIf(Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) > 0, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, 0)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, "2018", CInt(sRP))
            sUGResp = FormatarUGResponsavel(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 5, 2)
            'sData = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 7, 4) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 4, 2) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 1, 2)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).NumberFormat = "@"
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = sCodMT
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = sUGResp
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).NumberFormat = "mm/dd/yyyy"
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
   
: Return
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
       
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "AnoLançamento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RPrimario"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "UG Responsavel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Data de Emissao"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "CNPJ"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "Empresa"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloquedo"
           
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_NE_2020_3(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        'lSeq = 0             'Zerando a sequencia de registro no primeiro arquivo da série
                    
        'Estrutura do arquivo:
        '1)Sequencial, 2)AnoLaçamento, 3)ContaCorrente, 4)Funcional, 5)Ano, 6)PTRES, 7)CodMT, 8)Natureza, 9)RPrimario, 10)UG Responsável, 11) NE Dia Emissão
        '12) cnpj, 13)Empresa 14)29(VL_EMPENHADO), 15)31(VL_Liquidado), 16)34(VL_Pago),
        '17)35(RP_Pro_Inscrito), 18)36(RP_Pro_Reinscrito), 19)37(RP_Pro_Cancelado), 20)38(RP_Pro_Pago),
        '21)39(RP_Pro_APagar), 22)40(RP_NPro_Inscrito), 23)41(RP_NPro_Reinscrito), 24)42(RP_NPro_Cancelado), 25)43(RP_NPro_ALiquidar),
        '26)44(RP_NPro_Liquidado), 27)45(RP_NPro_Liq_APagar), 28)46(RP_NPro_Pago), 29)47(RP_NPro_APagar), 30)48(RP_NPro_Bloquedo)
                
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
               
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
                                    
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value)
            sRP = IIf(Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) > 0, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, 0)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, "2018", CInt(sRP))
            sUGResp = FormatarUGResponsavel(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 5, 2)
            'sData = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 7, 4) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 4, 2) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 1, 2)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).NumberFormat = "@"
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = sCodMT
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = sUGResp
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).NumberFormat = "mm/dd/yyyy"
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
   
: Return
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
       
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "AnoLançamento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RPrimario"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "UG Responsavel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Data de Emissao"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "CNPJ"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "Empresa"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloquedo"
           
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub TB_IMP_NE_2022_0(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        lSeq = 0             'Zerando a sequencia de registro no primeiro arquivo da série
                    
        'Estrutura do arquivo:
        '1)Sequencial, 2)AnoLaçamento, 3)ContaCorrente, 4)Funcional, 5)Ano, 6)PTRES, 7)CodMT, 8)Natureza, 9)RPrimario, 10)UG Responsável, 11) NE Dia Emissão
        '12) cnpj, 13)Empresa 14)29(VL_EMPENHADO), 15)31(VL_Liquidado), 16)34(VL_Pago),
        '17)35(RP_Pro_Inscrito), 18)36(RP_Pro_Reinscrito), 19)37(RP_Pro_Cancelado), 20)38(RP_Pro_Pago),
        '21)39(RP_Pro_APagar), 22)40(RP_NPro_Inscrito), 23)41(RP_NPro_Reinscrito), 24)42(RP_NPro_Cancelado), 25)43(RP_NPro_ALiquidar),
        '26)44(RP_NPro_Liquidado), 27)45(RP_NPro_Liq_APagar), 28)46(RP_NPro_Pago), 29)47(RP_NPro_APagar), 30)48(RP_NPro_Bloquedo)
                
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
               
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
                                    
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value)
            sRP = IIf(Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) > 0, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, 0)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, "2018", CInt(sRP))
            sUGResp = FormatarUGResponsavel(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 5, 2)
            'sData = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 7, 4) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 4, 2) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 1, 2)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).NumberFormat = "@"
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = sCodMT
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = sUGResp
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).NumberFormat = "mm/dd/yyyy"
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
   
: Return
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
       
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "AnoLançamento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RPrimario"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "UG Responsavel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Data de Emissao"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "CNPJ"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "Empresa"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloquedo"
           
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub
Sub TB_IMP_NE_2022_1(lm As Integer)
       
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre o arquivo que será processado
    sArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 1)
    'Primeira linha do arquivo contendo o cabeçalho
    iLinhaCabecalho = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 3)
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivo
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivo, UpdateLinks:=False
        
        'Desmesclar qualquer célula
        Rows.Select
        Selection.UnMerge
                
        'Apagar as linhas sobre a legenda do WBArquivo
        Rows("1:" & iLinhaCabecalho - 1).Select
        Selection.Delete Shift:=xlUp
        
        'APAGANDO A SEGUNDA LINHA DE CABEÇALHO: SOMENTE ARQUIVOS COM MAIS DE UMA LINHA NO CABEÇALHO
        Rows(2).Delete
        
        'Inicializar o contador de linhas do WBArquivo
        iLinha = 1
        'lSeq = 0             'Zerando a sequencia de registro no primeiro arquivo da série
                    
        'Estrutura do arquivo:
        '1)Sequencial, 2)AnoLaçamento, 3)ContaCorrente, 4)Funcional, 5)Ano, 6)PTRES, 7)CodMT, 8)Natureza, 9)RPrimario, 10)UG Responsável, 11) NE Dia Emissão
        '12) cnpj, 13)Empresa 14)29(VL_EMPENHADO), 15)31(VL_Liquidado), 16)34(VL_Pago),
        '17)35(RP_Pro_Inscrito), 18)36(RP_Pro_Reinscrito), 19)37(RP_Pro_Cancelado), 20)38(RP_Pro_Pago),
        '21)39(RP_Pro_APagar), 22)40(RP_NPro_Inscrito), 23)41(RP_NPro_Reinscrito), 24)42(RP_NPro_Cancelado), 25)43(RP_NPro_ALiquidar),
        '26)44(RP_NPro_Liquidado), 27)45(RP_NPro_Liq_APagar), 28)46(RP_NPro_Pago), 29)47(RP_NPro_APagar), 30)48(RP_NPro_Bloquedo)
                
        GoSub CriandoColunas
        GoSub ConferirColunasValores
        GoSub FormatarCabecalho
               
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Incrementar o contador de registro do arquivo
            lSeq = lSeq + 1
                                    
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = FormatarFuncional(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value)
            sRP = IIf(Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value) > 0, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value, 0)
            sCodMT = FormatarCodMT(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value, "2018", CInt(sRP))
            sUGResp = FormatarUGResponsavel(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value)
            sNatureza = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 1, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 2, 1) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 3, 2) + "." + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value, 5, 2)
            'sData = Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 7, 4) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value, 4, 2) + "-" + Mid(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value, 1, 2)
            
            If sFuncional <> "" Then
                'Armazenar a informação da rubrica da linha em questão
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = lSeq
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = sFuncional
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).NumberFormat = "@"
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = sCodMT
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = sNatureza
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = sUGResp
                Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).NumberFormat = "mm/dd/yyyy"
                
                'atribuindo valor zero onde não existe nada
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = 0
                If Len(Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value) = 0 Then Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = 0
            End If
        
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
        
        'Voltar o contado de linha interno
        lSeq = lSeq - 1
        
        GoTo SalvarArquivo
    Else
        Exit Sub
    End If
    
CriandoColunas:
    'Inserir coluna empurrando para direita
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).EntireColumn.Insert
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Select
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).NumberFormat = "0"
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
    Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 1), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    
   
: Return
    
ConferirColunasValores:
    
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value <> 29 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 14), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value <> 31 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 15), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value <> 34 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 16), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value <> 35 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 17), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value <> 36 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 18), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value <> 37 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 19), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value <> 38 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 20), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value <> 39 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 21), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value <> 40 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 22), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value <> 41 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 23), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value <> 42 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 24), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value <> 43 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 25), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value <> 44 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 26), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value <> 45 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 27), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value <> 46 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 28), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value <> 47 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 29), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
    If Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value <> 48 Then
        'Inserir coluna empurrando para direita
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).EntireColumn.Insert
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Select
        Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloqueado"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).HorizontalAlignment = xlHAlignRight
        Workbooks(sArquivo).Worksheets(1).Range(Cells(iLinha + 1, 30), Selection.End(xlDown)).VerticalAlignment = xlVAlignTop
    End If
       
: Return
   
FormatarCabecalho:
        
    'Itens a mudar no cabeçalho
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 1).Value = "Sequencia"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 2).Value = "AnoLançamento"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 3).Value = "ContaCorrente"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 4).Value = "Funcional"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 5).Value = "Ano"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 6).Value = "PTRES"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 7).Value = "CodMT"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 8).Value = "Natureza"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 9).Value = "RPrimario"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 10).Value = "UG Responsavel"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 11).Value = "Data de Emissao"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 12).Value = "CNPJ"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 13).Value = "Empresa"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 14).Value = "VL_Empenhado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 15).Value = "VL_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 16).Value = "VL_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 17).Value = "RP_Pro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 18).Value = "RP_Pro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 19).Value = "RP_Pro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 20).Value = "RP_Pro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 21).Value = "RP_Pro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 22).Value = "RP_NPro_Inscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 23).Value = "RP_NPro_Reinscrito"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 24).Value = "RP_NPro_Cancelado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 25).Value = "RP_NPro_ALiquidar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 26).Value = "RP_NPro_Liquidado"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 27).Value = "RP_NPro_Liq_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 28).Value = "RP_NPro_Pago"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 29).Value = "RP_NPro_APagar"
    Workbooks(sArquivo).Worksheets(1).Cells(iLinha, 30).Value = "RP_NPro_Bloquedo"
           
: Return

SalvarArquivo:
    'Renomear a Planilha Ativa com o mesmo nome do Arquivo (sem o .xlsx - 5 caracteres finais)
    sPlanWBArquivo = Workbooks(sWBMacro).Worksheets(sPlanAtiva).Range("B1").Value & "_" & sArquivo
    Workbooks(sArquivo).Activate
    ActiveSheet.Name = Left(sPlanWBArquivo, Len(sPlanWBArquivo) - 5)
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivo & "..."
    Workbooks(sArquivo).SaveAs Filename:=sCaminho & "\" & sPlanWBArquivo, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks(sPlanWBArquivo).Close SaveChanges:=False
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo " + sArquivo + " gerou erro: " + Error + " na linha : " + CStr(lSeq), vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub


'Função que formata a funcional programática
Function FormatarFuncional(PT As String) As String
    If (Len(PT) > 0) Then
        FormatarFuncional = Left(PT, 2) & "." & Right(Left(PT, 5), 3) & "." & Right(Left(PT, 9), 4) & "." & Right(Left(PT, 13), 4) & "." & Right(Left(PT, 17), 4)
    Else
        FormatarFuncional = ""
    End If
    
End Function

'Função que formata o Código MT
Function FormatarCodMT(MT As String, Ano As String, RP As Integer) As String
    MT = Trim(MT)      'Retirando espaços indesejados
    FormatarCodMT = MT 'Garantindo algum valor para o MT
    
    'Caso o empenho seja a partir de 2020
    If (Ano >= "2020") Then
        If ((MT = "'-8" Or MT = "-8") And (RP = 0 Or RP = 1 Or RP = 2 Or RP = 6 Or RP = 7 Or RP = 8 Or RP = 9)) Then
            FormatarCodMT = "MT99999"
        ElseIf (Len(MT) > 3) And (Len(MT) <= 8) Then
            FormatarCodMT = CStr(MT)
        ElseIf (Len(MT) = 3) And (MT = "DAQ") Then
            FormatarCodMT = "DAQ0001"
        Else
            FormatarCodMT = Left(MT, 8)
        End If
    Else  'Caso seja menor que 2020
        If ((MT = "'-8") Or (Left(MT, 1) = "B")) And (RP = 0 Or RP = 1 Or RP = 2 Or RP = 6 Or RP = 7 Or RP = 8 Or RP = 9) Then
            FormatarCodMT = "MT99999"
        ElseIf (MT = "'-8") And (RP = 3) Then
            FormatarCodMT = "MT00000"
        ElseIf (MT = "MT.01161") And (RP = 3) Then
            FormatarCodMT = "MT01161"
        ElseIf (Len(MT) = 7) Then
            FormatarCodMT = CStr(MT)
        ElseIf (Len(MT) = 8 And MT <> "MT.01161") Then
            FormatarCodMT = CStr(MT)
        ElseIf (Len(MT) = 11) And (RP = 2) Then
            FormatarCodMT = "MT99999"
        ElseIf (Len(MT) = 11) And (RP = 7) Then
            FormatarCodMT = Left(MT, 8)
        Else
            FormatarCodMT = "        "
        End If
    End If
    
End Function

'Função que formata Fonte SOF
Function FormatarFonteSOF(FT As String, DT As String) As String
    
    FormatarFonteSOF = FT + "." + DT
    
End Function

'Função que formata a UG Responsável
Function FormatarUGResponsavel(ugResp As String, ugExec As String) As String
    
    If Len(ugResp) = 6 Then
       FormatarUGResponsavel = CStr(ugResp)
    Else
       FormatarUGResponsavel = CStr(Mid(ugExec, 1, 6))
    End If
End Function

'Função que formata Mês em nímero
Function FormatarMes(MS As String) As Integer
    Select Case MS
            Case "JAN"
                FormatarMes = 1
            Case "FEV"
                FormatarMes = 2
            Case "MAR"
                FormatarMes = 3
            Case "ABR"
                FormatarMes = 4
            Case "MAI"
                FormatarMes = 5
            Case "JUN"
                FormatarMes = 6
            Case "JUL"
                FormatarMes = 7
            Case "AGO"
                FormatarMes = 8
            Case "SET"
                FormatarMes = 9
            Case "OUT"
                FormatarMes = 10
            Case "NOV"
                FormatarMes = 11
            Case "DEZ"
                FormatarMes = 12
            Case Else
                FormatarMes = 0
    End Select
End Function

'Função que formata a NaturezaDetalhada
Function FormatarNaturezaDetalhada(Natureza As String) As String
    Dim valor As Integer
  
    If Len(Natureza) < 8 Then 'se não existe natureza definida
       FormatarNaturezaDetalhada = "0.0.00.00.00"
    Else
       valor = Right(Natureza, 2)
       If valor = -9 Then
          FormatarNaturezaDetalhada = Mid(Natureza, 1, 1) + "." + Mid(Natureza, 2, 1) + "." + Mid(Natureza, 3, 2) + "." + Mid(Natureza, 5, 2) + ".00"
       Else
          FormatarNaturezaDetalhada = Mid(Natureza, 1, 1) + "." + Mid(Natureza, 2, 1) + "." + Mid(Natureza, 3, 2) + "." + Mid(Natureza, 5, 2) + "." + Mid(Natureza, 7, 2)
       End If
    End If
End Function

Sub JUNTARVALORESEXECUCAO(lm As Integer)
    Dim Funcional(1) As String, sFuncionalOrigem As String, i As Integer, Valores(1 To 25) As Double, ln As Long, _
        sArquivoOrigem As String, sArquivoDestino As String, x As Integer
    
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre os arquivos que serão processados
    sArquivoDestino = "EDIT_TB_IMP_EXECUCAO.xlsx"
    sArquivoOrigem = "EDIT_TB_IMP_RECEBIDO.xlsx"
    Funcional(0) = "26.784.2086.212A.0030"
    
        
    'zerando o vetor que vai receber os valores
    For x = 1 To 25
        Valores(x) = 0
        x = x + 1
    Next
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivoDestino
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivoDestino, UpdateLinks:=False
        
        'Inicializar o contador de linhas do Arquivo destino e origem
        iLinha = 1
        i = 0
        ln = 1
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 1).Value
                                    
            If (sFuncional = Funcional(i)) Then
                'Recebendo os valores da linha selecionada
                Valores(1) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 3).Value)
                Valores(2) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 4).Value)
                Valores(3) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 5).Value)
                Valores(4) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 6).Value)
                Valores(5) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 7).Value)
                Valores(6) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 8).Value)
                Valores(7) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 9).Value)
                Valores(8) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 10).Value)
                Valores(9) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 11).Value)
                Valores(10) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 12).Value)
                Valores(11) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 13).Value)
                Valores(12) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 14).Value)
                Valores(13) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 15).Value)
                Valores(14) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 16).Value)
                Valores(15) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 17).Value)
                Valores(16) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 18).Value)
                Valores(17) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 19).Value)
                Valores(18) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 20).Value)
                Valores(19) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 21).Value)
                Valores(20) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 22).Value)
                Valores(21) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 23).Value)
                Valores(22) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 24).Value)
                Valores(23) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 25).Value)
                Valores(24) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 26).Value)
                Valores(25) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 27).Value)
                
                'Abrir o Arquivo Origem para importação dos valores
                Application.StatusBar = "Abrindo " & sArquivoOrigem
                DoEvents  'Passa operação para o Sistema operacional
                Workbooks.Open Filename:=sCaminho & "\" & sArquivoOrigem, UpdateLinks:=False
                
                Do
                    'Incrementar o contador de linhas do WBArquivo
                    ln = ln + 1
            
                    'Armazenar a informação da rubrica da linha em questão
                    sFuncionalOrigem = Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 1).Value
                    If (sFuncionalOrigem = Funcional(i)) Then
                        'Recebendo os valores da linha selecionada
                        Valores(1) = Valores(1) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 3).Value)
                        Valores(2) = Valores(2) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 4).Value)
                        Valores(3) = Valores(3) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 5).Value)
                        Valores(4) = Valores(4) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 6).Value)
                        Valores(5) = Valores(5) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 7).Value)
                        Valores(6) = Valores(6) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 8).Value)
                        Valores(7) = Valores(7) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 9).Value)
                        Valores(8) = Valores(8) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 10).Value)
                        Valores(9) = Valores(9) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 11).Value)
                        Valores(10) = Valores(10) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 12).Value)
                        Valores(11) = Valores(11) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 13).Value)
                        Valores(12) = Valores(12) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 14).Value)
                        Valores(13) = Valores(13) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 15).Value)
                        Valores(14) = Valores(14) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 16).Value)
                        Valores(15) = Valores(15) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 17).Value)
                        Valores(16) = Valores(16) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 18).Value)
                        Valores(17) = Valores(17) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 19).Value)
                        Valores(18) = Valores(18) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 20).Value)
                        Valores(19) = Valores(19) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 21).Value)
                        Valores(20) = Valores(20) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 22).Value)
                        Valores(21) = Valores(21) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 23).Value)
                        Valores(22) = Valores(22) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 24).Value)
                        Valores(23) = Valores(23) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 25).Value)
                        Valores(24) = Valores(24) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 26).Value)
                        Valores(25) = Valores(25) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 27).Value)
                        
                        'Selecionando e apagando a linha copiada
                        Rows(ln).Select
                        Rows(ln).Delete
                        
                        'Fechar o WBarquivo, salvando como
                        DoEvents
                        Application.StatusBar = "Fechando o arquivo " & sArquivoOrigem & "..."
                        Workbooks(sArquivoOrigem).Close SaveChanges:=True
                        Application.StatusBar = "Voltando para o arquivo " & sArquivoDestino & "..."
                        
                        'Atualizar Arquivo Destino
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 3).Value = Valores(1)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 4).Value = Valores(2)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 5).Value = Valores(3)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 6).Value = Valores(4)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 7).Value = Valores(5)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 8).Value = Valores(6)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 9).Value = Valores(7)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 10).Value = Valores(8)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 11).Value = Valores(9)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 12).Value = Valores(10)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 13).Value = Valores(11)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 14).Value = Valores(12)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 15).Value = Valores(13)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 16).Value = Valores(14)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 17).Value = Valores(15)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 18).Value = Valores(16)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 19).Value = Valores(17)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 20).Value = Valores(18)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 21).Value = Valores(19)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 22).Value = Valores(20)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 23).Value = Valores(21)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 24).Value = Valores(22)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 25).Value = Valores(23)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 26).Value = Valores(24)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 27).Value = Valores(25)
                        Workbooks(sArquivoDestino).Save
        
                        iLinha = 1
                        ln = 1
                        'zerando o vetor que vai receber os valores
                        For x = 1 To 25
                            Valores(x) = 0
                        x = x + 1
                        Next
                        i = i + 1
                        
                        If i = 1 Then
                            GoTo SalvarArquivo
                        End If
                        Exit Do
                    End If
                Loop Until Len(sFuncionalOrigem) = 0
            End If
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
    Else
        Exit Sub
    End If
   

SalvarArquivo:
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivoDestino & "..."
    Workbooks(sArquivoDestino).Close SaveChanges:=True
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo JUNTARVALORESEXECUCAO() gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub JUNTARVALORESEXECUCAOPTRES(lm As Integer)
    Dim PTRES(1) As String, sPTRESOrigem As String, i As Integer, Valores(1 To 25) As Double, ln As Long, _
        sArquivoOrigem As String, sArquivoDestino As String, x As Integer, sPTRES As String
    
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre os arquivos que serão processados
    sArquivoDestino = "EDIT_TB_IMP_EXECUCAO_PTRES.xlsx"
    sArquivoOrigem = "EDIT_TB_IMP_RECEBIDO_PTRES.xlsx"
    PTRES(0) = "194825"
            
    'zerando o vetor que vai receber os valores
    For x = 1 To 25
        Valores(x) = 0
        x = x + 1
    Next
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivoDestino
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivoDestino, UpdateLinks:=False
        
        'Inicializar o contador de linhas do Arquivo destino e origem
        iLinha = 1
        i = 0
        ln = 1
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sPTRES = Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 3).Value
                                    
            If (sPTRES = PTRES(i)) Then
                'Recebendo os valores da linha selecionada
                Valores(1) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 9).Value)
                Valores(2) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 10).Value)
                Valores(3) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 11).Value)
                Valores(4) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 12).Value)
                Valores(5) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 13).Value)
                Valores(6) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 14).Value)
                Valores(7) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 15).Value)
                Valores(8) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 16).Value)
                Valores(9) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 17).Value)
                Valores(10) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 18).Value)
                Valores(11) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 19).Value)
                Valores(12) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 20).Value)
                Valores(13) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 21).Value)
                Valores(14) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 22).Value)
                Valores(15) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 23).Value)
                Valores(16) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 24).Value)
                Valores(17) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 25).Value)
                Valores(18) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 26).Value)
                Valores(19) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 27).Value)
                Valores(20) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 28).Value)
                Valores(21) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 29).Value)
                Valores(22) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 30).Value)
                Valores(23) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 31).Value)
                Valores(24) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 32).Value)
                Valores(25) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 33).Value)
                
                'Abrir o Arquivo Origem para importação dos valores
                Application.StatusBar = "Abrindo " & sArquivoOrigem
                DoEvents  'Passa operação para o Sistema operacional
                Workbooks.Open Filename:=sCaminho & "\" & sArquivoOrigem, UpdateLinks:=False
                
                Do
                    'Incrementar o contador de linhas do WBArquivo
                    ln = ln + 1
            
                    'Armazenar a informação da rubrica da linha em questão
                    sPTRESOrigem = Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 3).Value
                    If (sPTRESOrigem = PTRES(i)) Then
                        'Recebendo os valores da linha selecionada
                        Valores(1) = Valores(1) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 9).Value)
                        Valores(2) = Valores(2) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 10).Value)
                        Valores(3) = Valores(3) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 11).Value)
                        Valores(4) = Valores(4) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 12).Value)
                        Valores(5) = Valores(5) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 13).Value)
                        Valores(6) = Valores(6) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 14).Value)
                        Valores(7) = Valores(7) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 15).Value)
                        Valores(8) = Valores(8) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 16).Value)
                        Valores(9) = Valores(9) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 17).Value)
                        Valores(10) = Valores(10) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 18).Value)
                        Valores(11) = Valores(11) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 19).Value)
                        Valores(12) = Valores(12) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 20).Value)
                        Valores(13) = Valores(13) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 21).Value)
                        Valores(14) = Valores(14) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 22).Value)
                        Valores(15) = Valores(15) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 23).Value)
                        Valores(16) = Valores(16) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 24).Value)
                        Valores(17) = Valores(17) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 25).Value)
                        Valores(18) = Valores(18) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 26).Value)
                        Valores(19) = Valores(19) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 27).Value)
                        Valores(20) = Valores(20) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 28).Value)
                        Valores(21) = Valores(21) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 29).Value)
                        Valores(22) = Valores(22) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 30).Value)
                        Valores(23) = Valores(23) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 31).Value)
                        Valores(24) = Valores(24) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 32).Value)
                        Valores(25) = Valores(25) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 33).Value)
                        
                        'Selecionando e apagando a linha copiada
                        Rows(ln).Select
                        Rows(ln).Delete
                        
                        'Fechar o WBarquivo, salvando como
                        DoEvents
                        Application.StatusBar = "Fechando o arquivo " & sArquivoOrigem & "..."
                        Workbooks(sArquivoOrigem).Close SaveChanges:=True
                        Application.StatusBar = "Voltando para o arquivo " & sArquivoDestino & "..."
                        
                        'Atualizar Arquivo Destino
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 9).Value = Valores(1)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 10).Value = Valores(2)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 11).Value = Valores(3)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 12).Value = Valores(4)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 13).Value = Valores(5)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 14).Value = Valores(6)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 15).Value = Valores(7)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 16).Value = Valores(8)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 17).Value = Valores(9)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 18).Value = Valores(10)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 19).Value = Valores(11)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 20).Value = Valores(12)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 21).Value = Valores(13)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 22).Value = Valores(14)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 23).Value = Valores(15)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 24).Value = Valores(16)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 25).Value = Valores(17)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 26).Value = Valores(18)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 27).Value = Valores(19)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 28).Value = Valores(20)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 29).Value = Valores(21)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 30).Value = Valores(22)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 31).Value = Valores(23)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 32).Value = Valores(24)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 33).Value = Valores(25)
                        Workbooks(sArquivoDestino).Save
        
                        iLinha = 1
                        ln = 1
                        'zerando o vetor que vai receber os valores
                        For x = 1 To 25
                            Valores(x) = 0
                        x = x + 1
                        Next
                        i = i + 1
                        If i = 2 Then
                            GoTo SalvarArquivo
                        End If
                        Exit Do
                    End If
                Loop Until Len(sPTRESOrigem) = 0
            End If
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sPTRES) = 0
    Else
        Exit Sub
    End If
   

SalvarArquivo:
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivoDestino & "..."
    Workbooks(sArquivoDestino).Close SaveChanges:=True
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo JUNTARVALORESEXECUCAOPTRES() gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub JUNTARVALORESEXECUCAORP(lm As Integer)
    Dim Funcional(1) As String, RPrimario(0 To 1) As Integer, UGestora(0 To 1) As String, sFuncionalOrigem As String, i As Integer, Valores(1 To 19) As Double, ln As Long, _
        sArquivoOrigem As String, sArquivoDestino As String, x As Integer, sRP As Integer, sUG As String, sRPOrigem As Integer, sUGOrigem As String
    
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre os arquivos que serão processados
    sArquivoDestino = "EDIT_TB_IMP_EXECUCAO_RP.xlsx"
    sArquivoOrigem = "EDIT_TB_IMP_RECEBIDO_RP.xlsx"
    
    Funcional(0) = "26.784.2086.212A.0030"
    RPrimario(0) = 3
    UGestora(0) = "393003"
    
    'zerando o vetor que vai receber os valores
    For x = 1 To 19
        Valores(x) = 0
        x = x + 1
    Next
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivoDestino
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivoDestino, UpdateLinks:=False
        
        'Inicializar o contador de linhas do Arquivo destino e origem
        iLinha = 1
        i = 0
        ln = 1
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 1).Value
            sRP = Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 3).Value
            sUG = Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 4).Value
                                    
            If (sFuncional = Funcional(i) And sRP = RPrimario(i) And sUG = UGestora(i)) Then
                'Recebendo os valores da linha selecionada
                Valores(1) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 5).Value)
                Valores(2) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 6).Value)
                Valores(3) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 7).Value)
                Valores(4) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 8).Value)
                Valores(5) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 9).Value)
                Valores(6) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 10).Value)
                Valores(7) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 11).Value)
                Valores(8) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 12).Value)
                Valores(9) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 13).Value)
                Valores(10) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 14).Value)
                Valores(11) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 15).Value)
                Valores(12) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 16).Value)
                Valores(13) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 17).Value)
                Valores(14) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 18).Value)
                Valores(15) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 19).Value)
                Valores(16) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 20).Value)
                Valores(17) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 21).Value)
                Valores(18) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 22).Value)
                Valores(19) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 23).Value)
                
                'Abrir o Arquivo Origem para importação dos valores
                Application.StatusBar = "Abrindo " & sArquivoOrigem
                DoEvents  'Passa operação para o Sistema operacional
                Workbooks.Open Filename:=sCaminho & "\" & sArquivoOrigem, UpdateLinks:=False
                
                Do
                    'Incrementar o contador de linhas do WBArquivo
                    ln = ln + 1
            
                    'Armazenar a informação da rubrica da linha em questão
                    sFuncionalOrigem = Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 1).Value
                    sRPOrigem = Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 3).Value
                    sUGOrigem = Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 4).Value
                    If (sFuncionalOrigem = Funcional(i) And sRPOrigem = RPrimario(i) And sUGOrigem = UGestora(i)) Then
                        'Recebendo os valores da linha selecionada
                        Valores(1) = Valores(1) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 5).Value)
                        Valores(2) = Valores(2) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 6).Value)
                        Valores(3) = Valores(3) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 7).Value)
                        Valores(4) = Valores(4) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 8).Value)
                        Valores(5) = Valores(5) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 9).Value)
                        Valores(6) = Valores(6) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 10).Value)
                        Valores(7) = Valores(7) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 11).Value)
                        Valores(8) = Valores(8) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 12).Value)
                        Valores(9) = Valores(9) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 13).Value)
                        Valores(10) = Valores(10) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 14).Value)
                        Valores(11) = Valores(11) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 15).Value)
                        Valores(12) = Valores(12) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 16).Value)
                        Valores(13) = Valores(13) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 17).Value)
                        Valores(14) = Valores(14) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 18).Value)
                        Valores(15) = Valores(15) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 19).Value)
                        Valores(16) = Valores(16) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 20).Value)
                        Valores(17) = Valores(17) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 21).Value)
                        Valores(18) = Valores(18) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 22).Value)
                        Valores(19) = Valores(19) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 23).Value)
                        
                        'Selecionando e apagando a linha copiada
                        Rows(ln).Select
                        Rows(ln).Delete
                        
                        'Fechar o WBarquivo, salvando como
                        DoEvents
                        Application.StatusBar = "Fechando o arquivo " & sArquivoOrigem & "..."
                        Workbooks(sArquivoOrigem).Close SaveChanges:=True
                        Application.StatusBar = "Voltando para o arquivo " & sArquivoDestino & "..."
                        
                        'Atualizar Arquivo Destino
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 5).Value = Valores(1)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 6).Value = Valores(2)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 7).Value = Valores(3)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 8).Value = Valores(4)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 9).Value = Valores(5)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 10).Value = Valores(6)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 11).Value = Valores(7)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 12).Value = Valores(8)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 13).Value = Valores(9)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 14).Value = Valores(10)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 15).Value = Valores(11)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 16).Value = Valores(12)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 17).Value = Valores(13)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 18).Value = Valores(14)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 19).Value = Valores(15)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 20).Value = Valores(16)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 21).Value = Valores(17)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 22).Value = Valores(18)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 23).Value = Valores(19)
                        Workbooks(sArquivoDestino).Save
        
                        iLinha = 1
                        ln = 1
                        'zerando o vetor que vai receber os valores
                        For x = 1 To 19
                            Valores(x) = 0
                        x = x + 1
                        Next
                        i = i + 1
                        
                        If i = 1 Then
                            GoTo SalvarArquivo
                        End If
                        Exit Do
                    End If
                Loop Until Len(sFuncionalOrigem) = 0
            End If
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
    Else
        Exit Sub
    End If
   

SalvarArquivo:
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivoDestino & "..."
    Workbooks(sArquivoDestino).Close SaveChanges:=True
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo JUNTARVALORESEXECUCAORP() gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

Sub JUNTARVALORESNATUREZA(lm As Integer)
    Dim Funcional(0 To 6) As String, PTRES(0 To 6) As String, Natureza(0 To 6) As String, Fonte(0 To 6) As String, RPrimario(0 To 6) As Integer, _
        sFuncionalOrigem As String, i As Integer, sRP As Integer, sNatureza As String, sFonte As String, _
        Valores(1 To 25) As Double, ln As Long, sArquivoOrigem As String, sArquivoDestino As String, _
        x As Integer, sPTRES As String, sPTRESOrigem As String, sRPOrigem As Integer, _
        sFonteOrigem As String, sNaturezaOrigem As String
   
    On Error GoTo DeuErro
    
    'Armazenar as informações iniciais sobre os arquivos que serão processados
    sArquivoDestino = "EDIT_TB_IMP_NATUREZA_PTRES.xlsx"
    sArquivoOrigem = "EDIT_TB_IMP_NATUREZA_REC.xlsx"
    
    Funcional(0) = "26.122.0032.20TP.0001"
    PTRES(0) = "194825"
    Natureza(0) = "3.1.90.00"
    Fonte(0) = "0100.000000"
    RPrimario(0) = 1
    
    Funcional(1) = "26.122.0032.20TP.0001"
    PTRES(1) = "194825"
    Natureza(1) = "3.1.90.00"
    Fonte(1) = "0944.000000"
    RPrimario(1) = 1
    
    Funcional(2) = "26.122.0032.20TP.0001"
    PTRES(2) = "194825"
    Natureza(2) = "3.1.90.96"
    Fonte(2) = "0100.000000"
    RPrimario(2) = 1
    
    Funcional(3) = "26.301.0032.212B.0001"
    PTRES(3) = "194832"
    Natureza(3) = "3.3.90.00"
    Fonte(3) = "0100.000000"
    RPrimario(3) = 1
    
    Funcional(4) = "26.301.0032.212B.0001"
    PTRES(4) = "194832"
    Natureza(4) = "3.3.90.00"
    Fonte(4) = "0944.000000"
    RPrimario(4) = 1
    
    Funcional(5) = "26.301.0032.212B.0001"
    PTRES(5) = "194834"
    Natureza(5) = "3.3.90.00"
    Fonte(5) = "0100.000000"
    RPrimario(5) = 1
    
    Funcional(6) = "26.301.0032.212B.0001"
    PTRES(6) = "194834"
    Natureza(6) = "3.3.90.00"
    Fonte(6) = "0944.000000"
    RPrimario(6) = 1
    
          
    'zerando o vetor que vai receber os valores
    For x = 1 To 25
        Valores(x) = 0
        x = x + 1
    Next
    
    If UCase(Workbooks(sWBMacro).Worksheets(sPlanAtiva).Cells(lm, 2)) = "SIM" Then
    
        'Abrir o WBArquivo em questão
        Application.StatusBar = "Abrindo " & sArquivoDestino
        DoEvents  'Passa operação para o Sistema operacional
        Workbooks.Open Filename:=sCaminho & "\" & sArquivoDestino, UpdateLinks:=False
        
        'Inicializar o contador de linhas do Arquivo destino e origem
        iLinha = 1
        i = 0
        ln = 1
        
        'Loop 1.1: Loop nas linhas de dados do arquivo
        Do
        
            'Incrementar o contador de linhas do WBArquivo
            iLinha = iLinha + 1
            
            'Armazenar a informação da rubrica da linha em questão
            sFuncional = Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 1).Value
            sPTRES = Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 3).Value
            sRP = Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 4).Value
            sNatureza = Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 5).Value
            sFonte = Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 6).Value
                                                
            If (sFuncional = Funcional(i) And sPTRES = PTRES(i) And sRP = RPrimario(i) And sNatureza = Natureza(i) And sFonte = Fonte(i)) Then
                'Recebendo os valores da linha selecionada
                Valores(1) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 11).Value)
                Valores(2) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 12).Value)
                Valores(3) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 13).Value)
                Valores(4) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 14).Value)
                Valores(5) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 15).Value)
                Valores(6) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 16).Value)
                Valores(7) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 17).Value)
                Valores(8) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 18).Value)
                Valores(9) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 19).Value)
                Valores(10) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 20).Value)
                Valores(11) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 21).Value)
                Valores(12) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 22).Value)
                Valores(13) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 23).Value)
                Valores(14) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 24).Value)
                Valores(15) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 25).Value)
                Valores(16) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 26).Value)
                Valores(17) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 27).Value)
                Valores(18) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 28).Value)
                Valores(19) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 29).Value)
                Valores(20) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 30).Value)
                Valores(21) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 31).Value)
                Valores(22) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 32).Value)
                Valores(23) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 33).Value)
                Valores(24) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 34).Value)
                Valores(25) = (Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 35).Value)
                
                'Abrir o Arquivo Origem para importação dos valores
                Application.StatusBar = "Abrindo " & sArquivoOrigem
                DoEvents  'Passa operação para o Sistema operacional
                Workbooks.Open Filename:=sCaminho & "\" & sArquivoOrigem, UpdateLinks:=False
                
                Do
                    'Incrementar o contador de linhas do WBArquivo
                    ln = ln + 1
            
                    'Armazenar a informação da rubrica da linha em questão
                    sFuncionalOrigem = Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 1).Value
                    sPTRESOrigem = Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 3).Value
                    sRPOrigem = Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 4).Value
                    sNaturezaOrigem = Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 5).Value
                    sFonteOrigem = Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 6).Value
                    If (sFuncionalOrigem = Funcional(i) And sPTRESOrigem = PTRES(i) And sRPOrigem = RPrimario(i) And sNaturezaOrigem = Natureza(i) And sFonteOrigem = Fonte(i)) Then
                        'Recebendo os valores da linha selecionada
                        Valores(1) = Valores(1) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 11).Value)
                        Valores(2) = Valores(2) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 12).Value)
                        Valores(3) = Valores(3) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 13).Value)
                        Valores(4) = Valores(4) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 14).Value)
                        Valores(5) = Valores(5) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 15).Value)
                        Valores(6) = Valores(6) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 16).Value)
                        Valores(7) = Valores(7) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 17).Value)
                        Valores(8) = Valores(8) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 18).Value)
                        Valores(9) = Valores(9) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 19).Value)
                        Valores(10) = Valores(10) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 20).Value)
                        Valores(11) = Valores(11) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 21).Value)
                        Valores(12) = Valores(12) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 22).Value)
                        Valores(13) = Valores(13) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 23).Value)
                        Valores(14) = Valores(14) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 24).Value)
                        Valores(15) = Valores(15) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 25).Value)
                        Valores(16) = Valores(16) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 26).Value)
                        Valores(17) = Valores(17) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 27).Value)
                        Valores(18) = Valores(18) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 28).Value)
                        Valores(19) = Valores(19) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 29).Value)
                        Valores(20) = Valores(20) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 30).Value)
                        Valores(21) = Valores(21) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 31).Value)
                        Valores(22) = Valores(22) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 32).Value)
                        Valores(23) = Valores(23) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 33).Value)
                        Valores(24) = Valores(24) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 34).Value)
                        Valores(25) = Valores(25) + (Workbooks(sArquivoOrigem).Worksheets(1).Cells(ln, 35).Value)
                        
                        'Selecionando e apagando a linha copiada
                        Rows(ln).Select
                        Rows(ln).Delete
                        
                        'Fechar o WBarquivo, salvando como
                        DoEvents
                        Application.StatusBar = "Fechando o arquivo " & sArquivoOrigem & "..."
                        Workbooks(sArquivoOrigem).Close SaveChanges:=True
                        Application.StatusBar = "Voltando para o arquivo " & sArquivoDestino & "..."
                        
                        'Atualizar Arquivo Destino
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 11).Value = Valores(1)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 12).Value = Valores(2)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 13).Value = Valores(3)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 14).Value = Valores(4)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 15).Value = Valores(5)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 16).Value = Valores(6)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 17).Value = Valores(7)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 18).Value = Valores(8)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 19).Value = Valores(9)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 20).Value = Valores(10)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 21).Value = Valores(11)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 22).Value = Valores(12)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 23).Value = Valores(13)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 24).Value = Valores(14)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 25).Value = Valores(15)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 26).Value = Valores(16)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 27).Value = Valores(17)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 28).Value = Valores(18)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 29).Value = Valores(19)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 30).Value = Valores(20)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 31).Value = Valores(21)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 32).Value = Valores(22)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 33).Value = Valores(23)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 34).Value = Valores(24)
                        Workbooks(sArquivoDestino).Worksheets(1).Cells(iLinha, 35).Value = Valores(25)
                        Workbooks(sArquivoDestino).Save
        
                        iLinha = 1
                        ln = 1
                        'zerando o vetor que vai receber os valores
                        For x = 1 To 25
                            Valores(x) = 0
                        x = x + 1
                        Next
                        i = i + 1
                        
                        If i = 7 Then
                            GoTo SalvarArquivo
                        End If
                        Exit Do
                    End If
                Loop Until Len(sFuncionalOrigem) = 0
            End If
        'Fim do Loop 1.1: Procurar a primeira linha com os dados (tratamento de arquivos nos quais a legenda possui mais de uma linha)
        Loop Until Len(sFuncional) = 0
    Else
        Exit Sub
    End If
   

SalvarArquivo:
        
    'Fechar o WBarquivo, salvando como
    DoEvents
    Application.StatusBar = "Fechando o arquivo " & sArquivoDestino & "..."
    Workbooks(sArquivoDestino).Close SaveChanges:=True
               
    Exit Sub
    
DeuErro:
    MsgBox "O Arquivo JUNTARVALORESNATUREZA() gerou erro: " + Error, vbCritical, "Erro na Execução da Macro"
    Exit Sub
End Sub

