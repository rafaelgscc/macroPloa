USE DBPLOAWEB 
GO

/* Tabela Temporária de contratos */
DELETE FROM [DBPLOAWEB].[dbo].[TB_IMP_CONTRATO]  
GO

INSERT INTO [DBPLOAWEB].[dbo].[TB_IMP_CONTRATO]
      
	SELECT DISTINCT LTRIM(C.NU_CON_FORMATADO) AS NU_CON_FORMATADO
		  ,IIF(RTRIM(C.DS_TIP_CONTRATO) = 'SUPERVISAO', 'SUPERVISÃO', RTRIM(C.DS_TIP_CONTRATO)) AS DS_TIP_CONTRATO
		  ,IIF(LEFT(RTRIM(C.SG_UND_GESTORA), 8) = 'S.R.E - ', 'SRE-'+ RTRIM(SUBSTRING(C.SG_UND_GESTORA,9,10)), RTRIM(C.SG_UND_GESTORA)) AS SG_UND_GESTORA
		  ,RTRIM(C.ds_fas_contrato) AS DS_FAS_CONTRATO
		  ,CONVERT(date, C.DT_BASE, 103) AS DT_BASE
		  ,CONVERT(date, C.DT_INICIO, 103) AS DT_INICIO
		  ,CONVERT(date, C.DT_TERMINO_VIGENCIA, 103) AS DT_TERMINO
		  ,C.Valor_Inicial
		  ,C.Valor_Total_de_Aditivos
		  ,C.Valor_Total_de_Reajuste
		  ,0
	      ,0
	      ,0
		  ,0
		  ,0
	      ,0
	      ,''
	      ,(SUBSTRING(C.NU_CNPJ_CPF,1,2) + '.' + SUBSTRING(C.NU_CNPJ_CPF,3,3) + '.' + SUBSTRING(C.NU_CNPJ_CPF,6,3) + '/' + SUBSTRING(C.NU_CNPJ_CPF,9,4) + '-' + SUBSTRING(C.NU_CNPJ_CPF,13,2)) AS CNPJ 
	      ,C.NO_EMPRESA
		  ,0
    FROM [SEDEDF315BSA\MSSQL].[SIMDNIT].[dbo].[Dados_Contrato] C INNER JOIN [dbo].VW_CONTRATO V ON (V.NumContrato = C.NU_CON_FORMATADO)
    WHERE LEN(C.NU_CON_FORMATADO) = 13 AND LEN(C.SG_UND_GESTORA) <= 10 AND YEAR(C.DT_Base) >= 2010 AND LTRIM(C.NU_CON_FORMATADO) NOT IN ('01 00458/2014','07 00835/2018','16 00503/2015')
GO

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('VW_TEMP_MEDICAO') AND sysstat & 0xf = 2)
	DROP VIEW [VW_TEMP_MEDICAO]
GO

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('VW_TEMP_DADOSEMPENHO') AND sysstat & 0xf = 2)
	DROP VIEW [VW_TEMP_DADOSEMPENHO]
GO

CREATE VIEW [VW_TEMP_MEDICAO]AS(
SELECT M.NU_CON_FORMATADO AS NumContrato, ISNULL(SUM(M.VL_PI_MEDICAO), 0) AS VL_PI_Medicao, ISNULL(SUM(M.VL_PI_R_MEDICAO), 0) AS VL_Medicao_PI_R,
       ISNULL(SUM(M.VL_REA_MEDICAO), 0) AS VL_Reajuste_Medicao, MAX(M.NU_MEDICAO) AS NumMedicaoAtual, MAX(M.DS_ANO_MES) AS DT_Medicao  
FROM [SEDEDF315BSA\MSSQL].[SIMDNIT].[dbo].[Dados_Medicao] M
WHERE LEN(M.NU_CON_FORMATADO) = 13 AND M.NU_MEDICAO > 0
GROUP BY M.NU_CON_FORMATADO
)
GO

CREATE VIEW [VW_TEMP_DADOSEMPENHO]AS(
SELECT E.NU_CON_FORMATADO AS NumContrato, ISNULL(SUM(E.VLR_EMPENHO_CONSUMIDO), 0) AS VL_Consumido, ISNULL(SUM(E.VLR_EMPENHO_INICIAL) + SUM(E.VLR_EMPENHO_AJUSTES), 0) AS VL_Empenhado,
       ISNULL(SUM(E.VLR_EMPENHO_SALDO), 0) AS VL_SaldoEmpenho  
FROM [SEDEDF315BSA\MSSQL].[SIMDNIT].[dbo].[Dados_Empenho] E
WHERE LEN(E.NU_CON_FORMATADO) = 13 
GROUP BY E.NU_CON_FORMATADO
)
GO

UPDATE [dbo].[TB_IMP_CONTRATO]
   SET [VL_PI_Medicao] = ISNULL(V.VL_PI_Medicao,0)
      ,[VL_Medicao_PI_R] = ISNULL(V.VL_Medicao_PI_R,0)
      ,[VL_Reajuste_Medicao] = ISNULL(V.VL_Reajuste_Medicao,0)
      ,[VL_Consumido] = ISNULL(E.VL_Consumido,0)
      ,[VL_Empenhado] = ISNULL(E.VL_Empenhado,0)
      ,[NumMedicaoAtual] = ISNULL(V.NumMedicaoAtual,0)
      ,[DT_Medicao] = ISNULL((SUBSTRING(V.DT_Medicao,6,2) + '/' + SUBSTRING(V.DT_Medicao,1,4)), '') 
      ,[VL_SaldoEmpenho] = ISNULL(E.VL_SaldoEmpenho,0)
   FROM TB_IMP_CONTRATO I LEFT JOIN VW_TEMP_MEDICAO V ON (V.NumContrato = I.NumContrato)
                          LEFT JOIN VW_TEMP_DADOSEMPENHO E ON (E.NumContrato = I.NumContrato)
GO

UPDATE [dbo].[TB_PROCESSO]
   SET [Situacao] = C.Situacao
	  ,[MesReajuste] = CASE MONTH(C.DT_BASE) WHEN 1 THEN 'Janeiro'  
		                         WHEN 2 THEN 'Fevereiro'
								 WHEN 3 THEN 'Março'
								 WHEN 4 THEN 'Abril'
								 WHEN 5 THEN 'Maio'
								 WHEN 6 THEN 'Junho'
								 WHEN 7 THEN 'Julho'
								 WHEN 8 THEN 'Agosto' 
								 WHEN 9 THEN 'Setembro'
								 WHEN 10 THEN 'Outubro'
								 WHEN 11 THEN 'Novembro'
								 WHEN 12 THEN 'Dezembro'
								 END
	  ,[DT_Inicio] = C.DT_Inicio
      ,[DT_Termino] = C.DT_Termino
      ,[VL_Inicial] = C.VL_Inicial
      ,[VL_Aditivo] = C.VL_Aditivo
      ,[VL_Reajustamento] = C.VL_Reajustamento
	  ,[VL_PI_Medicao] = C.VL_PI_Medicao
      ,[VL_Medicao_PI_R] = C.VL_Medicao_PI_R
      ,[VL_Reajuste_Medicao] = C.VL_Reajuste_Medicao
	  ,[NumMedicaoAtual] = C.NumMedicaoAtual
      ,[DT_MedicaoAtual] = C.DT_Medicao
	  ,[DT_Atualizacao] = CONVERT(datetime, GETDATE(), 103)
	  ,[Usuario] = 'SISTEMA'
   FROM [DBPLOAWEB].[dbo].[TB_PROCESSO] P INNER JOIN [DBPLOAWEB].[dbo].[TB_IMP_CONTRATO] C ON (C.NumContrato = P.NumContrato)
   WHERE LEN(P.NumContrato ) = 13 AND P.AtualizarContrato = 1
GO
