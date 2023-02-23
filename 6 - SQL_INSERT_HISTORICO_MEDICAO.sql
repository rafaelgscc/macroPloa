USE DBPLOAWEB 
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/* Tabela Temporária de Histórico de Medição */
DELETE FROM [dbo].[TB_IMP_HISTORICO_MEDICAO]   
GO

INSERT INTO [DBPLOAWEB].[dbo].[TB_IMP_HISTORICO_MEDICAO]

SELECT LTRIM(M.NU_CON_FORMATADO) AS NU_CON_FORMATADO
      ,(SUBSTRING(M.DS_ANO_MES,6,2) + '/' + SUBSTRING(M.DS_ANO_MES,1,4)) AS DT_MEDICAO
	  ,CONVERT(datetime, M.DT_INICIO_MEDICAO, 103) AS DT_INICIO 
	  ,CONVERT(datetime, M.DT_TERMINO_MEDICAO, 103) AS DT_TERMINO  
	  ,CONVERT(datetime, M.DT_PROCESSAMENTO_MEDICAO, 103) AS DT_PROCESSAMENTO 
	  ,CONVERT(int, M.NU_MEDICAO) AS NU_MEDICAO 
	  ,SUM(M.VLR_MEDICAO_PI) AS VLR_MEDICAO_PI
	  ,SUM(M.VLR_MEDICAO_PIR) AS VLR_MEDICAO_PIR
	  ,SUM(M.VLR_REAJUSTE_MEDICAO) AS VLR_REAJUSTE_MEDICAO
	  ,C.Processo
FROM [10.100.10.65\MSSQL].[SIMDNIT].[dbo].[Dados_Medicao] M INNER JOIN VW_CONTRATO C ON (C.NumContrato = M.NU_CON_FORMATADO)
WHERE LEN(LTRIM(M.NU_CON_FORMATADO)) = 13 AND LEN(M.DS_ANO_MES) = 7 AND M.NU_MEDICAO > 0 AND (M.VLR_MEDICAO_PI + M.VLR_MEDICAO_PIR + M.VLR_REAJUSTE_MEDICAO) > 0 AND
      M.NU_CON_FORMATADO NOT IN ('10 00012/2002', '00 00001/2014', '00 00948/2010', '00 00002/2014', '11 00128/2018', '11 00955/2017', '14 00508/2016')
	                             
GROUP BY M.NU_CON_FORMATADO, M.DS_ANO_MES, M.DT_INICIO_MEDICAO, M.DT_TERMINO_MEDICAO, M.DT_PROCESSAMENTO_MEDICAO, M.NU_MEDICAO, C.Processo 

GO

INSERT INTO [DBPLOAWEB].[dbo].[TB_HISTORICO_MEDICAO]

SELECT I.Processo
      ,I.DT_Medicao
      ,I.DT_Inicio_Medicao
      ,I.DT_Termino_Medicao
      ,I.DT_Processamento
      ,I.NumMedicao
      ,I.VL_PI_Medicao
      ,I.VL_Medicao_PI_R
      ,I.VL_Reajuste_Medicao
	  ,'SISTEMA'
	  ,CONVERT(datetime, GETDATE(), 103)
	  
  FROM TB_IMP_HISTORICO_MEDICAO I 
  WHERE NOT EXISTS (SELECT * FROM TB_HISTORICO_MEDICAO H WHERE H.Processo = I.Processo AND H.NumMedicao = I.NumMedicao)

GO


UPDATE [DBPLOAWEB].[dbo].[TB_HISTORICO_MEDICAO]
   SET [Processo] = I.Processo
      ,[DT_Medicao] = I.DT_Medicao
      ,[DT_Inicio_Medicao] = I.DT_Inicio_Medicao
      ,[DT_Termino_Medicao] = I.DT_Termino_Medicao
      ,[DT_Processamento] = I.DT_Processamento
      ,[NumMedicao] = I.NumMedicao
      ,[VL_PI_Medicao] = I.VL_PI_Medicao
      ,[VL_Medicao_PI_R] = I.VL_Medicao_PI_R
      ,[VL_Reajuste_Medicao] = I.VL_Reajuste_Medicao
	  ,[Usuario] = 'SISTEMA'
	  ,[DT_Atualizacao] = CONVERT(datetime, GETDATE(), 103)
	FROM [dbo].[TB_HISTORICO_MEDICAO] H INNER JOIN TB_IMP_HISTORICO_MEDICAO I ON (I.Processo = H.Processo AND I.NumMedicao = H.NumMedicao) 
GO

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('VW_TEMP_HISTORICO_MEDICAO') AND sysstat & 0xf = 2)
DROP VIEW [VW_TEMP_HISTORICO_MEDICAO]
GO

CREATE VIEW VW_TEMP_HISTORICO_MEDICAO AS (
SELECT H.Processo
      ,H.DT_Medicao
      ,H.DT_Inicio_Medicao
      ,H.DT_Termino_Medicao
      ,H.DT_Processamento
      ,MAX(H.NumMedicao) AS NumMedicao
      ,H.VL_PI_Medicao
      ,H.VL_Medicao_PI_R
      ,H.VL_Reajuste_Medicao
FROM TB_HISTORICO_MEDICAO H
GROUP BY H.Processo, H.DT_Medicao, H.DT_Inicio_Medicao, H.DT_Termino_Medicao, H.DT_Processamento, H.VL_PI_Medicao, H.VL_Medicao_PI_R, H.VL_Reajuste_Medicao)
GO

UPDATE [DBPLOAWEB].[dbo].[TB_PROCESSO]
   SET [NumMedicaoAtual] = V.NumMedicao
      ,[DT_MedicaoAtual] = V.DT_Medicao 
      ,[VL_PI_Medicao] = V.VL_PI_Medicao
      ,[VL_Medicao_PI_R] = V.VL_Medicao_PI_R
      ,[VL_Reajuste_Medicao] = V.VL_Reajuste_Medicao
	FROM [dbo].[TB_PROCESSO] P INNER JOIN [dbo].[VW_TEMP_HISTORICO_MEDICAO] V ON (V.Processo = P.Processo) 
GO