USE DBPLOAWEB 
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/* Deletando registro da tabela Temporária de Histórico de Medição */
DELETE FROM [dbo].[TB_IMP_HISTORICO_MEDICAO]   
GO

/*Importando os registros de medição de outro servidor */
/*Inserindo os registros na tabela temporária */
INSERT INTO [DBPLOAWEB].[dbo].[TB_IMP_HISTORICO_MEDICAO]

SELECT LTRIM(M.NU_CON_FORMATADO) AS NU_CON_FORMATADO
      ,(SUBSTRING(M.DS_ANO_MES,6,2) + '/' + SUBSTRING(M.DS_ANO_MES,1,4)) AS DT_MEDICAO
	  ,CONVERT(datetime, M.DT_INICIO_MEDICAO, 103) AS DT_INICIO 
	  ,CONVERT(datetime, M.DT_TERMINO_MEDICAO, 103) AS DT_TERMINO  
	  ,CONVERT(datetime, M.DT_PROCESSAMENTO_MEDICAO, 103) AS DT_PROCESSAMENTO 
	  ,CONVERT(int, M.NU_MEDICAO) AS NU_MEDICAO 
	  ,SUM(M.VL_PI_MEDICAO) AS VLR_MEDICAO_PI
	  ,SUM(M.VL_PI_MEDICAO + M.VL_REA_MEDICAO) AS VLR_MEDICAO_PIR
	  ,SUM(M.VL_REA_MEDICAO) AS VLR_REAJUSTE_MEDICAO
	  ,C.Processo
FROM [SEDEDF315BSA\MSSQL].[SIMDNIT].[dbo].[Dados_Medicao] M INNER JOIN VW_CONTRATO C ON (C.NumContrato = M.NU_CON_FORMATADO)
WHERE LEN(LTRIM(M.NU_CON_FORMATADO)) = 13 AND LEN(M.DS_ANO_MES) = 7 AND M.NU_MEDICAO > 0 AND (M.VL_PI_MEDICAO + M.VL_PI_R_MEDICAO + M.VL_REA_MEDICAO) > 0 AND 
      NOT EXISTS(SELECT * FROM [DBPLOAWEB].[dbo].[TB_IMP_HISTORICO_MEDICAO] H WHERE H.NumContrato = M.NU_CON_FORMATADO AND H.NumMedicao = CONVERT(int, M.NU_MEDICAO))
GROUP BY M.NU_CON_FORMATADO, M.DS_ANO_MES, M.DT_INICIO_MEDICAO, M.DT_TERMINO_MEDICAO, M.DT_PROCESSAMENTO_MEDICAO, M.NU_MEDICAO, C.Processo 

GO

/*identificando e eliminando os registros duplicados */
SELECT Processo, NumContrato, NumMedicao, COUNT(*) AS Qtd FROM TB_IMP_HISTORICO_MEDICAO GROUP BY Processo, NumContrato, NumMedicao HAVING COUNT(*) > 1 
GO

WITH cte AS (SELECT Processo, NumContrato NumMedicao, ROW_NUMBER() OVER (PARTITION BY Processo, NumContrato, NumMedicao ORDER BY Processo, NumContrato, NumMedicao) linha FROM TB_IMP_HISTORICO_MEDICAO)
SELECT * FROM cte WHERE linha > 1
GO

WITH cte AS (SELECT Processo, NumContrato, NumMedicao, ROW_NUMBER() OVER (PARTITION BY Processo, NumContrato, NumMedicao ORDER BY Processo, NumContrato, NumMedicao) linha FROM TB_IMP_HISTORICO_MEDICAO)
DELETE FROM cte WHERE linha > 1
GO

/*Após eliminar duplicidade inserir novos registros na tabela principal */
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
	  ,NULL
	  
  FROM TB_IMP_HISTORICO_MEDICAO I 
  WHERE NOT EXISTS (SELECT * FROM TB_HISTORICO_MEDICAO H WHERE H.Processo = I.Processo AND H.NumMedicao = I.NumMedicao)

GO

/*Atualizar todos os registros existentes na tabela principal */
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

/*Deletar uma view temporária */
IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('VW_TEMP_ULTIMA_MEDICAO') AND sysstat & 0xf = 2)
DROP VIEW [VW_TEMP_ULTIMA_MEDICAO]
GO

/*Criar uma view temporária para processamentos */
CREATE VIEW VW_TEMP_ULTIMA_MEDICAO AS (
SELECT H.Processo
      ,MAX(H.NumMedicao) AS NumMedicao
FROM TB_HISTORICO_MEDICAO H
GROUP BY H.Processo)
GO

/*Atualizar valores de medição na tabela processo */
UPDATE [DBPLOAWEB].[dbo].[TB_PROCESSO]
   SET [NumMedicaoAtual] = V.NumMedicao
      ,[DT_MedicaoAtual] = M.DT_Medicao 
      ,[VL_PI_Medicao] = M.VL_PI_Medicao
      ,[VL_Medicao_PI_R] = M.VL_Medicao_PI_R
      ,[VL_Reajuste_Medicao] = M.VL_Reajuste_Medicao
	FROM [dbo].[TB_PROCESSO] P INNER JOIN [dbo].[VW_TEMP_ULTIMA_MEDICAO] V ON (V.Processo = P.Processo) 
	                           INNER JOIN [dbo].[TB_HISTORICO_MEDICAO] M ON (M.Processo = P.Processo AND M.NumMedicao = V.NumMedicao)
GO