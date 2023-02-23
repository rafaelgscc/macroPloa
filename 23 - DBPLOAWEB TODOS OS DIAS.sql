USE DBPLOAWEB 
GO

--CONFIGURANDO À INSTÂNCIA SQL PARA ACEITAR OPÇÕES AVANÇADAS
EXEC sp_configure 'show advanced options', 1
RECONFIGURE
GO

--HABILITANDO O USO DE CONSULTAS DISTRIBUÍDAS
EXEC sp_configure 'Ad Hoc Distributed Queries', 1
RECONFIGURE
GO

EXEC DBPLOAWEB.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1
GO

EXEC DBPLOAWEB.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 1
GO

/* 1) Tabela Temporária Nota Empenho Detalhada */
DELETE FROM [dbo].[TB_IMP_NE]  
GO

/* 2) Tabela Temporária Nota Empenho Detalhada */
INSERT INTO TB_IMP_NE
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_NE_2022_0.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_NE_2022_0$]') 
GO

/* 3) Tabela Temporária Nota Empenho Detalhada */
INSERT INTO TB_IMP_NE
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_NE_2022_1.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_NE_2022_1$]') 
GO

/* 4) Tabela Temporária Nota Empenho Detalhada */
INSERT INTO TB_IMP_NE
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_NE_2023_0.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_NE_2023_0$]') 
GO

/* 5) Tabela Temporária Nota Empenho Detalhada 
INSERT INTO TB_IMP_NE
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_NE_2023_1.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_NE_2023_1$]') 
GO*/

/* 6) Tabela Temporária de Execução */
DELETE FROM [dbo].[TB_IMP_EXECUCAO]
GO

/* 7) Tabela Temporária de Execução */
INSERT INTO TB_IMP_EXECUCAO
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_EXECUCAO.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_EXECUCAO$]') 
GO

/* 8) Tabela Temporária de TB_IMP_RECEBIDO */
INSERT INTO TB_IMP_EXECUCAO 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_RECEBIDO.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_RECEBIDO$]') 
GO

/* 9) Tabela Temporária de Execução por PTRES */
DELETE FROM [dbo].[TB_IMP_EXECUCAO_PTRES]
GO

/* 10) Tabela Temporária de Execução por PTRES */
INSERT INTO TB_IMP_EXECUCAO_PTRES
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_EXECUCAO_PTRES.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_EXECUCAO_PTRES$]') 
GO

/* 11) Tabela Temporária de TB_IMP_RECEBIDO_PTRES */
INSERT INTO TB_IMP_EXECUCAO_PTRES 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_RECEBIDO_PTRES.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_RECEBIDO_PTRES$]') 
GO

/* 12) Tabela Temporária de Execução com Resultado Primário */
DELETE FROM [dbo].[TB_IMP_EXECUCAO_RP]
GO

/* 13) Tabela Temporária de Execução com Resultado Primário */
INSERT INTO TB_IMP_EXECUCAO_RP
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_EXECUCAO_RP.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_EXECUCAO_RP$]') 
GO

/* 14) Tabela Temporária de Execução com Resultado Primário */
INSERT INTO TB_IMP_EXECUCAO_RP
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_RECEBIDO_RP.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_RECEBIDO_RP$]') 
GO

/* 15) Tabela Temporária de Execução por Natureza de Despesa */
DELETE FROM [dbo].[TB_IMP_NATUREZA_PTRES]
GO

/* 16) Tabela Temporária de Execução por Natureza de Despesa */
INSERT INTO TB_IMP_NATUREZA_PTRES 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_NATUREZA_PTRES.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_NATUREZA_PTRES$]') 
GO

/* 17) Tabela Temporária de Execução por Natureza de Despesa */
INSERT INTO TB_IMP_NATUREZA_PTRES 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_NATUREZA_REC.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_NATUREZA_REC$]') 
GO

/* 18) Tabela Temporária de Execução de Destaque */
DELETE FROM [dbo].[TB_IMP_DESTAQUE]
GO

/* 19) Tabela Temporária de Execução de Destaque */
INSERT INTO TB_IMP_DESTAQUE 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_DESTAQUE.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_DESTAQUE$]') 
GO

/* 20) Tabela Temporária de Execução de Destaque por PTRES */
DELETE FROM [dbo].[TB_IMP_DESTAQUE_PTRES]
GO

/* 21) Tabela Temporária de Execução de Destaque por PTRES */
INSERT INTO TB_IMP_DESTAQUE_PTRES 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_DESTAQUE_PTRES.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_DESTAQUE_PTRES$]') 
GO

/* 22) Tabela Temporária de TB_IMP_BLOQUEIO_NATUREZA */
DELETE FROM [dbo].[TB_IMP_BLOQUEIO_NATUREZA]   
GO

/* 23) Tabela Temporária de TB_IMP_BLOQUEIO_NATUREZA*/
INSERT INTO TB_IMP_BLOQUEIO_NATUREZA  
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_BLOQUEIO_NATUREZA.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_BLOQUEIO_NATUREZA$]') 
GO 

/* 24) Tabela Temporária de TB_IMP_HISTORICO_PGTO */
DELETE FROM [dbo].[TB_IMP_HISTORICO_PGTO]   
GO

/* 25) Tabela Temporária de TB_IMP_HISTORICO_PGTO */
INSERT INTO TB_IMP_HISTORICO_PGTO  
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_HISTORICO_PGTO.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_HISTORICO_PGTO$]') 
GO

/* 26) Tabela Temporária de TB_IMP_HISTORICO_PGTO2 */
INSERT INTO TB_IMP_HISTORICO_PGTO  
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_HISTORICO_PGTO2.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_HISTORICO_PGTO2$]') 
GO

/* 27) Tabela Temporária de TB_IMP_PGTO */
DELETE FROM [dbo].[TB_IMP_PGTO]    
GO

/* 28) Tabela Temporária de TB_IMP_PGTO */
INSERT INTO TB_IMP_PGTO  
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_PGTO.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_PGTO$]') 
GO

/* 29) Tabela Temporária de TB_IMP_LIMITE */
DELETE FROM [dbo].[TB_IMP_LIMITE]    
GO

/* 30) Tabela Temporária de TB_IMP_LIMITE */
INSERT INTO TB_IMP_LIMITE 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_LIMITE.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_LIMITE$]') 
GO

/* 31) Tabela Temporária de TB_IMP_SRE_ADM */
DELETE FROM [dbo].[TB_IMP_SRE_ADM]    
GO

/* 32) Tabela Temporária de TB_IMP_SRE_ADM */
INSERT INTO TB_IMP_SRE_ADM 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_SRE_ADM.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_SRE_ADM$]') 
GO

/* 33) Tabela Temporária de TB_IMP_RAP_EMPENHO */
DELETE FROM [dbo].[TB_IMP_RAP_EMPENHO]    
GO

/* 34) Tabela Temporária de TB_IMP_RAP_EMPENHO */
INSERT INTO TB_IMP_RAP_EMPENHO 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_RAP_EMPENHO.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_RAP_EMPENHO$]') 
GO

/* 35) Tabela Temporária de TB_IMP_NATUREZADETALHADA */
DELETE FROM [dbo].[TB_IMP_NATUREZADETALHADA]    
GO

/* 36) Tabela Temporária de TB_IMP_NATUREZADETALHADA */
INSERT INTO TB_IMP_NATUREZADETALHADA 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_NATUREZADETALHADA.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_NATUREZADETALHADA$]') 
GO

/* 37) Tabela Temporária de TB_IMP_DISPONIVELSRE */
DELETE FROM [dbo].[TB_IMP_DISPONIVELSRE]    
GO

/* 38) Tabela Temporária de TB_IMP_DISPONIVELSRE */
INSERT INTO TB_IMP_DISPONIVELSRE 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_DISPONIVELSRE.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_DISPONIVELSRE$]') 
GO

/* 39) Tabela Temporária de TB_IMP_INDISPONIVEL */
DELETE FROM [dbo].[TB_IMP_INDISPONIVEL]    
GO

/* 40) Tabela Temporária de TB_IMP_INDISPONIVEL */
INSERT INTO TB_IMP_INDISPONIVEL 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_INDISPONIVEL.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_INDISPONIVEL$]') 
GO

/* 41) Tabela Temporária de TB_IMP_LOA_DETALHADA */
DELETE FROM [dbo].[TB_IMP_LOA_DETALHADA]    
GO

/* 42) Tabela Temporária de TB_IMP_LOA_DETALHADA */
INSERT INTO TB_IMP_LOA_DETALHADA 
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_LOA_DETALHADA.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_LOA_DETALHADA$]') 
GO

/* 43) Tabela temporária de TB_IMP_CORFIN_RECEITA */
DELETE FROM [dbo].[TB_IMP_CORFIN_RECEITA]
GO

/* 44) Tabela Temporária de TB_IMP_CORFIN_RECEITA */
INSERT INTO TB_IMP_CORFIN_RECEITA
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_RECEITAS.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_RECEITAS$]')
GO

/* 45) Tabela Temporária de TB_IMP_CREDITOS */
DELETE FROM [dbo].[TB_IMP_CREDITOS]    
GO

/* 46) Tabela Temporária de TB_IMP_CREDITOS */
INSERT INTO TB_IMP_CREDITOS
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_CREDITOS.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_CREDITOS$]') 
GO

/* 47) Tabela Temporária de TB_IMP_CORFIN_NE_DIARIO */
DELETE FROM [dbo].[TB_IMP_CORFIN_NE_DIARIO]    
GO

/* 48) Tabela Temporária de TB_IMP_CORFIN_NE_DIARIO */
INSERT INTO TB_IMP_CORFIN_NE_DIARIO
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_CORFIN_NE_DIARIO.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_CORFIN_NE_DIARIO$]') 
GO

/* 49) Tabela Temporária de TB_IMP_UG_DESTAQUE */
DELETE FROM [dbo].[TB_IMP_UG_DESTAQUE]    
GO

/* 50) Tabela Temporária de TB_IMP_UG_DESTAQUE */
INSERT INTO TB_IMP_UG_DESTAQUE
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
'Excel 12.0; Database=\\10.100.10.174\ploa_carga\EDIT_TB_IMP_UG_DESTAQUE.xlsx; HDR=YES; IMEX=1',
'SELECT * FROM [EDIT_TB_IMP_UG_DESTAQUE$]') 
GO


/***************** INICIANDO TB_IMP_NE ****************************/

/*1) INSERINDO NA TB_UG_EXECUTORA QUANDO NÃO TIVER REFERÊNCIA */
INSERT INTO [dbo].[TB_UG_EXECUTORA]
       SELECT DISTINCT I.UGResponsavel, ('A CLASSIFICAR ' + I.UGResponsavel) AS DescricaoUG, 0, 0, 0, '', 0 FROM TB_IMP_NE I
       WHERE NOT EXISTS (SELECT * FROM TB_UG_EXECUTORA U WHERE U.UG_Executora = I.UGResponsavel)
GO

/*2) INSERINDO NA TB_NATUREZAS QUANDO NÃO TIVER REFERÊNCIA */
INSERT INTO [dbo].[TB_NATUREZAS]
       SELECT DISTINCT I.Natureza FROM TB_IMP_NE I
       WHERE NOT EXISTS (SELECT * FROM TB_NATUREZAS N WHERE N.Natureza = I.Natureza)
GO

/*3) INSERINDO NA TB_EMPREENDIMENTO QUANDO NÃO TIVER REFERÊNCIA */
INSERT INTO [dbo].[TB_EMPREENDIMENTO]
       SELECT DISTINCT I.CodMT, 'A CLASSIFICAR' AS DescricaoMT, 0 AS VL_Total, 'SISTEMA' AS Usuario, CONVERT(datetime, GETDATE(), 103) AS DT_Atualizacao FROM TB_IMP_NE I
       WHERE NOT EXISTS (SELECT * FROM TB_EMPREENDIMENTO E WHERE E.CodMT = I.CodMT)
GO

/*4) INSERINDO NA TB_NE OS EMPENHOS QUE NÃO EXISTEM (NOVOS)*/
INSERT INTO [dbo].[TB_NE]
	SELECT [AnoLancamento]
      ,[ContaCorrente]
      ,[Funcional]
      ,[Ano]
      ,[PTRES]
      ,[CodMT]
      ,[Natureza]
      ,[RPrimario]
      ,[UGResponsavel]
      ,[VL_Empenhado]
      ,[VL_Liquidado]
      ,[VL_Pago]
      ,[RP_Pro_Inscrito]
      ,[RP_Pro_Reinscrito]
      ,[RP_Pro_Cancelado]
      ,[RP_Pro_Pago]
      ,[RP_Pro_APagar]
      ,[RP_NPro_Inscrito]
      ,[RP_NPro_Reinscrito]
      ,[RP_NPro_Cancelado]
      ,[RP_NPro_ALiquidar]
      ,[RP_NPro_Liquidado]
      ,[RP_NPro_Liq_APagar]
      ,[RP_NPro_Pago]
      ,[RP_NPro_APagar]
      ,[RP_NPro_Bloqueado]
    FROM [DBPLOAWEB].[dbo].[TB_IMP_NE]
    WHERE NOT EXISTS (SELECT * FROM [DBPLOAWEB].[dbo].[TB_NE] WHERE [DBPLOAWEB].[dbo].[TB_NE].[AnoLancamento] = [DBPLOAWEB].[dbo].[TB_IMP_NE].[AnoLancamento] AND [DBPLOAWEB].[dbo].[TB_NE].[ContaCorrente] = [DBPLOAWEB].[dbo].[TB_IMP_NE].[ContaCorrente])     
GO

/*5) ATUALIZANDO NA TB_NE OS EMPENHOS QUE NÃO EXISTEM (NOVOS)*/
UPDATE [dbo].[TB_NE]
   SET [VL_Empenhado] = I.VL_Empenhado
      ,[VL_Liquidado] = I.VL_Liquidado
      ,[VL_Pago] = I.VL_Pago
      ,[RP_Pro_Inscrito] = I.RP_Pro_Inscrito
      ,[RP_Pro_Reinscrito] = I.RP_Pro_Reinscrito
      ,[RP_Pro_Cancelado] = I.RP_Pro_Cancelado
      ,[RP_Pro_Pago] = I.RP_Pro_Pago
      ,[RP_Pro_APagar] = I.RP_Pro_APagar
      ,[RP_NPro_Inscrito] = I.RP_NPro_Inscrito
      ,[RP_NPro_Reinscrito] = I.RP_NPro_Reinscrito
      ,[RP_NPro_Cancelado] = I.RP_NPro_Cancelado
      ,[RP_NPro_ALiquidar] = I.RP_NPro_ALiquidar
      ,[RP_NPro_Liquidado] = I.RP_NPro_Liquidado
      ,[RP_NPro_Liq_APagar] = I.RP_NPro_Liq_APagar
      ,[RP_NPro_Pago] = I.RP_NPro_Pago
      ,[RP_NPro_APagar] = I.RP_NPro_APagar
      ,[RP_NPro_Bloqueado] = I.RP_NPro_Bloqueado
   FROM [DBPLOAWEB].[dbo].[TB_NE] N INNER JOIN TB_IMP_NE I ON (I.AnoLancamento = N.AnoLancamento AND I.ContaCorrente = N.ContaCorrente)
GO

/*6) INSERINDO NA TB_IMP_EMPENHO OS EMPENHOS QUE NÃO EXISTEM (NOVOS) EXERCÍCIO 2022*/
INSERT INTO [dbo].[TB_IMP_EMPENHO]
	SELECT N.[ContaCorrente], N.[Funcional], N.[Ano], N.[PTRES], N.[CodMT], N.[RPrimario]
          ,N.[VL_Empenhado], N.[VL_Liquidado], N.[VL_Pago], 0, 0, 0, 0, 0, 0, ''
          ,N.[UGResponsavel], N.[Natureza], 0, convert(date,N.[DT_String],103), N.CNPJFavorecido, N.RzSocial
    FROM TB_IMP_NE N
	WHERE NOT EXISTS(SELECT * FROM TB_IMP_EMPENHO E WHERE E.ContaCorrente = N.ContaCorrente)
GO	

/*Atualizando data de emissão, cnpj e Favorecido TODOS INDEPENDENTE DO EXERCÍCIO*/
UPDATE [dbo].[TB_IMP_EMPENHO]
   SET [DT_Emissao] = convert(date,N.[DT_String],103)
      ,[CNPJFavorecido] = N.CNPJFavorecido 
	  ,[RzSocial] = N.RzSocial
   FROM TB_IMP_EMPENHO E INNER JOIN TB_IMP_NE N ON (N.ContaCorrente = E.ContaCorrente)
GO

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('VW_TEMP_NE') AND sysstat & 0xf = 2)
DROP VIEW [VW_TEMP_NE]
GO

CREATE VIEW [VW_TEMP_NE]AS(
       SELECT E.ContaCorrente, ISNULL(SUM(E.RP_Pro_Pago + E.RP_NPro_Pago), 0) AS RP_Pago, ISNULL(SUM(E.RP_Pro_Cancelado + E.RP_NPro_Cancelado), 0) AS RP_Cancelado, 
	          ISNULL(SUM(E.RP_Pro_Inscrito + E.RP_NPro_Liquidado), 0) AS RP_Liquidado 
	   FROM TB_NE E GROUP BY E.ContaCorrente
)
GO

/*7) ATUALIZA NA TB_IMP_EMPENHO OS EMPENHOS EXISTENTES DE EXECÍCIOS ANTERIORES*/
UPDATE [dbo].[TB_IMP_EMPENHO]
   SET [RP_Pago] = N.RP_Pago
      ,[RP_Cancelado] = N.RP_Cancelado
	  ,[RP_Liquidado] = N.RP_Liquidado
   FROM [DBPLOAWEB].[dbo].[TB_IMP_EMPENHO] I INNER JOIN VW_TEMP_NE N ON (I.ContaCorrente = N.ContaCorrente)
   WHERE I.Ano <= YEAR(GETDATE()-1)
GO

/*7) ATUALIZA NA TB_IMP_EMPENHO OS EMPENHOS EXISTENTES DE EXECÍCIOS ANTERIORES*/
UPDATE [dbo].[TB_IMP_EMPENHO]
   SET [RP_NPro_ALiquidar] = ISNULL(E.RP_NPro_ALiquidar,0)
      ,[RP_Bloqueado] = ISNULL(E.RP_NPro_Bloqueado,0)
   FROM [DBPLOAWEB].[dbo].[TB_IMP_EMPENHO] I LEFT JOIN TB_NE E ON (E.ContaCorrente = I.ContaCorrente AND E.AnoLancamento = YEAR(GETDATE()))
   WHERE I.Ano <= YEAR(GETDATE()-1)
GO

/*8) ATUALIZA NA TB_IMP_EMPENHO OS EMPENHOS DO EXERCICIO CORRENTE*/
UPDATE [dbo].[TB_IMP_EMPENHO]
   SET [VL_Empenhado] = E.VL_Empenhado
      ,[VL_Liquidado] = E.VL_Liquidado
      ,[VL_Pago] = E.VL_Pago
   FROM [DBPLOAWEB].[dbo].[TB_IMP_EMPENHO] I INNER JOIN TB_NE E ON (E.ContaCorrente = I.ContaCorrente AND E.AnoLancamento = I.Ano)
GO

/*9) ATUALIZA NA TB_EMPENHO CONFORME TB_IMP_EMPENHO*/
UPDATE [dbo].[TB_EMPENHO]
   SET [VL_Empenhado] = I.VL_Empenhado
      ,[VL_Liquidado] = I.VL_Liquidado
      ,[VL_Pago] = I.VL_Pago
	  ,[RP_Inscrito] = I.RP_Inscrito
      ,[RP_Liquidado] = I.RP_Liquidado
      ,[RP_Pago] = I.RP_Pago
      ,[RP_Cancelado] = I.RP_Cancelado
      ,[RP_Bloqueado] = I.RP_Bloqueado
   FROM [DBPLOAWEB].[dbo].[TB_EMPENHO] E INNER JOIN TB_IMP_EMPENHO I ON (E.ContaCorrente = I.ContaCorrente)
GO

DELETE FROM [dbo].[TB_IMP_NE]
      WHERE EXISTS(SELECT * FROM TB_NE E WHERE E.AnoLancamento = [dbo].[TB_IMP_NE].[AnoLancamento] AND E.ContaCorrente = [dbo].[TB_IMP_NE].[ContaCorrente]) AND
	        EXISTS(SELECT * FROM TB_IMP_EMPENHO E WHERE E.ContaCorrente = [dbo].[TB_IMP_NE].[ContaCorrente])
GO

SELECT * FROM [dbo].[TB_IMP_NE]
GO

/***************** FINALIZANDO TB_IMP_NE *************************/

/***************** INICIANDO TB_IMP_EXECUCAO *********************************/

UPDATE [dbo].[TB_ETAPAS]
   SET [MomentoAtual] = 'LOA'
      ,[Execucao] = IIF(((C.VL_Orc_Ini + C.VL_Orc_Aut) > 0 AND (C.RP_Pro_Inscrito + C.RP_Pro_Reinscrito + C.RP_NPro_Inscrito + C.RP_NPro_Reinscrito) = 0) , 'EXECUÇÃO', IIF(((C.VL_Orc_Ini + C.VL_Orc_Aut) > 0 AND (C.RP_Pro_Inscrito + C.RP_Pro_Reinscrito + C.RP_NPro_Inscrito + C.RP_NPro_Reinscrito) > 0), 'AMBOS', IIF(((C.VL_Orc_Ini + C.VL_Orc_Aut) = 0 AND (C.RP_Pro_Inscrito + C.RP_Pro_Reinscrito + C.RP_NPro_Inscrito + C.RP_NPro_Reinscrito) > 0), 'RAP', 'NENHUMA')))
	  ,[Externo] = C.Externo
      ,[VL_Orc_Ini] = C.VL_Orc_Ini
      ,[VL_Orc_Aut] = C.VL_Orc_Aut
      ,[VL_Empenhado] = C.VL_Empenhado
      ,[VL_Liquidado] = C.VL_Liquidado
      ,[VL_Pago] = C.VL_Pago
      ,[VL_Disponivel] = C.VL_Disponivel
      ,[VL_Contido] = C.VL_Contido
	  ,[VL_Cred_Sup] = C.VL_Cred_Sup
      ,[VL_Cred_Esp] = C.VL_Cred_Esp
      ,[VL_Cred_Ext] = C.VL_Cred_Ext
      ,[VL_Cred_Can] = C.VL_Cred_Can
	  ,[RP_Pro_Inscrito] = C.RP_Pro_Inscrito
      ,[RP_Pro_Reinscrito] = C.RP_Pro_Reinscrito
      ,[RP_Pro_Cancelado] = C.RP_Pro_Cancelado
      ,[RP_Pro_Pago] = C.RP_Pro_Pago
      ,[RP_Pro_APagar] = C.RP_Pro_APagar
      ,[RP_NPro_Inscrito] = C.RP_NPro_Inscrito
      ,[RP_NPro_Reinscrito] = C.RP_NPro_Reinscrito
      ,[RP_NPro_Cancelado] = C.RP_NPro_Cancelado
      ,[RP_NPro_ALiquidar] = C.RP_NPro_ALiquidar
      ,[RP_NPro_Liquidado] = C.RP_NPro_Liquidado
      ,[RP_NPro_Liq_APagar] = C.RP_NPro_Liq_APagar
      ,[RP_NPro_Pago] = C.RP_NPro_Pago
      ,[RP_NPro_APagar] = C.RP_NPro_APagar
      ,[RP_NPro_Bloqueado] = C.RP_NPro_Bloqueado
	  ,[RP_Inscrito] = (C.RP_Pro_Inscrito + C.RP_Pro_Reinscrito + C.RP_NPro_Inscrito + C.RP_NPro_Reinscrito)
	  ,[RP_Pago] = (C.RP_Pro_Pago + C.RP_NPro_Pago)
	  ,[RP_ExAnt] = C.RP_Pro_Inscrito
	  ,[RP_Liquidado] = C.RP_NPro_Liquidado
	  ,[RP_ALiquidar] = C.RP_NPro_ALiquidar
	  ,[RP_Cancelado] = (C.RP_Pro_Cancelado + C.RP_NPro_Cancelado)
	  ,[RP_Bloqueado] = C.RP_NPro_Bloqueado
	  ,[DT_Atualizacao] = CONVERT(datetime, GETDATE(), 103)
   FROM TB_ETAPAS E INNER JOIN TB_IMP_EXECUCAO C ON (E.Funcional = C.Funcional AND E.Ano = C.Ano)
GO

DELETE FROM [dbo].[TB_IMP_EXECUCAO]
      WHERE EXISTS(SELECT * FROM TB_ETAPAS WHERE TB_ETAPAS.Funcional = TB_IMP_EXECUCAO.Funcional AND TB_ETAPAS.Ano = TB_IMP_EXECUCAO.Ano)
GO

SELECT * FROM [dbo].[TB_IMP_EXECUCAO]
GO

/***************** FINALIZANDO TB_IMP_EXECUCAO *******************************/

/***************** INICIANDO TB_IMP_EXECUCAO_PTRES **************************/

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD') AND sysstat & 0xf = 8)
DROP TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD]
GO

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2]') AND sysstat & 0xf = 8)
DROP TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2]
GO

INSERT INTO [dbo].[TB_FUNCIONAL_PTRES]
       SELECT E.Funcional, E.Ano, E.PTRES, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '0 - Nenhuma', 0, 0, 0, 'A CLASSIFICAR', '', '', P.SiglaDir, 0, 0, '' 
	   FROM TB_IMP_UG_DESTAQUE E INNER JOIN TB_ETAPAS P ON (P.Funcional = E.Funcional AND P.Ano = E.Ano)
	   WHERE NOT EXISTS(SELECT * FROM TB_FUNCIONAL_PTRES F WHERE F.Funcional = E.Funcional AND F.Ano = E.Ano AND F.PTRES = E.PTRES)
GO

INSERT INTO [dbo].[TB_FUNCIONAL_PTRES]
       SELECT E.Funcional, E.Ano, E.PTRES, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '0 - Nenhuma', 0, 0, 0, E.DescricaoPO, '', '', P.SiglaDir, 0, 0, '' 
	   FROM TB_IMP_EXECUCAO_PTRES E INNER JOIN TB_ETAPAS P ON (P.Funcional = E.Funcional AND P.Ano = E.Ano)
	   WHERE NOT EXISTS(SELECT * FROM TB_FUNCIONAL_PTRES F WHERE F.Funcional = E.Funcional AND F.Ano = E.Ano AND F.PTRES = E.PTRES)
GO

UPDATE [dbo].[TB_FUNCIONAL_PTRES]
   SET [VL_Orc_Ini] = E.VL_Orc_Ini
      ,[VL_Orc_Aut] = E.VL_Orc_Aut
      ,[VL_Empenhado] = E.VL_Empenhado
      ,[VL_Liquidado] = E.VL_Liquidado
      ,[VL_Pago] = E.VL_Pago
      ,[VL_Disponivel] = E.VL_Disponivel
      ,[VL_Cred_Sup] = E.VL_Cred_Sup
      ,[VL_Cred_Esp] = E.VL_Cred_Esp
      ,[VL_Cred_Ext] = E.VL_Cred_Ext
      ,[VL_Cred_Can] = E.VL_Cred_Can
      ,[VL_Contido] = E.VL_Contido
	  ,[RP_Pro_Inscrito] = E.RP_Pro_Inscrito
      ,[RP_Pro_Reinscrito] = E.RP_Pro_Reinscrito
      ,[RP_Pro_Cancelado] = E.RP_Pro_Cancelado
      ,[RP_Pro_Pago] = E.RP_Pro_Pago
      ,[RP_Pro_APagar] = E.RP_Pro_APagar
      ,[RP_NPro_Inscrito] = E.RP_NPro_Inscrito
      ,[RP_NPro_Reinscrito] = E.RP_NPro_Reinscrito
      ,[RP_NPro_Cancelado] = E.RP_NPro_Cancelado
      ,[RP_NPro_ALiquidar] = E.RP_NPro_ALiquidar
      ,[RP_NPro_Liquidado] = E.RP_NPro_Liquidado
      ,[RP_NPro_Liq_APagar] = E.RP_NPro_Liq_APagar
      ,[RP_NPro_Pago] = E.RP_NPro_Pago
      ,[RP_NPro_APagar] = E.RP_NPro_APagar
      ,[RP_NPro_Bloqueado] = E.RP_NPro_Bloqueado
	  ,[RP_Inscrito] = (E.RP_Pro_Inscrito + E.RP_Pro_Reinscrito + E.RP_NPro_Inscrito + E.RP_NPro_Reinscrito)
	  ,[RP_Pago] = (E.RP_Pro_Pago + E.RP_NPro_Pago)
	  ,[RP_ExAnt] = E.RP_Pro_Inscrito
	  ,[RP_Liquidado] = E.RP_NPro_Liquidado
	  ,[RP_ALiquidar] = E.RP_NPro_ALiquidar
	  ,[RP_Cancelado] = (E.RP_Pro_Cancelado + E.RP_NPro_Cancelado)
	  ,[RP_Bloqueado] = E.RP_NPro_Bloqueado
	  ,[RPrimario] = E.RPrimario
	  ,[PlanoOrcamentario] = E.PlanoOrcamentario 
	  ,[DescricaoPO] = E.DescricaoPO 
	  ,[NumEmenda] = E.NumEmenda
	  ,[Autor] = E.Autor
	  ,[SiglaDir] = T.SiglaDir
   FROM TB_FUNCIONAL_PTRES P INNER JOIN TB_IMP_EXECUCAO_PTRES E ON (P.Funcional = E.Funcional AND P.Ano = E.Ano AND P.PTRES = E.PTRES) INNER JOIN TB_ETAPAS T ON (T.Funcional = P.Funcional AND T.Ano = P.Ano)
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD] ON [TB_FUNCIONAL_PTRES] FOR UPDATE AS
	IF ((SELECT INSERTED.[Funcional] FROM INSERTED) <> (SELECT DELETED.[Funcional] FROM DELETED) OR (SELECT INSERTED.[Ano] FROM INSERTED) <> (SELECT DELETED.[Ano] FROM DELETED) OR (SELECT INSERTED.[PTRES] FROM INSERTED) <> (SELECT DELETED.[PTRES] FROM DELETED))
	BEGIN
		IF (SELECT COUNT(*) FROM deleted INNER JOIN [TB_CRONOGRAMA] ON deleted.[Funcional] = [TB_CRONOGRAMA].[Funcional] AND deleted.[Ano] = [TB_CRONOGRAMA].[Ano] AND deleted.[PTRES] = [TB_CRONOGRAMA].[PTRES]) > 0
		BEGIN
			SET NOCOUNT ON
			UPDATE [TB_CRONOGRAMA]
			SET [TB_CRONOGRAMA].[Funcional] = (SELECT inserted.[Funcional] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES]),[TB_CRONOGRAMA].[Ano] = (SELECT inserted.[Ano] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES]),[TB_CRONOGRAMA].[PTRES] = (SELECT inserted.[PTRES] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES])
			FROM deleted INNER JOIN [TB_CRONOGRAMA] ON deleted.[Funcional] = [TB_CRONOGRAMA].[Funcional] AND deleted.[Ano] = [TB_CRONOGRAMA].[Ano] AND deleted.[PTRES] = [TB_CRONOGRAMA].[PTRES]
		END
	END
GO

CREATE TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2] ON [TB_CRONOGRAMA] FOR UPDATE AS
	BEGIN
	IF (SELECT COUNT(*) FROM inserted) != (SELECT COUNT(*) FROM [TB_FUNCIONAL_PTRES] INNER JOIN inserted ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES])
		BEGIN
			RAISERROR('TB_FUNCIONAL_PTRES não cadastrado!', 16, 1)
			ROLLBACK TRANSACTION
			RETURN
		END
	END
GO

DELETE FROM [dbo].[TB_IMP_EXECUCAO_PTRES]
   WHERE EXISTS(SELECT * FROM TB_FUNCIONAL_PTRES WHERE TB_FUNCIONAL_PTRES.Funcional = TB_IMP_EXECUCAO_PTRES.Funcional AND  TB_FUNCIONAL_PTRES.Ano = TB_IMP_EXECUCAO_PTRES.Ano AND TB_FUNCIONAL_PTRES.PTRES = TB_IMP_EXECUCAO_PTRES.PTRES)
GO

SELECT * FROM [dbo].[TB_IMP_EXECUCAO_PTRES]
GO

/***************** FINALIZANDO TB_IMP_EXECUCAO_PTRES ************************/

/***************** INICIANDO TB_IMP_EXECUCAO_RP *********************************/

INSERT INTO [dbo].[TB_UG_EXECUTORA]
	SELECT E.[UG_Executora], E.UG_Executora + ' - CLASSIFICAR', 0, 0, 0, '', 0  
	FROM TB_IMP_EXECUCAO_RP E 
	WHERE NOT EXISTS(SELECT U.UG_EXECUTORA FROM TB_UG_EXECUTORA U WHERE E.UG_Executora = U.UG_Executora)
GO

INSERT INTO [dbo].[TB_EXECUCAO_RP]
   SELECT [Funcional], [Ano], [RPrimario], [UG_Executora], [VL_Empenhado], [VL_Liquidado], [VL_Pago], [VL_Disponivel], [VL_Contido] 
   , RP_Pro_Inscrito, RP_Pro_Reinscrito, RP_Pro_Cancelado, RP_Pro_Pago, RP_Pro_APagar, RP_NPro_Inscrito, RP_NPro_Reinscrito, [RP_NPro_Cancelado]
   ,[RP_NPro_ALiquidar], [RP_NPro_Liquidado], [RP_NPro_Liq_APagar], [RP_NPro_Pago], [RP_NPro_APagar], [RP_NPro_Bloqueado], 'Nenhuma' 
   FROM TB_IMP_EXECUCAO_RP R
   WHERE NOT EXISTS(SELECT * FROM TB_EXECUCAO_RP P WHERE P.Funcional = R.Funcional AND P.Ano = R.Ano AND P.RPrimario = R.RPrimario AND P.UG_Executora = R.UG_Executora)
GO

UPDATE [dbo].[TB_EXECUCAO_RP]
   SET [VL_Empenhado] = R.VL_Empenhado
      ,[VL_Liquidado] = R.VL_Liquidado
      ,[VL_Pago] = R.VL_Pago
      ,[VL_Disponivel] = R.VL_Disponivel
      ,[VL_Contido] = R.VL_Contido
	  ,[RP_Pro_Inscrito] = R.RP_Pro_Inscrito 
      ,[RP_Pro_Reinscrito] = R.RP_Pro_Reinscrito 
      ,[RP_Pro_Cancelado] = R.RP_Pro_Cancelado 
      ,[RP_Pro_Pago] = R.RP_Pro_Pago 
      ,[RP_Pro_APagar] = R.RP_Pro_APagar 
      ,[RP_NPro_Inscrito] = R.RP_NPro_Inscrito 
      ,[RP_NPro_Reinscrito] = R.RP_NPro_Reinscrito 
      ,[RP_NPro_Cancelado] = R.RP_NPro_Cancelado 
      ,[RP_NPro_ALiquidar] = R.RP_NPro_ALiquidar 
      ,[RP_NPro_Liquidado] = R.RP_NPro_Liquidado 
      ,[RP_NPro_Liq_APagar] = R.RP_NPro_Liq_APagar 
      ,[RP_NPro_Pago] = R.RP_NPro_Pago 
      ,[RP_NPro_APagar] = R.RP_NPro_APagar 
      ,[RP_NPro_Bloqueado] = R.RP_NPro_Bloqueado
   FROM TB_EXECUCAO_RP P INNER JOIN TB_IMP_EXECUCAO_RP R ON (P.Funcional = R.Funcional AND P.Ano = R.Ano AND P.RPrimario = R.RPrimario AND P.UG_Executora = R.UG_Executora)
GO

DELETE FROM [dbo].[TB_IMP_EXECUCAO_RP]
      WHERE EXISTS(SELECT * FROM TB_EXECUCAO_RP P WHERE P.Funcional = TB_IMP_EXECUCAO_RP.Funcional AND P.Ano = TB_IMP_EXECUCAO_RP.Ano AND P.RPrimario = TB_IMP_EXECUCAO_RP.RPrimario AND P.UG_Executora = TB_IMP_EXECUCAO_RP.UG_Executora)
GO

SELECT * FROM [dbo].[TB_IMP_EXECUCAO_RP]
GO

/***************** FINALIZANDO TB_IMP_EXECUCAO_RP *******************************/

/***************** INICIANDO TB_IMP_NATUREZA_PTRES ***********************/

/*************************************/
/********ALTEREI AQUI ****************/
/*************************************/

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD') AND sysstat & 0xf = 8)
DROP TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD]
GO

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2]') AND sysstat & 0xf = 8)
DROP TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2]
GO

/*** INÍCIO DO PROCESSO DE INSERÇÃO ***/

INSERT INTO [dbo].[TB_FONTE]
       SELECT DISTINCT N.FonteSOF, 'A CLASSIFICAR' + ' ' + N.FonteSOF FROM TB_IMP_NATUREZA_PTRES N
       WHERE NOT EXISTS(SELECT * FROM TB_FONTE WHERE TB_FONTE.FonteSOF = N.FonteSOF)
GO

INSERT INTO [dbo].[TB_NATUREZAS]
       SELECT DISTINCT N.Natureza FROM TB_IMP_NATUREZA_PTRES N
       WHERE NOT EXISTS(SELECT * FROM TB_NATUREZAS WHERE TB_NATUREZAS.Natureza = N.Natureza)
GO

INSERT INTO [dbo].[TB_FUNCIONAL_PTRES]
       SELECT E.Funcional, E.Ano, E.PTRES, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '0 - Nenhuma', 0, 0, 0, '', '', '', P.SiglaDir, 0, 0, '' 
	   FROM TB_IMP_NATUREZA_PTRES E INNER JOIN TB_ETAPAS P ON (P.Funcional = E.Funcional AND P.Ano = E.Ano)
	   WHERE NOT EXISTS(SELECT * FROM TB_FUNCIONAL_PTRES F WHERE F.Funcional = E.Funcional AND F.Ano = E.Ano AND F.PTRES = E.PTRES)
GO

INSERT INTO [dbo].[TB_NATUREZA_PTRES]
     SELECT I.[Funcional], I.[Ano], I.[PTRES], I.[RPrimario], I.[Natureza], I.[FonteSOF], I.[CatEco], I.[GND]
           ,I.[ModDesp], I.[Elemento], I.[VL_Orc_Ini], I.[VL_Cred_Sup], I.[VL_Cred_Ext], I.[VL_Cred_Esp], I.[VL_Orc_Aut]
           ,I.[VL_Cred_Can], I.[VL_Disponivel], I.[VL_Contido], I.[VL_Empenhado], I.[VL_Liquidado], I.[VL_Pago]
           ,I.[RP_Pro_Inscrito], I.[RP_Pro_Reinscrito], I.[RP_Pro_Cancelado], I.[RP_Pro_Pago], I.[RP_Pro_APagar]
           ,I.[RP_NPro_Inscrito], I.[RP_NPro_Reinscrito], I.[RP_NPro_Cancelado], I.[RP_NPro_ALiquidar], I.[RP_NPro_Liquidado]
           ,I.[RP_NPro_Liq_APagar], I.[RP_NPro_Pago], I.[RP_NPro_APagar], I.[RP_NPro_Bloqueado],0,0,0,0,0,0,0
     FROM TB_IMP_NATUREZA_PTRES I
     WHERE NOT EXISTS(SELECT * FROM TB_NATUREZA_PTRES P WHERE P.Funcional = I.Funcional AND P.Ano = I.Ano AND P.PTRES = I.PTRES AND P.RPrimario = I.RPrimario AND P.Natureza = I.Natureza AND P.FonteSOF = I.FonteSOF)
GO

UPDATE [dbo].[TB_NATUREZA_PTRES]
     SET [VL_Orc_Ini] = I.[VL_Orc_Ini]
        ,[VL_Cred_Sup] = I.[VL_Cred_Sup]
        ,[VL_Cred_Ext] = I.[VL_Cred_Ext]
        ,[VL_Cred_Esp] = I.[VL_Cred_Esp]
        ,[VL_Orc_Aut] = I.[VL_Orc_Aut]
        ,[VL_Cred_Can] = I.[VL_Cred_Can]
        ,[VL_Disponivel] = I.[VL_Disponivel]
        ,[VL_Contido] = I.[VL_Contido]
        ,[VL_Empenhado] = I.[VL_Empenhado]
        ,[VL_Liquidado] = I.[VL_Liquidado]
        ,[VL_Pago] = I.[VL_Pago]
        ,[RP_Pro_Inscrito] = I.[RP_Pro_Inscrito]
        ,[RP_Pro_Reinscrito] = I.[RP_Pro_Reinscrito]
        ,[RP_Pro_Cancelado] = I.[RP_Pro_Cancelado]
        ,[RP_Pro_Pago] = I.[RP_Pro_Pago]
        ,[RP_Pro_APagar] = I.[RP_Pro_APagar]
        ,[RP_NPro_Inscrito] = I.[RP_NPro_Inscrito]
        ,[RP_NPro_Reinscrito] = I.[RP_NPro_Reinscrito]
        ,[RP_NPro_Cancelado] = I.[RP_NPro_Cancelado]
        ,[RP_NPro_ALiquidar] = I.[RP_NPro_ALiquidar]
        ,[RP_NPro_Liquidado] = I.[RP_NPro_Liquidado]
        ,[RP_NPro_Liq_APagar] = I.[RP_NPro_Liq_APagar]
        ,[RP_NPro_Pago] = I.[RP_NPro_Pago]
        ,[RP_NPro_APagar] = I.[RP_NPro_APagar]
        ,[RP_NPro_Bloqueado] = I.[RP_NPro_Bloqueado]
     FROM TB_NATUREZA_PTRES N INNER JOIN TB_IMP_NATUREZA_PTRES I ON (N.Funcional = I.Funcional AND N.Ano = I.Ano AND N.PTRES = I.PTRES AND N.RPrimario = I.RPrimario AND N.Natureza = I.Natureza AND N.FonteSOF = I.FonteSOF)
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD] ON [TB_FUNCIONAL_PTRES] FOR UPDATE AS
	IF ((SELECT INSERTED.[Funcional] FROM INSERTED) <> (SELECT DELETED.[Funcional] FROM DELETED) OR (SELECT INSERTED.[Ano] FROM INSERTED) <> (SELECT DELETED.[Ano] FROM DELETED) OR (SELECT INSERTED.[PTRES] FROM INSERTED) <> (SELECT DELETED.[PTRES] FROM DELETED))
	BEGIN
		IF (SELECT COUNT(*) FROM deleted INNER JOIN [TB_CRONOGRAMA] ON deleted.[Funcional] = [TB_CRONOGRAMA].[Funcional] AND deleted.[Ano] = [TB_CRONOGRAMA].[Ano] AND deleted.[PTRES] = [TB_CRONOGRAMA].[PTRES]) > 0
		BEGIN
			SET NOCOUNT ON
			UPDATE [TB_CRONOGRAMA]
			SET [TB_CRONOGRAMA].[Funcional] = (SELECT inserted.[Funcional] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES]),[TB_CRONOGRAMA].[Ano] = (SELECT inserted.[Ano] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES]),[TB_CRONOGRAMA].[PTRES] = (SELECT inserted.[PTRES] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES])
			FROM deleted INNER JOIN [TB_CRONOGRAMA] ON deleted.[Funcional] = [TB_CRONOGRAMA].[Funcional] AND deleted.[Ano] = [TB_CRONOGRAMA].[Ano] AND deleted.[PTRES] = [TB_CRONOGRAMA].[PTRES]
		END
	END
GO

CREATE TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2] ON [TB_CRONOGRAMA] FOR UPDATE AS
	BEGIN
	IF (SELECT COUNT(*) FROM inserted) != (SELECT COUNT(*) FROM [TB_FUNCIONAL_PTRES] INNER JOIN inserted ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES])
		BEGIN
			RAISERROR('TB_FUNCIONAL_PTRES não cadastrado!', 16, 1)
			ROLLBACK TRANSACTION
			RETURN
		END
	END
GO

DELETE FROM [dbo].[TB_IMP_NATUREZA_PTRES] 
   WHERE EXISTS(SELECT * FROM TB_NATUREZA_PTRES N WHERE N.Funcional = TB_IMP_NATUREZA_PTRES.Funcional AND N.Ano = TB_IMP_NATUREZA_PTRES.Ano AND N.PTRES = TB_IMP_NATUREZA_PTRES.PTRES AND N.RPrimario = TB_IMP_NATUREZA_PTRES.RPrimario AND N.Natureza = TB_IMP_NATUREZA_PTRES.Natureza AND N.FonteSOF = TB_IMP_NATUREZA_PTRES.FonteSOF)
GO

SELECT * FROM [dbo].[TB_IMP_NATUREZA_PTRES]
GO

/***************** FINALIZANDO TB_IMP_NATUREZA_PTRES *********************/

/***************** INICIANDO TB_IMP_DESTAQUE **********************/

UPDATE [dbo].[TB_ETAPAS]
   SET [VL_Destacado] = (D.VL_Provisao + D.VL_Destacado) 
      ,[VL_DestaqueRecebido] = (D.Rec_Destacado + D.Rec_Provisao)
	  
   FROM TB_ETAPAS E INNER JOIN TB_IMP_DESTAQUE D ON (E.Funcional = D.Funcional AND E.Ano = D.Ano)					     
GO

DELETE FROM [dbo].[TB_IMP_DESTAQUE] 
   WHERE EXISTS(SELECT * FROM TB_ETAPAS WHERE TB_ETAPAS.Funcional = TB_IMP_DESTAQUE.Funcional AND TB_ETAPAS.Ano = TB_IMP_DESTAQUE.Ano)
GO

SELECT * FROM [dbo].[TB_IMP_DESTAQUE]
GO

/***************** FINALIZANDO TB_IMP_DESTAQUE *******************/

/***************** INICIANDO TB_IMP_DESTAQUE_PTRES ***************/

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD') AND sysstat & 0xf = 8)
DROP TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD]
GO

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2') AND sysstat & 0xf = 8)
DROP TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2]
GO

UPDATE [dbo].[TB_FUNCIONAL_PTRES] 
   SET [VL_Destacado] = (D.VL_Provisao + D.VL_Destacado) 
	  ,[VL_DestaqueRecebido] = (D.Rec_Destacado + D.Rec_Provisao)
   FROM TB_FUNCIONAL_PTRES E INNER JOIN TB_IMP_DESTAQUE_PTRES D ON (E.Funcional = D.Funcional AND E.Ano = D.Ano AND E.PTRES = D.PTRES)					     
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD] ON [TB_FUNCIONAL_PTRES] FOR UPDATE AS
	IF ((SELECT INSERTED.[Funcional] FROM INSERTED) <> (SELECT DELETED.[Funcional] FROM DELETED) OR (SELECT INSERTED.[Ano] FROM INSERTED) <> (SELECT DELETED.[Ano] FROM DELETED) OR (SELECT INSERTED.[PTRES] FROM INSERTED) <> (SELECT DELETED.[PTRES] FROM DELETED))
	BEGIN
		IF (SELECT COUNT(*) FROM deleted INNER JOIN [TB_CRONOGRAMA] ON deleted.[Funcional] = [TB_CRONOGRAMA].[Funcional] AND deleted.[Ano] = [TB_CRONOGRAMA].[Ano] AND deleted.[PTRES] = [TB_CRONOGRAMA].[PTRES]) > 0
		BEGIN
			SET NOCOUNT ON
			UPDATE [TB_CRONOGRAMA]
			SET [TB_CRONOGRAMA].[Funcional] = (SELECT inserted.[Funcional] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES]),[TB_CRONOGRAMA].[Ano] = (SELECT inserted.[Ano] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES]),[TB_CRONOGRAMA].[PTRES] = (SELECT inserted.[PTRES] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES])
			FROM deleted INNER JOIN [TB_CRONOGRAMA] ON deleted.[Funcional] = [TB_CRONOGRAMA].[Funcional] AND deleted.[Ano] = [TB_CRONOGRAMA].[Ano] AND deleted.[PTRES] = [TB_CRONOGRAMA].[PTRES]
		END
	END
GO

CREATE TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2] ON [TB_CRONOGRAMA] FOR UPDATE AS
	BEGIN
	IF (SELECT COUNT(*) FROM inserted) != (SELECT COUNT(*) FROM [TB_FUNCIONAL_PTRES] INNER JOIN inserted ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES])
		BEGIN
			RAISERROR('TB_FUNCIONAL_PTRES não cadastrado!', 16, 1)
			ROLLBACK TRANSACTION
			RETURN
		END
	END
GO

DELETE FROM [dbo].[TB_IMP_DESTAQUE_PTRES]  
   WHERE EXISTS(SELECT * FROM TB_FUNCIONAL_PTRES F WHERE F.Funcional = TB_IMP_DESTAQUE_PTRES.Funcional AND F.Ano = TB_IMP_DESTAQUE_PTRES.Ano AND F.PTRES = TB_IMP_DESTAQUE_PTRES.PTRES)
GO

SELECT * FROM [dbo].[TB_IMP_DESTAQUE_PTRES]
GO

/***************** FINALIZANDO TB_IMP_DESTAQUE_PTRES *************/

/***************** INICIANDO TB_IMP_BLOQUEIO_NATUREZA ***************/

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('VW_TEMP_BLOQUEIO') AND sysstat & 0xf = 2)
DROP VIEW [VW_TEMP_BLOQUEIO]
GO

CREATE VIEW [VW_TEMP_BLOQUEIO] AS(
       SELECT B.Funcional,B.Ano,SUM(B.VL_Cred_Bloq_Rema) AS VL_Cred_Bloq_Rema,SUM(B.VL_Cred_Bloq_Contido) AS VL_Cred_Bloq_Contido 
			,SUM(B.VL_Cred_Bloq_SOF) AS VL_Cred_Bloq_SOF,SUM(B.VL_Cred_Bloq_RemaSOF) AS VL_Cred_Bloq_RemaSOF,SUM(B.VL_Cred_Bloq_Controle) AS VL_Cred_Bloq_Controle
			,SUM(B.VL_Cred_Bloq_RP) AS VL_Cred_Bloq_RP,SUM(B.VL_Cred_Bloq_PreEmp) AS VL_Cred_Bloq_PreEmp
	   FROM TB_IMP_BLOQUEIO_NATUREZA B
       GROUP BY B.Funcional,B.Ano
)
GO

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('VW_TEMP_BLOQUEIO_PTRES') AND sysstat & 0xf = 2)
DROP VIEW [VW_TEMP_BLOQUEIO_PTRES]
GO

CREATE VIEW [VW_TEMP_BLOQUEIO_PTRES]AS(
       SELECT B.Funcional,B.Ano,B.PTRES,SUM(B.VL_Cred_Bloq_Rema) AS VL_Cred_Bloq_Rema,SUM(B.VL_Cred_Bloq_Contido) AS VL_Cred_Bloq_Contido 
			,SUM(B.VL_Cred_Bloq_SOF) AS VL_Cred_Bloq_SOF,SUM(B.VL_Cred_Bloq_RemaSOF) AS VL_Cred_Bloq_RemaSOF,SUM(B.VL_Cred_Bloq_Controle) AS VL_Cred_Bloq_Controle
			,SUM(B.VL_Cred_Bloq_RP) AS VL_Cred_Bloq_RP,SUM(B.VL_Cred_Bloq_PreEmp) AS VL_Cred_Bloq_PreEmp
	   FROM TB_IMP_BLOQUEIO_NATUREZA B
       GROUP BY B.Funcional,B.Ano,B.PTRES
)
GO

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD') AND sysstat & 0xf = 8)
DROP TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD]
GO

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2') AND sysstat & 0xf = 8)
DROP TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2]
GO

UPDATE [dbo].[TB_ETAPAS] 
   SET [VL_Cred_Bloq_Rema] = 0 
      ,[VL_Cred_Bloq_SOF] = 0 
      ,[VL_Cred_Bloq_RemaSOF] = 0 
      ,[VL_Cred_Bloq_RP] = 0 
	  ,[VL_Cred_Bloq_Contido] = 0
	  ,[VL_Cred_Bloq_PreEmp] = 0
	  ,[VL_Cred_Bloq_Controle] = 0
    WHERE Ano = YEAR(GETDATE())					     
GO

UPDATE [dbo].[TB_FUNCIONAL_PTRES] 
   SET [VL_Cred_Bloq_Rema] = 0 
      ,[VL_Cred_Bloq_SOF] = 0 
      ,[VL_Cred_Bloq_RemaSOF] = 0 
      ,[VL_Cred_Bloq_RP] = 0 
	  ,[VL_Cred_Bloq_Contido] = 0
	  ,[VL_Cred_Bloq_PreEmp] = 0
	  ,[VL_Cred_Bloq_Controle] = 0
    WHERE Ano = YEAR(GETDATE())					     
GO

UPDATE [dbo].[TB_NATUREZA_PTRES] 
   SET [VL_Cred_Bloq_Rema] = 0 
      ,[VL_Cred_Bloq_SOF] = 0 
      ,[VL_Cred_Bloq_RemaSOF] = 0 
      ,[VL_Cred_Bloq_RP] = 0 
	  ,[VL_Cred_Bloq_Contido] = 0
	  ,[VL_Cred_Bloq_PreEmp] = 0
	  ,[VL_Cred_Bloq_Controle] = 0
    WHERE Ano = YEAR(GETDATE())					     
GO

UPDATE [dbo].[TB_ETAPAS] 
   SET [VL_Cred_Bloq_Rema] = C.VL_Cred_Bloq_Rema 
      ,[VL_Cred_Bloq_SOF] = C.VL_Cred_Bloq_SOF 
      ,[VL_Cred_Bloq_RemaSOF] = C.VL_Cred_Bloq_RemaSOF 
      ,[VL_Cred_Bloq_RP] = C.VL_Cred_Bloq_RP 
	  ,[VL_Cred_Bloq_Contido] = C.VL_Cred_Bloq_Contido
	  ,[VL_Cred_Bloq_PreEmp] = C.VL_Cred_Bloq_PreEmp
	  ,[VL_Cred_Bloq_Controle] = C.VL_Cred_Bloq_Controle
   FROM TB_ETAPAS E INNER JOIN VW_TEMP_BLOQUEIO C ON (E.Funcional = C.Funcional AND E.Ano = C.Ano)					     
GO

UPDATE [dbo].[TB_FUNCIONAL_PTRES] 
   SET [VL_Cred_Bloq_Rema] = C.VL_Cred_Bloq_Rema 
      ,[VL_Cred_Bloq_SOF] = C.VL_Cred_Bloq_SOF 
      ,[VL_Cred_Bloq_RemaSOF] = C.VL_Cred_Bloq_RemaSOF 
      ,[VL_Cred_Bloq_RP] = C.VL_Cred_Bloq_RP 
	  ,[VL_Cred_Bloq_Contido] = C.VL_Cred_Bloq_Contido
	  ,[VL_Cred_Bloq_PreEmp] = C.VL_Cred_Bloq_PreEmp
	  ,[VL_Cred_Bloq_Controle] = C.VL_Cred_Bloq_Controle
   FROM TB_FUNCIONAL_PTRES E INNER JOIN VW_TEMP_BLOQUEIO_PTRES C ON (E.Funcional = C.Funcional AND E.Ano = C.Ano AND E.PTRES = C.PTRES)					     
GO

UPDATE [dbo].[TB_NATUREZA_PTRES] 
   SET [VL_Cred_Bloq_Rema] = C.VL_Cred_Bloq_Rema 
      ,[VL_Cred_Bloq_SOF] = C.VL_Cred_Bloq_SOF 
      ,[VL_Cred_Bloq_RemaSOF] = C.VL_Cred_Bloq_RemaSOF 
      ,[VL_Cred_Bloq_RP] = C.VL_Cred_Bloq_RP 
	  ,[VL_Cred_Bloq_Contido] = C.VL_Cred_Bloq_Contido
	  ,[VL_Cred_Bloq_PreEmp] = C.VL_Cred_Bloq_PreEmp
	  ,[VL_Cred_Bloq_Controle] = C.VL_Cred_Bloq_Controle
   FROM TB_NATUREZA_PTRES E INNER JOIN TB_IMP_BLOQUEIO_NATUREZA C ON (E.Funcional = C.Funcional AND E.Ano = C.Ano AND E.PTRES = C.PTRES AND E.RPrimario = C.RPrimario AND E.Natureza = C.Natureza AND E.FonteSOF = C.FonteSOF)					     
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD] ON [TB_FUNCIONAL_PTRES] FOR UPDATE AS
	IF ((SELECT INSERTED.[Funcional] FROM INSERTED) <> (SELECT DELETED.[Funcional] FROM DELETED) OR (SELECT INSERTED.[Ano] FROM INSERTED) <> (SELECT DELETED.[Ano] FROM DELETED) OR (SELECT INSERTED.[PTRES] FROM INSERTED) <> (SELECT DELETED.[PTRES] FROM DELETED))
	BEGIN
		IF (SELECT COUNT(*) FROM deleted INNER JOIN [TB_CRONOGRAMA] ON deleted.[Funcional] = [TB_CRONOGRAMA].[Funcional] AND deleted.[Ano] = [TB_CRONOGRAMA].[Ano] AND deleted.[PTRES] = [TB_CRONOGRAMA].[PTRES]) > 0
		BEGIN
			SET NOCOUNT ON
			UPDATE [TB_CRONOGRAMA]
			SET [TB_CRONOGRAMA].[Funcional] = (SELECT inserted.[Funcional] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES]),[TB_CRONOGRAMA].[Ano] = (SELECT inserted.[Ano] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES]),[TB_CRONOGRAMA].[PTRES] = (SELECT inserted.[PTRES] FROM INSERTED INNER JOIN [TB_FUNCIONAL_PTRES] ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES])
			FROM deleted INNER JOIN [TB_CRONOGRAMA] ON deleted.[Funcional] = [TB_CRONOGRAMA].[Funcional] AND deleted.[Ano] = [TB_CRONOGRAMA].[Ano] AND deleted.[PTRES] = [TB_CRONOGRAMA].[PTRES]
		END
	END
GO

CREATE TRIGGER [FK_TB_CRONOGRAMA_TB_FUCIONAL_PTRES_UPD2] ON [TB_CRONOGRAMA] FOR UPDATE AS
	BEGIN
	IF (SELECT COUNT(*) FROM inserted) != (SELECT COUNT(*) FROM [TB_FUNCIONAL_PTRES] INNER JOIN inserted ON inserted.[Funcional] = [TB_FUNCIONAL_PTRES].[Funcional] AND inserted.[Ano] = [TB_FUNCIONAL_PTRES].[Ano] AND inserted.[PTRES] = [TB_FUNCIONAL_PTRES].[PTRES])
		BEGIN
			RAISERROR('TB_FUNCIONAL_PTRES não cadastrado!', 16, 1)
			ROLLBACK TRANSACTION
			RETURN
		END
	END
GO

DELETE FROM [dbo].[TB_IMP_BLOQUEIO_NATUREZA]
   WHERE EXISTS(SELECT * FROM TB_NATUREZA_PTRES C WHERE TB_IMP_BLOQUEIO_NATUREZA.Funcional = C.Funcional AND TB_IMP_BLOQUEIO_NATUREZA.Ano = C.Ano AND TB_IMP_BLOQUEIO_NATUREZA.PTRES = C.PTRES AND TB_IMP_BLOQUEIO_NATUREZA.RPrimario = C.RPrimario AND TB_IMP_BLOQUEIO_NATUREZA.Natureza = C.Natureza AND TB_IMP_BLOQUEIO_NATUREZA.FonteSOF = C.FonteSOF)
GO

SELECT * FROM [dbo].[TB_IMP_BLOQUEIO_NATUREZA]
GO

/***************** FINALIZANDO TB_IMP_BLOQUEIO_NATUREZA *************/

/***************** INICIANDO TB_IMP_HISTORICO_PGTO **************************/

/*---------------------------------------------------------------*/
/*                        Criação de VIEW                        */
/*                            VW_TEMP_HISTORICO                  */
/*---------------------------------------------------------------*/


IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('VW_TEMP_HISTORICO') AND sysstat & 0xf = 2)
DROP VIEW [VW_TEMP_HISTORICO]
GO

CREATE VIEW [VW_TEMP_HISTORICO]AS(
       SELECT H.ContaCorrente, H.Ano, H.MesPgto, SUM(H.VL_Pago) AS VL_Pago, SUM(H.RP_Pago) AS RP_Pago
	   FROM TB_IMP_HISTORICO_PGTO H
	   WHERE H.MesPgto > 0
	   GROUP BY H.ContaCorrente, H.Ano, H.MesPgto
)
GO

INSERT INTO [dbo].[TB_HISTORICO_PGTO]
       SELECT H.ContaCorrente, H.Ano, H.MesPgto, H.VL_Pago, H.RP_Pago
	   FROM VW_TEMP_HISTORICO H 
	   WHERE NOT EXISTS(SELECT * FROM TB_HISTORICO_PGTO G WHERE G.ContaCorrente = H.ContaCorrente AND G.Ano = H.Ano AND G.MesPgto = H.MesPgto) AND 
	             EXISTS(SELECT * FROM TB_IMP_EMPENHO E WHERE E.ContaCorrente = H.ContaCorrente)
GO

UPDATE [dbo].[TB_HISTORICO_PGTO]
   SET [VL_Pago] = H.VL_Pago,
       [RP_Pago] = H.RP_Pago
   FROM TB_HISTORICO_PGTO P INNER JOIN VW_TEMP_HISTORICO H ON (H.ContaCorrente = P.ContaCorrente AND H.Ano = P.Ano AND H.MesPgto = P.MesPgto) 
GO

DELETE FROM [dbo].[TB_IMP_HISTORICO_PGTO]
   WHERE EXISTS(SELECT * FROM TB_HISTORICO_PGTO WHERE TB_HISTORICO_PGTO.ContaCorrente = TB_IMP_HISTORICO_PGTO.ContaCorrente AND TB_HISTORICO_PGTO.Ano = TB_IMP_HISTORICO_PGTO.Ano AND TB_HISTORICO_PGTO.MesPgto = TB_IMP_HISTORICO_PGTO.MesPgto)
GO

DELETE FROM [dbo].[TB_IMP_HISTORICO_PGTO]
   WHERE VL_Pago = 0 AND RP_Pago = 0
GO

SELECT * FROM [dbo].[TB_IMP_HISTORICO_PGTO]
GO

/***************** FINALIZANDO TB_IMP_HISTORICO_PGTO ************************/

/***************** INICIANDO TB_IMP_PGTO *********************************/

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('VW_TEMP_PGTO') AND sysstat & 0xf = 2)
DROP VIEW [VW_TEMP_PGTO]
GO

CREATE VIEW [VW_TEMP_PGTO]AS(
       SELECT P.Funcional, P.Ano, P.Mes, P.RPrimario, P.PlanoOrcamentario, SUM(P.VL_Pago) AS VL_Pago 
       FROM TB_IMP_PGTO P
       GROUP BY P.Funcional, P.Ano, P.Mes, P.RPrimario, P.PlanoOrcamentario
)
GO

INSERT INTO [dbo].[TB_PGTO]
       SELECT P.Funcional, P.Ano, P.Mes, P.RPrimario, P.PlanoOrcamentario, P.VL_Pago 
	   FROM VW_TEMP_PGTO P WHERE NOT EXISTS(SELECT * FROM TB_PGTO G WHERE G.Funcional  = P.Funcional AND G.Ano = P.Ano AND G.Mes = P.Mes AND G.PlanoOrcamentario = P.PlanoOrcamentario)
GO

UPDATE [dbo].[TB_PGTO]
   SET [VL_Pago] = P.VL_Pago
   FROM TB_PGTO G INNER JOIN VW_TEMP_PGTO P ON (G.Funcional  = P.Funcional AND G.Ano = P.Ano AND G.Mes = P.Mes AND G.PlanoOrcamentario = P.PlanoOrcamentario)
GO

DELETE FROM [dbo].[TB_IMP_PGTO]
      WHERE EXISTS(SELECT * FROM TB_PGTO WHERE TB_PGTO.Funcional = TB_IMP_PGTO.Funcional AND TB_PGTO.Ano = TB_IMP_PGTO.Ano AND TB_PGTO.Mes = TB_IMP_PGTO.Mes AND TB_PGTO.PlanoOrcamentario = TB_IMP_PGTO.PlanoOrcamentario)
GO

SELECT * FROM [dbo].[TB_IMP_PGTO]
GO

/***************** FINALIZANDO TB_IMP_PGTO *******************************/

/***************** INICIANDO TB_IMP_LIMITE *******************************/

DELETE FROM [dbo].[TB_IMP_LIMITE]
      WHERE VL_Indisponivel = 0
GO

/***************** FINALIZANDO TB_IMP_LIMITE ****************************/

/***************** INICIANDO TB_IMP_SRE_ADM *********************************/

INSERT INTO [dbo].[TB_SRE_ADM]
       SELECT P.Funcional, P.Ano, P.PTRES, P.CodMT, P.UG_Executora, P.Mes, P.VL_Empenhado, P.VL_Liquidado 
	   FROM TB_IMP_SRE_ADM P WHERE NOT EXISTS(SELECT * FROM TB_SRE_ADM G WHERE G.Funcional  = P.Funcional AND G.Ano = P.Ano AND G.PTRES = P.PTRES AND G.CodMT = P.CodMT AND G.UG_Executora = P.UG_Executora AND G.Mes = P.Mes)
GO

UPDATE [dbo].[TB_SRE_ADM]
   SET [VL_Empenhado] = P.VL_Empenhado
      ,[VL_Liquidado] = P.VL_Liquidado  
   FROM TB_SRE_ADM G INNER JOIN TB_IMP_SRE_ADM P ON (G.Funcional  = P.Funcional AND G.Ano = P.Ano AND G.PTRES = P.PTRES AND G.CodMT = P.CodMT AND G.UG_Executora = P.UG_Executora AND G.Mes = P.Mes)
GO

DELETE FROM [dbo].[TB_IMP_SRE_ADM]
      WHERE EXISTS(SELECT * FROM TB_SRE_ADM WHERE TB_SRE_ADM.Funcional = TB_IMP_SRE_ADM.Funcional AND TB_SRE_ADM.Ano = TB_IMP_SRE_ADM.Ano AND TB_SRE_ADM.PTRES = TB_IMP_SRE_ADM.PTRES AND TB_SRE_ADM.CodMT = TB_IMP_SRE_ADM.CodMT AND TB_SRE_ADM.UG_Executora = TB_IMP_SRE_ADM.UG_Executora AND TB_SRE_ADM.Mes = TB_IMP_SRE_ADM.Mes)
GO

SELECT * FROM [dbo].[TB_IMP_SRE_ADM]
GO

/***************** FINALIZANDO TB_IMP_SRE_ADM *******************************/

/***************** INICIANDO TB_IMP_RAP_EMPENHO *********************************/

DELETE FROM [dbo].[TB_IMP_RAP_EMPENHO]
       WHERE RP_Inscrito = 0 AND RP_ExAnt = 0 AND RP_Liquidado = 0 AND RP_Cancelado = 0 AND RP_Pago = 0 AND RP_Saldo = 0 AND RP_Bloqueado = 0
GO

IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('VW_RAP_NE') AND sysstat & 0xf = 2)
DROP VIEW [VW_RAP_NE]
GO

CREATE VIEW [VW_RAP_NE]AS(SELECT ContaCorrente, Ano, IIF(LEFT(Mes, 3) = '000', ('JAN' + RIGHT(Mes, 5)), Mes) AS Mes, 
  SUM(RP_ExAnt) AS RP_ExAnt, SUM(RP_Liquidado) AS RP_Liquidado, SUM(RP_Inscrito) AS RP_Inscrito, SUM(RP_Cancelado) AS RP_Cancelado, 
  SUM(RP_Pago) AS RP_Pago, SUM(RP_Saldo) AS RP_Saldo, SUM(RP_Bloqueado) AS RP_Bloqueado  
  FROM TB_IMP_RAP_EMPENHO GROUP BY ContaCorrente, Ano, IIF(LEFT(Mes, 3) = '000', ('JAN' + RIGHT(Mes, 5)), Mes))
GO

INSERT INTO [dbo].[TB_RAP_EMPENHO]
	SELECT V.ContaCorrente, V.Ano, V.Mes, V.RP_Inscrito, V.RP_Liquidado, V.RP_Pago, V.RP_Cancelado, V.RP_Bloqueado, V.RP_Saldo 
	FROM [DBPLOAWEB].[dbo].[VW_RAP_NE] V  
    WHERE NOT EXISTS (SELECT * FROM [DBPLOAWEB].[dbo].[TB_RAP_EMPENHO] P WHERE P.[ContaCorrente] = V.[ContaCorrente] AND P.[Mes] = V.[Mes])     
GO

UPDATE [dbo].[TB_RAP_EMPENHO]
   SET [RP_Inscrito] = R.[RP_Inscrito]
      ,[RP_Liquidado] = (R.[RP_ExAnt] + R.[RP_Liquidado])
      ,[RP_Pago] = R.[RP_Pago]
      ,[RP_Cancelado] = R.[RP_Cancelado]
	  ,[RP_Bloqueado] = R.[RP_Bloqueado]
	  ,[RP_Saldo] = R.[RP_Saldo]
   FROM TB_RAP_EMPENHO E INNER JOIN VW_RAP_NE R ON (E.ContaCorrente = R.ContaCorrente AND E.Mes = R.Mes) 
GO

DELETE FROM [dbo].[TB_IMP_RAP_EMPENHO]
       WHERE (EXISTS(SELECT * FROM [dbo].[TB_RAP_EMPENHO] E WHERE E.ContaCorrente = [dbo].[TB_IMP_RAP_EMPENHO].[ContaCorrente] AND E.Mes = [dbo].[TB_IMP_RAP_EMPENHO].[Mes]) OR
	         Mes LIKE '000/%')
GO

SELECT * FROM [dbo].[TB_IMP_RAP_EMPENHO]
GO

/***************** FINALIZANDO TB_IMP_RAP_EMPENHO *******************************/

/***************** INICIANDO TB_IMP_NATUREZADETALHADA ***************************/

INSERT INTO [dbo].[TB_UG_EXECUTORA]
       SELECT DISTINCT I.UG_Executora, I.DetalhamentoNatureza, 0, 0, 0, '', 0 FROM TB_IMP_NATUREZADETALHADA I
       WHERE NOT EXISTS (SELECT * FROM TB_UG_EXECUTORA U WHERE U.UG_Executora = I.UG_Executora)
GO

INSERT INTO [dbo].[TB_NATUREZADETALHADA]
	SELECT DISTINCT '00', I.NaturezaDetalhada, I.DetalhamentoNatureza
	FROM TB_IMP_NATUREZADETALHADA I
	WHERE NOT EXISTS (SELECT * FROM TB_NATUREZADETALHADA N WHERE N.NaturezaDetalhada = I.NaturezaDetalhada)
GO

INSERT INTO [dbo].[TB_MOV_NATUREZA]
	SELECT I.Funcional, I.Ano, I.PTRES, I.PlanoOrcamentario, I.DescricaoPO, I.UG_Executora, I.NaturezaDetalhada, I.CodMT,
           I.VL_Empenhado, I.VL_Liquidado, I.VL_Pago, I.RP_Pro_Inscrito, I.RP_Pro_Reinscrito, I.RP_Pro_Cancelado,
           I.RP_Pro_Pago, I.RP_NPro_Inscrito, I.RP_NPro_Reinscrito, I.RP_NPro_Cancelado, I.RP_NPro_Liquidado,
		   I.RP_NPro_Pago, I.RP_NPro_Bloqueado
	FROM TB_IMP_NATUREZADETALHADA I
	WHERE NOT EXISTS(SELECT * FROM TB_MOV_NATUREZA N WHERE N.Funcional = I.Funcional AND N.Ano = I.Ano AND N.PTRES = I.PTRES AND N.UG_Executora = I.UG_Executora AND N.NaturezaDetalhada = I.NaturezaDetalhada AND N.CodMT = I.CodMT)     
GO

UPDATE [dbo].[TB_MOV_NATUREZA]
   SET [VL_Empenhado] = I.VL_Empenhado
      ,[VL_Liquidado] = I.VL_Liquidado
      ,[VL_Pago] = I.VL_Pago
      ,[RP_Pro_Inscrito] = I.RP_Pro_Inscrito
      ,[RP_Pro_Reinscrito] = I.RP_Pro_Reinscrito
      ,[RP_Pro_Cancelado] = I.RP_Pro_Cancelado
      ,[RP_Pro_Pago] = I.RP_Pro_Pago
      ,[RP_NPro_Inscrito] = I.RP_NPro_Inscrito
      ,[RP_NPro_Reinscrito] = I.RP_NPro_Reinscrito
      ,[RP_NPro_Cancelado] = I.RP_NPro_Cancelado
      ,[RP_NPro_Liquidado] = I.RP_NPro_Liquidado
      ,[RP_NPro_Pago] = I.RP_NPro_Pago
      ,[RP_NPro_Bloqueado] = I.RP_NPro_Bloqueado
	FROM TB_MOV_NATUREZA N INNER JOIN TB_IMP_NATUREZADETALHADA I ON (N.Funcional = I.Funcional AND N.Ano = I.Ano AND N.PTRES = I.PTRES AND N.UG_Executora = I.UG_Executora AND N.NaturezaDetalhada = I.NaturezaDetalhada AND N.CodMT = I.CodMT)
GO

DELETE FROM [dbo].[TB_IMP_NATUREZADETALHADA]
	WHERE EXISTS(SELECT * FROM TB_MOV_NATUREZA N WHERE N.Funcional = TB_IMP_NATUREZADETALHADA.Funcional AND N.Ano = TB_IMP_NATUREZADETALHADA.Ano AND N.PTRES = TB_IMP_NATUREZADETALHADA.PTRES AND N.UG_Executora = TB_IMP_NATUREZADETALHADA.UG_Executora AND N.NaturezaDetalhada = TB_IMP_NATUREZADETALHADA.NaturezaDetalhada AND N.CodMT = TB_IMP_NATUREZADETALHADA.CodMT)           
GO

SELECT * FROM [dbo].[TB_IMP_NATUREZADETALHADA]
GO

/***************** FINALIZANDO TB_IMP_NATUREZADETALHADA **************************/

/***************** INICIANDO TB_IMP_DISPONIVELSRE ****************************/

INSERT INTO [dbo].[TB_UG_EXECUTORA]
       SELECT DISTINCT D.UG_Executora, ('A CLASSIFICAR ' + D.UG_Executora) AS DescricaoUG, 0, 0, 0, '', 0 FROM TB_IMP_DISPONIVELSRE D
       WHERE NOT EXISTS (SELECT * FROM TB_UG_EXECUTORA U WHERE U.UG_Executora = D.UG_Executora)
GO

INSERT INTO [dbo].[TB_DISPONIVELSRE]
       SELECT I.Funcional, I.Ano, I.PTRES, I.UG_Executora, I.RPrimario, I.GND, I.PlanoOrcamentario,
              I.DescricaoPO, I.CodMT, I.VL_Provisao_Recebida, I.VL_Provisao_Concedida, I.VL_Destacado,
              I.VL_Disponivel, I.VL_Contido
	   FROM TB_IMP_DISPONIVELSRE I
	   WHERE NOT EXISTS(SELECT * FROM TB_DISPONIVELSRE D WHERE D.Funcional = I.Funcional AND D.Ano = I.Ano AND D.PTRES = I.PTRES AND D.UG_Executora = I.UG_Executora AND D.GND = I.GND AND D.CodMT = I.CodMT)
     
GO

UPDATE [dbo].[TB_DISPONIVELSRE]
   SET [VL_Provisao_Recebida] = I.VL_Provisao_Recebida
      ,[VL_Provisao_Concedida] = I.VL_Provisao_Concedida
      ,[VL_Destacado] = I.VL_Destacado
      ,[VL_Disponivel] = I.VL_Disponivel
      ,[VL_Contido] = I.VL_Contido
   FROM TB_DISPONIVELSRE D INNER JOIN TB_IMP_DISPONIVELSRE I ON (D.Funcional = I.Funcional AND D.Ano = I.Ano AND D.PTRES = I.PTRES AND D.UG_Executora = I.UG_Executora AND D.GND = I.GND AND D.CodMT = I.CodMT) 
GO

DELETE FROM [dbo].[TB_IMP_DISPONIVELSRE]
	WHERE EXISTS(SELECT * FROM TB_DISPONIVELSRE D WHERE D.Funcional = TB_IMP_DISPONIVELSRE.Funcional AND D.Ano = TB_IMP_DISPONIVELSRE.Ano AND D.PTRES = TB_IMP_DISPONIVELSRE.PTRES AND D.UG_Executora = TB_IMP_DISPONIVELSRE.UG_Executora AND D.GND = TB_IMP_DISPONIVELSRE.GND AND D.CodMT = TB_IMP_DISPONIVELSRE.CodMT)           
GO

SELECT * FROM [dbo].[TB_IMP_DISPONIVELSRE]
GO

/***************** FINALIZANDO TB_IMP_DISPONIVELSRE ***********************/

/***************** INICIANDO TB_IMP_INDISPONIVEL **************************/

INSERT INTO [dbo].[TB_LOA_DETALHADA]
       SELECT I.Funcional, I.Ano, I.PTRES, I.CodUO, I.ModDesp, I.GND, I.FonteSOF, I.RPrimario,
			  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, I.VL_Cred_Bloq_Rema, I.VL_Cred_Bloq_Contido,
              I.VL_Cred_Bloq_SOF, I.VL_Cred_Bloq_RemaSOF, I.VL_Cred_Bloq_Controle, I.VL_Cred_Bloq_RP,
              I.VL_Cred_Bloq_PreEmp
	   FROM TB_IMP_INDISPONIVEL I
       WHERE NOT EXISTS(SELECT * FROM TB_LOA_DETALHADA D WHERE D.Funcional = I.Funcional AND D.Ano = I.Ano AND D.PTRES = I.PTRES AND
                                                               D.CodUO = I.CodUO AND D.ModDesp = I.ModDesp AND D.GND = I.GND AND 
    			  									           D.FonteSOF = I.FonteSOF AND D.RPrimario = I.RPrimario)
GO

INSERT INTO [dbo].[TB_FUNCIONAL_PTRES]
       SELECT E.Funcional, E.Ano, E.PTRES, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '0 - Nenhuma', 0, 0, 0, '', '', '', P.SiglaDir, 0, 0, '' 
	   FROM TB_IMP_LOA_DETALHADA E INNER JOIN TB_ETAPAS P ON (P.Funcional = E.Funcional AND P.Ano = E.Ano)
	   WHERE NOT EXISTS(SELECT * FROM TB_FUNCIONAL_PTRES F WHERE F.Funcional = E.Funcional AND F.Ano = E.Ano AND F.PTRES = E.PTRES)
GO


INSERT INTO [dbo].[TB_LOA_DETALHADA]
       SELECT D.Funcional, D.Ano, D.PTRES, D.CodUO, D.ModDesp, D.GND, D.FonteSOF, D.RPrimario,
			  D.VL_PLOA, D.VL_Orc_Ini, D.VL_Cred_Sup, D.VL_Cred_Esp, D.VL_Cred_Ext, D.VL_Orc_Aut,
			  D.VL_Cred_Can, D.VL_Disponivel, D.VL_Empenhado, D.VL_Liquidado, D.VL_Pago, 
	          0, 0, 0, 0, 0, 0, 0 
	   FROM TB_IMP_LOA_DETALHADA D
	   WHERE NOT EXISTS(SELECT * FROM TB_LOA_DETALHADA L WHERE L.Funcional = D.Funcional AND L.Ano = D.Ano AND L.PTRES = D.PTRES AND
                                                               L.CodUO = D.CodUO AND L.ModDesp = D.ModDesp AND L.GND = D.GND AND 
    												           L.FonteSOF = D.FonteSOF AND L.RPrimario = D.RPrimario)
GO

UPDATE [dbo].[TB_LOA_DETALHADA]
   SET [VL_Cred_Bloq_Rema] = I.VL_Cred_Bloq_Rema
      ,[VL_Cred_Bloq_Contido] = I.VL_Cred_Bloq_Contido
      ,[VL_Cred_Bloq_SOF] = I.VL_Cred_Bloq_SOF
      ,[VL_Cred_Bloq_RemaSOF] = I.VL_Cred_Bloq_RemaSOF
      ,[VL_Cred_Bloq_Controle] = I.VL_Cred_Bloq_Controle
      ,[VL_Cred_Bloq_RP] = I.VL_Cred_Bloq_RP
      ,[VL_Cred_Bloq_PreEmp] = I.VL_Cred_Bloq_PreEmp
   FROM TB_LOA_DETALHADA D INNER JOIN TB_IMP_INDISPONIVEL I ON (D.Funcional = I.Funcional AND D.Ano = I.Ano AND D.PTRES = I.PTRES AND
                                                                D.CodUO = I.CodUO AND D.ModDesp = I.ModDesp AND D.GND = I.GND AND 
      			  									            D.FonteSOF = I.FonteSOF AND D.RPrimario = I.RPrimario)
GO

UPDATE [dbo].[TB_LOA_DETALHADA]
   SET [VL_PLOA] = I.VL_PLOA
      ,[VL_Orc_Ini] = I.VL_Orc_Ini
      ,[VL_Cred_Sup] = I.VL_Cred_Sup
      ,[VL_Cred_Esp] = I.VL_Cred_Esp
      ,[VL_Cred_Ext] = I.VL_Cred_Ext
      ,[VL_Orc_Aut] = I.VL_Orc_Aut
      ,[VL_Cred_Can] = I.VL_Cred_Can
      ,[VL_Disponivel] = I.VL_Disponivel
      ,[VL_Empenhado] = I.VL_Empenhado
      ,[VL_Liquidado] = I.VL_Liquidado
      ,[VL_Pago] = I.VL_Pago
   FROM TB_LOA_DETALHADA D INNER JOIN TB_IMP_LOA_DETALHADA I ON (D.Funcional = I.Funcional AND D.Ano = I.Ano AND D.PTRES = I.PTRES AND
                                                                 D.CodUO = I.CodUO AND D.ModDesp = I.ModDesp AND D.GND = I.GND AND 
     			  									             D.FonteSOF = I.FonteSOF AND D.RPrimario = I.RPrimario)
GO

DELETE FROM [dbo].[TB_IMP_INDISPONIVEL]
       WHERE EXISTS(SELECT * FROM TB_LOA_DETALHADA L WHERE L.Funcional = TB_IMP_INDISPONIVEL.Funcional AND L.Ano = TB_IMP_INDISPONIVEL.Ano AND 
	                                                       L.PTRES = TB_IMP_INDISPONIVEL.PTRES AND L.CodUO = TB_IMP_INDISPONIVEL.CodUO AND 
														   L.ModDesp = TB_IMP_INDISPONIVEL.ModDesp AND L.GND = TB_IMP_INDISPONIVEL.GND AND 
    												       L.FonteSOF = TB_IMP_INDISPONIVEL.FonteSOF AND L.RPrimario = TB_IMP_INDISPONIVEL.RPrimario)
GO

DELETE FROM [dbo].[TB_IMP_LOA_DETALHADA]
       WHERE EXISTS(SELECT * FROM TB_LOA_DETALHADA L WHERE L.Funcional = TB_IMP_LOA_DETALHADA.Funcional AND L.Ano = TB_IMP_LOA_DETALHADA.Ano AND 
	                                                       L.PTRES = TB_IMP_LOA_DETALHADA.PTRES AND L.CodUO = TB_IMP_LOA_DETALHADA.CodUO AND 
														   L.ModDesp = TB_IMP_LOA_DETALHADA.ModDesp AND L.GND = TB_IMP_LOA_DETALHADA.GND AND 
    												       L.FonteSOF = TB_IMP_LOA_DETALHADA.FonteSOF AND L.RPrimario = TB_IMP_LOA_DETALHADA.RPrimario)
GO

SELECT * FROM [dbo].[TB_IMP_INDISPONIVEL]

SELECT * FROM [dbo].[TB_IMP_LOA_DETALHADA]

/***************** FINALIZANDO TB_IMP_INDISPONIVEL **********************/

/***************** INICIANDO TB_IMP_CORFIN_RECEITA **************************/

INSERT INTO [dbo].[TB_FONTE]
    SELECT DISTINCT FonteSOF, FonteSOF + 'A CLASSIFICAR' FROM TB_IMP_CORFIN_RECEITA I
	WHERE NOT EXISTS(SELECT * FROM TB_FONTE F WHERE F.FonteSOF = I.FonteSOF)
GO

INSERT INTO [dbo].[TB_NATUREZADETALHADA]
     SELECT  DISTINCT '27', NaturezaDetalhada, NaturezaDetalhada + 'A CLASSIFICAR'
     FROM TB_IMP_CORFIN_RECEITA I
	 WHERE NOT EXISTS (SELECT * FROM TB_NATUREZADETALHADA D WHERE D.Grupo = '27' AND D.NaturezaDetalhada = I.NaturezaDetalhada)
GO

INSERT INTO [dbo].[TB_CORFIN_RECEITA]
       SELECT I.MesAnoString,I.FonteSOF,I.NaturezaDetalhada,I.EspecieReceita
             ,I.VL_Previsao_Receita,I.VL_Receita_Bruta,I.VL_Receita_Deducao
	         ,I.VL_Receita_Liquida,I.VL_Receita_Executada
			 ,0, 'SISTEMA', CONVERT(datetime, GETDATE(), 103)
	   FROM TB_IMP_CORFIN_RECEITA I 
	   WHERE NOT EXISTS(SELECT * FROM TB_CORFIN_RECEITA R WHERE R.MesAnoString = I.MesAnoString AND R.FonteSOF = I.FonteSOF AND R.NaturezaDetalhada = I.NaturezaDetalhada)
GO

UPDATE [dbo].[TB_CORFIN_RECEITA]
   SET [VL_Previsao_Receita] = I.VL_Previsao_Receita
      ,[VL_Receita_Bruta] = I.VL_Receita_Bruta
      ,[VL_Receita_Deducao] = I.VL_Receita_Deducao
      ,[VL_Receita_Liquida] = I.VL_Receita_Liquida
      ,[VL_Receita_Executada] = I.VL_Receita_Executada
   FROM TB_CORFIN_RECEITA R INNER JOIN TB_IMP_CORFIN_RECEITA I ON (R.MesAnoString = I.MesAnoString AND R.FonteSOF = I.FonteSOF AND R.NaturezaDetalhada = I.NaturezaDetalhada)
GO

DELETE FROM [dbo].[TB_IMP_CORFIN_RECEITA]
      WHERE EXISTS(SELECT * FROM TB_CORFIN_RECEITA R WHERE R.MesAnoString = TB_IMP_CORFIN_RECEITA.MesAnoString AND R.FonteSOF = TB_IMP_CORFIN_RECEITA.FonteSOF AND R.NaturezaDetalhada = TB_IMP_CORFIN_RECEITA.NaturezaDetalhada)
GO

SELECT * FROM TB_IMP_CORFIN_RECEITA

/***************** FINALIZANDO TB_IMP_CORFIN_RECEITA **********************/

/***************** INICIANDO TB_IMP_CREDITOS  *****************************/

INSERT INTO [dbo].[TB_CREDITOS]
	SELECT I.Funcional, I.Ano, I.PTRES, I.GND, I.FonteSOF, I.RPrimario, I.VL_PLOA, I.VL_Orc_Ini, I.VL_Cred_Sup, I.VL_Cred_Esp,
           I.VL_Cred_Ext, I.VL_Orc_Aut, I.VL_Cred_Can 
	FROM TB_IMP_CREDITOS I
	WHERE NOT EXISTS(SELECT * FROM TB_CREDITOS C WHERE C.Funcional = I.Funcional AND C.Ano = I.Ano AND C.PTRES = I.PTRES AND C.GND = I.GND AND C.FonteSOF = I.FonteSOF AND C.RPrimario = I.RPrimario)
GO

UPDATE [dbo].[TB_CREDITOS]
   SET [Funcional] = I.Funcional
      ,[Ano] = I.Ano
      ,[PTRES] = I.PTRES
      ,[GND] = I.GND
      ,[FonteSOF] = I.FonteSOF
      ,[RPrimario] = I.RPrimario
      ,[VL_PLOA] = I.VL_PLOA
      ,[VL_Orc_Ini] = I.VL_Orc_Ini
      ,[VL_Cred_Sup] = I.VL_Cred_Sup
      ,[VL_Cred_Esp] = I.VL_Cred_Esp
      ,[VL_Cred_Ext] = I.VL_Cred_Ext
      ,[VL_Orc_Aut] = I.VL_Orc_Aut
      ,[VL_Cred_Can] = I.VL_Cred_Can
	FROM TB_CREDITOS C INNER JOIN TB_IMP_CREDITOS I ON (C.Funcional = I.Funcional AND C.Ano = I.Ano AND C.PTRES = I.PTRES AND C.GND = I.GND AND C.FonteSOF = I.FonteSOF AND C.RPrimario = I.RPrimario)
GO

DELETE FROM [dbo].[TB_IMP_CREDITOS]
	WHERE EXISTS(SELECT * FROM TB_CREDITOS C WHERE C.Funcional = TB_IMP_CREDITOS.Funcional AND C.Ano = TB_IMP_CREDITOS.Ano AND C.PTRES = TB_IMP_CREDITOS.PTRES AND C.GND = TB_IMP_CREDITOS.GND AND C.FonteSOF = TB_IMP_CREDITOS.FonteSOF AND C.RPrimario = TB_IMP_CREDITOS.RPrimario)
GO

SELECT * FROM TB_IMP_CREDITOS


/***************** FINALIZANDO TB_IMP_CREDITOS **********************/

/***************** INICIANDO TB_IMP_CORFIN_NE_DIARIO  *****************************/

INSERT INTO TB_CORFIN_NE_DIARIO
SELECT I.[Ano]
      ,CONVERT(DATE,I.[DT_String],103)
      ,I.[CodEspecie]
      ,I.[VL_Empenhado]
  FROM TB_IMP_CORFIN_NE_DIARIO I
  WHERE NOT EXISTS(SELECT * FROM TB_CORFIN_NE_DIARIO C WHERE C.Ano = I.Ano AND C.CodEspecie = I.CodEspecie AND C.DT_Emissao = CONVERT(DATE,I.DT_String,103))
GO

DELETE FROM [dbo].[TB_IMP_CORFIN_NE_DIARIO]
	WHERE EXISTS(SELECT * FROM TB_CORFIN_NE_DIARIO C WHERE C.Ano = TB_IMP_CORFIN_NE_DIARIO.Ano AND C.CodEspecie = TB_IMP_CORFIN_NE_DIARIO.CodEspecie AND C.DT_Emissao = CONVERT(DATE,TB_IMP_CORFIN_NE_DIARIO.DT_String,103))
GO

SELECT * FROM TB_IMP_CORFIN_NE_DIARIO


/***************** FINALIZANDO TB_IMP_CORFIN_NE_DIARIO **********************/

/***************** INICIANDO TB_IMP_UG_DESTAQUE  *****************************/

INSERT INTO [dbo].[TB_UG_DESTAQUE]
      SELECT I.Funcional,I.Ano,I.PTRES,I.VL_Orc_Aut,I.VL_Provisao_Concedida,I.VL_Destacado,I.VL_Disponivel 
        ,I.VL_Indisponivel,I.VL_Cred_Bloq_PreEmp,I.VL_Empenhado,I.VL_Liquidado,I.VL_Pago 
      FROM TB_IMP_UG_DESTAQUE I
      WHERE NOT EXISTS(SELECT * FROM TB_UG_DESTAQUE U WHERE U.Funcional = I.Funcional AND U.Ano = I.Ano AND U.PTRES = I.PTRES)
GO

UPDATE [dbo].[TB_UG_DESTAQUE]
  SET  [VL_Orc_Aut] = I.VL_Orc_Aut
      ,[VL_Provisao_Concedida] = I.VL_Provisao_Concedida
      ,[VL_Destacado] = I.VL_Destacado
      ,[VL_Disponivel] = I.VL_Disponivel
      ,[VL_Indisponivel] = I.VL_Indisponivel
      ,[VL_Cred_Bloq_PreEmp] = I.VL_Cred_Bloq_PreEmp
      ,[VL_Empenhado] = I.VL_Empenhado
      ,[VL_Liquidado] = I.VL_Liquidado
      ,[VL_Pago] = I.VL_Pago
  FROM TB_UG_DESTAQUE U INNER JOIN TB_IMP_UG_DESTAQUE I ON (U.Funcional = I.Funcional AND U.Ano = I.Ano AND U.PTRES = I.PTRES)
GO

DELETE FROM [dbo].[TB_IMP_UG_DESTAQUE]
	WHERE EXISTS(SELECT * FROM TB_UG_DESTAQUE U WHERE U.Funcional = TB_IMP_UG_DESTAQUE.Funcional AND U.Ano = TB_IMP_UG_DESTAQUE.Ano AND U.PTRES = TB_IMP_UG_DESTAQUE.PTRES)
GO

SELECT * FROM TB_IMP_UG_DESTAQUE

/***************** FINALIZANDO TB_IMP_UG_DESTAQUE **********************/

/***************** INICIANDO TB_CGOR_PROGRAMACAO ******************************/

DELETE FROM [dbo].[TB_CGOR_PROGRAMACAO]
DECLARE @Mes as integer = 1
WHILE (@Mes <= 12)
BEGIN
INSERT INTO [dbo].[TB_CGOR_PROGRAMACAO]
	SELECT C.Processo
		  ,C.Funcional
		  ,C.Ano
		  ,E.CodAcao
	      ,@Mes AS Mes 
		  ,(SELECT SUM(CASE @Mes
				WHEN 1 THEN X.VL_Mes_01
				WHEN 2 THEN X.VL_Mes_02
				WHEN 3 THEN X.VL_Mes_03
				WHEN 4 THEN X.VL_Mes_04
				WHEN 5 THEN X.VL_Mes_05
				WHEN 6 THEN X.VL_Mes_06
				WHEN 7 THEN X.VL_Mes_07
				WHEN 8 THEN X.VL_Mes_08
				WHEN 9 THEN X.VL_Mes_09
				WHEN 10 THEN X.VL_Mes_10
				WHEN 11 THEN X.VL_Mes_11
				WHEN 12 THEN X.VL_Mes_12
				ELSE 0
				END)  FROM TB_CRONOGRAMA X WHERE X.Processo = C.Processo AND X.Funcional = C.Funcional AND X.Ano = C.Ano) AS VL_Programado
    FROM TB_CRONOGRAMA C INNER JOIN TB_ETAPAS E ON (E.Funcional = C.Funcional AND E.Ano = c.Ano)
	GROUP BY C.Processo, C.Funcional, C.Ano, E.CodAcao
    SET @Mes += 1
END
GO

/***************** FINALIZANDO TB_CGOR_PROGRAMACAO **********************/

/***************** INICIANDO TB_PARAMETRO ******************************/

UPDATE [dbo].[TB_PARAMETRO]
   SET [DT_Atualizacao_Base] = CONVERT(datetime, GETDATE(), 103)
GO

/***************** FINALIZANDO TB_PARAMETRO **********************/