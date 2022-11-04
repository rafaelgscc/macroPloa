import os, os.path, sys
import win32com.client, win32api

#############################################
# Script automatizador de tabelas do PLOAWEB#
# Programador: Rafael Gomes da Silva        # 
# Versão: Alpha 0.1                         #
# Data: 03/11/2022                          #
#############################################

#1) baixando os arquivos do e-mail



#2) Executar o macro

print("Executando o MACRO")
win32api.MessageBox(0, 'Executando o Macro', 'Script')

if os.path.exists("Macro_PLOAWEB_V49.xlsm"):
    plan=win32com.client.Dispatch("Excel.Application")
    plan.Workbooks.Open(os.path.abspath("Macro_PLOAWEB_V49.xlsm"), ReadOnly=1)
    plan.Application.Run("Macro_PLOAWEB_V49.xlsm!Macro_PLOAWEB")
##  plan.Application.Save() # Caso queira salvar, descomeente esta linha e apague o ", ReadOnly=1" da função Open.
    plan.Application.Quit() # Comment this out if your excel script closes
    del plan
    sys.exit(0)

#3) Mover os arquivos para o Z


