import os, os.path, sys, shutil, subprocess
import win32com.client, win32api

#############################################
# Script automatizador de tabelas do PLOAWEB#
# Programador: Rafael Gomes da Silva        # 
# Versão: Beta 1.0                          #
# Data: 09/03/2023                          #
#############################################

#1) Executar o macro

print("Executando o MACRO")
win32api.MessageBox(0, 'Executando o Macro', 'Script')

if os.path.exists("Macro_PLOAWEB_V52.xlsm"):
    plan=win32com.client.Dispatch("Excel.Application")
    plan.Workbooks.Open(os.path.abspath("Macro_PLOAWEB_V52.xlsm"), ReadOnly=1)
    plan.Application.Run("Macro_PLOAWEB_V52.xlsm!Macro_PLOAWEB")
##  plan.Application.Save() # Caso queira salvar, descomeente esta linha e apague o ", ReadOnly=1" da função Open.
    plan.Application.Quit() # Comment this out if your excel script closes
    del plan

#2) Mover os arquivos para o Z

# set the path to the source folder
source_folder = "C:/Users/rafael.sgomes/Documents/ploaTabelas"

# set the path to the destination folder
#destination_folder = "\\\\10.100.10.174\\ploa_carga"
destination_folder= "C:/Users/rafael.sgomes/Documents/ploaTabelas/destino"

# list all the files in the source folder
files = os.listdir(source_folder)

# count the number of files found
count = 0
eof = 0

#contando os arquivos EDIT

for file in files:
    if "EDIT" in files:
        eof +=1

# Loop sobre os arquivos  e move os arquivos EDIT para a pasta de destino
for file in files:
    if "EDIT" in file:
        #Move o arquivo para 
        shutil.move(os.path.join(source_folder, file), os.path.join(destination_folder, file))
        count += 1

        # break the loop when files have been moved
        if count == eof:
            break

# run the SQL queries
#queries = ["6 - SQL_INSERT_HISTORICO_MEDICAO.sql", "7 - SQL_INSERT_CONTRATO_SIAC.sql", "23 - DBPLOAWEB ALL THE DAYS.sql"]

#for query in queries:
    # set the path to the SQL query file
#    query_path = "C:/Users/rafael.sgomes/Documents/ploaTabelas" + query

    # run the SQL query using subprocess
#    subprocess.run(["sqlcmd", "-S", "server_name", "-d", "database_name", "-U", "usuario", "-P", "password" "-i", query_path])


