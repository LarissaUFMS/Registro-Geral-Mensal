# Conector de sql recomendado na documentação oficial.
import pyodbc

import pandas as pd

if __name__ == '__main__':
	# Conectando com o banco de dados
	con = pyodbc.connect(
	'DRIVER={SQL Server};SERVER=**.***.**.**;PORT=1433;DATABASE=SCI;Trusted_Connection=yes;')
	#Consultando a tabela do banco
	conHistorico= 'select REGIONAL, LOCAL as LOCALIDADE, REFERENCIA as DATA, RDA, RCE, ECON, DDIFF from dbo.historico_resultado ORDER BY REGIONAL, LOCAL, DATAGERACAO desc'
	conResultado ='select REGIONAL, LOCAL as LOCALIDADE, REFERENCIA as DATA, RDA, RCE, ECON, DDIFF from dbo.resultado ORDER BY REGIONAL, LOCAL, DATAGERACAO desc;'
	#Convertendo em dataframe
	dfHistorico = pd.read_sql(conHistorico, con)
	dfResultado = pd.read_sql(conResultado, con)
	#Agrupando por Localidade e Data
	dfgroup = dfResultado.groupby(["LOCALIDADE", "DATA"])
	#Selecionando a primeira linha de cada grupo
	dfResultado = dfgroup.head(1)
	#Concatenando os dataframes
	df = pd.concat([dfResultado, dfHistorico], ignore_index=True, sort=False)
	#Apagando as linhas duplicadas
	df = df.drop_duplicates()
	#Ordenando o dataframe primeiro pela redional e depois por localidade
	df = df.sort_values(by = ['REGIONAL', 'LOCALIDADE'])
	#Exportando para o excel
	RGmensal = df.to_excel(excel_writer = "RegistroGeral.xlsx", sheet_name = 'Registro_Geral')
