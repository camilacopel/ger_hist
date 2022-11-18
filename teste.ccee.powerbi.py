import pymongo
import pandas
from pymongo import MongoClient
from comum.shareplum import Site
from comum.shareplum.site import Version
from comum.shareplum import Office365
import io

client = pymongo.MongoClient("mongodb://user_rw:us3r_rw@nwautomhml/")
db = client["historico_oficial"]
col = db["ccee_geracao_tableau"]
dadostotais = col.find({}, {"sigla_ativo": 1, "GFIS": 1, "mes": 1, "fonte_energia_primaria": 1, "submercado": 1, "unidade_federativa": 1, "CAP_T": 1})
total = pandas.DataFrame(list(dadostotais))

total.loc[total['submercado'] == 'SE', 'submercado'] = "submercado sudeste"
total.loc[total['submercado'] == 'N', 'submercado'] = "submercado norte"
total.loc[total['submercado'] == 'S', 'submercado'] = "submercado sul"
total.loc[total['submercado'] == 'NE', 'submercado'] = "submercado nordeste"

total.loc[total['unidade_federativa'] == 'AM', 'unidade_federativa'] = "Amazonas, BR"
total.loc[total['unidade_federativa'] == 'RO', 'unidade_federativa'] = "Rondônia, BR"
total.loc[total['unidade_federativa'] == 'MG', 'unidade_federativa'] = "Minas Gerais, BR"
total.loc[total['unidade_federativa'] == 'PE', 'unidade_federativa'] = "Pernambuco, BR"
total.loc[total['unidade_federativa'] == 'SE', 'unidade_federativa'] = "Sergipe, BR"

#total.to_csv('DADOS_TOTAIS.csv', decimal=',', sep= ';', index= False)

endereco_sharepoint='https://4c80k2.sharepoint.com' 
endereco_site='https://4c80k2.sharepoint.com/sites/EstudoPCH'
chave_produto='rafaelpbaptista@4c80k2.onmicrosoft.com'
senha_chave_produto='t3st3C0p3l'
pasta_sharepoint_resultados='Documentos Partilhados/Estudo PCH'
nome_arquivo= 'DADOS_TOTAIS.csv'


authcookie = Office365(endereco_sharepoint, username= chave_produto,  password= senha_chave_produto).GetCookies()
site = Site(endereco_site, version=Version.v2016, authcookie=authcookie, verify_ssl=False)

folder = site.Folder(pasta_sharepoint_resultados)

bytes_io = io.BytesIO()

total.to_csv(bytes_io, sep=';', decimal=',', index=False)

bytes_io.seek(0)

folder.upload_file(bytes_io.read(), nome_arquivo)
