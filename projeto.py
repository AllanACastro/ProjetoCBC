import os, shutil
import pandas as pd

def limpar_pasta(folder):
  for filename in os.listdir(folder):
    file_path = os.path.join(folder, filename)
    try:
        if os.path.isfile(file_path) or os.path.islink(file_path):
            os.unlink(file_path)
        elif os.path.isdir(file_path):
            shutil.rmtree(file_path)
    except Exception as e:
        print('Failed to delete %s. Reason: %s' % (file_path, e))

def calcula_grupo(row):
      if row['Unnamed: 5'] in ['FFL']:
        return 'FLOAT'
      elif row['Unnamed: 5'] in ['ECB','EG1','EG10','FFP']:
          return 'ESPELHO'
      elif row['Unnamed: 5'] in ['LCL','CTV','LBB','LBN','LBS','LBZ','LFL','LHN','LHR','LRF']:
          return 'LAMINADO'
      elif row['Unnamed: 5']:
          return 'CONTROLE SOLAR'

def calcula_producao(row):
      if row['Unnamed: 4']<2039999:
        return 'EXTERNA'
      elif row['Unnamed: 4']:
        return 'CBC'

def calcula_mercado(row):
      if row['Unnamed: 4']<2039999:
        return 'MI'
      elif row['Unnamed: 4']<2049999:
        return 'ME'

def calcula_corte(row):
      if row['Unnamed: 9']<4400:
        return 'TR'
      elif row['Unnamed: 9']:
        return 'GV'

def calcula_ano(row):
    ano = pd.to_datetime(row['Unnamed: 22'], format='%d.%m.%Y').year
    return ano

def calcula_t(row):
    tTotal = float(row['Unnamed: 18']) + float(row['Unnamed: 20'])
    return round(tTotal,1)

def calcula_mes(row):
    mes = pd.to_datetime(row['Unnamed: 22'], format='%d.%m.%Y').month
    return mes

def days_diff(dataIni, dataEnd):
      delta = dataEnd - dataIni
      return delta.days

def calcula_idadeMes(row, AU):
    datarow = pd.to_datetime(row['Unnamed: 22'], format='%d.%m.%Y')
    idadeMes = days_diff(datarow, AU)/30
    return round(idadeMes, 1)

def calcula_idade(row):
      if row['Idade Meses']< 6:
        return '<6 Meses'
      elif row['Idade Meses']>24:
        return '>24 Meses'
      elif row['Idade Meses']>12:
        return 'De 12 a 24 Meses'
      elif row['Idade Meses']:
        return 'De 6 a 12 Meses'

def calcula_tempo(row):
      if row['Idade Meses']<2:
        return '<2Meses'
      elif row['Idade Meses']:
        return '<2Meses'

def calcula_idadeDias(row, AU):
    datarow = pd.to_datetime(row['Unnamed: 22'], format='%d.%m.%Y')
    idadeMes = days_diff(datarow, AU)
    return idadeMes


def organizaTabela(file, dataOrganize):

  AU = pd.to_datetime((dataOrganize), format='%Y/%m/%d')

  print('iniciado processo')

  df = pd.read_excel(file)

  df_ = df[11:]
  linhaCorte = df_[df_[df_.columns[1]].isna() == False].index[0]
  df_ = df_[df_.index <= linhaCorte-2].drop(df_.columns[:3], axis = 1)
  df_.head()
  print('criando colunas')


  df_['Familia'] = df_.apply(lambda x: calcula_grupo(x), axis=1)

  df_['Produção'] = df_.apply(lambda x: calcula_producao(x), axis=1)

  df_['Mercado'] = df_.apply(lambda x: calcula_mercado(x), axis=1)

  df_['Corte'] = df_.apply(lambda x: calcula_corte(x), axis=1)

  df_['(t)'] = df_.apply(lambda x: calcula_t(x), axis=1)

  df_['Ano'] = df_.apply(lambda x: calcula_ano(x), axis=1)

  df_['Mês'] = df_.apply(lambda x: calcula_mes(x), axis=1)

  df_['Idade Meses'] = df_.apply(lambda x: calcula_idadeMes(x, AU), axis=1)

  df_['Idade'] = df_.apply(lambda x: calcula_idade(x), axis=1)

  df_['Tempo'] = df_.apply(lambda x: calcula_tempo(x), axis=1)

  df_['Idade (dias)'] = df_.apply(lambda x: calcula_idadeDias(x, AU), axis=1)

  headS= ['Tipo','Material','Grup','Espes','Cor','Qual','Alt','Lar','Chap','Emb','Int','Lote','Posição','Disponivel','Entregue','Bloqueado','Data fabr/','Hora Prod','Status','Origem','Saída','Lado Banho','Trat.Químico','Defeito','Hist.Bloqueio','Hist.Desbloq','Cor_','Met','Est','2ª','St Ant','Cód.Obra','Desc.Obra','DataVen','Data prod.float']
  names= ['Unnamed: 3','Unnamed: 4','Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 10','Unnamed: 12','Unnamed: 13','Unnamed: 14','Unnamed: 16','Unnamed: 17','Unnamed: 18','Unnamed: 19','Unnamed: 20','Unnamed: 22','Unnamed: 23','Unnamed: 24','Unnamed: 25','Unnamed: 26','Unnamed: 27','Unnamed: 29','Unnamed: 30','Unnamed: 31','Unnamed: 32','Unnamed: 33','Unnamed: 34','Unnamed: 35','Unnamed: 36','Unnamed: 37','Unnamed: 38','Unnamed: 39','Unnamed: 40','Unnamed: 42']

  dict(zip(headS, names))
  print('separando cabeçalhos')

  a_dict = dict(zip(names, headS ))

  df_ = df_.rename((a_dict), axis='columns')

  print('organizando cabeçalhos')

  df_= df_[['Tipo', 'Material', 'Produção', 'Mercado' ,'Grup','Familia', 'Espes','Cor','Qual','Alt','Lar','Corte','Chap','Emb','Int','Lote','Posição','Disponivel','Entregue','Bloqueado','(t)', 'Data fabr/', 'Ano','Mês','Idade Meses','Idade','Tempo','Idade (dias)','Hora Prod','Status','Origem','Saída','Lado Banho','Trat.Químico','Defeito','Hist.Bloqueio','Hist.Desbloq','Cor_','Met','Est','2ª','St Ant','Cód.Obra','Desc.Obra','DataVen','Data prod.float']]
  pd.set_option('display.max_columns', None)
  df_.head(20)

  print('criando arquivo')

  limpar_pasta('Arquivos')
  df_.to_excel('Arquivos/ESTOQUE_GERAL-FILTRADO_{0}.xlsx'.format(AU.strftime('%d-%m-%Y')), index = False)
