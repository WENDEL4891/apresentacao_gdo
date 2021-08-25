import pandas as pd
import numpy as np
import os
from pprint import pprint

def get_files_path(ano, mes):
    files_path = dict()
    aux_path = f'files/{ano}/{mes}/dados/'
    for file_name in os.listdir(f'files/{ano}/{mes}/dados'):
        if 'IAF' in file_name:
            files_path['IAF'] = os.path.abspath(aux_path + file_name)
        elif 'TCV' in file_name:
            files_path['TCV'] = os.path.abspath(aux_path + file_name)
        elif 'THC' in file_name:
            files_path['THC'] = os.path.abspath(aux_path + file_name)
        elif 'TRI' in file_name:
                files_path['TRI'] = os.path.abspath(aux_path + file_name)
        elif 'OLS' in file_name:
                files_path['OLS'] = os.path.abspath(aux_path + file_name)
        elif 'TQF' in file_name:
                files_path['TQF'] = os.path.abspath(aux_path + file_name)
    return files_path


def definir_campos(df):
    for field_name in df.columns:
        if 'Qtd' in field_name:
            qtd = field_name
        if 'Número REDS' in field_name:
            NREDS = field_name
    
    df = df[
        [
            NREDS, 'Ano Fato', 'Mês Numérico Fato', 'Data Fato',
            'Município', 'Bairro', 'Logradouro Ocorrência', 'Bairro Não Cadastrado',
            'Logradouro Ocorrência Não Cadastrado', 'Complemento Endereço',
            'Ponto de Referência', qtd
        ]
    ]
    df.columns = [
            'NREDS', 'ANO_FATO', 'MES_FATO', 'DATA_FATO',
            'MUNICIPIO', 'BAIRRO', 'LOGRADOURO', 'BAIRRO_NAO_CADASTRADO',
            'LOGRADOURO_NAO_CADASTRADO', 'COMPLEMENTO_ENDERECO',
            'PONTO_DE_REFERENCIA', 'QTDE'
        ]
    return df


def get_df_classif():
        df_classif = pd.read_csv('files/classificadores.csv')
        df_classif['MUNICIPIO'] = df_classif['MUNICIPIO'].str.upper()
        df_classif['CLASSIFICADOR'] = df_classif['CLASSIFICADOR'].str.upper()
        df_classif.set_index(df_classif['MUNICIPIO'] + " " + df_classif['CLASSIFICADOR_TIPO'] + " " + df_classif['CLASSIFICADOR'], inplace=True)    
        df_classif.fillna('', inplace=True)
        df_classif = df_classif[~df_classif.index.duplicated(keep='first')]        
        return df_classif


def classifica_setor(row, df_classif):
    municipio = row['MUNICIPIO']    
    if ( municipio + 'NREDS' + row.name ) in df_classif.index:
        return df_classif.loc[municipio+" NREDS "+row['NREDS'], 'SETOR']
    elif municipio == 'CLAUDIO':
        return 'CLAUDIO'
    elif municipio == 'ITATIAIUCU':
            return 'LOURDES/ITATIAIUCU'
    elif municipio in ('CARMO DO CAJURU', 'SAO GONCALO DO PARA'):             
        return 'CARMO DO CAJURU/SAO GONCALO DO PARA'    
    elif municipio + ' BAIRRO ' + row['BAIRRO'] in df_classif.index:       
        return ( df_classif.loc[municipio + ' BAIRRO ' + row['BAIRRO'].upper(), 'SETOR'] ).upper()
    elif municipio + ' BAIRRO_NAO_CADASTRADO ' + row['BAIRRO_NAO_CADASTRADO'] in df_classif.index:       
        return ( df_classif.loc[municipio + ' BAIRRO_NAO_CADASTRADO ' + row['BAIRRO_NAO_CADASTRADO'].upper(), 'SETOR'] ).upper()
    elif municipio + ' LOGRADOURO ' + row['LOGRADOURO'] in df_classif.index:
        return df_classif.loc[municipio + ' LOGRADOURO ' + row['LOGRADOURO'].upper(), 'SETOR']    
    elif municipio + ' LOGRADOURO_NAO_CADASTRADO ' + row['LOGRADOURO_NAO_CADASTRADO'] in df_classif.index:
        return df_classif.loc[municipio + ' LOGRADOURO_NAO_CADASTRADO ' + row['LOGRADOURO_NAO_CADASTRADO'].upper(), 'SETOR']    
    elif municipio + ' COMPLEMENTO_ENDERECO ' + row['COMPLEMENTO_ENDERECO'] in df_classif.index:
        return df_classif.loc[municipio + ' COMPLEMENTO_ENDERECO ' + row['COMPLEMENTO_ENDERECO'].upper(), 'SETOR']
    elif ( municipio + ' PONTO_DE_REFERENCIA ' + row['PONTO_DE_REFERENCIA'] ) in df_classif.index:
        return df_classif.loc[municipio + ' PONTO_DE_REFERENCIA ' + row['PONTO_DE_REFERENCIA'], 'SETOR']    
    else:
        return 'SETOR_INDEFINIDO'

def classifica_cia(df):
        conds=[
            df['MUNICIPIO'].isin(['ITAUNA','ITATIAIUCU']),
            df['SETOR'].isin(['HIPER CENTRO','BOM PASTOR','ALTO GOIAS']),
            df['SETOR'].isin(['PLANALTO','SAO JOSE','CLAUDIO']),
            df['SETOR'].isin(['NITEROI','PORTO VELHO','CARMO DO CAJURU/SAO GONCALO DO PARA']),
            
        ]
        res=['51 CIA','53 CIA','139 CIA','142 CIA']
        df['CIA'] = np.select(conds,res,default='CIA_INDEFINIDA')