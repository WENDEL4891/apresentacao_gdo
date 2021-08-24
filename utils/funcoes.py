import pandas as pd
import numpy as np

def definir_campos(df):
    df = df[
        [
            'Número REDS', 'Ano Fato', 'Mês Numérico Fato', 'Data Fato',
            'Município', 'Bairro', 'Logradouro Ocorrência', 'Bairro Não Cadastrado',
            'Logradouro Ocorrência Não Cadastrado', 'Complemento Endereço',
            'Ponto de Referência', 'Qtde Ocorrências'
        ]
    ]
    df.columns = [
            'NREDS', 'ANO_FATO', 'MES_FATO', 'DATA_FATO',
            'MUNICIPIO', 'BAIRRO', 'LOGRADOURO', 'BAIRRO_NAO_CADASTRADO',
            'LOGRADOURO_NAO_CADASTRADO', 'COMPLEMENTO_ENDERECO',
            'PONTO_DE_REFERENCIA', 'QTDE OCORRENCIAS'
        ]
    return df


def get_df_classif():
        df_classif = pd.read_csv('files/dados/classificadores.csv')
        df_classif.set_index( df_classif['MUNICIPIO'] + " " + df_classif['CLASSIFICADOR_TIPO'] + " " + df_classif['CLASSIFICADOR'], inplace=True)    
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