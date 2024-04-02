#biblioteca pandas
import pandas as pd

#importando os dados de uma guia de uma planilha:
df = pd.read_excel('C:/ltd/IPPGR3_teste.xlsx', sheet_name='Total de produções')

# excluindo a terceira coluna, pois não tem dados utilizados nos cálculos
df = df.drop(df.columns[2], axis=1)

#eliminando as 8 primeiras linhas e criando um novo data frame
df_ippgr3 = df.drop(range(0, 9))

#algumas colunas finais em branco foram retiradas
df_ippgr3 = df_ippgr3.iloc[:, :21]

# renomeando as colunas para facilitar o trabalho
df_ippgr3.rename(columns={df_ippgr3.columns[0]: 'nome'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[1]: 'nome_maisculas'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[2]: 'qtde_orient_mestrado'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[3]: 'qtde_orient_iniciacao'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[4]: 'qtde_orient_pos'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[5]: 'qtde_orient_doutorado'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[6]: 'qtde_orient_graduacao'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[7]: 'artigos_a1'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[8]: 'artigos_a2'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[9]: 'artigos_a3'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[10]: 'artigos_a4'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[11]: 'artigos_b1'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[12]: 'artigos_b2'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[13]: 'artigos_b3'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[14]: 'artigos_b4'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[15]: 'artigos_c'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[16]: 'artigos_sem_pontuacao'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[17]: 'qtde_cap_livros'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[18]: 'qtde_livros'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[19]: 'qtde_trabs_anais'}, inplace=True);
df_ippgr3.rename(columns={df_ippgr3.columns[20]: 'ippgr3'}, inplace=True);

#Cálculo do IPPGR3
df_ippgr3.iloc[:, 20] = (((((((df_ippgr3.iloc[:, 7]*35*2.5) +    #artigo_a1
(df_ippgr3.iloc[:, 8]*35*1.5) +       #artigo_a2
(df_ippgr3.iloc[:, 9]*35*1.5) +       #artigo_a3
(df_ippgr3.iloc[:, 10]*35*1.5) +      #artigo_a4
(df_ippgr3.iloc[:, 11]*35*1.15) +     #artigo_b1
(df_ippgr3.iloc[:, 12]*35*1.1) +      #artigo_b2
(df_ippgr3.iloc[:, 13]*35*1.05) +     #artigo_b3
(df_ippgr3.iloc[:, 14]*35*1.03) +     #artigo_b4
(df_ippgr3.iloc[:, 15]*35*1.01) +     #artigo_c
(df_ippgr3.iloc[:, 16]*35*1) +        #artigo_np
(df_ippgr3.iloc[:, 18]*30) +          #qtde_livros
(df_ippgr3.iloc[:, 17]*10) +          #qtde_cap_livros
(df_ippgr3.iloc[:, 19]*15))/95)*3) +  #qtde_anaislivros
((((df_ippgr3.iloc[:, 5]*50) +        #qtde_tese
(df_ippgr3.iloc[:, 2]*25) +           #qtde_mestrado
(df_ippgr3.iloc[:, 3]*10) +           #qtde_iniciacao
(df_ippgr3.iloc[:, 6]*5) +            #qtde_tcc
(df_ippgr3.iloc[:, 4]*5))/95)*1))/4)/36)*10000 #qtde_monografia
     
# ***Exportar em .xlsx com os nomes das colunas originais***
#Determinando o nome da planilha com o cálculo - Está na pasta UNESA/LTD
arquivo = 'IPPGR3_calculado.xlsx'
  
# saving the excel
df_ippgr3.to_excel(arquivo)
