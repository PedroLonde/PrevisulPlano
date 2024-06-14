import streamlit as st
import pandas as pd

# Função para carregar os dados dos arquivos Excel
@st.cache
def load_data():
    nov = pd.read_excel("C:/Python/RelatorioNovo.xlsx", header=1)
    ant = pd.read_excel("C:/Python/RelatorioAntigo.xlsx")
    return nov, ant

# Carregar os dados
nov, ant = load_data()

# Processar os dados
nov = nov.drop(['Data de Cadastro', 'Data de Adesão', 'Data de Ativação', 
                'Tipo de Benefício', 'Sequencial de Benefício','Sequencial','Cartão de Desconto','Identidade','Unidade', 'Dia de Vencimento', 
                'Situação', 'Consultor', 'Bairro', 'Cidade', 'UF', 'Telefone', 
                'Email', 'Dados Adicionais'], axis=1)

nov = nov.dropna(subset=['CPF/CNPJ'])
nov.reset_index(drop=True, inplace=True)

if 'Idade' in nov.columns:
    for x in range(len(nov)):
        if 'Idade' in ant.columns:
            cpf_cnpj = nov.loc[x, 'CPF/CNPJ']
            ant_beneficiario = ant.loc[ant['CPF/CNPJ'] == cpf_cnpj]
            if not ant_beneficiario.empty:
                idade_antiga = ant_beneficiario.iloc[0]['Idade']
                if idade_antiga > 69:
                    nov.loc[x, 'Idade'] = idade_antiga

nov['Titular'] = None

x = 0
while x < len(nov):
    if nov.loc[x, 'Tipo de Beneficiário'] == 'Titular':
        nov.loc[x, 'Titular'] = nov.loc[x, 'CPF/CNPJ']
    elif nov.loc[x, 'Tipo de Beneficiário'] == 'Dependente':
        nov.loc[x, 'Titular'] = nov.loc[x-1, 'Titular']
    x += 1
    
x = 0
while x < len(nov):
    if nov.loc[x, 'Tipo de Beneficiário'] == 'Titular':
        y = x + 1
        while y < len(nov) and nov.loc[y, 'Tipo de Beneficiário'] != 'Titular':
            if (nov.loc[y, 'Idade'] <= 81 and nov.loc[y, 'Parentesco'] in ['Sogro(a)', 'Cônjuge', 'Mãe/Pai']) or (nov.loc[y, 'Idade'] <= 18 and nov.loc[y, 'Parentesco'] == 'Filho(a)'):
                break
            y += 1
        
        if y == len(nov) or nov.loc[y, 'Tipo de Beneficiário'] == 'Titular':
            if 'Plus' in nov.loc[x, 'Plano']:
                nov.loc[x, 'Plano'] = 'Pleno'
            elif 'Vital' in nov.loc[x, 'Plano']:
                nov.loc[x, 'Plano'] = 'Vital'
        
    x += 1

idade_maxima_titular = 69
excluir = []

for x in range(len(nov)):
    if x > 0 and nov.loc[x, 'Tipo de Beneficiário'] == 'Titular' and nov.loc[x-1, 'Tipo de Beneficiário'] == 'Titular' and nov.loc[x, 'Idade'] > idade_maxima_titular:
        excluir.append(x)
    elif nov.loc[x, 'Idade'] < 18 or (nov.loc[x, 'Idade'] > 69 and nov.loc[x, 'Tipo de Beneficiário'] == 'Dependente'):
        excluir.append(x)

excluir = list(set(excluir))
indices_para_excluir = [index for index in excluir if index in nov.index]
nov.drop(indices_para_excluir, inplace=True)

nov = nov[~nov['Parentesco'].isin(['Sogro(a)', 'Cônjuge', 'Mãe/Pai'])]
nov.reset_index(drop=True, inplace=True)

condicao_pleno = ((nov['Tipo de Beneficiário'] == 'Titular') & 
                  (nov['Plano'].str.contains('Prime|Pleno|DEPENDENTES PRIME|DEPENDENTES PLENO')))

condicao_plus = ((nov['Tipo de Beneficiário'] == 'Titular') & 
                 (nov['Plano'].str.contains('Plus')))

condicao_vital = ((nov['Tipo de Beneficiário'] == 'Titular') & 
                  (nov['Plano'] == 'Vital'))

nov.loc[condicao_pleno, 'Plano'] = 'Pleno'
nov.loc[condicao_plus, 'Plano'] = 'Plus'
nov.loc[condicao_vital, 'Plano'] = 'Vital'
nov.loc[~(condicao_pleno | condicao_plus | condicao_vital), 'Plano'] = 'Essencial'

nov = nov.drop(['Tipo de Beneficiário','Parentesco','Idade','Titular'], axis=1)
nov['Data de Nascimento'] = pd.to_datetime(nov['Data de Nascimento']).dt.strftime('%d/%m/%Y')

nov.to_excel("C:/Python/RelatorioAntigo.xlsx", index=False)
Pleno = nov[nov['Plano'].str.contains('Pleno')]
Pleno = Pleno.drop('Plano', axis=1)

novos_dados = {
    'Nome do Beneficiário': ['SOLANGE MORAES LONDE', 'MAURICO BATISTA LONDE'],
    'Sexo': ['Feminino', 'Masculino'],
    'CPF/CNPJ': ['006.042.546-60', '490.615.186-87'],
    'Data de Nascimento': ['17/05/1961', '17/10/1963'],
    'Endereço': ['Rua Lirio Montanhes, 66 Casa', 'Rua Marcio Garcia,30 Casa'],
    'CEP': ['30555-180', '34003-074']
}

novos_df = pd.DataFrame(novos_dados)
Pleno = pd.concat([Pleno, novos_df], ignore_index=True)

Plus = nov[nov['Plano'].str.contains('Plus')]
Plus = Plus.drop('Plano', axis=1)
Vital = nov[nov['Plano'].str.contains('Vital')]
Vital = Vital.drop('Plano', axis=1)
Essencial = nov[nov['Plano'].str.contains('Essencial')]
Essencial = Essencial.drop('Plano', axis=1)

Essencial.to_excel("C:/Python/12009820007990.xlsx", index=False)
Plus.to_excel("C:/Python/12009820007976.xlsx", index=False)
Vital.to_excel("C:/Python/12009820007972.xlsx", index=False)
Pleno.to_excel("C:/Python/12009820007974.xlsx", index=False)

conple = 0
conplu = 0
convit = 0
coness = 0

for x in range(len(nov)):
    if nov.loc[x,'Plano'] == 'Pleno':
        conple+=1
    elif nov.loc[x,'Plano'] == 'Plus':
        conplu+=1
    elif nov.loc[x,'Plano'] == 'Vital':
        convit+=1
    elif nov.loc[x,'Plano'] == 'Essencial':
        coness+=1

total = convit + conple + conplu + coness

valvit = convit*3.01
valple = conple*1.73
valplu = conplu*3.41
valess = coness*1.08

valtotal= valvit + valple + valplu + valess

dados = {
    'Vital': {'Quantidade': convit, 'Valor Total': valvit},
    'Pleno': {'Quantidade': conple, 'Valor Total': valple},
    'Plus': {'Quantidade': conplu, 'Valor Total': valplu},
    'Essencial': {'Quantidade': coness, 'Valor Total': valess},
    'Total': {'Quantidade': total, 'Valor Total': valtotal}
}

df = pd.DataFrame(dados)

# Interface do Streamlit
st.title("Relatório de Beneficiários")
st.write("### Dados Processados")
st.dataframe(nov)

st.write("### Quantidades e Valores por Plano")
st.dataframe(df)

st.write("### Beneficiários Pleno")
st.dataframe(Pleno)

st.write("### Beneficiários Plus")
st.dataframe(Plus)

st.write("### Beneficiários Vital")
st.dataframe(Vital)

st.write("### Beneficiários Essencial")
st.dataframe(Essencial)

# Baixar arquivos Excel
st.write("### Baixar Arquivos Excel")
st.download_button(label="Baixar Relatório Completo", data=nov.to_excel(index=False), file_name="RelatorioCompleto.xlsx")
st.download_button(label="Baixar Beneficiários Pleno", data=Pleno.to_excel(index=False), file_name="BeneficiariosPleno.xlsx")
st.download_button(label="Baixar Beneficiários Plus", data=Plus.to_excel(index=False), file_name="BeneficiariosPlus.xlsx")
st.download_button(label="Baixar Beneficiários Vital", data=Vital.to_excel(index=False), file_name="BeneficiariosVital.xlsx")
st.download_button(label="Baixar Beneficiários Essencial", data=Essencial.to_excel(index=False), file_name="BeneficiariosEssencial.xlsx")
