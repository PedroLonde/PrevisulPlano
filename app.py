import streamlit as st
import pandas as pd

def main():
    st.title("Processamento de Relatórios Excel")
    
    uploaded_file_nov = st.file_uploader("Escolha o arquivo Relatório Novo", type=["xlsx"], key="nov")
    uploaded_file_ant = st.file_uploader("Escolha o arquivo Relatório Antigo", type=["xlsx"], key="ant")
    uploaded_file_pend = st.file_uploader("Escolha o arquivo Pendentes Alteração", type=["xlsx"], key="pend")

    if uploaded_file_nov and uploaded_file_ant and uploaded_file_pend:
        nov = pd.read_excel(uploaded_file_nov, header=1)
        ant = pd.read_excel(uploaded_file_ant)
        dados_pendentes = pd.read_excel(uploaded_file_pend)

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
        
        for index, row in nov.iterrows():
            if pd.isnull(row['Sexo']):
                beneficiario = row['Nome do Beneficiário']
                sexo_correspondente = dados_pendentes.loc[dados_pendentes['Nome do Beneficiário'] == beneficiario, 'Sexo'].values
                if len(sexo_correspondente) > 0:
                    nov.at[index, 'Sexo'] = sexo_correspondente[0]

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

        Pleno = nov[nov['Plano'].str.contains('Pleno')].drop('Plano', axis=1)
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

        Plus = nov[nov['Plano'].str.contains('Plus')].drop('Plano', axis=1)
        Vital = nov[nov['Plano'].str.contains('Vital')].drop('Plano', axis=1)
        Essencial = nov[nov['Plano'].str.contains('Essencial')].drop('Plano', axis=1)

        st.write("Pleno")
        st.write(Pleno)
        st.write("Plus")
        st.write(Plus)
        st.write("Vital")
        st.write(Vital)
        st.write("Essencial")
        st.write(Essencial)

        st.download_button(label="Download Pleno", data=Pleno.to_csv(index=False), file_name="Pleno.csv", mime="text/csv")
        st.download_button(label="Download Plus", data=Plus.to_csv(index=False), file_name="Plus.csv", mime="text/csv")
        st.download_button(label="Download Vital", data=Vital.to_csv(index=False), file_name="Vital.csv", mime="text/csv")
        st.download_button(label="Download Essencial", data=Essencial.to_csv(index=False), file_name="Essencial.csv", mime="text/csv")

if __name__ == "__main__":
    main()

