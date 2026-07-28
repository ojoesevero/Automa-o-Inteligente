import streamlit as st
import pandas as pd
import io
import numpy as np
from thefuzz import process

# Configuração da página
st.set_page_config(page_title="Portal Financeiro - Saavedra", page_icon="📊")

# --- 1. INTELIGÊNCIA DE AGRUPAMENTO (Fuzzy Matching) ---
# Em vez de IFs, definimos os grupos "Mestres". 
# O sistema vai tentar encaixar nomes parecidos neles.
GRUPOS_MESTRES = [
    'UNIMED', 'CONCEICAO', 'PUC', 'SANTA CASA', 
    'EBSERH', 'HCAA', 'H DIVINA', 'SMS POA'
]

def classificar_cliente_inteligente(razao_social):
    razao = str(razao_social).upper()
    # Pega o grupo mais parecido e a pontuação de similaridade (0 a 100)
    melhor_match, pontuacao = process.extractOne(razao, GRUPOS_MESTRES)
    
    # Se for mais de 80% parecido, assume que é o grupo. Senão, mantém o nome original.
    if pontuacao >= 80:
        return melhor_match
    return razao.strip()

# --- 2. INTERFACE E UPLOAD ---
st.title("📊 Automação Inteligente Saavedra N3")
arquivo = st.file_uploader("1º Carregue o arquivo (Excel ou CSV)", type=['xlsx', 'xls', 'csv'])

if arquivo:
    st.info("Lendo arquivo e aplicando inteligência de agrupamento...")
    
    # Leitura dinâmica: procura a linha do cabeçalho
    # Lê as primeiras 20 linhas para encontrar onde estão as colunas certas
    df_temp = pd.read_excel(arquivo, nrows=20, header=None)
    linha_cabecalho = 0
    for i, row in df_temp.iterrows():
        linha_texto = "".join(str(val).upper() for val in row.values)
        if "RAZAOSOCIAL" in linha_texto or "CONVENIO" in linha_texto:
            linha_cabecalho = i
            break
            
    # Carrega o dataframe real a partir do cabeçalho correto
    df = pd.read_excel(arquivo, header=linha_cabecalho)
    
    # 1. Primeiro padronizamos tudo para maiúsculo e tiramos espaços das pontas
    df.columns = df.columns.str.upper().str.strip()
    
    # 2. Depois forçamos a renomeação do que importa, procurando palavras-chave
    novas_colunas = {}
    for col in df.columns:
        if 'REF' in col and 'PROD' in col: novas_colunas[col] = 'REFPROD'
        elif 'DESC' in col: novas_colunas[col] = 'DESCRICAO'
        elif 'QTD' in col: novas_colunas[col] = 'QTDCOM'
        elif 'VLR' in col or 'VALOR' in col: novas_colunas[col] = 'VLRTOTAL'
        elif 'RAZ' in col and 'SOC' in col: novas_colunas[col] = 'RAZAOSOCIAL'
        elif 'CONV' in col: novas_colunas[col] = 'CONVENIO'
        
    df = df.rename(columns=novas_colunas)  
    # Aplica a Inteligência (Gera a coluna GRUPO_CLIENTE)
    df['GRUPO_CLIENTE'] = df['RAZAOSOCIAL'].apply(classificar_cliente_inteligente)
    
    # --- 3. SELEÇÃO DE CLIENTES ---
    grupos_unicos = sorted(df['GRUPO_CLIENTE'].dropna().unique().tolist())
    st.subheader("2º Quais grupos terão aba própria?")
    
    clientes_selecionados = st.multiselect("Selecione os clientes:", grupos_unicos)

    # Adicione esta linha para debugar:
    st.write("Colunas identificadas na planilha:", df.columns.tolist())
    # --- 4. PROCESSAMENTO E EXPORTAÇÃO ---
    if st.button("Gerar Relatório Final", type="primary") and clientes_selecionados:
        
        # Cria um buffer na memória RAM (Não salva nada no HD)
        output = io.BytesIO()
        
        # Engine xlsxwriter permite colorir abas e células facilmente
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Formatos de Excel
            formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})
            formato_cabecalho = workbook.add_format({'bold': True, 'bg_color': '#333333', 'font_color': 'white'})
            
            # Divide os dados: O que é aba própria e o que vai para aba "NORMAL"
            df_normal = df[~df['GRUPO_CLIENTE'].isin(clientes_selecionados)]
            
            abas_para_criar = clientes_selecionados.copy()
            if not df_normal.empty:
                abas_para_criar.append('NORMAL')
            
            # Cria a aba RESUMO
            # (Aqui você faria o df.groupby() somando QTDCOM e VLRTOTAL)
            resumo_df = df.groupby(['GRUPO_CLIENTE', 'REFPROD', 'DESCRICAO']).agg(
                QTDCOM=('QTDCOM', 'sum'),
                VLRTOTAL=('VLRTOTAL', 'sum')
            ).reset_index()
            
            # Proteção contra divisão por zero
            resumo_df['VLR UNIT'] = np.where(resumo_df['QTDCOM'] > 0, resumo_df['VLRTOTAL'] / resumo_df['QTDCOM'], 0)
            
            # Salva aba Resumo
            resumo_df.to_excel(writer, sheet_name='RESUMO', index=False)
            worksheet_resumo = writer.sheets['RESUMO']
            worksheet_resumo.set_column('E:F', 15, formato_moeda) # Aplica moeda nas colunas de valor
            
            # Cria as abas individuais
            for cli in abas_para_criar:
                nome_aba = str(cli)[:31].replace(':', '').replace('/', '') # Regras do Excel
                
                if cli == 'NORMAL':
                    df_cli = df_normal
                else:
                    df_cli = df[df['GRUPO_CLIENTE'] == cli]
                
                df_cli.to_excel(writer, sheet_name=nome_aba, index=False)
                
                # Aplica estilos básicos
                worksheet_cli = writer.sheets[nome_aba]
                for col_num, value in enumerate(df_cli.columns.values):
                    worksheet_cli.write(0, col_num, value, formato_cabecalho)
        
        # Prepara o botão de download com o arquivo gerado na memória
        dados_excel = output.getvalue()
        
        st.success("Relatório gerado com sucesso!")
        st.download_button(
            label="📥 Baixar Planilha Processada",
            data=dados_excel,
            file_name="PROCESSADO_Relatorio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )