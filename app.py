import streamlit as st
import pandas as pd
import io
import numpy as np
import pdfplumber
from thefuzz import process

# Configuração da página
st.set_page_config(page_title="Portal Financeiro - Saavedra", page_icon="📊", layout="wide")

# --- 1. INTELIGÊNCIA DE AGRUPAMENTO E REGRAS ---
GRUPOS_MESTRES = [
    'UNIMED', 'CONCEICAO', 'PUC', 'SANTA CASA', 
    'EBSERH', 'HCAA', 'H DIVINA', 'SMS POA'
]

def classificar_cliente_inteligente(razao_social):
    razao = str(razao_social).upper()
    melhor_match, pontuacao = process.extractOne(razao, GRUPOS_MESTRES)
    if pontuacao >= 80:
        return melhor_match
    return razao.strip()

def extrair_precos_pdf(arquivos_pdf):
    dados_precos = []
    for arquivo in arquivos_pdf:
        with pdfplumber.open(arquivo) as pdf:
            # Identifica o cliente pelo texto da primeira página
            texto_pag1 = pdf.pages[0].extract_text().upper()
            grupo_cliente, pontuacao = process.extractOne(texto_pag1, GRUPOS_MESTRES)
            
            # Se a similaridade for baixa, marcamos para revisão
            if pontuacao < 70:
                grupo_cliente = "REVISAR"
                
            # Extrai as tabelas de todas as páginas
            for page in pdf.pages:
                tabelas = page.extract_tables()
                for tabela in tabelas:
                    for linha in tabela:
                        # Ignora linhas vazias ou muito curtas
                        if not linha or len(linha) < 6: continue
                        
                        celulas = [str(c).strip() if c else "" for c in linha]
                        ref_prod = celulas[0].replace('.', '').replace('-', '') 
                        
                        # Se a primeira coluna for um código (só números)
                        if ref_prod.isdigit() and len(ref_prod) > 4:
                            preco_str = ""
                            # O Preço Unitário costuma estar na coluna 6 ou 5 e contém "R$"
                            if len(celulas) > 6 and "R$" in celulas[6]:
                                preco_str = celulas[6]
                            elif "R$" in celulas[5]:
                                preco_str = celulas[5]
                            else:
                                # Varre a linha de trás pra frente buscando um valor monetário
                                for c in reversed(celulas):
                                    if "R$" in c:
                                        preco_str = c
                                        break
                                        
                            if preco_str:
                                # Limpa o R$, separadores de milhar e converte para float
                                valor_limpo = preco_str.replace("R$", "").replace(".", "").replace(",", ".").strip()
                                try:
                                    dados_precos.append({
                                        'GRUPO_CLIENTE': grupo_cliente,
                                        'REFPROD': ref_prod,
                                        'VALOR_TABELADO_BD': float(valor_limpo)
                                    })
                                except ValueError:
                                    pass
                                    
    return pd.DataFrame(dados_precos)

# --- 2. INTERFACE E UPLOADS ---
st.title("📊 Automação Inteligente Saavedra N3")
st.markdown("Faça o upload do seu relatório de vendas e das propostas da BD para gerar o relatório final consolidado.")

col1, col2 = st.columns(2)
with col1:
    arquivo_excel = st.file_uploader("1º Relatório de Vendas (Excel/CSV)", type=['xlsx', 'xls', 'csv'])
with col2:
    arquivos_pdf = st.file_uploader("2º Contratos e Propostas BD (PDF)", type=['pdf'], accept_multiple_files=True)

if arquivo_excel:
    st.info("🔄 Processando dados...")
    
    # Leitura dinâmica: procura a linha do cabeçalho
    df_temp = pd.read_excel(arquivo_excel, nrows=20, header=None)
    linha_cabecalho = 0
    for i, row in df_temp.iterrows():
        linha_texto = "".join(str(val).upper() for val in row.values)
        if "RAZAOSOCIAL" in linha_texto or "CONVENIO" in linha_texto:
            linha_cabecalho = i
            break
            
    df = pd.read_excel(arquivo_excel, header=linha_cabecalho)
    df.columns = df.columns.str.upper().str.strip()
    
    novas_colunas = {}
    for col in df.columns:
        if 'REF' in col and 'PROD' in col: novas_colunas[col] = 'REFPROD'
        elif 'DESC' in col: novas_colunas[col] = 'DESCRICAO'
        elif 'QTD' in col: novas_colunas[col] = 'QTDCOM'
        elif 'VLR' in col or 'VALOR' in col: novas_colunas[col] = 'VLRTOTAL'
        elif 'RAZ' in col and 'SOC' in col: novas_colunas[col] = 'RAZAOSOCIAL'
        elif 'CONV' in col: novas_colunas[col] = 'CONVENIO'
        
    df = df.rename(columns=novas_colunas)  
    df = df.loc[:, ~df.columns.duplicated()]
    
    # Converte REFPROD para string garantindo o match correto
    if 'REFPROD' in df.columns:
        df['REFPROD'] = df['REFPROD'].astype(str).str.replace(r'\.0$', '', regex=True)

    df['GRUPO_CLIENTE'] = df['RAZAOSOCIAL'].apply(classificar_cliente_inteligente)
    
    # --- PROCESSAMENTO DOS PDFs ---
    if arquivos_pdf:
        st.success(f"📄 {len(arquivos_pdf)} arquivo(s) PDF detectado(s). Lendo tabelas de preço...")
        df_precos = extrair_precos_pdf(arquivos_pdf)
        
        if not df_precos.empty:
            # Remove duplicatas caso o mesmo PDF seja lido duas vezes
            df_precos = df_precos.drop_duplicates(subset=['GRUPO_CLIENTE', 'REFPROD'])
            # O "PROCV" do Pandas: junta as vendas com os preços
            df = pd.merge(df, df_precos, on=['GRUPO_CLIENTE', 'REFPROD'], how='left')
        else:
            st.warning("⚠️ Não foi possível extrair preços legíveis dos PDFs enviados.")
    
    # --- 3. SELEÇÃO DE CLIENTES ---
    st.divider()
    grupos_unicos = sorted(df['GRUPO_CLIENTE'].dropna().unique().tolist())
    st.subheader("3º Quais grupos terão aba própria?")
    
    clientes_selecionados = st.multiselect("Selecione os clientes:", grupos_unicos, default=grupos_unicos)

    # --- 4. EXPORTAÇÃO ---
    if st.button("GERAR RELATÓRIO FINAL 📥", type="primary") and clientes_selecionados:
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})
            formato_cabecalho = workbook.add_format({'bold': True, 'bg_color': '#333333', 'font_color': 'white'})
            formato_destaque_bd = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'font_color': '#375623', 'num_format': 'R$ #,##0.00'})
            
            df_normal = df[~df['GRUPO_CLIENTE'].isin(clientes_selecionados)]
            abas_para_criar = clientes_selecionados.copy()
            if not df_normal.empty:
                abas_para_criar.append('NORMAL')
            
            # Aba RESUMO
            # Define o que vamos agregar
            agg_dict = {'QTDCOM': 'sum', 'VLRTOTAL': 'sum'}
            if 'VALOR_TABELADO_BD' in df.columns:
                agg_dict['VALOR_TABELADO_BD'] = 'first' # Traz o valor lido no PDF
                
            resumo_df = df.groupby(['GRUPO_CLIENTE', 'REFPROD', 'DESCRICAO']).agg(agg_dict).reset_index()
            resumo_df['VLR UNIT CALCULADO'] = np.where(resumo_df['QTDCOM'] > 0, resumo_df['VLRTOTAL'] / resumo_df['QTDCOM'], 0)
            
            # Reorganiza colunas para ficar visualmente lógico
            if 'VALOR_TABELADO_BD' in resumo_df.columns:
                cols = ['GRUPO_CLIENTE', 'REFPROD', 'DESCRICAO', 'QTDCOM', 'VLR UNIT CALCULADO', 'VALOR_TABELADO_BD', 'VLRTOTAL']
                resumo_df = resumo_df[cols]
            
            resumo_df.to_excel(writer, sheet_name='RESUMO', index=False)
            worksheet_resumo = writer.sheets['RESUMO']
            
            # Aplica larguras e formatos
            worksheet_resumo.set_column('A:A', 25)
            worksheet_resumo.set_column('B:B', 15)
            worksheet_resumo.set_column('C:C', 40)
            worksheet_resumo.set_column('D:D', 10)
            worksheet_resumo.set_column('E:E', 20, formato_moeda)
            
            if 'VALOR_TABELADO_BD' in resumo_df.columns:
                worksheet_resumo.set_column('F:F', 20, formato_destaque_bd) # Destaca a coluna do PDF
                worksheet_resumo.set_column('G:G', 15, formato_moeda)
            else:
                worksheet_resumo.set_column('F:F', 15, formato_moeda)
                
            for col_num, value in enumerate(resumo_df.columns.values):
                worksheet_resumo.write(0, col_num, value, formato_cabecalho)
            
            # Abas INDIVIDUAIS
            for cli in abas_para_criar:
                nome_aba = str(cli)[:31].replace(':', '').replace('/', '') 
                df_cli = df_normal if cli == 'NORMAL' else df[df['GRUPO_CLIENTE'] == cli]
                
                df_cli.to_excel(writer, sheet_name=nome_aba, index=False)
                worksheet_cli = writer.sheets[nome_aba]
                for col_num, value in enumerate(df_cli.columns.values):
                    worksheet_cli.write(0, col_num, value, formato_cabecalho)
        
        st.balloons()
        st.success("Tudo pronto! Seu relatório foi processado.")
        st.download_button(
            label="📄 Baixar Planilha Consolidada",
            data=output.getvalue(),
            file_name="PROCESSADO_Relatorio_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )