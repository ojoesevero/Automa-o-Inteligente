import streamlit as st
import pandas as pd
import io
import numpy as np
import pdfplumber

# Configuração da página
st.set_page_config(page_title="Portal Financeiro - Saavedra", page_icon="📊", layout="wide")

# --- 1. INTELIGÊNCIA DE AGRUPAMENTO E REGRAS ---
def agrupar_cliente(texto, fallback=None):
    r = str(texto).upper()
    if 'UNIMED' in r or '87096616' in r: return 'UNIMED'
    if 'CONCEICAO' in r or 'CONCEIÇÃO' in r: return 'CONCEICAO'
    if 'UNIAO BRASILEIRA' in r or 'PUC' in r or '88630413' in r: return 'PUC'
    if 'SANTA CASA' in r: return 'SANTA CASA'
    if 'EBSERH' in r or 'SERVICOS HOSPITALARES' in r or '15126437' in r: return 'EBSERH'
    if 'ASTROGILDO' in r or '95610887' in r: return 'HCAA'
    if 'DIVINA' in r or '87317764' in r: return 'H DIVINA'
    if 'PORTO ALEGRE' in r and ('PREF' in r or 'MUNICIPIO' in r or '92963560' in r): return 'SMS POA'
    if 'CLINICAS' in r or '87020517' in r: return 'HCPA'
    if 'GHC' in r or '450166419' in r: return 'GHC'
    return fallback if fallback is not None else r.strip()

def extrair_precos_pdf(arquivos_pdf):
    dados_precos = []
    for arquivo in arquivos_pdf:
        with pdfplumber.open(arquivo) as pdf:
            texto_pag1 = pdf.pages[0].extract_text().upper()
            grupo_cliente = agrupar_cliente(texto_pag1, fallback="REVISAR")
            
            for page in pdf.pages:
                tabelas = page.extract_tables()
                for tabela in tabelas:
                    for linha in tabela:
                        if not linha or len(linha) < 4: continue
                        
                        celulas = [str(c).strip() if c else "" for c in linha]
                        ref_prod = celulas[0].replace('.', '').replace('-', '') 
                        
                        if ref_prod.isalnum() and len(ref_prod) >= 5:
                            preco_str = next((c for c in celulas if "R$" in c), "")
                            if preco_str:
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
    
    with st.spinner('Lendo e cruzando os dados... Isso pode levar alguns segundos.'):
        
        # Lê o Excel
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
        
        if 'REFPROD' in df.columns:
            df['REFPROD'] = df['REFPROD'].astype(str).str.replace(r'\.0$', '', regex=True)

        df['GRUPO_CLIENTE'] = df['RAZAOSOCIAL'].apply(lambda x: agrupar_cliente(x))
        
        # Lê os PDFs e mescla
        if arquivos_pdf:
            df_precos = extrair_precos_pdf(arquivos_pdf)
            if not df_precos.empty:
                df_precos = df_precos.drop_duplicates(subset=['GRUPO_CLIENTE', 'REFPROD'])
                df = pd.merge(df, df_precos, on=['GRUPO_CLIENTE', 'REFPROD'], how='left')

        # Garante que a coluna de preços exista, mesmo se nenhum PDF for enviado ou faltar dado
        if 'VALOR_TABELADO_BD' not in df.columns:
            df['VALOR_TABELADO_BD'] = None

        grupos_unicos = sorted(df['GRUPO_CLIENTE'].dropna().unique().tolist())
    
    st.toast('Processamento concluído!', icon='✅')
    st.success("Tudo lido com sucesso! Você pode editar os preços em branco diretamente na tabela abaixo.")

    st.divider()
    
    # --- 3. SELEÇÃO E PRÉVIA EDITÁVEL ---
    st.subheader("3º Quais grupos terão aba própria?")
    clientes_selecionados = st.multiselect("Selecione os clientes (ou deixe todos):", grupos_unicos, default=grupos_unicos)

    if clientes_selecionados:
        
        df_normal = df[~df['GRUPO_CLIENTE'].isin(clientes_selecionados)]
        abas_para_criar = clientes_selecionados.copy()
        if not df_normal.empty:
            abas_para_criar.append('NORMAL')
            
        agg_dict = {'QTDCOM': 'sum', 'VLRTOTAL': 'sum', 'VALOR_TABELADO_BD': 'first'}
            
        resumo_df = df.groupby(['GRUPO_CLIENTE', 'REFPROD', 'DESCRICAO']).agg(agg_dict).reset_index()
        resumo_df['VLR UNIT CALCULADO'] = np.where(resumo_df['QTDCOM'] > 0, resumo_df['VLRTOTAL'] / resumo_df['QTDCOM'], 0)
        
        cols = ['GRUPO_CLIENTE', 'REFPROD', 'DESCRICAO', 'QTDCOM', 'VLR UNIT CALCULADO', 'VALOR_TABELADO_BD', 'VLRTOTAL']
        resumo_df = resumo_df[cols]
        
        # 👁️ EXIBE A PRÉVIA VISUAL EDITÁVEL
        st.subheader("👁️ Prévia do Relatório (Aba RESUMO)")
        st.markdown("✏️ **Dica:** Dê um **duplo-clique** na coluna `VALOR_TABELADO_BD` para preencher os valores vazios manualmente. As outras colunas estão travadas por segurança.")
        
        colunas_bloqueadas = ['GRUPO_CLIENTE', 'REFPROD', 'DESCRICAO', 'QTDCOM', 'VLR UNIT CALCULADO', 'VLRTOTAL']
        
        resumo_df = st.data_editor(
            resumo_df, 
            use_container_width=True, 
            hide_index=True,
            disabled=colunas_bloqueadas
        )
        
        # 🔄 REPASSA A EDIÇÃO PARA O DATAFRAME PRINCIPAL (Para as abas individuais)
        # Transforma a coluna editada em um dicionário para busca rápida
        precos_editados = resumo_df.set_index(['GRUPO_CLIENTE', 'REFPROD'])['VALOR_TABELADO_BD'].to_dict()
        
        # Atualiza o df principal com os valores que o Marcio digitou na tela
        df['VALOR_TABELADO_BD'] = df.apply(
            lambda row: precos_editados.get((row['GRUPO_CLIENTE'], row['REFPROD']), row['VALOR_TABELADO_BD']), 
            axis=1
        )
        
        # --- 4. GERAÇÃO DO ARQUIVO PARA DOWNLOAD ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})
            formato_cabecalho = workbook.add_format({'bold': True, 'bg_color': '#333333', 'font_color': 'white'})
            formato_destaque_bd = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'font_color': '#375623', 'num_format': 'R$ #,##0.00'})
            
            resumo_df.to_excel(writer, sheet_name='RESUMO', index=False)
            worksheet_resumo = writer.sheets['RESUMO']
            
            worksheet_resumo.set_column('A:A', 25)
            worksheet_resumo.set_column('B:B', 15)
            worksheet_resumo.set_column('C:C', 40)
            worksheet_resumo.set_column('D:D', 10)
            worksheet_resumo.set_column('E:E', 20, formato_moeda)
            worksheet_resumo.set_column('F:F', 20, formato_destaque_bd)
            worksheet_resumo.set_column('G:G', 15, formato_moeda)
                
            for col_num, value in enumerate(resumo_df.columns.values):
                worksheet_resumo.write(0, col_num, value, formato_cabecalho)
            
            for cli in abas_para_criar:
                nome_aba = str(cli)[:31].replace(':', '').replace('/', '') 
                df_cli = df_normal if cli == 'NORMAL' else df[df['GRUPO_CLIENTE'] == cli]
                
                df_cli.to_excel(writer, sheet_name=nome_aba, index=False)
                worksheet_cli = writer.sheets[nome_aba]
                for col_num, value in enumerate(df_cli.columns.values):
                    worksheet_cli.write(0, col_num, value, formato_cabecalho)
        
        st.divider()
        st.download_button(
            label="📥 Baixar Planilha Consolidada",
            data=output.getvalue(),
            file_name="PROCESSADO_Relatorio_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )