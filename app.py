import streamlit as st
import pandas as pd
import io
import numpy as np
import pdfplumber

st.set_page_config(page_title="Portal Financeiro - Saavedra", page_icon="📊", layout="wide")

# --- REGRAS E MAPEAMENTO ---
CONTRATOS_MAPPING = {
    'UNIMED': '450128261',
    'PUC': '450155842',
    'H DIVINA': '450146829',
    'HCPA': '450139832',
    'SMS POA': '450120243',
    'EBSERH': '450098658',
    'HCAA': '450137626',
    'GHC': '450166419',
    'CONCEICAO': '450166419'
}

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

# --- INTERFACE ---
st.title("📊 Automação Inteligente Saavedra N3")

col1, col2, col3 = st.columns(3)
with col1:
    arquivo_excel = st.file_uploader("1º Relatório de Vendas", type=['xlsx', 'xls', 'csv'])
with col2:
    arquivos_pdf = st.file_uploader("2º Contratos BD (PDFs)", type=['pdf'], accept_multiple_files=True)
with col3:
    arquivo_normal = st.file_uploader("3º Tabela Preços Normal", type=['xlsx', 'xls', 'csv'])

if arquivo_excel:
    with st.spinner('Lendo e aplicando regras de negócio...'):
        
        # 1. LEITURA VENDAS
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
            c_up = str(col).upper().strip()
            if 'REF' in c_up and 'PROD' in c_up: novas_colunas[col] = 'REFPROD'
            elif 'DESC' in c_up: novas_colunas[col] = 'DESCRICAO'
            elif 'QTD' in c_up: novas_colunas[col] = 'QTDCOM'
            elif 'VLR' in c_up or 'VALOR' in c_up: novas_colunas[col] = 'VLRTOTAL'
            elif 'RAZ' in c_up and 'SOC' in c_up: novas_colunas[col] = 'RAZAOSOCIAL'
            elif 'CONV' in c_up: novas_colunas[col] = 'CONVENIO'
        df = df.rename(columns=novas_colunas)  
        df = df.loc[:, ~df.columns.duplicated()]
        
        # Limpeza de códigos e CNPJs
        if 'REFPROD' in df.columns:
            df['REFPROD'] = df['REFPROD'].astype(str).str.replace(r'\.0$', '', regex=True)
        if 'CNPJPARCEIRO' in df.columns:
            df['CNPJPARCEIRO'] = df['CNPJPARCEIRO'].astype(str).str.replace(r'\D', '', regex=True)
        if 'CNPJPARCUSO' in df.columns:
            df['CNPJPARCUSO'] = df['CNPJPARCUSO'].astype(str).str.replace(r'\D', '', regex=True)

        df['GRUPO_CLIENTE'] = df['RAZAOSOCIAL'].apply(lambda x: agrupar_cliente(x))
        df['CONTRATO_ORIGINAL'] = df['CONTRATOCLIENTE'] if 'CONTRATOCLIENTE' in df.columns else ''
        
        # 2. LEITURA PDFs
        if arquivos_pdf:
            df_precos = extrair_precos_pdf(arquivos_pdf)
            if not df_precos.empty:
                df_precos = df_precos.drop_duplicates(subset=['GRUPO_CLIENTE', 'REFPROD'])
                df = pd.merge(df, df_precos, on=['GRUPO_CLIENTE', 'REFPROD'], how='left')
        
        if 'VALOR_TABELADO_BD' not in df.columns:
            df['VALOR_TABELADO_BD'] = np.nan
            
        # 3. LEITURA TABELA NORMAL
        if arquivo_normal:
            df_norm = pd.read_excel(arquivo_normal)
            norm_cols = {}
            for col in df_norm.columns:
                c_up = str(col).upper().strip()
                if 'REF' in c_up and 'PROD' in c_up: norm_cols[col] = 'REFPROD'
                elif 'VLR' in c_up or 'VALOR' in c_up or 'PRECO' in c_up or 'PREÇO' in c_up: norm_cols[col] = 'PRECO_NORMAL'
            df_norm = df_norm.rename(columns=norm_cols)
            
            if 'REFPROD' in df_norm.columns and 'PRECO_NORMAL' in df_norm.columns:
                df_norm['REFPROD'] = df_norm['REFPROD'].astype(str).str.replace(r'\.0$', '', regex=True)
                df_norm = df_norm.drop_duplicates(subset=['REFPROD'])
                df = pd.merge(df, df_norm[['REFPROD', 'PRECO_NORMAL']], on='REFPROD', how='left')
            else:
                df['PRECO_NORMAL'] = np.nan
        else:
            df['PRECO_NORMAL'] = np.nan
            
        # 4. APLICAÇÃO DA REGRA DE NEGÓCIO SUPREMA
        df['TEM_PRECO_BD'] = df['VALOR_TABELADO_BD'].notna()
        
        # Se não tiver preço de BD, destina para a aba NORMAL
        df['ABA_DESTINO'] = np.where(df['TEM_PRECO_BD'], df['GRUPO_CLIENTE'], 'NORMAL')
        df['SIGLA_RESUMO'] = np.where(df['TEM_PRECO_BD'], df['GRUPO_CLIENTE'], 'NORMAL')
        
        # Adiciona o número do contrato. Se não for contrato (ou for NORMAL), fica 'NORMAL'
        df['CONTRATO_BD_NUM'] = df['GRUPO_CLIENTE'].map(CONTRATOS_MAPPING).fillna('NORMAL')
        df['CONTRATO_FINAL'] = np.where(df['TEM_PRECO_BD'], df['CONTRATO_BD_NUM'], 'NORMAL')
        
        # O preço final puxa da BD, se for vazio puxa do Upload Normal
        df['PRECO_COMPRA_FINAL'] = df['VALOR_TABELADO_BD'].fillna(df['PRECO_NORMAL'])

    st.toast('Processamento concluído!', icon='✅')
    st.divider()

    # --- PRÉVIA EDITÁVEL ---
    st.subheader("👁️ Prévia do Relatório (Aba RESUMO)")
    st.markdown("✏️ **Dica:** Dê um duplo-clique na coluna `PRECO_COMPRA_FINAL` para preencher valores em branco.")

    agg_dict = {
        'QTDCOM': 'sum',
        'VLRTOTAL': 'sum',
        'PRECO_COMPRA_FINAL': 'first',
        'CONTRATO_FINAL': 'first',
        'SIGLA_RESUMO': 'first'
    }
    
    resumo_df = df.groupby(['ABA_DESTINO', 'GRUPO_CLIENTE', 'REFPROD', 'DESCRICAO']).agg(agg_dict).reset_index()
    resumo_df['VLR UNIT VENDA'] = np.where(resumo_df['QTDCOM'] > 0, resumo_df['VLRTOTAL'] / resumo_df['QTDCOM'], 0)
    
    cols_ui = ['CONTRATO_FINAL', 'SIGLA_RESUMO', 'GRUPO_CLIENTE', 'REFPROD', 'DESCRICAO', 'QTDCOM', 'VLR UNIT VENDA', 'VLRTOTAL', 'PRECO_COMPRA_FINAL']
    resumo_df_ui = resumo_df[cols_ui].copy()
    
    colunas_bloqueadas = ['CONTRATO_FINAL', 'SIGLA_RESUMO', 'GRUPO_CLIENTE', 'REFPROD', 'DESCRICAO', 'QTDCOM', 'VLR UNIT VENDA', 'VLRTOTAL']
    
    resumo_df_ui = st.data_editor(
        resumo_df_ui, 
        use_container_width=True, 
        hide_index=True,
        disabled=colunas_bloqueadas,
        column_config={
            "VLR UNIT VENDA": st.column_config.NumberColumn(format="R$ %.2f"),
            "VLRTOTAL": st.column_config.NumberColumn(format="R$ %.2f"),
            "PRECO_COMPRA_FINAL": st.column_config.NumberColumn(format="R$ %.2f"),
        }
    )
    
    # Atualiza o DataFrame principal com edições manuais
    precos_editados = resumo_df_ui.set_index(['GRUPO_CLIENTE', 'REFPROD'])['PRECO_COMPRA_FINAL'].to_dict()
    df['PRECO_COMPRA_FINAL'] = df.apply(
        lambda row: precos_editados.get((row['GRUPO_CLIENTE'], row['REFPROD']), row['PRECO_COMPRA_FINAL']), 
        axis=1
    )
    
    # --- GERAÇÃO DO ARQUIVO PARA DOWNLOAD ---
    if st.button("GERAR RELATÓRIO FINAL 📥", type="primary"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})
            formato_cabecalho = workbook.add_format({'bold': True, 'bg_color': '#F37021', 'font_color': 'white'}) # Laranja Saavedra
            formato_subtotal = workbook.add_format({'bold': True})
            formato_cabecalho_aba = workbook.add_format({'bold': True, 'bg_color': '#333333', 'font_color': 'white'})
            
            # 1. ABA RESUMO
            df_resumo_export = df.groupby(['ABA_DESTINO', 'GRUPO_CLIENTE', 'REFPROD', 'DESCRICAO']).agg({
                'CONTRATO_FINAL': 'first',
                'SIGLA_RESUMO': 'first',
                'QTDCOM': 'sum',
                'VLRTOTAL': 'sum',
                'PRECO_COMPRA_FINAL': 'first'
            }).reset_index()
            
            df_resumo_export = df_resumo_export.sort_values(by=['ABA_DESTINO', 'GRUPO_CLIENTE', 'REFPROD'])
            linhas_resumo = []
            
            # Cria os blocos de subtotal agrupados pelo destino
            for destino, group in df_resumo_export.groupby('ABA_DESTINO', sort=False):
                subtotal_venda = 0
                subtotal_compra = 0
                
                for _, row in group.iterrows():
                    vlr_unit_venda = row['VLRTOTAL'] / row['QTDCOM'] if row['QTDCOM'] > 0 else 0
                    vlr_total_compra = row['QTDCOM'] * (row['PRECO_COMPRA_FINAL'] if pd.notna(row['PRECO_COMPRA_FINAL']) else 0)
                    
                    linhas_resumo.append([
                        row['CONTRATO_FINAL'], 
                        row['SIGLA_RESUMO'], 
                        row['GRUPO_CLIENTE'], 
                        row['REFPROD'], 
                        row['DESCRICAO'], 
                        row['QTDCOM'], 
                        vlr_unit_venda, 
                        row['VLRTOTAL'], 
                        row['PRECO_COMPRA_FINAL'] if pd.notna(row['PRECO_COMPRA_FINAL']) else None, 
                        vlr_total_compra if vlr_total_compra > 0 else None
                    ])
                    subtotal_venda += row['VLRTOTAL']
                    subtotal_compra += vlr_total_compra
                
                # Linha de subtotal
                linhas_resumo.append([
                    '', '', f'TOTAL ACUMULADO {destino}', '', '', '', '', subtotal_venda, '', subtotal_compra if subtotal_compra > 0 else None
                ])
                
            df_excel_resumo = pd.DataFrame(linhas_resumo, columns=[
                'CONTRATO', 'SIGLA/GRUPO', 'PEDIDO', 'Ref Prod', 'Descrição', 'Qtd Com', 'VLR UNIT', 'VLR TOTAL', 'VLR UNIT COMPRA', 'VLR TOTAL COMPRA'
            ])
            
            df_excel_resumo.to_excel(writer, sheet_name='RESUMO', index=False)
            ws_resumo = writer.sheets['RESUMO']
            ws_resumo.set_column('A:B', 15)
            ws_resumo.set_column('C:E', 35)
            ws_resumo.set_column('F:F', 10)
            ws_resumo.set_column('G:J', 18, formato_moeda)
            
            for col_num, value in enumerate(df_excel_resumo.columns.values):
                ws_resumo.write(0, col_num, value, formato_cabecalho)
            
            # Negrito para os Totais Acumulados
            for row_num, row_data in enumerate(linhas_resumo):
                if 'TOTAL ACUMULADO' in str(row_data[2]):
                    ws_resumo.set_row(row_num + 1, None, formato_subtotal)
            
            # 2. ABAS INDIVIDUAIS
            abas_para_criar = df['ABA_DESTINO'].dropna().unique().tolist()
            
            for aba in abas_para_criar:
                df_aba = df[df['ABA_DESTINO'] == aba].copy()
                
                colunas_aba = []
                for _, r in df_aba.iterrows():
                    # Mantém colunas estritas da versão anterior
                    contrato_str = str(r.get('CONTRATO_ORIGINAL', ''))
                    cnpj_p = r.get('CNPJPARCEIRO', '')
                    uf = r.get('UF', '')
                    cod_cli = r.get('CODCLIUSO', '')
                    razao = r.get('RAZAOSOCIAL', '')
                    cnpj_u = r.get('CNPJPARCUSO', '')
                    conv = r.get('CONVENIO', '')
                    ref = r.get('REFPROD', '')
                    desc = r.get('DESCRICAO', '')
                    qtd = r.get('QTDCOM', '')
                    colunas_aba.append([contrato_str, f"'{cnpj_p}", uf, cod_cli, razao, f"'{cnpj_u}", conv, ref, desc, qtd])
                
                df_excel_aba = pd.DataFrame(colunas_aba, columns=[
                    'CONTRATO', 'CNPJ_PARCEIRO', 'UF', 'COD_CLI_USO', 'RAZÃO SOCIAL', 'CNPJ_PARC_USO', 'CONVENIO', 'Ref Prod', 'Descrição', 'Qtd'
                ])
                
                nome_sheet = str(aba)[:31].replace(':', '').replace('/', '')
                df_excel_aba.to_excel(writer, sheet_name=nome_sheet, index=False)
                ws_aba = writer.sheets[nome_sheet]
                
                for col_num, value in enumerate(df_excel_aba.columns.values):
                    ws_aba.write(0, col_num, value, formato_cabecalho_aba)
                
                # Gera o bloco de resumo de subtotais no final de cada aba
                start_r = len(df_excel_aba) + 4
                ws_aba.write_row(start_r - 1, 0, ['Ref Prod', 'Descrição', 'Qtd Com', 'VLR UNIT VENDA', 'VLR TOTAL VENDA'], formato_subtotal)
                
                resumo_cli = df_aba.groupby(['REFPROD', 'DESCRICAO']).agg({'QTDCOM': 'sum', 'VLRTOTAL': 'sum'}).reset_index()
                
                for idx, rcli in resumo_cli.iterrows():
                    v_unit = rcli['VLRTOTAL'] / rcli['QTDCOM'] if rcli['QTDCOM'] > 0 else 0
                    ws_aba.write_row(start_r + idx, 0, [rcli['REFPROD'], rcli['DESCRICAO'], rcli['QTDCOM'], v_unit, rcli['VLRTOTAL']])
                
                total_geral = resumo_cli['VLRTOTAL'].sum()
                ws_aba.write_row(start_r + len(resumo_cli), 0, ['TOTAL GERAL', '', '', '', total_geral], formato_subtotal)
                ws_aba.set_column('D:E', 15, formato_moeda)

        st.balloons()
        st.download_button(
            label="📄 Baixar Planilha Consolidada",
            data=output.getvalue(),
            file_name="PROCESSADO_Relatorio_Final_Definitivo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )