# 🎸 The Nehemizer (Portal Financeiro - Saavedra)

> *“O peso justo e a organização perfeita para os seus contratos financeiros.”* (Inspirado em Provérbios 11:1 e na precisão administrativa de Neemias).

## 🚀 Sobre o Projeto
**The Nehemizer** é uma aplicação web de alta performance desenvolvida em **Python (Streamlit + Pandas)** para automatizar o fluxo logístico e de compras da Saavedra. 

O sistema processa relatórios de vendas brutos, extrai de forma inteligente os preços e contratos vigentes de propostas comerciais em PDF (Becton Dickinson / BD), cruza os dados com tabelas de preço normal e gera planilhas consolidadas e formatadas em segundos, eliminando horas de trabalho manual.

## ✨ Principais Funcionalidades
- **Leitura Dinâmica de Vendas:** Identifica automaticamente o cabeçalho correto em planilhas Excel ou CSV.
- **Extração Inteligente de PDFs:** Varre propostas comerciais em PDF para capturar preços unitários e condições de contrato.
- **Regra Suprema de Exceção:** Produtos sem preço tabelado em contrato são automaticamente direcionados para a aba `NORMAL`.
- **Interface Editável (Data Editor):** Permite o ajuste manual de preços faltantes diretamente na prévia visual antes da exportação.
- **Geração Automatizada de Excel:** Cria abas individuais e uma aba `RESUMO` estruturada com subtotais acumulados, formatação de moeda e padrão visual corporativo.

## 🛠️ Tecnologias Utilizadas
- [Python](https://www.python.org/)
- [Streamlit](https://streamlit.io/)
- [Pandas](https://pandas.pydata.org/)
- [pdfplumber](https://github.com/jsvine/pdfplumber)
- [XlsxWriter](https://xlsxwriter.readthedocs.io/)