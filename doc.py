# -*- coding: utf-8 -*-
import streamlit as st
import requests
from docx import Document
import re
from datetime import datetime
import os
import io

# ========= CONFIGURAÇÃO =========
# Nome do seu arquivo modelo (já editado com tags)
# Usa caminho absoluto baseado no diretório do script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MODELO_PATH = os.path.join(SCRIPT_DIR, "Documento Contrato Serviço - Modelo.docx")

# API de CNPJ (gratuita)
API_CNPJ = "https://brasilapi.com.br/api/cnpj/v1/"

# ========= FUNÇÕES =========

def numero_para_extenso(valor):
    """
    Retorna o valor em formato extenso simples.
    Exemplo: 3400 -> 'três mil e quatrocentos reais'
    (Simplificado; para uso oficial, use num2words)
    """
    try:
        from num2words import num2words
        return num2words(valor, lang='pt_BR') + " reais"
    except ImportError:
        return f"{valor:.2f} reais"

def preencher_contrato(cnpj, valor, data_inicio, local_execucao, funcoes, observacoes):
    """
    Gera um contrato preenchido automaticamente com base em um modelo DOCX.
    Retorna o documento em bytes para download.
    """

    # --- 1️⃣ Buscar dados do CNPJ na API ---
    cnpj = re.sub(r'\D', '', cnpj)
    
    try:
        with st.spinner(f"Consultando dados do CNPJ {cnpj}..."):
            url = f"{API_CNPJ}{cnpj}"
            st.info(f"🔍 Consultando: {url}")
            
            r = requests.get(url, timeout=10)
            
            st.info(f"📡 Status da resposta: {r.status_code}")
            
            if r.status_code != 200:
                st.error(f"❌ Erro ao consultar CNPJ (Status {r.status_code})")
                st.code(r.text[:500])
                return None, None
            
            dados = r.json()
            st.success(f"✅ Dados encontrados: {dados.get('razao_social', 'N/A')}")
            
    except requests.exceptions.Timeout:
        st.error("❌ Tempo esgotado ao consultar API. Tente novamente.")
        return None, None
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Erro na requisição: {str(e)}")
        return None, None
    except Exception as e:
        st.error(f"❌ Erro inesperado: {str(e)}")
        return None, None

    # --- 2️⃣ Extrair campos ---
    nome_cliente = dados.get("razao_social", "")
    if not nome_cliente:
        st.warning("⚠️ Razão social não encontrada nos dados")
        
    logradouro = dados.get('logradouro', '')
    numero = dados.get('numero', '')
    bairro = dados.get('bairro', '')
    municipio = dados.get('municipio', '')
    uf = dados.get('uf', '')
    cep = dados.get('cep', '')
    
    endereco_cliente = f"{logradouro}, {numero}, {bairro} - {municipio}/{uf}, CEP {cep}"

    valor_extenso = numero_para_extenso(float(valor))

    # --- 3️⃣ Abrir o modelo e substituir tags ---
    # Verificar se o arquivo existe
    if not os.path.exists(MODELO_PATH):
        st.error(f"❌ Erro: Arquivo modelo não encontrado em: {MODELO_PATH}")
        st.info(f"Diretório atual: {os.getcwd()}")
        st.info(f"Arquivos disponíveis: {os.listdir(SCRIPT_DIR)}")
        return None, None
    
    doc = Document(MODELO_PATH)
    substituicoes = {
        "{{nome_cliente}}": nome_cliente,
        "{{cnpj_cliente}}": cnpj,
        "{{endereco_cliente}}": endereco_cliente,
        "{{valor_num}}": f"R$ {valor}",
        "{{valor_extenso}}": valor_extenso.capitalize(),
        "{{data_inicio}}": data_inicio,
        "{{local_execucao}}": local_execucao,
        "{{funcoes}}": funcoes,
        "{{observacoes}}": observacoes
    }

    for p in doc.paragraphs:
        for tag, valor in substituicoes.items():
            if tag in p.text:
                p.text = p.text.replace(tag, str(valor))

    # --- 4️⃣ Salvar o contrato em memória ---
    nome_arquivo = f"contrato_{nome_cliente[:20].strip().replace(' ', '_')}.docx"
    
    # Salvar em bytes para download
    doc_bytes = io.BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    
    return doc_bytes.getvalue(), nome_arquivo

# ========= INTERFACE STREAMLIT =========
def main():
    st.set_page_config(page_title="Gerador de Contratos", page_icon="📄", layout="wide")
    
    st.title("📄 Gerador Automático de Contratos")
    st.markdown("Preencha os dados abaixo para gerar um contrato personalizado automaticamente.")
    
    with st.form("contrato_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            cnpj = st.text_input("CNPJ do Cliente", value="65035552000180", 
                                help="Digite apenas números ou com formatação")
            valor = st.text_input("Valor do Contrato (R$)", value="3400.00")
            data_inicio = st.text_input("Data de Início", value="03/11/2025")
        
        with col2:
            local_execucao = st.text_area("Local de Execução", 
                                         value="Rua Joaquim Murtinho, 225, Bom Retiro - São Paulo/SP",
                                         height=100)
        
        funcoes = st.text_area("Funções e Horários", 
                              value="Supervisora Operacional / Encarregada – 8h, 4 Auxiliares de Limpeza – 8h",
                              height=100)
        
        observacoes = st.text_area("Observações", 
                                  value="Serviços de limpeza geral realizados bimestralmente aos sábados.",
                                  height=100)
        
        submitted = st.form_submit_button("🔄 Gerar Contrato", use_container_width=True)
    
    if submitted:
        if not cnpj or not valor or not data_inicio:
            st.error("Por favor, preencha todos os campos obrigatórios!")
        else:
            doc_bytes, nome_arquivo = preencher_contrato(
                cnpj=cnpj,
                valor=valor,
                data_inicio=data_inicio,
                local_execucao=local_execucao,
                funcoes=funcoes,
                observacoes=observacoes
            )
            
            if doc_bytes and nome_arquivo:
                st.success("✅ Contrato gerado com sucesso!")
                st.download_button(
                    label="📥 Baixar Contrato",
                    data=doc_bytes,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

if __name__ == "__main__":
    main()
