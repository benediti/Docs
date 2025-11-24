# -*- coding: utf-8 -*-
import streamlit as st
import requests
from docx import Document
from docx2pdf import convert
import re
from datetime import datetime
import os
import io
import tempfile
import platform

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

def consultar_cnpj(cnpj):
    """
    Consulta dados do CNPJ na API BrasilAPI
    Retorna os dados ou None em caso de erro
    """
    cnpj_limpo = re.sub(r'\D', '', cnpj)
    
    try:
        url = f"{API_CNPJ}{cnpj_limpo}"
        r = requests.get(url, timeout=10)
        
        if r.status_code != 200:
            st.error(f"❌ Erro ao consultar CNPJ (Status {r.status_code})")
            return None
        
        dados = r.json()
        return dados
            
    except requests.exceptions.Timeout:
        st.error("❌ Tempo esgotado ao consultar API. Tente novamente.")
        return None
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Erro na requisição: {str(e)}")
        return None
    except Exception as e:
        st.error(f"❌ Erro inesperado: {str(e)}")
        return None

def converter_docx_para_pdf(docx_bytes, nome_arquivo_base):
    """
    Converte um documento DOCX em bytes para PDF
    Retorna os bytes do PDF ou None se houver erro
    """
    try:
        # Criar arquivos temporários
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
            tmp_docx.write(docx_bytes)
            tmp_docx_path = tmp_docx.name
        
        tmp_pdf_path = tmp_docx_path.replace('.docx', '.pdf')
        
        # Converter para PDF (funciona apenas no Windows com Word instalado)
        if platform.system() == 'Windows':
            try:
                convert(tmp_docx_path, tmp_pdf_path)
                
                # Ler o PDF gerado
                with open(tmp_pdf_path, 'rb') as f:
                    pdf_bytes = f.read()
                
                # Limpar arquivos temporários
                os.unlink(tmp_docx_path)
                os.unlink(tmp_pdf_path)
                
                return pdf_bytes
            except Exception as e:
                st.warning(f"⚠️ Não foi possível converter para PDF: {str(e)}")
                st.info("💡 A conversão para PDF requer Microsoft Word instalado no Windows.")
                # Limpar arquivo temporário
                if os.path.exists(tmp_docx_path):
                    os.unlink(tmp_docx_path)
                return None
        else:
            st.info("💡 Conversão para PDF disponível apenas no Windows com Word instalado.")
            os.unlink(tmp_docx_path)
            return None
            
    except Exception as e:
        st.error(f"❌ Erro ao converter para PDF: {str(e)}")
        return None

def preencher_contrato(tipo_servico, nome_servico, cnpj, ie_cliente, valor, data_inicio, 
                       local_execucao, funcoes, observacoes, dados_cnpj=None):
    """
    Gera um contrato preenchido automaticamente com base em um modelo DOCX.
    Retorna o documento DOCX e PDF em bytes para download.
    """
    
    # --- 1️⃣ Usar dados do CNPJ já consultados ou informados manualmente ---
    if dados_cnpj:
        nome_cliente = dados_cnpj.get("razao_social", "")
        nome_fantasia = dados_cnpj.get("nome_fantasia", "")
        
        # Montar endereço completo
        # A API BrasilAPI não retorna tipo de logradouro separado, apenas o nome
        logradouro = dados_cnpj.get('logradouro', '')  # Ex: BRIGADEIRO FARIA LIMA
        numero = dados_cnpj.get('numero', '')
        complemento = dados_cnpj.get('complemento', '')
        bairro = dados_cnpj.get('bairro', '')
        municipio = dados_cnpj.get('municipio', '')
        uf = dados_cnpj.get('uf', '')
        cep = dados_cnpj.get('cep', '')
        
        # Construir endereço
        partes_endereco = []
        if logradouro:
            partes_endereco.append(logradouro)
        if numero:
            partes_endereco.append(numero)
        if complemento:
            partes_endereco.append(complemento)
        
        endereco_base = ', '.join(partes_endereco)
        
        # Adicionar bairro, cidade e CEP
        if bairro:
            endereco_base += f" - {bairro}"
        if municipio and uf:
            endereco_base += f" - {municipio}/{uf}"
        if cep:
            # Formatar CEP (XXXXX-XXX)
            if len(cep) == 8:
                cep_formatado = f"{cep[:5]}-{cep[5:]}"
                endereco_base += f", CEP {cep_formatado}"
            else:
                endereco_base += f", CEP {cep}"
        
        endereco_cliente = endereco_base
        cnpj_formatado = re.sub(r'\D', '', cnpj)
    else:
        nome_cliente = "NÃO INFORMADO"
        nome_fantasia = ""
        endereco_cliente = "NÃO INFORMADO"
        cnpj_formatado = re.sub(r'\D', '', cnpj)

    # --- 2️⃣ Formatar valor ---
    valor_extenso = numero_para_extenso(float(valor))

    # --- 3️⃣ Verificar se o modelo existe ---
    if not os.path.exists(MODELO_PATH):
        st.error(f"❌ Erro: Arquivo modelo não encontrado em: {MODELO_PATH}")
        st.info(f"Diretório atual: {os.getcwd()}")
        st.info(f"Arquivos disponíveis: {os.listdir(SCRIPT_DIR) if os.path.exists(SCRIPT_DIR) else 'N/A'}")
        return None, None, None, None
    
    # --- 4️⃣ Abrir o modelo e substituir tags ---
    doc = Document(MODELO_PATH)
    substituicoes = {
        "{{tipo_servico}}": tipo_servico,
        "{{nome_servico}}": nome_servico,
        "{{nome_cliente}}": nome_cliente,
        "{{nome_fantasia}}": nome_fantasia if nome_fantasia else "",
        "{{endereco_cliente}}": endereco_cliente,
        "{{cnpj}}": cnpj_formatado,
        "{{ie_cliente}}": ie_cliente,
        "{{funcoes}}": funcoes,
        "{{observacoes}}": observacoes,
        "{{local_execucao}}": local_execucao,
        "{{valor_num}}": f"R$ {valor}",
        "{{valor_extenso}}": valor_extenso.capitalize(),
        "{{data_inicio}}": data_inicio
    }

    # Substituir em parágrafos
    for p in doc.paragraphs:
        for tag, val in substituicoes.items():
            if tag in p.text:
                p.text = p.text.replace(tag, str(val))
    
    # Substituir em tabelas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for tag, val in substituicoes.items():
                    if tag in cell.text:
                        cell.text = cell.text.replace(tag, str(val))

    # --- 5️⃣ Salvar DOCX em memória ---
    nome_base = f"contrato_{nome_cliente[:20].strip().replace(' ', '_')}"
    nome_arquivo_docx = f"{nome_base}.docx"
    nome_arquivo_pdf = f"{nome_base}.pdf"
    
    # Salvar DOCX em bytes
    docx_bytes_io = io.BytesIO()
    doc.save(docx_bytes_io)
    docx_bytes_io.seek(0)
    docx_bytes = docx_bytes_io.getvalue()
    
    # --- 6️⃣ Converter para PDF ---
    pdf_bytes = converter_docx_para_pdf(docx_bytes, nome_base)
    
    return docx_bytes, nome_arquivo_docx, pdf_bytes, nome_arquivo_pdf

# ========= INTERFACE STREAMLIT =========
def main():
    st.set_page_config(page_title="Gerador de Contratos", page_icon="📄", layout="wide")
    
    st.title("📄 Gerador Automático de Contratos")
    st.markdown("Preencha os dados abaixo para gerar um contrato personalizado automaticamente.")
    
    # Inicializar session state
    if 'dados_cnpj' not in st.session_state:
        st.session_state.dados_cnpj = None
    if 'cnpj_consultado' not in st.session_state:
        st.session_state.cnpj_consultado = ""
    
    # Seção para consulta de CNPJ
    st.markdown("### 🔍 Consulta de CNPJ")
    col_cnpj1, col_cnpj2 = st.columns([3, 1])
    
    with col_cnpj1:
        cnpj_input = st.text_input("Digite o CNPJ para consulta automática", 
                                   value="65035552000180",
                                   help="Digite o CNPJ e clique em Consultar")
    with col_cnpj2:
        st.markdown("<br>", unsafe_allow_html=True)
        btn_consultar = st.button("🔎 Consultar CNPJ", use_container_width=True, type="primary")
    
    # Consultar CNPJ quando o botão for clicado
    if btn_consultar and cnpj_input:
        with st.spinner("Consultando CNPJ..."):
            dados = consultar_cnpj(cnpj_input)
            if dados:
                st.session_state.dados_cnpj = dados
                st.session_state.cnpj_consultado = cnpj_input
                st.success(f"✅ Dados encontrados: {dados.get('razao_social', 'N/A')}")
                
                # Mostrar dados encontrados
                with st.expander("📋 Dados da Empresa", expanded=True):
                    col_info1, col_info2 = st.columns(2)
                    with col_info1:
                        st.write(f"**Razão Social:** {dados.get('razao_social', 'N/A')}")
                        nome_fant = dados.get('nome_fantasia', '')
                        st.write(f"**Nome Fantasia:** {nome_fant if nome_fant else 'Não informado'}")
                        st.write(f"**CNPJ:** {dados.get('cnpj', 'N/A')}")
                    with col_info2:
                        # Montar endereço completo
                        log = dados.get('logradouro', '')
                        num = dados.get('numero', '')
                        comp = dados.get('complemento', '')
                        bairro = dados.get('bairro', '')
                        mun = dados.get('municipio', '')
                        uf = dados.get('uf', '')
                        cep = dados.get('cep', '')
                        
                        # Formatar CEP
                        if cep and len(cep) == 8:
                            cep_formatado = f"{cep[:5]}-{cep[5:]}"
                        else:
                            cep_formatado = cep
                        
                        endereco_partes = []
                        if log:
                            endereco_partes.append(log)
                        if num:
                            endereco_partes.append(num)
                        if comp:
                            endereco_partes.append(comp)
                        
                        endereco_linha1 = ', '.join(endereco_partes)
                        
                        st.write(f"**Endereço:** {endereco_linha1}")
                        st.write(f"**Bairro:** {bairro}")
                        st.write(f"**Cidade/UF:** {mun}/{uf}")
                        st.write(f"**CEP:** {cep_formatado}")
    
    st.markdown("---")
    
    # Formulário principal
    with st.form("contrato_form", clear_on_submit=False):
        st.markdown("### 📝 Dados do Serviço")
        
        col1, col2 = st.columns(2)
        
        with col1:
            tipo_servico = st.text_input("Tipo de Serviço *", 
                                        value="Prestação de Serviços de Limpeza",
                                        help="Ex: Prestação de Serviços de Limpeza")
            
            nome_servico = st.text_input("Nome do Serviço *", 
                                        value="Limpeza e Conservação",
                                        help="Ex: Limpeza e Conservação")
            
            valor = st.text_input("Valor Mensal (R$) *", 
                                 value="3400.00",
                                 help="Valor numérico, ex: 3400.00")
            
            data_inicio = st.text_input("Data de Início *", 
                                       value="03/11/2025",
                                       help="Formato: DD/MM/AAAA")
        
        with col2:
            cnpj = st.text_input("CNPJ *", 
                                value=st.session_state.cnpj_consultado if st.session_state.cnpj_consultado else "65035552000180",
                                help="Será preenchido automaticamente após consulta")
            
            ie_cliente = st.text_input("Inscrição Estadual", 
                                      value="",
                                      help="Deixe em branco se não houver")
            
            local_execucao = st.text_area("Local de Execução *", 
                                         value="Rua Joaquim Murtinho, 225, Bom Retiro - São Paulo/SP",
                                         height=100,
                                         help="Endereço onde o serviço será executado")
        
        st.markdown("### 👥 Funções e Quadro Funcional")
        funcoes = st.text_area("Funções e Quadro Funcional *", 
                              value="Supervisora Operacional / Encarregada – 8h\n4 Auxiliares de Limpeza – 8h",
                              height=120,
                              help="Descreva as funções e quantidade de funcionários")
        
        st.markdown("### 📌 Observações")
        observacoes = st.text_area("Observações", 
                                  value="Serviços de limpeza geral realizados bimestralmente aos sábados.",
                                  height=100,
                                  help="Informações adicionais sobre o contrato")
        
        st.markdown("---")
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            submitted = st.form_submit_button("🔄 Gerar Contrato (DOCX e PDF)", 
                                             use_container_width=True, 
                                             type="primary")
    
    # Processar formulário
    if submitted:
        if not all([tipo_servico, nome_servico, cnpj, valor, data_inicio, local_execucao, funcoes]):
            st.error("⚠️ Por favor, preencha todos os campos obrigatórios marcados com *")
        else:
            with st.container():
                st.markdown("---")
                
                with st.spinner("Gerando contrato..."):
                    docx_bytes, nome_docx, pdf_bytes, nome_pdf = preencher_contrato(
                        tipo_servico=tipo_servico,
                        nome_servico=nome_servico,
                        cnpj=cnpj,
                        ie_cliente=ie_cliente if ie_cliente else "NÃO INFORMADO",
                        valor=valor,
                        data_inicio=data_inicio,
                        local_execucao=local_execucao,
                        funcoes=funcoes,
                        observacoes=observacoes,
                        dados_cnpj=st.session_state.dados_cnpj
                    )
                
                if docx_bytes and nome_docx:
                    st.success("✅ Contrato gerado com sucesso!")
                    
                    col_down1, col_down2, col_down3 = st.columns([1, 1, 1])
                    
                    with col_down1:
                        st.download_button(
                            label="📥 Baixar DOCX",
                            data=docx_bytes,
                            file_name=nome_docx,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            type="secondary"
                        )
                    
                    with col_down2:
                        if pdf_bytes and nome_pdf:
                            st.download_button(
                                label="📄 Baixar PDF",
                                data=pdf_bytes,
                                file_name=nome_pdf,
                                mime="application/pdf",
                                use_container_width=True,
                                type="primary"
                            )
                        else:
                            st.info("PDF não disponível")
                    
                    with col_down3:
                        st.button("🔄 Novo Contrato", 
                                 use_container_width=True,
                                 on_click=lambda: st.session_state.clear())

if __name__ == "__main__":
    main()
