# -*- coding: utf-8 -*-
import requests
from docx import Document
import re
from datetime import datetime

# ========= CONFIGURAÇÃO =========
# Nome do seu arquivo modelo (já editado com tags)
MODELO_PATH = "contrato_modelo.docx"

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
    """

    # --- 1️⃣ Buscar dados do CNPJ na API ---
    cnpj = re.sub(r'\D', '', cnpj)
    print(f"Consultando dados do CNPJ {cnpj}...")
    r = requests.get(f"{API_CNPJ}{cnpj}")
    if r.status_code != 200:
        print("❌ Erro ao consultar CNPJ:", r.text)
        return
    dados = r.json()

    # --- 2️⃣ Extrair campos ---
    nome_cliente = dados.get("razao_social", "")
    endereco_cliente = f"{dados.get('logradouro', '')}, {dados.get('numero', '')}, {dados.get('bairro', '')} - {dados.get('municipio', '')}/{dados.get('uf', '')}, CEP {dados.get('cep', '')}"

    valor_extenso = numero_para_extenso(float(valor))

    # --- 3️⃣ Abrir o modelo e substituir tags ---
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

    # --- 4️⃣ Salvar o contrato preenchido ---
    nome_arquivo = f"contrato_{nome_cliente[:20].strip().replace(' ', '_')}.docx"
    doc.save(nome_arquivo)
    print(f"✅ Contrato gerado com sucesso: {nome_arquivo}")

# ========= EXEMPLO DE USO =========
if __name__ == "__main__":
    preencher_contrato(
        cnpj="65035552000180",  # CNPJ de exemplo (Equippe)
        valor="3400.00",
        data_inicio="03/11/2025",
        local_execucao="Rua Joaquim Murtinho, 225, Bom Retiro - São Paulo/SP",
        funcoes="Supervisora Operacional / Encarregada – 8h, 4 Auxiliares de Limpeza – 8h",
        observacoes="Serviços de limpeza geral realizados bimestralmente aos sábados."
    )
