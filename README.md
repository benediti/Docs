# 📄 Gerador Automático de Contratos

Sistema web para geração automatizada de contratos com consulta de CNPJ e exportação para DOCX e PDF.

## 🚀 Funcionalidades

- ✅ **Consulta automática de CNPJ** via API BrasilAPI
- ✅ **Preenchimento automático** de dados da empresa (razão social, endereço, etc.)
- ✅ **Geração direta por botão** para `Contrato` ou `Adendo 2026`
- ✅ **Todos os campos personalizáveis:**
  - Tipo de Serviço
  - Nome do Serviço
  - Razão Social (CNPJ)
  - Endereço (CNPJ)
  - CNPJ
  - Inscrição Estadual
  - Funções e Quadro Funcional
  - Observações
  - Local de Execução
  - Valor Mensal (numérico e por extenso)
  - Data de Início
- ✅ **Geração de DOCX** (Word)
- ✅ **Conversão para PDF** com fallback compatível (quando Word não estiver disponível)
- ✅ **Interface intuitiva** com Streamlit

## 📋 Placeholders no Template DOCX

Certifique-se de que seu arquivo `Documento Contrato Serviço - Modelo.docx` contenha os seguintes placeholders:

```
{{tipo_servico}}         - Tipo de Serviço
{{nome_servico}}         - Nome do Serviço
{{nome_cliente}}         - Razão Social (via CNPJ)
{{nome_fantasia}}        - Nome Fantasia (via CNPJ - pode estar vazio)
{{endereco_cliente}}     - Endereço completo (Logradouro, Número, Complemento, Bairro, Cidade/UF, CEP)
{{cnpj}}                 - CNPJ
{{ie_cliente}}           - Inscrição Estadual
{{funcoes}}              - Funções e Quadro Funcional
{{observacoes}}          - Observações
{{local_execucao}}       - Local de Execução
{{valor_num}}            - Valor numérico (R$ X.XXX,XX)
{{valor_extenso}}        - Valor por extenso
{{data_inicio}}          - Data de Início
```

**Exemplo de endereço gerado via API:**
- `BRIGADEIRO FARIA LIMA, 201, CONJ 121 E 122 - PINHEIROS - SAO PAULO/SP, CEP 05426-100`

**Nota:** A API BrasilAPI não retorna o tipo de logradouro (Rua, Avenida, etc.) separadamente. Se precisar, adicione manualmente no documento.

## 🛠️ Instalação

### Requisitos
- Python 3.8+
- pip

### Passos

1. Clone o repositório:
```bash
git clone https://github.com/benediti/Docs.git
cd Docs
```

2. Instale as dependências:
```bash
pip install -r requirements.txt
```

3. Execute o aplicativo:
```bash
streamlit run doc.py
```

4. Acesse no navegador: `http://localhost:8501`

## 📦 Deploy no Streamlit Cloud

1. Faça push do código para o GitHub
2. Acesse [share.streamlit.io](https://share.streamlit.io)
3. Conecte seu repositório GitHub
4. Configure:
   - Repository: `benediti/Docs`
   - Branch: `main`
   - Main file: `doc.py`
5. Clique em "Deploy"

## 📝 Como Usar

1. **Consultar CNPJ:**
   - Digite o CNPJ no campo de consulta
   - Clique em "🔎 Consultar CNPJ"
   - Os dados da empresa serão carregados automaticamente

2. **Preencher Formulário:**
   - Complete todos os campos obrigatórios (*)
   - Os dados do CNPJ já estarão preenchidos se consultados
   - Ajuste valores conforme necessário

3. **Gerar Contrato:**
   - Clique em `📄 Gerar Contrato (DOCX e PDF)` ou `📑 Gerar Adendo 2026 (DOCX e PDF)`
   - Faça o download do DOCX e/ou PDF gerado

## ⚠️ Observações

- **Conversão PDF:** Quando Microsoft Word não estiver disponível, o sistema gera um PDF compatível com o conteúdo, podendo simplificar a formatação visual.
- **API CNPJ:** Usa a API gratuita BrasilAPI (pode ter limitações de uso)
- **Template:** O arquivo `Documento Contrato Serviço - Modelo.docx` deve estar no mesmo diretório do `doc.py`

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

## 📄 Licença

MIT License