# api-totvs-powerapps-integration
Automação integrada de extração de dados via API TOTVS, envio para SharePoint com Microsoft Graph e consumo via Power Apps.

# Automação Integrada: TOTVS API + Power Apps + Microsoft Graph

Este projeto automatiza o processo de extração de dados de colaboradores via API TOTVS, transforma os dados em JSON e os disponibiliza em uma Lista do SharePoint. Essas informações são consumidas por um aplicativo desenvolvido no Power Apps, promovendo uma integração moderna e sem dependência de Excel.

---

## Fluxo da Solução

1. **Extração via API TOTVS**
   - O código `extrair_dados_api.py` autentica na API da TOTVS, consulta as informações dos colaboradores e organiza os dados.
   - A resposta da API é tratada e transformada em um arquivo JSON.

2. **Upload para o SharePoint via Microsoft Graph API**
   - O código `subir_para_sharepoint.py` realiza a autenticação usando `msal`, carrega o JSON, divide em duas partes (caso o tamanho seja muito grande) e envia os dados para a Lista do SharePoint.
   - O conteúdo é atualizado diretamente no SharePoint, permitindo uso no ambiente corporativo.

3. **Consumo no Power Apps**
   - Um aplicativo desenvolvido dentro da Power Platform consome os dados da Lista SharePoint criada, exibindo e filtrando informações conforme as necessidades dos usuários internos.

---

## Estrutura dos Arquivos

```bash
automacao-integrada-api-powerapps/
│
├── extrair_dados_api.py          # Código que extrai e trata os dados da API TOTVS
├── subir_para_sharepoint.py      # Código que envia o JSON ao SharePoint usando Microsoft Graph
├── parameters.json               # Exemplo de estrutura de parâmetros (sem credenciais reais)
├── requirements.txt              # Lista de bibliotecas necessárias
└── README.md                     # Este arquivo
 

## Tecnologias Utilizadas

Python: Scripts de automação e tratamento de dados
Requests: Requisições HTTP à API TOTVS
MSAL (Microsoft Authentication Library): Autenticação no Microsoft Graph
Microsoft Graph API: Integração com o SharePoint
Power Apps: Aplicativo corporativo que consome os dados
SharePoint Online: Armazenamento acessível dos dados em lista


## Requisitos

pip install requests
pip install msal
pip install pandas
pip install openpyxl


## Benefícios da Solução

Integração direta com a API TOTVS
Eliminação do uso manual de Excel
Automação ponta-a-ponta no ecossistema Microsoft
Maior confiabilidade e agilidade na entrega dos dados
Interface moderna via Power Apps
