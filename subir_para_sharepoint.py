import json
import requests
import msal
import os

username = os.getlogin()

# Função para obter o token de acesso via Azure AD
def obter_token(client_id, authority, client_secret, scope):
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )

    result = app.acquire_token_silent(scope, account=None)

    if not result:
        print("Nenhum token encontrado no cache. Obtendo novo token...")
        result = app.acquire_token_for_client(scopes=scope)

    if "access_token" in result:
        return result["access_token"]
    else:
        print(result.get("error"))
        print(result.get("error_description"))
        print(result.get("correlation_id"))
        return None


# Função para atualizar um item em uma lista do SharePoint
def atualizar_item_na_lista(site_id, lista_id, item_id, headers, json_string):
    update_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lista_id}/items/{item_id}/fields"

    dados_sharepoint = {
        "ValoresJson": json_string  # Nome da coluna no SharePoint que será atualizada
    }

    response = requests.patch(update_url, headers=headers, json=dados_sharepoint)

    if response.status_code == 200:
        print(f"Item com ID {item_id} atualizado com sucesso.")
    else:
        print(f"Falha ao atualizar item com ID {item_id}. Código: {response.status_code}")
        print("Resposta:", response.text)


# Caminho local do arquivo de configuração JSON
caminho_parameters = (f"C:/Users/{username}/CAMINHO/parameters.json")

# Carrega configurações sensíveis (Client ID, Secret, etc.)
with open(caminho_parameters, "r") as file:
    config = json.load(file)

# Autentica e obtém token para a API do Microsoft Graph
token = obter_token(
    config["client_id"], config["authority"], config["secret"], config["scope"]
)

if token:
    # IDs do site e da lista no SharePoint
    site_id = "SEU_SITE_ID"
    lista_id = "SUA_LISTA_ID"

    # Caminho para o arquivo JSON com os dados
    caminho_downloads = f"C:\\Users\\{username}\\CAMINHO\\Dcolab.json"
    
    with open(caminho_downloads, "r", encoding="utf-8") as file:
        dados_json = json.load(file)

    # Divide os dados JSON em duas partes (para evitar limites de tamanho no SharePoint)
    if isinstance(dados_json, list):
        meio = len(dados_json) // 2
        parte1 = dados_json[:meio]
        parte2 = dados_json[meio:]
    elif isinstance(dados_json, dict):
        chaves = list(dados_json.keys())
        meio = len(chaves) // 2
        parte1 = {chave: dados_json[chave] for chave in chaves[:meio]}
        parte2 = {chave: dados_json[chave] for chave in chaves[meio:]}
    else:
        print("Formato JSON não suportado.")
        exit()

    json_string1 = json.dumps(parte1, ensure_ascii=False)
    json_string2 = json.dumps(parte2, ensure_ascii=False)

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    # Atualiza dois itens na lista do SharePoint com os dados divididos
    atualizar_item_na_lista(site_id, lista_id, 1, headers, json_string1)
    atualizar_item_na_lista(site_id, lista_id, 2, headers, json_string2)

else:
    print("Falha ao obter token de acesso.")
