# -*- coding: utf-8 -*-
"""
Script para gerar o 1º CSV e enviar para o Google Drive, SEM salvar arquivo local.

Regras lógicas:
1) Lê todas as obras da Carteira
2) Exclui obras não aptas conforme Carteira:
   - Carteira!D = OBRA RETIRADA
   - Carteira!U = CONCLUÍDA
   - Carteira!AX vazia
   - Carteira!AY vazia
3) Para toda obra apta da Carteira, cria os pontos:
   - _VIST
   - _P0
   - _LPT
   desde que ainda não existam em BD_Obras_GPM!A
4) Também inclui códigos de ATIVIDADES_POR_PONTO_BASE!K2:K que não existam em BD_Obras_GPM!A
5) Gera CSV principal:
   - modelo_importar_obras_ponto.csv
6) Gera CSV de auditoria das obras que NÃO entraram no CSV principal:
   - obras_excluidas_import_gpm.csv

O CSV de auditoria inclui:
- Obras excluídas pela Carteira
- Obras aptas, mas sem novos códigos para importar
- Motivo técnico da não geração
- Situação em BD_Obras_GPM!A
- Situação em BD_Obras_GPM!B
- Situação em ATIVIDADES_POR_PONTO_BASE!K

AJUSTE PARA GITHUB:
- Lê credenciais do Secret GOOGLE_CREDENTIALS (JSON) se existir
- Caso contrário, usa credenciais.json local
"""

import os
import json
import csv
import re
import unicodedata
from datetime import datetime
from io import StringIO, BytesIO

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError


# =============== CONFIGURAÇÕES GERAIS ===============

SERVICE_ACCOUNT_FILE = "credenciais.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",
]

ID_CARTEIRA = "1T6HVLBQi21CIeS64tAjI314TYi2795COOCAakzLV-q0"
ID_OBRAS_GPM = "189JPWONK4hSpziocviwSQOtj59rWl9tbhkVvrxb6Lds"
ID_ATIVIDADES = "1Ipp454Clq0lKik8G5LjMMmV-8eA0R6if4FGG555K1j8"

DRIVE_FOLDER_ID = "1r2-jH6hF5jtaO7UVq7HnYX9Vquod6YDB"

CSV_FIXED_NAME = "modelo_importar_obras_ponto.csv"
CSV_EXCLUSOES_FIXED_NAME = "obras_excluidas_import_gpm.csv"

CSV_HEADER = [
    "Projeto",
    "OT Principal",
    "Titulo da obra",
    "Municipio",
    "CONTRATO",
    "TIPO DE OBRA",
    "NATUREZA OBRAS",
    "CENTRO SERVICO",
    "RESPONSAVEL",
    "TIPO SERVICO",
    "DATA_RECEBIMENTO",
    "DATA_PREVISTO_RECEBIMENTO",
    "OBRIGA VISTORIA",
    "DATA_PREVISAO_INICIO",
    "DATA_PREVISAO_FIM",
    "STATUS",
    "OBSERVACAO",
    "Postes",
    "fiscal",
    "LATITUDE",
    "LONGITUDE",
    "MEDIDOR",
]

CSV_EXCLUSOES_HEADER = [
    "Base Obra",
    "Codigo Carteira",
    "Linha Carteira",
    "Classificacao Auditoria",
    "Motivo Principal",
    "Motivos Detalhados",
    "Carteira D",
    "Carteira U",
    "Latitude AX",
    "Longitude AY",
    "Existe em BD_Obras_GPM!B?",
    "Qtd Projetos Existentes em BD_Obras_GPM!A",
    "Projetos Existentes em BD_Obras_GPM!A",
    "Sufixos Padrao Existentes",
    "Sufixos Padrao Faltantes",
    "Qtd Projetos em ATIVIDADES.K",
    "Projetos em ATIVIDADES.K",
    "Qtd Projetos Novos em ATIVIDADES.K Ausentes em Obras.A",
    "Projetos Novos em ATIVIDADES.K Ausentes em Obras.A",
    "Qtd Codigos Gerados no CSV Principal",
    "Codigos Gerados no CSV Principal",
    "Qtd Codigos Removidos por Exclusao",
    "Codigos Removidos por Exclusao",
]


# =============== LOG ===============

LOGAR_EXCLUSOES_DETALHADO = True

# Defina como None para exibir todas as bases excluídas no log.
LIMITE_LOG_EXCLUSOES = 300


# =============== AUTH ===============

def _load_credentials():
    """
    Prioridade:
    1) GOOGLE_CREDENTIALS, com JSON inteiro em env/secret
    2) credenciais.json local
    """
    env_json = os.environ.get("GOOGLE_CREDENTIALS", "").strip()

    if env_json:
        info = json.loads(env_json)
        return service_account.Credentials.from_service_account_info(
            info,
            scopes=SCOPES
        )

    return service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES
    )


# =============== NORMALIZAÇÃO ===============

def normalizar_chave(valor) -> str:
    if valor is None:
        return ""

    s = str(valor)
    s = s.replace("\u00A0", " ").replace("\u200B", " ").strip()
    s = re.sub(r"\s+", " ", s)

    if re.fullmatch(r"-?\d+\.0+", s):
        s = s.split(".", 1)[0]

    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.upper().strip()

    return s


# =============== FUNÇÕES AUXILIARES ===============

def get_services():
    creds = _load_credentials()
    sheets_service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return sheets_service, drive_service


def read_single_column(sheets_service, spreadsheet_id: str, range_a1: str):
    result = (
        sheets_service.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=range_a1
        )
        .execute()
    )

    values = result.get("values", [])
    col = []

    for row in values:
        if not row:
            continue

        val = str(row[0]).strip()

        if val != "":
            col.append(val)

    return col


def read_carteira_info(sheets_service):
    ranges = [
        "Carteira!B6:B",    # 0 - Código/base da obra
        "Carteira!AA6:AA",  # 1 - Título
        "Carteira!Y6:Y",    # 2 - Município
        "Carteira!W6:W",    # 3 - Chave Config
        "Carteira!Z6:Z",    # 4 - Tipo de obra
        "Carteira!BX6:BX",  # 5 - Natureza
        "Carteira!AX6:AX",  # 6 - Latitude
        "Carteira!AY6:AY",  # 7 - Longitude
    ]

    resp = (
        sheets_service.spreadsheets()
        .values()
        .batchGet(
            spreadsheetId=ID_CARTEIRA,
            ranges=ranges
        )
        .execute()
    )

    vrs = resp.get("valueRanges", [])

    def get_col(idx):
        if idx >= len(vrs):
            return []
        return vrs[idx].get("values", [])

    col_b = get_col(0)
    col_aa = get_col(1)
    col_y = get_col(2)
    col_w = get_col(3)
    col_z = get_col(4)
    col_bx = get_col(5)
    col_ax = get_col(6)
    col_ay = get_col(7)

    def get_val(col, i):
        if i >= len(col):
            return ""
        row = col[i]
        if not row:
            return ""
        return str(row[0]).strip()

    lista_carteira_b = []
    mapa_info = {}

    for i in range(len(col_b)):
        cod = get_val(col_b, i)

        if cod == "":
            continue

        base9 = cod[:9]

        titulo = get_val(col_aa, i)
        municipio = get_val(col_y, i)
        w = get_val(col_w, i)
        tipo_obra = get_val(col_z, i)
        natureza = get_val(col_bx, i)
        lat = get_val(col_ax, i)
        lon = get_val(col_ay, i)

        lista_carteira_b.append(cod)

        if base9 not in mapa_info:
            mapa_info[base9] = {
                "titulo": titulo,
                "municipio": municipio,
                "w": w,
                "tipo_obra": tipo_obra,
                "natureza": natureza,
                "lat": lat,
                "lon": lon,
            }

    return lista_carteira_b, mapa_info


def read_carteira_auditoria(sheets_service):
    """
    Lê a Carteira e retorna:
    1) auditoria_carteira: todas as bases encontradas
    2) exclusoes_info: apenas bases excluídas pelas regras da Carteira

    Critérios atuais de exclusão:
    - Carteira!D = OBRA RETIRADA
    - Carteira!U = CONCLUÍDA
    - Carteira!AX vazia
    - Carteira!AY vazia
    """

    ranges = [
        "Carteira!B6:B",
        "Carteira!D6:D",
        "Carteira!U6:U",
        "Carteira!AX6:AX",
        "Carteira!AY6:AY",
    ]

    resp = (
        sheets_service.spreadsheets()
        .values()
        .batchGet(
            spreadsheetId=ID_CARTEIRA,
            ranges=ranges
        )
        .execute()
    )

    vrs = resp.get("valueRanges", [])

    col_b = vrs[0].get("values", []) if len(vrs) > 0 else []
    col_d = vrs[1].get("values", []) if len(vrs) > 1 else []
    col_u = vrs[2].get("values", []) if len(vrs) > 2 else []
    col_ax = vrs[3].get("values", []) if len(vrs) > 3 else []
    col_ay = vrs[4].get("values", []) if len(vrs) > 4 else []

    def get_val(col, i):
        if i >= len(col):
            return ""
        row = col[i]
        if not row:
            return ""
        return str(row[0]).strip()

    auditoria_carteira = {}
    exclusoes_info = {}

    for i in range(len(col_b)):
        b = get_val(col_b, i)

        if b == "":
            continue

        base9 = b[:9]

        status_d = get_val(col_d, i).upper()
        status_u = get_val(col_u, i).upper()
        lat = get_val(col_ax, i)
        lon = get_val(col_ay, i)

        motivos = []

        if status_d == "OBRA RETIRADA":
            motivos.append("Carteira!D = OBRA RETIRADA")

        if status_u == "CONCLUÍDA":
            motivos.append("Carteira!U = CONCLUÍDA")

        if lat == "":
            motivos.append("Carteira!AX sem latitude")

        if lon == "":
            motivos.append("Carteira!AY sem longitude")

        if base9 not in auditoria_carteira:
            auditoria_carteira[base9] = {
                "codigo_original": b,
                "linha_carteira": i + 6,
                "motivos_exclusao": set(),
                "status_d": status_d,
                "status_u": status_u,
                "lat": lat,
                "lon": lon,
            }

        for motivo in motivos:
            auditoria_carteira[base9]["motivos_exclusao"].add(motivo)

        if motivos:
            if base9 not in exclusoes_info:
                exclusoes_info[base9] = {
                    "codigo_original": b,
                    "linha_carteira": i + 6,
                    "motivos": set(),
                    "status_d": status_d,
                    "status_u": status_u,
                    "lat": lat,
                    "lon": lon,
                }

            for motivo in motivos:
                exclusoes_info[base9]["motivos"].add(motivo)

    return auditoria_carteira, exclusoes_info


def read_config_map(sheets_service):
    result = (
        sheets_service.spreadsheets()
        .values()
        .get(
            spreadsheetId=ID_CARTEIRA,
            range="Config!L2:O"
        )
        .execute()
    )

    rows = result.get("values", [])
    config_map = {}

    for row in rows:
        if not row:
            continue

        chave_raw = str(row[0]).strip() if len(row) > 0 else ""
        chave = normalizar_chave(chave_raw)

        if chave == "":
            continue

        contrato_nome = str(row[1]).strip() if len(row) > 1 else ""
        centro_servico = str(row[2]).strip() if len(row) > 2 else ""
        responsavel = str(row[3]).strip() if len(row) > 3 else ""

        if chave not in config_map:
            config_map[chave] = {
                "contrato_nome": contrato_nome,
                "centro_servico": centro_servico,
                "responsavel": responsavel,
            }

    return config_map


# =============== CSV EM MEMÓRIA ===============

def build_csv_string(rows) -> str:
    sio = StringIO()
    writer = csv.writer(sio, delimiter=";", lineterminator="\n")

    for r in rows:
        writer.writerow(r)

    text = sio.getvalue()
    sio.close()

    return text


def montar_mapa_projetos_por_base(lista_codigos):
    mapa = {}

    for cod in lista_codigos:
        cod = str(cod).strip()

        if cod == "":
            continue

        base9 = cod[:9]
        mapa.setdefault(base9, set()).add(cod)

    return mapa


def montar_linhas_csv_exclusoes(
    auditoria_carteira,
    exclusoes_info,
    set_obras_b,
    projetos_obras_a_por_base,
    projetos_ativ_k_por_base,
    projetos_ativ_k_novos_por_base,
    codigos_gerados_csv_final_por_base,
    codigos_removidos_por_base
):
    """
    Monta auditoria das obras que NÃO entraram no CSV principal.

    Inclui:
    1) Bases excluídas pela Carteira
    2) Bases aptas que não geraram nenhum código no CSV principal
    """

    linhas = [CSV_EXCLUSOES_HEADER]

    total_excluidas_carteira = 0
    total_aptas_nao_geradas = 0

    for base9, info_cart in sorted(auditoria_carteira.items(), key=lambda x: x[0]):
        codigos_gerados = sorted(codigos_gerados_csv_final_por_base.get(base9, set()))

        # Se gerou algo no CSV principal e não foi excluída, não entra no arquivo de auditoria.
        if base9 not in exclusoes_info and codigos_gerados:
            continue

        codigo_original = info_cart.get("codigo_original", "")
        linha_carteira = info_cart.get("linha_carteira", "")
        status_d = info_cart.get("status_d", "")
        status_u = info_cart.get("status_u", "")
        lat = info_cart.get("lat", "")
        lon = info_cart.get("lon", "")

        existe_obras_b = "SIM" if base9 in set_obras_b or codigo_original in set_obras_b else "NÃO"

        projetos_obras_a = sorted(projetos_obras_a_por_base.get(base9, set()))
        projetos_ativ_k = sorted(projetos_ativ_k_por_base.get(base9, set()))
        projetos_ativ_k_novos = sorted(projetos_ativ_k_novos_por_base.get(base9, set()))
        codigos_removidos = sorted(set(codigos_removidos_por_base.get(base9, [])))

        sufixos_padrao = [
            f"{codigo_original}_VIST",
            f"{codigo_original}_P0",
            f"{codigo_original}_LPT",
        ]

        sufixos_existentes = [c for c in sufixos_padrao if c in projetos_obras_a]
        sufixos_faltantes = [c for c in sufixos_padrao if c not in projetos_obras_a]

        motivos_detalhados = []
        classificacao = ""
        motivo_principal = ""

        if base9 in exclusoes_info:
            total_excluidas_carteira += 1
            classificacao = "EXCLUIDA_CARTEIRA"
            motivos_exclusao = exclusoes_info[base9].get("motivos", set())
            motivo_principal = "Obra não apta conforme regras da Carteira"
            motivos_detalhados.extend(sorted(motivos_exclusao))

            if codigos_removidos:
                motivos_detalhados.append(
                    "Havia códigos candidatos para importação, mas foram removidos por regra de exclusão da Carteira"
                )

        else:
            total_aptas_nao_geradas += 1
            classificacao = "APTA_NAO_GERADA"

            if len(sufixos_faltantes) == 0 and len(projetos_ativ_k_novos) == 0:
                motivo_principal = (
                    "Obra apta, mas sem novos códigos para importar"
                )
                motivos_detalhados.append(
                    "_VIST, _P0 e _LPT já existem em BD_Obras_GPM!A"
                )
                motivos_detalhados.append(
                    "Não há projetos novos em ATIVIDADES_POR_PONTO_BASE!K ausentes em BD_Obras_GPM!A"
                )

            elif len(sufixos_faltantes) > 0 and len(codigos_gerados) == 0:
                motivo_principal = (
                    "Obra apta com sufixos faltantes, mas não gerou código - revisar lógica"
                )
                motivos_detalhados.append(
                    "Existem sufixos padrão faltantes em BD_Obras_GPM!A"
                )

            elif len(projetos_ativ_k_novos) > 0 and len(codigos_gerados) == 0:
                motivo_principal = (
                    "Obra apta com projetos novos no orçamento, mas não gerou código - revisar lógica"
                )
                motivos_detalhados.append(
                    "Existem projetos em ATIVIDADES.K ausentes em BD_Obras_GPM!A"
                )

            else:
                motivo_principal = (
                    "Obra apta não entrou no CSV principal por regra técnica não classificada"
                )
                motivos_detalhados.append(
                    "Revisar BD_Obras_GPM!A, BD_Obras_GPM!B e ATIVIDADES.K"
                )

            if existe_obras_b == "NÃO":
                motivos_detalhados.append(
                    "Base não encontrada em BD_Obras_GPM!B"
                )

            if len(projetos_ativ_k) == 0:
                motivos_detalhados.append(
                    "Nenhum projeto encontrado em ATIVIDADES_POR_PONTO_BASE!K para esta base"
                )

        linhas.append([
            base9,
            codigo_original,
            linha_carteira,
            classificacao,
            motivo_principal,
            " | ".join(motivos_detalhados),
            status_d,
            status_u,
            lat,
            lon,
            existe_obras_b,
            len(projetos_obras_a),
            " | ".join(projetos_obras_a),
            " | ".join(sufixos_existentes),
            " | ".join(sufixos_faltantes),
            len(projetos_ativ_k),
            " | ".join(projetos_ativ_k),
            len(projetos_ativ_k_novos),
            " | ".join(projetos_ativ_k_novos),
            len(codigos_gerados),
            " | ".join(codigos_gerados),
            len(codigos_removidos),
            " | ".join(codigos_removidos),
        ])

    print("📊 Resumo auditoria de obras não importadas:")
    print(f"   Obras excluídas pela Carteira: {total_excluidas_carteira}")
    print(f"   Obras aptas mas não geradas: {total_aptas_nao_geradas}")
    print(f"   Total no CSV de auditoria sem cabeçalho: {len(linhas) - 1}")

    return linhas


# =============== DRIVE ===============

def _escape_drive_q(s: str) -> str:
    return s.replace("\\", "\\\\").replace("'", "\\'")


def list_files_same_name_in_folder(drive_service, folder_id: str, filename: str):
    fn = _escape_drive_q(filename)

    q = f"'{folder_id}' in parents and trashed = false and name = '{fn}'"

    out = []
    page_token = None

    while True:
        resp = drive_service.files().list(
            q=q,
            fields="nextPageToken, files(id,name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            corpora="allDrives",
            pageSize=1000,
            pageToken=page_token,
        ).execute()

        out.extend(resp.get("files", []))

        page_token = resp.get("nextPageToken")

        if not page_token:
            break

    return out


def upload_csv_to_drive(drive_service, csv_text: str, filename: str, folder_id: str):
    try:
        folder = drive_service.files().get(
            fileId=folder_id,
            fields="id, name, mimeType",
            supportsAllDrives=True,
        ).execute()

        print(f"📁 Pasta encontrada no Drive: {folder.get('name')} ({folder.get('id')})")

    except HttpError as e:
        print(f"❌ Não foi possível acessar a pasta {folder_id}.")
        print(f"   Detalhes: {e}")
        return

    data = csv_text.encode("utf-8-sig")
    bio = BytesIO(data)
    media = MediaIoBaseUpload(bio, mimetype="text/csv", resumable=True)

    existing = list_files_same_name_in_folder(drive_service, folder_id, filename)

    if not existing:
        try:
            created = drive_service.files().create(
                body={
                    "name": filename,
                    "mimeType": "text/csv",
                    "parents": [folder_id],
                },
                media_body=media,
                fields="id, webViewLink, modifiedTime",
                supportsAllDrives=True,
            ).execute()

            print(f"✅ Criado no Drive: {filename}")
            print(f"   ID: {created.get('id')}")
            print(f"   Link: {created.get('webViewLink')}")
            print(f"   🕒 modifiedTime: {created.get('modifiedTime')}")

        except HttpError as e:
            print(f"❌ Erro ao criar arquivo '{filename}' no Drive.")
            print(f"   Detalhes: {e}")

        return

    main = existing[0]
    file_id = main["id"]

    try:
        updated = drive_service.files().update(
            fileId=file_id,
            media_body=media,
            fields="id, webViewLink, modifiedTime",
            supportsAllDrives=True,
        ).execute()

        print(f"♻️ Atualizado no Drive (substituído): {filename}")
        print(f"   ID: {updated.get('id')}")
        print(f"   Link: {updated.get('webViewLink')}")
        print(f"   🕒 modifiedTime: {updated.get('modifiedTime')}")

    except HttpError as e:
        msg = str(e)

        if "File not found" in msg or "notFound" in msg:
            print("⚠️ Arquivo listado mas não encontrado no update. Criando novo arquivo...")

            bio.seek(0)
            media2 = MediaIoBaseUpload(bio, mimetype="text/csv", resumable=True)

            try:
                created = drive_service.files().create(
                    body={
                        "name": filename,
                        "mimeType": "text/csv",
                        "parents": [folder_id],
                    },
                    media_body=media2,
                    fields="id, webViewLink, modifiedTime",
                    supportsAllDrives=True,
                ).execute()

                print(f"✅ Recriado no Drive: {filename}")
                print(f"   ID: {created.get('id')}")
                print(f"   Link: {created.get('webViewLink')}")
                print(f"   🕒 modifiedTime: {created.get('modifiedTime')}")

            except HttpError as e2:
                print(f"❌ Erro ao recriar arquivo '{filename}' no Drive.")
                print(f"   Detalhes: {e2}")

        else:
            print(f"❌ Erro ao atualizar arquivo existente '{filename}' ID {file_id}.")
            print(f"   Detalhes: {e}")


# =============== LÓGICA DO 1º CSV ===============

def montar_primeiro_csv():
    sheets_service, drive_service = get_services()

    print("🔄 Lendo dados da Carteira...")
    carteira_b, carteira_info = read_carteira_info(sheets_service)
    print(f"   Linhas/códigos lidos em Carteira!B: {len(carteira_b)}")

    print("🔄 Lendo auditoria da Carteira e bases excluídas...")
    auditoria_carteira, exclusoes_info = read_carteira_auditoria(sheets_service)
    set_base9_excluidas = set(exclusoes_info.keys())

    print(f"   Bases únicas auditadas na Carteira: {len(auditoria_carteira)}")
    print(f"   Base9 excluídas pela Carteira: {len(set_base9_excluidas)}")

    if LOGAR_EXCLUSOES_DETALHADO and exclusoes_info:
        print("📋 Detalhamento das bases excluídas pela Carteira:")

        itens = sorted(exclusoes_info.items(), key=lambda x: x[0])

        if LIMITE_LOG_EXCLUSOES is not None:
            itens_log = itens[:LIMITE_LOG_EXCLUSOES]
        else:
            itens_log = itens

        for base9, info_exc in itens_log:
            motivos_txt = " | ".join(sorted(info_exc["motivos"]))
            codigo_original = info_exc.get("codigo_original", "")
            linha_carteira = info_exc.get("linha_carteira", "")

            print(
                f"   🚫 Base: {base9} | "
                f"Código Carteira: {codigo_original} | "
                f"Linha Carteira: {linha_carteira} | "
                f"Motivo(s): {motivos_txt}"
            )

        if LIMITE_LOG_EXCLUSOES is not None and len(itens) > LIMITE_LOG_EXCLUSOES:
            print(
                f"   ... (+{len(itens) - LIMITE_LOG_EXCLUSOES} bases excluídas não exibidas no log)"
            )

    print("🔄 Lendo Config (lookup normalizado)...")
    config_map = read_config_map(sheets_service)
    print(f"   Chaves Config carregadas normalizadas: {len(config_map)}")

    print("🔄 Lendo BD_Obras_GPM (A e B) e ATIVIDADES (K)...")
    obras_b = read_single_column(
        sheets_service,
        ID_OBRAS_GPM,
        "BD_Obras_GPM!B2:B"
    )

    obras_a = read_single_column(
        sheets_service,
        ID_OBRAS_GPM,
        "BD_Obras_GPM!A2:A"
    )

    ativ_k = read_single_column(
        sheets_service,
        ID_ATIVIDADES,
        "ATIVIDADES_POR_PONTO_BASE!K2:K"
    )

    set_obras_b = set(obras_b)
    set_obras_a = set(obras_a)

    projetos_obras_a_por_base = montar_mapa_projetos_por_base(obras_a)
    projetos_ativ_k_por_base = montar_mapa_projetos_por_base(ativ_k)

    projetos_ativ_k_novos_por_base = {}

    for cod in ativ_k:
        cod = str(cod).strip()

        if cod == "":
            continue

        if cod in set_obras_a:
            continue

        base9 = cod[:9]
        projetos_ativ_k_novos_por_base.setdefault(base9, set()).add(cod)

    print(f"   Projetos existentes em BD_Obras_GPM!A: {len(set_obras_a)}")
    print(f"   Bases existentes em BD_Obras_GPM!B: {len(set_obras_b)}")
    print(f"   Códigos lidos em ATIVIDADES_POR_PONTO_BASE!K: {len(ativ_k)}")

    print("🧮 Parte 1: criando sufixos que faltam (_VIST, _P0, _LPT) para TODAS as obras aptas da Carteira...")

    sufixos = ["_VIST", "_P0", "_LPT"]
    criados = []
    criados_set = set()

    bases_carteira_unicas = []
    bases_carteira_set = set()

    for cod_base in carteira_b:
        cod_base = str(cod_base).strip()

        if cod_base == "":
            continue

        if cod_base in bases_carteira_set:
            continue

        bases_carteira_unicas.append(cod_base)
        bases_carteira_set.add(cod_base)

    print(f"   Bases únicas na Carteira: {len(bases_carteira_unicas)}")

    bases_aptas_carteira = 0
    bases_nao_aptas_excluidas = 0
    bases_aptas_existentes_obras_b = 0
    bases_aptas_nao_existentes_obras_b = 0

    sufixos_ja_existentes = 0
    sufixos_novos_criados = 0

    for cod_base in bases_carteira_unicas:
        base9 = cod_base[:9]

        if base9 in set_base9_excluidas:
            bases_nao_aptas_excluidas += 1
            continue

        bases_aptas_carteira += 1

        if cod_base in set_obras_b or base9 in set_obras_b:
            bases_aptas_existentes_obras_b += 1
        else:
            bases_aptas_nao_existentes_obras_b += 1

        for suf in sufixos:
            codigo_completo = f"{cod_base}{suf}"

            if codigo_completo in set_obras_a:
                sufixos_ja_existentes += 1
                continue

            if codigo_completo in criados_set:
                continue

            criados.append(codigo_completo)
            criados_set.add(codigo_completo)
            sufixos_novos_criados += 1

    print(f"   Bases aptas na Carteira: {bases_aptas_carteira}")
    print(f"   Bases não aptas/excluídas: {bases_nao_aptas_excluidas}")
    print(f"   Bases aptas que JÁ existem em BD_Obras_GPM!B: {bases_aptas_existentes_obras_b}")
    print(f"   Bases aptas que NÃO existem em BD_Obras_GPM!B: {bases_aptas_nao_existentes_obras_b}")
    print(f"   Sufixos _VIST/_P0/_LPT já existentes em BD_Obras_GPM!A: {sufixos_ja_existentes}")
    print(f"   Parte 1: {sufixos_novos_criados} códigos criados")

    print("🧮 Parte 2: incluindo projetos existentes em ATIVIDADES.K que não estão em Obras.A (sem repetição)...")

    faltantes_ativ = []
    faltantes_ativ_set = set()

    ignorados_ativ_ja_existentes = 0
    ignorados_ativ_base_fora_carteira = 0
    ignorados_ativ_vazios = 0

    for cod in ativ_k:
        if cod is None:
            ignorados_ativ_vazios += 1
            continue

        cod = str(cod).strip()

        if cod == "":
            ignorados_ativ_vazios += 1
            continue

        if cod in faltantes_ativ_set:
            continue

        if cod in set_obras_a:
            ignorados_ativ_ja_existentes += 1
            continue

        base9 = cod[:9]

        if base9 not in carteira_info:
            ignorados_ativ_base_fora_carteira += 1
            continue

        faltantes_ativ.append(cod)
        faltantes_ativ_set.add(cod)

    print(
        f"   Parte 2: {len(faltantes_ativ)} códigos únicos vindos de "
        f"ATIVIDADES.K e ausentes em Obras.A"
    )
    print(f"   ATIVIDADES.K ignorados por já existirem em Obras.A: {ignorados_ativ_ja_existentes}")
    print(f"   ATIVIDADES.K ignorados por base não existir na Carteira: {ignorados_ativ_base_fora_carteira}")
    print(f"   ATIVIDADES.K ignorados vazios/nulos: {ignorados_ativ_vazios}")

    codigos_final = []
    codigos_final_set = set()

    for c in criados:
        if c not in codigos_final_set:
            codigos_final.append(c)
            codigos_final_set.add(c)

    for c in faltantes_ativ:
        if c not in codigos_final_set:
            codigos_final.append(c)
            codigos_final_set.add(c)

    print(f"   Total de códigos para CSV antes dos filtros finais: {len(codigos_final)}")

    linhas_csv = [CSV_HEADER]
    hoje_str = datetime.now().strftime("%d/%m/%Y")

    removidos_excluidas = 0
    config_nao_encontrada = 0
    codigos_removidos_por_base = {}
    codigos_gerados_csv_final_por_base = {}

    print("🧹 Aplicando filtros finais e montando linhas do CSV principal...")

    for cod in codigos_final:
        base9 = cod[:9]

        if base9 in set_base9_excluidas:
            removidos_excluidas += 1

            codigos_removidos_por_base.setdefault(base9, []).append(cod)

            info_exc = exclusoes_info.get(base9, {})
            motivos = info_exc.get("motivos", set())
            motivos_txt = " | ".join(sorted(motivos)) if motivos else "Motivo não identificado"

            print(
                f"🚫 Removido do CSV final: {cod} | "
                f"Base: {base9} | "
                f"Motivo(s): {motivos_txt}"
            )

            continue

        info = carteira_info.get(base9, {})

        titulo = info.get("titulo", "")
        municipio = info.get("municipio", "")
        w_raw = info.get("w", "")
        w_key = normalizar_chave(w_raw)
        tipo_obra = info.get("tipo_obra", "")
        natureza = info.get("natureza", "")
        lat = info.get("lat", "")
        lon = info.get("lon", "")

        conf = config_map.get(w_key, {})

        if not conf and w_key:
            config_nao_encontrada += 1

        contrato_nome = conf.get("contrato_nome", "")
        centro_servico = conf.get("centro_servico", "")
        responsavel = conf.get("responsavel", "")

        ot_principal = base9
        tipo_servico = "OBRAS CCM"
        obriga_vistoria = "SIM" if cod.endswith("_VIST") else "NÃO"
        status = "Recebida" if (cod.endswith("_VIST") or cod.endswith("_LPT")) else "Em execução"

        linhas_csv.append([
            cod,
            ot_principal,
            titulo,
            municipio,
            contrato_nome,
            tipo_obra,
            natureza,
            centro_servico,
            responsavel,
            tipo_servico,
            hoje_str,
            hoje_str,
            obriga_vistoria,
            "",
            "",
            status,
            "",
            "",
            "",
            lat,
            lon,
            "",
        ])

        codigos_gerados_csv_final_por_base.setdefault(base9, set()).add(cod)

    print(f"   Linhas removidas por filtros D/U/AX/AY: {removidos_excluidas}")
    print(f"   Lookups de Config não encontrados mesmo normalizando: {config_nao_encontrada}")
    print(f"   Total final de linhas no CSV principal incluindo cabeçalho: {len(linhas_csv)}")

    print("📄 Gerando CSV de auditoria das obras que não entraram no CSV principal...")
    linhas_exclusoes_csv = montar_linhas_csv_exclusoes(
        auditoria_carteira=auditoria_carteira,
        exclusoes_info=exclusoes_info,
        set_obras_b=set_obras_b,
        projetos_obras_a_por_base=projetos_obras_a_por_base,
        projetos_ativ_k_por_base=projetos_ativ_k_por_base,
        projetos_ativ_k_novos_por_base=projetos_ativ_k_novos_por_base,
        codigos_gerados_csv_final_por_base=codigos_gerados_csv_final_por_base,
        codigos_removidos_por_base=codigos_removidos_por_base,
    )

    print(f"   Total de linhas no CSV de auditoria incluindo cabeçalho: {len(linhas_exclusoes_csv)}")

    csv_text = build_csv_string(linhas_csv)
    csv_exclusoes_text = build_csv_string(linhas_exclusoes_csv)

    print("💾 CSV principal gerado em memória. Enviando direto para o Drive...")
    upload_csv_to_drive(
        drive_service,
        csv_text,
        CSV_FIXED_NAME,
        DRIVE_FOLDER_ID
    )

    print("💾 CSV de auditoria gerado em memória. Enviando direto para o Drive...")
    upload_csv_to_drive(
        drive_service,
        csv_exclusoes_text,
        CSV_EXCLUSOES_FIXED_NAME,
        DRIVE_FOLDER_ID
    )


if __name__ == "__main__":
    print("🚀 Iniciando geração do 1º CSV...")
    montar_primeiro_csv()
    print("✅ Processo concluído (1º CSV).")
