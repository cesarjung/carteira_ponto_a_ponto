# -*- coding: utf-8 -*-
"""
Script para gerar o 1º CSV e enviar para o Google Drive, SEM salvar arquivo local.

Regras lógicas (mantidas):
1) bases_validas = Carteira!B6:B ∩ BD_Obras_GPM!B2:B
2) Parte 1: códigos <base9>_VIST, _P0, _LPT que não existam em BD_Obras_GPM!A2:A
3) Parte 2: códigos de ATIVIDADES_POR_PONTO_BASE!K2:K que não estejam em Obras.A
4) Exclusões por Carteira (D/U/AX/AY)
5) Lookup Config normalizado
6) Encoding utf-8-sig para Excel não quebrar acentos ("NÃO" etc.)
7) Envia DIRETO para a pasta do Drive (ID: 1r2-jH6hF5jtaO7UVq7HnYX9Vquod6YDB)

Não é criado nenhum .csv no disco local.

AJUSTE PARA GITHUB:
- Lê credenciais do Secret GOOGLE_CREDENTIALS (JSON) se existir
- Caso contrário, usa credenciais.json local (para rodar no PC também)
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

CSV_FIXED_NAME = "modelo_importar_obras_ponto.csv"


# =============== AUTH ===============

def _load_credentials():
    """
    Prioridade:
    1) GOOGLE_CREDENTIALS (JSON inteiro em env/secret)
    2) credenciais.json local
    """
    env_json = os.environ.get("GOOGLE_CREDENTIALS", "").strip()
    if env_json:
        info = json.loads(env_json)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)

    # fallback local
    return service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# =============== NORMALIZAÇÃO DE CHAVES ===============

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


# =============== FUNÇÕES AUXILIARES (Sheets) ===============

def get_services():
    creds = _load_credentials()
    sheets_service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return sheets_service, drive_service


def read_single_column(sheets_service, spreadsheet_id: str, range_a1: str):
    result = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=range_a1)
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
        "Carteira!B6:B",   # 0
        "Carteira!AA6:AA", # 1
        "Carteira!Y6:Y",   # 2
        "Carteira!W6:W",   # 3
        "Carteira!Z6:Z",   # 4
        "Carteira!BX6:BX", # 5
        "Carteira!AX6:AX", # 6
        "Carteira!AY6:AY", # 7
    ]
    resp = (
        sheets_service.spreadsheets()
        .values()
        .batchGet(spreadsheetId=ID_CARTEIRA, ranges=ranges)
        .execute()
    )
    vrs = resp.get("valueRanges", [])

    def get_col(idx):
        if idx >= len(vrs):
            return []
        return vrs[idx].get("values", [])

    col_b  = get_col(0)
    col_aa = get_col(1)
    col_y  = get_col(2)
    col_w  = get_col(3)
    col_z  = get_col(4)
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

        titulo    = get_val(col_aa, i)
        municipio = get_val(col_y, i)
        w         = get_val(col_w, i)
        tipo_obra = get_val(col_z, i)
        natureza  = get_val(col_bx, i)
        lat       = get_val(col_ax, i)
        lon       = get_val(col_ay, i)

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


def read_carteira_base9_excluidas(sheets_service):
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
        .batchGet(spreadsheetId=ID_CARTEIRA, ranges=ranges)
        .execute()
    )
    vrs = resp.get("valueRanges", [])
    col_b  = vrs[0].get("values", []) if len(vrs) > 0 else []
    col_d  = vrs[1].get("values", []) if len(vrs) > 1 else []
    col_u  = vrs[2].get("values", []) if len(vrs) > 2 else []
    col_ax = vrs[3].get("values", []) if len(vrs) > 3 else []
    col_ay = vrs[4].get("values", []) if len(vrs) > 4 else []

    def get_val(col, i):
        if i >= len(col):
            return ""
        row = col[i]
        if not row:
            return ""
        return str(row[0]).strip()

    excluidas = set()
    for i in range(len(col_b)):
        b = get_val(col_b, i)
        if b == "":
            continue
        status_d = get_val(col_d, i).upper()
        status_u = get_val(col_u, i).upper()
        lat = get_val(col_ax, i)
        lon = get_val(col_ay, i)
        if status_d == "OBRA RETIRADA" or status_u == "CONCLUÍDA" or lat == "" or lon == "":
            excluidas.add(b[:9])
    return excluidas


def read_config_map(sheets_service):
    result = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=ID_CARTEIRA, range="Config!L2:O")
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
        contrato_nome  = str(row[1]).strip() if len(row) > 1 else ""
        centro_servico = str(row[2]).strip() if len(row) > 2 else ""
        responsavel    = str(row[3]).strip() if len(row) > 3 else ""
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


# =============== DRIVE (UPLOAD/UPDATE DIRETO) ===============

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
                body={"name": filename, "mimeType": "text/csv", "parents": [folder_id]},
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
            print("⚠️ Arquivo listado mas não encontrado no update. Criando um novo do zero...")
            bio.seek(0)
            media2 = MediaIoBaseUpload(bio, mimetype="text/csv", resumable=True)
            try:
                created = drive_service.files().create(
                    body={"name": filename, "mimeType": "text/csv", "parents": [folder_id]},
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
            print(f"❌ Erro ao atualizar arquivo existente '{filename}' (ID {file_id}).")
            print(f"   Detalhes: {e}")


# =============== LÓGICA DO 1º CSV ===============

def montar_primeiro_csv():
    sheets_service, drive_service = get_services()

    print("🔄 Lendo dados da Carteira...")
    carteira_b, carteira_info = read_carteira_info(sheets_service)

    print("🔄 Lendo base9 excluídas (OBRA RETIRADA, CONCLUÍDA, sem AX/AY)...")
    set_base9_excluidas = read_carteira_base9_excluidas(sheets_service)
    print(f"   Base9 excluídas: {len(set_base9_excluidas)}")

    print("🔄 Lendo Config (lookup normalizado)...")
    config_map = read_config_map(sheets_service)
    print(f"   Chaves Config carregadas (normalizadas): {len(config_map)}")

    print("🔄 Lendo BD_Obras_GPM (A e B) e ATIVIDADES (K)...")
    obras_b = read_single_column(sheets_service, ID_OBRAS_GPM, "BD_Obras_GPM!B2:B")
    obras_a = read_single_column(sheets_service, ID_OBRAS_GPM, "BD_Obras_GPM!A2:A")
    ativ_k = read_single_column(sheets_service, ID_ATIVIDADES, "ATIVIDADES_POR_PONTO_BASE!K2:K")

    set_obras_b = set(obras_b)
    set_obras_a = set(obras_a)

    bases_validas = set(carteira_b).intersection(set_obras_b)
    print(f"   Bases válidas (Carteira∩Obras.B): {len(bases_validas)}")

    print("🧮 Parte 1: criando sufixos que faltam (_VIST, _P0, _LPT)...")
    sufixos = ["_VIST", "_P0", "_LPT"]
    criados = []
    criados_set = set()
    for cod_base in carteira_b:
        if cod_base not in bases_validas:
            continue
        for suf in sufixos:
            codigo_completo = f"{cod_base}{suf}"
            if codigo_completo in set_obras_a:
                continue
            if codigo_completo in criados_set:
                continue
            criados.append(codigo_completo)
            criados_set.add(codigo_completo)
    print(f"   Parte 1: {len(criados)} códigos criados")

    print("🧮 Parte 2: incluindo projetos existentes em ATIVIDADES.K que não estão em Obras.A (sem repetição)...")
    faltantes_ativ = []
    faltantes_ativ_set = set()
    for cod in ativ_k:
        if cod is None:
            continue
        cod = str(cod).strip()
        if cod == "":
            continue
        if cod in faltantes_ativ_set:
            continue
        if cod in set_obras_a:
            continue
        base9 = cod[:9]
        if base9 not in carteira_info:
            continue
        faltantes_ativ.append(cod)
        faltantes_ativ_set.add(cod)
    print(f"   Parte 2: {len(faltantes_ativ)} códigos únicos vindos de ATIVIDADES.K e ausentes em Obras.A")

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
    print(f"   Total de códigos para CSV (Parte1 + Parte2): {len(codigos_final)}")

    linhas_csv = [CSV_HEADER]
    hoje_str = datetime.now().strftime("%d/%m/%Y")

    removidos_excluidas = 0
    config_nao_encontrada = 0

    for cod in codigos_final:
        base9 = cod[:9]
        if base9 in set_base9_excluidas:
            removidos_excluidas += 1
            continue

        info = carteira_info.get(base9, {})
        titulo    = info.get("titulo", "")
        municipio = info.get("municipio", "")
        w_raw     = info.get("w", "")
        w_key     = normalizar_chave(w_raw)
        tipo_obra = info.get("tipo_obra", "")
        natureza  = info.get("natureza", "")
        lat       = info.get("lat", "")
        lon       = info.get("lon", "")

        conf = config_map.get(w_key, {})
        if not conf and w_key:
            config_nao_encontrada += 1
        contrato_nome  = conf.get("contrato_nome", "")
        centro_servico = conf.get("centro_servico", "")
        responsavel    = conf.get("responsavel", "")

        ot_principal   = base9
        tipo_servico   = "OBRAS CCM"
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

    print(f"   Linhas removidas por filtros (D/U/AX/AY): {removidos_excluidas}")
    print(f"   Lookups de Config não encontrados (mesmo normalizando): {config_nao_encontrada}")
    print(f"   Total final de linhas no CSV (incluindo cabeçalho): {len(linhas_csv)}")

    csv_text = build_csv_string(linhas_csv)
    print("💾 CSV gerado em memória. Enviando direto para o Drive...")
    upload_csv_to_drive(drive_service, csv_text, CSV_FIXED_NAME, DRIVE_FOLDER_ID)


if __name__ == "__main__":
    print("🚀 Iniciando geração do 1º CSV...")
    montar_primeiro_csv()
    print("✅ Processo concluído (1º CSV).")
