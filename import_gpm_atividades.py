# -*- coding: utf-8 -*-
"""
CSV Atividades por Projeto (usando lista do arquivo modelo_importar_obras_ponto sempre)

OBJETIVO:
- Gerar e enviar (overwrite/update) o CSV: modelo_imp_ativ_obra_lote.csv

MAPEAMENTO (aba ATIVIDADES_POR_PONTO):
- Coluna A do CSV = Projeto     <- Coluna K
- Coluna B do CSV = Atividade   <- Coluna D
- Coluna C do CSV = Quantidade  <- Coluna G

MANTÉM:
- Lista de projetos vem SEMPRE do arquivo "modelo_importar_obras_ponto" no Drive
- Planilhas a consultar vêm de BD_Config!A3:A (na planilha ID_ATIVIDADES)
- Overwrite no Drive via files.update (sem gerar (1)), cria se não existir
- Retry robusto e logs no CMD

AJUSTE PARA GITHUB:
- Lê credenciais do Secret GOOGLE_CREDENTIALS (JSON) se existir
- Caso contrário, usa credenciais.json local (para rodar no PC também)
"""

import os
import json
import csv
import io
import re
import time
import random

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload


# ===================== CONFIG =====================

SERVICE_ACCOUNT_FILE = os.path.join(os.path.dirname(__file__), "credenciais.json")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",
]

ID_ATIVIDADES = "1Ipp454Clq0lKik8G5LjMMmV-8eA0R6if4FGG555K1j8"
BD_CONFIG_RANGE_SOURCES = "BD_Config!A3:A"

DRIVE_FOLDER_ID = "1r2-jH6hF5jtaO7UVq7HnYX9Vquod6YDB"

LISTA_PROJETOS_BASE_NAME = "modelo_importar_obras_ponto"
LISTA_PROJETOS_FILENAMES_TRY = [
    f"{LISTA_PROJETOS_BASE_NAME}.csv",
    LISTA_PROJETOS_BASE_NAME,
]

CSV_FINAL_NAME = "modelo_imp_ativ_obra_lote.csv"

ATIV_SHEET = "ATIVIDADES_POR_PONTO"
RANGE_COL_I = f"{ATIV_SHEET}!K2:K"
RANGE_COL_D = f"{ATIV_SHEET}!D2:D"
RANGE_COL_G = f"{ATIV_SHEET}!G2:G"

CSV_HEADER_FINAL = ["Projeto", "Atividade", "Quantidade"]


# ===================== UTIL =====================

def log(msg: str):
    print(msg, flush=True)


def _load_credentials():
    env_json = os.environ.get("GOOGLE_CREDENTIALS", "").strip()
    if env_json:
        info = json.loads(env_json)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)


def get_services():
    log("🔑 Carregando credenciais e serviços...")
    creds = _load_credentials()
    sheets = build("sheets", "v4", credentials=creds)
    drive = build("drive", "v3", credentials=creds)
    log("✅ Serviços carregados.")
    return sheets, drive


def retry(func, *, tries=6, base_sleep=1.2, jitter=0.6):
    last_err = None
    for attempt in range(1, tries + 1):
        try:
            return func()
        except HttpError as e:
            status = getattr(e.resp, "status", None)
            if status in (429, 500, 502, 503, 504):
                sleep_s = base_sleep * (2 ** (attempt - 1)) + random.random() * jitter
                log(f"⚠️ HTTP {status} (tentativa {attempt}/{tries}) — aguardando {sleep_s:.1f}s...")
                time.sleep(sleep_s)
                last_err = e
                continue
            raise
        except Exception as e:
            last_err = e
            sleep_s = base_sleep * (2 ** (attempt - 1)) + random.random() * jitter
            log(f"⚠️ Erro '{type(e).__name__}' (tentativa {attempt}/{tries}) — aguardando {sleep_s:.1f}s...")
            time.sleep(sleep_s)
    raise last_err


def clean(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\ufeff", "")
    s = s.replace("\u00A0", " ")
    s = s.replace("\u200b", "")
    return s.strip()


def norm_key(s: str) -> str:
    return clean(s).upper()


def base_key(s: str) -> str:
    s = norm_key(s)
    if "_" in s:
        return s.split("_", 1)[0].strip()
    return s


def extrair_sheet_id(valor: str) -> str:
    if not valor:
        return ""
    v = str(valor).strip()
    if not v:
        return ""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", v)
    if m:
        return m.group(1)
    return v


def read_single_column(sheets_service, spreadsheet_id: str, range_a1: str):
    def _call():
        return (
            sheets_service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=range_a1)
            .execute()
        )
    values = retry(_call).get("values", [])
    out = []
    for row in values:
        if not row:
            continue
        val = clean(row[0])
        if val:
            out.append(val)
    return out


def read_batch_columns(sheets_service, spreadsheet_id: str, ranges: list[str]):
    def _call():
        return (
            sheets_service.spreadsheets()
            .values()
            .batchGet(spreadsheetId=spreadsheet_id, ranges=ranges)
            .execute()
        )
    return retry(_call).get("valueRanges", [])


def write_csv_local(rows, filename: str) -> str:
    base_dir = os.path.dirname(__file__)
    filepath = os.path.join(base_dir, filename)
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerows(rows)
    return filepath


# ===================== DRIVE =====================

def drive_list_files_by_exact_name_in_folder(drive_service, folder_id: str, filename: str):
    safe_name = filename.replace("'", "\\'")
    q = (
        f"'{folder_id}' in parents and trashed = false "
        f"and name = '{safe_name}'"
    )

    all_files = []
    page_token = None
    while True:
        def _call():
            return drive_service.files().list(
                q=q,
                fields="nextPageToken, files(id,name,modifiedTime,size,mimeType)",
                pageSize=200,
                pageToken=page_token,
                orderBy="modifiedTime desc",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                corpora="allDrives",
            ).execute()

        resp = retry(_call)
        all_files.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    return all_files


def drive_get_latest_file_by_name(drive_service, folder_id: str, filename: str):
    files = drive_list_files_by_exact_name_in_folder(drive_service, folder_id, filename)
    if not files:
        return None
    return files[0]


def drive_download_file_to_bytes(drive_service, file_id: str) -> bytes:
    request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request, chunksize=1024 * 1024 * 2)

    done = False
    while not done:
        status, done = retry(lambda: downloader.next_chunk(), tries=6)
        if status:
            pct = int(status.progress() * 100)
            log(f"    ⬇️ Download: {pct}%")

    return fh.getvalue()


def parse_projetos_from_relatorio_csv(content_bytes: bytes) -> list[str]:
    text = content_bytes.decode("utf-8-sig", errors="replace")
    sample = text[:4096]
    delim = ";" if sample.count(";") >= sample.count(",") else ","

    reader = csv.reader(io.StringIO(text), delimiter=delim)
    projetos = []
    seen = set()

    for idx, row in enumerate(reader, start=1):
        if not row:
            continue
        a = clean(row[0])
        if not a:
            continue
        if idx == 1 and a.lower() in ("projeto", "codigo", "código", "obra", "id"):
            continue
        if a not in seen:
            projetos.append(a)
            seen.add(a)

    return projetos


def _drive_try_delete_or_trash(drive_service, fid: str):
    try:
        drive_service.files().delete(fileId=fid, supportsAllDrives=True).execute()
        return "deleted"
    except HttpError as e:
        status = getattr(e.resp, "status", None)
        if status == 404:
            return "notfound"
        if status == 403:
            drive_service.files().update(
                fileId=fid,
                body={"trashed": True},
                supportsAllDrives=True,
                fields="id,trashed"
            ).execute()
            return "trashed"
        raise


def drive_upload_or_update_unique_csv(drive_service, folder_id: str, filepath: str, filename: str) -> str:
    files = drive_list_files_by_exact_name_in_folder(drive_service, folder_id, filename)
    media = MediaFileUpload(filepath, mimetype="text/csv", resumable=True)
    kept_id = None

    if files:
        keep = files[0]
        keep_id = keep["id"]
        log(f"♻️ Atualizando arquivo existente no Drive (overwrite): {filename} | id={keep_id}")

        def _call_update():
            return drive_service.files().update(
                fileId=keep_id,
                media_body=media,
                supportsAllDrives=True,
                fields="id,webViewLink"
            ).execute()

        updated = retry(_call_update)
        kept_id = updated["id"]
        log(f"✅ Atualizado. Link: {updated.get('webViewLink')}")
    else:
        log(f"☁️ Criando novo arquivo no Drive: {filename}")
        metadata = {"name": filename, "parents": [folder_id], "mimeType": "text/csv"}

        def _call_create():
            return drive_service.files().create(
                body=metadata,
                media_body=media,
                supportsAllDrives=True,
                fields="id,webViewLink"
            ).execute()

        created = retry(_call_create)
        kept_id = created["id"]
        log(f"✅ Criado. Link: {created.get('webViewLink')}")

    files2 = drive_list_files_by_exact_name_in_folder(drive_service, folder_id, filename)
    extras = [f for f in files2 if f.get("id") != kept_id]
    if extras:
        log(f"🧹 Encontrados {len(extras)} duplicados com o mesmo nome. Tentando remover...")
        for f in extras:
            fid = f["id"]
            mtime = f.get("modifiedTime")
            size = f.get("size")
            log(f"   - removendo duplicado: id={fid} | modifiedTime={mtime} | size={size}")
            try:
                result = retry(lambda: _drive_try_delete_or_trash(drive_service, fid))
                if result == "deleted":
                    log("     ✅ removido")
                elif result == "trashed":
                    log("     ✅ enviado para lixeira")
                else:
                    log("     ⚠️ 404/sem acesso — não removido, seguindo")
            except HttpError as e:
                status = getattr(e.resp, "status", None)
                log(f"     ❌ erro removendo duplicado (HTTP {status}): {e}")

    return kept_id


# ===================== CONSULTA ATIVIDADES =====================

def ler_planilhas_origem_bd_config(sheets_service):
    log(f"🔄 Lendo lista de planilhas em {BD_CONFIG_RANGE_SOURCES} (planilha 1Ipp...)...")
    raw = read_single_column(sheets_service, ID_ATIVIDADES, BD_CONFIG_RANGE_SOURCES)

    ids = []
    for v in raw:
        sid = extrair_sheet_id(v)
        if sid:
            ids.append(sid)

    seen = set()
    unique = []
    for sid in ids:
        if sid not in seen:
            unique.append(sid)
            seen.add(sid)

    log(f"✅ Planilhas para consultar (únicas): {len(unique)}")
    return unique


def coletar_atividades_em_planilha(sheets_service, spreadsheet_id: str, projetos_norm: set, projetos_base: set):
    vrs = read_batch_columns(sheets_service, spreadsheet_id, [RANGE_COL_I, RANGE_COL_D, RANGE_COL_G])

    col_i = vrs[0].get("values", []) if len(vrs) > 0 else []
    col_d = vrs[1].get("values", []) if len(vrs) > 1 else []
    col_g = vrs[2].get("values", []) if len(vrs) > 2 else []

    max_len = max(len(col_i), len(col_d), len(col_g))

    linhas = []
    for idx in range(max_len):
        proj = clean(col_i[idx][0]) if idx < len(col_i) and col_i[idx] else ""
        if not proj:
            continue

        proj_n = norm_key(proj)
        proj_b = base_key(proj)

        atividade = clean(col_d[idx][0]) if idx < len(col_d) and col_d[idx] else ""
        qtd = clean(col_g[idx][0]) if idx < len(col_g) and col_g[idx] else ""

        if proj_n in projetos_norm or proj_b in projetos_base:
            linhas.append([proj, atividade, qtd])

    return linhas, len(col_i)


# ===================== MAIN =====================

def main():
    log("🚀 Iniciando geração do CSV de atividades...")
    sheets_service, drive_service = get_services()

    chosen = None

    log("🔎 Localizando no Drive a lista de projetos (modelo_importar_obras_ponto)...")
    for fname in LISTA_PROJETOS_FILENAMES_TRY:
        f = drive_get_latest_file_by_name(drive_service, DRIVE_FOLDER_ID, fname)
        if f:
            chosen = f
            break

    if not chosen:
        raise RuntimeError(
            "❌ Não encontrei no Drive o arquivo de lista de projetos: "
            f"{' OU '.join(LISTA_PROJETOS_FILENAMES_TRY)} "
            f"(na pasta ID {DRIVE_FOLDER_ID})."
        )

    rel_id = chosen["id"]
    rel_name = chosen["name"]
    rel_mtime = chosen.get("modifiedTime")

    log("\n✅ Lista escolhida:")
    log(f"   Nome: {rel_name}")
    log(f"   Modificado: {rel_mtime}")
    log(f"   FileID: {rel_id}")

    log("\n⬇️ Baixando lista...")
    rel_bytes = drive_download_file_to_bytes(drive_service, rel_id)

    log("\n🧾 Lendo projetos da coluna A da lista...")
    projetos = parse_projetos_from_relatorio_csv(rel_bytes)
    log(f"✅ Projetos na lista: {len(projetos)}")

    if not projetos:
        raise RuntimeError("❌ A lista não retornou nenhum projeto na coluna A.")

    projetos_norm = set(norm_key(p) for p in projetos if norm_key(p))
    projetos_base = set(base_key(p) for p in projetos if base_key(p))

    planilhas_origem = ler_planilhas_origem_bd_config(sheets_service)

    log("\n🔎 Consultando ATIVIDADES_POR_PONTO nas planilhas listadas...")
    linhas_csv = [CSV_HEADER_FINAL]

    encontrados_total = 0
    planilhas_ok = 0
    planilhas_erro = 0

    for idx, sid in enumerate(planilhas_origem, start=1):
        log(f"   ➜ ({idx}/{len(planilhas_origem)}) Planilha: {sid}")
        try:
            linhas, rows_i = coletar_atividades_em_planilha(
                sheets_service,
                sid,
                projetos_norm,
                projetos_base
            )
            planilhas_ok += 1
            linhas_csv.extend(linhas)
            encontrados_total += len(linhas)
            log(f"     ✅ Col I lida: {rows_i} | Linhas adicionadas: {len(linhas)}")

        except HttpError as e:
            planilhas_erro += 1
            status = getattr(e.resp, "status", None)
            log(f"     ❌ Erro HTTP {status}: {e}")
        except Exception as e:
            planilhas_erro += 1
            log(f"     ❌ Erro: {type(e).__name__}: {e}")

    csv_path = write_csv_local(linhas_csv, CSV_FINAL_NAME)

    log("\n================= RESUMO =================")
    log(f"Lista usada:                 {rel_name}")
    log(f"Planilhas OK:                {planilhas_ok}")
    log(f"Planilhas com erro:          {planilhas_erro}")
    log(f"Projetos na lista:           {len(projetos)}")
    log(f"Linhas (registros) geradas:  {encontrados_total}")
    log(f"CSV local:                   {csv_path}")

    log("\n♻️ Garantindo arquivo único no Drive (overwrite/update)...")
    file_id = drive_upload_or_update_unique_csv(
        drive_service,
        DRIVE_FOLDER_ID,
        csv_path,
        CSV_FINAL_NAME
    )
    log(f"✅ Concluído. FileID final: {file_id}")


if __name__ == "__main__":
    main()
