# -*- coding: utf-8 -*-
"""
sync_macro.py - Sincroniza o script ImportadorSAGE.py com a macro embutida
dentro de SageBonis.ods (Scripts/python/ImportadorSAGE.py).

Um arquivo .ods é apenas um ZIP. Esta ferramenta troca o script embutido
sem precisar abrir o LibreOffice, preservando a estrutura exigida pelo ODF
(o 'mimetype' precisa ser a primeira entrada e ficar SEM compressao).

Uso:
    python sync_macro.py status     # mostra se o .py e a macro embutida diferem
    python sync_macro.py inject      # .py do disco  -> macro dentro do .ods
    python sync_macro.py extract     # macro do .ods -> .py do disco

Opcionais:
    --ods   CAMINHO   (padrao: SageBonis.ods)
    --py    CAMINHO   (padrao: ImportadorSAGE.py)
    --no-backup       nao cria .bak antes de gravar (inject)
"""

import argparse
import difflib
import os
import shutil
import sys
import zipfile

CAMINHO_INTERNO = "Scripts/python/ImportadorSAGE.py"


def _ler_embutido(ods_path):
    with zipfile.ZipFile(ods_path, "r") as z:
        nomes = z.namelist()
        if CAMINHO_INTERNO not in nomes:
            sys.exit(f"ERRO: '{CAMINHO_INTERNO}' nao existe dentro de {ods_path}.")
        return z.read(CAMINHO_INTERNO).decode("utf-8")


def _ler_disco(py_path):
    if not os.path.isfile(py_path):
        sys.exit(f"ERRO: arquivo nao encontrado: {py_path}")
    with open(py_path, "r", encoding="utf-8") as f:
        return f.read()


def cmd_status(ods_path, py_path):
    embutido = _ler_embutido(ods_path)
    disco = _ler_disco(py_path)
    if embutido == disco:
        print(f"OK: {py_path} esta identico a macro embutida em {ods_path}.")
        return
    print(f"DIFERENTE: {py_path} e a macro embutida em {ods_path} divergem.\n")
    diff = difflib.unified_diff(
        disco.splitlines(), embutido.splitlines(),
        fromfile=f"disco:{py_path}", tofile=f"ods:{CAMINHO_INTERNO}", lineterm="",
    )
    n = 0
    for linha in diff:
        print(linha)
        n += 1
        if n > 200:
            print("... (diff truncado em 200 linhas)")
            break


def cmd_extract(ods_path, py_path):
    embutido = _ler_embutido(ods_path)
    with open(py_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(embutido)
    print(f"OK: macro embutida extraida para {py_path}.")


def cmd_inject(ods_path, py_path, fazer_backup=True):
    novo = _ler_disco(py_path).encode("utf-8")

    if fazer_backup:
        bak = ods_path + ".bak"
        shutil.copy2(ods_path, bak)
        print(f"Backup criado: {bak}")

    tmp = ods_path + ".tmp"
    with zipfile.ZipFile(ods_path, "r") as zin:
        infos = zin.infolist()
        # Reordena para garantir que 'mimetype' seja a primeira entrada.
        mimetype = [i for i in infos if i.filename == "mimetype"]
        resto = [i for i in infos if i.filename != "mimetype"]
        ordenadas = mimetype + resto

        with zipfile.ZipFile(tmp, "w") as zout:
            substituido = False
            for info in ordenadas:
                dados = novo if info.filename == CAMINHO_INTERNO else zin.read(info.filename)
                if info.filename == CAMINHO_INTERNO:
                    substituido = True
                # mimetype precisa ficar STORED (sem compressao); o resto mantem o original.
                compress = zipfile.ZIP_STORED if info.filename == "mimetype" else (
                    info.compress_type or zipfile.ZIP_DEFLATED
                )
                novo_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                novo_info.compress_type = compress
                novo_info.external_attr = info.external_attr
                novo_info.internal_attr = info.internal_attr
                novo_info.create_system = info.create_system
                zout.writestr(novo_info, dados)

    if not substituido:
        os.remove(tmp)
        sys.exit(f"ERRO: '{CAMINHO_INTERNO}' nao existe dentro de {ods_path}; nada injetado.")

    os.replace(tmp, ods_path)
    print(f"OK: {py_path} injetado em {ods_path} ({CAMINHO_INTERNO}).")
    print("Abra o LibreOffice e habilite as macros do documento para testar.")


def main():
    p = argparse.ArgumentParser(description="Sincroniza ImportadorSAGE.py com a macro embutida em SageBonis.ods")
    p.add_argument("acao", choices=["status", "inject", "extract"])
    p.add_argument("--ods", default="SageBonis.ods")
    p.add_argument("--py", default="ImportadorSAGE.py")
    p.add_argument("--no-backup", action="store_true")
    a = p.parse_args()

    if a.acao == "status":
        cmd_status(a.ods, a.py)
    elif a.acao == "extract":
        cmd_extract(a.ods, a.py)
    elif a.acao == "inject":
        cmd_inject(a.ods, a.py, fazer_backup=not a.no_backup)


if __name__ == "__main__":
    main()
