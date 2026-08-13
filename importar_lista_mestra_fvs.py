#!/usr/bin/env python3
"""Adiciona somente os serviços e critérios novos da Lista Mestre FVS.

Não sobrescreve itens existentes. Se o nome do serviço já foi customizado no banco,
este script ignora o serviço em vez de apagar ou alterar o que já existe.
"""

from __future__ import annotations

import re
import sqlite3
import unicodedata
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

PASTA = Path(__file__).resolve().parent
PLANILHA = PASTA / "Lista_Mestra_FVS_completa.xlsx"
BANCO = PASTA / "acompanhamento.db"
NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def normalizar_texto(valor: str) -> str:
    texto = unicodedata.normalize("NFD", str(valor or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto).strip().lower()


def ler_shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
    itens = []
    for si in raiz.findall("a:si", NS):
        texto = "".join(
            el.text or "" for el in si.iter(f"{{{NS['a']}}}t")
        )
        itens.append(texto)
    return itens


def ler_planilha_em_rows(caminho: Path):
    with ZipFile(caminho) as z:
        shared_strings = ler_shared_strings(z)
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target_map = {item.attrib["Id"]: item.attrib["Target"] for item in rels}
        workbook = ET.fromstring(z.read("xl/workbook.xml"))

        for sheet in workbook.findall("a:sheets/a:sheet", NS):
            nome = sheet.attrib["name"]
            rel_id = sheet.attrib[f"{{{NS['r']}}}id"]
            target = target_map[rel_id]
            if not target.startswith("xl/"):
                target = "xl/" + target
            planilha = ET.fromstring(z.read(target))
            linhas = []
            for linha in planilha.findall(".//a:sheetData/a:row", NS):
                valores = []
                for celula in linha.findall("a:c", NS):
                    tipo = celula.attrib.get("t")
                    valor_celula = celula.find("a:v", NS)
                    valor = "" if valor_celula is None else (valor_celula.text or "")
                    if tipo == "s" and valor:
                        indice = int(valor)
                        valor = shared_strings[indice] if 0 <= indice < len(shared_strings) else ""
                    valores.append(str(valor).strip())
                if any(v for v in valores):
                    linhas.append(valores)
            yield nome, linhas


def extrair_servicos_e_criterios(planilha: Path):
    servicos = {}
    for nome_planilha, linhas in ler_planilha_em_rows(planilha):
        if not linhas:
            continue
        if not linhas[0] or not any("FVS" in str(v).upper() for v in linhas[0]):
            continue
        titulo = next((valor for valor in linhas[0] if valor.strip()), "")
        if " - " not in titulo:
            continue
        nome_servico = titulo.split(" - ", 1)[1].strip()
        nome_servico = nome_servico.title()
        criterios = []
        for linha in linhas[1:]:
            if not linha:
                continue
            primeiro = linha[0].strip()
            if not primeiro or primeiro.lower().startswith("item / critério"):
                continue
            if primeiro.lower().startswith("fvs ") or primeiro.lower().startswith("código fvs"):
                continue
            criterios.append(primeiro)
        if nome_servico and criterios:
            servicos.setdefault(nome_servico, [])
            for criterio in criterios:
                if criterio not in servicos[nome_servico]:
                    servicos[nome_servico].append(criterio)
    return servicos


def atualizar_banco():
    if not PLANILHA.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {PLANILHA}")

    livros = extrair_servicos_e_criterios(PLANILHA)
    if not livros:
        raise RuntimeError("Nenhum serviço novo foi encontrado na Lista Mestre FVS.")

    conexao = sqlite3.connect(BANCO)
    conexao.row_factory = sqlite3.Row
    conexao.execute(
        """
        CREATE TABLE IF NOT EXISTS atividades_config (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            torre TEXT NOT NULL,
            andar INTEGER NOT NULL,
            unidade TEXT NOT NULL DEFAULT '*',
            criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(nome, torre, andar, unidade)
        )
        """
    )
    conexao.execute(
        """
        CREATE TABLE IF NOT EXISTS atividades_especificacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            atividade TEXT NOT NULL,
            especificacao TEXT NOT NULL,
            ordem INTEGER NOT NULL DEFAULT 0,
            UNIQUE(atividade, especificacao)
        )
        """
    )

    nomes_existentes = {
        normalizar_texto(nome[0])
        for nome in conexao.execute("SELECT DISTINCT nome FROM atividades_config").fetchall()
    }

    criados = 0
    esperados = 0
    for nome_servico, criterios in livros.items():
        chave = normalizar_texto(nome_servico)
        if chave in nomes_existentes:
            continue

        esperados += 1
        for torre in ("aurora", "horizonte"):
            max_andar = 36 if torre == "aurora" else 23
            andares = [-2, -1, 0] + list(range(1, max_andar + 1)) + [101, 102, 103]
            for andar in andares:
                conexao.execute(
                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                    (nome_servico, torre, andar),
                )
        for ordem, criterio in enumerate(criterios):
            conexao.execute(
                "INSERT OR IGNORE INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES (?, ?, ?)",
                (nome_servico, criterio, ordem),
            )
        criados += 1

    conexao.commit()
    conexao.close()
    print(f"Serviços novos incluídos: {criados} de {esperados} esperados.")


if __name__ == "__main__":
    atualizar_banco()
