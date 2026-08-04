#!/usr/bin/env python3
"""Converte a planilha de controle da obra para o formato usado pelo painel."""

import json
import hashlib
import re
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
TORRES = {"TORRE 1": "horizonte", "TORRE 2": "aurora"}
MAX_ANDARES = {"horizonte": 23, "aurora": 36}
STATUS = {
    "EXECUTADO": "concluido",
    "EM EXECUÇÃO": "em-andamento",
    "PROJEÇÃO MÊS": "nao-iniciado",
    "PENDÊNCIA": "pendente",
}
VERSAO_CONVERSOR = "5"


def coluna_numero(referencia):
    letras = re.match(r"[A-Z]+", referencia).group()
    numero = 0
    for letra in letras:
        numero = numero * 26 + ord(letra) - 64
    return numero


def ler_celulas(arquivo):
    with ZipFile(arquivo) as zipado:
        compartilhadas = []
        raiz = ET.fromstring(zipado.read("xl/sharedStrings.xml"))
        for item in raiz.findall("m:si", NS):
            compartilhadas.append(
                "".join(texto.text or "" for texto in item.iter(f"{{{NS['m']}}}t"))
            )

        workbook = ET.fromstring(zipado.read("xl/workbook.xml"))
        relacionamentos = ET.fromstring(zipado.read("xl/_rels/workbook.xml.rels"))
        destinos = {item.attrib["Id"]: item.attrib["Target"] for item in relacionamentos}

        for aba in workbook.find("m:sheets", NS):
            nome = aba.attrib["name"]
            if nome not in TORRES:
                continue
            destino = destinos[aba.attrib[f"{{{NS['r']}}}id"]]
            if not destino.startswith("xl/"):
                destino = "xl/" + destino
            planilha = ET.fromstring(zipado.read(destino))
            celulas = {}
            for celula in planilha.findall(".//m:sheetData/m:row/m:c", NS):
                referencia = celula.attrib["r"]
                valor = celula.find("m:v", NS)
                if valor is None:
                    continue
                conteudo = valor.text or ""
                if celula.attrib.get("t") == "s":
                    conteudo = compartilhadas[int(conteudo)]
                celulas[referencia] = conteudo.strip() if isinstance(conteudo, str) else conteudo
            yield nome, celulas


def converter(arquivo):
    digest = hashlib.sha256(arquivo.read_bytes()).hexdigest()[:16]
    resultado = {"versao": f"{arquivo.name}:{digest}:v{VERSAO_CONVERSOR}", "torres": {}}
    for nome_aba, celulas in ler_celulas(arquivo):
        torre = TORRES[nome_aba]
        linha_servicos = next(
            int(re.search(r"\d+", ref).group())
            for ref, valor in celulas.items()
            if ref.startswith("B") and valor == "SERVIÇO"
        )
        atividades = {
            coluna_numero(ref): valor.strip()
            for ref, valor in celulas.items()
            if int(re.search(r"\d+", ref).group()) == linha_servicos
            and coluna_numero(ref) >= 3
            and valor.strip()
        }
        andares = {}
        for ref, valor_andar in celulas.items():
            if not ref.startswith("B") or not str(valor_andar).isdigit():
                continue
            andar = int(valor_andar)
            if andar < 1 or andar > MAX_ANDARES[torre]:
                continue
            linha = int(re.search(r"\d+", ref).group())
            servicos = []
            for coluna, atividade in atividades.items():
                # Toda célula vazia da matriz representa serviço não iniciado.
                letras = ""
                numero = coluna
                while numero:
                    numero, resto = divmod(numero - 1, 26)
                    letras = chr(65 + resto) + letras
                valor = celulas.get(f"{letras}{linha}", "")
                servicos.append(
                    {"atividade": atividade, "status": STATUS.get(valor, "nao-iniciado")}
                )
            andares[str(andar)] = servicos
        resultado["torres"][torre] = andares
    return resultado


def main():
    if len(sys.argv) != 3:
        raise SystemExit("Uso: importar_planilha_xlsx.py entrada.xlsx saida.json")
    entrada, saida = map(Path, sys.argv[1:])
    dados = converter(entrada)
    saida.write_text(json.dumps(dados, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    total = sum(len(itens) for torre in dados["torres"].values() for itens in torre.values())
    print(f"{total} serviços ativos importados para {saida}")


if __name__ == "__main__":
    main()
