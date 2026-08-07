import json
import re
import unicodedata
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


PASTA = Path(__file__).resolve().parent
PLANILHA_OPCOES = next(PASTA.glob("OPC*ACABAMENTOS*.xlsx"))
PLANILHA_CATALOGO = next(PASTA.glob("PLANILHA_DE_ACABAMENTOS*.xlsx"))
ARQUIVO_SAIDA = PASTA / "acabamentos_unidades.json"
NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def normalizar(texto):
    texto = "".join(
        caractere for caractere in unicodedata.normalize("NFD", str(texto or ""))
        if unicodedata.category(caractere) != "Mn"
    ).upper()
    texto = re.sub(r"\b0+(\d+)\b", r"\1", texto)
    return re.sub(r"[^A-Z0-9]+", " ", texto).strip()


def numero_coluna(referencia):
    numero = 0
    for letra in re.match(r"[A-Z]+", referencia).group():
        numero = numero * 26 + ord(letra) - 64
    return numero


def ler_xlsx(caminho):
    folhas = {}
    with ZipFile(caminho) as arquivo:
        compartilhadas = []
        if "xl/sharedStrings.xml" in arquivo.namelist():
            raiz = ET.fromstring(arquivo.read("xl/sharedStrings.xml"))
            compartilhadas = [
                "".join(no.text or "" for no in item.iter(f"{{{NS['m']}}}t"))
                for item in raiz.findall("m:si", NS)
            ]
        workbook = ET.fromstring(arquivo.read("xl/workbook.xml"))
        relacoes_raiz = ET.fromstring(arquivo.read("xl/_rels/workbook.xml.rels"))
        relacoes = {item.attrib["Id"]: item.attrib["Target"] for item in relacoes_raiz}
        for folha in workbook.find("m:sheets", NS):
            alvo = relacoes[folha.attrib[f"{{{NS['r']}}}id"]]
            raiz_folha = ET.fromstring(arquivo.read("xl/" + alvo.lstrip("/")))
            linhas = []
            for linha in raiz_folha.findall(".//m:sheetData/m:row", NS):
                valores = {}
                for celula in linha.findall("m:c", NS):
                    valor = celula.find("m:v", NS)
                    interno = celula.find("m:is", NS)
                    if valor is None and interno is None:
                        continue
                    if interno is not None:
                        conteudo = "".join(no.text or "" for no in interno.iter(f"{{{NS['m']}}}t"))
                    else:
                        conteudo = valor.text or ""
                        if celula.attrib.get("t") == "s":
                            conteudo = compartilhadas[int(conteudo)]
                    valores[numero_coluna(celula.attrib["r"])] = str(conteudo).strip()
                linhas.append(valores)
            folhas[folha.attrib["name"]] = linhas
    return folhas


def catalogos_por_folha():
    catalogos = {}
    for nome, linhas in ler_xlsx(PLANILHA_CATALOGO).items():
        secoes = {}
        secao = ""
        for indice, linha in enumerate(linhas):
            titulo = linha.get(1, "").strip()
            descricao = linha.get(2, "").strip()
            if not titulo:
                continue
            proxima_opcao = indice + 1 < len(linhas) and re.match(
                r"OPÇÃO\s*([0-9]+)", linhas[indice + 1].get(1, ""), re.IGNORECASE
            )
            if not descricao and proxima_opcao:
                secao = titulo
                secoes[normalizar(secao)] = {}
                continue
            opcao = re.match(r"OPÇÃO\s*([0-9]+)", titulo, re.IGNORECASE)
            if secao and opcao and descricao:
                secoes[normalizar(secao)][int(opcao.group(1))] = descricao
        catalogos[nome] = secoes
    return catalogos


def folha_para_unidade(numero):
    final = int(numero[-1])
    andar = int(numero[:-2]) if len(numero) >= 3 else 0
    if andar >= 30 and final == 1:
        return "PENT FINAL 1"
    if andar >= 30 and final == 6:
        return "PENT FINAL 6"
    if final in (1, 2):
        return "TIPO FINAIS 1-2"
    if final in (3, 4, 9, 0) or (final == 5 and andar <= 2):
        return "TIPO FINAIS 3-4-9-10 E 5 1,2PAV"
    if final in (5, 8):
        return "TIPO FINAIS 5-8"
    return "TIPO FINAIS 6-7"


def chave_catalogo(ambiente, item):
    ambiente = re.sub(r"\b0+(\d+)\b", r"\1", ambiente.strip())
    item = item.strip().replace("Paredes Fora", "Parede Fora")
    return normalizar(f"{ambiente} - {item}")


def palavras(texto):
    retorno = set()
    for palavra in normalizar(texto).split():
        if len(palavra) > 3 and palavra.endswith("S"):
            palavra = palavra[:-1]
        if palavra not in {"DE", "DO", "DA", "E", "FORA"}:
            retorno.add(palavra)
    return retorno


def categoria_item(item):
    item_normalizado = normalizar(item)
    if "FECHADURA" in item_normalizado:
        return {"FECHADURA"}
    if "LOUCA" in item_normalizado:
        return {"LOUCA"}
    if "METAI" in item_normalizado:
        return {"METAI"}
    if "TAMPO" in item_normalizado or "CUBA" in item_normalizado or "BITS" in item_normalizado:
        return {"TAMPO"}
    if "PAREDE" in item_normalizado and "BOX" in item_normalizado:
        return {"PAREDE", "BOX"}
    if "PAREDE" in item_normalizado and "PISO" not in item_normalizado:
        return {"PAREDE"}
    if "PISO" in item_normalizado or "RODAPE" in item_normalizado:
        return {"PISO"}
    return palavras(item)


def buscar_descricao(catalogo, ambiente, item, numero_opcao):
    chave_exata = chave_catalogo(ambiente, item)
    if numero_opcao in catalogo.get(chave_exata, {}):
        return catalogo[chave_exata][numero_opcao]
    ambiente_palavras = palavras(ambiente)
    categoria = categoria_item(item)
    candidatos = []
    for secao, opcoes in catalogo.items():
        if numero_opcao not in opcoes:
            continue
        secao_palavras = palavras(secao)
        if not categoria.issubset(secao_palavras):
            continue
        encontrados = len(ambiente_palavras & secao_palavras)
        ausentes = len(ambiente_palavras - secao_palavras)
        if encontrados == 0:
            continue
        pontos = encontrados * 12 - ausentes * 10 - len(secao_palavras - ambiente_palavras - categoria)
        candidatos.append((pontos, opcoes[numero_opcao]))
    if not candidatos:
        return ""
    candidatos.sort(key=lambda candidato: candidato[0], reverse=True)
    return candidatos[0][1]


def gerar():
    catalogos = catalogos_por_folha()
    resultado = {}
    for nome_folha, linhas in ler_xlsx(PLANILHA_OPCOES).items():
        if len(linhas) < 3:
            continue
        torre = "horizonte" if nome_folha.endswith("T1") else "aurora"
        grupos, atual = {}, ""
        for coluna in range(3, max(linhas[0], default=3) + 1):
            if linhas[0].get(coluna):
                atual = linhas[0][coluna]
            grupos[coluna] = atual
        itens = linhas[1]
        for linha in linhas[2:]:
            numero = linha.get(1, "").strip()
            if not re.fullmatch(r"\d{3,4}", numero):
                continue
            acabamentos = []
            catalogo = catalogos.get(folha_para_unidade(numero), {}) if torre == "aurora" else {}
            for coluna in range(4, max(itens, default=4) + 1):
                escolha = linha.get(coluna, "").strip()
                if not escolha or escolha == "-":
                    continue
                ambiente, item = grupos.get(coluna, ""), itens.get(coluna, "")
                if any(termo in normalizar(item) for termo in ("VALOR CONTRATO", "ADITIVO", "MULTA PERSONALIZACAO")):
                    continue
                opcao = re.search(r"(\d+)", escolha)
                descricao = ""
                if opcao:
                    descricao = buscar_descricao(catalogo, ambiente, item, int(opcao.group(1)))
                acabamentos.append({
                    "ambiente": ambiente.strip(),
                    "item": item.strip(),
                    "opcao": escolha,
                    "descricao": descricao,
                })
            resultado[f"{torre}|Apto {numero}"] = {
                "planta": linha.get(3, "").strip(),
                "itens": acabamentos,
            }
    temporario = ARQUIVO_SAIDA.with_suffix(".json.tmp")
    temporario.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    temporario.replace(ARQUIVO_SAIDA)
    detalhados = sum(bool(item["descricao"]) for unidade in resultado.values() for item in unidade["itens"])
    total = sum(len(unidade["itens"]) for unidade in resultado.values())
    print(f"{len(resultado)} unidades; {total} escolhas; {detalhados} especificações detalhadas")


if __name__ == "__main__":
    gerar()
