import json
import re
import shutil
import subprocess
import tempfile
import unicodedata
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from PIL import Image, ImageChops
from pypdf import PdfReader, PdfWriter


PASTA = Path(__file__).resolve().parent
PLANILHA = next(PASTA.glob("RELA*PLANTA*.xlsx"))
PASTA_SAIDA = PASTA / "miniaturas_tipos_planta"
ARQUIVO_SAIDA = PASTA / "plantas_unidades.json"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def normalizar(texto):
    texto = "".join(
        caractere for caractere in unicodedata.normalize("NFD", str(texto or ""))
        if unicodedata.category(caractere) != "Mn"
    ).upper().replace("º", "")
    return re.sub(r"\s+", " ", texto).strip()


def coluna_numero(referencia):
    letras = re.match(r"[A-Z]+", referencia).group()
    numero = 0
    for letra in letras:
        numero = numero * 26 + ord(letra) - 64
    return numero


def ler_planilha():
    folhas = {}
    with ZipFile(PLANILHA) as arquivo:
        compartilhadas = []
        raiz = ET.fromstring(arquivo.read("xl/sharedStrings.xml"))
        for item in raiz.findall("m:si", NS):
            compartilhadas.append("".join(no.text or "" for no in item.iter(f"{{{NS['m']}}}t")))
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
                    if valor is None:
                        continue
                    conteudo = valor.text or ""
                    if celula.attrib.get("t") == "s":
                        conteudo = compartilhadas[int(conteudo)]
                    valores[coluna_numero(celula.attrib["r"])] = conteudo
                linhas.append(valores)
            folhas[folha.attrib["name"]] = linhas
    return folhas


def metadados_paginas(pdf):
    paginas = []
    leitor = PdfReader(pdf)
    for indice, pagina in enumerate(leitor.pages):
        texto_original = " ".join((pagina.extract_text() or "").split())
        texto = normalizar(texto_original)
        final_encontrado = re.search(r"APTO FINAL\s*0?(\d+)", texto)
        final = int(final_encontrado.group(1)) if final_encontrado else None
        trecho = texto[final_encontrado.end():final_encontrado.end() + 100] if final_encontrado else ""
        intervalo = re.search(r"(\d+)\s*AO\s*(\d+)", trecho)
        conjunto = re.search(r"(\d+)\s*E\s*(\d+)", trecho)
        unico = re.search(r"(\d+)\s*PAV", trecho)
        if intervalo:
            andares = set(range(int(intervalo.group(1)), int(intervalo.group(2)) + 1))
        elif conjunto:
            andares = {int(conjunto.group(1)), int(conjunto.group(2))}
        elif unico:
            andares = {int(unico.group(1))}
        else:
            andares = set()
        opcao = None
        opcao_encontrada = re.search(r"OPCAO\s*0?(\d+)", texto)
        if opcao_encontrada:
            opcao = int(opcao_encontrada.group(1))
        adaptada = "ADAPTADA" in texto or "ACESSIVEL" in texto
        titulo = ""
        titulo_inicio = texto.find("OPCAO")
        titulo_fim = texto.find("APTO FINAL")
        if titulo_inicio >= 0 and titulo_fim > titulo_inicio:
            titulo = texto_original[titulo_inicio:titulo_fim].strip()
        paginas.append({"indice": indice, "final": final, "andares": andares, "opcao": opcao, "adaptada": adaptada, "texto": texto, "titulo": titulo})
    return leitor, paginas


def pagina_da_unidade(paginas, final, andar, descricao):
    descricao_normalizada = normalizar(descricao)
    adaptada = "ACESSIVEL" in descricao_normalizada or "ADAPTADA" in descricao_normalizada
    opcao_encontrada = re.search(r"OPCAO\s*0?(\d+)", descricao_normalizada)
    opcao = int(opcao_encontrada.group(1)) if opcao_encontrada else None
    candidatos = [pagina for pagina in paginas if pagina["final"] == final]
    pontuados = []
    for pagina in candidatos:
        pontos = 0
        if andar in pagina["andares"]:
            pontos += 20
        elif pagina["andares"]:
            pontos -= 20
        if adaptada == pagina["adaptada"]:
            pontos += 10
        if opcao is not None and pagina["opcao"] == opcao:
            pontos += 8
        elif opcao is not None:
            pontos -= 8
        pontuados.append((pontos, pagina))
    if not pontuados:
        return None
    pontuados.sort(key=lambda item: item[0], reverse=True)
    return pontuados[0][1]


def renderizar_pagina(leitor, indice, identificador):
    PASTA_SAIDA.mkdir(exist_ok=True)
    destino = PASTA_SAIDA / f"{identificador}.png"
    if destino.exists() and destino.stat().st_mtime >= Path(__file__).stat().st_mtime:
        return f"/{PASTA_SAIDA.name}/{destino.name}"
    with tempfile.TemporaryDirectory() as temporaria:
        temporaria = Path(temporaria)
        pagina_pdf = temporaria / "pagina.pdf"
        escritor = PdfWriter()
        escritor.add_page(leitor.pages[indice])
        with pagina_pdf.open("wb") as saida:
            escritor.write(saida)
        subprocess.run(
            ["qlmanage", "-t", "-s", "1800", "-o", str(temporaria), str(pagina_pdf)],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        imagem_origem = next(temporaria.glob("*.png"))
        imagem = Image.open(imagem_origem).convert("RGB")
        largura, altura = imagem.size
        # Mantém a planta e eventuais legendas, removendo o carimbo/título inferior.
        imagem = imagem.crop((int(largura * .02), int(altura * .03), int(largura * .98), int(altura * .96)))
        cinza = imagem.convert("L")
        limite_inicial = int(imagem.height * .72)
        limite_final = int(imagem.height * .95)
        linhas = []
        for y in range(limite_inicial, limite_final):
            histograma = cinza.crop((0, y, imagem.width, y + 1)).histogram()
            proporcao_escura = sum(histograma[:255]) / imagem.width
            if proporcao_escura >= .64:
                linhas.append(y)
        if linhas:
            imagem = imagem.crop((0, 0, imagem.width, max(1, linhas[0] - 78)))
        fundo = Image.new("RGB", imagem.size, "white")
        diferenca = ImageChops.difference(imagem, fundo).convert("L")
        limite = diferenca.point(lambda valor: 255 if valor > 18 else 0).getbbox()
        if limite:
            margem = 18
            esquerda = max(0, limite[0] - margem)
            topo = max(0, limite[1] - margem)
            direita = min(imagem.width, limite[2] + margem)
            base = min(imagem.height, limite[3] + margem)
            imagem = imagem.crop((esquerda, topo, direita, base))
        imagem.thumbnail((1100, 850), Image.Resampling.LANCZOS)
        imagem.save(destino, "PNG", optimize=True)
    return f"/{PASTA_SAIDA.name}/{destino.name}"


def numero_unidade(andar, final, torre):
    if torre == "aurora" and andar == 36:
        especiais = {1: "3001", 3: "3003", 4: "3004", 5: "3005", 6: "3006", 8: "3008", 9: "3009", 10: "3010"}
        return especiais.get(final, "")
    largura = 2 if torre == "horizonte" else 2
    return f"{andar}{final:0{largura}d}"


def gerar():
    folhas = ler_planilha()
    configuracoes = [
        ("TORRE 1 - SMART", "horizonte", PASTA / "T1-R00.pdf"),
        ("TORRE 2 - HOME", "aurora", PASTA / "T2-R00.pdf"),
    ]
    plantas_alteradas = json.loads((PASTA / "personalizacoes_unidades.json").read_text(encoding="utf-8"))
    resultado = {}
    paginas_usadas = set()
    for nome_folha, torre, pdf in configuracoes:
        leitor, paginas = metadados_paginas(pdf)
        linhas = folhas[nome_folha]
        cabecalho = next(linha for linha in linhas if linha.get(1) == "Andar")
        finais = {coluna: int(valor) for coluna, valor in cabecalho.items() if coluna > 1 and str(valor).isdigit()}
        for linha in linhas:
            andar_encontrado = re.match(r"(\d+)", str(linha.get(1, "")))
            if not andar_encontrado:
                continue
            andar = int(andar_encontrado.group(1))
            for coluna, final in finais.items():
                descricao = str(linha.get(coluna, "")).strip()
                if not descricao or descricao == "*":
                    continue
                numero = numero_unidade(andar, final, torre)
                if not numero:
                    continue
                chave = f"{torre}|Apto {numero}"
                if "PROJETO ALTERADO" in normalizar(descricao):
                    miniatura = plantas_alteradas.get(chave, {}).get("plantaMiniatura", "")
                    resultado[chave] = {"tipo": "Projeto alterado", "plantaMiniatura": miniatura, "origem": "Projeto específico da unidade"}
                    continue
                pagina = pagina_da_unidade(paginas, final, andar, descricao)
                if not pagina:
                    continue
                identificador = f"{torre}-t{1 if torre == 'horizonte' else 2}-p{pagina['indice'] + 1:02d}"
                miniatura = renderizar_pagina(leitor, pagina["indice"], identificador)
                paginas_usadas.add((torre, pagina["indice"] + 1))
                resultado[chave] = {"tipo": descricao, "plantaMiniatura": miniatura, "origem": f"{pdf.name} · página {pagina['indice'] + 1}"}
    temporario = ARQUIVO_SAIDA.with_suffix(".json.tmp")
    temporario.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    temporario.replace(ARQUIVO_SAIDA)
    sem_imagem = sum(not item["plantaMiniatura"] for item in resultado.values())
    print(f"{len(resultado)} unidades mapeadas; {len(paginas_usadas)} páginas utilizadas; {sem_imagem} sem miniatura")


if __name__ == "__main__":
    gerar()
