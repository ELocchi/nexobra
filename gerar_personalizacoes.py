import json
import re
import shutil
import subprocess
import tempfile
import unicodedata
from pathlib import Path

from PIL import Image, ImageChops
from pypdf import PdfReader, PdfWriter


PASTA = Path(__file__).resolve().parent
ORIGEM = PASTA / "projetos alterados"
DESTINO = PASTA / "personalizacoes_unidades.json"
PASTA_MINIATURAS = PASTA / "miniaturas_plantas"


def texto_pdf(caminho):
    return " ".join(
        "\n".join((pagina.extract_text() or "") for pagina in PdfReader(caminho).pages).split()
    )


def sem_acentos(texto):
    return "".join(
        caractere for caractere in unicodedata.normalize("NFD", texto)
        if unicodedata.category(caractere) != "Mn"
    ).lower()


def trecho_tecnico(texto):
    normalizado = sem_acentos(texto)
    inicios = [
        "descricoes detalhadas abaixo:",
        "a fim de constar:",
    ]
    posicao = -1
    tamanho = 0
    for marcador in inicios:
        encontrada = normalizado.find(marcador)
        if encontrada >= 0 and (posicao < 0 or encontrada < posicao):
            posicao = encontrada
            tamanho = len(marcador)
    if posicao < 0:
        return ""
    inicio = posicao + tamanho
    finais = [
        "observacoes gerais:",
        "clausula terceira:",
        "valores e condicoes de pagamento:",
        "docusign envelope id:",
    ]
    candidatos = [normalizado.find(marcador, inicio) for marcador in finais]
    candidatos = [valor for valor in candidatos if valor >= 0]
    fim = min(candidatos) if candidatos else min(len(texto), inicio + 3000)
    trecho = texto[inicio:fim].strip(" :-")
    trecho = re.sub(r"Docusign Envelope ID:\s*[A-Z0-9-]+", "", trecho, flags=re.I)
    return trecho.strip()


def nome_documento(caminho):
    nome = caminho.stem.replace("_", " ")
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome


def localizar_contrato_com_planta(pasta):
    """Localiza a planta incorporada ao contrato, sem usar PDFs/DWGs anexos."""
    for arquivo in sorted(pasta.glob("*.pdf")):
        if "PROJETO" in sem_acentos(arquivo.name).upper():
            continue
        try:
            leitor = PdfReader(arquivo)
            if not leitor.pages:
                continue
            for indice in range(len(leitor.pages) - 1, -1, -1):
                texto_pagina = sem_acentos(leitor.pages[indice].extract_text() or "").upper()
                if "ARQUITETURA" in texto_pagina and "ALTERACAO DE PLANTA" in texto_pagina:
                    return arquivo, indice
        except Exception:
            continue
    return None


def localizar_projetos(pasta):
    projetos = []
    for arquivo in sorted(pasta.rglob("*.pdf")):
        nome = sem_acentos(arquivo.name).upper()
        if "PROJETO" in nome or arquivo.parent != pasta or nome.startswith(("ARQ_", "HID_", "ELE_")):
            projetos.append(arquivo)
    return projetos


def tipo_prancha(texto, indice):
    texto = sem_acentos(texto).upper()
    if "HIDRAUL" in texto or "HID_" in texto:
        return "Projeto hidráulico"
    if "ELETR" in texto and ("AUX" in texto or "ILUM" in texto or "TOMADA" in texto):
        return "Projeto elétrico auxiliar"
    if "ELETR" in texto:
        return "Projeto elétrico de distribuição"
    if "ARQUIT" in texto or "ARQ_" in texto:
        return "Projeto arquitetônico"
    nomes = ["Projeto arquitetônico", "Projeto hidráulico", "Projeto elétrico de distribuição", "Projeto elétrico auxiliar"]
    return nomes[indice] if indice < len(nomes) else f"Prancha {indice + 1}"


def gerar_miniatura(contrato_com_pagina, torre, numero):
    if not contrato_com_pagina:
        return ""
    contrato, indice_pagina = contrato_com_pagina
    PASTA_MINIATURAS.mkdir(exist_ok=True)
    nome = f"{torre}-{numero}.png"
    destino = PASTA_MINIATURAS / nome
    if destino.exists() and destino.stat().st_mtime >= Path(__file__).stat().st_mtime:
        return f"/miniaturas_plantas/{nome}"
    with tempfile.TemporaryDirectory() as temporaria:
        temporaria = Path(temporaria)
        pagina_pdf = temporaria / "planta-contrato.pdf"
        leitor = PdfReader(contrato)
        escritor = PdfWriter()
        escritor.add_page(leitor.pages[indice_pagina])
        with pagina_pdf.open("wb") as saida:
            escritor.write(saida)
        subprocess.run(
            ["qlmanage", "-t", "-s", "1800", "-o", str(temporaria), str(pagina_pdf)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        miniaturas = list(temporaria.glob("*.png"))
        if not miniaturas:
            return ""
        shutil.move(str(miniaturas[0]), destino)
    imagem = Image.open(destino).convert("RGB")
    largura, altura = imagem.size
    # A última página contém a prancha contratual; remove carimbo e margens,
    # preservando a área da planta arquitetônica.
    imagem = imagem.crop((int(largura * .03), int(altura * .04), int(largura * .70), int(altura * .96)))
    cinza = imagem.convert("L")
    linhas = []
    for y in range(int(imagem.height * .72), int(imagem.height * .95)):
        histograma = cinza.crop((0, y, imagem.width, y + 1)).histogram()
        if sum(histograma[:255]) / imagem.width >= .64:
            linhas.append(y)
    if linhas:
        imagem = imagem.crop((0, 0, imagem.width, max(1, linhas[0] - 78)))
    fundo = Image.new("RGB", imagem.size, "white")
    limite = ImageChops.difference(imagem, fundo).convert("L").point(lambda valor: 255 if valor > 18 else 0).getbbox()
    if limite:
        margem = 18
        imagem = imagem.crop((max(0, limite[0] - margem), max(0, limite[1] - margem), min(imagem.width, limite[2] + margem), min(imagem.height, limite[3] + margem)))
    imagem.thumbnail((1100, 850), Image.Resampling.LANCZOS)
    imagem.save(destino, "PNG", optimize=True)
    return f"/miniaturas_plantas/{nome}"


def gerar_miniaturas_projetos(pasta, torre, numero):
    projetos = []
    PASTA_MINIATURAS.mkdir(exist_ok=True)
    contador = 0
    for arquivo in localizar_projetos(pasta):
        leitor = PdfReader(arquivo)
        for indice, pagina in enumerate(leitor.pages):
            contador += 1
            nome = f"{torre}-{numero}-projeto-{contador:02d}.png"
            destino = PASTA_MINIATURAS / nome
            if not destino.exists():
                with tempfile.TemporaryDirectory() as temporaria:
                    temporaria = Path(temporaria)
                    pagina_pdf = temporaria / "prancha.pdf"
                    escritor = PdfWriter()
                    escritor.add_page(pagina)
                    with pagina_pdf.open("wb") as saida:
                        escritor.write(saida)
                    subprocess.run(
                        ["qlmanage", "-t", "-s", "1600", "-o", str(temporaria), str(pagina_pdf)],
                        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                    )
                    origem = next(temporaria.glob("*.png"), None)
                    if not origem:
                        continue
                    shutil.move(str(origem), destino)
                imagem = Image.open(destino).convert("RGB")
                largura, altura = imagem.size
                imagem = imagem.crop((int(largura * .015), int(altura * .02), int(largura * .985), int(altura * .94)))
                imagem.thumbnail((1400, 1000), Image.Resampling.LANCZOS)
                imagem.save(destino, "PNG", optimize=True)
            titulo = tipo_prancha((pagina.extract_text() or "") + " " + arquivo.name, indice)
            projetos.append({
                "titulo": titulo,
                "arquivo": nome_documento(arquivo),
                "imagem": f"/miniaturas_plantas/{nome}",
            })
    return projetos


def gerar():
    unidades = {}
    existentes = json.loads(DESTINO.read_text(encoding="utf-8")) if DESTINO.exists() else {}
    for pasta in sorted(ORIGEM.iterdir()):
        if not pasta.is_dir():
            continue
        identificador = re.search(r"APTO\.(\d+)-(\d+)", pasta.name, re.I)
        if not identificador:
            continue
        numero, torre_numero = identificador.groups()
        torre = "aurora" if torre_numero == "2" else "horizonte"
        unidade = f"Apto {numero}"
        andar = int(numero) // 100
        documentos = []
        for pdf in sorted(pasta.glob("*.pdf")):
            if "PROJETO" in sem_acentos(pdf.name).upper():
                continue
            try:
                descricao = trecho_tecnico(texto_pdf(pdf))
            except Exception:
                descricao = ""
            if not descricao and "NAO_INSTALACAO" in sem_acentos(pdf.name).upper():
                descricao = "Não instalação parcial conforme documento contratual da unidade."
            if descricao:
                documentos.append({"titulo": nome_documento(pdf), "descricao": descricao})
        quantidade_projetos = sum(
            1 for arquivo in pasta.rglob("*")
            if arquivo.is_file() and arquivo.suffix.lower() in {".pdf", ".dwg"}
            and (arquivo.parent != pasta or "PROJETO" in sem_acentos(arquivo.name).upper())
        )
        contrato_com_planta = localizar_contrato_com_planta(pasta)
        chave_unidade = f"{torre}|{unidade}"
        projetos = existentes.get(chave_unidade, {}).get("projetos") or gerar_miniaturas_projetos(pasta, torre, numero)
        unidades[chave_unidade] = {
            "torre": torre,
            "andar": andar,
            "unidade": unidade,
            "documentos": documentos,
            "quantidadeProjetos": quantidade_projetos,
            "plantaMiniatura": gerar_miniatura(contrato_com_planta, torre, numero),
            "plantaOrigem": nome_documento(contrato_com_planta[0]) if contrato_com_planta else "",
            "projetos": projetos,
        }
    temporario = DESTINO.with_suffix(".json.tmp")
    temporario.write_text(json.dumps(unidades, ensure_ascii=False, indent=2), encoding="utf-8")
    temporario.replace(DESTINO)
    print(f"{len(unidades)} unidades gravadas em {DESTINO.name}")


if __name__ == "__main__":
    gerar()
