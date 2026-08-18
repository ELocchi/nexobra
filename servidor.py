#!/usr/bin/env python3
import base64
import binascii
import gzip
import hashlib
import hmac
import json
import os
import re
import secrets
import sqlite3
import subprocess
import threading
import unicodedata
import webbrowser
from datetime import date, datetime, timedelta
from io import BytesIO
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse, unquote
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile

from reportlab.graphics import renderPDF, renderSVG
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image, KeepInFrame, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader

PASTA = Path(__file__).resolve().parent
ARQUIVO_PLANEJAMENTO = PASTA / "dados_planilha_atual.json"
ARQUIVO_ACABAMENTOS = PASTA / "acabamentos_unidades.json"
ARQUIVO_PLANTAS = PASTA / "plantas_unidades.json"
ARQUIVO_PERSONALIZACOES = PASTA / "personalizacoes_unidades.json"
ARQUIVO_LISTA_MESTRA_FVS = PASTA / "Lista_Mestra_FVS_completa.xlsx"
BANCO = Path(os.environ.get("OBRA_BANCO", str(PASTA / "acompanhamento.db")))
CACHE_REGISTROS_COMPACTOS = {"assinatura": None, "corpo": b"", "gzip": b"", "etag": ""}
LOCK_CACHE_REGISTROS = threading.Lock()
HOST = os.environ.get("HOST", "0.0.0.0")
PORTA = int(os.environ.get("PORT", "8000"))
TORRES_NOMES = {"aurora": "Torre Home", "horizonte": "Torre Smart"}
MARCADOR_SEM_SETORES = "__SEM_SETORES__"
_OPCOES_PLANTAS_PDF = {}


def opcoes_plantas_pdf(torre, andar, unidade):
    chave_cache = (torre, int(andar), unidade)
    if chave_cache in _OPCOES_PLANTAS_PDF:
        return _OPCOES_PLANTAS_PDF[chave_cache]
    pdf = PASTA / ("T1-R00.pdf" if torre == "horizonte" else "T2-R00.pdf")
    numero = re.search(r"(\d+)", unidade or "")
    final = int(numero.group(1)[-2:]) if numero else 0
    opcoes = []
    for indice_pagina, pagina in enumerate(PdfReader(pdf).pages):
        texto_original = " ".join((pagina.extract_text() or "").split())
        texto = unicodedata.normalize("NFD", texto_original).encode("ascii", "ignore").decode("ascii").upper()
        final_encontrado = re.search(r"APTO FINAL\s*0?(\d+)", texto)
        if not final_encontrado or int(final_encontrado.group(1)) != final:
            continue
        trecho = texto[final_encontrado.end():final_encontrado.end() + 100]
        intervalo = re.search(r"(\d+)\s*AO\s*(\d+)", trecho)
        conjunto = re.search(r"(\d+)\s*E\s*(\d+)", trecho)
        unico = re.search(r"(\d+)\s*PAV", trecho)
        andares = set(range(int(intervalo.group(1)), int(intervalo.group(2)) + 1)) if intervalo else {int(conjunto.group(1)), int(conjunto.group(2))} if conjunto else {int(unico.group(1))} if unico else set()
        if andares and int(andar) not in andares:
            continue
        inicio = texto.find("OPCAO")
        fim = texto.find("APTO FINAL")
        titulo = texto[inicio:fim].strip() if inicio >= 0 and fim > inicio else ""
        titulo = re.sub(r"\s+", " ", titulo).title().replace("Opcao", "Opção")
        for sem_acento, com_acento in (("Dormitorios", "Dormitórios"), ("Dormitorio", "Dormitório"), ("Suites", "Suítes"), ("Suite", "Suíte"), ("Acessivel", "Acessível")):
            titulo = titulo.replace(sem_acento, com_acento)
        if titulo and not any(item["titulo"] == titulo for item in opcoes):
            torre_numero = 1 if torre == "horizonte" else 2
            opcoes.append({"titulo": titulo, "imagem": f"/miniaturas_tipos_planta/{torre}-t{torre_numero}-p{indice_pagina + 1:02d}.png", "pagina": indice_pagina + 1})
    _OPCOES_PLANTAS_PDF[chave_cache] = opcoes
    return opcoes


def gerar_pdf_planta_unidade(torre, unidade, tipo="planta", indice=0, projeto_id=0):
    imagem = ""
    titulo = f"Planta {unidade}"
    if projeto_id:
        with conectar() as conexao:
            linha = conexao.execute(
                "SELECT titulo, imagem FROM projetos_unidade WHERE id=? AND torre=? AND unidade=?",
                (projeto_id, torre, unidade),
            ).fetchone()
        if not linha:
            raise ValueError("Projeto não encontrado")
        titulo, imagem = linha["titulo"], linha["imagem"]
    elif tipo == "personalizacao":
        dados = json.loads(ARQUIVO_PERSONALIZACOES.read_text(encoding="utf-8")) if ARQUIVO_PERSONALIZACOES.exists() else {}
        projetos = dados.get(f"{torre}|{unidade}", {}).get("projetos", [])
        if indice < 0 or indice >= len(projetos):
            raise ValueError("Projeto não encontrado")
        projeto = projetos[indice]
        titulo = projeto.get("titulo") or projeto.get("arquivo") or titulo
        imagem = projeto.get("imagem", "")
    else:
        dados = json.loads(ARQUIVO_PLANTAS.read_text(encoding="utf-8")) if ARQUIVO_PLANTAS.exists() else {}
        planta = dados.get(f"{torre}|{unidade}", {})
        titulo = planta.get("tipo") or titulo
        imagem = planta.get("plantaMiniatura", "")
    if not imagem:
        raise ValueError("Imagem da planta não encontrada")
    if imagem.startswith("data:image/"):
        conteudo = base64.b64decode(imagem.split(",", 1)[-1])
    else:
        caminho_imagem = (PASTA / unquote(urlparse(imagem).path).lstrip("/")).resolve()
        if PASTA.resolve() not in caminho_imagem.parents or not caminho_imagem.is_file():
            raise ValueError("Arquivo da planta não encontrado")
        conteudo = caminho_imagem.read_bytes()
    leitor = ImageReader(BytesIO(conteudo))
    largura_imagem, altura_imagem = leitor.getSize()
    pagina = landscape(A4) if largura_imagem > altura_imagem else A4
    margem = 12 * mm
    escala = min((pagina[0] - 2 * margem) / largura_imagem, (pagina[1] - 2 * margem) / altura_imagem)
    largura, altura = largura_imagem * escala, altura_imagem * escala
    saida = BytesIO()
    pdf = canvas.Canvas(saida, pagesize=pagina, pageCompression=1)
    pdf.setTitle(str(titulo))
    pdf.drawImage(leitor, (pagina[0] - largura) / 2, (pagina[1] - altura) / 2, largura, altura, preserveAspectRatio=True)
    pdf.showPage()
    pdf.save()
    return saida.getvalue()
SERVICOS_MANUAIS_BASE = [
    "Alvenaria",
    "Pedra Natural",
    "Check List",
    "Coifa de Churrasqueira",
    "Contrapiso",
    "Esquadrias",
    "Estrutura",
    "Gesso",
    "Impermeabilização",
    "Infra Ar Condicionado",
    "Instalações Elétricas",
    "Instalações Hidráulicas",
    "Limpeza Final",
    "Pintura - Fachada",
    "Pintura 1° Demão",
    "Pintura 2° Demão (Geral)",
    "Pintura Hall 1° Demão",
    "Porta Pronta",
    "Portas Shafts",
    "Produção de Argamassa",
    "Revestimento",
]
SERVICOS_LEGADOS_OCULTOS = {
    "Check List",
    "Coifa de Churrasqueira",
    "Infra Ar Condicionado",
    "Limpeza Final",
    "Pintura - Fachada",
    "Pintura 1° Demão",
    "Pintura 2° Demão (Geral)",
    "Pintura Hall 1° Demão",
    "Porta Pronta",
    "Portas Shafts",
}
NS_FVS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
PAVIMENTOS_TECNICOS = {101: "Barrilete", 102: "Reservatório", 103: "Cobertura"}
STATUS_NOMES = {
    "nao-iniciado": "Não iniciado",
    "em-andamento": "Em andamento",
    "pendente": "Pendente",
    "concluido": "Concluído",
}
CONFIGURACAO_LOCAL = PASTA / "configuracao.local.json"
configuracao_local = {}
if CONFIGURACAO_LOCAL.exists():
    try:
        configuracao_local = json.loads(CONFIGURACAO_LOCAL.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        configuracao_local = {}
USUARIO_ENGENHEIRO = os.environ.get("OBRA_USUARIO", configuracao_local.get("usuario_engenheiro", ""))
SENHA_ENGENHEIRO = os.environ.get("OBRA_SENHA", configuracao_local.get("senha_engenheiro", ""))
SEGREDO_VISITANTE = os.environ.get(
    "OBRA_SEGREDO_VISITANTE",
    configuracao_local.get("segredo_visitante", SENHA_ENGENHEIRO or "boulevard-dialogo-visitante-v1"),
).encode("utf-8")


def assinatura_visitante(torre, andar, unidade):
    mensagem = f"{torre}|{int(andar)}|{unidade}".encode("utf-8")
    return hmac.new(SEGREDO_VISITANTE, mensagem, hashlib.sha256).hexdigest()[:32]


def acesso_visitante_valido(torre, andar, unidade, assinatura):
    return bool(assinatura) and hmac.compare_digest(
        assinatura_visitante(torre, andar, unidade), str(assinatura)
    )


def nome_exibicao_usuario(usuario):
    texto = str(usuario or "").strip()
    nomes_conhecidos = {"LOCCHI": "Emanuel Locchi"}
    if texto.upper() in nomes_conhecidos:
        return nomes_conhecidos[texto.upper()]
    identificador = texto.split("@", 1)[0]
    partes = identificador.replace("_", ".").replace("-", ".").split(".")
    nome = " ".join(parte.capitalize() for parte in partes if parte)
    return nome or "Usuário autenticado"


NOME_USUARIO_ENGENHEIRO = nome_exibicao_usuario(USUARIO_ENGENHEIRO)
SESSOES_ENGENHEIRO = {}
PERFIS_ACESSO = {
    "administrador": "Administrador",
    "engenheiro-responsavel": "Engenheiro responsável",
    "operacional": "Operacional",
    "consulta": "Consulta",
    "visitante": "Visitante",
}
PERFIS_ADMINISTRACAO = {"administrador"}
ACOES_PERFIL = {
    "visualizar": "Visualizar dados da obra",
    "atualizar_acompanhamento": "Atualizar planilha de acompanhamento",
    "gerenciar_ocorrencias": "Cadastrar e atualizar ocorrências",
    "gerenciar_setores": "Cadastrar, editar e excluir setores",
    "gerenciar_atividades": "Cadastrar, editar e excluir atividades",
    "gerenciar_projetos": "Adicionar e remover projetos das unidades",
    "gerenciar_documentos": "Adicionar e remover documentos e links",
    "gerar_pdf": "Gerar relatórios em PDF",
    "administrar_acessos": "Administrar usuários, cargos e perfis",
}
PERMISSOES_PADRAO = {
    "administrador": list(ACOES_PERFIL),
    "engenheiro-responsavel": [acao for acao in ACOES_PERFIL if acao != "administrar_acessos"],
    "operacional": ["visualizar", "atualizar_acompanhamento", "gerenciar_ocorrencias", "gerar_pdf"],
    "consulta": ["visualizar", "gerar_pdf"],
    "visitante": ["visualizar"],
}


def gerar_hash_senha(senha):
    salt = secrets.token_bytes(16)
    derivada = hashlib.pbkdf2_hmac("sha256", str(senha).encode("utf-8"), salt, 310_000)
    return f"pbkdf2_sha256$310000${salt.hex()}${derivada.hex()}"


def conferir_senha(senha, valor):
    try:
        algoritmo, iteracoes, salt, esperado = str(valor).split("$", 3)
        if algoritmo != "pbkdf2_sha256":
            return False
        derivada = hashlib.pbkdf2_hmac(
            "sha256", str(senha).encode("utf-8"), bytes.fromhex(salt), int(iteracoes)
        )
        return hmac.compare_digest(derivada.hex(), esperado)
    except (ValueError, TypeError):
        return False


def registrar_auditoria(conexao, usuario_id, usuario_nome, acao, detalhes=""):
    conexao.execute(
        "INSERT INTO auditoria_acessos (usuario_id, usuario_nome, acao, detalhes) VALUES (?, ?, ?, ?)",
        (usuario_id, usuario_nome or "Usuário", acao, detalhes),
    )


def permissoes_do_perfil(conexao, perfil):
    linha = conexao.execute("SELECT permissoes FROM perfis_permissoes WHERE perfil=?", (perfil,)).fetchone()
    try:
        return [acao for acao in json.loads(linha["permissoes"] if linha else "[]") if acao in ACOES_PERFIL]
    except (json.JSONDecodeError, TypeError):
        return []


def rotulo_pavimento(andar):
    numero = int(andar)
    return {
        -2: "Fundação",
        -1: "1º Subsolo",
        0: "Térreo",
        **PAVIMENTOS_TECNICOS,
    }.get(numero, f"{numero}º andar")


def converter_filtro_andar(valor):
    if valor == "todos":
        return None
    try:
        return int(valor)
    except (TypeError, ValueError):
        raise ValueError("Pavimento inválido") from None


def atividades_planejadas(torre, andar=None):
    if not ARQUIVO_PLANEJAMENTO.exists():
        return set()
    dados = json.loads(ARQUIVO_PLANEJAMENTO.read_text(encoding="utf-8"))
    andares = dados.get("torres", {}).get(torre, {})
    grupos = [andares.get(str(andar), [])] if andar is not None else andares.values()
    atividades = {servico["atividade"] for grupo in grupos for servico in grupo}
    try:
        with conectar() as conexao:
            existe = conexao.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name='atividades_config'"
            ).fetchone()
            if existe:
                consulta = "SELECT DISTINCT nome FROM atividades_config WHERE torre = ?"
                parametros = [torre]
                if andar is not None:
                    consulta += " AND andar = ?"
                    parametros.append(int(andar))
                return {
                    linha["nome"] for linha in conexao.execute(consulta, parametros).fetchall()
                }
    except sqlite3.Error:
        pass
    return atividades


def somente_registros_planejados(registros, torre):
    cache = {}
    resultado = []
    for registro in registros:
        andar = registro["andar"]
        if andar not in cache:
            cache[andar] = atividades_planejadas(torre, andar)
        if registro["atividade"] in cache[andar]:
            resultado.append(registro)
    return resultado


def conectar():
    BANCO.parent.mkdir(parents=True, exist_ok=True)
    conexao = sqlite3.connect(BANCO, timeout=15.0)
    conexao.row_factory = sqlite3.Row
    conexao.execute("PRAGMA busy_timeout = 15000")
    conexao.execute("PRAGMA synchronous = NORMAL")
    return conexao


def gerar_relatorio_pdf(torre, andar, unidade, incluir_ocorrencias_pendentes=True):
    with conectar() as conexao:
        registros = conexao.execute(
            "SELECT andar, atividade, status, data_conclusao, observacao FROM registros WHERE torre = ? AND andar = ? AND unidade = ? ORDER BY atividade",
            (torre, andar, unidade),
        ).fetchall()
        registros = somente_registros_planejados(registros, torre)
        ocorrencias = conexao.execute(
            "SELECT atividade, status, data_ocorrencia, descricao FROM ocorrencias WHERE torre = ? AND andar = ? AND unidade = ? ORDER BY data_ocorrencia DESC, id DESC",
            (torre, andar, unidade),
        ).fetchall()
        if not incluir_ocorrencias_pendentes:
            ocorrencias = [item for item in ocorrencias if item["status"] != "pendente"]
    memoria = BytesIO()
    documento = SimpleDocTemplate(memoria, pagesize=A4, rightMargin=16*mm, leftMargin=16*mm, topMargin=14*mm, bottomMargin=14*mm)
    estilos = getSampleStyleSheet()
    elementos = []
    logo = PASTA / "logo dialogo.png"
    if logo.exists():
        elementos.extend([Image(str(logo), width=48*mm, height=16*mm), Spacer(1, 4*mm)])
    nome_unidade = unidade.replace("Apto ", "Apartamento ")
    elementos.append(Paragraph(f"<b>Relatório da unidade</b>", estilos["Title"]))
    elementos.append(Paragraph(f"{escape(TORRES_NOMES.get(torre, torre))} · {andar}º andar · {escape(nome_unidade)}", estilos["Heading2"]))
    elementos.append(Spacer(1, 4*mm))
    total = len(registros)
    concluidos = sum(1 for item in registros if item["status"] == "concluido")
    percentual = round(concluidos / total * 100) if total else 0
    elementos.append(Paragraph(f"<b>Andamento geral:</b> {percentual}% concluído ({concluidos} de {total} serviços)", estilos["BodyText"]))
    elementos.append(Spacer(1, 4*mm))
    tabela = [["Atividade", "Status", "Data", "Observação"]]
    for item in registros:
        data = item["data_conclusao"]
        data = "/".join(reversed(data.split("-"))) if data else "—"
        tabela.append([
            Paragraph(escape(item["atividade"]), estilos["BodyText"]),
            STATUS_NOMES.get(item["status"], item["status"]),
            data,
            Paragraph(escape(item["observacao"] or "—"), estilos["BodyText"]),
        ])
    if len(tabela) == 1:
        tabela.append(["Nenhum serviço registrado", "—", "—", "—"])
    quadro = Table(tabela, colWidths=[55*mm, 34*mm, 25*mm, 58*mm], repeatRows=1)
    quadro.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173a5e")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), .4, colors.HexColor("#cbd3da")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f4f6f8")]),
        ("LEFTPADDING", (0,0), (-1,-1), 5),
        ("RIGHTPADDING", (0,0), (-1,-1), 5),
    ]))
    elementos.extend([quadro, Spacer(1, 6*mm), Paragraph("<b>Ocorrências vinculadas</b>", estilos["Heading2"])])
    if ocorrencias:
        for ocorrencia in ocorrencias:
            data = "/".join(reversed((ocorrencia["data_ocorrencia"] or "").split("-")))
            elementos.append(Paragraph(
                f"<b>{escape(ocorrencia['atividade'])}</b> · {escape(STATUS_NOMES.get(ocorrencia['status'], ocorrencia['status']))} · {data}<br/>{escape(ocorrencia['descricao'])}",
                estilos["BodyText"],
            ))
            elementos.append(Spacer(1, 2*mm))
    else:
        elementos.append(Paragraph("Nenhuma ocorrência vinculada.", estilos["BodyText"]))
    documento.build(elementos)
    return memoria.getvalue()


def gerar_historico_ocorrencias_pdf(torre, andar="todos", unidade="todos", status="todos", tipo="todos"):
    with conectar() as conexao:
        consulta = "SELECT id, torre, andar, unidade, atividade, subatividade, especificacao, status, data_ocorrencia, descricao, foto, foto_nome, criado_em FROM ocorrencias"
        parametros = ()
        if torre != "todos":
            consulta += " WHERE torre = ?"
            parametros = (torre,)
        ocorrencias = conexao.execute(consulta + " ORDER BY data_ocorrencia DESC, id DESC", parametros).fetchall()
    andar_numero = converter_filtro_andar(andar)
    if andar_numero is not None:
        ocorrencias = [item for item in ocorrencias if item["andar"] == andar_numero]
    if unidade != "todos":
        ocorrencias = [item for item in ocorrencias if item["unidade"] == unidade]
    if status != "todos":
        ocorrencias = [item for item in ocorrencias if item["status"] == status]
    if tipo != "todos":
        ocorrencias = [item for item in ocorrencias if ("seguranca" if item["atividade"] == "Segurança" else "atividade") == tipo]
    memoria = BytesIO()
    documento = SimpleDocTemplate(
        memoria, pagesize=A4, rightMargin=18*mm, leftMargin=18*mm,
        topMargin=40*mm, bottomMargin=22*mm,
    )
    estilos = getSampleStyleSheet()
    estilo_texto = estilos["BodyText"].clone("OcorrenciaTexto")
    estilo_texto.fontName = "Helvetica"
    estilo_texto.fontSize = 8.5
    estilo_texto.leading = 10.5
    estilo_centro = estilo_texto.clone("OcorrenciaCentro")
    estilo_centro.alignment = 1
    total_paginas = max(1, (len(ocorrencias) + 3) // 4)
    gerado_em = datetime.now()
    identificador = gerado_em.strftime("%Y%m%d%H%M")

    def cabecalho_rodape(canvas, doc):
        largura, altura = A4
        canvas.saveState()
        canvas.setFont("Helvetica", 8.5)
        y = altura - 16*mm
        linhas = [
            ("Criado:", gerado_em.strftime("%d/%m/%Y %H:%M")),
            ("Localização:", TORRES_NOMES.get(torre, "Todas as torres")),
            ("Título:", "Histórico de ocorrências"),
            ("Filtros:", " · ".join([
                rotulo_pavimento(andar_numero) if andar_numero is not None else "Todos os pavimentos",
                unidade if unidade != "todos" else "Todas as unidades",
                STATUS_NOMES.get(status, status) if status != "todos" else "Todos os status",
            ])),
            ("Nº de itens:", str(len(ocorrencias))),
        ]
        for rotulo, valor in linhas:
            canvas.drawString(20*mm, y, rotulo)
            canvas.drawString(42*mm, y, valor)
            y -= 4.2*mm
        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(.7)
        canvas.line(20*mm, altura-38*mm, largura-20*mm, altura-38*mm)
        canvas.line(20*mm, 17*mm, largura-20*mm, 17*mm)
        logo = PASTA / "logo dialogo.png"
        if logo.exists():
            canvas.drawImage(str(logo), 20*mm, 6*mm, width=24*mm, height=8*mm, preserveAspectRatio=True, mask="auto")
        canvas.setFont("Helvetica", 7)
        canvas.drawString(48*mm, 11*mm, "Desenvolvido por Locchi Engenharia")
        canvas.drawString(48*mm, 7.5*mm, "BoulevarDiálogo Butantã")
        canvas.drawRightString(largura-20*mm, 11*mm, f"Doc. Id.: {identificador}")
        canvas.drawRightString(largura-20*mm, 7.5*mm, f"página {doc.page} de {total_paginas}")
        canvas.restoreState()

    def cartao(item, numero):
        conteudo = []
        if item["foto"] and "," in item["foto"]:
            try:
                dados_foto = base64.b64decode(item["foto"].split(",", 1)[1])
                conteudo.append(Image(BytesIO(dados_foto), width=61*mm, height=70*mm, kind="proportional"))
            except Exception:
                conteudo.append(Paragraph("<br/><br/><br/>Foto indisponível", estilo_centro))
        else:
            conteudo.append(Paragraph("<br/><br/><br/>Sem foto", estilo_centro))
        conteudo.append(Paragraph(escape(item["foto_nome"] or f"ocorrencia-{item['id']}.jpg"), estilo_centro))
        data = "/".join(reversed(item["data_ocorrencia"].split("-"))) if item["data_ocorrencia"] else "—"
        servico = " · ".join(filter(None, [
            item["atividade"], item["subatividade"], item["especificacao"],
        ]))
        detalhes = (
            f"<b>Criada:</b>&nbsp;&nbsp; {data}<br/>"
            f"<b>({numero})</b>&nbsp;&nbsp; {escape(TORRES_NOMES.get(item['torre'], item['torre']))} · {escape(item['unidade'])} · {escape(rotulo_pavimento(item['andar']))}<br/>"
            f"<b>Serviço:</b>&nbsp;&nbsp; {escape(servico)}<br/>"
            f"<b>Status:</b>&nbsp;&nbsp; {escape(STATUS_NOMES.get(item['status'], item['status']))}<br/>"
            f"{escape(item['descricao'])}"
        )
        conteudo.extend([Spacer(1, 2*mm), Paragraph(detalhes, estilo_texto)])
        return KeepInFrame(78*mm, 105*mm, conteudo, mode="shrink")

    elementos = []
    if not ocorrencias:
        elementos.append(Paragraph("Nenhuma ocorrência encontrada para os filtros selecionados.", estilo_texto))
    for inicio in range(0, len(ocorrencias), 4):
        lote = ocorrencias[inicio:inicio+4]
        celulas = [cartao(item, inicio+indice+1) for indice, item in enumerate(lote)]
        while len(celulas) < 4:
            celulas.append("")
        grade = Table(
            [[celulas[0], celulas[1]], [celulas[2], celulas[3]]],
            colWidths=[84*mm, 84*mm], rowHeights=[107*mm, 107*mm],
            hAlign="CENTER",
        )
        grade.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 4*mm),
            ("RIGHTPADDING", (0,0), (-1,-1), 4*mm),
            ("TOPPADDING", (0,0), (-1,-1), 1*mm),
            ("BOTTOMPADDING", (0,0), (-1,-1), 1*mm),
        ]))
        elementos.append(grade)
        if inicio + 4 < len(ocorrencias):
            elementos.append(PageBreak())
    documento.build(elementos, onFirstPage=cabecalho_rodape, onLaterPages=cabecalho_rodape)
    return memoria.getvalue()


def gerar_pagina_visitante(torre, andar, unidade, assinatura=""):
    with conectar() as conexao:
        registros = conexao.execute(
            "SELECT atividade, status, data_conclusao, observacao, foto, foto_nome FROM registros WHERE torre = ? AND andar = ? AND unidade = ? ORDER BY atividade",
            (torre, andar, unidade),
        ).fetchall()
        registros = somente_registros_planejados(registros, torre)
        ocorrencias = conexao.execute(
            "SELECT atividade, status, data_ocorrencia, descricao, foto, foto_nome FROM ocorrencias WHERE torre = ? AND andar = ? AND unidade = ? AND status != 'pendente' ORDER BY data_ocorrencia DESC, id DESC",
            (torre, andar, unidade),
        ).fetchall()
    total = len(registros)
    concluidos = sum(1 for item in registros if item["status"] == "concluido")
    percentual = round(concluidos / total * 100) if total else 0
    consulta = urlencode({"torre": torre, "andar": andar, "unidade": unidade, "acesso": assinatura})
    linhas = []
    for item in registros:
        data = "/".join(reversed(item["data_conclusao"].split("-"))) if item["data_conclusao"] else "—"
        foto = (
            f'<a class="foto" href="{item["foto"]}" download="{escape(item["foto_nome"] or "foto-servico.jpg")}">Ver/baixar foto</a>'
            if item["foto"] else "—"
        )
        linhas.append(
            f'<tr><td>{escape(item["atividade"])}</td><td><span class="status {escape(item["status"])}">{escape(STATUS_NOMES.get(item["status"], item["status"]))}</span></td>'
            f'<td>{data}</td><td>{escape(item["observacao"] or "—")}</td><td>{foto}</td></tr>'
        )
    cards = []
    for item in ocorrencias:
        data = "/".join(reversed(item["data_ocorrencia"].split("-"))) if item["data_ocorrencia"] else "—"
        imagem = f'<img src="{item["foto"]}" alt="Foto da ocorrência">' if item["foto"] else ""
        cards.append(
            f'<article class="ocorrencia {escape(item["status"])}"><strong>{escape(item["atividade"])}</strong>'
            f'<span class="status {escape(item["status"])}">{escape(STATUS_NOMES.get(item["status"], item["status"]))}</span>'
            f'<p>{escape(item["descricao"])}</p><small>{data}</small>{imagem}</article>'
        )
    nome_unidade = unidade.replace("Apto ", "Apartamento ")
    return f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{escape(nome_unidade)} — acompanhamento</title><style>
*{{box-sizing:border-box}}body{{margin:0;background:#f2f5f7;color:#243445;font-family:Arial,sans-serif}}header{{background:#173a5e;color:#fff;padding:18px}}header div,main{{max-width:1050px;margin:auto}}header img{{width:170px;background:#fff;border-radius:7px;padding:6px}}h1{{font-size:1.35rem;margin:15px 0 4px}}header p{{margin:0;color:#d9e3eb}}main{{padding:18px}}.avanco,.painel{{background:#fff;border:1px solid #d8dfe5;border-radius:11px;padding:16px;margin-bottom:15px}}.avanco strong{{font-size:1.8rem;color:#173a5e}}.barra{{height:10px;background:#e5e9ec;border-radius:8px;overflow:hidden;margin-top:8px}}.barra i{{display:block;height:100%;width:{percentual}%;background:#2e9e5b}}h2{{font-size:1rem;color:#173a5e}}table{{width:100%;border-collapse:collapse;font-size:.8rem}}th,td{{padding:9px;border-bottom:1px solid #e1e6ea;text-align:left;vertical-align:top}}th{{background:#eef3f6}}.status{{display:inline-block;border-radius:12px;padding:4px 7px;font-size:.68rem;font-weight:bold}}.status.nao-iniciado{{background:#e8ecef;color:#637582}}.status.em-andamento{{background:#fff1bd;color:#8a6a00}}.status.pendente{{background:#fbe0dd;color:#d9483d}}.status.concluido{{background:#def2e5;color:#2e9e5b}}.ocorrencias{{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:10px}}.ocorrencia{{background:#fff;border-left:4px solid #637582;border-radius:8px;padding:12px}}.ocorrencia.em-andamento{{border-color:#e0a800}}.ocorrencia.pendente{{border-color:#d9483d}}.ocorrencia.concluido{{border-color:#2e9e5b}}.ocorrencia strong{{display:block;margin:0 0 7px}}.ocorrencia p{{font-size:.8rem;line-height:1.45}}.ocorrencia img{{display:block;width:100%;max-height:260px;object-fit:cover;border-radius:7px;margin-top:9px}}.botao,.foto{{display:inline-block;color:#fff;background:#17608f;border-radius:7px;padding:9px 12px;text-decoration:none;font-size:.76rem;font-weight:bold}}.foto{{padding:5px 7px}}.aviso{{font-size:.72rem;color:#687887;margin-top:10px}}@media(max-width:700px){{main{{padding:10px}}.tabela{{overflow:auto}}table{{min-width:720px}}}}
</style></head><body><header><div><img src="/logo%20dialogo.png" alt="Diálogo Engenharia"><h1>{escape(nome_unidade)} · {andar}º andar</h1><p>{escape(TORRES_NOMES.get(torre, torre))} · visualização para visitantes</p></div></header>
<main><section class="avanco"><strong>{percentual}% concluído</strong><div class="barra"><i></i></div><p>{concluidos} de {total} serviços concluídos</p><a class="botao" href="/relatorio.pdf?{consulta}">Abrir relatório em PDF</a><div class="aviso">Página somente para consulta. Nenhuma informação pode ser alterada neste acesso.</div></section>
<section class="painel"><h2>Serviços da unidade</h2><div class="tabela"><table><thead><tr><th>Atividade</th><th>Status</th><th>Data</th><th>Observação</th><th>Foto</th></tr></thead><tbody>{''.join(linhas) if linhas else '<tr><td colspan="5">Nenhum serviço registrado.</td></tr>'}</tbody></table></div></section>
<section class="painel"><h2>Ocorrências</h2><div class="ocorrencias">{''.join(cards) if cards else '<p>Nenhuma ocorrência vinculada.</p>'}</div></section></main></body></html>""".encode("utf-8")


def registros_filtrados(torre, andar="todos", unidade="todos", atividade="todos", status="todos"):
    with conectar() as conexao:
        itens = conexao.execute(
            "SELECT andar, unidade, atividade, status, data_conclusao FROM registros WHERE torre = ? ORDER BY andar, unidade, atividade",
            (torre,),
        ).fetchall()
    itens = somente_registros_planejados(itens, torre)
    if andar != "todos":
        itens = [item for item in itens if item["andar"] == int(andar)]
    if unidade != "todos":
        itens = [item for item in itens if item["unidade"] == unidade]
    if atividade != "todos":
        itens = [item for item in itens if item["atividade"] == atividade]
    if status != "todos":
        itens = [item for item in itens if item["status"] == status]
    return itens


def gerar_pdf_visitante_filtros(torre, andar="todos", unidade="todos", atividade="todos", status="todos", acabamento="todos", ocorrencias="todas"):
    itens = registros_filtrados(torre, andar, unidade, atividade, status)
    if acabamento != "todos" and ARQUIVO_ACABAMENTOS.exists():
        dados_acabamentos = json.loads(ARQUIVO_ACABAMENTOS.read_text(encoding="utf-8"))
        def corresponde(item):
            escolhas = [str(escolha.get("opcao", "")) for escolha in item.get("itens", [])]
            alteradas = [escolha for escolha in escolhas if escolha.lower() != "opção 1"]
            if acabamento == "padrao":
                return bool(escolhas) and not alteradas
            if acabamento == "alterado":
                return bool(alteradas)
            if acabamento == "personalizada":
                return any("personalizada" in escolha.lower() for escolha in escolhas)
            if acabamento == "nao-instalar":
                return any("não instalar" in escolha.lower() or "nao instalar" in escolha.lower() for escolha in escolhas)
            numero = acabamento.replace("opcao-", "")
            return any(escolha.lower() == f"opção {numero}" for escolha in escolhas)
        unidades_acabamento = {
            chave.split("|", 1)[1] for chave, dados in dados_acabamentos.items()
            if chave.startswith(f"{torre}|") and corresponde(dados)
        }
        itens = [item for item in itens if item["unidade"] in unidades_acabamento]
    with conectar() as conexao:
        ocorrencias_banco = conexao.execute(
            "SELECT andar, unidade, atividade, descricao, status, data_ocorrencia FROM ocorrencias WHERE torre = ? ORDER BY data_ocorrencia DESC, id DESC",
            (torre,),
        ).fetchall()
    ocorrencias_filtradas = [item for item in ocorrencias_banco
        if (andar == "todos" or item["andar"] == int(andar))
        and (unidade == "todos" or item["unidade"] == unidade)
        and (atividade == "todos" or item["atividade"] == atividade)]
    if ocorrencias in {"pendente", "concluido"}:
        ocorrencias_filtradas = [item for item in ocorrencias_filtradas if item["status"] == ocorrencias]
    unidades_com_ocorrencia = {(item["andar"], item["unidade"]) for item in ocorrencias_filtradas}
    if ocorrencias in {"com", "pendente", "concluido"}:
        itens = [item for item in itens if (item["andar"], item["unidade"]) in unidades_com_ocorrencia]
    elif ocorrencias == "sem":
        todas_ocorrencias = {(item["andar"], item["unidade"]) for item in ocorrencias_banco}
        itens = [item for item in itens if (item["andar"], item["unidade"]) not in todas_ocorrencias]
    memoria = BytesIO()
    documento = SimpleDocTemplate(memoria, pagesize=landscape(A4), rightMargin=12*mm, leftMargin=12*mm, topMargin=12*mm, bottomMargin=12*mm)
    estilos = getSampleStyleSheet()
    elementos = []
    logo = PASTA / "logo dialogo.png"
    if logo.exists():
        elementos.extend([Image(str(logo), width=42*mm, height=14*mm), Spacer(1, 3*mm)])
    elementos.append(Paragraph("<b>Relatório de acompanhamento da obra</b>", estilos["Title"]))
    filtros = [
        TORRES_NOMES.get(torre, torre),
        f"Andar: {andar if andar != 'todos' else 'Todos'}",
        f"Unidade: {unidade.replace('Apto ', 'Apartamento ') if unidade != 'todos' else 'Todas'}",
        f"Serviço: {atividade if atividade != 'todos' else 'Todos'}",
        f"Situação: {STATUS_NOMES.get(status, 'Todas') if status != 'todos' else 'Todas'}",
        f"Acabamento: {acabamento.replace('-', ' ').title() if acabamento != 'todos' else 'Todos'}",
        f"Ocorrências: {ocorrencias.title()}",
    ]
    elementos.extend([Paragraph(" · ".join(escape(item) for item in filtros), estilos["BodyText"]), Spacer(1, 4*mm)])
    tabela = [["Andar", "Unidade", "Serviço", "Status", "Data da atualização"]]
    for item in itens:
        data = "/".join(reversed(item["data_conclusao"].split("-"))) if item["data_conclusao"] else "—"
        tabela.append([f"{item['andar']}º", escape(item["unidade"]), Paragraph(escape(item["atividade"]), estilos["BodyText"]), STATUS_NOMES.get(item["status"], item["status"]), data])
    if len(tabela) == 1:
        tabela.append(["—", "—", "Nenhum registro encontrado", "—", "—"])
    quadro = Table(tabela, colWidths=[24*mm, 48*mm, 90*mm, 50*mm, 48*mm], repeatRows=1)
    quadro.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173a5e")), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), .4, colors.HexColor("#cbd3da")), ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("FONTSIZE", (0,0), (-1,-1), 8), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f4f6f8")]),
    ]))
    elementos.extend([Paragraph(f"<b>Total de serviços:</b> {len(itens)}", estilos["BodyText"]), Spacer(1, 3*mm), quadro])
    if ocorrencias != "sem":
        elementos.extend([Spacer(1, 6*mm), Paragraph("<b>Ocorrências</b>", estilos["Heading2"])])
        tabela_ocorrencias = [["Andar", "Unidade", "Atividade", "Ocorrência", "Status", "Data"]]
        for item in ocorrencias_filtradas:
            tabela_ocorrencias.append([f"{item['andar']}º", item["unidade"], item["atividade"], Paragraph(escape(item["descricao"]), estilos["BodyText"]), STATUS_NOMES.get(item["status"], item["status"]), "/".join(reversed(item["data_ocorrencia"].split("-")))])
        if len(tabela_ocorrencias) == 1:
            tabela_ocorrencias.append(["—", "—", "—", "Nenhuma ocorrência encontrada", "—", "—"])
        quadro_ocorrencias = Table(tabela_ocorrencias, colWidths=[18*mm, 35*mm, 55*mm, 92*mm, 34*mm, 28*mm], repeatRows=1)
        quadro_ocorrencias.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173a5e")), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("GRID", (0,0), (-1,-1), .4, colors.HexColor("#cbd3da")), ("VALIGN", (0,0), (-1,-1), "TOP"), ("FONTSIZE", (0,0), (-1,-1), 7), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f4f6f8")])]))
        elementos.append(quadro_ocorrencias)
    documento.build(elementos)
    return memoria.getvalue()


def gerar_portal_relatorios(torre, andar="todos", unidade="todos", atividade="todos"):
    itens = registros_filtrados(torre, andar, unidade, atividade)
    with conectar() as conexao:
        unidades = conexao.execute(
            "SELECT DISTINCT unidade FROM registros WHERE torre = ? AND (? = 'todos' OR andar = ?) ORDER BY unidade",
            (torre, andar, int(andar) if andar != "todos" else 0),
        ).fetchall()
        atividades = conexao.execute(
            "SELECT DISTINCT atividade FROM registros WHERE torre = ? ORDER BY atividade",
            (torre,),
        ).fetchall()
    max_andares = 36 if torre == "aurora" else 23
    opcoes_andar = '<option value="todos">Todos os andares</option>' + "".join(
        f'<option value="{n}" {"selected" if str(n)==andar else ""}>{n}º andar</option>' for n in range(1, max_andares + 1)
    )
    opcoes_unidade = '<option value="todos">Todos os apartamentos/setores</option>' + "".join(
        f'<option value="{escape(item["unidade"])}" {"selected" if item["unidade"]==unidade else ""}>{escape(item["unidade"].replace("Apto ", "Apartamento "))}</option>' for item in unidades
    )
    opcoes_atividade = '<option value="todos">Todos os serviços</option>' + "".join(
        f'<option value="{escape(item["atividade"])}" {"selected" if item["atividade"]==atividade else ""}>{escape(item["atividade"])}</option>' for item in atividades
    )
    contagens = {status: sum(1 for item in itens if item["status"] == status) for status in STATUS_NOMES}
    total = len(itens)
    realizados = contagens["concluido"]
    servicos_pendentes = total - realizados
    avanco = round(realizados / total * 100) if total else 0
    andares_no_escopo = sorted({item["andar"] for item in itens})
    pavimentos_concluidos = sum(
        1 for numero in andares_no_escopo
        if all(item["status"] == "concluido" for item in itens if item["andar"] == numero)
    )
    percentuais = {
        status: (contagens[status] / total * 100 if total else 0)
        for status in STATUS_NOMES
    }
    limite_cinza = percentuais["nao-iniciado"]
    limite_amarelo = limite_cinza + percentuais["em-andamento"]
    limite_vermelho = limite_amarelo + percentuais["pendente"]
    gradiente_status = (
        f"conic-gradient(#637582 0 {limite_cinza:.2f}%,#e0a800 {limite_cinza:.2f}% {limite_amarelo:.2f}%,"
        f"#d9483d {limite_amarelo:.2f}% {limite_vermelho:.2f}%,#2e9e5b {limite_vermelho:.2f}% 100%)"
        if total else "#637582"
    )
    pavimentos_html = []
    for numero in range(max_andares, 0, -1):
        grupo = [item for item in itens if item["andar"] == numero]
        if grupo:
            totais = {status: sum(1 for item in grupo if item["status"] == status) / len(grupo) * 100 for status in STATUS_NOMES}
            p1 = totais["nao-iniciado"]
            p2 = p1 + totais["em-andamento"]
            p3 = p2 + totais["pendente"]
            fundo = f"linear-gradient(to right,#637582 0 {p1:.2f}%,#e0a800 {p1:.2f}% {p2:.2f}%,#d9483d {p2:.2f}% {p3:.2f}%,#2e9e5b {p3:.2f}% 100%)"
        else:
            fundo = "#455f72"
        pavimentos_html.append(f'<div class="andar-vis"><b>{numero}º</b><i style="background:{fundo}"></i></div>')
    linhas = "".join(
        f'<tr><td>{item["andar"]}º</td><td>{escape(item["unidade"])}</td><td>{escape(item["atividade"])}</td><td><span class="status {escape(item["status"])}">{escape(STATUS_NOMES.get(item["status"], item["status"]))}</span></td><td>{"/".join(reversed(item["data_conclusao"].split("-"))) if item["data_conclusao"] else "—"}</td></tr>'
        for item in itens
    )
    consulta_pdf = urlencode({"torre": torre, "andar": andar, "unidade": unidade, "atividade": atividade})
    return f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Relatórios — visão visitante</title><style>
*{{box-sizing:border-box}}body{{margin:0;background:#f2f5f7;color:#243445;font-family:Arial,sans-serif}}header{{background:#173a5e;color:#fff;padding:18px}}header div,main{{max-width:1150px;margin:auto}}header img{{width:170px;background:#fff;border-radius:7px;padding:6px}}h1{{font-size:1.3rem;margin:13px 0 3px}}main{{padding:18px}}.painel{{background:#fff;border:1px solid #d8dfe5;border-radius:11px;padding:16px;margin-bottom:15px}}form{{display:grid;grid-template-columns:repeat(3,1fr) auto;gap:10px;align-items:end}}label{{display:block;font-size:.68rem;font-weight:bold;margin-bottom:5px}}select{{width:100%;height:38px;border:1px solid #bdc8d0;border-radius:7px;background:#fff;padding:0 8px}}button,.pdf{{height:38px;border:0;border-radius:7px;background:#17608f;color:#fff;padding:0 14px;font-weight:bold;text-decoration:none;display:inline-flex;align-items:center;justify-content:center;cursor:pointer}}.dashboard{{background:#123554;border-radius:13px;padding:16px;margin-bottom:15px;color:#fff}}.kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px}}.kpi{{padding:14px;border:1px solid #49647d;border-radius:9px;background:#294965}}.kpi span{{display:block;color:#bdcbd6;font-size:.68rem;text-transform:uppercase}}.kpi strong{{display:block;font-size:1.45rem;margin-top:4px}}.grafico-status{{background:#fff;color:#243445;border-radius:10px;padding:14px;text-align:center;margin-bottom:14px}}.rosca{{width:190px;height:190px;border-radius:50%;margin:10px auto;display:grid;place-items:center}}.rosca:after{{content:"";width:112px;height:112px;border-radius:50%;background:#fff}}.legenda{{display:flex;justify-content:center;flex-wrap:wrap;gap:12px;font-size:.68rem}}.legenda i{{display:inline-block;width:10px;height:10px;margin-right:4px}}.evolucao{{background:#294965;border:1px solid #49647d;border-radius:10px;padding:14px}}.torre-vis{{width:min(100%,360px);margin:12px auto}}.andar-vis{{display:grid;grid-template-columns:30px 1fr;gap:5px;align-items:center;height:12px}}.andar-vis b{{font-size:.52rem;text-align:right;color:#fff}}.andar-vis i{{display:block;height:8px;border:1px solid rgba(255,255,255,.18)}}.nome-torre{{text-align:center;font-weight:bold;margin-top:10px}}table{{width:100%;border-collapse:collapse;font-size:.8rem}}th,td{{padding:9px;border-bottom:1px solid #e1e6ea;text-align:left}}th{{background:#eef3f6}}.tabela{{overflow:auto}}.status{{display:inline-block;border-radius:12px;padding:4px 7px;font-size:.68rem;font-weight:bold}}.nao-iniciado{{background:#e8ecef;color:#637582}}.em-andamento{{background:#fff1bd;color:#8a6a00}}.pendente{{background:#fbe0dd;color:#d9483d}}.concluido{{background:#def2e5;color:#2e9e5b}}.aviso{{font-size:.72rem;color:#687887;margin-top:8px}}@media(max-width:750px){{form{{grid-template-columns:1fr}}.kpis{{grid-template-columns:1fr 1fr}}main{{padding:10px}}table{{min-width:720px}}}}
</style></head><body><header><div><img src="/logo%20dialogo.png" alt="Diálogo Engenharia"><h1>Relatórios da obra — visão visitante</h1><p>{escape(TORRES_NOMES.get(torre, torre))}</p></div></header><main>
<section class="painel"><form method="get" action="/visitante-relatorios"><input type="hidden" name="torre" value="{escape(torre)}"><div><label>Pavimento</label><select name="andar">{opcoes_andar}</select></div><div><label>Apartamento / setor</label><select name="unidade">{opcoes_unidade}</select></div><div><label>Serviço</label><select name="atividade">{opcoes_atividade}</select></div><button type="submit">Aplicar filtros</button></form><div class="aviso">Acesso somente para consulta. Nenhuma informação pode ser alterada.</div></section>
<section class="dashboard"><div class="kpis"><div class="kpi"><span>Avanço geral</span><strong>{avanco}%</strong></div><div class="kpi"><span>Serviços realizados</span><strong>{realizados}</strong></div><div class="kpi"><span>Serviços pendentes</span><strong>{servicos_pendentes}</strong></div><div class="kpi"><span>Pavimentos concluídos</span><strong>{pavimentos_concluidos} / {len(andares_no_escopo)}</strong></div></div>
<div class="grafico-status"><h2>Status geral</h2><div class="rosca" style="background:{gradiente_status}"></div><div class="legenda"><span><i style="background:#637582"></i>Não iniciado</span><span><i style="background:#e0a800"></i>Em andamento</span><span><i style="background:#d9483d"></i>Pendente</span><span><i style="background:#2e9e5b"></i>Concluído</span></div></div>
<div class="evolucao"><h2>Evolução visual dos pavimentos</h2><div class="legenda"><span><i style="background:#637582"></i>Não iniciado</span><span><i style="background:#e0a800"></i>Em andamento</span><span><i style="background:#d9483d"></i>Pendente</span><span><i style="background:#2e9e5b"></i>Concluído</span></div><div class="torre-vis">{''.join(pavimentos_html)}</div><div class="nome-torre">{escape(TORRES_NOMES.get(torre, torre))} · {max_andares} pavimentos</div></div></section>
<section class="painel"><a class="pdf" href="/relatorio-visitante.pdf?{consulta_pdf}">Gerar PDF com estes filtros</a></section>
<section class="painel"><h2>Serviços encontrados: {len(itens)}</h2><div class="tabela"><table><thead><tr><th>Andar</th><th>Unidade</th><th>Serviço</th><th>Status</th><th>Data</th></tr></thead><tbody>{linhas if linhas else '<tr><td colspan="5">Nenhum serviço encontrado.</td></tr>'}</tbody></table></div></section>
</main></body></html>""".encode("utf-8")


def gerar_qr_svg(url):
    codigo = qr.QrCodeWidget(url)
    x1, y1, x2, y2 = codigo.getBounds()
    largura, altura = x2 - x1, y2 - y1
    tamanho = 280
    desenho = Drawing(tamanho, tamanho, transform=[tamanho/largura, 0, 0, tamanho/altura, 0, 0])
    desenho.add(codigo)
    return renderSVG.drawToString(desenho)


def desenho_qr(url, tamanho):
    codigo = qr.QrCodeWidget(url)
    x1, y1, x2, y2 = codigo.getBounds()
    largura, altura = x2 - x1, y2 - y1
    desenho = Drawing(tamanho, tamanho, transform=[tamanho/largura, 0, 0, tamanho/altura, -x1, -y1])
    desenho.add(codigo)
    return desenho


def desenhar_imagem_contida(pdf, arquivo, x, y, largura, altura):
    if not arquivo.exists():
        return
    leitor = ImageReader(str(arquivo))
    largura_original, altura_original = leitor.getSize()
    escala = min(largura / largura_original, altura / altura_original)
    final_largura, final_altura = largura_original * escala, altura_original * escala
    pdf.drawImage(
        leitor, x + (largura - final_largura) / 2, y + (altura - final_altura) / 2,
        final_largura, final_altura, preserveAspectRatio=True, mask="auto",
    )


def unidades_para_placas(torre, andar="todos", unidade="todos"):
    if unidade != "todos":
        numero = re.search(r"(\d+)", unidade)
        andar_unidade = int(andar) if andar != "todos" else int(numero.group(1)[:-2]) if numero and len(numero.group(1)) > 2 else 0
        return [(andar_unidade, unidade)]
    ultimo_andar = 36 if torre == "aurora" else 23
    andares = [int(andar)] if andar != "todos" else range(1, ultimo_andar + 1)
    return [
        (numero, nome) for numero in andares for nome in apartamentos_do_andar(torre, numero)
        if re.match(r"^(Apto|Apartamento)\s", nome, re.I)
    ]


def gerar_placas_pdf(torre, unidades, base_publica):
    memoria = BytesIO()
    tamanho_placa = (105*mm, 145*mm)
    pdf = canvas.Canvas(memoria, pagesize=tamanho_placa, pageCompression=1)
    largura, altura = tamanho_placa
    azul = colors.HexColor("#294d82")
    vermelho = colors.HexColor("#c82512")
    margem = 5 * mm
    for andar, unidade in unidades:
        pdf.setFillColor(colors.white)
        pdf.rect(0, 0, largura, altura, fill=1, stroke=0)
        pdf.setStrokeColor(azul)
        pdf.setLineWidth(1.1 * mm)
        pdf.rect(margem, margem, largura - 2*margem, altura - 2*margem, fill=0, stroke=1)
        pdf.setFillColor(azul)
        pdf.rect(margem, altura - 33*mm, largura - 2*margem, 28*mm, fill=1, stroke=0)
        pdf.rect(margem, margem, largura - 2*margem, 11*mm, fill=1, stroke=0)

        # A marca da obra fica em uma faixa branca para preservar as cores originais.
        pdf.setFillColor(colors.white)
        pdf.roundRect(10*mm, altura - 30*mm, largura - 20*mm, 22*mm, 1.5*mm, fill=1, stroke=0)
        desenhar_imagem_contida(pdf, PASTA / "boulevardialogo.png", 12*mm, altura - 28*mm, largura - 24*mm, 18*mm)

        pdf.setFillColor(vermelho)
        pdf.setFont("Helvetica-Bold", 17)
        pdf.drawCentredString(largura/2, altura - 43*mm, TORRES_NOMES.get(torre, torre).upper())
        nome_unidade = unidade.replace("Apto ", "").replace("Apartamento ", "")
        pdf.setFillColor(colors.black)
        tamanho_fonte = 48 if len(nome_unidade) <= 5 else 38
        pdf.setFont("Helvetica-Bold", tamanho_fonte)
        pdf.drawCentredString(largura/2, altura - 70*mm, nome_unidade)

        assinatura = assinatura_visitante(torre, andar, unidade)
        consulta = urlencode({"torre": torre, "andar": andar, "unidade": unidade, "acesso": assinatura})
        url = f"{base_publica}/visitante?{consulta}"
        tamanho_qr = 34*mm
        renderPDF.draw(desenho_qr(url, tamanho_qr), pdf, (largura-tamanho_qr)/2, 29*mm)
        desenhar_imagem_contida(pdf, PASTA / "logo dialogo.png", 7*mm, 17*mm, 29*mm, 11*mm)
        pdf.showPage()
    pdf.save()
    return memoria.getvalue()


def gerar_pasta_placas_zip(torre, unidades, base_publica, escopo):
    memoria = BytesIO()
    nome_escopo = re.sub(
        r"[^a-z0-9]+", "-",
        unicodedata.normalize("NFD", escopo).encode("ascii", "ignore").decode().lower(),
    ).strip("-")
    pasta_raiz = f"placas-{nome_escopo}"
    with ZipFile(memoria, "w") as arquivo_zip:
        for andar, unidade in unidades:
            numero = re.sub(r"[^0-9a-z]+", "-", unidade.lower()).strip("-").replace("apto-", "")
            caminho = f"{pasta_raiz}/placa-{numero}.pdf"
            arquivo_zip.writestr(caminho, gerar_placas_pdf(torre, [(andar, unidade)], base_publica))
    return memoria.getvalue(), f"{pasta_raiz}.zip"


def endereco_rede():
    try:
        return subprocess.check_output(["ipconfig", "getifaddr", "en0"], text=True).strip()
    except Exception:
        return ""


def gerar_inicio_visitante():
    return """<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Acesso visitante</title><style>
*{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;background:#123554;font-family:Arial,sans-serif;color:#243445}.caixa{width:min(92%,720px);background:#fff;border-radius:15px;padding:28px;box-shadow:0 18px 50px #0005;text-align:center}.logo{width:190px;margin-bottom:18px}h1{color:#173a5e}.torres{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:22px}.torre{display:block;padding:25px 15px;border:1px solid #ccd6de;border-radius:11px;text-decoration:none;color:#173a5e;font-weight:bold;background:#f4f7f9}.torre:hover{background:#e8f1f6;border-color:#17608f}.torre span{display:block;font-size:.75rem;color:#687887;margin-top:7px;font-weight:normal}.aviso{font-size:.75rem;color:#687887;margin-top:20px}@media(max-width:550px){.torres{grid-template-columns:1fr}}
</style></head><body><main class="caixa"><img class="logo" src="/logo%20dialogo.png" alt="Diálogo Engenharia"><h1>Visão visitante</h1><p>Selecione a torre que deseja consultar.</p><div class="torres"><a class="torre" href="/visitante-relatorios?torre=aurora">Torre Home<span>36 pavimentos</span></a><a class="torre" href="/visitante-relatorios?torre=horizonte">Torre Smart<span>23 pavimentos</span></a></div><div class="aviso">Acesso somente para leitura. Não é possível alterar informações.</div></main></body></html>""".encode("utf-8")


def preparar_banco():
    with conectar() as conexao:
        conexao.execute("PRAGMA journal_mode = WAL")
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS registros (
                chave TEXT PRIMARY KEY,
                torre TEXT NOT NULL,
                andar INTEGER NOT NULL,
                unidade TEXT NOT NULL,
                atividade TEXT NOT NULL,
                concluido INTEGER NOT NULL DEFAULT 0,
                data_conclusao TEXT NOT NULL DEFAULT '',
                observacao TEXT NOT NULL DEFAULT '',
                foto TEXT NOT NULL DEFAULT '',
                foto_nome TEXT NOT NULL DEFAULT '',
                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        colunas_registros = {
            linha["name"]
            for linha in conexao.execute("PRAGMA table_info(registros)").fetchall()
        }
        if "status" not in colunas_registros:
            conexao.execute(
                "ALTER TABLE registros ADD COLUMN status TEXT NOT NULL DEFAULT 'nao-iniciado'"
            )
            conexao.execute(
                """
                UPDATE registros
                SET status = CASE WHEN concluido = 1 THEN 'concluido' ELSE 'nao-iniciado' END
                """
            )
        if "especificacoes" not in colunas_registros:
            conexao.execute(
                "ALTER TABLE registros ADD COLUMN especificacoes TEXT NOT NULL DEFAULT '{}'"
            )
        conexao.execute(
            "UPDATE registros SET data_conclusao = date('now', 'localtime') WHERE status != 'nao-iniciado' AND data_conclusao = ''"
        )
        conexao.execute(
            "CREATE INDEX IF NOT EXISTS idx_registros_local ON registros(torre, andar, unidade)"
        )
        conexao.execute(
            "CREATE INDEX IF NOT EXISTS idx_registros_atividade_local ON registros(atividade, torre, andar)"
        )
        conexao.execute(
            "CREATE INDEX IF NOT EXISTS idx_registros_atualizado ON registros(atualizado_em)"
        )
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS ocorrencias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                torre TEXT NOT NULL,
                andar INTEGER NOT NULL,
                unidade TEXT NOT NULL,
                descricao TEXT NOT NULL,
                foto TEXT NOT NULL DEFAULT '',
                foto_nome TEXT NOT NULL DEFAULT '',
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        colunas = {
            linha["name"]
            for linha in conexao.execute("PRAGMA table_info(ocorrencias)").fetchall()
        }
        if "data_ocorrencia" not in colunas:
            conexao.execute(
                "ALTER TABLE ocorrencias ADD COLUMN data_ocorrencia TEXT NOT NULL DEFAULT ''"
            )
            conexao.execute(
                "UPDATE ocorrencias SET data_ocorrencia = substr(criado_em, 1, 10) WHERE data_ocorrencia = ''"
            )
        if "status" not in colunas:
            conexao.execute(
                "ALTER TABLE ocorrencias ADD COLUMN status TEXT NOT NULL DEFAULT 'pendente'"
            )
        conexao.execute(
            "UPDATE ocorrencias SET status = 'pendente' WHERE status NOT IN ('pendente', 'concluido')"
        )
        if "atividade" not in colunas:
            conexao.execute(
                "ALTER TABLE ocorrencias ADD COLUMN atividade TEXT NOT NULL DEFAULT 'Não informada'"
            )
        if "subatividade" not in colunas:
            conexao.execute(
                "ALTER TABLE ocorrencias ADD COLUMN subatividade TEXT NOT NULL DEFAULT ''"
            )
        if "especificacao" not in colunas:
            conexao.execute(
                "ALTER TABLE ocorrencias ADD COLUMN especificacao TEXT NOT NULL DEFAULT ''"
            )
        conexao.execute(
            "CREATE INDEX IF NOT EXISTS idx_ocorrencias_local ON ocorrencias(torre, andar, unidade)"
        )
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS projetos_unidade (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                torre TEXT NOT NULL,
                andar INTEGER NOT NULL,
                unidade TEXT NOT NULL,
                titulo TEXT NOT NULL,
                imagem TEXT NOT NULL,
                arquivo_nome TEXT NOT NULL DEFAULT '',
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conexao.execute(
            "CREATE INDEX IF NOT EXISTS idx_projetos_unidade ON projetos_unidade(torre, andar, unidade)"
        )
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS comentarios_unidade (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                torre TEXT NOT NULL,
                andar INTEGER NOT NULL,
                unidade TEXT NOT NULL,
                comentario TEXT NOT NULL,
                autor TEXT NOT NULL DEFAULT '',
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conexao.execute(
            "CREATE INDEX IF NOT EXISTS idx_comentarios_unidade ON comentarios_unidade(torre, andar, unidade, criado_em)"
        )
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS projetos_ocultos (
                torre TEXT NOT NULL,
                andar INTEGER NOT NULL,
                unidade TEXT NOT NULL,
                imagem TEXT NOT NULL,
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(torre, andar, unidade, imagem)
            )
            """
        )
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS usuarios_acesso (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                email TEXT NOT NULL COLLATE NOCASE UNIQUE,
                cargo TEXT NOT NULL DEFAULT '',
                perfil TEXT NOT NULL,
                senha_hash TEXT NOT NULL,
                torres TEXT NOT NULL DEFAULT '["aurora", "horizonte"]',
                ativo INTEGER NOT NULL DEFAULT 1,
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                ultimo_acesso TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS auditoria_acessos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario_id INTEGER,
                usuario_nome TEXT NOT NULL,
                acao TEXT NOT NULL,
                detalhes TEXT NOT NULL DEFAULT '',
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conexao.execute(
            "CREATE INDEX IF NOT EXISTS idx_auditoria_acessos_data ON auditoria_acessos(criado_em DESC)"
        )
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS cargos_perfis (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cargo TEXT NOT NULL COLLATE NOCASE UNIQUE,
                perfil TEXT NOT NULL,
                ativo INTEGER NOT NULL DEFAULT 1,
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        cargos_iniciais = [
            ("Administrador da obra", "administrador"),
            ("Engenheiro responsável", "engenheiro-responsavel"),
            ("Engenheiro", "operacional"),
            ("Assistente de engenharia", "operacional"),
            ("Mestre de obras", "operacional"),
            ("Encarregado", "operacional"),
            ("Consultor", "consulta"),
            ("Cliente", "visitante"),
        ]
        conexao.executemany(
            "INSERT OR IGNORE INTO cargos_perfis (cargo, perfil) VALUES (?, ?)",
            cargos_iniciais,
        )
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS perfis_permissoes (
                perfil TEXT PRIMARY KEY,
                permissoes TEXT NOT NULL,
                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conexao.executemany(
            "INSERT OR IGNORE INTO perfis_permissoes (perfil, permissoes) VALUES (?, ?)",
            [(perfil, json.dumps(permissoes)) for perfil, permissoes in PERMISSOES_PADRAO.items()],
        )
        if USUARIO_ENGENHEIRO and SENHA_ENGENHEIRO:
            usuario_inicial = conexao.execute(
                "SELECT id FROM usuarios_acesso WHERE email = ? COLLATE NOCASE",
                (USUARIO_ENGENHEIRO,),
            ).fetchone()
            if not usuario_inicial:
                cursor = conexao.execute(
                    """
                    INSERT INTO usuarios_acesso (nome, email, cargo, perfil, senha_hash)
                    VALUES (?, ?, 'Administrador da obra', 'administrador', ?)
                    """,
                    (NOME_USUARIO_ENGENHEIRO, USUARIO_ENGENHEIRO, gerar_hash_senha(SENHA_ENGENHEIRO)),
                )
                registrar_auditoria(
                    conexao, cursor.lastrowid, NOME_USUARIO_ENGENHEIRO,
                    "Cadastro inicial", "Administrador inicial criado a partir da configuração do sistema."
                )
        importar_planejamento_inicial(conexao)
        preparar_atividades_config(conexao)


def apartamentos_do_andar(torre, andar):
    if andar in PAVIMENTOS_TECNICOS:
        return []
    if torre == "horizonte":
        return [f"Apto {andar}{numero:02d}" for numero in range(1, 15)] if andar >= 4 else []
    if andar == 1:
        return ["Apto 103", "Apto 104", "Apto 105"]
    if andar == 2:
        return ["Apto 203", "Apto 204", "Apto 205"]
    if 4 <= andar <= 35:
        return [f"Apto {andar}{numero:02d}" for numero in range(1, 11)]
    if andar == 36:
        return [f"Apto {numero}" for numero in (3001, 3003, 3004, 3005, 3006, 3008, 3009, 3010)]
    return []


def importar_planejamento_inicial(conexao):
    """Importa uma versão da planilha uma única vez, sem apagar edições posteriores."""
    if not ARQUIVO_PLANEJAMENTO.exists():
        return
    planejamento = json.loads(ARQUIVO_PLANEJAMENTO.read_text(encoding="utf-8"))
    conexao.execute(
        "CREATE TABLE IF NOT EXISTS configuracoes (chave TEXT PRIMARY KEY, valor TEXT NOT NULL)"
    )
    versao = planejamento.get("versao", "planilha-atual")
    chave_versao = "planejamento_importado"
    atual = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = ?", (chave_versao,)
    ).fetchone()
    if atual and atual["valor"] == versao:
        return
    for torre, andares in planejamento.get("torres", {}).items():
        for andar_texto, servicos in andares.items():
            andar = int(andar_texto)
            unidades = apartamentos_do_andar(torre, andar) or ["Área comum"]
            for unidade in unidades:
                for servico in servicos:
                    atividade = servico["atividade"]
                    status = servico["status"]
                    chave = f"{torre}|{andar}|{unidade}|{atividade}"
                    conexao.execute(
                        """
                        INSERT INTO registros
                        (chave, torre, andar, unidade, atividade, concluido,
                         data_conclusao, observacao, foto, foto_nome, status, atualizado_em)
                        VALUES (?, ?, ?, ?, ?, ?, '', '', '', '', ?, CURRENT_TIMESTAMP)
                        ON CONFLICT(chave) DO UPDATE SET
                            concluido=excluded.concluido,
                            status=excluded.status,
                            atualizado_em=CURRENT_TIMESTAMP
                        """,
                        (chave, torre, andar, unidade, atividade, status == "concluido", status),
                    )
    conexao.execute(
        "INSERT INTO configuracoes (chave, valor) VALUES (?, ?) "
        "ON CONFLICT(chave) DO UPDATE SET valor=excluded.valor",
        (chave_versao, versao),
    )


def normalizar_texto_fvs(valor):
    texto = unicodedata.normalize("NFD", str(valor or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto).strip().lower()


def ler_shared_strings_fvs(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
    itens = []
    for si in raiz.findall("a:si", NS_FVS):
        texto = "".join(el.text or "" for el in si.iter(f"{{{NS_FVS['a']}}}t"))
        itens.append(texto)
    return itens


def ler_planilha_fvs_em_rows(caminho):
    with ZipFile(caminho) as z:
        shared_strings = ler_shared_strings_fvs(z)
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target_map = {item.attrib["Id"]: item.attrib["Target"] for item in rels}
        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        for sheet in workbook.findall("a:sheets/a:sheet", NS_FVS):
            nome = sheet.attrib["name"]
            rel_id = sheet.attrib[f"{{{NS_FVS['r']}}}id"]
            target = target_map[rel_id]
            if not target.startswith("xl/"):
                target = "xl/" + target
            planilha = ET.fromstring(z.read(target))
            linhas = []
            for linha in planilha.findall(".//a:sheetData/a:row", NS_FVS):
                valores = []
                for celula in linha.findall("a:c", NS_FVS):
                    tipo = celula.attrib.get("t")
                    valor_celula = celula.find("a:v", NS_FVS)
                    valor = "" if valor_celula is None else (valor_celula.text or "")
                    if tipo == "s" and valor:
                        indice = int(valor)
                        valor = shared_strings[indice] if 0 <= indice < len(shared_strings) else ""
                    valores.append(str(valor).strip())
                if any(v for v in valores):
                    linhas.append(valores)
            yield nome, linhas


def extrair_servicos_fvs_da_planilha(caminho):
    servicos = {}
    for _, linhas in ler_planilha_fvs_em_rows(caminho):
        if not linhas:
            continue
        if not any("FVS" in str(v).upper() for v in linhas[0]):
            continue
        titulo = next((valor for valor in linhas[0] if valor.strip()), "")
        if " - " not in titulo:
            continue
        nome_servico = titulo.split(" - ", 1)[1].strip().title()
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


def servico_fvs_ja_representado_no_modelo(nome, nomes_existentes):
    nome_norm = normalizar_texto_fvs(nome)
    if not nome_norm:
        return True
    for existente in sorted(nomes_existentes, key=lambda item: (-len(item.split()), item)):
        if existente == nome_norm:
            return True
        if nome_norm.startswith(existente + " ") or nome_norm.endswith(" " + existente):
            return True
        if len(existente.split()) == 1 and existente in nome_norm.split():
            return True
    return False


def limpar_servicos_fvs_duplicados(conexao):
    # Atividades cadastradas pelo usuário são parte da configuração permanente da
    # obra. A rotina antiga removia tudo que não constasse numa lista fixa ao
    # reiniciar o servidor, apagando também registros e ocorrências relacionados.
    return 0


def sincronizar_atividades_fvs(conexao):
    return 0


def preparar_atividades_config(conexao):
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
        "CREATE INDEX IF NOT EXISTS idx_atividades_config_escopo ON atividades_config(torre, andar, unidade)"
    )
    conexao.execute(
        """CREATE TABLE IF NOT EXISTS subservicos_setores (
            atividade TEXT NOT NULL, subservico TEXT NOT NULL, torre TEXT NOT NULL,
            andar INTEGER NOT NULL, unidade TEXT NOT NULL,
            UNIQUE(atividade, subservico, torre, andar, unidade)
        )"""
    )
    conexao.execute("CREATE INDEX IF NOT EXISTS idx_subservicos_setores_local ON subservicos_setores(atividade, subservico, torre, andar)")
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
    marcador = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'atividades_config_inicial'"
    ).fetchone()
    if not marcador and ARQUIVO_PLANEJAMENTO.exists():
        dados = json.loads(ARQUIVO_PLANEJAMENTO.read_text(encoding="utf-8"))
        for torre, andares in dados.get("torres", {}).items():
            for andar, servicos in andares.items():
                for servico in servicos:
                    conexao.execute(
                        "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                        (servico["atividade"], torre, int(andar)),
                    )
    if not marcador:
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('atividades_config_inicial', '1')"
        )
    sincronizar_atividades_fvs(conexao)
    marcador_registros = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'atividades_registros_config_v1'"
    ).fetchone()
    if not marcador_registros:
        linhas = conexao.execute(
            "SELECT nome, torre, andar, unidade FROM atividades_config ORDER BY nome"
        ).fetchall()
        por_atividade = {}
        for linha in linhas:
            por_atividade.setdefault(linha["nome"], []).append(
                (linha["torre"], linha["andar"], linha["unidade"])
            )
        for nome, escopos in por_atividade.items():
            garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('atividades_registros_config_v1', '1')"
        )
    marcador_hidraulica = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'instalacoes_hidraulicas_v1'"
    ).fetchone()
    if not marcador_hidraulica and ARQUIVO_PLANEJAMENTO.exists():
        nome = "Instalações Hidráulicas"
        especificacoes = [
            "Condição para início dos serviços",
            "Posicionamento dos Ralos",
            "Posicionamento dos Pontos (Consumo de Água, Registros, Válvulas, Sistema de Esgoto e Gás)",
            "Prumadas",
            "Chumbamento das Tubulações e Ralos",
            "Posicionamento e Distribuição dos Esgotos (Aranha)",
            "Kit Hidráulico",
            "Tubos para Exaustão dos banhos com Ventilação Forçada",
            "Chumbamento de Passantes Fachada",
            "Encamisamento/Ventilação do Gás",
            "Cavaletes de Água e Redutoras de Pressão",
            "Pressurização da Tubulação de Gás (Antes do Fechamento)",
            "Pressurização da Tubulação de Gás (Depois da Conclusão dos Revestimentos Cerâmicos)",
            "Funcionamento do Sistema de Esgoto",
            "Funcionamento do Sistema de Combate a Incêndio",
            "Proteção dos Registros",
            "Medidor de Água Fria",
            "Limpeza e Proteção dos Ralos",
            "Instalação de Louças (Bacias)",
            "Acabamento de Registros",
            "Metais",
            "Sifão e Acessórios",
        ]
        planejamento = json.loads(ARQUIVO_PLANEJAMENTO.read_text(encoding="utf-8"))
        escopos = []
        for torre, andares in planejamento.get("torres", {}).items():
            for andar in andares:
                escopo = (torre, int(andar), "*")
                escopos.append(escopo)
                conexao.execute(
                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                    (nome, torre, int(andar)),
                )
        salvar_especificacoes_atividade(conexao, nome, especificacoes)
        garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('instalacoes_hidraulicas_v1', '1')"
        )
    marcador_remocao_hidraulica = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'remover_servicos_hidraulicos_legados_v1'"
    ).fetchone()
    if not marcador_remocao_hidraulica:
        servicos_removidos = [
            "Kit Hidráulico",
            "Louças e Metais",
            "Prumadas AP, AF, ES e Ventilação",
            "Prumadas de Incêndio, AQ e Gás",
            "Ramais e Aranhas",
            "Teste PEX",
            "Pex Aéreo",
        ]
        marcadores = ",".join("?" for _ in servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_config WHERE nome IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_especificacoes WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM registros WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('remover_servicos_hidraulicos_legados_v1', '1')"
        )
    marcador_eletrica = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'instalacoes_eletricas_v1'"
    ).fetchone()
    if not marcador_eletrica and ARQUIVO_PLANEJAMENTO.exists():
        nome = "Instalações Elétricas"
        especificacoes = [
            "Condição para início dos serviços",
            "Passagem e Diâmetro dos Eletrodutos Embutidos nas Paredes e Caixas Elétricas",
            "Quadros Elétricos",
            "Tomadas, Pontos de TV, Interruptores, Telefone e Interfone",
            "Tomadas 220V",
            "Prumadas de Cabos Busway",
            "Fiação e Arame Guia",
            "Proteção de Caixas de Elétrica",
        ]
        planejamento = json.loads(ARQUIVO_PLANEJAMENTO.read_text(encoding="utf-8"))
        escopos = []
        for torre, andares in planejamento.get("torres", {}).items():
            for andar in andares:
                escopos.append((torre, int(andar), "*"))
                conexao.execute(
                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                    (nome, torre, int(andar)),
                )
        salvar_especificacoes_atividade(conexao, nome, especificacoes)
        garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('instalacoes_eletricas_v1', '1')"
        )
    marcador_remocao_eletrica = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'remover_servicos_eletricos_legados_v1'"
    ).fetchone()
    if not marcador_remocao_eletrica:
        servicos_removidos = ["Acabamentos Elétricos", "Caixinha Elétrica"]
        marcadores = ",".join("?" for _ in servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_config WHERE nome IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_especificacoes WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM registros WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('remover_servicos_eletricos_legados_v1', '1')"
        )
    marcador_andares_especiais = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'todos_servicos_andares_especiais_v1'"
    ).fetchone()
    if not marcador_andares_especiais:
        atividades = [
            linha["nome"]
            for linha in conexao.execute("SELECT DISTINCT nome FROM atividades_config ORDER BY nome").fetchall()
        ]
        for nome in atividades:
            escopos = []
            for torre in TORRES_NOMES:
                for andar in (-1, 0):
                    escopo = (torre, andar, "*")
                    escopos.append(escopo)
                    conexao.execute(
                        "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                        (nome, torre, andar),
                    )
            garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('todos_servicos_andares_especiais_v1', '1')"
        )
    marcador_fundacao = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'andar_fundacao_estrutura_v1'"
    ).fetchone()
    if not marcador_fundacao:
        nome = "Estrutura"
        escopos = []
        for torre in TORRES_NOMES:
            escopo = (torre, -2, "*")
            escopos.append(escopo)
            conexao.execute(
                "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, -2, '*')",
                (nome, torre),
            )
        garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('andar_fundacao_estrutura_v1', '1')"
        )
    marcador_remocao_estrutura = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'remover_servicos_estrutura_legados_v1'"
    ).fetchone()
    if not marcador_remocao_estrutura:
        servicos_removidos = ["Perfil", "Hélice", "Escavação", "Bloco"]
        marcadores = ",".join("?" for _ in servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_config WHERE nome IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_especificacoes WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM registros WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('remover_servicos_estrutura_legados_v1', '1')"
        )
    marcador_alvenaria = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'servico_alvenaria_subservicos_v1'"
    ).fetchone()
    if not marcador_alvenaria:
        nome = "Alvenaria"
        especificacoes = [
            "Marcação Alvenaria :: Execução do serviço",
            "Elevação Alvenaria :: Execução do serviço",
            "Encunhamento :: Execução do serviço",
        ]
        escopos = []
        for torre in TORRES_NOMES:
            andares = [-1, 0] + list(range(1, 37 if torre == "aurora" else 24))
            for andar in andares:
                escopo = (torre, andar, "*")
                escopos.append(escopo)
                conexao.execute(
                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                    (nome, torre, andar),
                )
        salvar_especificacoes_atividade(conexao, nome, especificacoes)
        garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('servico_alvenaria_subservicos_v1', '1')"
        )
    marcador_marcacao_alvenaria = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_marcacao_alvenaria_v1'"
    ).fetchone()
    if not marcador_marcacao_alvenaria:
        salvar_especificacoes_atividade(
            conexao,
            "Alvenaria",
            [
                "Marcação Alvenaria :: Condições para Início dos Serviços",
                "Marcação Alvenaria :: Posicionamento da Fiada de Marcação",
                "Marcação Alvenaria :: Esquadro da Fiada de Marcação",
                "Marcação Alvenaria :: Alinhamento da Fiada de Marcação",
                "Marcação Alvenaria :: Tubulação de Dreno",
                "Marcação Alvenaria :: Condutes Elétricos",
                "Marcação Alvenaria :: Nivelamento da Fiada de Marcação",
                "Elevação Alvenaria :: Execução do serviço",
                "Encunhamento :: Execução do serviço",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_marcacao_alvenaria_v1', '1')"
        )
    marcador_elevacao_alvenaria = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_elevacao_alvenaria_v1'"
    ).fetchone()
    if not marcador_elevacao_alvenaria:
        salvar_especificacoes_atividade(
            conexao,
            "Alvenaria",
            [
                "Marcação Alvenaria :: Condições para Início dos Serviços",
                "Marcação Alvenaria :: Posicionamento da Fiada de Marcação",
                "Marcação Alvenaria :: Esquadro da Fiada de Marcação",
                "Marcação Alvenaria :: Alinhamento da Fiada de Marcação",
                "Marcação Alvenaria :: Tubulação de Dreno",
                "Marcação Alvenaria :: Condutes Elétricos",
                "Marcação Alvenaria :: Nivelamento da Fiada de Marcação",
                "Elevação Alvenaria :: Condições para Início dos Serviços",
                "Elevação Alvenaria :: Dreno do Ar-Condicionado",
                "Elevação Alvenaria :: Tela Galvanizada",
                "Elevação Alvenaria :: Preenchimento com Argamassa dos Blocos da Marcação",
                "Elevação Alvenaria :: Prumo da Alvenaria",
                "Elevação Alvenaria :: Planeza da Alvenaria",
                "Elevação Alvenaria :: Prumo de Vãos de Portas e Dimensões dos Vãos de Janelas e Portas",
                "Elevação Alvenaria :: Aspecto Geral da Alvenaria",
                "Elevação Alvenaria :: Espessura do Vão para Encunhamento",
                "Encunhamento :: Execução do serviço",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_elevacao_alvenaria_v1', '1')"
        )
    marcador_encunhamento_alvenaria = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_encunhamento_alvenaria_v1'"
    ).fetchone()
    if not marcador_encunhamento_alvenaria:
        salvar_especificacoes_atividade(
            conexao,
            "Alvenaria",
            [
                "Marcação Alvenaria :: Condições para Início dos Serviços",
                "Marcação Alvenaria :: Posicionamento da Fiada de Marcação",
                "Marcação Alvenaria :: Esquadro da Fiada de Marcação",
                "Marcação Alvenaria :: Alinhamento da Fiada de Marcação",
                "Marcação Alvenaria :: Tubulação de Dreno",
                "Marcação Alvenaria :: Condutes Elétricos",
                "Marcação Alvenaria :: Nivelamento da Fiada de Marcação",
                "Elevação Alvenaria :: Condições para Início dos Serviços",
                "Elevação Alvenaria :: Dreno do Ar-Condicionado",
                "Elevação Alvenaria :: Tela Galvanizada",
                "Elevação Alvenaria :: Preenchimento com Argamassa dos Blocos da Marcação",
                "Elevação Alvenaria :: Prumo da Alvenaria",
                "Elevação Alvenaria :: Planeza da Alvenaria",
                "Elevação Alvenaria :: Prumo de Vãos de Portas e Dimensões dos Vãos de Janelas e Portas",
                "Elevação Alvenaria :: Aspecto Geral da Alvenaria",
                "Elevação Alvenaria :: Espessura do Vão para Encunhamento",
                "Encunhamento :: Condições para Início dos Serviços",
                "Encunhamento :: Encunhamento",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_encunhamento_alvenaria_v1', '1')"
        )
    marcador_remocao_alvenaria_legada = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'remover_servicos_alvenaria_legados_v1'"
    ).fetchone()
    if not marcador_remocao_alvenaria_legada:
        servicos_removidos = [
            "Elevação da Alvenaria",
            "Encunhamento",
            "Marcação da Alvenaria",
        ]
        marcadores = ",".join("?" for _ in servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_config WHERE nome IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_especificacoes WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM registros WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('remover_servicos_alvenaria_legados_v1', '1')"
        )
    marcador_revestimento = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'servico_revestimento_subservicos_v1'"
    ).fetchone()
    if not marcador_revestimento:
        nome = "Revestimento"
        especificacoes = [
            "Massa Interna :: Execução do serviço",
            "Gesso Liso :: Execução do serviço",
            "Argamassa Externa :: Execução do serviço",
        ]
        escopos = []
        for torre in TORRES_NOMES:
            ultimo_andar = 36 if torre == "aurora" else 23
            for andar in [-2, -1, 0] + list(range(1, ultimo_andar + 1)):
                escopo = (torre, andar, "*")
                escopos.append(escopo)
                conexao.execute(
                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                    (nome, torre, andar),
                )
        salvar_especificacoes_atividade(conexao, nome, especificacoes)
        garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('servico_revestimento_subservicos_v1', '1')"
        )
    marcador_revestimento_ceramico = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'subservicos_revestimento_ceramico_v1'"
    ).fetchone()
    if not marcador_revestimento_ceramico:
        salvar_especificacoes_atividade(
            conexao,
            "Revestimento",
            [
                "Massa Interna :: Execução do serviço",
                "Gesso Liso :: Execução do serviço",
                "Argamassa Externa :: Execução do serviço",
                "Piso Cerâmico :: Execução do serviço",
                "Parede em Cerâmica :: Execução do serviço",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('subservicos_revestimento_ceramico_v1', '1')"
        )
    marcador_gesso_liso = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_gesso_liso_v1'"
    ).fetchone()
    if not marcador_gesso_liso:
        salvar_especificacoes_atividade(
            conexao,
            "Revestimento",
            [
                "Massa Interna :: Execução do serviço",
                "Gesso Liso :: Condições para Início do Serviço",
                "Gesso Liso :: Planicidade",
                "Gesso Liso :: Nivelamento e Prumo",
                "Gesso Liso :: Esquadro",
                "Gesso Liso :: Cantos Riscados",
                "Gesso Liso :: Tela Entre Alvenaria e Estrutura",
                "Gesso Liso :: Aspecto Final",
                "Argamassa Externa :: Execução do serviço",
                "Piso Cerâmico :: Execução do serviço",
                "Parede em Cerâmica :: Execução do serviço",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_gesso_liso_v1', '1')"
        )
    marcador_massa_interna = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_massa_interna_v1'"
    ).fetchone()
    if not marcador_massa_interna:
        salvar_especificacoes_atividade(
            conexao,
            "Revestimento",
            [
                "Massa Interna :: Condições para Início do Serviço",
                "Massa Interna :: Planicidade",
                "Massa Interna :: Nivelamento e Prumo",
                "Massa Interna :: Esquadro",
                "Massa Interna :: Cantos Riscados",
                "Massa Interna :: Tela Entre Alvenaria e Estrutura",
                "Massa Interna :: Aspecto Final",
                "Gesso Liso :: Condições para Início do Serviço",
                "Gesso Liso :: Planicidade",
                "Gesso Liso :: Nivelamento e Prumo",
                "Gesso Liso :: Esquadro",
                "Gesso Liso :: Cantos Riscados",
                "Gesso Liso :: Tela Entre Alvenaria e Estrutura",
                "Gesso Liso :: Aspecto Final",
                "Argamassa Externa :: Execução do serviço",
                "Piso Cerâmico :: Execução do serviço",
                "Parede em Cerâmica :: Execução do serviço",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_massa_interna_v1', '1')"
        )
    marcador_argamassa_externa = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_externo_argamassa_v1'"
    ).fetchone()
    if not marcador_argamassa_externa:
        salvar_especificacoes_atividade(
            conexao,
            "Revestimento",
            [
                "Massa Interna :: Condições para Início do Serviço",
                "Massa Interna :: Planicidade",
                "Massa Interna :: Nivelamento e Prumo",
                "Massa Interna :: Esquadro",
                "Massa Interna :: Cantos Riscados",
                "Massa Interna :: Tela Entre Alvenaria e Estrutura",
                "Massa Interna :: Aspecto Final",
                "Gesso Liso :: Condições para Início do Serviço",
                "Gesso Liso :: Planicidade",
                "Gesso Liso :: Nivelamento e Prumo",
                "Gesso Liso :: Esquadro",
                "Gesso Liso :: Cantos Riscados",
                "Gesso Liso :: Tela Entre Alvenaria e Estrutura",
                "Gesso Liso :: Aspecto Final",
                "Externo em Argamassa :: Condições para Início do Serviço",
                "Externo em Argamassa :: Posicionamento dos Arames",
                "Externo em Argamassa :: Requadração dos Vãos",
                "Externo em Argamassa :: Tela Metálica",
                "Externo em Argamassa :: Passantes na Fachada",
                "Externo em Argamassa :: Espessura do Revestimento",
                "Externo em Argamassa :: Planicidade, Prumo e Alinhamento",
                "Externo em Argamassa :: Aspecto Final",
                "Externo em Argamassa :: Ensaio de Aderência",
                "Externo em Argamassa :: Relatório de Execução",
                "Piso Cerâmico :: Execução do serviço",
                "Parede em Cerâmica :: Execução do serviço",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_externo_argamassa_v1', '1')"
        )
    marcador_piso_ceramico = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_piso_ceramico_v1'"
    ).fetchone()
    if not marcador_piso_ceramico:
        conexao.execute(
            "DELETE FROM atividades_especificacoes WHERE atividade = 'Revestimento' AND especificacao LIKE 'Piso Cerâmico :: %'"
        )
        conexao.execute(
            "UPDATE atividades_especificacoes SET ordem = 31 WHERE atividade = 'Revestimento' AND especificacao LIKE 'Parede em Cerâmica :: %'"
        )
        especificacoes = [
            "Condição para Início dos Serviços",
            "Juntas",
            "Planicidade e Nivelamento",
            "Condição para Rejuntamento",
            "Aspecto Final",
            "Caimento",
            "Proteção",
        ]
        for ordem, especificacao in enumerate(especificacoes, start=24):
            conexao.execute(
                "INSERT INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES ('Revestimento', ?, ?)",
                (f"Piso Cerâmico :: {especificacao}", ordem),
            )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_piso_ceramico_v1', '1')"
        )
    marcador_parede_ceramica = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_parede_ceramica_v1'"
    ).fetchone()
    if not marcador_parede_ceramica:
        conexao.execute(
            "DELETE FROM atividades_especificacoes WHERE atividade = 'Revestimento' AND especificacao LIKE 'Parede em Cerâmica :: %'"
        )
        especificacoes = [
            "Condição para Início dos Serviços",
            "Juntas",
            "Planicidade e Prumo do Revestimento",
            "Aderência das Placas",
            "Condição para Rejuntamento",
            "Aspecto Final",
        ]
        for ordem, especificacao in enumerate(especificacoes, start=31):
            conexao.execute(
                "INSERT INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES ('Revestimento', ?, ?)",
                (f"Parede em Cerâmica :: {especificacao}", ordem),
            )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_parede_ceramica_v1', '1')"
        )
    marcador_remocao_revestimentos_legados = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'remover_servicos_revestimentos_legados_v1'"
    ).fetchone()
    if not marcador_remocao_revestimentos_legados:
        servicos_removidos = [
            "Azulejo",
            "Piso Cerâmico - Hall",
            "Piso Interno",
            "Piso Sacada",
            "Rejunte",
        ]
        marcadores = ",".join("?" for _ in servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_config WHERE nome IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_especificacoes WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM registros WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('remover_servicos_revestimentos_legados_v1', '1')"
        )
    marcador_remocao_massas_legadas = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'remover_servicos_massas_legados_v1'"
    ).fetchone()
    if not marcador_remocao_massas_legadas:
        servicos_removidos = [
            "Gesso",
            "Massa Interna",
            "Massa Fachada",
            "Chapisco Fachada",
        ]
        marcadores = ",".join("?" for _ in servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_config WHERE nome IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_especificacoes WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM registros WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('remover_servicos_massas_legados_v1', '1')"
        )
    marcador_contrapiso = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'servico_contrapiso_subservicos_v1'"
    ).fetchone()
    if not marcador_contrapiso:
        nome = "Contrapiso"
        especificacoes = [
            "Contrapiso :: Execução do serviço",
            "Contrapiso Acústico :: Execução do serviço",
        ]
        escopos = []
        for torre in TORRES_NOMES:
            ultimo_andar = 36 if torre == "aurora" else 23
            for andar in [-1, 0] + list(range(1, ultimo_andar + 1)) + list(PAVIMENTOS_TECNICOS):
                escopo = (torre, andar, "*")
                escopos.append(escopo)
                conexao.execute(
                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                    (nome, torre, andar),
                )
        salvar_especificacoes_atividade(conexao, nome, especificacoes)
        garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('servico_contrapiso_subservicos_v1', '1')"
        )
    marcador_especificacoes_contrapiso = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_contrapiso_v1'"
    ).fetchone()
    if not marcador_especificacoes_contrapiso:
        salvar_especificacoes_atividade(
            conexao,
            "Contrapiso",
            [
                "Contrapiso :: Condições para Início dos Serviços",
                "Contrapiso :: Tela Soldada (para Locais de Execução sobre Camada de Regularização)",
                "Contrapiso :: Nível das Taliscas",
                "Contrapiso :: Acabamento da Superfície",
                "Contrapiso :: Planicidade",
                "Contrapiso :: Caimento em Direção aos Ralos",
                "Contrapiso :: Aderência da Base",
                "Contrapiso Acústico :: Execução do serviço",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_contrapiso_v1', '1')"
        )
    marcador_especificacoes_contrapiso_acustico = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_contrapiso_acustico_v1'"
    ).fetchone()
    if not marcador_especificacoes_contrapiso_acustico:
        salvar_especificacoes_atividade(
            conexao,
            "Contrapiso",
            [
                "Contrapiso :: Condições para Início dos Serviços",
                "Contrapiso :: Tela Soldada (para Locais de Execução sobre Camada de Regularização)",
                "Contrapiso :: Nível das Taliscas",
                "Contrapiso :: Acabamento da Superfície",
                "Contrapiso :: Planicidade",
                "Contrapiso :: Caimento em Direção aos Ralos",
                "Contrapiso :: Aderência da Base",
                "Contrapiso Acústico :: Condições para Início dos Serviços",
                "Contrapiso Acústico :: Nível das Taliscas",
                "Contrapiso Acústico :: Manta Acústica",
                "Contrapiso Acústico :: Tela Metálica",
                "Contrapiso Acústico :: Planicidade",
                "Contrapiso Acústico :: Acabamento da Superfície",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_contrapiso_acustico_v1', '1')"
        )
    marcador_esquadrias = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'servico_esquadrias_subservicos_v1'"
    ).fetchone()
    if not marcador_esquadrias:
        nome = "Esquadrias"
        especificacoes = [
            "Taliscas e Contramarco :: Execução do serviço",
            "Colocação de Batente Metálico :: Execução do serviço",
            "Colocação de Gradil :: Execução do serviço",
        ]
        escopos = []
        for torre in TORRES_NOMES:
            ultimo_andar = 36 if torre == "aurora" else 23
            for andar in [-1, 0] + list(range(1, ultimo_andar + 1)) + list(PAVIMENTOS_TECNICOS):
                escopo = (torre, andar, "*")
                escopos.append(escopo)
                conexao.execute(
                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                    (nome, torre, andar),
                )
        salvar_especificacoes_atividade(conexao, nome, especificacoes)
        garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('servico_esquadrias_subservicos_v1', '1')"
        )
    marcador_taliscas_contramarco = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_taliscas_contramarco_v1'"
    ).fetchone()
    if not marcador_taliscas_contramarco:
        salvar_especificacoes_atividade(
            conexao,
            "Esquadrias",
            [
                "Taliscas e Contramarco :: Condições para Início do Serviço",
                "Taliscas e Contramarco :: Personalização",
                "Taliscas e Contramarco :: Prumo das Taliscas",
                "Taliscas e Contramarco :: Esquadro das Taliscas",
                "Taliscas e Contramarco :: Esquadro do Contramarco em Relação à Fachada",
                "Taliscas e Contramarco :: Altura do Vão em Relação ao Piso Acabado",
                "Taliscas e Contramarco :: Grapas Nivelamento das Travessas",
                "Taliscas e Contramarco :: Checar Prumo e Esquadro do Contramarco",
                "Taliscas e Contramarco :: Checar o Chumbamento com o Gabarito e Remoção das Fitas Hellerman",
                "Colocação de Batente Metálico :: Execução do serviço",
                "Colocação de Gradil :: Execução do serviço",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_taliscas_contramarco_v1', '1')"
        )
    marcador_batente_metalico = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_batente_metalico_v1'"
    ).fetchone()
    if not marcador_batente_metalico:
        salvar_especificacoes_atividade(
            conexao,
            "Esquadrias",
            [
                "Taliscas e Contramarco :: Condições para Início do Serviço",
                "Taliscas e Contramarco :: Personalização",
                "Taliscas e Contramarco :: Prumo das Taliscas",
                "Taliscas e Contramarco :: Esquadro das Taliscas",
                "Taliscas e Contramarco :: Esquadro do Contramarco em Relação à Fachada",
                "Taliscas e Contramarco :: Altura do Vão em Relação ao Piso Acabado",
                "Taliscas e Contramarco :: Grapas Nivelamento das Travessas",
                "Taliscas e Contramarco :: Checar Prumo e Esquadro do Contramarco",
                "Taliscas e Contramarco :: Checar o Chumbamento com o Gabarito e Remoção das Fitas Hellerman",
                "Colocação de Batente Metálico :: Sentido de Abertura da Porta",
                "Colocação de Batente Metálico :: Preenchimento da Argamassa",
                "Colocação de Batente Metálico :: Prumo e Nível",
                "Colocação de Batente Metálico :: Fixação",
                "Colocação de Gradil :: Execução do serviço",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_batente_metalico_v1', '1')"
        )
    marcador_colocacao_gradil = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_colocacao_gradil_v1'"
    ).fetchone()
    if not marcador_colocacao_gradil:
        salvar_especificacoes_atividade(
            conexao,
            "Esquadrias",
            [
                "Taliscas e Contramarco :: Condições para Início do Serviço",
                "Taliscas e Contramarco :: Personalização",
                "Taliscas e Contramarco :: Prumo das Taliscas",
                "Taliscas e Contramarco :: Esquadro das Taliscas",
                "Taliscas e Contramarco :: Esquadro do Contramarco em Relação à Fachada",
                "Taliscas e Contramarco :: Altura do Vão em Relação ao Piso Acabado",
                "Taliscas e Contramarco :: Grapas Nivelamento das Travessas",
                "Taliscas e Contramarco :: Checar Prumo e Esquadro do Contramarco",
                "Taliscas e Contramarco :: Checar o Chumbamento com o Gabarito e Remoção das Fitas Hellerman",
                "Colocação de Batente Metálico :: Sentido de Abertura da Porta",
                "Colocação de Batente Metálico :: Preenchimento da Argamassa",
                "Colocação de Batente Metálico :: Prumo e Nível",
                "Colocação de Batente Metálico :: Fixação",
                "Colocação de Gradil :: Verificar o Prumo da Fachada e Locação",
                "Colocação de Gradil :: Verificar Alinhamento do Gradil",
                "Colocação de Gradil :: Nivelamento do Gradil",
                "Colocação de Gradil :: Prumo do Gradil",
                "Colocação de Gradil :: Verificar o Chumbamento/Firmeza (Fixação no Concreto/Alvenaria)",
                "Colocação de Gradil :: Distância e Medidas",
                "Colocação de Gradil :: Pintura ou Proteção de Gradis de Alumínio, Aço e suas Ligas",
                "Colocação de Gradil :: Ensaios",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_colocacao_gradil_v1', '1')"
        )
    marcador_caixilho_aluminio = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'subatividade_caixilho_aluminio_v1'"
    ).fetchone()
    if not marcador_caixilho_aluminio:
        proxima_ordem = conexao.execute(
            "SELECT COALESCE(MAX(ordem), -1) + 1 AS ordem FROM atividades_especificacoes WHERE atividade = 'Esquadrias'"
        ).fetchone()["ordem"]
        conexao.execute(
            "INSERT OR IGNORE INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES ('Esquadrias', ?, ?)",
            ("Caixilho de Alumínio :: Execução do serviço", proxima_ordem),
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('subatividade_caixilho_aluminio_v1', '1')"
        )
    marcador_remocao_esquadrias_legadas = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'remover_servicos_esquadrias_legados_v1'"
    ).fetchone()
    if not marcador_remocao_esquadrias_legadas:
        servicos_removidos = ["Caixilho", "Gradil de Ferro", "Talisca", "Contramarco"]
        marcadores = ",".join("?" for _ in servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_config WHERE nome IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_especificacoes WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(f"DELETE FROM registros WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('remover_servicos_esquadrias_legados_v1', '1')"
        )
    marcador_bancada_pedra_natural = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'servico_bancada_pedra_natural_v1'"
    ).fetchone()
    if not marcador_bancada_pedra_natural:
        nome = "Bancada de Pedra Natural"
        escopos = []
        for torre in TORRES_NOMES:
            ultimo_andar = 36 if torre == "aurora" else 23
            for andar in [-1, 0] + list(range(1, ultimo_andar + 1)) + list(PAVIMENTOS_TECNICOS):
                escopo = (torre, andar, "*")
                escopos.append(escopo)
                conexao.execute(
                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                    (nome, torre, andar),
                )
        salvar_especificacoes_atividade(conexao, nome, ["Execução do serviço"])
        garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('servico_bancada_pedra_natural_v1', '1')"
        )
    marcador_especificacoes_bancada_pedra = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_bancada_pedra_natural_v1'"
    ).fetchone()
    if not marcador_especificacoes_bancada_pedra:
        salvar_especificacoes_atividade(
            conexao,
            "Bancada de Pedra Natural",
            [
                "Condição de Início de Serviço",
                "Nível da Bancada",
                "Aspecto Final da Bancada",
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_bancada_pedra_natural_v1', '1')"
        )
    marcador_pedra_natural_subservicos = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'pedra_natural_subservicos_v1'"
    ).fetchone()
    if not marcador_pedra_natural_subservicos:
        nome_antigo = "Bancada de Pedra Natural"
        nome_novo = "Pedra Natural"
        conexao.execute(
            "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) "
            "SELECT ?, torre, andar, unidade FROM atividades_config WHERE nome = ?",
            (nome_novo, nome_antigo),
        )
        conexao.execute("DELETE FROM atividades_config WHERE nome = ?", (nome_antigo,))
        for linha in conexao.execute(
            "SELECT chave, torre, andar, unidade, especificacoes FROM registros WHERE atividade = ?", (nome_antigo,)
        ).fetchall():
            try:
                estados_anteriores = json.loads(linha["especificacoes"] or "{}")
            except (TypeError, json.JSONDecodeError):
                estados_anteriores = {}
            estados_novos = {
                (chave if " :: " in chave else f"Bancadas :: {chave}"): valor
                for chave, valor in estados_anteriores.items()
            }
            chave_nova = f'{linha["torre"]}|{linha["andar"]}|{linha["unidade"]}|{nome_novo}'
            conexao.execute(
                "UPDATE registros SET chave = ?, atividade = ?, especificacoes = ? WHERE chave = ?",
                (chave_nova, nome_novo, json.dumps(estados_novos, ensure_ascii=False), linha["chave"]),
            )
        conexao.execute(
            "UPDATE ocorrencias SET atividade = ?, subatividade = CASE "
            "WHEN subatividade IS NULL OR TRIM(subatividade) = '' THEN 'Bancadas' ELSE subatividade END, "
            "especificacao = CASE WHEN especificacao IS NOT NULL AND especificacao != '' "
            "AND instr(especificacao, ' :: ') = 0 THEN 'Bancadas :: ' || especificacao ELSE especificacao END "
            "WHERE atividade = ?",
            (nome_novo, nome_antigo),
        )
        salvar_especificacoes_atividade(
            conexao,
            nome_novo,
            [
                "Bancadas :: Condição de Início de Serviço",
                "Bancadas :: Nível da Bancada",
                "Bancadas :: Aspecto Final da Bancada",
                "Piso :: Execução do serviço",
            ],
        )
        conexao.execute("DELETE FROM atividades_especificacoes WHERE atividade = ?", (nome_antigo,))
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('pedra_natural_subservicos_v1', '1')"
        )
    marcador_servico_gesso = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'servico_gesso_subservicos_v1'"
    ).fetchone()
    if not marcador_servico_gesso:
        nome = "Gesso"
        especificacoes_gesso_liso = [
            "Condições para Início do Serviço",
            "Planicidade",
            "Nivelamento e Prumo",
            "Esquadro",
            "Cantos Riscados",
            "Tela Entre Alvenaria e Estrutura",
            "Aspecto Final",
        ]
        especificacoes = [f"Gesso Liso :: {item}" for item in especificacoes_gesso_liso] + [
            "Paredes em Drywall :: Execução do serviço",
            "Shaft em Drywall :: Execução do serviço",
            "Forro em Placas de Gesso Acartonado :: Execução do serviço",
        ]
        conexao.execute(
            """
            INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade)
            SELECT ?, torre, andar, unidade
            FROM atividades_config
            WHERE nome = 'Revestimento'
            """,
            (nome,),
        )
        escopos = [
            (linha["torre"], linha["andar"], linha["unidade"])
            for linha in conexao.execute(
                "SELECT torre, andar, unidade FROM atividades_config WHERE nome = ?",
                (nome,),
            ).fetchall()
        ]
        salvar_especificacoes_atividade(conexao, nome, especificacoes)
        garantir_registros_atividade(conexao, nome, escopos)

        for registro in conexao.execute(
            "SELECT torre, andar, unidade, especificacoes, data_conclusao FROM registros WHERE atividade = 'Revestimento'"
        ).fetchall():
            try:
                estados_revestimento = json.loads(registro["especificacoes"] or "{}")
            except (json.JSONDecodeError, TypeError):
                estados_revestimento = {}
            estados_gesso = {
                chave: valor for chave, valor in estados_revestimento.items()
                if chave.startswith("Gesso Liso :: ")
            }
            if not estados_gesso:
                continue
            chave_gesso = f"{registro['torre']}|{registro['andar']}|{registro['unidade']}|Gesso"
            linha_gesso = conexao.execute(
                "SELECT especificacoes FROM registros WHERE chave = ?", (chave_gesso,)
            ).fetchone()
            try:
                mapa_gesso = json.loads(linha_gesso["especificacoes"] or "{}") if linha_gesso else {}
            except (json.JSONDecodeError, TypeError):
                mapa_gesso = {}
            mapa_gesso.update(estados_gesso)
            estados_aplicaveis = [valor for valor in estados_gesso.values() if valor != "nao-aplicavel"]
            if any(valor == "pendente" for valor in estados_aplicaveis):
                status_gesso = "pendente"
            elif any(valor == "em-andamento" for valor in estados_aplicaveis):
                status_gesso = "em-andamento"
            elif estados_aplicaveis and all(valor == "concluido" for valor in estados_aplicaveis):
                status_gesso = "concluido"
            else:
                status_gesso = "nao-iniciado"
            conexao.execute(
                "UPDATE registros SET especificacoes=?, status=?, concluido=?, data_conclusao=?, atualizado_em=CURRENT_TIMESTAMP WHERE chave=?",
                (json.dumps(mapa_gesso, ensure_ascii=False), status_gesso, status_gesso == "concluido", registro["data_conclusao"], chave_gesso),
            )
            estados_revestimento = {
                chave: valor for chave, valor in estados_revestimento.items()
                if not chave.startswith("Gesso Liso :: ")
            }
            conexao.execute(
                "UPDATE registros SET especificacoes=?, atualizado_em=CURRENT_TIMESTAMP WHERE torre=? AND andar=? AND unidade=? AND atividade='Revestimento'",
                (json.dumps(estados_revestimento, ensure_ascii=False), registro["torre"], registro["andar"], registro["unidade"]),
            )

        conexao.execute(
            "DELETE FROM atividades_especificacoes WHERE atividade = 'Revestimento' AND especificacao LIKE 'Gesso Liso :: %'"
        )
        conexao.execute(
            "UPDATE ocorrencias SET atividade='Gesso' WHERE atividade='Revestimento' AND (subatividade='Gesso Liso' OR especificacao LIKE 'Gesso Liso :: %')"
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('servico_gesso_subservicos_v1', '1')"
        )
    marcador_forro_gesso_acartonado = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_forro_gesso_acartonado_v1'"
    ).fetchone()
    if not marcador_forro_gesso_acartonado:
        conexao.execute(
            "DELETE FROM atividades_especificacoes WHERE atividade='Gesso' AND especificacao='Forro em Placas de Gesso Acartonado :: Execução do serviço'"
        )
        proxima_ordem = conexao.execute(
            "SELECT COALESCE(MAX(ordem), -1) + 1 AS ordem FROM atividades_especificacoes WHERE atividade='Gesso'"
        ).fetchone()["ordem"]
        for deslocamento, especificacao in enumerate([
            "Marcação do Nível do Forro",
            "Tipo de Placa",
            "Fixação da Estrutura",
            "Planicidade",
            "Aspecto Geral",
        ]):
            conexao.execute(
                "INSERT OR IGNORE INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES ('Gesso', ?, ?)",
                (f"Forro em Placas de Gesso Acartonado :: {especificacao}", proxima_ordem + deslocamento),
            )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_forro_gesso_acartonado_v1', '1')"
        )
    marcador_paredes_drywall = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_paredes_drywall_v1'"
    ).fetchone()
    if not marcador_paredes_drywall:
        conexao.execute(
            "DELETE FROM atividades_especificacoes WHERE atividade='Gesso' AND especificacao='Paredes em Drywall :: Execução do serviço'"
        )
        proxima_ordem = conexao.execute(
            "SELECT COALESCE(MAX(ordem), -1) + 1 AS ordem FROM atividades_especificacoes WHERE atividade='Gesso'"
        ).fetchone()["ordem"]
        for deslocamento, especificacao in enumerate([
            "Condição para Início do Serviço",
            "Marcação das Paredes e Vão de Portas",
            "Reforço de Madeira",
            "Esquadro das Paredes",
            "Prumo das Paredes e Pontos Elétricos Antes do Fechamento",
            "Planicidade",
            "Cantoneira de Canto",
            "Aspecto Geral",
        ]):
            conexao.execute(
                "INSERT OR IGNORE INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES ('Gesso', ?, ?)",
                (f"Paredes em Drywall :: {especificacao}", proxima_ordem + deslocamento),
            )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_paredes_drywall_v1', '1')"
        )
    marcador_shaft_drywall = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_shaft_drywall_v1'"
    ).fetchone()
    if not marcador_shaft_drywall:
        conexao.execute(
            "DELETE FROM atividades_especificacoes WHERE atividade='Gesso' AND especificacao='Shaft em Drywall :: Execução do serviço'"
        )
        proxima_ordem = conexao.execute(
            "SELECT COALESCE(MAX(ordem), -1) + 1 AS ordem FROM atividades_especificacoes WHERE atividade='Gesso'"
        ).fetchone()["ordem"]
        for deslocamento, especificacao in enumerate([
            "Condição para Início do Serviço",
            "Transpasse e Posicionamento dos Montantes",
            "Prumo",
            "Planicidade",
            "Esquadro",
            "Reforço de Madeira",
            "Aspecto Geral",
        ]):
            conexao.execute(
                "INSERT OR IGNORE INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES ('Gesso', ?, ?)",
                (f"Shaft em Drywall :: {especificacao}", proxima_ordem + deslocamento),
            )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_shaft_drywall_v1', '1')"
        )
    marcador_impermeabilizacao = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'servico_impermeabilizacao_subservicos_v1'"
    ).fetchone()
    if not marcador_impermeabilizacao:
        nome = "Impermeabilização"
        especificacoes = [
            "Impermeabilização com Manta :: Execução do serviço",
            "Impermeabilização Interna :: Execução do serviço",
        ]
        escopos = []
        for torre in TORRES_NOMES:
            ultimo_andar = 36 if torre == "aurora" else 23
            for andar in [-1, 0] + list(range(1, ultimo_andar + 1)) + list(PAVIMENTOS_TECNICOS):
                escopo = (torre, andar, "*")
                escopos.append(escopo)
                conexao.execute(
                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                    (nome, torre, andar),
                )
        salvar_especificacoes_atividade(conexao, nome, especificacoes)
        garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('servico_impermeabilizacao_subservicos_v1', '1')"
        )
    marcador_impermeabilizacao_manta = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_impermeabilizacao_manta_v1'"
    ).fetchone()
    if not marcador_impermeabilizacao_manta:
        conexao.execute(
            "DELETE FROM atividades_especificacoes WHERE atividade='Impermeabilização' AND especificacao='Impermeabilização com Manta :: Execução do serviço'"
        )
        proxima_ordem = conexao.execute(
            "SELECT COALESCE(MAX(ordem), -1) + 1 AS ordem FROM atividades_especificacoes WHERE atividade='Impermeabilização'"
        ).fetchone()["ordem"]
        for deslocamento, especificacao in enumerate([
            "Condições de Início dos Serviços",
            "Regularização Acabada",
            "Pintura da Base",
            "Cola das Emendas",
            "Cantos e Detalhes",
            "Ralos",
            "Juntas de Dilatação",
            "Estanqueidade",
            "Proteção Mecânica",
        ]):
            conexao.execute(
                "INSERT OR IGNORE INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES ('Impermeabilização', ?, ?)",
                (f"Impermeabilização com Manta :: {especificacao}", proxima_ordem + deslocamento),
            )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_impermeabilizacao_manta_v1', '1')"
        )
    marcador_impermeabilizacao_interna = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_impermeabilizacao_interna_v1'"
    ).fetchone()
    if not marcador_impermeabilizacao_interna:
        conexao.execute(
            "DELETE FROM atividades_especificacoes WHERE atividade='Impermeabilização' AND especificacao='Impermeabilização Interna :: Execução do serviço'"
        )
        proxima_ordem = conexao.execute(
            "SELECT COALESCE(MAX(ordem), -1) + 1 AS ordem FROM atividades_especificacoes WHERE atividade='Impermeabilização'"
        ).fetchone()["ordem"]
        for deslocamento, especificacao in enumerate([
            "Condições de Início dos Serviços",
            "Reforço nos Ralos e Aplicação do Produto",
            "Falhas na Superfície",
            "Estanqueidade",
            "Proteção, Piso Acabado, Contrapiso com Caída e Piso “OCO” após 7 dias",
        ]):
            conexao.execute(
                "INSERT OR IGNORE INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES ('Impermeabilização', ?, ?)",
                (f"Impermeabilização Interna :: {especificacao}", proxima_ordem + deslocamento),
            )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_impermeabilizacao_interna_v1', '1')"
        )
    marcador_producao_argamassa = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'servico_producao_argamassa_subservicos_v1'"
    ).fetchone()
    if not marcador_producao_argamassa:
        nome = "Produção de Argamassa"
        subatividades = [
            "Revestimento Externo",
            "Contrapiso",
            "Contrapiso Acústico",
            "Revestimento Interno",
            "Assentamento de Alvenaria",
        ]
        itens = ["Equipamento", "Balde Dosador", "Traço"]
        especificacoes = [
            f"{subatividade} :: {item}"
            for subatividade in subatividades
            for item in itens
        ]
        escopos = []
        for torre in TORRES_NOMES:
            ultimo_andar = 36 if torre == "aurora" else 23
            for andar in [-1, 0] + list(range(1, ultimo_andar + 1)) + list(PAVIMENTOS_TECNICOS):
                escopo = (torre, andar, "*")
                escopos.append(escopo)
                conexao.execute(
                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                    (nome, torre, andar),
                )
        salvar_especificacoes_atividade(conexao, nome, especificacoes)
        garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('servico_producao_argamassa_subservicos_v1', '1')"
        )
    marcador_remocao_servicos_acabamentos_legados = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'remover_servicos_acabamentos_legados_v1'"
    ).fetchone()
    if not marcador_remocao_servicos_acabamentos_legados:
        servicos_removidos = [
            "Dry Wall - Áreas Secas",
            "Fechamento Shaft - A. Molhada + Terraço",
            "Forro Interno",
            "Forro Sacada",
            "Imperm. Interna - Piso",
            "Imperm. Shaft",
            "Montante Drywall",
            "Tampo",
        ]
        marcadores = ",".join("?" for _ in servicos_removidos)
        conexao.execute(f"DELETE FROM atividades_config WHERE nome IN ({marcadores})", servicos_removidos)
        conexao.execute(
            f"DELETE FROM atividades_especificacoes WHERE atividade IN ({marcadores})",
            servicos_removidos,
        )
        conexao.execute(f"DELETE FROM registros WHERE atividade IN ({marcadores})", servicos_removidos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('remover_servicos_acabamentos_legados_v1', '1')"
        )
    marcador_classificacao_servicos = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'classificacao_servicos_usuario_v1'"
    ).fetchone()
    if not marcador_classificacao_servicos:
        classificacao = {
            "Produção de Argamassa": ["Colante", "Revestimento Externo (Piscina)"],
            "Serviços Preliminares": ["Locação de Obra (Gabarito)", "Escavação"],
            "Revestimento": ["Pastilha", "Piso Intertravado"],
            "Esquadrias": ["Caixilho de Alumínio", "Portas Shafts"],
            "Gesso": ["Forro de Gesso"],
            "Pintura": ["PVA e Acrílica", "Tinta Esmalte Verniz"],
            "Fundação": ["Sapata Isolada", "Tubulão e Broca", "Hélice Contínua", "Estaca Strauss", "Perfil Metálico", "Parede Diafragma", "Compactação de Aterro", "Tirantes"],
            "Alvenaria": ["Muros Externos"],
            "Acabamentos": ["Lareira e Churrasqueira", "Louças e Metais"],
            "Piscina": [],
            "Paisagismo": [],
        }
        escopos_modelo = [
            (linha["torre"], linha["andar"], linha["unidade"])
            for linha in conexao.execute(
                "SELECT torre, andar, unidade FROM atividades_config WHERE nome='Contrapiso'"
            ).fetchall()
        ]
        for atividade, subservicos in classificacao.items():
            existe = conexao.execute(
                "SELECT 1 FROM atividades_config WHERE nome=? LIMIT 1", (atividade,)
            ).fetchone()
            if not existe:
                for torre, andar, unidade in escopos_modelo:
                    conexao.execute(
                        "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, ?)",
                        (atividade, torre, andar, unidade),
                    )
                garantir_registros_atividade(conexao, atividade, escopos_modelo)
            conexao.execute("DELETE FROM atividades_especificacoes WHERE atividade=?", (atividade,))
            especificacoes = [f"{subservico} :: Execução do serviço" for subservico in subservicos] or ["Execução do serviço"]
            for ordem, especificacao in enumerate(especificacoes):
                conexao.execute(
                    "INSERT INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES (?, ?, ?)",
                    (atividade, especificacao, ordem),
                )
        conexao.execute("UPDATE ocorrencias SET atividade='Esquadrias', subatividade='Portas Shafts' WHERE atividade='Portas Shafts'")
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('classificacao_servicos_usuario_v1', '1')"
        )
    marcador_pavimentos_tecnicos = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'pavimentos_tecnicos_todos_servicos_v1'"
    ).fetchone()
    if not marcador_pavimentos_tecnicos:
        atividades = [
            linha["nome"]
            for linha in conexao.execute("SELECT DISTINCT nome FROM atividades_config ORDER BY nome").fetchall()
        ]
        for nome in atividades:
            escopos = []
            for torre in TORRES_NOMES:
                for andar in PAVIMENTOS_TECNICOS:
                    escopo = (torre, andar, "*")
                    escopos.append(escopo)
                    conexao.execute(
                        "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, '*')",
                        (nome, torre, andar),
                    )
            garantir_registros_atividade(conexao, nome, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('pavimentos_tecnicos_todos_servicos_v1', '1')"
        )
    marcador_correcao_pavimentos_tecnicos = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'corrigir_unidades_pavimentos_tecnicos_v1'"
    ).fetchone()
    if not marcador_correcao_pavimentos_tecnicos:
        marcadores = ",".join("?" for _ in PAVIMENTOS_TECNICOS)
        conexao.execute(
            f"DELETE FROM registros WHERE andar IN ({marcadores}) AND unidade != 'Área comum'",
            tuple(PAVIMENTOS_TECNICOS),
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('corrigir_unidades_pavimentos_tecnicos_v1', '1')"
        )
    marcador_servicos_pdf = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_servicos_pdf_v1'"
    ).fetchone()
    if not marcador_servicos_pdf:
        especificacoes_pdf = {
            "Instalações Hidráulicas": [
                "Condição para início dos serviços", "Posicionamento dos Ralos",
                "Posicionamento dos Pontos (Consumo de Água, Registros, Válvulas, Sistema de Esgoto e Gás)",
                "Prumadas", "Chumbamento das Tubulações e Ralos",
                "Posicionamento dos Drenos na Marcação da Alvenaria",
                "Posicionamento e Distribuição dos Esgotos (Aranha)", "Kit Hidráulico",
                "Tubos para Exaustão dos banhos com Ventilação Forçada", "Chumbamento de Passantes Fachada",
                "Encamisamento/Ventilação do Gás", "Cavaletes de Água e Redutoras de Pressão",
                "Pressurização de Tubulações de Água Fria e Quente",
                "Pressurização da Tubulação de Gás (Antes do Fechamento)",
                "Pressurização da Tubulação de Gás (Depois da Conclusão dos Revestimentos Cerâmicos)",
                "Funcionamento do Sistema de Esgoto", "Funcionamento do Sistema de Combate a Incêndio",
                "Proteção dos Registros", "Medidor de Água Fria", "Limpeza e Proteção dos Ralos",
                "Instalação de Louças (Bacias)", "Acabamento de Registros", "Metais", "Sifão e Acessórios",
            ],
            "Esquadrias": [
                "Taliscas e Contramarco :: Condições para Início do Serviço",
                "Taliscas e Contramarco :: Personalização", "Taliscas e Contramarco :: Prumo das Taliscas",
                "Taliscas e Contramarco :: Esquadro das Taliscas",
                "Taliscas e Contramarco :: Esquadro do Contramarco em Relação à Fachada",
                "Taliscas e Contramarco :: Altura do Vão em Relação ao Piso Acabado",
                "Taliscas e Contramarco :: Grapas", "Taliscas e Contramarco :: Nivelamento das Travessas",
                "Taliscas e Contramarco :: Checar Prumo e Esquadro do Contramarco",
                "Taliscas e Contramarco :: Checar o Chumbamento com o Gabarito e Remoção das Fitas Hellerman",
                "Colocação de Batente Metálico :: Sentido de Abertura da Porta",
                "Colocação de Batente Metálico :: Preenchimento da Argamassa",
                "Colocação de Batente Metálico :: Prumo e Nível", "Colocação de Batente Metálico :: Fixação",
                "Colocação de Gradil :: Verificar o Prumo da Fachada e Locação",
                "Colocação de Gradil :: Verificar Alinhamento do Gradil",
                "Colocação de Gradil :: Nivelamento do Gradil", "Colocação de Gradil :: Prumo do Gradil",
                "Colocação de Gradil :: Verificar o Chumbamento/Firmeza (Fixação no Concreto/Alvenaria)",
                "Colocação de Gradil :: Distância e Medidas",
                "Colocação de Gradil :: Pintura ou Proteção de Gradis de Alumínio, Aço e suas Ligas",
                "Colocação de Gradil :: Ensaios",
            ],
            "Estrutura": [
                "Piso Armado :: Condições para Início do Serviço", "Piso Armado :: Espessura da Camada de Brita",
                "Piso Armado :: Lona Plástica Preta Reforçada", "Piso Armado :: Isopor", "Piso Armado :: Armadura",
                "Piso Armado :: Lançamento do concreto", "Piso Armado :: Cura Química",
                "Piso Armado :: Dreno nas Cortinas", "Piso Armado :: Aplicação do Selante (1,0cm)",
                "Piso Armado :: Juntas de Dilatação (Juntas de Construção)",
                "Lajes :: Cabo de Aço para Atrelar o Cinto de Segurança – Linha de Vida",
                "Lajes :: Transferência dos Eixos Principais (Para Assoalho)",
                "Lajes :: Posicionamento do Assoalho em Relação aos Eixos Principais", "Lajes :: Escoramento",
                "Lajes :: Armação", "Lajes :: Arranques", "Lajes :: Itens de Segurança",
                "Lajes :: Passagens em Lajes", "Lajes :: Travamento e Encaixe dos Painéis e Nivelamento das Lajes",
                "Lajes :: Limpeza do Assoalho", "Lajes :: Componentes Embutidos", "Lajes :: Lançamento do Concreto",
                "Lajes :: Cura Química", "Lajes :: Reescoramento",
                "Lajes :: Falhas de Concretagem e Aspecto Geral da Laje", "Lajes :: Limpeza dos Painéis",
            ],
        }
        for atividade, especificacoes in especificacoes_pdf.items():
            conexao.execute("DELETE FROM atividades_especificacoes WHERE atividade = ?", (atividade,))
            conexao.executemany(
                "INSERT INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES (?, ?, ?)",
                [(atividade, especificacao, ordem) for ordem, especificacao in enumerate(especificacoes)],
            )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_servicos_pdf_v1', '1')"
        )
    marcador_loucas_metais = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_loucas_metais_v1'"
    ).fetchone()
    if not marcador_loucas_metais:
        especificacoes = [
            "Lareira e Churrasqueira :: Execução do serviço",
            "Louças e Metais :: Condição de Início de Serviço",
            "Louças e Metais :: Posicionamento da Tubulação de Esgoto da Bacia",
            "Louças e Metais :: Fixação dos Metais e Louças Sanitários",
            "Louças e Metais :: Teste de Funcionamento e Vazamento",
            "Louças e Metais :: Aparência do Conjunto Após Instalação: Metais, Louças e Acabamento dos Rejuntes das Louças",
            "Louças e Metais :: Distância entre a Caixa Acoplada e Parede",
            "Louças e Metais :: Nivelamento das Louças (Bacia c/ Caixa Acoplada, Lavatório e Tanque)",
            "Louças e Metais :: Altura do Lavatório e Tanque (Conforme Medidas de Projeto)",
            "Louças e Metais :: Proteção dos Metais (Torneiras e Misturadores)",
        ]
        conexao.execute("DELETE FROM atividades_especificacoes WHERE atividade = 'Acabamentos'")
        conexao.executemany(
            "INSERT INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES ('Acabamentos', ?, ?)",
            [(especificacao, ordem) for ordem, especificacao in enumerate(especificacoes)],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_loucas_metais_v1', '1')"
        )
    marcador_sapata_isolada = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_sapata_isolada_v1'"
    ).fetchone()
    if not marcador_sapata_isolada:
        especificacoes_sapata = [
            "Sapata Isolada :: Checar a Locação da Sapata",
            "Sapata Isolada :: Cota do Fundo da Vala",
            "Sapata Isolada :: Forma da Sapata",
            "Sapata Isolada :: Armação da Sapata",
            "Sapata Isolada :: Gastalho",
            "Sapata Isolada :: Arranque do Pilar (ou Vigas Saindo das Sapatas)",
            "Sapata Isolada :: Concretagem",
            "Sapata Isolada :: Mapeamento",
        ]
        outras_especificacoes = [
            "Tubulão e Broca :: Execução do serviço", "Hélice Contínua :: Execução do serviço",
            "Estaca Strauss :: Execução do serviço", "Perfil Metálico :: Execução do serviço",
            "Parede Diafragma :: Execução do serviço", "Compactação de Aterro :: Execução do serviço",
            "Tirantes :: Execução do serviço",
        ]
        conexao.execute("DELETE FROM atividades_especificacoes WHERE atividade = 'Fundação'")
        conexao.executemany(
            "INSERT INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES ('Fundação', ?, ?)",
            [(especificacao, ordem) for ordem, especificacao in enumerate(especificacoes_sapata + outras_especificacoes)],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_sapata_isolada_v1', '1')"
        )
    marcador_tubulao_broca = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'especificacoes_tubulao_broca_v1'"
    ).fetchone()
    if not marcador_tubulao_broca:
        especificacoes_tubulao_broca = [
            "Tubulão e Broca :: Condições de Início",
            "Tubulão e Broca :: Diâmetro",
            "Tubulão e Broca :: Cota de Apoio",
            "Tubulão e Broca :: Dimensões da Base (Apenas para Tubulão)",
            "Tubulão e Broca :: Checagem da armadura",
            "Tubulão e Broca :: Cota de Arrasamento",
        ]
        linhas_fundacao = conexao.execute(
            "SELECT especificacao FROM atividades_especificacoes WHERE atividade = 'Fundação' ORDER BY ordem, id"
        ).fetchall()
        especificacoes_fundacao = [
            linha["especificacao"] for linha in linhas_fundacao
            if not linha["especificacao"].startswith("Tubulão e Broca ::")
        ]
        indice_tubulao = next(
            (indice + 1 for indice, especificacao in enumerate(especificacoes_fundacao)
             if especificacao.startswith("Sapata Isolada ::")),
            len(especificacoes_fundacao),
        )
        while (
            indice_tubulao < len(especificacoes_fundacao)
            and especificacoes_fundacao[indice_tubulao].startswith("Sapata Isolada ::")
        ):
            indice_tubulao += 1
        especificacoes_fundacao[indice_tubulao:indice_tubulao] = especificacoes_tubulao_broca
        conexao.execute("DELETE FROM atividades_especificacoes WHERE atividade = 'Fundação'")
        conexao.executemany(
            "INSERT INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES ('Fundação', ?, ?)",
            [(especificacao, ordem) for ordem, especificacao in enumerate(especificacoes_fundacao)],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('especificacoes_tubulao_broca_v1', '1')"
        )
    marcador_lista_mestra_fvs = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'lista_mestra_fvs_completa_v1'"
    ).fetchone()
    catalogo_fvs = PASTA / "catalogo_fvs.json"
    if not marcador_lista_mestra_fvs and catalogo_fvs.exists():
        dados_fvs = json.loads(catalogo_fvs.read_text(encoding="utf-8"))
        for atividade, subservicos in dados_fvs.get("catalogo", {}).items():
            especificacoes = []
            for subservico, itens in subservicos.items():
                especificacoes.extend(
                    f"{subservico} :: {item}" if subservico else item
                    for item in itens
                )
            if especificacoes:
                salvar_especificacoes_atividade(conexao, atividade, especificacoes)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('lista_mestra_fvs_completa_v1', '1')"
        )
    marcador_fundacao_somente_pavimento = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'fundacao_somente_pavimento_fundacao_v1'"
    ).fetchone()
    if not marcador_fundacao_somente_pavimento:
        atividade = "Fundação"
        escopos = []
        for torre in TORRES_NOMES:
            conexao.execute(
                "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, -2, '*')",
                (atividade, torre),
            )
            escopos.append((torre, -2, "*"))
        conexao.execute(
            "DELETE FROM atividades_config WHERE nome = ? AND andar != -2", (atividade,)
        )
        conexao.execute(
            "DELETE FROM registros WHERE atividade = ? AND andar != -2", (atividade,)
        )
        garantir_registros_atividade(conexao, atividade, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('fundacao_somente_pavimento_fundacao_v1', '1')"
        )
    marcador_preliminares_somente_fundacao = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'preliminares_somente_pavimento_fundacao_v1'"
    ).fetchone()
    if not marcador_preliminares_somente_fundacao:
        atividade = "Serviços Preliminares"
        escopos = []
        for torre in TORRES_NOMES:
            conexao.execute(
                "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, -2, '*')",
                (atividade, torre),
            )
            escopos.append((torre, -2, "*"))
        conexao.execute(
            "DELETE FROM atividades_config WHERE nome = ? AND andar != -2", (atividade,)
        )
        conexao.execute(
            "DELETE FROM registros WHERE atividade = ? AND andar != -2", (atividade,)
        )
        garantir_registros_atividade(conexao, atividade, escopos)
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('preliminares_somente_pavimento_fundacao_v1', '1')"
        )
    marcador_pavimento_fundacao_exclusivo = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'pavimento_fundacao_exclusivo_v1'"
    ).fetchone()
    if not marcador_pavimento_fundacao_exclusivo:
        atividades_permitidas = ("Fundação", "Serviços Preliminares")
        conexao.execute(
            "DELETE FROM atividades_config WHERE andar = -2 AND nome NOT IN (?, ?)",
            atividades_permitidas,
        )
        conexao.execute(
            "DELETE FROM registros WHERE andar = -2 AND atividade NOT IN (?, ?)",
            atividades_permitidas,
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('pavimento_fundacao_exclusivo_v1', '1')"
        )
    marcador_setores_revestimento_externo = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'setores_revestimento_externo_blocos_v1'"
    ).fetchone()
    if not marcador_setores_revestimento_externo:
        atividade = "Revestimento"
        subservico = "Externo em Argamassa"
        setores_por_torre = {
            "horizonte": [f"B{numero:02d}" for numero in range(1, 16)],
            "aurora": [f"B{numero:02d}" for numero in range(1, 20)],
        }
        conexao.execute(
            "DELETE FROM subservicos_setores WHERE atividade=? AND subservico=?",
            (atividade, subservico),
        )
        pavimentos = conexao.execute(
            "SELECT DISTINCT torre, andar FROM atividades_config WHERE nome=? ORDER BY torre, andar",
            (atividade,),
        ).fetchall()
        conexao.executemany(
            "INSERT INTO subservicos_setores (atividade, subservico, torre, andar, unidade) VALUES (?, ?, ?, ?, ?)",
            [
                (atividade, subservico, linha["torre"], linha["andar"], setor)
                for linha in pavimentos
                for setor in setores_por_torre.get(linha["torre"], [])
            ],
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('setores_revestimento_externo_blocos_v1', '1')"
        )
    marcador_ocultar_forro_gesso = conexao.execute(
        "SELECT valor FROM configuracoes WHERE chave = 'ocultar_forro_gesso_obra_v1'"
    ).fetchone()
    if not marcador_ocultar_forro_gesso:
        conexao.execute(
            "DELETE FROM subservicos_setores WHERE atividade='Gesso' AND subservico='Forro de Gesso'"
        )
        conexao.execute(
            "INSERT INTO configuracoes (chave, valor) VALUES ('ocultar_forro_gesso_obra_v1', '1')"
        )


def listar_atividades_config():
    with conectar() as conexao:
        marcadores = ",".join("?" for _ in SERVICOS_LEGADOS_OCULTOS)
        linhas = conexao.execute(
            f"SELECT id, nome, torre, andar, unidade FROM atividades_config WHERE nome NOT IN ({marcadores}) ORDER BY nome, torre, andar, unidade",
            tuple(SERVICOS_LEGADOS_OCULTOS),
        ).fetchall()
        linhas_especificacoes = conexao.execute(
            "SELECT atividade, especificacao FROM atividades_especificacoes ORDER BY atividade, ordem, id"
        ).fetchall()
        linhas_setores_subservicos = conexao.execute(
            "SELECT atividade, subservico, torre, andar, unidade FROM subservicos_setores ORDER BY atividade, subservico, torre, andar, unidade"
        ).fetchall()
    agrupadas = {}
    for linha in linhas:
        item = agrupadas.setdefault(linha["nome"], {"nome": linha["nome"], "escopos": [], "especificacoes": [], "setoresSubservicos": [], "subservicosSemSetores": []})
        item["escopos"].append(
            {"id": linha["id"], "torre": linha["torre"], "andar": linha["andar"], "unidade": linha["unidade"]}
        )
    for linha in linhas_especificacoes:
        if linha["atividade"] in agrupadas:
            agrupadas[linha["atividade"]]["especificacoes"].append(linha["especificacao"])
    for linha in linhas_setores_subservicos:
        if linha["atividade"] in agrupadas:
            if linha["unidade"] == MARCADOR_SEM_SETORES:
                agrupadas[linha["atividade"]]["subservicosSemSetores"].append({
                    "subservico": linha["subservico"], "torre": linha["torre"], "andar": linha["andar"]
                })
            else:
                agrupadas[linha["atividade"]]["setoresSubservicos"].append(dict(linha))
    return list(agrupadas.values())


def setores_efetivos_subservico(conexao, atividade, subservico, torre, andar):
    personalizados = conexao.execute(
        "SELECT unidade FROM subservicos_setores WHERE atividade=? AND subservico=? AND torre=? AND andar=? ORDER BY unidade",
        (atividade, subservico, torre, andar),
    ).fetchall()
    if personalizados:
        unidades = [linha["unidade"] for linha in personalizados]
        return [] if MARCADOR_SEM_SETORES in unidades else unidades
    escopos = conexao.execute(
        "SELECT unidade FROM atividades_config WHERE nome=? AND torre=? AND andar=? ORDER BY unidade",
        (atividade, torre, andar),
    ).fetchall()
    unidades = []
    for linha in escopos:
        if linha["unidade"] == "*":
            unidades.extend(apartamentos_do_andar(torre, andar) + ["Área comum"])
        else:
            unidades.append(linha["unidade"])
    return list(dict.fromkeys(unidades))


def salvar_setores_subservico(conexao, atividade, subservico, torre, andar, unidades):
    conexao.execute(
        "DELETE FROM subservicos_setores WHERE atividade=? AND subservico=? AND torre=? AND andar=?",
        (atividade, subservico, torre, andar),
    )
    unidades_unicas = list(dict.fromkeys(unidades)) or [MARCADOR_SEM_SETORES]
    conexao.executemany(
        "INSERT INTO subservicos_setores (atividade,subservico,torre,andar,unidade) VALUES (?,?,?,?,?)",
        [(atividade, subservico, torre, andar, unidade) for unidade in unidades_unicas],
    )


def validar_escopos(escopos):
    resultado = []
    for escopo in escopos:
        torre = str(escopo.get("torre", ""))
        andar = int(escopo.get("andar", 0))
        unidade = str(escopo.get("unidade", "*")).strip() or "*"
        maximo = 36 if torre == "aurora" else 23 if torre == "horizonte" else 0
        if not maximo or andar < -2 or (andar > maximo and andar not in PAVIMENTOS_TECNICOS):
            raise ValueError("Escopo de torre ou pavimento inválido")
        resultado.append((torre, andar, unidade))
    if not resultado:
        raise ValueError("Adicione ao menos um pavimento, apartamento ou setor")
    return list(dict.fromkeys(resultado))


def validar_especificacoes(especificacoes):
    resultado = []
    for especificacao in especificacoes:
        texto = str(especificacao).strip()
        if texto and texto not in resultado:
            resultado.append(texto)
    if not resultado:
        raise ValueError("Adicione ao menos uma especificação do serviço")
    return resultado


def salvar_especificacoes_atividade(conexao, nome, especificacoes):
    conexao.execute("DELETE FROM atividades_especificacoes WHERE atividade = ?", (nome,))
    for ordem, especificacao in enumerate(especificacoes):
        conexao.execute(
            "INSERT INTO atividades_especificacoes (atividade, especificacao, ordem) VALUES (?, ?, ?)",
            (nome, especificacao, ordem),
        )


def garantir_registros_atividade(conexao, nome, escopos):
    for torre, andar, unidade in escopos:
        if unidade == "*":
            unidades = apartamentos_do_andar(torre, andar) + ["Área comum"]
            unidades += [
                linha["unidade"]
                for linha in conexao.execute(
                    "SELECT DISTINCT unidade FROM atividades_config WHERE torre = ? AND andar = ? AND unidade != '*'",
                    (torre, andar),
                ).fetchall()
            ]
            unidades = list(dict.fromkeys(unidades))
        else:
            unidades = [unidade]
        for destino in unidades:
            chave = f"{torre}|{andar}|{destino}|{nome}"
            conexao.execute(
                """
                INSERT OR IGNORE INTO registros
                (chave, torre, andar, unidade, atividade, concluido, data_conclusao,
                 observacao, foto, foto_nome, status, atualizado_em)
                VALUES (?, ?, ?, ?, ?, 0, '', '', '', '', 'nao-iniciado', CURRENT_TIMESTAMP)
                """,
                (chave, torre, andar, destino, nome),
            )


class ServidorObra(SimpleHTTPRequestHandler):
    def url_publica(self, caminho, consulta=""):
        base_configurada = os.environ.get("OBRA_URL_PUBLICA", "").rstrip("/")
        if base_configurada:
            base = base_configurada
        else:
            protocolo = self.headers.get("X-Forwarded-Proto", "http").split(",", 1)[0].strip()
            host = self.headers.get("X-Forwarded-Host", self.headers.get("Host", f"127.0.0.1:{PORTA}"))
            base = f"{protocolo}://{host}"
        sufixo = f"?{consulta}" if consulta else ""
        return f"{base}{caminho}{sufixo}"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(PASTA), **kwargs)

    def end_headers(self):
        if urlparse(self.path).path in {"/", "/acompanhamento_obra.html"}:
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
        super().end_headers()

    def enviar_json(self, dados, status=200):
        corpo = json.dumps(dados, ensure_ascii=False).encode("utf-8")
        aceita_gzip = "gzip" in self.headers.get("Accept-Encoding", "").lower()
        compactado = aceita_gzip and len(corpo) >= 1024
        if compactado:
            corpo = gzip.compress(corpo, compresslevel=5)
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        if compactado:
            self.send_header("Content-Encoding", "gzip")
            self.send_header("Vary", "Accept-Encoding")
        self.send_header("Content-Length", str(len(corpo)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(corpo)

    def enviar_json_cacheado(self, corpo, corpo_gzip, etag):
        if self.headers.get("If-None-Match", "").strip() == etag:
            self.send_response(304)
            self.send_header("ETag", etag)
            self.send_header("Cache-Control", "private, no-cache")
            self.end_headers()
            return
        aceita_gzip = "gzip" in self.headers.get("Accept-Encoding", "").lower()
        resposta = corpo_gzip if aceita_gzip else corpo
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("ETag", etag)
        self.send_header("Cache-Control", "private, no-cache")
        if aceita_gzip:
            self.send_header("Content-Encoding", "gzip")
            self.send_header("Vary", "Accept-Encoding")
        self.send_header("Content-Length", str(len(resposta)))
        self.end_headers()
        self.wfile.write(resposta)

    def sessao_usuario(self):
        cookies = self.headers.get("Cookie", "")
        token = next(
            (parte.split("=", 1)[1] for parte in cookies.split("; ") if parte.startswith("sessao_obra=")),
            "",
        )
        sessao = SESSOES_ENGENHEIRO.get(token)
        if not sessao:
            return None
        with conectar() as conexao:
            linha = conexao.execute(
                "SELECT id, nome, email, perfil, ativo FROM usuarios_acesso WHERE id=?",
                (sessao["id"],),
            ).fetchone()
            permissoes = permissoes_do_perfil(conexao, linha["perfil"]) if linha and linha["ativo"] else []
        if not linha or not linha["ativo"]:
            SESSOES_ENGENHEIRO.pop(token, None)
            return None
        usuario = {"id": linha["id"], "nome": linha["nome"], "email": linha["email"], "perfil": linha["perfil"], "permissoes": permissoes}
        SESSOES_ENGENHEIRO[token] = usuario
        return usuario

    def sessao_engenheiro(self):
        return bool(self.sessao_usuario())

    def exigir_engenheiro(self):
        if self.sessao_engenheiro():
            return True
        self.enviar_json({"erro": "Acesso restrito ao engenheiro"}, 401)
        return False

    def exigir_administracao(self):
        usuario = self.sessao_usuario()
        if usuario and usuario.get("perfil") in PERFIS_ADMINISTRACAO and "administrar_acessos" in usuario.get("permissoes", []):
            return usuario
        self.enviar_json({"erro": "Acesso restrito à administração"}, 403 if usuario else 401)
        return None

    def exigir_permissao(self, permissao):
        usuario = self.sessao_usuario()
        if usuario and permissao in usuario.get("permissoes", []):
            return usuario
        self.enviar_json({"erro": "Seu perfil não possui permissão para esta ação"}, 403 if usuario else 401)
        return None

    def do_GET(self):
        url = urlparse(self.path)
        caminho = url.path
        parametros = parse_qs(url.query)
        if caminho == "/api/sessao":
            usuario = self.sessao_usuario()
            self.enviar_json({
                "autenticado": bool(usuario),
                "usuario": usuario.get("nome", "") if usuario else "",
                "perfil": usuario.get("perfil", "") if usuario else "",
                "perfilNome": PERFIS_ACESSO.get(usuario.get("perfil"), "") if usuario else "",
                "permissoes": usuario.get("permissoes", []) if usuario else [],
                "podeAdministrar": bool(usuario and usuario.get("perfil") in PERFIS_ADMINISTRACAO and "administrar_acessos" in usuario.get("permissoes", [])),
            })
            return
        if caminho == "/api/cargos-publicos":
            with conectar() as conexao:
                linhas = conexao.execute(
                    "SELECT cargo, perfil FROM cargos_perfis WHERE ativo=1 ORDER BY cargo COLLATE NOCASE"
                ).fetchall()
            self.enviar_json([
                {"cargo": linha["cargo"], "perfil": linha["perfil"], "perfilNome": PERFIS_ACESSO.get(linha["perfil"], linha["perfil"])}
                for linha in linhas
            ])
            return
        if caminho == "/api/admin/usuarios":
            if not self.exigir_administracao():
                return
            with conectar() as conexao:
                linhas = conexao.execute(
                    "SELECT id, nome, email, cargo, perfil, torres, ativo, criado_em, atualizado_em, ultimo_acesso FROM usuarios_acesso ORDER BY nome COLLATE NOCASE"
                ).fetchall()
            self.enviar_json({"perfis": PERFIS_ACESSO, "usuarios": [dict(linha) for linha in linhas]})
            return
        if caminho == "/api/admin/cargos":
            if not self.exigir_administracao():
                return
            with conectar() as conexao:
                linhas = conexao.execute(
                    "SELECT id, cargo, perfil, ativo, criado_em, atualizado_em FROM cargos_perfis ORDER BY cargo COLLATE NOCASE"
                ).fetchall()
            self.enviar_json({"perfis": PERFIS_ACESSO, "cargos": [dict(linha) for linha in linhas]})
            return
        if caminho == "/api/admin/perfis":
            if not self.exigir_administracao():
                return
            with conectar() as conexao:
                perfis = {
                    perfil: permissoes_do_perfil(conexao, perfil)
                    for perfil in PERFIS_ACESSO
                }
            self.enviar_json({"perfis": PERFIS_ACESSO, "acoes": ACOES_PERFIL, "permissoes": perfis})
            return
        if caminho == "/api/admin/historico":
            if not self.exigir_administracao():
                return
            with conectar() as conexao:
                linhas = conexao.execute(
                    "SELECT id, usuario_nome, acao, detalhes, criado_em FROM auditoria_acessos ORDER BY id DESC LIMIT 500"
                ).fetchall()
            self.enviar_json([dict(linha) for linha in linhas])
            return
        if caminho == "/visitante-inicio":
            corpo = gerar_inicio_visitante()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(corpo)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(corpo)
            return
        if caminho in {"/visitante-relatorios", "/relatorio-visitante.pdf", "/api/qrcode-relatorios"}:
            torre = parametros.get("torre", [""])[0]
            andar = parametros.get("andar", ["todos"])[0]
            unidade = parametros.get("unidade", ["todos"])[0]
            atividade = parametros.get("atividade", ["todos"])[0]
            status = parametros.get("status", ["todos"])[0]
            acabamento = parametros.get("acabamento", ["todos"])[0]
            ocorrencias = parametros.get("ocorrencias", ["todas"])[0]
            if torre not in TORRES_NOMES or (andar != "todos" and not andar.isdigit()):
                self.enviar_json({"erro": "Filtros inválidos"}, 400)
                return
            if status not in {*STATUS_NOMES, "todos"}:
                self.enviar_json({"erro": "Situação inválida"}, 400)
                return
            acabamentos_validos = {"todos", "padrao", "alterado", "personalizada", "nao-instalar", *(f"opcao-{numero}" for numero in range(2, 8))}
            if acabamento not in acabamentos_validos or ocorrencias not in {"todas", "com", "sem", "pendente", "concluido"}:
                self.enviar_json({"erro": "Filtro de acabamento ou ocorrência inválido"}, 400)
                return
            if caminho == "/visitante-relatorios":
                corpo = gerar_portal_relatorios(torre, andar, unidade, atividade)
                tipo = "text/html; charset=utf-8"
            elif caminho == "/relatorio-visitante.pdf":
                corpo = gerar_pdf_visitante_filtros(torre, andar, unidade, atividade, status, acabamento, ocorrencias)
                tipo = "application/pdf"
            else:
                corpo = gerar_qr_svg(self.url_publica(
                    "/visitante-relatorios", urlencode({"torre": torre})
                ))
                if isinstance(corpo, str):
                    corpo = corpo.encode("utf-8")
                tipo = "image/svg+xml; charset=utf-8"
            self.send_response(200)
            self.send_header("Content-Type", tipo)
            self.send_header("Content-Length", str(len(corpo)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(corpo)
            return
        if caminho == "/relatorio-ocorrencias.pdf":
            torre = parametros.get("torre", [""])[0]
            andar = parametros.get("andar", ["todos"])[0]
            unidade = parametros.get("unidade", ["todos"])[0]
            status = parametros.get("status", ["todos"])[0]
            tipo = parametros.get("tipo", ["todos"])[0]
            try:
                converter_filtro_andar(andar)
            except ValueError:
                self.enviar_json({"erro": "Filtros inválidos"}, 400)
                return
            if torre not in {*TORRES_NOMES, "todos"} or status not in {*STATUS_NOMES.keys(), "todos"} or tipo not in {"todos", "atividade", "seguranca"}:
                self.enviar_json({"erro": "Filtros inválidos"}, 400)
                return
            corpo = gerar_historico_ocorrencias_pdf(torre, andar, unidade, status, tipo)
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", 'attachment; filename="historico-ocorrencias.pdf"')
            self.send_header("Content-Length", str(len(corpo)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(corpo)
            return
        if caminho == "/placas-unidades.pdf":
            if not self.exigir_permissao("gerar_pdf"):
                return
            torre = parametros.get("torre", [""])[0]
            andar = parametros.get("andar", ["todos"])[0]
            unidade = parametros.get("unidade", ["todos"])[0]
            if torre not in TORRES_NOMES or (andar != "todos" and not andar.isdigit()):
                self.enviar_json({"erro": "Filtros inválidos para as placas"}, 400)
                return
            unidades = unidades_para_placas(torre, andar, unidade)
            if not unidades:
                self.enviar_json({"erro": "Nenhuma unidade encontrada para as placas"}, 404)
                return
            escopo = unidade if unidade != "todos" else f"andar-{andar}" if andar != "todos" else TORRES_NOMES[torre]
            base_publica = self.url_publica("").rstrip("/")
            if unidade == "todos":
                corpo, nome_arquivo = gerar_pasta_placas_zip(torre, unidades, base_publica, escopo)
                tipo_resposta = "application/zip"
            else:
                corpo = gerar_placas_pdf(torre, unidades, base_publica)
                nome = re.sub(r"[^a-z0-9]+", "-", unicodedata.normalize("NFD", escopo).encode("ascii", "ignore").decode().lower()).strip("-")
                nome_arquivo = f"placa-{nome}.pdf"
                tipo_resposta = "application/pdf"
            self.send_response(200)
            self.send_header("Content-Type", tipo_resposta)
            self.send_header("Content-Disposition", f'attachment; filename="{nome_arquivo}"')
            self.send_header("Content-Length", str(len(corpo)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(corpo)
            return
        if caminho in {"/relatorio.pdf", "/api/qrcode", "/visitante"}:
            torre = parametros.get("torre", [""])[0]
            unidade = parametros.get("unidade", [""])[0]
            try:
                andar = int(parametros.get("andar", ["0"])[0])
            except ValueError:
                andar = 0
            if torre not in TORRES_NOMES or andar < -2 or not unidade:
                self.enviar_json({"erro": "Unidade inválida"}, 400)
                return
            if caminho == "/visitante":
                assinatura = parametros.get("acesso", [""])[0]
                if not acesso_visitante_valido(torre, andar, unidade, assinatura):
                    corpo = b"<!doctype html><html lang='pt-BR'><meta charset='utf-8'><title>Acesso inv\xc3\xa1lido</title><body style='font-family:Arial;padding:40px'><h1>Acesso inv\xc3\xa1lido</h1><p>Use o QR Code impresso na placa desta unidade.</p></body></html>"
                    self.send_response(403)
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                    self.send_header("Content-Length", str(len(corpo)))
                    self.end_headers()
                    self.wfile.write(corpo)
                    return
                corpo = gerar_pagina_visitante(torre, andar, unidade, assinatura)
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(corpo)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(corpo)
                return
            if caminho == "/relatorio.pdf":
                assinatura = parametros.get("acesso", [""])[0]
                usuario = self.sessao_usuario()
                pode_gerar = bool(usuario and "gerar_pdf" in usuario.get("permissoes", []))
                if not pode_gerar and not acesso_visitante_valido(torre, andar, unidade, assinatura):
                    self.enviar_json({"erro": "Acesso restrito a esta unidade"}, 403)
                    return
                corpo = gerar_relatorio_pdf(
                    torre, andar, unidade,
                    incluir_ocorrencias_pendentes=pode_gerar,
                )
                nome = unidade.replace("Apto ", "apartamento-").replace(" ", "-").lower()
                self.send_response(200)
                self.send_header("Content-Type", "application/pdf")
                self.send_header("Content-Disposition", f'attachment; filename="relatorio-{nome}.pdf"')
                self.send_header("Content-Length", str(len(corpo)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(corpo)
                return
            if not self.exigir_permissao("gerar_pdf"):
                return
            consulta = urlencode({"torre": torre, "andar": andar, "unidade": unidade, "acesso": assinatura_visitante(torre, andar, unidade)})
            corpo = gerar_qr_svg(self.url_publica("/visitante", consulta))
            if isinstance(corpo, str):
                corpo = corpo.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "image/svg+xml; charset=utf-8")
            self.send_header("Content-Length", str(len(corpo)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(corpo)
            return
        if caminho == "/planta-unidade.pdf":
            torre = parametros.get("torre", [""])[0]
            unidade = parametros.get("unidade", [""])[0]
            tipo = parametros.get("tipo", ["planta"])[0]
            try:
                indice = int(parametros.get("indice", ["0"])[0])
                projeto_id = int(parametros.get("projeto", ["0"])[0])
                if torre not in TORRES_NOMES or not unidade or tipo not in {"planta", "personalizacao"}:
                    raise ValueError("Planta inválida")
                corpo = gerar_pdf_planta_unidade(torre, unidade, tipo, indice, projeto_id)
                nome = re.sub(r"[^a-z0-9-]+", "-", unicodedata.normalize("NFD", unidade).encode("ascii", "ignore").decode("ascii").lower()).strip("-")
                self.send_response(200)
                self.send_header("Content-Type", "application/pdf")
                self.send_header("Content-Disposition", f'attachment; filename="planta-{nome}.pdf"')
                self.send_header("Content-Length", str(len(corpo)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(corpo)
            except (ValueError, OSError, json.JSONDecodeError, binascii.Error) as erro:
                self.enviar_json({"erro": str(erro)}, 404)
            return
        if caminho == "/api/registros":
            parametros = parse_qs(urlparse(self.path).query)
            formato_compacto = parametros.get("formato", [""])[0] == "compacto"
            if formato_compacto:
                def compactar_linhas(linhas):
                    registros_compactos = []
                    for linha in linhas:
                        try:
                            especificacoes = json.loads(linha["especificacoes"] or "{}")
                        except (json.JSONDecodeError, TypeError):
                            especificacoes = {}
                        item = [
                            linha["chave"], linha["status"] or ("concluido" if linha["concluido"] else "nao-iniciado"),
                            linha["data_conclusao"] or "", linha["observacao"] or "", linha["foto"] or "",
                            linha["foto_nome"] or "", especificacoes if isinstance(especificacoes, dict) else {},
                        ]
                        while len(item) > 2 and item[-1] in ("", {}, None):
                            item.pop()
                        registros_compactos.append(item)
                    return registros_compactos
                desde = parametros.get("desde", [""])[0].strip()
                if desde:
                    with conectar() as conexao:
                        linhas = conexao.execute(
                            "SELECT chave, concluido, status, data_conclusao, observacao, foto, foto_nome, especificacoes FROM registros WHERE atualizado_em >= datetime(?, '-2 seconds')",
                            (desde,),
                        ).fetchall()
                    self.enviar_json({
                        "formato": "compacto-v1", "parcial": True,
                        "sincronizadoEm": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
                        "registros": compactar_linhas(linhas),
                    })
                    return
                caminhos_banco = [BANCO, Path(f"{BANCO}-wal")]
                def assinatura_banco():
                    return tuple(
                        (arquivo.stat().st_mtime_ns, arquivo.stat().st_size) if arquivo.exists() else (0, 0)
                        for arquivo in caminhos_banco
                    )
                assinatura = assinatura_banco()
                with LOCK_CACHE_REGISTROS:
                    if CACHE_REGISTROS_COMPACTOS["assinatura"] != assinatura:
                        with conectar() as conexao:
                            linhas = conexao.execute(
                                "SELECT chave, concluido, status, data_conclusao, observacao, foto, foto_nome, especificacoes FROM registros"
                            ).fetchall()
                        assinatura = assinatura_banco()
                        corpo = json.dumps({
                            "formato": "compacto-v1",
                            "sincronizadoEm": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
                            "registros": compactar_linhas(linhas),
                        }, ensure_ascii=False).encode("utf-8")
                        CACHE_REGISTROS_COMPACTOS.update({
                            "assinatura": assinatura, "corpo": corpo, "gzip": gzip.compress(corpo, compresslevel=5),
                            "etag": f'"{hashlib.sha256(corpo).hexdigest()[:24]}"',
                        })
                    cache = dict(CACHE_REGISTROS_COMPACTOS)
                self.enviar_json_cacheado(cache["corpo"], cache["gzip"], cache["etag"])
                return
            with conectar() as conexao:
                linhas = conexao.execute(
                    "SELECT chave, concluido, status, data_conclusao, observacao, foto, foto_nome, especificacoes FROM registros"
                ).fetchall()
            registros = {}
            for linha in linhas:
                try:
                    especificacoes = json.loads(linha["especificacoes"] or "{}")
                except (json.JSONDecodeError, TypeError):
                    especificacoes = {}
                registros[linha["chave"]] = {
                    "concluido": bool(linha["concluido"]),
                    "status": linha["status"],
                    "data": linha["data_conclusao"],
                    "observacao": linha["observacao"],
                    "foto": linha["foto"],
                    "fotoNome": linha["foto_nome"],
                    "especificacoes": especificacoes if isinstance(especificacoes, dict) else {},
                }
            self.enviar_json(registros)
            return
        if caminho == "/api/planejamento":
            if not ARQUIVO_PLANEJAMENTO.exists():
                self.enviar_json({"erro": "Planejamento não encontrado"}, 404)
                return
            self.enviar_json(json.loads(ARQUIVO_PLANEJAMENTO.read_text(encoding="utf-8")))
            return
        if caminho == "/api/atividades":
            self.enviar_json(listar_atividades_config())
            return
        if caminho == "/api/opcoes-plantas":
            try:
                torre = parametros.get("torre", [""])[0]
                andar = int(parametros.get("andar", ["0"])[0])
                unidade = parametros.get("unidade", [""])[0]
                if torre not in TORRES_NOMES or not unidade:
                    raise ValueError("Unidade inválida")
                self.enviar_json(opcoes_plantas_pdf(torre, andar, unidade))
            except ValueError as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            return
        if caminho == "/api/ocorrencias":
            with conectar() as conexao:
                linhas = conexao.execute(
                    "SELECT * FROM ocorrencias ORDER BY data_ocorrencia DESC, criado_em DESC, id DESC"
                ).fetchall()
            self.enviar_json(
                [
                    {
                        "id": linha["id"],
                        "torre": linha["torre"],
                        "andar": linha["andar"],
                        "unidade": linha["unidade"],
                        "descricao": linha["descricao"],
                        "foto": linha["foto"],
                        "fotoNome": linha["foto_nome"],
                        "dataOcorrencia": linha["data_ocorrencia"],
                        "status": linha["status"],
                        "atividade": linha["atividade"],
                        "subatividade": linha["subatividade"],
                        "especificacao": linha["especificacao"],
                        "criadoEm": linha["criado_em"],
                    }
                    for linha in linhas
                ]
            )
            return
        if caminho == "/api/setores/importar":
            if not self.exigir_engenheiro():
                return
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                linhas = dados.get("linhas", [])
                if not isinstance(linhas, list) or not linhas or len(linhas) > 5000:
                    raise ValueError("A planilha deve conter entre 1 e 5.000 linhas")
                mapa_torres = {
                    "home": "aurora", "torre home": "aurora", "aurora": "aurora",
                    "smart": "horizonte", "torre smart": "horizonte", "horizonte": "horizonte",
                }
                mapa_pavimentos = {
                    "fundacao": -2, "fundação": -2, "1 subsolo": -1, "1º subsolo": -1,
                    "subsolo": -1, "terreo": 0, "térreo": 0, "barrilete": 101,
                    "reservatorio": 102, "reservatório": 102, "cobertura": 103,
                }
                criados = 0
                ignorados = 0
                erros = []
                with conectar() as conexao:
                    for indice, linha in enumerate(linhas, start=2):
                        try:
                            torre_texto = str(linha.get("torre", "")).strip().lower()
                            torre = mapa_torres.get(torre_texto)
                            pavimento_texto = str(linha.get("pavimento", "")).strip().lower()
                            pavimento_normalizado = pavimento_texto.replace("º andar", "").replace(" andar", "").strip()
                            andar = mapa_pavimentos.get(pavimento_texto)
                            if andar is None:
                                andar = int(pavimento_normalizado)
                            setor = str(linha.get("setor", "")).strip()
                            atividade_texto = str(linha.get("atividade", "")).strip()
                            subservico = str(linha.get("subservico", "")).strip()
                            validar_escopos([{"torre": torre or "", "andar": andar, "unidade": setor}])
                            if not setor:
                                raise ValueError("setor não informado")
                            if atividade_texto.lower() in {"todas", "todos", "*"}:
                                atividades = [
                                    item["nome"] for item in conexao.execute(
                                        "SELECT DISTINCT nome FROM atividades_config WHERE torre=? AND andar=? ORDER BY nome",
                                        (torre, andar),
                                    ).fetchall()
                                ]
                            else:
                                atividades = [atividade_texto]
                            if not atividades:
                                raise ValueError("nenhuma atividade disponível no pavimento")
                            for atividade in atividades:
                                existe = conexao.execute(
                                    "SELECT 1 FROM atividades_config WHERE nome=? AND torre=? AND andar=? LIMIT 1",
                                    (atividade, torre, andar),
                                ).fetchone()
                                if not existe:
                                    raise ValueError(f"atividade '{atividade}' não disponível no pavimento")
                                conflito = conexao.execute(
                                    "SELECT 1 FROM registros WHERE torre=? AND andar=? AND unidade=? AND atividade=? LIMIT 1",
                                    (torre, andar, setor, atividade),
                                ).fetchone()
                                if conflito:
                                    ignorados += 1
                                    continue
                                conexao.execute(
                                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, ?)",
                                    (atividade, torre, andar, setor),
                                )
                                garantir_registros_atividade(conexao, atividade, [(torre, andar, setor)])
                                criados += 1
                        except (ValueError, TypeError) as erro:
                            erros.append(f"Linha {indice}: {erro}")
                self.enviar_json({"ok": True, "criados": criados, "ignorados": ignorados, "erros": erros[:30]}, 201)
            except (ValueError, json.JSONDecodeError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            return
        if caminho == "/api/projetos":
            with conectar() as conexao:
                linhas = conexao.execute(
                    "SELECT * FROM projetos_unidade ORDER BY criado_em, id"
                ).fetchall()
            self.enviar_json([
                {"id": linha["id"], "torre": linha["torre"], "andar": linha["andar"],
                 "unidade": linha["unidade"], "titulo": linha["titulo"],
                 "imagem": linha["imagem"], "arquivoNome": linha["arquivo_nome"]}
                for linha in linhas
            ])
            return
        if caminho == "/api/comentarios-unidade":
            with conectar() as conexao:
                linhas = conexao.execute(
                    "SELECT id, torre, andar, unidade, comentario, autor, criado_em FROM comentarios_unidade ORDER BY criado_em, id"
                ).fetchall()
            self.enviar_json([dict(linha) for linha in linhas])
            return
        if caminho == "/api/projetos-ocultos":
            with conectar() as conexao:
                linhas = conexao.execute("SELECT torre, andar, unidade, imagem FROM projetos_ocultos").fetchall()
            self.enviar_json([dict(linha) for linha in linhas])
            return
        if caminho == "/":
            self.path = "/acompanhamento_obra.html"
        super().do_GET()

    def do_PUT(self):
        caminho = urlparse(self.path).path
        partes = caminho.strip("/").split("/")
        permissao_rota = {
            "/api/setores": "gerenciar_setores",
            "/api/atividades": "gerenciar_atividades",
            "/api/registros": "atualizar_acompanhamento",
        }.get(caminho)
        if permissao_rota and not self.exigir_permissao(permissao_rota):
            return
        if len(partes) == 4 and partes[:3] == ["api", "admin", "perfis"]:
            administrador = self.exigir_administracao()
            if not administrador:
                return
            try:
                perfil = partes[3]
                if perfil not in PERFIS_ACESSO:
                    raise ValueError("Perfil inválido")
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                permissoes = list(dict.fromkeys(acao for acao in dados.get("permissoes", []) if acao in ACOES_PERFIL))
                if "visualizar" not in permissoes:
                    permissoes.insert(0, "visualizar")
                if perfil in PERFIS_ADMINISTRACAO and "administrar_acessos" not in permissoes:
                    permissoes.append("administrar_acessos")
                with conectar() as conexao:
                    conexao.execute(
                        "INSERT INTO perfis_permissoes (perfil,permissoes,atualizado_em) VALUES (?,?,CURRENT_TIMESTAMP) ON CONFLICT(perfil) DO UPDATE SET permissoes=excluded.permissoes,atualizado_em=CURRENT_TIMESTAMP",
                        (perfil, json.dumps(permissoes)),
                    )
                    nomes = ", ".join(ACOES_PERFIL[item] for item in permissoes)
                    registrar_auditoria(conexao, administrador["id"], administrador["nome"], "Permissões de perfil alteradas", f"{PERFIS_ACESSO[perfil]} · {nomes}")
                self.enviar_json({"ok": True, "perfil": perfil, "permissoes": permissoes})
            except (ValueError, json.JSONDecodeError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            return
        if len(partes) == 4 and partes[:3] == ["api", "admin", "cargos"]:
            administrador = self.exigir_administracao()
            if not administrador:
                return
            try:
                cargo_id = int(partes[3])
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                cargo = str(dados.get("cargo", "")).strip()
                perfil = str(dados.get("perfil", "")).strip()
                ativo = 1 if dados.get("ativo", True) else 0
                if not cargo or perfil not in PERFIS_ACESSO:
                    raise ValueError("Informe o cargo e um perfil válido")
                with conectar() as conexao:
                    anterior = conexao.execute("SELECT cargo FROM cargos_perfis WHERE id=?", (cargo_id,)).fetchone()
                    if not anterior:
                        raise ValueError("Cargo não encontrado")
                    conexao.execute(
                        "UPDATE cargos_perfis SET cargo=?,perfil=?,ativo=?,atualizado_em=CURRENT_TIMESTAMP WHERE id=?",
                        (cargo, perfil, ativo, cargo_id),
                    )
                    atualizados = conexao.execute(
                        "UPDATE usuarios_acesso SET cargo=?,perfil=?,atualizado_em=CURRENT_TIMESTAMP WHERE cargo=? COLLATE NOCASE",
                        (cargo, perfil, anterior["cargo"]),
                    ).rowcount
                    registrar_auditoria(conexao, administrador["id"], administrador["nome"], "Cargo e perfil alterados", f"{cargo} · {PERFIS_ACESSO[perfil]} · {atualizados} usuário(s) atualizado(s)")
                self.enviar_json({"ok": True, "id": cargo_id, "usuariosAtualizados": atualizados})
            except (ValueError, json.JSONDecodeError, sqlite3.IntegrityError) as erro:
                mensagem = "Este cargo já está cadastrado" if isinstance(erro, sqlite3.IntegrityError) else str(erro)
                self.enviar_json({"erro": mensagem}, 400)
            return
        if len(partes) == 4 and partes[:3] == ["api", "admin", "usuarios"]:
            administrador = self.exigir_administracao()
            if not administrador:
                return
            try:
                usuario_id = int(partes[3])
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                nome = str(dados.get("nome", "")).strip()
                email = str(dados.get("email", "")).strip().lower()
                cargo = str(dados.get("cargo", "")).strip()
                senha = str(dados.get("senha", ""))
                ativo = 1 if dados.get("ativo", True) else 0
                torres = [torre for torre in dados.get("torres", []) if torre in TORRES_NOMES]
                if not nome or "@" not in email or not cargo or not torres:
                    raise ValueError("Informe nome, e-mail, cargo e ao menos uma torre")
                if senha and len(senha) < 8:
                    raise ValueError("A nova senha deve ter ao menos 8 caracteres")
                with conectar() as conexao:
                    cargo_config = conexao.execute(
                        "SELECT cargo, perfil FROM cargos_perfis WHERE cargo=? COLLATE NOCASE AND ativo=1", (cargo,)
                    ).fetchone()
                    if not cargo_config:
                        raise ValueError("Selecione um cargo ativo cadastrado")
                    cargo, perfil = cargo_config["cargo"], cargo_config["perfil"]
                    if usuario_id == administrador["id"] and (not ativo or perfil not in PERFIS_ADMINISTRACAO):
                        raise ValueError("Você não pode retirar o próprio acesso administrativo")
                    anterior = conexao.execute("SELECT * FROM usuarios_acesso WHERE id=?", (usuario_id,)).fetchone()
                    if not anterior:
                        raise ValueError("Usuário não encontrado")
                    campos = [nome, email, cargo, perfil, json.dumps(torres), ativo]
                    consulta = "UPDATE usuarios_acesso SET nome=?,email=?,cargo=?,perfil=?,torres=?,ativo=?,atualizado_em=CURRENT_TIMESTAMP"
                    if senha:
                        consulta += ",senha_hash=?"
                        campos.append(gerar_hash_senha(senha))
                    consulta += " WHERE id=?"
                    campos.append(usuario_id)
                    conexao.execute(consulta, campos)
                    registrar_auditoria(conexao, administrador["id"], administrador["nome"], "Usuário alterado", f"{nome} · {email} · {PERFIS_ACESSO[perfil]} · {'ativo' if ativo else 'bloqueado'}")
                self.enviar_json({"ok": True, "id": usuario_id})
            except (ValueError, json.JSONDecodeError, sqlite3.IntegrityError) as erro:
                mensagem = "Este e-mail já está cadastrado" if isinstance(erro, sqlite3.IntegrityError) else str(erro)
                self.enviar_json({"erro": mensagem}, 400)
            return
        if caminho == "/api/setores":
            if not self.exigir_engenheiro():
                return
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                torre = str(dados.get("torre", ""))
                andar = int(dados.get("andar", 0))
                atividade = str(dados.get("atividade", "")).strip()
                unidade = str(dados.get("unidade", "")).strip()
                nova_unidade = str(dados.get("novaUnidade", "")).strip()
                subservico = str(dados.get("subservico", "")).strip()
                if torre not in TORRES_NOMES or not unidade or not nova_unidade or andar < -2:
                    raise ValueError("Setor inválido")
                with conectar() as conexao:
                    if subservico:
                        if not atividade:
                            raise ValueError("Informe a atividade do subserviço")
                        unidades = setores_efetivos_subservico(conexao, atividade, subservico, torre, andar)
                        if unidade not in unidades:
                            raise ValueError("Setor não encontrado neste subserviço")
                        if nova_unidade in unidades:
                            raise ValueError("Já existe um setor com esse nome neste subserviço")
                        unidades = [nova_unidade if item == unidade else item for item in unidades]
                        salvar_setores_subservico(conexao, atividade, subservico, torre, andar, unidades)
                        garantir_registros_atividade(conexao, atividade, [(torre, andar, nova_unidade)])
                        conexao.execute("UPDATE ocorrencias SET unidade=? WHERE torre=? AND andar=? AND unidade=? AND atividade=? AND subatividade=?", (nova_unidade, torre, andar, unidade, atividade, subservico))
                        self.enviar_json({"ok": True, "unidade": nova_unidade})
                        return
                    conflito = conexao.execute("SELECT 1 FROM registros WHERE torre=? AND andar=? AND unidade=? LIMIT 1", (torre, andar, nova_unidade)).fetchone()
                    if conflito:
                        raise ValueError("Já existe um setor com esse nome neste andar")
                    linhas = conexao.execute("SELECT chave, atividade FROM registros WHERE torre=? AND andar=? AND unidade=?", (torre, andar, unidade)).fetchall()
                    if not linhas:
                        raise ValueError("Setor não encontrado")
                    atividades_afetadas = list(dict.fromkeys(linha["atividade"] for linha in linhas))
                    for atividade in atividades_afetadas:
                        possui_escopo_geral = conexao.execute(
                            "SELECT 1 FROM atividades_config WHERE torre=? AND andar=? AND nome=? AND unidade='*' LIMIT 1",
                            (torre, andar, atividade),
                        ).fetchone()
                        if not possui_escopo_geral:
                            continue
                        destinos_automaticos = apartamentos_do_andar(torre, andar) + ["Área comum"]
                        destinos_explicitos = [
                            linha_escopo["unidade"] for linha_escopo in conexao.execute(
                                "SELECT DISTINCT unidade FROM atividades_config WHERE torre=? AND andar=? AND nome=? AND unidade!='*'",
                                (torre, andar, atividade),
                            ).fetchall()
                        ]
                        destinos = list(dict.fromkeys(
                            nova_unidade if destino == unidade else destino
                            for destino in destinos_automaticos + destinos_explicitos
                        ))
                        conexao.execute(
                            "DELETE FROM atividades_config WHERE torre=? AND andar=? AND nome=?",
                            (torre, andar, atividade),
                        )
                        conexao.executemany(
                            "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, ?)",
                            [(atividade, torre, andar, destino) for destino in destinos],
                        )
                    for linha in linhas:
                        nova_chave = f"{torre}|{andar}|{nova_unidade}|{linha['atividade']}"
                        conexao.execute("UPDATE registros SET chave=?, unidade=? WHERE chave=?", (nova_chave, nova_unidade, linha["chave"]))
                    for tabela in ("ocorrencias", "projetos_unidade", "projetos_ocultos", "atividades_config"):
                        conexao.execute(f"UPDATE {tabela} SET unidade=? WHERE torre=? AND andar=? AND unidade=?", (nova_unidade, torre, andar, unidade))
                self.enviar_json({"ok": True, "unidade": nova_unidade})
            except (ValueError, json.JSONDecodeError, sqlite3.IntegrityError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            return
        if caminho == "/api/atividades":
            if not self.exigir_engenheiro():
                return
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                nome_original = str(dados.get("nomeOriginal", "")).strip()
                nome = str(dados.get("nome", "")).strip()
                if not nome_original or not nome:
                    raise ValueError("Informe o nome da atividade")
                escopos = validar_escopos(dados.get("escopos", []))
                especificacoes = validar_especificacoes(dados.get("especificacoes", []))
                with conectar() as conexao:
                    conexao.execute("DELETE FROM atividades_config WHERE nome = ?", (nome_original,))
                    conexao.execute("DELETE FROM atividades_especificacoes WHERE atividade = ?", (nome_original,))
                    for torre, andar, unidade in escopos:
                        conexao.execute(
                            "INSERT INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, ?)",
                            (nome, torre, andar, unidade),
                        )
                    salvar_especificacoes_atividade(conexao, nome, especificacoes)
                    garantir_registros_atividade(conexao, nome, escopos)
                self.enviar_json({"ok": True, "nome": nome})
            except (ValueError, json.JSONDecodeError, sqlite3.IntegrityError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            except Exception as erro:
                self.enviar_json({"erro": f"Falha ao atualizar atividade: {erro}"}, 500)
            return
        if caminho != "/api/registros":
            self.enviar_json({"erro": "Rota não encontrada"}, 404)
            return
        if not self.exigir_engenheiro():
            return
        try:
            tamanho = int(self.headers.get("Content-Length", "0"))
            if tamanho > 80 * 1024 * 1024:
                self.enviar_json({"erro": "Dados acima do limite permitido"}, 413)
                return
            registros = json.loads(self.rfile.read(tamanho).decode("utf-8"))
            if not isinstance(registros, dict):
                raise ValueError("Formato inválido")
            with conectar() as conexao:
                for chave, registro in registros.items():
                    partes = chave.split("|", 3)
                    if len(partes) != 4:
                        continue
                    torre, andar, unidade, atividade = partes
                    conexao.execute(
                        """
                        INSERT INTO registros
                        (chave, torre, andar, unidade, atividade, concluido,
                         data_conclusao, observacao, foto, foto_nome, status, especificacoes, atualizado_em)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                        ON CONFLICT(chave) DO UPDATE SET
                            concluido=excluded.concluido,
                            data_conclusao=excluded.data_conclusao,
                            observacao=excluded.observacao,
                            foto=excluded.foto,
                            foto_nome=excluded.foto_nome,
                            status=excluded.status,
                            especificacoes=excluded.especificacoes,
                            atualizado_em=CURRENT_TIMESTAMP
                        """,
                        (
                            chave,
                            torre,
                            int(andar),
                            unidade,
                            atividade,
                            1 if registro.get("concluido") else 0,
                            registro.get("data", ""),
                            registro.get("observacao", ""),
                            registro.get("foto", ""),
                            registro.get("fotoNome", ""),
                            registro.get(
                                "status",
                                "concluido" if registro.get("concluido") else "nao-iniciado",
                            ),
                            json.dumps(
                                registro.get("especificacoes", {}),
                                ensure_ascii=False,
                            ),
                        ),
                    )
            self.enviar_json({"ok": True, "registros": len(registros)})
        except (ValueError, json.JSONDecodeError) as erro:
            self.enviar_json({"erro": str(erro)}, 400)
        except Exception as erro:
            self.enviar_json({"erro": f"Falha ao salvar: {erro}"}, 500)

    def do_POST(self):
        caminho = urlparse(self.path).path
        permissao_rota = {
            "/api/setores": "gerenciar_setores",
            "/api/setores/importar": "gerenciar_setores",
            "/api/personalizacoes/acabamentos": "gerenciar_projetos",
            "/api/personalizacoes/planta": "gerenciar_projetos",
            "/api/personalizacoes/importar-planilhas": "gerenciar_projetos",
            "/api/atividades": "gerenciar_atividades",
            "/api/projetos": "gerenciar_projetos",
            "/api/comentarios-unidade": "gerenciar_projetos",
            "/api/projetos-ocultos": "gerenciar_projetos",
            "/api/ocorrencias": "gerenciar_ocorrencias",
        }.get(caminho)
        if permissao_rota and not self.exigir_permissao(permissao_rota):
            return
        if caminho == "/api/personalizacoes/importar-planilhas":
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                arquivos = {}
                for campo, padrao in (("plantas", "RELACAO_OPCOES_DE_PLANTA_IMPORTADA.xlsx"), ("acabamentos", "OPCOES_DE_ACABAMENTOS_IMPORTADA.xlsx")):
                    item = dados.get(campo) or {}
                    conteudo = str(item.get("conteudo", ""))
                    if not conteudo:
                        continue
                    nome = Path(str(item.get("nome") or padrao)).name
                    if not nome.lower().endswith(".xlsx"):
                        raise ValueError(f"O arquivo {nome} deve estar no formato XLSX.")
                    destino = PASTA / nome
                    destino.write_bytes(base64.b64decode(conteudo.split(",", 1)[-1], validate=True))
                    arquivos[campo] = destino
                if not arquivos:
                    raise ValueError("Selecione ao menos uma planilha para importar.")
                plantas_antes = json.loads(ARQUIVO_PLANTAS.read_text(encoding="utf-8")) if ARQUIVO_PLANTAS.exists() else {}
                acabamentos_antes = json.loads(ARQUIVO_ACABAMENTOS.read_text(encoding="utf-8")) if ARQUIVO_ACABAMENTOS.exists() else {}
                ambiente = os.environ.copy()
                if "plantas" in arquivos:
                    ambiente["PLANILHA_PLANTAS"] = str(arquivos["plantas"])
                if "acabamentos" in arquivos:
                    ambiente["PLANILHA_ACABAMENTOS"] = str(arquivos["acabamentos"])
                scripts = (["gerar_plantas_tipos.py"] if "plantas" in arquivos else []) + (["gerar_acabamentos_unidades.py"] if "acabamentos" in arquivos else [])
                for script in scripts:
                    processo = subprocess.run(["python3", str(PASTA / script)], cwd=PASTA, env=ambiente, capture_output=True, text=True)
                    if processo.returncode:
                        raise RuntimeError(processo.stderr.strip() or f"Falha ao processar {script}")
                plantas_depois = json.loads(ARQUIVO_PLANTAS.read_text(encoding="utf-8"))
                acabamentos_depois = json.loads(ARQUIVO_ACABAMENTOS.read_text(encoding="utf-8"))
                alteracoes = []
                agora = datetime.now().astimezone().isoformat()
                for chave in (sorted(set(plantas_antes) | set(plantas_depois)) if "plantas" in arquivos else []):
                    if plantas_antes.get(chave) == plantas_depois.get(chave):
                        continue
                    torre, unidade = chave.split("|", 1)
                    anterior = plantas_antes.get(chave, {}).get("tipo", "Não cadastrada")
                    atual = plantas_depois.get(chave, {}).get("tipo", "Removida")
                    alteracoes.append({"chave": f"importacao-planta-{int(datetime.now().timestamp())}-{chave}", "aba": "personalizacao", "texto": f"Planta atualizada · {unidade.replace('Apto ', 'Apartamento ')} · {TORRES_NOMES.get(torre, torre)} · {anterior} → {atual}", "data": agora, "alvo": {"torre": torre, "unidade": unidade}})
                for chave in (sorted(set(acabamentos_antes) | set(acabamentos_depois)) if "acabamentos" in arquivos else []):
                    if acabamentos_antes.get(chave) == acabamentos_depois.get(chave):
                        continue
                    torre, unidade = chave.split("|", 1)
                    alteracoes.append({"chave": f"importacao-acabamento-{int(datetime.now().timestamp())}-{chave}", "aba": "personalizacao", "texto": f"Acabamento atualizado · {unidade.replace('Apto ', 'Apartamento ')} · {TORRES_NOMES.get(torre, torre)}", "data": agora, "alvo": {"torre": torre, "unidade": unidade}})
                self.enviar_json({"ok": True, "alteracoes": alteracoes, "quantidade": len(alteracoes)})
            except (ValueError, json.JSONDecodeError, binascii.Error) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            except Exception as erro:
                self.enviar_json({"erro": f"Falha ao importar planilhas: {erro}"}, 500)
            return
        if caminho == "/api/personalizacoes/acabamentos":
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                torre = str(dados.get("torre", ""))
                unidade = str(dados.get("unidade", "")).strip()
                opcao = str(dados.get("opcao", "padrao"))
                escolhas = {str(ambiente).strip(): str(escolha).strip() for ambiente, escolha in dados.get("escolhas", {}).items() if str(ambiente).strip() and str(escolha).strip()}
                if torre not in TORRES_NOMES or not unidade or opcao not in {"padrao", "personalizado"}:
                    raise ValueError("Personalização de acabamento inválida")
                acabamentos = json.loads(ARQUIVO_ACABAMENTOS.read_text(encoding="utf-8")) if ARQUIVO_ACABAMENTOS.exists() else {}
                chave = f"{torre}|{unidade}"
                acabamento = acabamentos.get(chave, {"planta": "Opção 1", "itens": []})
                acabamento["planta"] = "Opção 1" if opcao == "padrao" else "Opção Personalizada"
                catalogo = {}
                for acabamento_catalogo in acabamentos.values():
                    for item_catalogo in acabamento_catalogo.get("itens", []):
                        chave_catalogo = (item_catalogo.get("ambiente"), item_catalogo.get("item"), item_catalogo.get("opcao"))
                        if item_catalogo.get("descricao"):
                            catalogo[chave_catalogo] = item_catalogo["descricao"]
                for item in acabamento.get("itens", []):
                    chave_item = f'{item.get("ambiente", "")}|||{item.get("item", "")}'
                    escolha = (escolhas.get(chave_item) or escolhas.get(item.get("ambiente"))) if opcao == "personalizado" else "Opção 1"
                    if escolha:
                        item["opcao"] = escolha
                        descricao = catalogo.get((item.get("ambiente"), item.get("item"), escolha))
                        if descricao is not None:
                            item["descricao"] = descricao
                acabamentos[chave] = acabamento
                temporario = ARQUIVO_ACABAMENTOS.with_suffix(".json.tmp")
                temporario.write_text(json.dumps(acabamentos, ensure_ascii=False, indent=2), encoding="utf-8")
                temporario.replace(ARQUIVO_ACABAMENTOS)
                self.enviar_json({"ok": True, "acabamento": acabamento}, 201)
            except (ValueError, json.JSONDecodeError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            return
        if caminho == "/api/personalizacoes/planta":
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                torre = str(dados.get("torre", "")); unidade = str(dados.get("unidade", "")).strip()
                titulo = str(dados.get("titulo", "")).strip(); imagem = str(dados.get("imagem", "")).strip()
                if torre not in TORRES_NOMES or not unidade or not titulo:
                    raise ValueError("Opção de planta inválida")
                plantas = json.loads(ARQUIVO_PLANTAS.read_text(encoding="utf-8")) if ARQUIVO_PLANTAS.exists() else {}
                planta = {"tipo": titulo, "plantaMiniatura": imagem, "origem": dados.get("origem", "Personalização cadastrada")}
                plantas[f"{torre}|{unidade}"] = planta
                temporario = ARQUIVO_PLANTAS.with_suffix(".json.tmp")
                temporario.write_text(json.dumps(plantas, ensure_ascii=False, indent=2), encoding="utf-8")
                temporario.replace(ARQUIVO_PLANTAS)
                self.enviar_json({"ok": True, "planta": planta}, 201)
            except (ValueError, json.JSONDecodeError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            return
        if caminho == "/api/setores/importar":
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                linhas = dados.get("linhas", [])
                if not isinstance(linhas, list) or not linhas or len(linhas) > 5000:
                    raise ValueError("O cadastro deve conter entre 1 e 5.000 linhas")
                mapa_torres = {
                    "home": "aurora", "torre home": "aurora", "aurora": "aurora",
                    "smart": "horizonte", "torre smart": "horizonte", "horizonte": "horizonte",
                }
                mapa_andares = {
                    "fundacao": -2, "fundação": -2, "1 subsolo": -1, "1º subsolo": -1,
                    "subsolo": -1, "terreo": 0, "térreo": 0, "barrilete": 101,
                    "reservatorio": 102, "reservatório": 102, "cobertura": 103,
                }
                criados = 0
                ignorados = 0
                erros = []
                with conectar() as conexao:
                    for indice, linha in enumerate(linhas, start=1):
                        try:
                            torre_texto = str(linha.get("torre", "")).strip().lower()
                            torre = mapa_torres.get(torre_texto)
                            andar_texto = str(linha.get("pavimento", linha.get("andar", ""))).strip().lower()
                            andar_normalizado = andar_texto.replace("º andar", "").replace(" andar", "").strip()
                            andar = mapa_andares.get(andar_texto)
                            if andar is None:
                                andar = int(andar_normalizado)
                            setor = str(linha.get("setor", "")).strip()
                            atividade_texto = str(linha.get("atividade", "")).strip()
                            subservico = str(linha.get("subservico", "")).strip()
                            validar_escopos([{"torre": torre or "", "andar": andar, "unidade": setor}])
                            if not setor:
                                raise ValueError("setor não informado")
                            if atividade_texto.lower() in {"todas", "todos", "*"}:
                                atividades = [
                                    item["nome"] for item in conexao.execute(
                                        "SELECT DISTINCT nome FROM atividades_config WHERE torre=? AND andar=? ORDER BY nome",
                                        (torre, andar),
                                    ).fetchall()
                                ]
                            else:
                                atividades = [atividade_texto]
                            if not atividades:
                                raise ValueError("nenhuma atividade disponível no andar")
                            for atividade in atividades:
                                existe = conexao.execute(
                                    "SELECT 1 FROM atividades_config WHERE nome=? AND torre=? AND andar=? LIMIT 1",
                                    (atividade, torre, andar),
                                ).fetchone()
                                if not existe:
                                    raise ValueError(f"atividade '{atividade}' não disponível no andar")
                                if subservico:
                                    unidades = setores_efetivos_subservico(conexao, atividade, subservico, torre, andar)
                                    if setor in unidades:
                                        ignorados += 1
                                        continue
                                    salvar_setores_subservico(conexao, atividade, subservico, torre, andar, unidades + [setor])
                                    garantir_registros_atividade(conexao, atividade, [(torre, andar, setor)])
                                    criados += 1
                                    continue
                                conflito = conexao.execute(
                                    "SELECT 1 FROM registros WHERE torre=? AND andar=? AND unidade=? AND atividade=? LIMIT 1",
                                    (torre, andar, setor, atividade),
                                ).fetchone()
                                if conflito:
                                    ignorados += 1
                                    continue
                                conexao.execute(
                                    "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, ?)",
                                    (atividade, torre, andar, setor),
                                )
                                garantir_registros_atividade(conexao, atividade, [(torre, andar, setor)])
                                criados += 1
                        except (ValueError, TypeError) as erro:
                            erros.append(f"Linha {indice}: {erro}")
                self.enviar_json({"ok": True, "criados": criados, "ignorados": ignorados, "erros": erros[:30]}, 201)
            except (ValueError, json.JSONDecodeError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            except Exception as erro:
                self.enviar_json({"erro": f"Falha ao salvar setores: {erro}"}, 500)
            return
        if caminho == "/api/cadastro":
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                nome = str(dados.get("nome", "")).strip()
                email = str(dados.get("email", "")).strip().lower()
                cargo = str(dados.get("cargo", "")).strip()
                senha = str(dados.get("senha", ""))
                if not nome or not email.endswith("@dialogo.com.br") or len(senha) < 8:
                    raise ValueError("Informe nome, e-mail @dialogo.com.br e senha com ao menos 8 caracteres")
                with conectar() as conexao:
                    cargo_config = conexao.execute(
                        "SELECT cargo, perfil FROM cargos_perfis WHERE cargo=? COLLATE NOCASE AND ativo=1", (cargo,)
                    ).fetchone()
                    if not cargo_config:
                        raise ValueError("Selecione um cargo válido")
                    cursor = conexao.execute(
                        """
                        INSERT INTO usuarios_acesso (nome,email,cargo,perfil,senha_hash,torres,ativo)
                        VALUES (?,?,?,?,?,'["aurora", "horizonte"]',0)
                        """,
                        (nome, email, cargo_config["cargo"], cargo_config["perfil"], gerar_hash_senha(senha)),
                    )
                    registrar_auditoria(conexao, cursor.lastrowid, nome, "Cadastro solicitado", f"{email} · {cargo_config['cargo']} · aguardando aprovação")
                self.enviar_json({"ok": True, "mensagem": "Cadastro enviado para aprovação."}, 201)
            except (ValueError, json.JSONDecodeError, sqlite3.IntegrityError) as erro:
                mensagem = "Este e-mail já está cadastrado" if isinstance(erro, sqlite3.IntegrityError) else str(erro)
                self.enviar_json({"erro": mensagem}, 400)
            return
        if caminho == "/api/login":
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                identificador = str(dados.get("usuario", "")).strip()
                with conectar() as conexao:
                    usuario = conexao.execute(
                        "SELECT * FROM usuarios_acesso WHERE email = ? COLLATE NOCASE AND ativo = 1",
                        (identificador,),
                    ).fetchone()
                    if not usuario or not conferir_senha(dados.get("senha", ""), usuario["senha_hash"]):
                        registrar_auditoria(conexao, None, identificador or "Não informado", "Acesso recusado", "Credenciais inválidas ou usuário inativo.")
                        self.enviar_json({"erro": "Usuário ou senha inválidos"}, 401)
                        return
                    conexao.execute("UPDATE usuarios_acesso SET ultimo_acesso=CURRENT_TIMESTAMP WHERE id=?", (usuario["id"],))
                    registrar_auditoria(conexao, usuario["id"], usuario["nome"], "Acesso realizado", usuario["email"])
                    permissoes = permissoes_do_perfil(conexao, usuario["perfil"])
                if not usuario:
                    self.enviar_json({"erro": "Usuário ou senha inválidos"}, 401)
                    return
                token = secrets.token_urlsafe(32)
                SESSOES_ENGENHEIRO[token] = {"id": usuario["id"], "nome": usuario["nome"], "email": usuario["email"], "perfil": usuario["perfil"], "permissoes": permissoes}
                corpo = json.dumps({"ok": True, "usuario": usuario["nome"], "perfil": usuario["perfil"], "permissoes": permissoes, "podeAdministrar": usuario["perfil"] in PERFIS_ADMINISTRACAO and "administrar_acessos" in permissoes}, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                seguro = "; Secure" if self.headers.get("X-Forwarded-Proto", "").split(",", 1)[0].strip() == "https" else ""
                self.send_header("Set-Cookie", f"sessao_obra={token}; Path=/; HttpOnly; SameSite=Strict{seguro}")
                self.send_header("Content-Length", str(len(corpo)))
                self.end_headers()
                self.wfile.write(corpo)
            except Exception:
                self.enviar_json({"erro": "Não foi possível realizar o login"}, 400)
            return
        if caminho == "/api/logout":
            cookies = self.headers.get("Cookie", "")
            token = next((parte.split("=", 1)[1] for parte in cookies.split("; ") if parte.startswith("sessao_obra=")), "")
            usuario = SESSOES_ENGENHEIRO.pop(token, None)
            if usuario:
                with conectar() as conexao:
                    registrar_auditoria(conexao, usuario["id"], usuario["nome"], "Saída do sistema", usuario["email"])
            corpo = b'{"ok":true}'
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Set-Cookie", "sessao_obra=; Path=/; Max-Age=0; HttpOnly; SameSite=Strict")
            self.send_header("Content-Length", str(len(corpo)))
            self.end_headers()
            self.wfile.write(corpo)
            return
        if caminho == "/api/admin/usuarios":
            administrador = self.exigir_administracao()
            if not administrador:
                return
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                nome = str(dados.get("nome", "")).strip()
                email = str(dados.get("email", "")).strip().lower()
                cargo = str(dados.get("cargo", "")).strip()
                senha = str(dados.get("senha", ""))
                torres = [torre for torre in dados.get("torres", []) if torre in TORRES_NOMES]
                if not nome or "@" not in email or not cargo or len(senha) < 8:
                    raise ValueError("Informe nome, e-mail, cargo e senha com ao menos 8 caracteres")
                if not torres:
                    raise ValueError("Selecione ao menos uma torre")
                with conectar() as conexao:
                    cargo_config = conexao.execute(
                        "SELECT cargo, perfil FROM cargos_perfis WHERE cargo=? COLLATE NOCASE AND ativo=1", (cargo,)
                    ).fetchone()
                    if not cargo_config:
                        raise ValueError("Selecione um cargo ativo cadastrado")
                    cargo, perfil = cargo_config["cargo"], cargo_config["perfil"]
                    cursor = conexao.execute(
                        "INSERT INTO usuarios_acesso (nome,email,cargo,perfil,senha_hash,torres) VALUES (?,?,?,?,?,?)",
                        (nome, email, cargo, perfil, gerar_hash_senha(senha), json.dumps(torres)),
                    )
                    registrar_auditoria(conexao, administrador["id"], administrador["nome"], "Usuário cadastrado", f"{nome} · {email} · {PERFIS_ACESSO[perfil]}")
                self.enviar_json({"ok": True, "id": cursor.lastrowid}, 201)
            except (ValueError, json.JSONDecodeError, sqlite3.IntegrityError) as erro:
                mensagem = "Este e-mail já está cadastrado" if isinstance(erro, sqlite3.IntegrityError) else str(erro)
                self.enviar_json({"erro": mensagem}, 400)
            return
        if caminho == "/api/admin/cargos":
            administrador = self.exigir_administracao()
            if not administrador:
                return
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                cargo = str(dados.get("cargo", "")).strip()
                perfil = str(dados.get("perfil", "")).strip()
                if not cargo or perfil not in PERFIS_ACESSO:
                    raise ValueError("Informe o cargo e um perfil válido")
                with conectar() as conexao:
                    cursor = conexao.execute("INSERT INTO cargos_perfis (cargo, perfil) VALUES (?, ?)", (cargo, perfil))
                    registrar_auditoria(conexao, administrador["id"], administrador["nome"], "Cargo cadastrado", f"{cargo} · {PERFIS_ACESSO[perfil]}")
                self.enviar_json({"ok": True, "id": cursor.lastrowid}, 201)
            except (ValueError, json.JSONDecodeError, sqlite3.IntegrityError) as erro:
                mensagem = "Este cargo já está cadastrado" if isinstance(erro, sqlite3.IntegrityError) else str(erro)
                self.enviar_json({"erro": mensagem}, 400)
            return
        if caminho == "/api/setores":
            if not self.exigir_engenheiro():
                return
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                torre = str(dados.get("torre", ""))
                andar = int(dados.get("andar", 0))
                atividade = str(dados.get("atividade", "")).strip()
                unidade = str(dados.get("unidade", "")).strip()
                subservico = str(dados.get("subservico", "")).strip()
                validar_escopos([{"torre": torre, "andar": andar, "unidade": unidade}])
                if not atividade or not unidade:
                    raise ValueError("Informe o serviço e o nome do setor")
                with conectar() as conexao:
                    atividade_no_andar = conexao.execute(
                        "SELECT 1 FROM atividades_config WHERE nome=? AND torre=? AND andar=? LIMIT 1",
                        (atividade, torre, andar),
                    ).fetchone()
                    if not atividade_no_andar:
                        raise ValueError("O serviço não está disponível neste pavimento")
                    if subservico:
                        unidades = setores_efetivos_subservico(conexao, atividade, subservico, torre, andar)
                        if unidade in unidades:
                            raise ValueError("Este setor já existe para o subserviço neste pavimento")
                        salvar_setores_subservico(conexao, atividade, subservico, torre, andar, unidades + [unidade])
                        garantir_registros_atividade(conexao, atividade, [(torre, andar, unidade)])
                        self.enviar_json({"ok": True, "unidade": unidade}, 201)
                        return
                    conflito = conexao.execute(
                        "SELECT 1 FROM registros WHERE torre=? AND andar=? AND unidade=? AND atividade=? LIMIT 1",
                        (torre, andar, unidade, atividade),
                    ).fetchone()
                    if conflito:
                        raise ValueError("Este setor já existe para o serviço neste pavimento")
                    conexao.execute(
                        "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, ?)",
                        (atividade, torre, andar, unidade),
                    )
                    garantir_registros_atividade(conexao, atividade, [(torre, andar, unidade)])
                self.enviar_json({"ok": True, "unidade": unidade}, 201)
            except (ValueError, json.JSONDecodeError, sqlite3.IntegrityError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            return
        if caminho == "/api/atividades":
            if not self.exigir_engenheiro():
                return
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                nome = str(dados.get("nome", "")).strip()
                if not nome:
                    raise ValueError("Informe o nome da atividade")
                escopos = validar_escopos(dados.get("escopos", []))
                especificacoes = validar_especificacoes(dados.get("especificacoes", []))
                with conectar() as conexao:
                    for torre, andar, unidade in escopos:
                        conexao.execute(
                            "INSERT INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, ?)",
                            (nome, torre, andar, unidade),
                        )
                    salvar_especificacoes_atividade(conexao, nome, especificacoes)
                    garantir_registros_atividade(conexao, nome, escopos)
                self.enviar_json({"ok": True, "nome": nome}, 201)
            except (ValueError, json.JSONDecodeError, sqlite3.IntegrityError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            except Exception as erro:
                self.enviar_json({"erro": f"Falha ao cadastrar atividade: {erro}"}, 500)
            return
        if caminho == "/api/projetos":
            if not self.exigir_engenheiro():
                return
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                if tamanho > 12 * 1024 * 1024:
                    self.enviar_json({"erro": "Imagem acima do limite permitido"}, 413)
                    return
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                campos = ("torre", "andar", "unidade", "titulo", "imagem")
                if any(not dados.get(campo) for campo in campos):
                    raise ValueError("Informe a unidade, o título e a imagem do projeto")
                if not str(dados["imagem"]).startswith("data:image/"):
                    raise ValueError("Envie uma imagem PNG, JPG ou WEBP")
                with conectar() as conexao:
                    cursor = conexao.execute(
                        "INSERT INTO projetos_unidade (torre, andar, unidade, titulo, imagem, arquivo_nome) VALUES (?, ?, ?, ?, ?, ?)",
                        (dados["torre"], int(dados["andar"]), dados["unidade"],
                         str(dados["titulo"]).strip(), dados["imagem"], dados.get("arquivoNome", "")),
                    )
                self.enviar_json({"ok": True, "id": cursor.lastrowid}, 201)
            except (ValueError, json.JSONDecodeError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            except Exception as erro:
                self.enviar_json({"erro": f"Falha ao salvar projeto: {erro}"}, 500)
            return
        if caminho == "/api/comentarios-unidade":
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                torre = str(dados.get("torre", "")).strip()
                andar = int(dados.get("andar", 0))
                unidade = str(dados.get("unidade", "")).strip()
                comentario = str(dados.get("comentario", "")).strip()
                if torre not in TORRES_NOMES or not unidade or not comentario:
                    raise ValueError("Informe a unidade e o comentário")
                if len(comentario) > 2000:
                    raise ValueError("O comentário deve ter no máximo 2.000 caracteres")
                usuario = self.sessao_usuario() or {}
                autor = str(usuario.get("nome", "")).strip()
                with conectar() as conexao:
                    cursor = conexao.execute(
                        "INSERT INTO comentarios_unidade (torre, andar, unidade, comentario, autor) VALUES (?, ?, ?, ?, ?)",
                        (torre, andar, unidade, comentario, autor),
                    )
                    linha = conexao.execute(
                        "SELECT id, torre, andar, unidade, comentario, autor, criado_em FROM comentarios_unidade WHERE id=?",
                        (cursor.lastrowid,),
                    ).fetchone()
                self.enviar_json(dict(linha), 201)
            except (ValueError, json.JSONDecodeError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            except Exception as erro:
                self.enviar_json({"erro": f"Falha ao salvar comentário: {erro}"}, 500)
            return
        if caminho == "/api/projetos-ocultos":
            if not self.exigir_engenheiro():
                return
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                campos = ("torre", "andar", "unidade", "imagem")
                if any(not dados.get(campo) for campo in campos):
                    raise ValueError("Projeto inválido")
                with conectar() as conexao:
                    conexao.execute(
                        "INSERT OR IGNORE INTO projetos_ocultos (torre, andar, unidade, imagem) VALUES (?, ?, ?, ?)",
                        (dados["torre"], int(dados["andar"]), dados["unidade"], dados["imagem"]),
                    )
                self.enviar_json({"ok": True}, 201)
            except (ValueError, json.JSONDecodeError) as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            return
        if caminho != "/api/ocorrencias":
            self.enviar_json({"erro": "Rota não encontrada"}, 404)
            return
        if not self.exigir_engenheiro():
            return
        try:
            tamanho = int(self.headers.get("Content-Length", "0"))
            if tamanho > 20 * 1024 * 1024:
                self.enviar_json({"erro": "Foto acima do limite permitido"}, 413)
                return
            dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
            campos_texto = ("torre", "unidade", "atividade", "especificacao", "descricao", "dataOcorrencia")
            if any(not dados.get(campo) for campo in campos_texto) or dados.get("andar") is None:
                self.enviar_json({"erro": "Preencha torre, andar, unidade e descrição"}, 400)
                return
            ocorrencia_seguranca = str(dados.get("atividade", "")).strip().casefold() == "segurança".casefold()
            if ocorrencia_seguranca:
                dados["atividade"] = "Segurança"
                dados["subatividade"] = ""
                dados["especificacao"] = "Ocorrência de segurança"
            with conectar() as conexao:
                cursor = conexao.execute(
                    """
                    INSERT INTO ocorrencias
                    (torre, andar, unidade, atividade, subatividade, especificacao, descricao, foto, foto_nome, data_ocorrencia, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        dados["torre"],
                        int(dados["andar"]),
                        dados["unidade"],
                        dados["atividade"],
                        dados.get("subatividade", ""),
                        dados.get("especificacao", ""),
                        dados["descricao"],
                        dados.get("foto", ""),
                        dados.get("fotoNome", ""),
                        dados["dataOcorrencia"],
                        "pendente",
                    ),
                )
                identificador = cursor.lastrowid
                if not ocorrencia_seguranca:
                    chave_registro = f"{dados['torre']}|{int(dados['andar'])}|{dados['unidade']}|{dados['atividade']}"
                    linha_registro = conexao.execute(
                        "SELECT especificacoes FROM registros WHERE chave = ?",
                        (chave_registro,),
                    ).fetchone()
                    if not linha_registro:
                        raise ValueError("O serviço selecionado não existe para este apartamento ou setor")
                    try:
                        estados_especificacoes = json.loads(linha_registro["especificacoes"] or "{}")
                    except (json.JSONDecodeError, TypeError):
                        estados_especificacoes = {}
                    estados_especificacoes[dados["especificacao"]] = "pendente"
                    conexao.execute(
                        """
                        UPDATE registros
                        SET especificacoes=?, status='pendente', concluido=0,
                            data_conclusao=?, atualizado_em=CURRENT_TIMESTAMP
                        WHERE chave=?
                        """,
                        (json.dumps(estados_especificacoes, ensure_ascii=False), dados["dataOcorrencia"], chave_registro),
                    )
            self.enviar_json({"ok": True, "id": identificador}, 201)
        except (ValueError, json.JSONDecodeError) as erro:
            self.enviar_json({"erro": str(erro)}, 400)
        except Exception as erro:
            self.enviar_json({"erro": f"Falha ao salvar ocorrência: {erro}"}, 500)

    def do_PATCH(self):
        caminho = urlparse(self.path).path
        if caminho.startswith("/api/ocorrencias/") and not self.exigir_permissao("gerenciar_ocorrencias"):
            return
        partes = caminho.strip("/").split("/")
        if len(partes) != 3 or partes[:2] != ["api", "ocorrencias"]:
            self.enviar_json({"erro": "Rota não encontrada"}, 404)
            return
        if not self.exigir_engenheiro():
            return
        try:
            identificador = int(partes[2])
            tamanho = int(self.headers.get("Content-Length", "0"))
            dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
            status = dados.get("status", "")
            permitidos = {"pendente", "concluido"}
            if status not in permitidos:
                self.enviar_json({"erro": "Status inválido"}, 400)
                return
            with conectar() as conexao:
                cursor = conexao.execute(
                    "UPDATE ocorrencias SET status = ? WHERE id = ?",
                    (status, identificador),
                )
            if not cursor.rowcount:
                self.enviar_json({"erro": "Ocorrência não encontrada"}, 404)
                return
            self.enviar_json({"ok": True, "id": identificador, "status": status})
        except (ValueError, json.JSONDecodeError) as erro:
            self.enviar_json({"erro": str(erro)}, 400)
        except Exception as erro:
            self.enviar_json({"erro": f"Falha ao atualizar ocorrência: {erro}"}, 500)

    def do_DELETE(self):
        caminho = urlparse(self.path).path
        partes = caminho.strip("/").split("/")
        if len(partes) == 4 and partes[:3] == ["api", "admin", "cargos"]:
            administrador = self.exigir_administracao()
            if not administrador:
                return
            try:
                cargo_id = int(partes[3])
                with conectar() as conexao:
                    cargo = conexao.execute("SELECT cargo FROM cargos_perfis WHERE id=?", (cargo_id,)).fetchone()
                    if not cargo:
                        raise ValueError("Cargo não encontrado")
                    vinculados = conexao.execute(
                        "SELECT COUNT(*) AS total FROM usuarios_acesso WHERE cargo=? COLLATE NOCASE", (cargo["cargo"],)
                    ).fetchone()["total"]
                    if vinculados:
                        raise ValueError(f"Este cargo possui {vinculados} usuário(s) vinculado(s). Altere o cargo desses usuários antes de excluir.")
                    conexao.execute("DELETE FROM cargos_perfis WHERE id=?", (cargo_id,))
                    registrar_auditoria(conexao, administrador["id"], administrador["nome"], "Cargo excluído", cargo["cargo"])
                self.enviar_json({"ok": True, "id": cargo_id})
            except ValueError as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            return
        if caminho == "/api/setores" and not self.exigir_permissao("gerenciar_setores"):
            return
        if caminho == "/api/atividades" and not self.exigir_permissao("gerenciar_atividades"):
            return
        if len(partes) == 3 and partes[:2] == ["api", "projetos"] and not self.exigir_permissao("gerenciar_projetos"):
            return
        if len(partes) == 3 and partes[:2] == ["api", "ocorrencias"] and not self.exigir_permissao("gerenciar_ocorrencias"):
            return
        if caminho == "/api/setores":
            if not self.exigir_engenheiro():
                return
            try:
                parametros = parse_qs(urlparse(self.path).query)
                torre = parametros.get("torre", [""])[0]
                andar = int(parametros.get("andar", ["0"])[0])
                unidade = parametros.get("unidade", [""])[0].strip()
                atividade = parametros.get("atividade", [""])[0].strip()
                subservico = parametros.get("subservico", [""])[0].strip()
                if torre not in TORRES_NOMES or not unidade or andar < -2:
                    raise ValueError("Setor inválido")
                with conectar() as conexao:
                    if atividade and subservico:
                        unidades = setores_efetivos_subservico(conexao, atividade, subservico, torre, andar)
                        if unidade not in unidades:
                            raise ValueError("Setor não encontrado neste subserviço")
                        salvar_setores_subservico(conexao, atividade, subservico, torre, andar, [item for item in unidades if item != unidade])
                        conexao.execute("DELETE FROM ocorrencias WHERE torre=? AND andar=? AND unidade=? AND atividade=? AND subatividade=?", (torre, andar, unidade, atividade, subservico))
                        self.enviar_json({"ok": True})
                        return
                    if atividade:
                        destinos_registrados = [
                            linha["unidade"] for linha in conexao.execute(
                                "SELECT DISTINCT unidade FROM registros WHERE torre=? AND andar=? AND atividade=? AND unidade!=? ORDER BY unidade",
                                (torre, andar, atividade, unidade),
                            ).fetchall()
                        ]
                        possui_escopo_geral = conexao.execute(
                            "SELECT 1 FROM atividades_config WHERE torre=? AND andar=? AND nome=? AND unidade='*' LIMIT 1",
                            (torre, andar, atividade),
                        ).fetchone()
                        if possui_escopo_geral:
                            destinos_automaticos = apartamentos_do_andar(torre, andar) + ["Área comum"]
                            destinos_explicitos = [
                                linha["unidade"] for linha in conexao.execute(
                                    "SELECT DISTINCT unidade FROM atividades_config WHERE torre=? AND andar=? AND nome=? AND unidade!='*'",
                                    (torre, andar, atividade),
                                ).fetchall()
                            ]
                            destinos_restantes = list(dict.fromkeys(
                                destino for destino in destinos_automaticos + destinos_explicitos + destinos_registrados
                                if destino != unidade
                            ))
                            conexao.execute(
                                "DELETE FROM atividades_config WHERE torre=? AND andar=? AND nome=?",
                                (torre, andar, atividade),
                            )
                            conexao.executemany(
                                "INSERT OR IGNORE INTO atividades_config (nome, torre, andar, unidade) VALUES (?, ?, ?, ?)",
                                [(atividade, torre, andar, destino) for destino in destinos_restantes],
                            )
                        else:
                            cursor_config = conexao.execute(
                                "DELETE FROM atividades_config WHERE torre=? AND andar=? AND nome=? AND unidade=?",
                                (torre, andar, atividade, unidade),
                            )
                        cursor = conexao.execute(
                            "DELETE FROM registros WHERE torre=? AND andar=? AND unidade=? AND atividade=?",
                            (torre, andar, unidade, atividade),
                        )
                        conexao.execute(
                            "DELETE FROM ocorrencias WHERE torre=? AND andar=? AND unidade=? AND atividade=?",
                            (torre, andar, unidade, atividade),
                        )
                        setor_encontrado = bool(cursor.rowcount or possui_escopo_geral or (not possui_escopo_geral and cursor_config.rowcount))
                    else:
                        cursor = conexao.execute("DELETE FROM registros WHERE torre=? AND andar=? AND unidade=?", (torre, andar, unidade))
                        for tabela in ("ocorrencias", "projetos_unidade", "projetos_ocultos"):
                            conexao.execute(f"DELETE FROM {tabela} WHERE torre=? AND andar=? AND unidade=?", (torre, andar, unidade))
                        conexao.execute("DELETE FROM atividades_config WHERE torre=? AND andar=? AND unidade=?", (torre, andar, unidade))
                        setor_encontrado = bool(cursor.rowcount)
                if not setor_encontrado:
                    self.enviar_json({"erro": "Setor não encontrado"}, 404)
                    return
                self.enviar_json({"ok": True})
            except ValueError as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            return
        if len(partes) == 3 and partes[:2] == ["api", "projetos"]:
            if not self.exigir_engenheiro():
                return
            try:
                identificador = int(partes[2])
                with conectar() as conexao:
                    cursor = conexao.execute("DELETE FROM projetos_unidade WHERE id = ?", (identificador,))
                if not cursor.rowcount:
                    self.enviar_json({"erro": "Projeto não encontrado"}, 404)
                    return
                self.enviar_json({"ok": True})
            except ValueError:
                self.enviar_json({"erro": "Projeto inválido"}, 400)
            return
        if caminho == "/api/atividades":
            if not self.exigir_engenheiro():
                return
            try:
                parametros = parse_qs(urlparse(self.path).query)
                nome = parametros.get("nome", [""])[0].strip()
                if not nome:
                    raise ValueError("Informe a atividade")
                with conectar() as conexao:
                    cursor = conexao.execute("DELETE FROM atividades_config WHERE nome = ?", (nome,))
                    conexao.execute("DELETE FROM atividades_especificacoes WHERE atividade = ?", (nome,))
                if not cursor.rowcount:
                    self.enviar_json({"erro": "Atividade não encontrada"}, 404)
                    return
                self.enviar_json({"ok": True, "nome": nome})
            except ValueError as erro:
                self.enviar_json({"erro": str(erro)}, 400)
            except Exception as erro:
                self.enviar_json({"erro": f"Falha ao excluir atividade: {erro}"}, 500)
            return
        partes = caminho.strip("/").split("/")
        if len(partes) != 3 or partes[:2] != ["api", "ocorrencias"]:
            self.enviar_json({"erro": "Rota não encontrada"}, 404)
            return
        if not self.exigir_engenheiro():
            return
        try:
            identificador = int(partes[2])
            with conectar() as conexao:
                cursor = conexao.execute(
                    "DELETE FROM ocorrencias WHERE id = ?",
                    (identificador,),
                )
            if not cursor.rowcount:
                self.enviar_json({"erro": "Ocorrência não encontrada"}, 404)
                return
            self.enviar_json({"ok": True, "id": identificador})
        except ValueError as erro:
            self.enviar_json({"erro": str(erro)}, 400)
        except Exception as erro:
            self.enviar_json({"erro": f"Falha ao excluir ocorrência: {erro}"}, 500)

    def log_message(self, formato, *args):
        print(f"[servidor] {self.address_string()} - {formato % args}")


if __name__ == "__main__":
    preparar_banco()
    endereco = f"http://127.0.0.1:{PORTA}"
    try:
        ip_rede = endereco_rede()
    except Exception:
        ip_rede = ""
    servidor = ThreadingHTTPServer((HOST, PORTA), ServidorObra)
    print(f"Banco SQLite: {BANCO}")
    print(f"Painel neste Mac: {endereco}")
    if ip_rede:
        print(f"Painel no celular: http://{ip_rede}:{PORTA}")
    print("Para encerrar, pressione Control + C.")
    if not os.environ.get("RENDER"):
        threading.Timer(0.7, lambda: subprocess.run(["open", "-a", "Safari", endereco], check=False)).start()
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor encerrado.")
    finally:
        servidor.server_close()
