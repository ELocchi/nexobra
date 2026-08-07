#!/usr/bin/env python3
import base64
import json
import os
import secrets
import sqlite3
import subprocess
import threading
import webbrowser
from datetime import date, datetime, timedelta
from io import BytesIO
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse
from xml.sax.saxutils import escape

from reportlab.graphics import renderSVG
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image, KeepInFrame, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

PASTA = Path(__file__).resolve().parent
ARQUIVO_PLANEJAMENTO = PASTA / "dados_planilha_atual.json"
ARQUIVO_ACABAMENTOS = PASTA / "acabamentos_unidades.json"
BANCO = Path(os.environ.get("OBRA_BANCO", str(PASTA / "acompanhamento.db")))
HOST = "0.0.0.0"
PORTA = int(os.environ.get("PORT", "8000"))
TORRES_NOMES = {"aurora": "Torre Home", "horizonte": "Torre Smart"}
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
SESSOES_ENGENHEIRO = set()


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
    conexao = sqlite3.connect(BANCO)
    conexao.row_factory = sqlite3.Row
    return conexao


def gerar_relatorio_pdf(torre, andar, unidade):
    with conectar() as conexao:
        registros = conexao.execute(
            "SELECT atividade, status, data_conclusao, observacao FROM registros WHERE torre = ? AND andar = ? AND unidade = ? ORDER BY atividade",
            (torre, andar, unidade),
        ).fetchall()
        registros = somente_registros_planejados(registros, torre)
        ocorrencias = conexao.execute(
            "SELECT atividade, status, data_ocorrencia, descricao FROM ocorrencias WHERE torre = ? AND andar = ? AND unidade = ? ORDER BY data_ocorrencia DESC, id DESC",
            (torre, andar, unidade),
        ).fetchall()
    memoria = BytesIO()
    documento = SimpleDocTemplate(memoria, pagesize=A4, rightMargin=16*mm, leftMargin=16*mm, topMargin=14*mm, bottomMargin=14*mm)
    estilos = getSampleStyleSheet()
    elementos = []
    logo = PASTA / "LOGO-DIALOGO.png"
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


def gerar_historico_ocorrencias_pdf(torre, andar="todos", status="todos", periodo="todas"):
    with conectar() as conexao:
        ocorrencias = conexao.execute(
            "SELECT id, andar, unidade, atividade, status, data_ocorrencia, descricao, foto, foto_nome, criado_em FROM ocorrencias WHERE torre = ? ORDER BY data_ocorrencia DESC, id DESC",
            (torre,),
        ).fetchall()
    if andar != "todos":
        ocorrencias = [item for item in ocorrencias if item["andar"] == int(andar)]
    if status != "todos":
        ocorrencias = [item for item in ocorrencias if item["status"] == status]
    if periodo != "todas":
        limite = date.today() - timedelta(days=max(int(periodo) - 1, 0))
        ocorrencias = [
            item for item in ocorrencias
            if item["data_ocorrencia"] and date.fromisoformat(item["data_ocorrencia"]) >= limite
        ]
    memoria = BytesIO()
    documento = SimpleDocTemplate(
        memoria, pagesize=A4, rightMargin=18*mm, leftMargin=18*mm,
        topMargin=36*mm, bottomMargin=22*mm,
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
            ("Localização:", TORRES_NOMES.get(torre, torre)),
            ("Título:", "Histórico de ocorrências"),
            ("No. Items:", str(len(ocorrencias))),
        ]
        for rotulo, valor in linhas:
            canvas.drawString(20*mm, y, rotulo)
            canvas.drawString(42*mm, y, valor)
            y -= 4.2*mm
        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(.7)
        canvas.line(20*mm, altura-34*mm, largura-20*mm, altura-34*mm)
        canvas.line(20*mm, 17*mm, largura-20*mm, 17*mm)
        logo = PASTA / "LOGO-DIALOGO.png"
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
        detalhes = (
            f"<b>Criada:</b>&nbsp;&nbsp; {data}<br/>"
            f"<b>({numero})</b>&nbsp;&nbsp; {escape(item['unidade'])} · {item['andar']}º andar<br/>"
            f"<b>Atividade:</b>&nbsp;&nbsp; {escape(item['atividade'])}<br/>"
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


def gerar_pagina_visitante(torre, andar, unidade):
    with conectar() as conexao:
        registros = conexao.execute(
            "SELECT atividade, status, data_conclusao, observacao, foto, foto_nome FROM registros WHERE torre = ? AND andar = ? AND unidade = ? ORDER BY atividade",
            (torre, andar, unidade),
        ).fetchall()
        registros = somente_registros_planejados(registros, torre)
        ocorrencias = conexao.execute(
            "SELECT atividade, status, data_ocorrencia, descricao, foto, foto_nome FROM ocorrencias WHERE torre = ? AND andar = ? AND unidade = ? ORDER BY data_ocorrencia DESC, id DESC",
            (torre, andar, unidade),
        ).fetchall()
    total = len(registros)
    concluidos = sum(1 for item in registros if item["status"] == "concluido")
    percentual = round(concluidos / total * 100) if total else 0
    consulta = urlencode({"torre": torre, "andar": andar, "unidade": unidade})
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
</style></head><body><header><div><img src="/LOGO-DIALOGO.png" alt="Diálogo Engenharia"><h1>{escape(nome_unidade)} · {andar}º andar</h1><p>{escape(TORRES_NOMES.get(torre, torre))} · visualização para visitantes</p></div></header>
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
    logo = PASTA / "LOGO-DIALOGO.png"
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
</style></head><body><header><div><img src="/LOGO-DIALOGO.png" alt="Diálogo Engenharia"><h1>Relatórios da obra — visão visitante</h1><p>{escape(TORRES_NOMES.get(torre, torre))}</p></div></header><main>
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


def endereco_rede():
    try:
        return subprocess.check_output(["ipconfig", "getifaddr", "en0"], text=True).strip()
    except Exception:
        return ""


def gerar_inicio_visitante():
    return """<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Acesso visitante</title><style>
*{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;background:#123554;font-family:Arial,sans-serif;color:#243445}.caixa{width:min(92%,720px);background:#fff;border-radius:15px;padding:28px;box-shadow:0 18px 50px #0005;text-align:center}.logo{width:190px;margin-bottom:18px}h1{color:#173a5e}.torres{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:22px}.torre{display:block;padding:25px 15px;border:1px solid #ccd6de;border-radius:11px;text-decoration:none;color:#173a5e;font-weight:bold;background:#f4f7f9}.torre:hover{background:#e8f1f6;border-color:#17608f}.torre span{display:block;font-size:.75rem;color:#687887;margin-top:7px;font-weight:normal}.aviso{font-size:.75rem;color:#687887;margin-top:20px}@media(max-width:550px){.torres{grid-template-columns:1fr}}
</style></head><body><main class="caixa"><img class="logo" src="/LOGO-DIALOGO.png" alt="Diálogo Engenharia"><h1>Visão visitante</h1><p>Selecione a torre que deseja consultar.</p><div class="torres"><a class="torre" href="/visitante-relatorios?torre=aurora">Torre Home<span>36 pavimentos</span></a><a class="torre" href="/visitante-relatorios?torre=horizonte">Torre Smart<span>23 pavimentos</span></a></div><div class="aviso">Acesso somente para leitura. Não é possível alterar informações.</div></main></body></html>""".encode("utf-8")


def preparar_banco():
    with conectar() as conexao:
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
        importar_planejamento_inicial(conexao)
        preparar_atividades_config(conexao)


def apartamentos_do_andar(torre, andar):
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


def listar_atividades_config():
    with conectar() as conexao:
        linhas = conexao.execute(
            "SELECT id, nome, torre, andar, unidade FROM atividades_config ORDER BY nome, torre, andar, unidade"
        ).fetchall()
        linhas_especificacoes = conexao.execute(
            "SELECT atividade, especificacao FROM atividades_especificacoes ORDER BY atividade, ordem, id"
        ).fetchall()
    agrupadas = {}
    for linha in linhas:
        item = agrupadas.setdefault(linha["nome"], {"nome": linha["nome"], "escopos": [], "especificacoes": []})
        item["escopos"].append(
            {"id": linha["id"], "torre": linha["torre"], "andar": linha["andar"], "unidade": linha["unidade"]}
        )
    for linha in linhas_especificacoes:
        if linha["atividade"] in agrupadas:
            agrupadas[linha["atividade"]]["especificacoes"].append(linha["especificacao"])
    return list(agrupadas.values())


def validar_escopos(escopos):
    resultado = []
    for escopo in escopos:
        torre = str(escopo.get("torre", ""))
        andar = int(escopo.get("andar", 0))
        unidade = str(escopo.get("unidade", "*")).strip() or "*"
        maximo = 36 if torre == "aurora" else 23 if torre == "horizonte" else 0
        if not maximo or andar < 1 or andar > maximo:
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

    def enviar_json(self, dados, status=200):
        corpo = json.dumps(dados, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(corpo)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(corpo)

    def sessao_engenheiro(self):
        cookies = self.headers.get("Cookie", "")
        token = next(
            (parte.split("=", 1)[1] for parte in cookies.split("; ") if parte.startswith("sessao_obra=")),
            "",
        )
        return token in SESSOES_ENGENHEIRO

    def exigir_engenheiro(self):
        if self.sessao_engenheiro():
            return True
        self.enviar_json({"erro": "Acesso restrito ao engenheiro"}, 401)
        return False

    def do_GET(self):
        url = urlparse(self.path)
        caminho = url.path
        parametros = parse_qs(url.query)
        if caminho == "/api/sessao":
            self.enviar_json({"autenticado": self.sessao_engenheiro()})
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
            status = parametros.get("status", ["todos"])[0]
            periodo = parametros.get("periodo", ["todas"])[0]
            if torre not in TORRES_NOMES or (andar != "todos" and not andar.isdigit()):
                self.enviar_json({"erro": "Filtros inválidos"}, 400)
                return
            if status not in {*STATUS_NOMES.keys(), "todos"} or periodo not in {"7", "30", "90", "todas"}:
                self.enviar_json({"erro": "Filtros inválidos"}, 400)
                return
            corpo = gerar_historico_ocorrencias_pdf(torre, andar, status, periodo)
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", 'inline; filename="historico-ocorrencias.pdf"')
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
            if torre not in TORRES_NOMES or andar < 1 or not unidade:
                self.enviar_json({"erro": "Unidade inválida"}, 400)
                return
            if caminho == "/visitante":
                corpo = gerar_pagina_visitante(torre, andar, unidade)
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(corpo)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(corpo)
                return
            if caminho == "/relatorio.pdf":
                corpo = gerar_relatorio_pdf(torre, andar, unidade)
                nome = unidade.replace("Apto ", "apartamento-").replace(" ", "-").lower()
                self.send_response(200)
                self.send_header("Content-Type", "application/pdf")
                self.send_header("Content-Disposition", f'inline; filename="relatorio-{nome}.pdf"')
                self.send_header("Content-Length", str(len(corpo)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(corpo)
                return
            consulta = urlencode({"torre": torre, "andar": andar, "unidade": unidade})
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
        if caminho == "/api/registros":
            with conectar() as conexao:
                linhas = conexao.execute("SELECT * FROM registros").fetchall()
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
                        "criadoEm": linha["criado_em"],
                    }
                    for linha in linhas
                ]
            )
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
        if caminho == "/api/login":
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
                dados = json.loads(self.rfile.read(tamanho).decode("utf-8"))
                if str(dados.get("usuario", "")).upper() != USUARIO_ENGENHEIRO or str(dados.get("senha", "")) != SENHA_ENGENHEIRO:
                    self.enviar_json({"erro": "Usuário ou senha inválidos"}, 401)
                    return
                token = secrets.token_urlsafe(32)
                SESSOES_ENGENHEIRO.add(token)
                corpo = json.dumps({"ok": True, "usuario": USUARIO_ENGENHEIRO}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Set-Cookie", f"sessao_obra={token}; Path=/; HttpOnly; SameSite=Strict")
                self.send_header("Content-Length", str(len(corpo)))
                self.end_headers()
                self.wfile.write(corpo)
            except Exception:
                self.enviar_json({"erro": "Não foi possível realizar o login"}, 400)
            return
        if caminho == "/api/logout":
            cookies = self.headers.get("Cookie", "")
            token = next((parte.split("=", 1)[1] for parte in cookies.split("; ") if parte.startswith("sessao_obra=")), "")
            SESSOES_ENGENHEIRO.discard(token)
            corpo = b'{"ok":true}'
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Set-Cookie", "sessao_obra=; Path=/; Max-Age=0; HttpOnly; SameSite=Strict")
            self.send_header("Content-Length", str(len(corpo)))
            self.end_headers()
            self.wfile.write(corpo)
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
            campos = ("torre", "andar", "unidade", "atividade", "descricao", "dataOcorrencia")
            if any(not dados.get(campo) for campo in campos):
                self.enviar_json({"erro": "Preencha torre, andar, unidade e descrição"}, 400)
                return
            with conectar() as conexao:
                cursor = conexao.execute(
                    """
                    INSERT INTO ocorrencias
                    (torre, andar, unidade, atividade, descricao, foto, foto_nome, data_ocorrencia, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        dados["torre"],
                        int(dados["andar"]),
                        dados["unidade"],
                        dados["atividade"],
                        dados["descricao"],
                        dados.get("foto", ""),
                        dados.get("fotoNome", ""),
                        dados["dataOcorrencia"],
                        "pendente",
                    ),
                )
                identificador = cursor.lastrowid
            self.enviar_json({"ok": True, "id": identificador}, 201)
        except (ValueError, json.JSONDecodeError) as erro:
            self.enviar_json({"erro": str(erro)}, 400)
        except Exception as erro:
            self.enviar_json({"erro": f"Falha ao salvar ocorrência: {erro}"}, 500)

    def do_PATCH(self):
        caminho = urlparse(self.path).path
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
        threading.Timer(0.7, lambda: webbrowser.open(endereco)).start()
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor encerrado.")
    finally:
        servidor.server_close()
