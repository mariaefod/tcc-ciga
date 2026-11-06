import docx
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from markdown_it import MarkdownIt
from pathlib import Path

from estilo import CAPA, CENTRO, FOLHA_DE_ROSTO, LARGURA_UTIL, aplicar, numerar

PASTA = Path(__file__).parent.parent / 'projeto'
ARQUIVO = 'Projeto - Maria Eduarda.docx'
LEGENDAS = ('Tabela ', 'Figura ', 'Quadro ', 'Gráfico ', 'Mapa ')
PRE_TEXTUAIS_CORRIDOS = ('RESUMO', 'ABSTRACT')

texto = (PASTA / 'projeto.md').read_text(encoding='utf-8')
_, cabecalho, corpo = texto.split('---\n', 2)
dados = dict(linha.split(': ', 1) for linha in cabecalho.strip().splitlines())

documento = docx.Document()
aplicar(documento)
quebrar = False


def paragrafo(nome):
    global quebrar
    p = documento.add_paragraph(style=nome)
    if quebrar:
        p.paragraph_format.page_break_before = True
        quebrar = False
    return p


paginas = [CAPA]
if 'natureza' in dados:
    paginas.append(FOLHA_DE_ROSTO)
for pagina in paginas:
    for chave, antes in pagina.items():
        if chave not in dados:
            continue
        if chave in ('natureza', 'orientadora'):
            p = paragrafo('Natureza')
        else:
            p = paragrafo('Capa')
        p.paragraph_format.space_before = Pt(antes)
        if chave == 'orientadora':
            p.add_run('Orientadora: ')
        p.add_run(dados[chave])
        if chave == 'titulo' and 'subtitulo' in dados:
            p.add_run(': ')
            p.add_run(dados['subtitulo']).font.all_caps = False
    quebrar = True


def escrever(paragrafo, inline):
    negrito = False
    italico = False
    for filho in inline.children:
        if filho.type == 'strong_open':
            negrito = True
        elif filho.type == 'strong_close':
            negrito = False
        elif filho.type == 'em_open':
            italico = True
        elif filho.type == 'em_close':
            italico = False
        elif filho.type == 'softbreak':
            paragrafo.add_run(' ')
        elif filho.type == 'hardbreak':
            paragrafo.add_run().add_break()
        elif filho.type == 'image':
            imagem = paragrafo.add_run().add_picture(str(PASTA / filho.attrGet('src')))
            largura = Cm(LARGURA_UTIL)
            if imagem.width > largura:
                imagem.height = int(imagem.height * largura / imagem.width)
                imagem.width = largura
        elif filho.content:
            run = paragrafo.add_run(filho.content)
            run.bold = negrito or None
            run.italic = italico or None


def campo(paragrafo, codigo, texto):
    run = paragrafo.add_run()._r
    inicio = OxmlElement('w:fldChar')
    inicio.set(qn('w:fldCharType'), 'begin')
    instrucao = OxmlElement('w:instrText')
    instrucao.text = codigo
    separador = OxmlElement('w:fldChar')
    separador.set(qn('w:fldCharType'), 'separate')
    resultado = OxmlElement('w:t')
    resultado.text = texto
    fim = OxmlElement('w:fldChar')
    fim.set(qn('w:fldCharType'), 'end')
    for elemento in (inicio, instrucao, separador, resultado, fim):
        run.append(elemento)


def borda(pai, nome, lados):
    bordas = OxmlElement(nome)
    for lado in lados:
        linha = OxmlElement(f'w:{lado}')
        linha.set(qn('w:val'), 'single')
        linha.set(qn('w:sz'), '6')
        bordas.append(linha)
    aparencia = pai.find(qn('w:tblLook'))
    if aparencia is None:
        pai.append(bordas)
    else:
        aparencia.addprevious(bordas)


def criar_tabela(linhas):
    tabela = documento.add_table(rows=len(linhas), cols=len(linhas[0]))
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    borda(tabela._tbl.tblPr, 'w:tblBorders', ('top', 'bottom'))
    tamanhos = [max(8, *(min(len(c[0].content), 30) for c in coluna)) for coluna in zip(*linhas)]
    larguras = [Cm(LARGURA_UTIL * tamanho / sum(tamanhos)) for tamanho in tamanhos]
    for coluna, largura in zip(tabela.columns, larguras):
        coluna.width = largura
    for i, linha in enumerate(linhas):
        for j, (inline, alinhamento) in enumerate(linha):
            celula = tabela.cell(i, j)
            celula.width = larguras[j]
            p = celula.paragraphs[0]
            p.style = 'Tabela'
            p.paragraph_format.keep_with_next = True
            if alinhamento == 'text-align:center':
                p.alignment = CENTRO
            escrever(p, inline)
            if i == 0:
                for run in p.runs:
                    run.bold = True
                borda(celula._tc.get_or_add_tcPr(), 'w:tcBorders', ('bottom',))


tokens = MarkdownIt('commonmark').enable('table').parse(corpo)
listas = []
linhas = []
citacao = 0
textual = False
secao_atual = ''
for i, token in enumerate(tokens):
    if token.type == 'bullet_list_open':
        listas.append('List Bullet')
    elif token.type == 'ordered_list_open':
        listas.append('List Number')
        # cada lista numerada recomeça do 1
        numeracao = documento.part.numbering_part.element
        estilo_lista = documento.styles['List Number'].element
        lista = numeracao.add_num(numeracao.num_having_numId(
            estilo_lista.pPr.numPr.numId.val).abstractNumId.val)
        lista.add_lvlOverride(ilvl=0).add_startOverride(1)
    elif token.type in ('bullet_list_close', 'ordered_list_close'):
        listas.pop()
    elif token.type == 'blockquote_open':
        citacao += 1
    elif token.type == 'blockquote_close':
        citacao -= 1
    elif token.type == 'tr_open':
        linhas.append([])
    elif token.type in ('th_open', 'td_open'):
        linhas[-1].append((tokens[i + 1], token.attrGet('style')))
    elif token.type == 'table_close':
        criar_tabela(linhas)
        linhas = []
    elif token.type == 'inline' and linhas:
        continue
    elif token.type == 'inline' and tokens[i - 1].type == 'heading_open':
        secao_atual = token.content
        nivel = int(tokens[i - 1].tag[1])
        numerado = token.content[0].isdigit()
        inicio_textual = nivel == 1 and numerado and not textual
        if inicio_textual:
            # a numeração aparece a partir da primeira seção numerada
            numerar(documento.add_section(WD_SECTION.NEW_PAGE))
            quebrar = False
            textual = True
        if nivel == 1 and not textual:
            p = paragrafo('Titulo pre-textual')
        else:
            p = paragrafo(f'Heading {nivel}')
        if inicio_textual:
            p.paragraph_format.page_break_before = False
        if nivel == 1 and not numerado:
            p.alignment = CENTRO
        if nivel >= 3 and numerado:
            numero = token.content.split(' ', 1)[0]
            p.add_run(f'{numero} ').bold = True
            token.children[0].content = token.children[0].content[len(numero) + 1:]
        escrever(p, token)
        if nivel == 1 and not numerado:
            # mantém a caixa escrita, como em 'APÊNDICE A – Questionário'
            for run in p.runs:
                run.font.all_caps = False
        if secao_atual == 'SUMÁRIO':
            campo(paragrafo('Normal'), "TOC \\o '1-3' \\h \\z \\u",
                  'O sumário é preenchido no push para a main.')
    elif token.type == 'inline':
        if listas:
            nome = listas[-1]
        elif citacao:
            nome = 'Citacao'
        elif any(filho.type == 'image' for filho in token.children):
            nome = 'Figura'
        elif token.content.startswith(LEGENDAS):
            nome = 'Legenda'
        elif token.content.startswith('Fonte:'):
            nome = 'Fonte'
        elif secao_atual in PRE_TEXTUAIS_CORRIDOS:
            nome = 'Resumo'
        elif secao_atual == 'REFERÊNCIAS':
            nome = 'Referencia'
        else:
            nome = 'Normal'
        p = paragrafo(nome)
        if nome == 'List Number':
            p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = (lista.numId)
        escrever(p, token)

documento.save(PASTA / ARQUIVO)
print(f'Gerado: projeto/{ARQUIVO}')
