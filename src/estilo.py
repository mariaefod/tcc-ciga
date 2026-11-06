from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

FONTE = 'Arial'
FORMATO_NUMERO = '%1)'
TAMANHO_NUMERO_PAGINA = 10

ESQUERDA = WD_ALIGN_PARAGRAPH.LEFT
CENTRO = WD_ALIGN_PARAGRAPH.CENTER
DIREITA = WD_ALIGN_PARAGRAPH.RIGHT
JUSTIFICADO = WD_ALIGN_PARAGRAPH.JUSTIFY

MARGEM_SUPERIOR = 3
MARGEM_INFERIOR = 2
MARGEM_ESQUERDA = 3
MARGEM_DIREITA = 2
LARGURA_UTIL = 21 - MARGEM_ESQUERDA - MARGEM_DIREITA

PADRAO = dict(
    tamanho=12,
    negrito=False,
    maiusculas=False,
    alinhamento=JUSTIFICADO,
    recuo=0,
    recuo_esquerdo=0,
    entrelinha=1.5,
    antes=0,
    depois=0,
    nova_pagina=False,
)

ESTILOS = {
    'Normal': dict(recuo=1.25, depois=6),
    # 1 TÍTULO DE SEÇÃO PRIMÁRIA: maiúsculas, negrito, nova página
    'Heading 1': dict(negrito=True, maiusculas=True, alinhamento=ESQUERDA, depois=18, nova_pagina=True),
    # 1.1 Título de seção secundária: negrito
    'Heading 2': dict(negrito=True, alinhamento=ESQUERDA, antes=18, depois=18),
    # 1.1.1 Título de seção terciária: só o número em negrito
    'Heading 3': dict(alinhamento=ESQUERDA, antes=18, depois=18),
    # RESUMO, ABSTRACT, SUMÁRIO: centralizados e fora do sumário
    'Titulo pre-textual': dict(negrito=True, maiusculas=True, alinhamento=CENTRO, depois=18, nova_pagina=True),
    # listas
    'List Bullet': dict(recuo=-0.635, recuo_esquerdo=1.27, depois=6),
    'List Number': dict(recuo=-0.635, recuo_esquerdo=1.27, depois=6),
    # citação com mais de 3 linhas
    'Citacao': dict(tamanho=10, recuo_esquerdo=4, entrelinha=1, antes=12, depois=12),
    'Resumo': dict(entrelinha=1, depois=12),
    'Capa': dict(negrito=True, maiusculas=True, alinhamento=CENTRO),
    # natureza do trabalho na folha de rosto
    'Natureza': dict(recuo_esquerdo=8, entrelinha=1),
    # título da tabela, figura, quadro ou gráfico
    'Legenda': dict(tamanho=10, alinhamento=ESQUERDA, entrelinha=1, antes=12),
    'Figura': dict(alinhamento=CENTRO, entrelinha=1),
    'Tabela': dict(tamanho=10, alinhamento=ESQUERDA, entrelinha=1, antes=1, depois=1),
    'Fonte': dict(tamanho=10, alinhamento=ESQUERDA, entrelinha=1, antes=2, depois=12),
    'Referencia': dict(alinhamento=ESQUERDA, entrelinha=1, depois=12),
    # linhas do sumário
    'toc 1': dict(negrito=True, maiusculas=True, alinhamento=ESQUERDA),
    'toc 2': dict(negrito=True, alinhamento=ESQUERDA),
    'toc 3': dict(alinhamento=ESQUERDA),
}

# espaço antes de cada linha (pt)
CAPA = {
    'instituicao': 0,
    'escola': 0,
    'curso': 0,
    'autor': 120,
    'titulo': 150,
    'cidade': 220,
    'ano': 0,
}

FOLHA_DE_ROSTO = {
    'autor': 0,
    'titulo': 150,
    'natureza': 60,
    'orientadora': 12,
    'cidade': 180,
    'ano': 0,
}


def aplicar(documento):
    for nome, valores in ESTILOS.items():
        v = {**PADRAO, **valores}
        if nome in documento.styles:
            s = documento.styles[nome]
        else:
            s = documento.styles.add_style(nome, WD_STYLE_TYPE.PARAGRAPH)
            s.base_style = documento.styles['Normal']
        s.font.name = FONTE
        s.font.size = Pt(v['tamanho'])
        s.font.bold = v['negrito']
        s.font.all_caps = v['maiusculas']
        s.font.italic = False
        rpr = s.element.get_or_add_rPr()
        fontes = rpr.get_or_add_rFonts()
        for atributo in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
            fontes.attrib.pop(qn(f'w:{atributo}'), None)
        fontes.set(qn('w:eastAsia'), FONTE)
        cor = rpr.get_or_add_color()
        cor.attrib.pop(qn('w:themeColor'), None)
        cor.set(qn('w:val'), 'auto')
        f = s.paragraph_format
        f.alignment = v['alinhamento']
        f.first_line_indent = Cm(v['recuo'])
        f.left_indent = Cm(v['recuo_esquerdo'])
        f.line_spacing = v['entrelinha']
        f.space_before = Pt(v['antes'])
        f.space_after = Pt(v['depois'])
        f.page_break_before = v['nova_pagina']
        if nome.startswith('toc'):
            f.tab_stops.add_tab_stop(Cm(LARGURA_UTIL), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

    idioma = OxmlElement('w:lang')
    idioma.set(qn('w:val'), 'pt-BR')
    documento.styles['Normal'].element.get_or_add_rPr().append(idioma)
    documento.styles['Legenda'].paragraph_format.keep_with_next = True

    numeracao = documento.part.numbering_part.element
    for nome in ('List Bullet', 'List Number'):
        s = documento.styles[nome]
        num_id = s.element.pPr.numPr.numId.val
        abstrato = numeracao.num_having_numId(num_id).abstractNumId.val
        nivel = (f"w:abstractNum[@w:abstractNumId='{abstrato}']/w:lvl[@w:ilvl='0']")
        esquerda = str(s.paragraph_format.left_indent.twips)
        tab = numeracao.xpath(f'{nivel}/w:pPr/w:tabs/w:tab')[0]
        tab.set(qn('w:pos'), esquerda)
        recuo = numeracao.xpath(f'{nivel}/w:pPr/w:ind')[0]
        recuo.set(qn('w:left'), esquerda)
        recuo.set(qn('w:hanging'), str(-s.paragraph_format.first_line_indent.twips))
        if nome == 'List Number':
            texto = numeracao.xpath(f'{nivel}/w:lvlText')[0]
            texto.set(qn('w:val'), FORMATO_NUMERO)

    secao = documento.sections[0]
    secao.page_width = Cm(21)
    secao.page_height = Cm(29.7)
    secao.top_margin = Cm(MARGEM_SUPERIOR)
    secao.bottom_margin = Cm(MARGEM_INFERIOR)
    secao.left_margin = Cm(MARGEM_ESQUERDA)
    secao.right_margin = Cm(MARGEM_DIREITA)
    # a capa não conta; a contagem começa na folha de rosto
    numeracao = OxmlElement('w:pgNumType')
    numeracao.set(qn('w:start'), '0')
    secao._sectPr.append(numeracao)


def numerar(secao):
    # continua a contagem da seção anterior
    secao._sectPr.remove(secao._sectPr.find(qn('w:pgNumType')))
    secao.header.is_linked_to_previous = False
    pagina = secao.header.paragraphs[0]
    pagina.alignment = DIREITA
    pagina.paragraph_format.first_line_indent = Cm(0)
    run = pagina.add_run()
    run.font.name = FONTE
    run.font.size = Pt(TAMANHO_NUMERO_PAGINA)
    for tipo in ('begin', 'instrText', 'end'):
        if tipo == 'instrText':
            elemento = OxmlElement('w:instrText')
            elemento.text = 'PAGE'
        else:
            elemento = OxmlElement('w:fldChar')
            elemento.set(qn('w:fldCharType'), tipo)
        run._r.append(elemento)
