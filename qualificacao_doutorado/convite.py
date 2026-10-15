"""Preenche o convite e ajusta o corpo sem mover a arte ou a assinatura."""
from zipfile import ZipFile

from docxtpl import DocxTemplate
from jinja2 import Environment, StrictUndefined
from lxml import etree
import pymupdf

from layout import NS, FONTE_MINIMA, canonico, elemento, quebrar_linhas


# Caixa do corpo no modelo, em pontos da página. A assinatura fica abaixo dela.
CAIXA_CORPO = (129.1, 144.6, 567.7, 554.6)
FONTE_MAXIMA = 12


def gerar_docx_convite(modelo, destino, contexto, limite=FONTE_MAXIMA):
    tpl = DocxTemplate(modelo)
    tpl.render(contexto, jinja_env=Environment(undefined=StrictUndefined), autoescape=True)
    tpl.save(destino)
    with ZipFile(destino) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    xml = etree.fromstring(dict((i.filename, data) for i, data in items)["word/document.xml"])
    corpos = xml.xpath('//wp:anchor[wp:docPr/@name="PPEB_convite"]//w:txbxContent | '
                       '//v:shape[@id="PPEB_convite"]//w:txbxContent', namespaces=NS)
    if len(corpos) != 2:
        raise ValueError("Modelo de convite incompatível. Use o modelo com a caixa PPEB_convite.")
    paragrafos = ["".join(p.xpath('.//w:t/text()', namespaces=NS))
                  for p in corpos[0].findall('w:p', NS)]
    largura = (CAIXA_CORPO[2] - CAIXA_CORPO[0] - 14.4) * 0.94
    altura = CAIXA_CORPO[3] - CAIXA_CORPO[1] - 7.2
    for meio_ponto in range(int(limite * 2), FONTE_MINIMA * 2 - 1, -1):
        fonte = meio_ponto / 2
        linhas = sum(len(quebrar_linhas(t, largura, fonte)) for t in paragrafos)
        if linhas * fonte * 1.2 + max(0, len(paragrafos) - 1) * 8 <= altura:
            break
    else:
        raise ValueError("O texto do convite não cabe nem com fonte 8. Reduza o texto na planilha.")
    for corpo in corpos:
        for p in corpo.findall('w:p', NS):
            props = p.find('w:pPr', NS)
            if props is None:
                props = elemento('w:pPr')
                p.insert(0, props)
            for spacing in props.findall('w:spacing', NS):
                props.remove(spacing)
            props.append(elemento('w:spacing', before=0, after=160,
                                  line=round(fonte * 1.2 * 20), lineRule='exact'))
            for r in p.findall('w:r', NS):
                rprops = r.find('w:rPr', NS)
                if rprops is None:
                    rprops = elemento('w:rPr')
                    r.insert(0, rprops)
                for tag in ('sz', 'szCs'):
                    for old in rprops.findall('w:' + tag, NS):
                        rprops.remove(old)
                    rprops.append(elemento('w:' + tag, val=round(fonte * 2)))
    with ZipFile(destino, 'w') as z:
        for item, data in items:
            if item.filename == 'word/document.xml':
                data = etree.tostring(xml, xml_declaration=True, encoding='UTF-8', standalone=True)
            z.writestr(item, data)
    return fonte, " ".join(paragrafos)


def conferir_pdf_convite(pdf, esperado):
    """Uma página só pode esconder texto cortado na caixa; confira o corpo todo."""
    with pymupdf.open(pdf) as doc:
        if len(doc) != 1:
            return False
        visivel = doc[0].get_text('text', clip=pymupdf.Rect(CAIXA_CORPO), sort=True)
        return canonico(visivel) == canonico(esperado)
